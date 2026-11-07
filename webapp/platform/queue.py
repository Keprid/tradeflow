"""Durable background job worker.

The queue lives in SQLite, so a restart does not lose work: rows left in
``running`` are requeued at startup and picked up again.  The actual build
function is injected, keeping this module independent of the pipelines.
"""

import logging
import threading
import time

from . import metrics, registry, settings

LOG = logging.getLogger("tradeflow.queue")


class JobQueue:
    def __init__(self, runner):
        self.runner = runner
        self._stop = threading.Event()
        self._wake = threading.Event()
        self._threads = []

    def start(self):
        registry.init_db()
        registry.reset_running()
        metrics.set_gauge("tradeflow_queue_depth", registry.queue_depth())
        if self._threads:
            return
        for index in range(settings.WORKER_COUNT):
            thread = threading.Thread(
                target=self._loop, name=f"tradeflow-worker-{index}",
                daemon=True)
            thread.start()
            self._threads.append(thread)
        LOG.info("Job queue started with %d worker(s)", settings.WORKER_COUNT)

    def stop(self):
        self._stop.set()
        self._wake.set()
        for thread in self._threads:
            thread.join(timeout=5)
        self._threads = []

    def notify(self):
        self._wake.set()

    def _loop(self):
        while not self._stop.is_set():
            job = None
            try:
                job = registry.claim_next()
            except Exception:  # noqa: BLE001 - keep the worker alive
                LOG.exception("Could not claim a queued job")
            if job is None:
                self._wake.wait(0.6)
                self._wake.clear()
                continue
            metrics.set_gauge("tradeflow_queue_depth", registry.queue_depth())
            metrics.inc("tradeflow_jobs_started_total")
            started = time.time()
            try:
                self.runner(job)
            except Exception as exc:  # noqa: BLE001
                LOG.exception("Job %s crashed in the worker", job["id"])
                attempts = int(job.get("attempts") or 0)
                max_attempts = int(job.get("max_attempts") or 1)
                if attempts < max_attempts and not self._stop.is_set():
                    registry.requeue(job["id"], str(exc))
                    metrics.inc("tradeflow_jobs_retried_total")
                else:
                    registry.set_status(job["id"], "error", error=str(exc))
                    metrics.inc("tradeflow_jobs_failed_total")
            finally:
                metrics.set_gauge(
                    "tradeflow_job_duration_seconds", time.time() - started)
                metrics.set_gauge("tradeflow_queue_depth",
                                  registry.queue_depth())

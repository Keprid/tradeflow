"""Dependency-free counters rendered in Prometheus text format."""

import threading
import time
from collections import defaultdict

_LOCK = threading.Lock()
_COUNTERS = defaultdict(float)
_GAUGES = {}
_START = time.time()


def inc(name, value=1.0):
    with _LOCK:
        _COUNTERS[name] += value


def set_gauge(name, value):
    with _LOCK:
        _GAUGES[name] = float(value)


def get(name):
    with _LOCK:
        return _COUNTERS.get(name, 0.0)


def render(extra_gauges=None):
    gauges = dict(_GAUGES)
    if extra_gauges:
        gauges.update({k: float(v) for k, v in extra_gauges.items()})
    gauges.setdefault("tradeflow_uptime_seconds", time.time() - _START)
    lines = []
    with _LOCK:
        counters = dict(_COUNTERS)
    for name, value in sorted(counters.items()):
        lines.append(f"# TYPE {name} counter")
        lines.append(f"{name} {value:g}")
    for name, value in sorted(gauges.items()):
        lines.append(f"# TYPE {name} gauge")
        lines.append(f"{name} {value:g}")
    return "\n".join(lines) + "\n"

"""KNBS reference series connector.

The portal is a Hub-and-Node SPA, not the SDMX REST service this connector
once assumed: cubes come from ``POST .../datasets/{flow}/data`` as JSON-stat
2.0.  ``knbs_reference`` already owns that client and the unit scales verified
against the printed Economic Survey, so this connector rebuilds the same
snapshot the report pipelines read and hands it back as a download.

A reference source cannot build a report.  Its snapshot is installed into
``data/knbs_reference.json`` by ``POST /api/fetches/{id}/install``, which is
what makes the next build read the new vintage.
"""

import csv
import io
import json
import sys
import time
import urllib.error
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import knbs_reference as K  # noqa: E402

from .base import Connector, ConnectorError, FetchResult  # noqa: E402

SNAPSHOT_FILE = "knbs_reference.json"
OBSERVATIONS_FILE = "knbs_observations.csv"

# Where an installed snapshot actually shows up.  KNBS is a cross-check layer,
# not a report input, so the app has to say which builds read it and whereabouts
# in the document to look.
USED_BY = (
    "Trade Brief - Table 4a (bilateral cross-check) and Table 7a (exports by "
    "region)",
    "Services Brief - Figure 12's 2025 tail and Table 13a (services balance "
    "cross-check)",
    "Product Profile - Table 12a (export cross-check) and Table 12b (sectoral "
    "context: area, output, yield, farm-gate prices)",
)


class KnbsConnector(Connector):
    id = "knbs"
    label = "KNBS reference series"
    description = ("Kenya National Bureau of Statistics, the Economic Survey "
                   "tables behind the open data portal. Rebuilds the snapshot "
                   "reports cross-check their dollar figures against, and the "
                   "sectoral context a product profile prints. A cross-check "
                   "layer, not a report input: install it, then generate from "
                   "your ITC uploads or a UNCTADstat fetch.")
    report_types = ("reference",)
    used_by = USED_BY
    params = ()

    def fetch(self, values):
        self.validate(values)
        try:
            snapshot = K.build_snapshot()
        except K.KnbsError as exc:
            raise ConnectorError(f"KNBS snapshot failed: {exc}",
                                 status_code=502) from exc
        except (urllib.error.URLError, OSError, ValueError) as exc:
            raise ConnectorError(f"Could not read the KNBS portal: {exc}",
                                 status_code=502) from exc

        datasets = snapshot["datasets"]
        observations = sum(len(d["rows"]) for d in datasets.values())
        if not observations:
            raise ConnectorError(
                "The KNBS portal returned no observations; the snapshot was "
                "not built.", status_code=502)

        files = {
            SNAPSHOT_FILE: json.dumps(snapshot,
                                      separators=(",", ":")).encode("utf-8"),
            OBSERVATIONS_FILE: _observations_csv(datasets),
        }
        notes = [
            f"Rebuilt {len(datasets)} datasets, {observations} observations.",
            f"Snapshot taken {snapshot['generated_at']}; values are in base "
            "units (KSh, tonnes, headcount) at the scales verified against the "
            "printed Economic Survey.",
            "KNBS is a cross-check layer, not a report input. Generate a "
            "report from your ITC uploads or a UNCTADstat fetch; once "
            "installed, these series appear in:",
            *USED_BY,
        ]
        provenance = {
            "source": "Kenya National Bureau of Statistics",
            "dataset": "KNBS reference series (Economic Survey 2026)",
            "url": K.API_ROOT,
            "http_status": 200,
            "license": "KNBS terms of use",
            "retrieved_at": time.time(),
            "generated_at": snapshot["generated_at"],
            "datasets": sorted(datasets),
            "observations": observations,
            "install_file": SNAPSHOT_FILE,
        }
        return FetchResult(files=files, provenance=provenance, notes=notes)


def _observations_csv(datasets):
    """One row per observation, readable beside the dense snapshot JSON."""
    dims = set()
    rows = []
    for name, entry in sorted(datasets.items()):
        for key, value in entry["rows"].items():
            coords = dict(part.split("=", 1) for part in key.split("|"))
            dims.update(coords)
            rows.append(dict(coords, dataset=name,
                             survey_table=entry.get("survey_table", ""),
                             value=value))
    fields = (["dataset", "survey_table", "TIME_PERIOD"]
              + sorted(dims - {"TIME_PERIOD"}) + ["value"])
    stream = io.StringIO()
    writer = csv.DictWriter(stream, fieldnames=fields, restval="",
                            extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return stream.getvalue().encode("utf-8")

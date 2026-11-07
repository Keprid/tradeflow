"""Small platform layer shared by the web app and the data connectors.

This package is deliberately dependency-free (stdlib only) so the web app
keeps working on the existing requirements.  It adds:

* ``settings``   - environment-driven configuration and a persistent secret.
* ``registry``   - SQLite job/queue/provenance metadata.
* ``auth``       - per-job owner tokens and signed, expiring share links.
* ``provenance`` - SHA-256 fingerprints for uploaded/fetched files.
* ``metrics``    - in-process counters rendered as Prometheus text.
* ``queue``      - a durable, restart-safe worker for heavy report builds.
"""

"""Sanctioned open-data connectors exposed by the web app."""

from .knbs import KnbsConnector
from .unctad import UnctadConnector

_CONNECTORS = {
    connector.id: connector
    for connector in (UnctadConnector(), KnbsConnector())
}


def all_connectors():
    return sorted(_CONNECTORS.values(), key=lambda c: c.label)


def get_connector(connector_id):
    return _CONNECTORS.get(connector_id)


def connector_catalog():
    out = []
    for connector in all_connectors():
        out.append({
            "id": connector.id,
            "label": connector.label,
            "description": connector.description,
            "report_types": list(connector.report_types),
            "used_by": list(connector.used_by),
            "requires_credentials": connector.requires_credentials,
            "experimental": connector.experimental,
            "params": [p.as_dict() for p in connector.params],
        })
    return out

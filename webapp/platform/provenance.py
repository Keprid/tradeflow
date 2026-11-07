"""File fingerprints and provenance records for reproducible runs."""

import hashlib
import os
import time
from pathlib import Path


def sha256_file(path, chunk=1024 * 1024):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        while True:
            block = handle.read(chunk)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


def files_provenance(directory, source="upload"):
    directory = Path(directory)
    items = []
    if not directory.is_dir():
        return {"source": source, "captured_at": time.time(), "files": items}
    for path in sorted(directory.rglob("*")):
        if not path.is_file() or path.name.startswith(("~$", ".")):
            continue
        try:
            stat = path.stat()
            items.append({
                "name": str(path.relative_to(directory)).replace(os.sep, "/"),
                "bytes": stat.st_size,
                "sha256": sha256_file(path),
                "source": source,
            })
        except OSError:
            continue
    return {
        "source": source,
        "captured_at": time.time(),
        "file_count": len(items),
        "total_bytes": sum(i["bytes"] for i in items),
        "files": items,
    }

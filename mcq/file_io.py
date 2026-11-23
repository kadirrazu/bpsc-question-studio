"""Atomic writes shared by the studio modules."""

import json
from pathlib import Path
from tempfile import NamedTemporaryFile


def write_bytes(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=path.parent, delete=False) as temporary:
        staging = Path(temporary.name)
        try:
            temporary.write(data)
        except BaseException:
            temporary.close()
            staging.unlink(missing_ok=True)
            raise
    try:
        staging.replace(path)
    finally:
        staging.unlink(missing_ok=True)


def save_json(path, data):
    write_bytes(path, json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8"))

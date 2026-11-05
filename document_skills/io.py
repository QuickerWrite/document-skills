"""Bounded reads and atomic, non-overwriting output by default."""

import os
import tempfile
from pathlib import Path
from .render import MAX_BYTES, _bounded_bytes


def read(path):
    with Path(path).open("rb") as stream:
        return _bounded_bytes(stream.read(MAX_BYTES + 1))


def write(path, data, overwrite=False):
    target = Path(path)
    fd, temporary = tempfile.mkstemp(prefix=".documents-", dir=target.parent)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        if overwrite:
            os.replace(temporary, target)
        else:
            os.link(temporary, target)  # atomic fail-if-exists; no check/use race
    finally:
        Path(temporary).unlink(missing_ok=True)

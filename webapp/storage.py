"""Short-lived, in-memory result store.

Generated .docx bytes are handed to the browser via a two-step
render-result-page -> GET /download/<token> flow, so a token needs to
survive between those two requests without ever touching disk (avoids the
kind of file-lock problems a viewer/editor holding the file open caused
before). A small in-memory dict with a TTL is enough for a single-process
personal/small-team tool; it is not meant to survive a server restart.
"""

import threading
import time
import uuid

_TTL_SECONDS = 20 * 60
_lock = threading.Lock()
_store = {}


def put(data: bytes, filename: str) -> str:
    token = uuid.uuid4().hex
    with _lock:
        _prune_locked()
        _store[token] = {"data": data, "filename": filename, "created": time.monotonic()}
    return token


def get(token: str):
    with _lock:
        _prune_locked()
        return _store.get(token)


def _prune_locked():
    cutoff = time.monotonic() - _TTL_SECONDS
    expired = [k for k, v in _store.items() if v["created"] < cutoff]
    for k in expired:
        del _store[k]

"""
Concurrency-safe presentation state for the remote MCP server.

Replaces the upstream module-level ``presentations = {}`` dict and the
``current_presentation_id`` global, which were designed for a single local
stdio client and are unsafe for a shared multi-tenant server.

Design (workstream 5.3):
- Presentation handles are UUID4 hex strings generated server-side.
  They act as unguessable capabilities: a conversation can only touch
  decks whose handle it was given. Client-chosen IDs are not accepted.
- The store is a drop-in ``MutableMapping`` so the ~27 existing
  ``pres_id in presentations`` / ``presentations[pres_id]`` call sites in
  tools/ keep working unchanged.
- Entries expire after ``PPT_MCP_STATE_TTL_SECONDS`` (default 3600) of
  inactivity and the store holds at most ``PPT_MCP_STATE_MAX_PRESENTATIONS``
  (default 50) decks, evicting least-recently-used — a remote server must
  bound its memory.
- Each presentation has an ``RLock``; ``lock_for(pres_id)`` lets the server
  serialize tool calls that target the same deck (python-pptx objects are
  not thread-safe). Calls on different decks run concurrently.
"""
import os
import threading
import time
import uuid
from collections.abc import MutableMapping


class PresentationStore(MutableMapping):
    def __init__(self, ttl_seconds=None, max_items=None):
        self._ttl = ttl_seconds if ttl_seconds is not None else int(
            os.environ.get("PPT_MCP_STATE_TTL_SECONDS", "3600"))
        self._max = max_items if max_items is not None else int(
            os.environ.get("PPT_MCP_STATE_MAX_PRESENTATIONS", "50"))
        self._lock = threading.RLock()
        self._items = {}  # pres_id -> {"pres": ..., "lock": RLock, "last_used": float}

    @staticmethod
    def new_id():
        return uuid.uuid4().hex

    def _purge(self):
        """Drop expired entries; then LRU-evict down to max size. Caller holds _lock."""
        now = time.monotonic()
        expired = [k for k, v in self._items.items()
                   if now - v["last_used"] > self._ttl]
        for k in expired:
            del self._items[k]
        while len(self._items) > self._max:
            oldest = min(self._items, key=lambda k: self._items[k]["last_used"])
            del self._items[oldest]

    def __setitem__(self, pres_id, pres):
        with self._lock:
            self._items[pres_id] = {
                "pres": pres,
                "lock": threading.RLock(),
                "last_used": time.monotonic(),
            }
            self._purge()

    def __getitem__(self, pres_id):
        with self._lock:
            self._purge()
            entry = self._items[pres_id]  # KeyError propagates
            entry["last_used"] = time.monotonic()
            return entry["pres"]

    def __delitem__(self, pres_id):
        with self._lock:
            del self._items[pres_id]

    def __contains__(self, pres_id):
        with self._lock:
            self._purge()
            return pres_id in self._items

    def __iter__(self):
        with self._lock:
            return iter(list(self._items))

    def __len__(self):
        with self._lock:
            return len(self._items)

    def lock_for(self, pres_id):
        """Per-presentation lock, or a throwaway lock for unknown/missing IDs
        (the tool will then return its normal not-found error)."""
        with self._lock:
            entry = self._items.get(pres_id)
            return entry["lock"] if entry else threading.RLock()


def serialize_per_presentation(app, store):
    """Wrap every registered tool so calls holding the same presentation_id
    are serialized on that deck's lock (python-pptx is not thread-safe).

    Uses the FastMCP 1.x tool registry (mcp[cli] is pinned <2.0 in
    requirements.txt); degrades to a no-op with a warning if the SDK's
    internals change.
    """
    try:
        tools = app._tool_manager._tools
    except AttributeError:
        print("Warning: could not install per-presentation locking "
              "(unexpected FastMCP internals); same-deck concurrent calls "
              "are not serialized.")
        return

    def wrap(fn):
        def wrapper(*args, **kwargs):
            pres_id = kwargs.get("presentation_id")
            with store.lock_for(pres_id):
                return fn(*args, **kwargs)
        wrapper.__name__ = getattr(fn, "__name__", "tool")
        return wrapper

    for tool in tools.values():
        tool.fn = wrap(tool.fn)

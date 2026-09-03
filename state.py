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
- ``serialize_per_presentation`` also dispatches every tool call to a worker
  thread (see its docstring) so that one blocking call — a vision-LLM review,
  a DIAL upload, a LibreOffice conversion — cannot stall every other call the
  process is handling. Concurrency is bounded by
  ``PPT_MCP_MAX_CONCURRENT_TOOL_CALLS`` to keep a single pod from being
  flooded.
"""
import logging
import os
import threading
import time
import uuid
from collections.abc import MutableMapping

import anyio
import anyio.to_thread

from logging_utils import get_logger, flatten

logger = get_logger("state")


def short_id(pres_id):
    """Handles are unguessable capabilities — log only a prefix, never the
    whole thing, so logs can be correlated without leaking a usable handle."""
    return f"{pres_id[:8]}…" if isinstance(pres_id, str) and pres_id else "-"


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
            logger.info("presentation_expired presentation_id=%s ttl_seconds=%d "
                        "held=%d", short_id(k), self._ttl, len(self._items))
        while len(self._items) > self._max:
            oldest = min(self._items, key=lambda k: self._items[k]["last_used"])
            del self._items[oldest]
            logger.warning("presentation_evicted presentation_id=%s reason=lru "
                           "max=%d", short_id(oldest), self._max)

    def __setitem__(self, pres_id, pres):
        with self._lock:
            self._items[pres_id] = {
                "pres": pres,
                "lock": threading.RLock(),
                "last_used": time.monotonic(),
                # A deck that was never inspected (or edited since its last
                # passed inspection) needs visual QA before export.
                "needs_inspection": True,
            }
            self._purge()
            logger.info("presentation_stored presentation_id=%s held=%d",
                        short_id(pres_id), len(self._items))

    def __getitem__(self, pres_id):
        with self._lock:
            self._purge()
            entry = self._items[pres_id]  # KeyError propagates
            entry["last_used"] = time.monotonic()
            return entry["pres"]

    def __delitem__(self, pres_id):
        with self._lock:
            del self._items[pres_id]
            logger.debug("presentation_dropped presentation_id=%s held=%d",
                         short_id(pres_id), len(self._items))

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

    def mark_dirty(self, pres_id):
        with self._lock:
            entry = self._items.get(pres_id)
            if entry:
                entry["needs_inspection"] = True

    def clear_dirty(self, pres_id):
        with self._lock:
            entry = self._items.get(pres_id)
            if entry:
                entry["needs_inspection"] = False
                logger.debug("presentation_inspection_cleared presentation_id=%s",
                             short_id(pres_id))

    def is_dirty(self, pres_id):
        with self._lock:
            entry = self._items.get(pres_id)
            return bool(entry and entry["needs_inspection"])

    def lock_for(self, pres_id):
        """Per-presentation lock, or a throwaway lock for unknown/missing IDs
        (the tool will then return its normal not-found error)."""
        with self._lock:
            entry = self._items.get(pres_id)
            return entry["lock"] if entry else threading.RLock()


# Tools that must not re-flag a passed deck as needing inspection: exports,
# and the visual QA tools — visual_repair_slides does edit the deck, but it
# re-inspects what it edited and manages the flag itself.
#
# manage_speaker_notes is here because the flag tracks *visual* staleness and
# notes are never rendered on a slide: writing them cannot invalidate a
# passing inspection, and marking the deck dirty would cost a whole re-review
# for text the audience never sees.
_NON_EDITING_TOOLS = {"export_presentation", "save_presentation",
                      "visual_inspect_slides", "visual_repair_slides",
                      "manage_speaker_notes"}


def _elapsed_ms(started):
    return int((time.monotonic() - started) * 1000)


# Tool arguments can carry a whole base64 template or slide text; log only
# names and sizes, never the payloads.
def _arg_summary(kwargs, limit=200):
    parts = []
    for key, value in kwargs.items():
        if key == "presentation_id":
            continue
        if isinstance(value, str):
            parts.append(f"{key}=<str:{len(value)}>" if len(value) > 40
                         else f"{key}={flatten(value)}")
        elif isinstance(value, (list, tuple, dict)):
            parts.append(f"{key}=<{type(value).__name__}:{len(value)}>")
        elif value is not None:
            parts.append(f"{key}={value}")
    return " ".join(parts)[:limit] or "-"


def _max_concurrent_tool_calls():
    """How many tool calls may run at once, process-wide. Every tool in this
    codebase is a plain ``def`` (see serialize_per_presentation) doing
    blocking work — python-pptx edits, PyMuPDF rasterization, a LibreOffice
    subprocess, or a synchronous httpx call to DIAL/a vision LLM. Dispatching
    them to worker threads lets independent calls overlap instead of
    serializing on the single asyncio event loop, but an unbounded thread
    pool would let a burst of requests (many tenants, or one agent firing a
    dozen parallel render_svg_icon calls) pile up enough CPU-bound work to
    starve the pod. The default assumes most calls are network-bound (vision
    review, DIAL I/O) rather than CPU-bound, so it leaves headroom above the
    CPU count; override with PPT_MCP_MAX_CONCURRENT_TOOL_CALLS if a
    deployment's mix runs heavier on rendering/conversion."""
    raw = os.environ.get("PPT_MCP_MAX_CONCURRENT_TOOL_CALLS", "").strip()
    if raw:
        try:
            return max(1, int(raw))
        except ValueError:
            logger.warning("invalid_max_concurrent_tool_calls value=%r "
                           "falling_back_to_default", raw)
    return max(4, (os.cpu_count() or 4) * 2)


def serialize_per_presentation(app, store):
    """Wrap every registered tool so that (a) it runs on a worker thread,
    bounded by a process-wide capacity limiter, instead of blocking the
    single asyncio event loop FastMCP dispatches tools on — every tool here
    is a synchronous function doing blocking I/O or CPU work, and the SDK
    calls a sync tool's fn directly rather than off-loading it (see
    mcp.server.fastmcp.utilities.func_metadata.call_fn_with_arg_validation),
    so without this a single slow call (a vision-LLM review, a DIAL upload, a
    LibreOffice conversion) stalls every other call the process is handling,
    including ones for an unrelated deck; (b) calls holding the same
    presentation_id are still serialized on that deck's lock (python-pptx is
    not thread-safe — the lock is acquired inside the worker thread, so it
    only blocks other calls for the same deck, never the event loop); and
    (c) successful editing calls mark the deck as needing visual inspection
    before its next export (see export gate in tools/presentation_tools.py).

    Marks each wrapped tool ``is_async = True`` after replacing its ``fn``:
    FastMCP decides whether to await a tool from a flag captured at
    registration time from the *original* function, not from whatever ``fn``
    holds when it is actually called, so the flag has to be corrected by hand
    once ``fn`` becomes a coroutine function.

    Uses the FastMCP 1.x tool registry (mcp[cli] is pinned <2.0 in
    requirements.txt); degrades to a no-op with a warning if the SDK's
    internals change.
    """
    try:
        tools = app._tool_manager._tools
    except AttributeError:
        logger.warning(
            "tool_locking_unavailable reason=unexpected_fastmcp_internals "
            "impact=same_deck_calls_not_serialized_and_not_parallelized")
        return

    max_concurrent = _max_concurrent_tool_calls()
    limiter = anyio.CapacityLimiter(max_concurrent)
    logger.info("tool_concurrency_configured max_concurrent=%d", max_concurrent)

    def wrap(name, fn, read_only):
        marks_dirty = not read_only and name not in _NON_EDITING_TOOLS

        async def wrapper(*args, **kwargs):
            pres_id = kwargs.get("presentation_id")
            call_logger = get_logger(f"tool.{name}")
            call_logger.debug("tool_start tool=%s presentation_id=%s args=%s",
                              name, short_id(pres_id), _arg_summary(kwargs))
            started = time.monotonic()

            def run_locked():
                with store.lock_for(pres_id):
                    return fn(*args, **kwargs)

            try:
                result = await anyio.to_thread.run_sync(run_locked,
                                                         limiter=limiter)
            except Exception as e:
                call_logger.error(
                    "tool_raised tool=%s presentation_id=%s duration_ms=%d "
                    "error=%s", name, short_id(pres_id),
                    _elapsed_ms(started), e,
                    exc_info=call_logger.isEnabledFor(logging.DEBUG))
                raise
            failed = isinstance(result, dict) and "error" in result
            if failed:
                call_logger.warning(
                    "tool_error tool=%s presentation_id=%s duration_ms=%d "
                    "error=%s", name, short_id(pres_id),
                    _elapsed_ms(started), flatten(str(result["error"]))[:300])
            else:
                call_logger.info("tool_ok tool=%s presentation_id=%s duration_ms=%d",
                                 name, short_id(pres_id), _elapsed_ms(started))
                if marks_dirty and pres_id is not None:
                    store.mark_dirty(pres_id)
            return result
        wrapper.__name__ = getattr(fn, "__name__", "tool")
        return wrapper

    for name, tool in tools.items():
        annotations = getattr(tool, "annotations", None)
        read_only = bool(annotations and getattr(annotations, "readOnlyHint", False))
        tool.fn = wrap(name, tool.fn, read_only)
        tool.is_async = True

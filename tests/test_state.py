"""Unit tests for the concurrency-safe PresentationStore (workstream 5.3)."""
import asyncio
import sys
import threading
import time
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from state import PresentationStore, serialize_per_presentation


class TestPresentationStore(unittest.TestCase):
    def test_uuid_handles(self):
        store = PresentationStore(ttl_seconds=60, max_items=10)
        a, b = store.new_id(), store.new_id()
        self.assertNotEqual(a, b)
        self.assertEqual(len(a), 32)

    def test_roundtrip_and_membership(self):
        store = PresentationStore(ttl_seconds=60, max_items=10)
        pid = store.new_id()
        store[pid] = "deck"
        self.assertIn(pid, store)
        self.assertEqual(store[pid], "deck")
        self.assertNotIn("presentation_1", store)

    def test_ttl_expiry(self):
        store = PresentationStore(ttl_seconds=0, max_items=10)
        pid = store.new_id()
        store[pid] = "deck"
        import time
        time.sleep(0.01)
        self.assertNotIn(pid, store)

    def test_lru_eviction(self):
        store = PresentationStore(ttl_seconds=60, max_items=2)
        ids = [store.new_id() for _ in range(3)]
        for i, pid in enumerate(ids):
            store[pid] = i
        self.assertNotIn(ids[0], store)  # oldest evicted
        self.assertIn(ids[1], store)
        self.assertIn(ids[2], store)

    def test_per_presentation_lock_serializes(self):
        store = PresentationStore(ttl_seconds=60, max_items=10)
        pid = store.new_id()
        store[pid] = []
        order = []

        def worker(tag):
            with store.lock_for(pid):
                order.append((tag, "in"))
                import time
                time.sleep(0.02)
                order.append((tag, "out"))

        threads = [threading.Thread(target=worker, args=(t,)) for t in "ab"]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        # Critical sections must not interleave
        self.assertEqual([e[1] for e in order], ["in", "out", "in", "out"])


if __name__ == "__main__":
    unittest.main()


class _FakeTool:
    """Minimal stand-in for the SDK's Tool object (name, fn, annotations)."""

    def __init__(self, fn, read_only=False):
        self.fn = fn
        self.annotations = type("A", (), {"readOnlyHint": read_only})()


class _FakeApp:
    def __init__(self, tools):
        self._tool_manager = type("TM", (), {"_tools": tools})()


class TestToolCallLogging(unittest.TestCase):
    """Every tool call leaves exactly one line, at a level matching outcome."""

    def _run(self, fn, read_only=False, **kwargs):
        store = PresentationStore(ttl_seconds=60, max_items=10)
        pid = store.new_id()
        store[pid] = object()
        tools = {"demo_tool": _FakeTool(fn, read_only)}
        serialize_per_presentation(_FakeApp(tools), store)
        # Wrapped tools are coroutine functions now — see is_async below.
        self.assertTrue(tools["demo_tool"].is_async)
        with self.assertLogs("dial_pptx.tool.demo_tool", level="DEBUG") as caught:
            try:
                result = asyncio.run(
                    tools["demo_tool"].fn(presentation_id=pid, **kwargs))
            except RuntimeError:
                result = None
        return caught.records, result, store, pid

    def test_success_logs_one_info_line_and_marks_dirty(self):
        store_records, result, store, pid = self._run(lambda **kw: {"ok": True})
        info = [r for r in store_records if r.levelname == "INFO"]
        self.assertEqual(len(info), 1)
        self.assertIn("tool_ok tool=demo_tool", info[0].getMessage())
        self.assertIn("duration_ms=", info[0].getMessage())
        self.assertTrue(store.is_dirty(pid))

    def test_error_result_logs_a_warning(self):
        records, _, store, pid = self._run(lambda **kw: {"error": "bad input"})
        warnings = [r for r in records if r.levelname == "WARNING"]
        self.assertEqual(len(warnings), 1)
        self.assertIn("tool_error", warnings[0].getMessage())
        self.assertIn("bad input", warnings[0].getMessage())

    def test_raised_exception_logs_an_error_and_propagates(self):
        def boom(**kwargs):
            raise RuntimeError("kaboom")

        records, _, _, _ = self._run(boom)
        errors = [r for r in records if r.levelname == "ERROR"]
        self.assertEqual(len(errors), 1)
        self.assertIn("tool_raised", errors[0].getMessage())

    def test_debug_line_summarizes_arguments_without_payloads(self):
        records, _, _, _ = self._run(lambda **kw: {"ok": True},
                                     template_content="A" * 5000, slide_index=2)
        debug = [r for r in records if r.levelname == "DEBUG"]
        message = debug[0].getMessage()
        self.assertIn("template_content=<str:5000>", message)
        self.assertIn("slide_index=2", message)
        self.assertNotIn("AAAA", message)  # never the payload itself

    def test_handles_are_truncated_in_logs(self):
        records, _, _, pid = self._run(lambda **kw: {"ok": True})
        for record in records:
            self.assertNotIn(pid, record.getMessage())
            self.assertIn(pid[:8], record.getMessage())


class TestToolConcurrency(unittest.TestCase):
    """Wrapped tools run on worker threads, bounded by
    PPT_MCP_MAX_CONCURRENT_TOOL_CALLS, instead of blocking the event loop."""

    def _register(self, fn, monkeypatch_env=None, presentation_ids=()):
        store = PresentationStore(ttl_seconds=60, max_items=10)
        for pid in presentation_ids:
            store[pid] = object()
        tools = {"demo_tool": _FakeTool(fn, read_only=True)}
        if monkeypatch_env is not None:
            import os
            old = os.environ.get("PPT_MCP_MAX_CONCURRENT_TOOL_CALLS")
            os.environ["PPT_MCP_MAX_CONCURRENT_TOOL_CALLS"] = monkeypatch_env
            try:
                serialize_per_presentation(_FakeApp(tools), store)
            finally:
                if old is None:
                    del os.environ["PPT_MCP_MAX_CONCURRENT_TOOL_CALLS"]
                else:
                    os.environ["PPT_MCP_MAX_CONCURRENT_TOOL_CALLS"] = old
        else:
            serialize_per_presentation(_FakeApp(tools), store)
        return tools["demo_tool"].fn

    def test_independent_calls_overlap_instead_of_serializing(self):
        # Two blocking (time.sleep) calls with no shared presentation_id
        # must run concurrently on separate worker threads: total wall time
        # close to one sleep, not the sum of both.
        def slow(**kw):
            time.sleep(0.15)
            return {"ok": True}

        fn = self._register(slow, monkeypatch_env="8")

        async def run_both():
            await asyncio.gather(fn(presentation_id="a"), fn(presentation_id="b"))

        started = time.monotonic()
        asyncio.run(run_both())
        elapsed = time.monotonic() - started
        self.assertLess(elapsed, 0.28)  # well under 2x0.15 if serialized

    def test_concurrency_capped_by_limiter(self):
        # With the limiter set to 1, two blocking calls must NOT overlap.
        def slow(**kw):
            time.sleep(0.1)
            return {"ok": True}

        fn = self._register(slow, monkeypatch_env="1")

        async def run_both():
            await asyncio.gather(fn(presentation_id="a"), fn(presentation_id="b"))

        started = time.monotonic()
        asyncio.run(run_both())
        elapsed = time.monotonic() - started
        self.assertGreaterEqual(elapsed, 0.2)  # serialized behind the cap

    def test_same_presentation_calls_still_serialize_despite_threading(self):
        order = []

        def worker(**kw):
            order.append(("start", kw["tag"]))
            time.sleep(0.05)
            order.append(("end", kw["tag"]))
            return {"ok": True}

        fn = self._register(worker, monkeypatch_env="8",
                            presentation_ids=["same"])

        async def run_both():
            await asyncio.gather(
                fn(presentation_id="same", tag="a"),
                fn(presentation_id="same", tag="b"),
            )

        asyncio.run(run_both())
        # Critical sections for the same deck must not interleave.
        self.assertEqual([e[0] for e in order], ["start", "end", "start", "end"])

    def test_invalid_env_value_falls_back_to_default(self):
        fn = self._register(lambda **kw: {"ok": True}, monkeypatch_env="not-a-number")
        result = asyncio.run(fn(presentation_id="a"))
        self.assertEqual(result, {"ok": True})

"""Unit tests for the concurrency-safe PresentationStore (workstream 5.3)."""
import sys
import threading
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from state import PresentationStore


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

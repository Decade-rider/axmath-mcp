"""Check the thread/lifecycle properties required for safe Office automation."""

import asyncio
import threading

import pytest

from axmath_mcp.sta import StaExecutor


def test_calls_and_lifecycle_share_one_thread():
    threads = []
    events = []

    class Backend:
        def __init__(self):
            threads.append(threading.get_ident())

        def add(self, number):
            threads.append(threading.get_ident())
            events.append(number)
            return number + 1

        def shutdown(self):
            threads.append(threading.get_ident())

    async def run():
        worker = StaExecutor(Backend, use_com=False)
        try:
            results = await asyncio.gather(*(worker.call("add", number=i) for i in range(8)))
            assert results == list(range(1, 9))
        finally:
            await worker.close()
        with pytest.raises(RuntimeError, match="shutting down"):
            await worker.call("add", number=10)

    asyncio.run(run())
    assert len(set(threads)) == 1
    assert threads[0] != threading.get_ident()
    assert events == list(range(8))


def test_cancelled_request_does_not_allow_overlapping_mutations():
    entered = threading.Event()
    release = threading.Event()
    events = []

    class Backend:
        def block(self):
            events.append("first_started")
            entered.set()
            assert release.wait(5)
            events.append("first_finished")

        def second(self):
            events.append("second_started")

        def shutdown(self):
            events.append("shutdown")

    async def run():
        worker = StaExecutor(Backend, use_com=False)
        try:
            first = asyncio.create_task(worker.call("block"))
            assert await asyncio.to_thread(entered.wait, 5)
            first.cancel()
            with pytest.raises(asyncio.CancelledError):
                await first
            second = asyncio.create_task(worker.call("second"))
            await asyncio.sleep(0)
            assert events == ["first_started"]
            release.set()
            await second
        finally:
            release.set()
            await worker.close()

    asyncio.run(run())
    assert events == ["first_started", "first_finished", "second_started", "shutdown"]


def test_request_failure_does_not_kill_worker():
    class Backend:
        def fail(self):
            raise ValueError("expected failure")

        def good(self):
            return {"ok": True}

        def shutdown(self):
            pass

    async def run():
        worker = StaExecutor(Backend, use_com=False)
        try:
            with pytest.raises(ValueError, match="expected failure"):
                await worker.call("fail")
            assert await worker.call("good") == {"ok": True}
        finally:
            await worker.close()

    asyncio.run(run())

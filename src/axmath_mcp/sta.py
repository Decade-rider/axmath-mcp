"""Serialize every Office call on one COM single-threaded apartment."""

from __future__ import annotations

import asyncio
import queue
import sys
import threading
from collections.abc import Callable
from concurrent.futures import Future
from dataclasses import dataclass
from typing import Any


def _make_backend() -> Any:
    from axmath_mcp.office import OfficeBackend

    return OfficeBackend()


@dataclass(frozen=True)
class _Request:
    method: str
    kwargs: dict[str, Any]
    result: Future[Any]


class StaExecutor:
    """Own a backend and its COM references for the whole server lifetime.

    Cancellation of an awaiting client never releases the serialization gate:
    already queued operations finish in order. No COM proxy leaves this thread;
    backend methods must return ordinary Python values. ``use_com=False`` exists
    only for testing the queue with a non-COM backend.
    """

    def __init__(self, factory: Callable[[], Any] = _make_backend, *, use_com: bool = True) -> None:
        self._factory = factory
        self._use_com = use_com
        self._queue: queue.Queue[_Request | None] = queue.Queue()
        self._ready: Future[None] = Future()
        self._stopped: Future[None] = Future()
        self._lock = threading.Lock()
        self._closing = False
        self._thread = threading.Thread(target=self._run, name="axmath-office-sta", daemon=True)
        self._thread.start()

    async def start(self) -> None:
        await asyncio.shield(asyncio.wrap_future(self._ready))

    async def call(self, method: str, **kwargs: Any) -> Any:
        await self.start()
        result: Future[Any] = Future()
        with self._lock:
            if self._closing:
                raise RuntimeError("The Office automation worker is shutting down.")
            self._queue.put(_Request(method, kwargs, result))
        return await asyncio.shield(asyncio.wrap_future(result))

    async def close(self) -> None:
        with self._lock:
            if not self._closing:
                self._closing = True
                self._queue.put(None)
        await asyncio.shield(asyncio.wrap_future(self._stopped))

    def _run(self) -> None:
        com = None
        backend = None
        shutdown_error: BaseException | None = None
        try:
            if self._use_com:
                if sys.platform != "win32":
                    raise RuntimeError("AxMath MCP requires Windows desktop Word and AxMath.")
                import pythoncom

                pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
                com = pythoncom
            backend = self._factory()
            self._ready.set_result(None)
            while True:
                try:
                    request = self._queue.get(timeout=0.05)
                except queue.Empty:
                    if com is not None:
                        com.PumpWaitingMessages()
                    continue
                if request is None:
                    break
                try:
                    value = getattr(backend, request.method)(**request.kwargs)
                except BaseException as exc:
                    request.result.set_exception(exc)
                    if not isinstance(exc, Exception):
                        raise
                else:
                    request.result.set_result(value)
                if com is not None:
                    com.PumpWaitingMessages()
        except BaseException as exc:  # noqa: BLE001 - resolve readiness even on thread failure.
            shutdown_error = exc
            if not self._ready.done():
                self._ready.set_exception(exc)
        finally:
            with self._lock:
                self._closing = True
            while True:
                try:
                    pending = self._queue.get_nowait()
                except queue.Empty:
                    break
                if pending is not None:
                    pending.result.set_exception(
                        RuntimeError("The Office automation worker has stopped.")
                    )
            try:
                if backend is not None:
                    backend.shutdown()
                backend = None
            except BaseException as exc:  # noqa: BLE001 - always uninitialize this COM apartment.
                shutdown_error = exc
            finally:
                if com is not None:
                    com.CoUninitialize()
                if shutdown_error is not None:
                    self._stopped.set_exception(shutdown_error)
                else:
                    self._stopped.set_result(None)

"""Threaded work with live progress — lifted from wfmodelingworkbench
(mofo/engine.py). In THIS app the plan-LR engine recalculates in
microseconds, so nothing here guards recalc; it exists for the slow paths —
Excel export via the generator (seconds) and optional COM recalc (tens of
seconds) — and for the OneShot/defer_to_loaded primitives the binding kit's
debouncers depend on.

Every scar in these docstrings was measured in the workbench: periodic
callbacks that strobe the busy spinner, off-thread stops that orphan tornado
timers under --num-threads, double onload registrations, progress bursts
that fire hundreds of watcher fans per drain.
"""
from __future__ import annotations

import queue
import threading
import traceback
from typing import Callable, Optional

import panel as pn
import param

_DONE = "__done__"
_FAIL = "__fail__"


def defer_to_loaded(fn: Callable) -> None:
    """Run ``fn`` immediately when the session is already attached (or
    headless), else at document load. Registering a session callback DURING
    the initial document build makes Bokeh add it twice ('A callback of the
    same type has already been added with this ID'). EVERY periodic
    registration routes through here."""
    try:
        pn.state.onload(fn)     # immediate once loaded; deferred pre-attach
    except Exception:
        fn()


def _quiet_periodic(fn: Callable, period_ms: int):
    """Register a periodic callback that never blinks the header busy
    spinner. Panel wraps EVERY tick of a normal periodic in a global busy
    event, so a 150ms drain strobes the spinner at ~7Hz for as long as it
    lives. background=True is Panel's own bookkeeping-not-user-work escape
    hatch; pn.state.add_periodic_callback doesn't forward it, so construct
    the callback directly and replicate its state._periodic bookkeeping."""
    from panel.io.callbacks import PeriodicCallback
    cb = PeriodicCallback(callback=fn, period=period_ms, background=True)
    cb.start()
    try:
        doc = pn.state.curdoc
        if doc is not None:
            pn.state._periodic.setdefault(doc, []).append(cb)
    except Exception:
        pass
    return cb


def _safe_stop(cb) -> None:
    """Stop a periodic callback from ANY thread, tolerating repeats. Under
    --num-threads Panel runs tick bodies on the executor, so a stop() issued
    from inside a tick runs without the document lock; Bokeh's
    remove_session_callback can then fail partway, leaving a tornado-level
    timer nobody holds a working handle to. Off the loop thread, schedule
    the stop onto the document's own next tick (which runs locked); callers
    in tick bodies re-invoke this until the timer is actually dead."""
    if cb is None:
        return
    doc = getattr(cb, "_doc", None)
    if doc is not None and pn.state._thread_pool is not None:
        try:
            doc.add_next_tick_callback(cb.stop)
            return
        except Exception:   # destroyed doc — its timers died with it
            pass
    try:
        cb.stop()
    except Exception:
        pass


class OneShot:
    """Run ``fn`` once, ``delay_ms`` from now, on the server loop — the
    replacement for count=1 periodics, whose count-based self-stop executes
    on the thread pool under --num-threads and can fail like any off-thread
    stop. Every tick re-attempts a safe stop until the timer is dead; a
    fired/cancelled shot can never fire again."""

    def __init__(self, fn: Callable, delay_ms: int):
        self._fn = fn
        self._fired = False
        self._cb = _quiet_periodic(self._tick, delay_ms)

    def _tick(self) -> None:
        _safe_stop(self._cb)
        if self._fired:
            return          # zombie tick — keep re-stopping, never re-fire
        self._fired = True
        self._fn()

    def cancel(self) -> None:
        self._fired = True
        _safe_stop(self._cb)


def run_async(fn: Callable, on_done: Callable,
              on_error: Optional[Callable] = None,
              period_ms: int = 150) -> None:
    """Run ``fn()`` on a worker thread; deliver its return value to
    ``on_done(result)`` (or the exception to ``on_error(exc)``) on the
    server loop. For short-but-not-instant work (workbook import, scenario
    load) where a full RunHandle is overkill but freezing the UI is rude."""
    q: "queue.Queue[tuple]" = queue.Queue()

    def _work():
        try:
            q.put((_DONE, fn()))
        except Exception as e:                          # noqa: BLE001
            q.put((_FAIL, e))

    st = {"cb": None, "done": False}

    def _drain():
        if st["done"]:              # delivered, yet still ticking: an
            _safe_stop(st["cb"])    # off-thread stop failed — retry
            return                  # until the timer is actually dead
        try:
            kind, payload = q.get_nowait()
        except queue.Empty:
            return
        st["done"] = True
        _safe_stop(st["cb"])
        if kind == _DONE:
            on_done(payload)
        elif on_error is not None:
            on_error(payload)

    def _register():
        if st["done"]:
            return
        # double onload (page rebuilt pre-attach): never orphan a ticking
        # callback — an unstoppable periodic lives forever
        _safe_stop(st["cb"])
        st["cb"] = _quiet_periodic(_drain, period_ms)
        _drain()      # the worker may have finished before registration

    defer_to_loaded(_register)
    threading.Thread(target=_work, daemon=True, name="plw-async").start()


class RunHandle(param.Parameterized):
    """Observable state of one long-running job (Excel export / COM recalc).

    Contract: ``execute(config, df=df, progress=cb, **kw)`` returns
    ``(results, messages)`` and calls ``progress(fraction, message)``
    densely from the worker."""

    running = param.Boolean(default=False)
    fraction = param.Number(default=0.0, bounds=(0.0, 1.0))
    message = param.String(default="")
    error = param.String(default="")

    def __init__(self, **params):
        super().__init__(**params)
        self._q: "queue.Queue[tuple]" = queue.Queue()
        self._cb = None
        self._on_done: Optional[Callable] = None

    def start(self, execute: Callable, config, df,
              on_done: Callable[[dict, list], None],
              period_ms: int = 150, **exec_kwargs) -> bool:
        """Single-flight: returns False if a run is already going."""
        if self.running:
            return False
        self.running, self.fraction, self.message, self.error = \
            True, 0.0, "Starting…", ""
        self._on_done = on_done

        def _progress(fraction: float, message: str) -> None:
            self._q.put(("p", min(float(fraction), 1.0), str(message)))

        def _work() -> None:
            try:
                results, messages = execute(config, df=df, progress=_progress,
                                            **exec_kwargs)
                self._q.put((_DONE, results, messages))
            except Exception as e:                      # noqa: BLE001
                self._q.put((_FAIL, f"{e}", traceback.format_exc()))

        defer_to_loaded(lambda: self._register_drain(period_ms))
        threading.Thread(target=_work, daemon=True,
                         name="plw-export-run").start()
        return True

    def _register_drain(self, period_ms: int) -> None:
        # idempotent: start() + reattach() can BOTH queue a registration at
        # onload — the second used to overwrite self._cb and orphan the
        # first, which then ticked every 150ms forever (permanent busy)
        _safe_stop(self._cb)
        self._cb = _quiet_periodic(self._drain, period_ms)
        self._drain()      # anything queued before registration

    def _drain(self) -> None:
        if not self.running and self._q.empty():
            _safe_stop(self._cb)   # outlived _finish — keep re-stopping
            return
        # progress tuples COALESCE to the newest per tick: applying every
        # queued one fired two param watchers per tuple, and a burst can
        # queue hundreds between drains
        last_p = None
        try:
            while True:
                item = self._q.get_nowait()
                if item[0] == "p":
                    last_p = item
                elif item[0] == _DONE:
                    self._finish()
                    if self._on_done is not None:
                        self._on_done(item[1], item[2])
                    return
                else:   # _FAIL
                    self.error = item[1]
                    self.message = "Run failed"
                    self._finish()
                    return
        except queue.Empty:
            pass
        if last_p is not None:
            _, self.fraction, self.message = last_p

    def reattach(self, on_done: Optional[Callable] = None,
                 period_ms: int = 150) -> None:
        """Re-home a handle that outlived its Bokeh document (the handle
        lives on session state and survives page reloads via adopt(); the
        old document's drain died with the old session, but the worker keeps
        queueing). Purges watchers from the dead page graph and, if work is
        still in flight, registers a fresh drain on the CURRENT document."""
        self._purge_watchers()
        if on_done is not None:
            self._on_done = on_done
        _safe_stop(self._cb)
        self._cb = None
        if self.running or not self._q.empty():
            defer_to_loaded(lambda: self._register_drain(period_ms))

    def _purge_watchers(self) -> None:
        try:
            for kinds in list(self.param.watchers.values()):
                for ws in list(kinds.values()):
                    for w in list(ws):
                        try:
                            self.param.unwatch(w)
                        except Exception:
                            pass
        except Exception:
            pass

    def _finish(self) -> None:
        self.running = False
        self.fraction = 1.0
        _safe_stop(self._cb)
        # keep self._cb: if the stop only lands on a later tick, the zombie
        # branch in _drain keeps re-stopping the same handle

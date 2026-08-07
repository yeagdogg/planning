"""Config <-> widget binding kit — lifted from wfmodelingworkbench
(mofo/bindings.py), whose docstrings carry the scar tissue; trimmed of its
registry-driven field renderer, which this app does not need.

THE PATTERN (every bound widget in every page):

1. SEED at construction: the widget's value is initialized from the config
   attribute (or dict key) it represents.
2. WRITE BACK on change: a watcher assigns the new value into the config
   object and bumps the scope's ConfigBus so derived displays refresh.
3. REBUILD on config replacement: loading a scenario replaces config objects
   wholesale — callers rebuild the affected widget scope.

Bindings are one-directional at runtime (widget -> config). Nothing here ever
writes config -> widget after construction, which is what makes update loops
structurally impossible. If a page needs to change config programmatically
(bulk adjust, scenario load), it mutates the config and rebuilds that scope.
"""
from __future__ import annotations

from typing import Any, Callable, Optional

import panel as pn
import param

# --------------------------------------------------------------- change bus


class ConfigBus(param.Parameterized):
    """Per-scope change signal. Widget write-backs call ``bump()``; derived
    displays are watched/bound to ``rev`` so they recompute exactly when the
    scenario changed."""

    rev = param.Integer(default=0)

    def bump(self) -> None:
        self.rev += 1


# ------------------------------------------------------------ watcher bag

class WatcherBag:
    """Watchers registered on LONG-LIVED objects (bus, run handle) by
    REBUILDABLE components leak: the page rebuilds the component, the old
    watchers stay registered, and every bump fans out to dead UI — work per
    event grows with every rebuild. Components route those registrations
    through a bag; whoever replaces the component closes the bag first.

    ``close()`` empties the bag but keeps it usable, so an in-place
    rebuild-loop (build -> close -> build ...) can reuse one bag."""

    def __init__(self) -> None:
        self._watchers: list = []
        self._closers: list = []

    def watch(self, owner, fn, parameter_names, **kwargs):
        names = ([parameter_names] if isinstance(parameter_names, str)
                 else list(parameter_names))
        w = owner.param.watch(fn, names, **kwargs)
        self._watchers.append((owner, w))
        return w

    def add_closer(self, fn) -> None:
        """Register an extra teardown (e.g. a nested bag's close)."""
        self._closers.append(fn)

    def close(self) -> None:
        watchers, self._watchers = self._watchers, []
        for owner, w in watchers:
            try:
                owner.param.unwatch(w)
            except Exception:
                pass
        closers, self._closers = self._closers, []
        for fn in closers:
            try:
                fn()
            except Exception:
                pass


# ------------------------------------------------------------- debouncing

class _DebouncedRev(param.Parameterized):
    rev = param.Integer(default=0)


def debounce(trigger, delay_ms: int = 300,
             bag: Optional[WatcherBag] = None):
    """A rev param that follows ``trigger`` (e.g. ``bus.param.rev``) but
    coalesces bursts: consumers bind to the RETURNED object's ``.param.rev``
    instead of the raw trigger, so a keystroke storm costs one recompute
    shortly after the first event instead of one per event."""
    d = _DebouncedRev()
    pending = {"on": False}

    def _fire() -> None:
        pending["on"] = False
        d.rev += 1

    def _on(_event) -> None:
        if pending["on"]:
            return
        if pn.state.curdoc is None:
            # headless (tests / scripts): no loop will ever run the
            # deferred callback — degrade to synchronous passthrough
            _fire()
            return
        pending["on"] = True

        def _register():
            # OneShot, not a count=1 periodic: Panel's count self-stop
            # runs on the executor under --num-threads and can fail,
            # leaving a zombie ticker; OneShot re-stops until dead and
            # never blinks the busy spinner while waiting
            from .engineio import OneShot
            try:
                OneShot(_fire, delay_ms)
            except Exception:      # no running loop after all
                _fire()

        # registering a session callback during the initial document
        # build double-adds it at attach — same rule as engineio.run_async
        from .engineio import defer_to_loaded
        defer_to_loaded(_register)

    if bag is not None:
        bag.watch(trigger.owner, _on, trigger.name)
    else:
        trigger.owner.param.watch(_on, trigger.name)
    return d


def coalesce(fn, delay_ms: int = 400):
    """Trailing-edge coalescer for bursty UI events (a MultiChoice fires
    once per chip added): call the returned trigger any number of times —
    ``fn`` runs ONCE, ``delay_ms`` after the burst's last call. Headless
    (no document): synchronous passthrough."""
    pending = {"shot": None}

    def _fire():
        pending["shot"] = None
        fn()

    def trigger(*_a, **_k):
        if pn.state.curdoc is None:
            _fire()
            return
        if pending["shot"] is not None:
            # cancel-then-rearm; a cancelled OneShot can never fire,
            # even if its underlying timer takes a tick to die
            pending["shot"].cancel()
            pending["shot"] = None

        def _register():
            from .engineio import OneShot
            try:
                pending["shot"] = OneShot(_fire, delay_ms)
            except Exception:
                _fire()

        from .engineio import defer_to_loaded
        defer_to_loaded(_register)

    return trigger


# ---------------------------------------------------------- option syncing

def sync_options(selector, compute_options: Callable[[], list],
                 *triggers, bus: Optional[ConfigBus] = None,
                 keep_value: bool = True,
                 bag: Optional[WatcherBag] = None) -> None:
    """Keep a selector's options fresh as upstream state changes.

    ``triggers`` are param objects (``widget.param.value``,
    ``bus.param.rev``, ...). On any trigger change the options are recomputed
    IN PLACE (no widget rebuild — focus and scroll survive); a current value
    that fell out of the new options resets to the first option (or None/[]
    for multi-selects) and a stale value can never crash a render, by
    construction."""

    def _refresh(*_events):
        computed = compute_options()
        # dict options must SURVIVE as dicts: {label: value} selectors carry
        # their real value in the dict VALUES
        if isinstance(computed, dict):
            options: Any = dict(computed)
            valid = list(computed.values())
        else:
            options = list(computed)
            valid = options
        selector.options = options
        if not keep_value:
            return
        # Read the value AFTER the options assignment, not before: a
        # bind_attr-bound widget re-seeds itself from the config on an
        # options change. Pruning the pre-assignment value would then work
        # on a stale copy.
        old = selector.value
        if isinstance(old, list):
            # "no options known" must never DESTROY a bound value: clearing
            # here would write [] into the config. Keep the value; it prunes
            # correctly the moment real options exist.
            if not valid:
                return
            kept = [v for v in old if v in valid]
            if kept != old:
                selector.value = kept
        elif old not in valid:
            selector.value = valid[0] if valid else None
        if bus is not None and selector.value != old:
            bus.bump()

    for t in triggers:
        if bag is not None:
            bag.watch(t.owner, _refresh, t.name)
        else:
            t.owner.param.watch(_refresh, t.name)
    _refresh()


# ------------------------------------------------------------ tab retitle

def retitle_tab(strip, index: int, title: str) -> None:
    """Rename one ``pn.Tabs`` pane IN PLACE without re-rendering its content.
    Panel 1.9.3 has no supported path for this — see the workbench original
    for the full autopsy. The fix swaps a FRESH TabPanel around the SAME
    child into (a) the rendered model's tabs list — the list change makes
    BokehJS rebuild headers while the content view is reused — and (b)
    Panel's per-root panel cache, so later refills carry the new title."""
    from bokeh.models import TabPanel as BkTabPanel
    strip._names[index] = title                # server-side truth
    try:
        pane = strip.objects[index]
    except IndexError:
        return
    pref = id(pane)
    for ref, (model, _parent) in list(getattr(strip, "_models", {}).items()):
        try:
            if index >= len(model.tabs):
                continue
            old = model.tabs[index]
            fresh = BkTabPanel(title=title, name=old.name, child=old.child,
                               closable=old.closable)
            tabs = list(model.tabs)
            tabs[index] = fresh
            model.tabs = tabs
            panels = strip._panels.get(ref)
            if panels is not None and pref in panels:
                panels[pref] = fresh
        except Exception:                               # noqa: BLE001
            pass


# ------------------------------------------------------------- core binders

def _option_values(widget) -> Optional[list]:
    """Valid values for an options-bearing widget, or None if the widget has
    no options concept. Dict options carry their real value in the VALUES."""
    opts = getattr(widget, "options", None)
    if opts is None:
        return None
    return list(opts.values()) if isinstance(opts, dict) else list(opts)


def _seed_widget(widget, value) -> None:
    """Seed ``widget.value`` — but NEVER seed a value the widget's current
    options can't represent. A Select handed an out-of-options value gets
    coerced client-side, and that coercion fires the write-back as if the
    user had cleared the field — silently wiping the config value."""
    valid = _option_values(widget)
    if valid is None:
        widget.value = value                         # no options concept
        return
    if isinstance(value, list):
        widget.value = [v for v in value if v in valid]
        return
    if value in valid:
        widget.value = value
    # else: leave the caller's constructed (options-valid) value untouched;
    # the config keeps its value and re-seeds correctly once the option
    # exists (see bind_attr's options re-seed).


def bind_attr(widget, obj: Any, attr: str, *,
              bus: Optional[ConfigBus] = None,
              to_config: Optional[Callable] = None,
              from_config: Optional[Callable] = None):
    """Seed ``widget.value`` from ``obj.attr`` and write every subsequent
    widget change back into ``obj.attr``. Returns the widget (chainable)."""
    cur = getattr(obj, attr)
    _seed_widget(widget, from_config(cur) if from_config else cur)

    def _write_back(event):
        value = to_config(event.new) if to_config else event.new
        old = getattr(obj, attr)
        setattr(obj, attr, value)
        # bump only on a REAL config change: to_config transforms can map
        # distinct widget states to the same config value, and every bump
        # fans out to the whole page
        if bus is not None and not _same(old, value):
            bus.bump()

    widget.param.watch(_write_back, "value")

    if "options" in widget.param:
        def _reseed_on_options(_event):
            """Re-apply the CONFIG value whenever the options change — a
            config adopted before options exist would otherwise render an
            empty widget over a populated config (the model runs on values
            the user cannot see). Idempotent after a real user edit."""
            want = getattr(obj, attr)
            _seed_widget(widget, from_config(want) if from_config else want)

        widget.param.watch(_reseed_on_options, "options")
    return widget


def _same(a: Any, b: Any) -> bool:
    try:
        return bool(a == b)
    except Exception:
        return a is b


def bind_key(widget, mapping: dict, key: str, *,
             default: Any = None,
             bus: Optional[ConfigBus] = None,
             to_config: Optional[Callable] = None,
             from_config: Optional[Callable] = None):
    """dict-flavored ``bind_attr``. Seeds from ``mapping.get(key, default)``;
    write-back assigns ``mapping[key]``."""
    cur = mapping.get(key, default)
    widget.value = from_config(cur) if from_config else cur

    def _write_back(event):
        value = to_config(event.new) if to_config else event.new
        old = mapping.get(key, default)
        mapping[key] = value
        if bus is not None and not _same(old, value):
            bus.bump()

    widget.param.watch(_write_back, "value")
    return widget

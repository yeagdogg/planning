"""Per-browser-session state — lifted from wfmodelingworkbench
(mofo/session.py) and reduced to this app's one-page-family shape.

One PlanSession is constructed per Panel session (browser tab) inside
main.py's page factory; every page component receives it explicitly. The app
is single-route with internal navigation, so the object graph survives page
switches; reload survival across F5/theme-toggle rides the ?sid= store in
main.py plus ``adopt`` here.

The config (a Scenario, defined in app/state.py) remains the single source
of truth; results hold live engine outputs keyed by combo.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Optional

import param

from .bindings import ConfigBus


class PageContext(param.Parameterized):
    """Reactive facts many components depend on. Components bind to these
    (or hand them to sync_options) instead of coupling to whichever widget
    happens to change them:

    - data_rev bumps when the SCENARIO IS REPLACED (import / load / new) —
      widget scopes rebuild on this
    - results_rev bumps when engine results refresh (bus-driven recompute)
    - combos tracks the scenario's (bu, state) roster for selectors
    """
    data_rev = param.Integer(default=0)
    results_rev = param.Integer(default=0)
    combos = param.List(default=[])
    # one-shot combo adoption: the Book page's click-through sets this to a
    # 'BU|State' key just before activating a line; the Combo page consumes
    # it on the resulting data_rev and clears it (plain attr, not a param —
    # nothing should ever bind to it)
    focus = None


@dataclass
class PageState:
    """The app's mutable state."""
    config: Any = None                # the ACTIVE Scenario (app/state.py)
    # P4: every loaded line, keyed by LOB name. ``config`` is always one of
    # these objects (identity, not a copy) — Inputs/Combo edit the active
    # line IN the book, so the Book page sees edits with no sync step.
    book: Dict[str, Any] = field(default_factory=dict)
    data_version: int = 0
    results: Dict[str, Any] = field(default_factory=dict)  # combo key -> EngineResult
    caches: Dict[str, Any] = field(default_factory=dict)
    # the export RunHandle — lives HERE (not on a page object) so a reload
    # mid-export keeps the run: adopt() carries it and the new page re-homes
    # it via RunHandle.reattach()
    run_handle: Optional[Any] = None


def _run_in_flight(handle) -> bool:
    """True if a RunHandle has work actively running or results queued —
    the only case where sharing it across an adopt matters. An idle handle
    is not shared, so a duplicated tab takes a fresh one instead of
    reattach-purging the live original's watchers."""
    if handle is None:
        return False
    try:
        if getattr(handle, "running", False):
            return True
        q = getattr(handle, "_q", None)
        return bool(q is not None and not q.empty())
    except Exception:
        return True        # unknown -> err toward preserving the run


class PlanSession:
    """Everything one user session holds."""

    def __init__(self) -> None:
        self.page = PageState()
        self.bus = ConfigBus()       # scenario-input change signal
        self.ctx = PageContext()
        # set by main.build(): navigate the SPA to a page by nav label.
        # Deliberately NOT carried by adopt() — it belongs to the live page
        # graph, not the state.
        self.goto = None

    # ---- reload survival ---------------------------------------------------
    def adopt(self, prev: "PlanSession", *, isolate: bool = False) -> None:
        """Carry STATE (scenario, results, caches) from the previous
        page-load's session into this fresh one. The reactive objects (bus,
        ctx) stay NEW so nothing from the dead page graph's watchers
        survives; callable cache entries are dropped — their owners
        re-register on build.

        ``isolate=True`` is the claimed-tab path (true duplicate OR a fast
        reload whose old socket hasn't torn down — indistinguishable at
        attach). Deep-copy the scenario so a real duplicate's edits stop
        clobbering the original; fork the results/caches dicts (values are
        read-only engine outputs, shared not duplicated). The run handle is
        shared ONLY when work is actually in flight."""
        import copy
        src, dst = prev.page, self.page
        # the book carries WITH the active config's identity intact: after a
        # deepcopy the active scenario must be the copy IN the copied book,
        # or Inputs would edit an object the Book page no longer shows
        dst.book = copy.deepcopy(src.book) if isolate else src.book
        if src.config is None:
            dst.config = None
        elif isolate:
            dst.config = dst.book.get(getattr(src.config, "lob", None))
            if dst.config is None:
                dst.config = copy.deepcopy(src.config)
        else:
            dst.config = src.config
        dst.data_version = src.data_version
        dst.results = dict(src.results) if isolate else src.results
        dst.caches = {k: v for k, v in src.caches.items() if not callable(v)}
        if isolate and not _run_in_flight(src.run_handle):
            dst.run_handle = None
        else:
            dst.run_handle = src.run_handle
        if dst.config is not None:
            try:
                self.ctx.combos = list(dst.config.combo_keys())
            except Exception:
                pass

    # ---- config replacement (scenario load / import / new) -----------------
    def replace_config(self, config: Any) -> None:
        """Load a scenario: it joins the book under its line's key AND
        becomes the active one. Replacing the object invalidates every
        widget bound to the old one, so pages register a rebuild hook on
        ctx.data_rev; results are stale by definition."""
        if config is not None and getattr(config, "lob", None):
            self.page.book[config.lob] = config
        self._swap(config)

    def activate(self, lob: str) -> bool:
        """Focus an already-loaded line: Inputs/Combo work this scenario;
        the book keeps every line. Always swaps (even to the current line)
        so a click-through re-renders predictably."""
        scn = self.page.book.get(lob)
        if scn is None:
            return False
        self._swap(scn)
        return True

    def _swap(self, config: Any) -> None:
        self.page.config = config
        self.page.results = {}
        self.page.data_version += 1
        try:
            self.ctx.combos = (list(config.combo_keys())
                               if config is not None else [])
        except Exception:
            self.ctx.combos = []
        self.ctx.data_rev += 1
        self.bus.bump()

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


@dataclass
class PageState:
    """The app's mutable state."""
    config: Any = None                # the Scenario (app/state.py)
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
        dst.config = copy.deepcopy(src.config) if isolate else src.config
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
        """Swap the scenario wholesale. Replacing the object invalidates
        every widget bound to the old one, so pages register a rebuild hook
        on ctx.data_rev; results are stale by definition."""
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

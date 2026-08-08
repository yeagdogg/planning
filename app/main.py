"""Planning Workbench — Panel entry point.

    app\\serve.bat [port]        (default 8602)

ONE route, internal navigation (Inputs | Combo | Book | Scenarios | Guide):
in Panel each URL is its own session, so a single route keeps one live
object graph per browser tab. Architecture lifted from wfmodelingworkbench
(app.py there), whose docstrings carry the measurements behind every
decision; trimmed of what this app doesn't need (AI panels, chip sentinel,
busy-debug hooks).
"""
from __future__ import annotations

import os
import sys

# repo root on sys.path: the app imports src.engine (the calc core) and
# app.* by package path; nothing is pip-installed (no stale-shadow risk)
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

# ---- compiled-DLL preflight (Windows AppControl / WDAC) -------------------
# The numeric stack loads native DLLs on first import, and a corporate
# Application Control policy can block a DLL on FIRST TOUCH. When that
# happens inside a session build, Bokeh's script handler CACHES the failure
# and every later session serves an EMPTY page until the server restarts.
# Import the stack here, ONCE, at boot: a block surfaces immediately,
# retries once (first-touch blocks usually clear after the policy scan),
# and a persistent block stops the server AT STARTUP with a message.


def _preflight_compiled_stack() -> None:
    import importlib
    import time as _t
    for mod in ("numpy", "pandas", "bokeh", "holoviews"):
        for attempt in (1, 2):
            try:
                importlib.import_module(mod)
                break
            except ImportError as e:
                if attempt == 1 and "Application Control" in str(e):
                    _t.sleep(2.0)       # first-touch scan usually clears
                    continue
                if "Application Control" in str(e):
                    print("\n" + "=" * 72
                          + f"\nAppControl blocked a compiled DLL in "
                          f"'{mod}' twice:\n  {e}\nThe server is stopping "
                          "NOW instead of serving blank pages. Re-run "
                          "serve.bat — the DLL almost always loads once "
                          "the policy has scanned it.\n" + "=" * 72 + "\n",
                          flush=True)
                    raise
                raise               # unrelated import error — fail loud


_preflight_compiled_stack()

import traceback  # noqa: E402

import panel as pn  # noqa: E402

_last_err = {"msg": "", "t": 0.0}


def _surface_error(exc: BaseException) -> None:
    """Global callback-exception handler: without it, an exception in a
    widget watcher silently kills that event server-side — the widget just
    'stops responding' with the only evidence in a console nobody watches.
    Print the full traceback AND toast it."""
    import time as _t
    traceback.print_exception(exc)
    msg = f"{type(exc).__name__}: {exc}"
    if msg == _last_err["msg"] and _t.monotonic() - _last_err["t"] < 5:
        return                                  # storm guard
    _last_err["msg"], _last_err["t"] = msg, _t.monotonic()
    try:
        if pn.state.notifications is not None:
            pn.state.notifications.error(
                f"Internal error — a control may stop responding until "
                f"reload: {msg}", duration=0)
    except Exception:                           # noqa: BLE001
        pass


# disconnect_notification: a tab that outlives its websocket (server
# restart, laptop sleep) otherwise LOOKS alive while silently ignoring
# every click. The watcher JS below sees this banner and reloads
# automatically; the ?sid= session store brings the state back.
pn.extension("tabulator", "floatpanel", "filedropper", notifications=True,
             throttled=True,
             exception_handler=_surface_error,
             disconnect_notification=(
                 "Connection lost — reloading automatically to restore "
                 "your session…"))

import holoviews as hv  # noqa: E402

hv.extension("bokeh")

from collections import OrderedDict  # noqa: E402

from app.glue.session import PlanSession  # noqa: E402
from app.glue.theme import ACCENT, brand_kwargs, brand_raw_css  # noqa: E402

for _brand_css in brand_raw_css():
    if _brand_css not in pn.config.raw_css:
        pn.config.raw_css.append(_brand_css)

APP_TITLE = "Westfield Planning Workbench"


def _placeholder(title: str, blurb: str):
    """P0 stub page — replaced wave by wave with real modules."""
    def build(session):
        return {
            "main": pn.Column(
                pn.pane.Markdown(f"## {title}\n\n{blurb}"),
                sizing_mode="stretch_both"),
            "sidebar": pn.Column(sizing_mode="stretch_width"),
        }
    return build


from app.pages import book as _book_page  # noqa: E402
from app.pages import combo as _combo_page  # noqa: E402
from app.pages import inputs as _inputs_page  # noqa: E402

_PAGES = {
    "✏️ Inputs": _inputs_page.build,
    "🔬 Combo": _combo_page.build,
    "📚 Book": _book_page.build,
    "🗂 Scenarios": _placeholder(
        "Scenarios", "Named save / load / diff and the Excel export (wave P5)."),
    "📖 Guide": _placeholder(
        "Guide", "How the app relates to the workbook and the engine (wave P6)."),
}
# nav VALUES are URL-safe slugs so the current page lives in ?page=… —
# reloads land back where the user was, and pages become bookmarkable
_PAGE_SLUGS = {"✏️ Inputs": "inputs", "🔬 Combo": "combo", "📚 Book": "book",
               "🗂 Scenarios": "scenarios", "📖 Guide": "guide"}
_SLUG_BUILDERS = {slug: _PAGES[label] for label, slug in _PAGE_SLUGS.items()}

# ---- sessions that survive reloads ----------------------------------------
# A reload (the template's theme toggle NAVIGATES and drops query params,
# F5, reconnect after a server hiccup) normally starts a blank session. Key
# each browser TAB with a sessionStorage id, re-attach via a ?sid= redirect,
# and ADOPT the previous session's state into the fresh one. New tab = new
# session, by design. (Full autopsy in the workbench original.)
_STORE_KEY = "plw_sessions"
_MAX_KEPT = 4      # soft cap on DEAD sessions only; live tabs never evicted

_REDIRECT_JS = """
var sid = sessionStorage.getItem('plw_sid');
if (!sid) {
    sid = Math.random().toString(36).slice(2, 12);
    sessionStorage.setItem('plw_sid', sid);
}
var q = new URLSearchParams(window.location.search);
q.set('sid', sid);
window.location.replace(window.location.pathname + '?' + q.toString());
"""

# duplicating a tab COPIES sessionStorage: the duplicate presents the
# original's sid while the original is live. But "still live" needs a grace
# period — a Ctrl+R against a busy server can arrive before the old
# socket's disconnect is processed. Retry the SAME sid twice first.
_REMINT_JS = """
var n = parseInt(sessionStorage.getItem('plw_sid_retry') || '0', 10);
if (n < 2) {
    sessionStorage.setItem('plw_sid_retry', String(n + 1));
    setTimeout(function () { window.location.reload(); }, 1400);
} else {
    sessionStorage.removeItem('plw_sid_retry');
    var sid = Math.random().toString(36).slice(2, 12);
    sessionStorage.setItem('plw_sid', sid);
    var q = new URLSearchParams(window.location.search);
    q.set('sid', sid);
    window.location.replace(window.location.pathname + '?' + q.toString());
}
"""

_CLEAR_RETRY_JS = """
sessionStorage.removeItem('plw_sid_retry');
"""

# Watch for the disconnect banner and reload ONCE after a short pause —
# the session store restores state, so the user never Ctrl+Rs a dead tab.
_RECONNECT_JS = """
if (!window._plwReconnect) {
    window._plwReconnect = true;
    var fired = false;
    new MutationObserver(function () {
        if (fired) { return; }
        if ((document.body.innerText || '').indexOf(
                'Connection lost') !== -1) {
            fired = true;
            setTimeout(function () { window.location.reload(); }, 2000);
        }
    }).observe(document.body, {childList: true, subtree: true});
}
"""

# Liveness heartbeat — the banner only covers CLEAN socket closes. A
# half-open websocket (laptop sleep, proxy idle-kill) shows NO banner: the
# page looks alive while every event goes nowhere. The server pulses a
# hidden counter every 15s; this watchdog reloads when 60s+ pass with no
# pulse while the tab is visible. Reload storms are capped at 2, reset by
# a minute of healthy pulses.
_HEARTBEAT_JS = """
window._plwBeat = Date.now();
if (!window._plwHbWatch) {
    window._plwHbWatch = true;
    var check = function () {
        if (document.visibilityState !== 'visible') { return; }
        var quiet = Date.now() - window._plwBeat;
        if (quiet > 60000) {
            var n = parseInt(sessionStorage.getItem('plw_hb_n') || '0', 10);
            if (n >= 2) { return; }
            sessionStorage.setItem('plw_hb_n', String(n + 1));
            window.location.reload();
        } else if (quiet < 20000) {
            var since = parseInt(
                sessionStorage.getItem('plw_hb_ok') || '0', 10);
            if (!since) {
                sessionStorage.setItem('plw_hb_ok', String(Date.now()));
            } else if (Date.now() - since > 60000) {
                sessionStorage.removeItem('plw_hb_n');
            }
        }
    };
    setInterval(check, 10000);
    document.addEventListener('visibilitychange', check);
}
"""

_BEAT_SYNC_JS = """
window._plwBeat = Date.now();
"""

# Browser-side errors are invisible to the server: a JS exception can kill
# a component's listeners while the websocket stays healthy. Bridge
# window.onerror / unhandledrejection into a hidden widget the server
# watches. ResizeObserver-loop noise is benign and filtered server-side.
_ERR_BRIDGE_JS = """
if (!window._plwErrHook) {
    window._plwErrHook = true;
    var send = function (msg) {
        try {
            errlog.value = String(msg).slice(0, 500) + ' @' + Date.now();
        } catch (e) {}
    };
    window.addEventListener('error', function (e) {
        send((e.message || 'script error') + ' [' + (e.filename || '?')
             + ':' + (e.lineno || '?') + ']');
    });
    window.addEventListener('unhandledrejection', function (e) {
        send('unhandledrejection: ' + (e.reason && e.reason.message
             ? e.reason.message : e.reason));
    });
}
"""


def _session_id():
    args = pn.state.session_args or {}
    vals = args.get("sid")
    if not vals:
        return None
    try:
        return vals[0].decode("utf-8", "ignore") or None
    except Exception:
        return None


def _redirect_stub(js: str = _REDIRECT_JS,
                   message: str = "*Restoring your session…*"
                   ) -> pn.template.FastListTemplate:
    """Tiny page that mints/reads the per-tab id and reloads with ?sid=."""
    kick = pn.widgets.Toggle(visible=False)
    kick.jscallback(value=js)
    pn.state.onload(lambda: setattr(kick, "value", True))
    return pn.template.FastListTemplate(
        title=APP_TITLE, accent_base_color=ACCENT, **brand_kwargs(),
        main=[pn.Column(pn.pane.Markdown(message), kick)])


def _claimed_by_live_tab(prev) -> bool:
    """True when the stored session's document still has a CONNECTED
    websocket — the tab presenting this sid is a DUPLICATE of a live tab.
    A reload's old socket is already closed, so adoption keeps working."""
    try:
        doc = getattr(prev, "_doc", None)
        session = getattr(getattr(doc, "session_context", None),
                          "session", None)
        return bool(session is not None
                    and session.connection_count > 0)
    except Exception:
        return False


def _has_stuff(session) -> bool:
    try:
        return bool(session.page.config is not None or session.page.results)
    except Exception:
        return False


def _evict_over_cap(store, keep_sid, is_live=_claimed_by_live_tab) -> None:
    """Trim the store to ``_MAX_KEPT`` — but only ever evict a session
    whose tab is NO LONGER live (the numeric cap is soft for live tabs)."""
    while len(store) > _MAX_KEPT:
        victim = next((s for s in store
                       if s != keep_sid and not is_live(store[s])), None)
        if victim is None:
            break                      # all remaining are live — keep them
        store.pop(victim, None)


def _attach_session(sid: str):
    """Adopt (or create) the session for this sid; None = sid collision (a
    duplicated tab) — the caller serves the re-mint stub instead. A session
    HOLDING REAL STATE is never abandoned to the remint path: a slow
    teardown can outlive the stub's retry budget, and reminting would throw
    away the scenario. A stateful 'claimed' sid is adopted with isolation
    plus a duplicate-tab warning."""
    store = pn.state.cache.setdefault(_STORE_KEY, OrderedDict())
    prev = store.get(sid)
    claimed = False
    if prev is not None and _claimed_by_live_tab(prev):
        if not _has_stuff(prev):
            return None            # nothing to lose — remint as before
        claimed = True
        prev_ref = prev

        def _warn_if_still_claimed():
            if _claimed_by_live_tab(prev_ref):
                pn.state.notifications.warning(
                    "This session looks open in another tab — edits from "
                    "two tabs write over each other.", duration=8000)

        def _schedule_dup_check():
            # deferred ~4s: a reload's old socket reads as connected for a
            # second or two; a true twin holds its connection indefinitely
            from app.glue.engineio import OneShot
            OneShot(_warn_if_still_claimed, 4000)

        pn.state.onload(_schedule_dup_check)
    store.pop(sid, None)
    session = PlanSession()
    if prev is not None:
        session.adopt(prev, isolate=claimed)
        if _has_stuff(session) and not claimed:
            pn.state.onload(lambda: pn.state.notifications.info(
                "Session restored — your scenario and results survived "
                "the reload.", duration=4000))
    try:
        session._doc = pn.state.curdoc     # liveness anchor
    except Exception:
        session._doc = None
    store[sid] = session
    _evict_over_cap(store, keep_sid=sid)
    return session


def build() -> pn.template.FastListTemplate:
    sid = _session_id()
    if sid is None:
        return _redirect_stub()
    session = _attach_session(sid)
    if session is None:     # claimed sid: retry (busy-server reload),
        return _redirect_stub(  # then remint (true duplicate tab)
            _REMINT_JS, "*Reattaching to your session…*")

    # browser-side error bridge: JS exceptions land in this hidden widget;
    # the server logs and toasts, so a client-side death names itself
    errlog = pn.widgets.TextInput(value="", visible=False)
    _seen_ro = {"on": False}

    def _client_error(event):
        if not event.new:
            return
        msg = event.new.rsplit(" @", 1)[0]
        if "ResizeObserver" in msg:     # benign browser noise
            if not _seen_ro["on"]:
                _seen_ro["on"] = True
                print(f"[plw client-error] {msg} (benign — further "
                      "occurrences suppressed)", flush=True)
            return
        print(f"[plw client-error] {msg}", flush=True)
        try:
            pn.state.notifications.error(
                "Browser-side error — a control may stop responding "
                f"until reload: {msg}", duration=0)
        except Exception:                               # noqa: BLE001
            pass

    errlog.param.watch(_client_error, "value")

    watchdog = pn.widgets.Toggle(visible=False)
    watchdog.jscallback(value=_RECONNECT_JS + _CLEAR_RETRY_JS
                        + _HEARTBEAT_JS + _ERR_BRIDGE_JS,
                        args={"errlog": errlog})
    pn.state.onload(lambda: setattr(watchdog, "value", True))

    # server side of the liveness heartbeat: a quiet background periodic
    # pulses a hidden counter every 15s; a clientless session stops
    # pulsing after 3 unheard beats instead of ticking into the void
    beat = pn.widgets.IntInput(value=0, visible=False)
    beat.jscallback(value=_BEAT_SYNC_JS)
    _hb = {"dead": 0, "cb": None}

    def _pulse():
        from app.glue.engineio import _safe_stop
        try:
            sess = pn.state.curdoc.session_context.session
            alive = sess is None or sess.connection_count > 0
        except Exception:                               # noqa: BLE001
            alive = False       # torn-down doc = definitely clientless
        if not alive:
            _hb["dead"] += 1
            if _hb["dead"] >= 3:
                _safe_stop(_hb["cb"])
            return
        _hb["dead"] = 0
        beat.value += 1

    def _register_beat():
        from app.glue.engineio import _quiet_periodic
        _hb["cb"] = _quiet_periodic(_pulse, 15000)

    pn.state.onload(_register_beat)

    nav = pn.widgets.RadioButtonGroup(
        options=dict(_PAGE_SLUGS), value="inputs",
        button_type="primary", orientation="vertical",
        sizing_mode="stretch_width")
    session.goto = lambda name: setattr(nav, "value",
                                        _PAGE_SLUGS.get(name, name))

    # Pages are built LAZILY on first visit and STAY MOUNTED afterwards —
    # switching toggles `visible` instead of swapping holder.objects (an
    # objects swap makes the browser tear down and re-render every Bokeh
    # view; measured ~2.8s on a heavy page in the workbench).
    built: dict = {}
    main_holder = pn.Column(sizing_mode="stretch_both")
    side_holder = pn.Column(sizing_mode="stretch_width")

    def _switch(slug: str) -> None:
        if slug not in _SLUG_BUILDERS:
            return
        fresh = slug not in built
        if fresh:
            built[slug] = _SLUG_BUILDERS[slug](session)
            main_holder.append(built[slug]["main"])
            side_holder.append(built[slug]["sidebar"])
        for s, parts in built.items():
            show = s == slug
            if parts["main"].visible != show:
                parts["main"].visible = show
            if parts["sidebar"].visible != show:
                parts["sidebar"].visible = show
        # Tabulator's virtual renderer drops its rows while a page is hidden
        # and does not redraw on the visibility flip alone — pages that hold
        # grids expose on_show (a value rewrite) and it runs AFTER re-show.
        # A fresh build just rendered; only re-shows need the nudge.
        if not fresh:
            hook = built[slug].get("on_show")
            if callable(hook):
                try:
                    hook()
                except Exception:                       # noqa: BLE001
                    pass

    nav.param.watch(lambda e: _switch(e.new), "value")
    # the CURRENT PAGE lives in the URL: reloads restore it, pages are
    # linkable. The theme toggle still drops query params by design.
    try:
        if pn.state.location is not None:
            pn.state.location.sync(nav, {"value": "page"})
    except Exception:
        pass
    _switch(nav.value)

    template = pn.template.FastListTemplate(
        title=APP_TITLE,
        accent_base_color=ACCENT,
        **brand_kwargs(),
        sidebar=[nav, pn.layout.Divider(), side_holder, watchdog,
                 beat, errlog],
        main=[main_holder],
        sidebar_width=340,
    )
    return template


build().servable()

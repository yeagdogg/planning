"""Import walls for the app (mirrors wfmodelingworkbench's test_boundary).

Two directions, two interpreters:
- The ENGINE must stay UI-free: importing src.engine must not pull panel /
  bokeh / holoviews into the process. Runs everywhere (subprocess probe).
- The APP glue must import cleanly headless. Runs only where panel is
  installed (the app venv); the system interpreter — the one the Excel
  release pipeline uses — skips it, which is itself the non-interference
  property: the pipeline never needs the app's dependencies.
"""

import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]


def test_engine_is_ui_free():
    probe = (
        "import sys; sys.path.insert(0, r'%s')\n"
        "import src.engine\n"
        "bad = sorted({m.split('.')[0] for m in sys.modules} & "
        "{'panel', 'bokeh', 'holoviews', 'hvplot', 'param'})\n"
        "assert not bad, f'engine imported UI stack: {bad}'\n" % ROOT
    )
    r = subprocess.run([sys.executable, "-c", probe],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr


def test_app_never_imported_by_pipeline_modules():
    """No src/ or tools/ module may import app.* — the pipeline must build
    and verify identically whether app/ exists or not."""
    offenders = []
    for folder in ("src", "tools"):
        for py in (ROOT / folder).glob("*.py"):
            text = py.read_text(encoding="utf-8", errors="ignore")
            if "from app" in text or "import app." in text:
                offenders.append(str(py))
    assert not offenders, offenders


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_glue_imports_headless():
    sys.path.insert(0, str(ROOT))
    from app.glue import bindings, engineio, format as fmt, session, theme
    b = bindings.ConfigBus()
    b.bump()
    assert b.rev == 1
    bag = bindings.WatcherBag()
    bag.close()                      # empty close is safe and reusable
    s = session.PlanSession()
    s2 = session.PlanSession()
    s.page.caches["x"] = 1
    s.page.caches["fn"] = lambda: None
    s2.adopt(s)
    assert s2.page.caches == {"x": 1}        # callables dropped
    assert fmt.fmt_pct(0.634) == "63.4%"
    assert fmt.fmt_mod(1.0523) == "1.052"
    assert fmt.fmt_idx(0.95966) == "0.9597"
    assert theme.ACCENT.startswith("#")
    assert hasattr(engineio, "RunHandle")


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_session_isolate_deep_copies_config():
    sys.path.insert(0, str(ROOT))
    from app.glue.session import PlanSession
    a = PlanSession()
    a.page.config = {"rows": [1, 2]}          # any mutable stand-in
    b = PlanSession()
    b.adopt(a, isolate=True)
    b.page.config["rows"].append(3)
    assert a.page.config["rows"] == [1, 2]    # original untouched
    c = PlanSession()
    c.adopt(a)                                # plain reload: shared
    assert c.page.config is a.page.config

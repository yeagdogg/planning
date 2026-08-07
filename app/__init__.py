"""Planning Workbench — the Panel front end for the plan-LR engine.

Strictly additive to the repo (see the D-series entry when it ships): the
app imports src/engine.py and friends read-only, runs in its own pinned
venv (app/.venv, requirements-app.txt), and the Excel pipeline neither
imports nor depends on anything under app/.
"""

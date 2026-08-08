"""P3 deep-dive visuals for the Combo page.

Split the way the numbers split: ``bridge`` (the factor chain, tabulated and
as waterfalls), ``earning`` (the parallelogram and the earn-in lag),
``delivery`` (the monthly LR walk and the ec(k,P) band).

Every module keeps its FRAME functions pure (pandas + the engine's own
methods) so the numbers test on both interpreters; only the chart builders
touch holoviews/bokeh, and they import them lazily through ``ensure_hv``.
"""
from __future__ import annotations


def ensure_hv():
    """Import holoviews with its bokeh backend registered (idempotent).

    ``hv.extension`` is notebook-flavored; registering the plotting backend
    directly is the quiet server-side equivalent the workbench uses.
    """
    import holoviews as hv
    if "bokeh" not in hv.Store.renderers:
        import holoviews.plotting.bokeh  # noqa: F401 — registers on import
    return hv

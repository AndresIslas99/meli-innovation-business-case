"""
Caso 3 — Picking robótico productivity simulator.

Sliders: OWP, OWS, R2R, PPF.
Output: T_rack, racks/hr, items/hr/estación + Δ vs baseline + tornado de palancas.
"""

import json
from pyscript import document
from pyscript.ffi import create_proxy
from js import Plotly, JSON


def _js(obj):
    return JSON.parse(json.dumps(obj))


BASELINE = dict(owp=7.0, ows=6.0, r2r=8.5, ppf=1.8)
BASELINE_ITEMS_HR = (3600 / (BASELINE["owp"] + BASELINE["ows"] + BASELINE["r2r"])) * BASELINE["ppf"]


def _val(elem_id: str) -> float:
    el = document.querySelector(f"#{elem_id}")
    return float(el.value) if el else 0.0


def _set(elem_id: str, value: str, color: str = ""):
    el = document.querySelector(f"#{elem_id}")
    if el:
        el.textContent = value
        if color:
            el.style.color = color


def compute(owp, ows, r2r, ppf):
    t_rack = owp + ows + r2r
    racks_hr = 3600 / t_rack
    items_hr = racks_hr * ppf
    return dict(t_rack=t_rack, racks_hr=racks_hr, items_hr=items_hr)


def render(*_args):
    owp = _val("c3-owp")
    ows = _val("c3-ows")
    r2r = _val("c3-r2r")
    ppf = _val("c3-ppf")

    _set("c3-owp-val", f"{owp:.1f}")
    _set("c3-ows-val", f"{ows:.1f}")
    _set("c3-r2r-val", f"{r2r:.1f}")
    _set("c3-ppf-val", f"{ppf:.1f}")

    r = compute(owp, ows, r2r, ppf)
    delta_pct = (r["items_hr"] - BASELINE_ITEMS_HR) / BASELINE_ITEMS_HR * 100

    _set("c3-out-track", f"{r['t_rack']:.1f}")
    _set("c3-out-racks", f"{r['racks_hr']:.1f}")
    _set("c3-out-items", f"{r['items_hr']:.1f}")

    delta_color = "#00A650" if delta_pct >= 0 else "#E63946"
    delta_sign = "+" if delta_pct >= 0 else ""
    _set("c3-out-delta", f"{delta_sign}{delta_pct:.1f}%", delta_color)

    # Tornado — sensibilidad de items/hr a cada palanca (variando ±25% individualmente)
    levers = [
        ("OWP", "owp", owp),
        ("OWS", "ows", ows),
        ("R2R", "r2r", r2r),
        ("PPF", "ppf", ppf),
    ]
    tornado_x = []
    tornado_y = []
    tornado_colors = []
    for label, key, base in levers:
        # ±25%
        kwargs_low = dict(owp=owp, ows=ows, r2r=r2r, ppf=ppf)
        kwargs_high = dict(owp=owp, ows=ows, r2r=r2r, ppf=ppf)
        kwargs_low[key] = max(0.1, base * 0.75)
        kwargs_high[key] = base * 1.25
        items_low = compute(**kwargs_low)["items_hr"]
        items_high = compute(**kwargs_high)["items_hr"]

        # Para OWP/OWS/R2R: bajarlos sube items_hr → la "mejora" es la versión low
        # Para PPF: subirlo sube items_hr → la "mejora" es la versión high
        if key == "ppf":
            improvement = items_high - r["items_hr"]
            worsening = items_low - r["items_hr"]
        else:
            improvement = items_low - r["items_hr"]
            worsening = items_high - r["items_hr"]

        tornado_x.append(improvement)
        tornado_y.append(label)
        tornado_colors.append("#00A650")

        tornado_x.append(worsening)
        tornado_y.append(label)
        tornado_colors.append("#E63946")

    tornado = {
        "x": tornado_x,
        "y": tornado_y,
        "type": "bar",
        "orientation": "h",
        "marker": {"color": tornado_colors, "line": {"color": "#2D3277", "width": 1}},
        "text": [f"{x:+.1f}" for x in tornado_x],
        "textposition": "auto",
        "hovertemplate": "<b>%{y}</b><br>Δ items/hr: %{x:+.2f}<extra></extra>",
    }
    layout = {
        "title": {
            "text": f"<b>Tornado · sensibilidad ±25% por palanca (items/hr)</b><br>"
                    f"<sub>Baseline current: {r['items_hr']:.1f} items/hr/estación</sub>",
            "font": {"family": "Inter, sans-serif", "size": 14, "color": "#2D3277"},
        },
        "margin": {"l": 50, "r": 30, "t": 70, "b": 50},
        "paper_bgcolor": "#F7F7FB",
        "plot_bgcolor": "#FFFFFF",
        "font": {"family": "Inter, sans-serif", "color": "#1A1A2E"},
        "xaxis": {"title": "Δ items/hr (verde = mejora, rojo = empeora)", "zeroline": True, "zerolinewidth": 2, "zerolinecolor": "#2D3277"},
        "yaxis": {"categoryorder": "array", "categoryarray": ["PPF", "R2R", "OWS", "OWP"]},
        "barmode": "overlay",
        "showlegend": False,
        "height": 320,
    }
    config = {
        "displayModeBar": True,
        "modeBarButtonsToRemove": ["select2d", "lasso2d", "toggleSpikelines", "hoverClosestCartesian", "hoverCompareCartesian", "sendDataToCloud"],
        "displaylogo": False,
        "doubleClick": "reset",
        "responsive": True,
        "toImageButtonOptions": {"format": "png", "filename": "meli-picking-productivity"},
    }
    Plotly.react("c3-chart", _js([tornado]), _js(layout), _js(config))


_proxy = create_proxy(render)
for sid in ["c3-owp", "c3-ows", "c3-r2r", "c3-ppf"]:
    el = document.querySelector(f"#{sid}")
    if el:
        el.addEventListener("input", _proxy)

render()

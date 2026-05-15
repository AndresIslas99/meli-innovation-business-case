"""
Caso 2 — Shelves-to-Person bottleneck simulator.

Permite al evaluador mover utilización de Put Away, eficiencia de Picking,
horas/día y días/sem para ver cómo se mueve el bottleneck y si se alcanza el target.
"""

import json
from pyscript import document
from pyscript.ffi import create_proxy
from js import Plotly, JSON


def _js(obj):
    """Convert nested Python dict/list to a plain JS object.
    Plotly via Pyodide rejects Map-style dicts (AttributeError xaxis),
    so we round-trip through JSON.parse to get plain JS objects/arrays.
    """
    return JSON.parse(json.dumps(obj))


SEM_PER_MES = 4.3

# Parámetros fijos del enunciado
PUTAWAY_ITEMS_HR = 200
PUTAWAY_PUESTOS = 20
PUTAWAY_HR_DIA = 16
PUTAWAY_DIAS_SEM = 6

PICKING_ITEMS_MIN = 1.5
PICKING_PUESTOS = 30

SORTER_TOTES_MIN = 6
SORTER_ITEMS_TOTE = 8


def _val(elem_id: str) -> float:
    el = document.querySelector(f"#{elem_id}")
    return float(el.value) if el else 0.0


def _set(elem_id: str, value: str, color: str = ""):
    el = document.querySelector(f"#{elem_id}")
    if el:
        el.textContent = value
        if color:
            el.style.color = color


def compute(util_pa, eff_pk, hr_pk, dias_pk, target):
    putaway_bruta = PUTAWAY_ITEMS_HR * PUTAWAY_PUESTOS * PUTAWAY_HR_DIA * PUTAWAY_DIAS_SEM * SEM_PER_MES
    putaway_efectiva = putaway_bruta * util_pa
    picking_bruta = PICKING_ITEMS_MIN * 60 * PICKING_PUESTOS * hr_pk * dias_pk * SEM_PER_MES
    picking_efectiva = picking_bruta * eff_pk
    sorter_cap = SORTER_TOTES_MIN * SORTER_ITEMS_TOTE * 60 * hr_pk * dias_pk * SEM_PER_MES

    candidates = [
        ("Put Away", putaway_efectiva),
        ("Picking", picking_efectiva),
        ("Sorter", sorter_cap),
    ]
    bottleneck_name, bottleneck_val = min(candidates, key=lambda x: x[1])

    return {
        "putaway": putaway_efectiva,
        "picking": picking_efectiva,
        "sorter": sorter_cap,
        "bottleneck_name": bottleneck_name,
        "bottleneck_val": bottleneck_val,
        "target": target,
        "alcanza": bottleneck_val >= target,
    }


def fmt(n: float) -> str:
    if n >= 1_000_000:
        return f"{n/1_000_000:.2f}M"
    if n >= 1_000:
        return f"{n/1_000:.0f}K"
    return f"{n:,.0f}"


def render(*_args):
    util_pa = _val("c2-util") / 100.0
    eff_pk = _val("c2-eff") / 100.0
    hr_pk = _val("c2-hr")
    dias_pk = _val("c2-dias")
    target = _val("c2-target")

    _set("c2-util-val", f"{int(util_pa*100)}%")
    _set("c2-eff-val", f"{int(eff_pk*100)}%")
    _set("c2-hr-val", f"{int(hr_pk)}")
    _set("c2-dias-val", f"{int(dias_pk)}")
    _set("c2-target-val", f"{int(target):,}")

    r = compute(util_pa, eff_pk, hr_pk, dias_pk, target)

    _set("c2-out-pa", fmt(r["putaway"]))
    _set("c2-out-pk", fmt(r["picking"]))
    _set("c2-out-so", fmt(r["sorter"]))
    _set("c2-out-bn", r["bottleneck_name"])
    if r["alcanza"]:
        _set("c2-out-status", "✓ Alcanza", "#00A650")
    else:
        deficit = target - r["bottleneck_val"]
        _set("c2-out-status", f"✗ {fmt(deficit)} faltan", "#E63946")

    # Bar chart — capacidades vs target
    bars = {
        "x": ["Put Away", "Picking", "Sorter"],
        "y": [r["putaway"], r["picking"], r["sorter"]],
        "type": "bar",
        "marker": {
            "color": [
                "#E63946" if "Put Away" == r["bottleneck_name"] else "#3483FA",
                "#E63946" if "Picking" == r["bottleneck_name"] else "#3483FA",
                "#E63946" if "Sorter" == r["bottleneck_name"] else "#3483FA",
            ],
            "line": {"color": "#2D3277", "width": 1.2},
        },
        "text": [fmt(r["putaway"]), fmt(r["picking"]), fmt(r["sorter"])],
        "textposition": "outside",
        "hovertemplate": "<b>%{x}</b><br>%{y:,.0f} items/mes<extra></extra>",
        "name": "Capacidad",
    }
    target_line = {
        "x": ["Put Away", "Picking", "Sorter"],
        "y": [target, target, target],
        "type": "scatter",
        "mode": "lines",
        "line": {"color": "#FFE600", "width": 3, "dash": "dash"},
        "name": f"Target {fmt(target)}",
        "hovertemplate": "Target<extra></extra>",
    }
    layout = {
        "title": {
            "text": f"<b>Capacidad efectiva vs target {fmt(target)} items/mes</b><br>"
                    f"<sub style='color:#E63946;'>Bottleneck: {r['bottleneck_name']} = {fmt(r['bottleneck_val'])}</sub>",
            "font": {"family": "Inter, sans-serif", "size": 14, "color": "#2D3277"},
        },
        "margin": {"l": 60, "r": 30, "t": 70, "b": 40},
        "paper_bgcolor": "#F7F7FB",
        "plot_bgcolor": "#FFFFFF",
        "font": {"family": "Inter, sans-serif", "color": "#1A1A2E"},
        "yaxis": {"title": "items/mes", "rangemode": "tozero"},
        "showlegend": True,
        "legend": {"orientation": "h", "y": -0.15, "x": 0.5, "xanchor": "center"},
        "height": 380,
    }
    config = {"displayModeBar": False, "responsive": True}
    Plotly.react("c2-chart", _js([bars, target_line]), _js(layout), _js(config))


_proxy = create_proxy(render)
for sid in ["c2-util", "c2-eff", "c2-hr", "c2-dias", "c2-target"]:
    el = document.querySelector(f"#{sid}")
    if el:
        el.addEventListener("input", _proxy)

render()

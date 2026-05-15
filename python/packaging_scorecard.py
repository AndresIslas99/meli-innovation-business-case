"""
Caso 1 — AHP scorecard interactivo para selección de proveedor de empaquetado.

Pesos editables por el evaluador → ranking de candidatos en tiempo real.
Cada candidato tiene un score 0–100 por criterio (declarado abajo,
defendible desde la investigación en data/proveedores_empaquetado_mx.json).
"""

import json
from pyscript import document
from pyscript.ffi import create_proxy
from js import Plotly, JSON


def _js(obj):
    return JSON.parse(json.dumps(obj))

# ============================================================
#  Vendor candidates · scoring rubric
# ============================================================
# Columns:  [fit_op, soporte_mx, talento, economica, riesgo, sustent]
# Scoring scale: 0 = no fit, 100 = best-in-class.
# Notas en data/proveedores_empaquetado_mx.json + supuestos.json.
VENDORS = {
    "Packsize X-Series (auto-box)": {
        "scores": [85, 90, 70, 75, 70, 90],
        "color": "#FFE600",
        "note": "Auto-box right-sized; sitio MX activo; case Autopartes Tracto. Compró Sparck 2025.",
    },
    "Sparck CMC Pakto (auto-wrap)": {
        "scores": [80, 70, 60, 70, 55, 88],
        "color": "#FCD400",
        "note": "Auto-wrap CartonWrap; canal MX vía Packsize post-merge; transición de soporte.",
    },
    "Sealed Air AutoBag 850HB (bagging)": {
        "scores": [75, 95, 80, 85, 85, 65],
        "color": "#2D3277",
        "note": "Bagging hybrid; 2 plantas MX (Toluca + Monterrey); 700 órdenes/hr; menos right-sizing.",
    },
    "Quadient CVP Impack (auto-box)": {
        "scores": [78, 50, 60, 72, 75, 85],
        "color": "#3483FA",
        "note": "Auto-box; sin oficina MX identificada; soporte vía partner global.",
    },
    "Status quo manual + Smurfit": {
        "scores": [50, 100, 100, 60, 85, 40],
        "color": "#8A8AA0",
        "note": "Baseline cero CAPEX nuevo; alto OPEX humano; over-pack típico.",
    },
}

# Criterios en el mismo orden que el array de scores
CRITERIA = ["Fit operativo", "Soporte MX", "Talento", "Económica", "Riesgo", "Sustentabilidad"]
CRITERIA_SHORT = ["Fit", "Sop", "Tal", "Eco", "Rie", "Sus"]

# IDs de los sliders del HTML
SLIDER_IDS = ["ahp-w1", "ahp-w2", "ahp-w3", "ahp-w4", "ahp-w5", "ahp-w6"]


def _val(slider_id: str) -> float:
    el = document.querySelector(f"#{slider_id}")
    return float(el.value) if el else 0.0


def _set_text(elem_id: str, text: str):
    el = document.querySelector(f"#{elem_id}")
    if el:
        el.textContent = text


def normalize(weights):
    total = sum(weights) or 1.0
    return [w / total for w in weights]


def compute_rankings(weights_norm):
    """Return list of (vendor_name, total_score, criteria_scores, color, note) sorted desc."""
    out = []
    for name, v in VENDORS.items():
        total = sum(w * s for w, s in zip(weights_norm, v["scores"]))
        out.append((name, total, v["scores"], v["color"], v["note"]))
    out.sort(key=lambda x: -x[1])
    return out


def _make_el(tag, *, text=None, classes=None, style=None):
    """Create a DOM element with optional textContent and classes/style. Avoids innerHTML."""
    el = document.createElement(tag)
    if text is not None:
        el.textContent = text
    if classes:
        el.className = classes
    if style:
        el.setAttribute("style", style)
    return el


def render(*_args):
    raw = [_val(s) for s in SLIDER_IDS]
    # show raw % next to slider
    for i, s in enumerate(SLIDER_IDS):
        _set_text(f"{s}-val", f"{int(raw[i])}%")

    norm = normalize(raw)
    ranked = compute_rankings(norm)

    # ----- Bar chart -----
    bar = {
        "x": [r[1] for r in ranked],
        "y": [r[0] for r in ranked],
        "orientation": "h",
        "type": "bar",
        "marker": {
            "color": [r[3] for r in ranked],
            "line": {"color": "#2D3277", "width": 1.2},
        },
        "text": [f"{r[1]:.1f}" for r in ranked],
        "textposition": "outside",
        "hovertemplate": "<b>%{y}</b><br>Score: %{x:.1f}<extra></extra>",
    }
    layout = {
        "title": {
            "text": f"<b>Ranking AHP — pesos normalizados</b><br>"
                    f"<sub>Fit {norm[0]:.0%} · Soporte {norm[1]:.0%} · Talento {norm[2]:.0%} · "
                    f"Econ {norm[3]:.0%} · Riesgo {norm[4]:.0%} · Sustent {norm[5]:.0%}</sub>",
            "font": {"family": "Inter, sans-serif", "size": 14, "color": "#2D3277"},
        },
        "margin": {"l": 280, "r": 60, "t": 70, "b": 40},
        "paper_bgcolor": "#F7F7FB",
        "plot_bgcolor": "#FFFFFF",
        "font": {"family": "Inter, sans-serif", "color": "#1A1A2E"},
        "xaxis": {"title": "Score total (0–100)", "range": [0, 105]},
        "yaxis": {"autorange": "reversed"},
        "height": 360,
    }
    config = {"displayModeBar": False, "responsive": True}
    Plotly.react("ahp-chart", _js([bar]), _js(layout), _js(config))

    # ----- Ranking detail (DOM-safe, no innerHTML) -----
    target = document.querySelector("#ahp-ranking")
    if target:
        # clear children safely
        while target.firstChild:
            target.removeChild(target.firstChild)
        for rank, (name, total, scores, color, note) in enumerate(ranked, start=1):
            row = _make_el(
                "div",
                style=(
                    f"padding:0.6rem 0.9rem;border-left:3px solid {color};"
                    f"background:#FFFFFF;margin:0.35rem 0;border-radius:0 6px 6px 0;"
                ),
            )
            header = _make_el(
                "div",
                style="display:flex;justify-content:space-between;align-items:baseline;gap:1rem;",
            )
            title_box = _make_el("div")
            title_strong = _make_el("strong", text=f"#{rank} {name}")
            title_box.appendChild(title_strong)
            if rank == 1:
                winner_tag = _make_el(
                    "span",
                    text=" winner",
                    classes="brand-tag brand-tag--yellow",
                    style="font-size:0.65rem;margin-left:0.4em;",
                )
                title_box.appendChild(winner_tag)
            score_span = _make_el(
                "span",
                text=f"{total:.1f}",
                style="font-family:JetBrains Mono,monospace;font-weight:600;color:#2D3277;",
            )
            header.appendChild(title_box)
            header.appendChild(score_span)
            row.appendChild(header)

            # breakdown
            breakdown = _make_el(
                "div",
                style="font-size:0.78rem;color:#4A4A6A;margin-top:0.2em;",
            )
            parts = []
            for i in range(len(CRITERIA)):
                label = _make_el(
                    "span",
                    text=f"{CRITERIA_SHORT[i]} ",
                    style="color:#888;",
                )
                value = _make_el("span", text=f"{scores[i]}")
                parts.append((label, value))
            for j, (lbl, val) in enumerate(parts):
                breakdown.appendChild(lbl)
                breakdown.appendChild(val)
                if j < len(parts) - 1:
                    sep = _make_el("span", text=" · ", style="color:#CCC;")
                    breakdown.appendChild(sep)
            row.appendChild(breakdown)

            note_el = _make_el(
                "div",
                text=note,
                style="font-size:0.78rem;color:#8A8AA0;margin-top:0.15em;font-style:italic;",
            )
            row.appendChild(note_el)

            target.appendChild(row)


# Attach listeners (use create_proxy so JS can call the Python function)
_proxy = create_proxy(render)
for sid in SLIDER_IDS:
    el = document.querySelector(f"#{sid}")
    if el:
        el.addEventListener("input", _proxy)

# First paint
render()

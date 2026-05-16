"""
Caso 5 — Gantt simulator con CPM + Monte Carlo de duraciones.

Sliders:
- Multiplicador discovery (fases 1–4)
- Multiplicador POC (fase 5)
- Multiplicador licitación (fases 8–10)
- Fabricación (semanas)
- Ramp-up (semanas)

Output: weeks to contract, weeks to steady-state, p50 / p90 Monte Carlo.
"""

import json
import random
from pyscript import document
from pyscript.ffi import create_proxy
from js import Plotly, JSON


def _js(obj):
    return JSON.parse(json.dumps(obj))

# Fases en sem (base case)
PHASES_BASE = [
    {"name": "1. Definición + alineación", "dur": 2, "group": "discovery"},
    {"name": "2. Baseline + KPIs target",  "dur": 3, "group": "discovery"},
    {"name": "3. Long-list + RFI",         "dur": 4, "group": "discovery"},
    {"name": "4. Due diligence + visitas", "dur": 4, "group": "discovery"},
    {"name": "5. POC / piloto",            "dur": 6, "group": "poc"},
    {"name": "6. Business case + ROI",     "dur": 3, "group": "poc"},
    {"name": "7. Aprobación senior",       "dur": 2, "group": "poc"},
    {"name": "8. Spec técnica + RFP",      "dur": 3, "group": "lic"},
    {"name": "9. Licitación + selección",  "dur": 4, "group": "lic"},
    {"name": "10. Negociación contrato",   "dur": 3, "group": "lic"},
    {"name": "Fabricación (dado)",         "dur": 12, "group": "fab"},
    {"name": "11. Preparación sitio",      "dur": 4, "group": "sitio", "parallel_with": "fab"},
    {"name": "12. Instalación + commiss.", "dur": 4, "group": "install"},
    {"name": "13. Ramp-up + validación",   "dur": 6, "group": "rampup"},
]


def _val(elem_id: str) -> float:
    el = document.querySelector(f"#{elem_id}")
    return float(el.value) if el else 0.0


def _set(elem_id: str, value: str, color: str = ""):
    el = document.querySelector(f"#{elem_id}")
    if el:
        el.textContent = value
        if color:
            el.style.color = color


def apply_multipliers(mult_dis, mult_poc, mult_lic, fab_wk, ramp_wk):
    phases = []
    for p in PHASES_BASE:
        dur = p["dur"]
        g = p["group"]
        if g == "discovery":
            dur *= mult_dis
        elif g == "poc":
            dur *= mult_poc
        elif g == "lic":
            dur *= mult_lic
        elif g == "fab":
            dur = fab_wk
        elif g == "rampup":
            dur = ramp_wk
        new = dict(p)
        new["dur"] = max(0.5, dur)
        phases.append(new)
    return phases


def schedule(phases):
    """Compute start/end for each phase. Phase 'sitio' runs in parallel with 'fab'."""
    timeline = []
    cursor = 0.0
    fab_start = None
    fab_end = None
    for p in phases:
        if p.get("parallel_with") == "fab" and fab_start is not None:
            start = fab_start
            end = start + p["dur"]
            # don't advance cursor
        else:
            start = cursor
            end = start + p["dur"]
            cursor = end
            if p["group"] == "fab":
                fab_start = start
                fab_end = end
        timeline.append({"name": p["name"], "start": start, "end": end, "group": p["group"]})
    return timeline


def _beta_pert_sample(low, mode, high, lam=4.0):
    """
    Beta-PERT sampling (Vose D., "Risk Analysis", 3rd ed., Wiley 2008).
    Distribución suave (no kink en moda) que sobrepesa el modo realista vs colas.
    Más realista que triangular para project durations.
    α = 1 + λ·(mode-low)/(high-low),  β = 1 + λ·(high-mode)/(high-low),  λ=4 standard.
    """
    if high <= low:
        return low
    a = 1.0 + lam * (mode - low) / (high - low)
    b = 1.0 + lam * (high - mode) / (high - low)
    return low + (high - low) * random.betavariate(a, b)


def monte_carlo(mult_dis, mult_poc, mult_lic, fab_wk, ramp_wk, n=600):
    """Beta-PERT ±20% sobre cada fase ya multiplicada (Vose 2008, mejor que triangular)."""
    end_distribution = []
    contract_distribution = []
    for _ in range(n):
        phases = apply_multipliers(mult_dis, mult_poc, mult_lic, fab_wk, ramp_wk)
        perturbed = []
        for p in phases:
            d = p["dur"]
            low, mode, high = d * 0.8, d, d * 1.2
            d_perturbed = _beta_pert_sample(low, mode, high)
            perturbed.append({**p, "dur": d_perturbed})
        t = schedule(perturbed)
        # weeks to contract = end of phase 10 (negociación contrato)
        contract_idx = 9  # 0-indexed phase 10
        # weeks to steady state = end of last phase
        contract_distribution.append(t[contract_idx]["end"])
        end_distribution.append(t[-1]["end"])
    end_distribution.sort()
    contract_distribution.sort()
    p50 = end_distribution[n // 2]
    p90 = end_distribution[int(n * 0.9)]
    return p50, p90, contract_distribution, end_distribution


def render(*_args):
    mult_dis = _val("c5-mult-discovery")
    mult_poc = _val("c5-mult-poc")
    mult_lic = _val("c5-mult-lic")
    fab_wk = _val("c5-fab")
    ramp_wk = _val("c5-ramp")

    _set("c5-mult-discovery-val", f"{mult_dis:.1f}×")
    _set("c5-mult-poc-val", f"{mult_poc:.1f}×")
    _set("c5-mult-lic-val", f"{mult_lic:.1f}×")
    _set("c5-fab-val", f"{int(fab_wk)}")
    _set("c5-ramp-val", f"{int(ramp_wk)}")

    phases = apply_multipliers(mult_dis, mult_poc, mult_lic, fab_wk, ramp_wk)
    timeline = schedule(phases)
    weeks_to_contract = timeline[9]["end"]
    weeks_to_steady = timeline[-1]["end"]

    p50, p90, _, _ = monte_carlo(mult_dis, mult_poc, mult_lic, fab_wk, ramp_wk, n=500)

    _set("c5-out-pre", f"{weeks_to_contract:.0f} sem")
    _set("c5-out-go", f"{weeks_to_steady:.0f} sem")
    _set("c5-out-mc-p50", f"{p50:.0f} sem")
    _set("c5-out-mc-p90", f"{p90:.0f} sem")

    # Gantt — horizontal bars
    group_colors = {
        "discovery": "#3483FA",
        "poc": "#00A650",
        "lic": "#2D3277",
        "fab": "#FFB400",
        "sitio": "#FCD400",
        "install": "#2D3277",
        "rampup": "#FFE600",
    }

    bars = []
    for t in timeline:
        bars.append({
            "x": [t["end"] - t["start"]],
            "y": [t["name"]],
            "base": [t["start"]],
            "orientation": "h",
            "type": "bar",
            "marker": {
                "color": group_colors.get(t["group"], "#8A8AA0"),
                "line": {"color": "#2D3277", "width": 1},
            },
            "text": [f"{t['end'] - t['start']:.1f}w"],
            "textposition": "inside",
            "insidetextanchor": "middle",
            "name": t["name"],
            "showlegend": False,
            "hovertemplate": f"<b>{t['name']}</b><br>Sem {t['start']:.1f} → {t['end']:.1f}<extra></extra>",
        })

    # Marca el momento "contract signed"
    marker_contract = {
        "x": [weeks_to_contract, weeks_to_contract],
        "y": [0, len(timeline) - 1],
        "type": "scatter",
        "mode": "lines",
        "line": {"color": "#E63946", "width": 2, "dash": "dot"},
        "name": "Contract signed",
        "hoverinfo": "skip",
        "showlegend": False,
    }

    layout = {
        "title": {
            "text": f"<b>Gantt + critical path</b> · Contrato sem {weeks_to_contract:.0f} · Steady state sem {weeks_to_steady:.0f}<br>"
                    f"<sub>Monte Carlo (n=500): p50 = {p50:.0f} sem · p90 = {p90:.0f} sem</sub>",
            "font": {"family": "Inter, sans-serif", "size": 14, "color": "#2D3277"},
        },
        "margin": {"l": 230, "r": 30, "t": 80, "b": 50},
        "paper_bgcolor": "#F7F7FB",
        "plot_bgcolor": "#FFFFFF",
        "font": {"family": "Inter, sans-serif", "color": "#1A1A2E"},
        "xaxis": {"title": "Semanas desde inicio del proyecto"},
        "yaxis": {"autorange": "reversed", "categoryorder": "trace"},
        "barmode": "stack",
        "showlegend": False,
        "height": 480,
    }
    config = {
        "displayModeBar": True,
        "modeBarButtonsToRemove": ["select2d", "lasso2d", "toggleSpikelines", "hoverClosestCartesian", "hoverCompareCartesian", "sendDataToCloud"],
        "displaylogo": False,
        "doubleClick": "reset",
        "responsive": True,
        "toImageButtonOptions": {"format": "png", "filename": "meli-gantt-roadmap"},
    }

    Plotly.react("c5-chart", _js(bars + [marker_contract]), _js(layout), _js(config))


_proxy = create_proxy(render)
for sid in ["c5-mult-discovery", "c5-mult-poc", "c5-mult-lic", "c5-fab", "c5-ramp"]:
    el = document.querySelector(f"#{sid}")
    if el:
        el.addEventListener("input", _proxy)

render()

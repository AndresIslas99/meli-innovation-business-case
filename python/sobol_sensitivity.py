"""
Sobol indices widget — variance-based sensitivity analysis pure-Python.

Saltelli (2002) estimator descompone la varianza del NPV del business case
en contribuciones por variable. Identifica qué variable domina (aplicar
foco) vs cuáles son ruido (no investir más en estimación).

Referencias:
- Sobol I. M. "Sensitivity estimates for nonlinear mathematical models". 1993.
- Saltelli A. "Making best use of model evaluations to compute sensitivity indices". 2002.
- SALib (Python tool): https://salib.readthedocs.io/

Implementación pure stdlib (sin SALib) para portabilidad PyScript.
"""

import math
import random
import json
from pyscript import document
from pyscript.ffi import create_proxy
from js import Plotly, JSON


def _js(obj):
    return JSON.parse(json.dumps(obj))


def _val(eid):
    el = document.querySelector(f"#{eid}")
    return float(el.value) if el else 0.0


def _set(eid, txt, color=""):
    el = document.querySelector(f"#{eid}")
    if el:
        el.textContent = txt
        if color:
            el.style.color = color


# ============================================================
#  Business case NPV model
# ============================================================
# 6 variables driving NPV de proyecto automatización CEDIS:
VAR_NAMES = [
    "CAPEX auto-box",
    "CAPEX STP / AMR",
    "Adoption rate Y1",
    "Cost-per-unit base",
    "Volumen mensual",
    "Discount rate",
]
VAR_BOUNDS = [
    (200_000, 800_000),       # CAPEX auto-box USD
    (500_000, 3_000_000),     # CAPEX STP USD
    (0.30, 0.90),             # Adoption %
    (1.5, 3.5),               # Cost per unit USD/item
    (1_000_000, 3_000_000),   # Volumen items/mes
    (0.08, 0.18),             # Discount rate decimal
]

HORIZON_MONTHS = 36
COST_REDUCTION_PCT = 0.15  # supuesto: -15% cost-per-unit cuando aplicado


def model_npv(x):
    """Returns NPV (USD) over HORIZON_MONTHS."""
    capex_box, capex_stp, adoption, cpu_base, volume, disc = x
    monthly_save = cpu_base * COST_REDUCTION_PCT * volume * adoption
    npv = -(capex_box + capex_stp)
    monthly_rate = disc / 12.0
    for m in range(1, HORIZON_MONTHS + 1):
        npv += monthly_save / ((1.0 + monthly_rate) ** m)
    return npv


# ============================================================
#  Saltelli (2002) Sobol estimator
# ============================================================

def saltelli_estimate(n_samples):
    """
    Returns first-order S_i and total-order S_Ti indices.
    Uses 2N samples (matrices A and B) plus N*k = N*6 'ABi' samples.
    Total model evaluations: N*(k+2) = N*8 for 6 vars.
    """
    k = len(VAR_BOUNDS)

    def _sample():
        return [random.uniform(*VAR_BOUNDS[i]) for i in range(k)]

    A = [_sample() for _ in range(n_samples)]
    B = [_sample() for _ in range(n_samples)]
    yA = [model_npv(row) for row in A]
    yB = [model_npv(row) for row in B]

    yA_mean = sum(yA) / n_samples
    var_y = sum((y - yA_mean) ** 2 for y in yA) / max(1, n_samples - 1)
    if var_y < 1e-6:
        return [0.0] * k, [0.0] * k

    s1 = []
    sT = []
    for i in range(k):
        yABi = []
        for kk in range(n_samples):
            row = list(A[kk])
            row[i] = B[kk][i]
            yABi.append(model_npv(row))
        # Saltelli 2010 best-practice estimators
        # First-order: V_i = (1/N) Σ yB·(yABi - yA)
        v_i = sum(yB[m] * (yABi[m] - yA[m]) for m in range(n_samples)) / n_samples
        # Total-order: V_Ti = (1/2N) Σ (yA - yABi)^2
        v_ti = sum((yA[m] - yABi[m]) ** 2 for m in range(n_samples)) / (2 * n_samples)
        s1.append(max(0.0, min(1.0, v_i / var_y)))
        sT.append(max(0.0, min(1.0, v_ti / var_y)))
    return s1, sT


# ============================================================
#  Render
# ============================================================

def render(*_args):
    try:
        n_samples = max(64, min(int(_val("sobol-n-samples")), 1024))
        _set("sobol-n-samples-val", str(n_samples))

        s1, sT = saltelli_estimate(n_samples)

        # Find dominant variable by total order
        max_idx = max(range(len(sT)), key=lambda i: sT[i])
        _set("sobol-dominant", VAR_NAMES[max_idx])
        _set("sobol-dominant-pct", f"S_T = {sT[max_idx]:.0%}")

        # Bar chart
        bar_s1 = {
            "x": VAR_NAMES,
            "y": s1,
            "name": "S₁ (first-order)",
            "type": "bar",
            "marker": {"color": "#2D3277"},
            "text": [f"{v:.2f}" for v in s1],
            "textposition": "outside",
        }
        bar_sT = {
            "x": VAR_NAMES,
            "y": sT,
            "name": "S_T (total)",
            "type": "bar",
            "marker": {"color": "#FFE600", "line": {"color": "#2D3277", "width": 1}},
            "text": [f"{v:.2f}" for v in sT],
            "textposition": "outside",
        }
        layout = {
            "title": {
                "text": f"<b>Sobol indices · sensibilidad varianza-based del NPV</b>"
                        f"<br><sub>Saltelli 2002 estimator · n={n_samples} muestras · {n_samples*8} model evals</sub>",
                "font": {"family": "Inter, sans-serif", "size": 13, "color": "#2D3277"},
            },
            "margin": {"l": 60, "r": 30, "t": 70, "b": 90},
            "paper_bgcolor": "#F7F7FB",
            "plot_bgcolor": "#FFFFFF",
            "font": {"family": "Inter, sans-serif", "color": "#1A1A2E"},
            "yaxis": {"title": "Sobol index (fracción de varianza explicada)", "range": [0, 1.1]},
            "xaxis": {"tickangle": -25},
            "barmode": "group",
            "showlegend": True,
            "legend": {"orientation": "h", "y": -0.32, "x": 0.5, "xanchor": "center"},
            "height": 400,
        }
        config = {
            "displayModeBar": True,
            "modeBarButtonsToRemove": ["select2d", "lasso2d", "toggleSpikelines", "hoverClosestCartesian", "hoverCompareCartesian", "sendDataToCloud"],
            "displaylogo": False,
            "doubleClick": "reset",
            "responsive": True,
            "toImageButtonOptions": {"format": "png", "filename": "meli-sobol-sensitivity"},
        }
        Plotly.react("sobol-chart", _js([bar_s1, bar_sT]), _js(layout), _js(config))
    except Exception as e:
        _set("sobol-dominant", "Error")
        _set("sobol-dominant-pct", f"{str(e)[:60]}")


_proxy = create_proxy(render)
for sid in ["sobol-n-samples"]:
    el = document.querySelector(f"#{sid}")
    if el:
        el.addEventListener("input", _proxy)

render()

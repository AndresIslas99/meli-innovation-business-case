"""
Real Options binomial tree (Cox-Ross-Rubinstein 1979).

Valora el gate de Fase 7 del roadmap (caso 5) como una CALL option:
MELI tiene el DERECHO (no la obligación) de invertir el CAPEX completo de
fabricación si el POC valida. La flexibilidad tiene valor — NPV "punto" lo subestima.

Cita: Cox, Ross, Rubinstein. "Option pricing: A simplified approach". J Financial Economics 7(3), 1979.
"""

import math
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


def binomial_tree(S, K, r, sigma, T, N):
    """
    Cox-Ross-Rubinstein binomial tree pricing — American call (early exercise allowed).
    S = current PV of underlying (USD K)
    K = strike (CAPEX) (USD K)
    r = risk-free rate (annual, decimal)
    sigma = volatility (annual, decimal)
    T = time to expiration (years)
    N = number of time steps
    """
    dt = T / N
    u = math.exp(sigma * math.sqrt(dt))
    d = 1.0 / u
    disc = math.exp(-r * dt)
    p = (math.exp(r * dt) - d) / (u - d)
    p = max(0.0, min(1.0, p))  # numerical clamp

    # Prices at each node (i = time step, j = number of down moves)
    prices = [[S * (u ** (i - j)) * (d ** j) for j in range(i + 1)] for i in range(N + 1)]

    # Terminal payoffs (call: max(S - K, 0))
    values_terminal = [max(prices[N][j] - K, 0.0) for j in range(N + 1)]

    # Initialize full grid of node values for visualization
    node_values = [[0.0] * (i + 1) for i in range(N + 1)]
    node_values[N] = list(values_terminal)
    early_exercise = [[False] * (i + 1) for i in range(N + 1)]

    # Backward induction
    values = list(values_terminal)
    for i in range(N - 1, -1, -1):
        new_values = []
        for j in range(i + 1):
            continuation = disc * (p * values[j] + (1.0 - p) * values[j + 1])
            payoff_now = max(prices[i][j] - K, 0.0)
            v = max(continuation, payoff_now)
            new_values.append(v)
            node_values[i][j] = v
            early_exercise[i][j] = payoff_now > continuation and payoff_now > 0
        values = new_values

    option_value = values[0]
    intrinsic_value = max(S - K, 0.0)
    time_value = option_value - intrinsic_value
    return dict(
        option_value=option_value,
        intrinsic_value=intrinsic_value,
        time_value=time_value,
        u=u, d=d, p=p,
        prices=prices,
        node_values=node_values,
        early_exercise=early_exercise,
    )


def render(*_args):
    try:
        S = _val("ro-S")  # current PV (USD K)
        K = _val("ro-K")  # strike CAPEX (USD K)
        sigma = _val("ro-sigma") / 100.0
        T = _val("ro-T")
        r = _val("ro-r") / 100.0
        N = max(2, int(_val("ro-N")))

        _set("ro-S-val", f"${S:.0f}K")
        _set("ro-K-val", f"${K:.0f}K")
        _set("ro-sigma-val", f"{sigma:.0%}")
        _set("ro-T-val", f"{T:.2f}y")
        _set("ro-r-val", f"{r:.1%}")
        _set("ro-N-val", str(N))

        result = binomial_tree(S, K, r, sigma, T, N)

        ov = result["option_value"]
        iv = result["intrinsic_value"]
        tv = result["time_value"]

        _set("ro-out-value", f"${ov:,.1f}K")
        _set("ro-out-intrinsic", f"${iv:,.1f}K")
        _set("ro-out-timevalue", f"${tv:,.1f}K", "#00A650" if tv > 1e-3 else "#E63946")
        _set("ro-out-params", f"u={result['u']:.3f} · d={result['d']:.3f} · p={result['p']:.3f}")

        # Build tree scatter
        prices = result["prices"]
        node_vals = result["node_values"]
        early = result["early_exercise"]

        x_pts, y_pts, txt, colors = [], [], [], []
        for i in range(N + 1):
            for j in range(i + 1):
                x_pts.append(i * T / N)
                y_pts.append(prices[i][j])
                txt.append(
                    f"t={i*T/N:.2f}y<br>S=${prices[i][j]:.0f}K<br>V=${node_vals[i][j]:.1f}K"
                )
                colors.append("#00A650" if early[i][j] else "#3483FA")

        scatter = {
            "x": x_pts,
            "y": y_pts,
            "mode": "markers+text",
            "type": "scatter",
            "marker": {"size": 14, "color": colors, "line": {"color": "#2D3277", "width": 1.5}},
            "text": [f"{v:.0f}" for v in [node_vals[i][j] for i in range(N + 1) for j in range(i + 1)]],
            "textposition": "top center",
            "textfont": {"size": 9, "color": "#2D3277"},
            "hovertemplate": "%{customdata}<extra></extra>",
            "customdata": txt,
            "showlegend": False,
        }

        # Connecting lines
        lines_x, lines_y = [], []
        for i in range(N):
            for j in range(i + 1):
                lines_x.extend([i * T / N, (i + 1) * T / N, None])
                lines_y.extend([prices[i][j], prices[i + 1][j], None])
                lines_x.extend([i * T / N, (i + 1) * T / N, None])
                lines_y.extend([prices[i][j], prices[i + 1][j + 1], None])

        lines = {
            "x": lines_x,
            "y": lines_y,
            "mode": "lines",
            "type": "scatter",
            "line": {"color": "#D8D8E3", "width": 1},
            "hoverinfo": "skip",
            "showlegend": False,
        }

        # Strike line
        strike_line = {
            "x": [0, T],
            "y": [K, K],
            "mode": "lines",
            "type": "scatter",
            "line": {"color": "#E63946", "width": 1.5, "dash": "dash"},
            "name": f"Strike (CAPEX) = ${K:.0f}K",
            "hovertemplate": "Strike<extra></extra>",
        }

        layout = {
            "title": {
                "text": f"<b>Binomial tree · Cox-Ross-Rubinstein 1979</b><br>"
                        f"<sub>● mantener opción · ● ejercer temprano · ━ línea roja = strike (CAPEX)</sub>",
                "font": {"family": "Inter, sans-serif", "size": 13, "color": "#2D3277"},
            },
            "margin": {"l": 60, "r": 30, "t": 70, "b": 50},
            "paper_bgcolor": "#F7F7FB",
            "plot_bgcolor": "#FFFFFF",
            "font": {"family": "Inter, sans-serif", "color": "#1A1A2E"},
            "xaxis": {"title": "Tiempo (años hasta gate Fase 7)"},
            "yaxis": {"title": "Valor del subyacente (USD K)"},
            "showlegend": True,
            "legend": {"orientation": "h", "y": -0.18, "x": 0.5, "xanchor": "center"},
            "height": 420,
        }
        config = {
            "displayModeBar": True,
            "modeBarButtonsToRemove": ["select2d", "lasso2d", "toggleSpikelines", "hoverClosestCartesian", "hoverCompareCartesian", "sendDataToCloud"],
            "displaylogo": False,
            "doubleClick": "reset",
            "responsive": True,
            "toImageButtonOptions": {"format": "png", "filename": "meli-real-options-binomial"},
        }
        Plotly.react("ro-chart", _js([lines, strike_line, scatter]), _js(layout), _js(config))
    except Exception as e:
        _set("ro-out-value", f"Error")
        _set("ro-out-params", f"{str(e)[:80]}")


_proxy = create_proxy(render)
for sid in ["ro-S", "ro-K", "ro-sigma", "ro-T", "ro-r", "ro-N"]:
    el = document.querySelector(f"#{sid}")
    if el:
        el.addEventListener("input", _proxy)

render()

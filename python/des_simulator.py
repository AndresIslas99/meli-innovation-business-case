"""
Discrete Event Simulation widget — pure Python stdlib (heapq).

Modela el flow Put Away → Picking → Sorter del Caso 2 con variabilidad
estocástica (llegadas Poisson, servicios Gamma). Captura lo que TOC lineal
asume como promedio — el bottleneck DINÁMICO puede ser distinto al promedio.

Referencia: Banks, Carson, Nelson, Nicol. "Discrete-Event System Simulation" (5th ed., 2009).
En producción se usaría SimPy o AnyLogic; aquí stdlib para portabilidad PyScript.
"""

import math
import random
import heapq
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
#  DES core (heapq-based)
# ============================================================

class _Event:
    __slots__ = ("t", "seq", "action")

    def __init__(self, t, seq, action):
        self.t = t
        self.seq = seq
        self.action = action

    def __lt__(self, other):
        return (self.t, self.seq) < (other.t, other.seq)


class DESModel:
    """3 estaciones en serie: PutAway → Picking → Sorter."""

    def __init__(self, rate_arr, cap_pa, cap_pk, cap_so, cv, sim_hours):
        self.rate_arr = rate_arr   # llegadas por hora
        self.cap_pa = cap_pa       # tasa servicio Put Away (items/hr)
        self.cap_pk = cap_pk       # tasa servicio Picking
        self.cap_so = cap_so       # tasa servicio Sorter
        self.cv = max(0.05, cv)
        self.sim_hours = sim_hours

        self.now = 0.0
        self.events = []
        self.seq = 0

        self.q_pa = 0
        self.q_pk = 0
        self.q_so = 0
        self.busy_pa = 0
        self.busy_pk = 0
        self.busy_so = 0

        self.completed = 0
        self.max_q_pk_to_so = 0
        self.peak_bottleneck_counts = {"PutAway": 0, "Picking": 0, "Sorter": 0}

    def _gamma(self, rate_per_hr):
        # service time = 1/rate (mean), gamma con cv configurable
        mean_t = 1.0 / max(rate_per_hr, 1e-6)
        k = 1.0 / (self.cv ** 2)
        theta = mean_t / k
        return random.gammavariate(k, theta)

    def _schedule(self, delay, action):
        self.seq += 1
        heapq.heappush(self.events, _Event(self.now + delay, self.seq, action))

    def _bottleneck_snapshot(self):
        qs = [("PutAway", self.q_pa), ("Picking", self.q_pk), ("Sorter", self.q_so)]
        name, _ = max(qs, key=lambda x: x[1])
        self.peak_bottleneck_counts[name] += 1

    def _arrival(self):
        if self.now < self.sim_hours:
            inter = random.expovariate(self.rate_arr)
            self._schedule(inter, self._arrival)
        if self.busy_pa == 0:
            self.busy_pa = 1
            self._schedule(self._gamma(self.cap_pa), self._complete_pa)
        else:
            self.q_pa += 1

    def _complete_pa(self):
        self.busy_pa = 0
        if self.q_pa > 0:
            self.q_pa -= 1
            self.busy_pa = 1
            self._schedule(self._gamma(self.cap_pa), self._complete_pa)
        # forward to picking
        if self.busy_pk == 0:
            self.busy_pk = 1
            self._schedule(self._gamma(self.cap_pk), self._complete_pk)
        else:
            self.q_pk += 1

    def _complete_pk(self):
        self.busy_pk = 0
        if self.q_pk > 0:
            self.q_pk -= 1
            self.busy_pk = 1
            self._schedule(self._gamma(self.cap_pk), self._complete_pk)
        # forward to sorter
        if self.q_so > self.max_q_pk_to_so:
            self.max_q_pk_to_so = self.q_so
        if self.busy_so == 0:
            self.busy_so = 1
            self._schedule(self._gamma(self.cap_so), self._complete_so)
        else:
            self.q_so += 1

    def _complete_so(self):
        self.busy_so = 0
        self.completed += 1
        if self.q_so > 0:
            self.q_so -= 1
            self.busy_so = 1
            self._schedule(self._gamma(self.cap_so), self._complete_so)

    def run(self):
        first_inter = random.expovariate(self.rate_arr)
        self._schedule(first_inter, self._arrival)
        snapshot_every = max(self.sim_hours / 50, 1.0)
        next_snapshot = snapshot_every
        while self.events and self.events[0].t <= self.sim_hours:
            evt = heapq.heappop(self.events)
            self.now = evt.t
            if self.now >= next_snapshot:
                self._bottleneck_snapshot()
                next_snapshot += snapshot_every
            evt.action()
        return dict(
            throughput=self.completed,
            max_q_buf=self.max_q_pk_to_so,
            bottleneck_counts=self.peak_bottleneck_counts,
        )


# ============================================================
#  Render
# ============================================================

def render(*_args):
    try:
        rate_arr = _val("des-rate-pa")          # llegadas/hr
        cv = _val("des-cv")
        sim_hours = _val("des-sim-hours")
        n_runs = max(1, int(_val("des-n-runs")))

        _set("des-rate-pa-val", f"{rate_arr:.0f}/hr")
        _set("des-cv-val", f"{cv:.2f}")
        _set("des-sim-hours-val", f"{int(sim_hours)}h")
        _set("des-n-runs-val", str(n_runs))

        # Capacities efectivas (de Caso 2 — escaladas a items/hr)
        # Put Away: 200 items/hr/puesto × 20 puestos × 10% util = 400 items/hr
        # Picking: 1.5 × 60 × 30 × 90% eff = 2,430 items/hr
        # Sorter: 6 totes/min × 8 items/tote × 60 = 2,880 items/hr
        cap_pa = 200 * 20 * 0.10
        cap_pk = 1.5 * 60 * 30 * 0.90
        cap_so = 6 * 8 * 60

        throughputs = []
        bottleneck_counts_agg = {"PutAway": 0, "Picking": 0, "Sorter": 0}
        for _ in range(n_runs):
            m = DESModel(rate_arr, cap_pa, cap_pk, cap_so, cv, sim_hours)
            r = m.run()
            throughputs.append(r["throughput"])
            for k, v in r["bottleneck_counts"].items():
                bottleneck_counts_agg[k] += v

        throughputs.sort()
        mean = sum(throughputs) / n_runs
        p10 = throughputs[max(0, int(n_runs * 0.10))]
        p50 = throughputs[int(n_runs * 0.50)]
        p90 = throughputs[min(n_runs - 1, int(n_runs * 0.90))]

        # Linear baseline projection: capacidad efectiva más baja × sim_hours
        # PA es bottleneck: 400 items/hr × sim_hours
        linear_proj = min(cap_pa, cap_pk, cap_so) * sim_hours

        delta_vs_linear = (mean - linear_proj) / linear_proj * 100 if linear_proj > 0 else 0

        _set("des-out-mean", f"{mean:,.0f}")
        _set("des-out-p10", f"{p10:,.0f}")
        _set("des-out-p50", f"{p50:,.0f}")
        _set("des-out-p90", f"{p90:,.0f}")
        _set("des-out-linear", f"{linear_proj:,.0f}")
        _set("des-out-delta", f"{delta_vs_linear:+.1f}%",
             "#00A650" if delta_vs_linear >= -2 else "#E63946")

        # Identify dynamic bottleneck (highest aggregate count)
        dyn = max(bottleneck_counts_agg.items(), key=lambda x: x[1])
        _set("des-out-bottleneck", f"{dyn[0]} ({dyn[1]/sum(bottleneck_counts_agg.values()):.0%} snapshots)")

        # Histogram chart
        hist = {
            "x": throughputs,
            "type": "histogram",
            "marker": {"color": "#3483FA", "line": {"color": "#2D3277", "width": 1}},
            "name": "Throughput simulado",
            "nbinsx": min(15, max(5, n_runs // 2)),
        }
        baseline_line = {
            "x": [linear_proj, linear_proj],
            "y": [0, max(1, n_runs // 3)],
            "type": "scatter",
            "mode": "lines",
            "line": {"color": "#E63946", "width": 3, "dash": "dash"},
            "name": f"TOC lineal {linear_proj:,.0f}",
        }
        layout = {
            "title": {
                "text": f"<b>DES · throughput distribution (n={n_runs} runs · sim {int(sim_hours)}h)</b>"
                        f"<br><sub>Bottleneck dinámico: {dyn[0]} · vs lineal Δ {delta_vs_linear:+.1f}%</sub>",
                "font": {"family": "Inter, sans-serif", "size": 13, "color": "#2D3277"},
            },
            "margin": {"l": 60, "r": 30, "t": 75, "b": 50},
            "paper_bgcolor": "#F7F7FB",
            "plot_bgcolor": "#FFFFFF",
            "font": {"family": "Inter, sans-serif", "color": "#1A1A2E"},
            "xaxis": {"title": "Items procesados en el horizonte simulado"},
            "yaxis": {"title": "Frecuencia (runs)"},
            "showlegend": True,
            "legend": {"orientation": "h", "y": -0.18, "x": 0.5, "xanchor": "center"},
            "height": 360,
        }
        config = {
            "displayModeBar": True,
            "modeBarButtonsToRemove": ["select2d", "lasso2d", "toggleSpikelines", "hoverClosestCartesian", "hoverCompareCartesian", "sendDataToCloud"],
            "displaylogo": False,
            "doubleClick": "reset",
            "responsive": True,
            "toImageButtonOptions": {"format": "png", "filename": "meli-des-stp"},
        }
        Plotly.react("des-chart", _js([hist, baseline_line]), _js(layout), _js(config))
    except Exception as e:
        _set("des-out-mean", "Error")
        _set("des-out-p10", f"{str(e)[:60]}")


_proxy = create_proxy(render)
for sid in ["des-rate-pa", "des-cv", "des-sim-hours", "des-n-runs"]:
    el = document.querySelector(f"#{sid}")
    if el:
        el.addEventListener("input", _proxy)

render()

"""
Caso 1 — AHP scorecard interactivo + rúbrica completa visible.

Diseño:
1. Sliders de pesos por criterio (6).
2. Chart Plotly del ranking.
3. Panel de RÚBRICA (definitions + anchors 0/25/50/75/100 por criterio) — visible al cargar.
4. Panel de EVIDENCIA POR VENDOR — cada vendor expone sus 6 evidencias con source links — visible al cargar.

Toda la rúbrica está en `data/rubrica_ahp.json`; aquí se embebe como dict Python para que
el widget sea self-contained (no async fetch). Sincronización con el JSON canónico es manual.
"""

import json
from pyscript import document
from pyscript.ffi import create_proxy
from js import Plotly, JSON


def _js(obj):
    return JSON.parse(json.dumps(obj))


# ============================================================
#  Rúbrica embebida (espejo de data/rubrica_ahp.json)
# ============================================================

CRITERIA = [
    ("fit_operativo",   "①", "Fit operativo",   "Cómo encaja la tecnología con el SKU mix, throughput requerido e integraciones (WMS/MES, OPC UA).", {
        0:   "No soporta el SKU mix; throughput insuficiente.",
        25:  "Soporta solo un sub-segmento; requiere bypass manual.",
        50:  "Cubre rango típico; performa al nominal; integraciones básicas.",
        75:  "Cubre todo el rango + márgenes; integraciones nativas WMS.",
        100: "Cubre rango ampliado x1.5; integraciones + OPC UA + APIs abiertas."
    }),
    ("soporte_mx",      "②", "Soporte MX",      "Oficinas / plantas / distribuidores MX, SLA, stock refacciones >24mo, MTTR contractual, casos LATAM.", {
        0:   "Sin canal LATAM; SLA >72 hrs.",
        25:  "Solo vía partner regional Brasil/Chile; SLA 48-72 hrs.",
        50:  "Distribuidor MX certificado; SLA <48 hrs; sin stock local refacciones.",
        75:  "Oficina/distribuidor MX directo; SLA <24 hrs; stock local >24 mo.",
        100: "Planta o assembly MX; FSE residente; MTTR ≤4 hrs contractual."
    }),
    ("talento",         "③", "Talento (pipeline)", "Operadores capacitables y técnicos formados cercanos; universidades; capacitación; certificación vendor.", {
        0:   "Tech exótica >12 sem capacitación; sin pipeline universitario.",
        25:  "1-2 universidades cercanas marginales; capacitación 6-12 sem.",
        50:  "3-5 universidades (UPIITA, ESIME, Tec); capacitación 3-6 sem.",
        75:  "5+ universidades cercanas; capacitación 2-4 sem; programa vendor maduro.",
        100: "Pool laboral abundante (operador manual estándar); capacitación 1-2 sem."
    }),
    ("economica",       "④", "Económica",       "CAPEX + integración + OPEX (consumibles, energía, labor) + payback vs benchmark industria.", {
        0:   "CAPEX prohibitivo >USD 5M; payback >5 años.",
        25:  "CAPEX alto USD 1-5M; payback 3-5 años.",
        50:  "CAPEX medio USD 200K-1M; payback 2-3 años.",
        75:  "CAPEX optimizado <USD 500K; payback 1-2 años; bundled OPEX.",
        100: "CAPEX cero (status quo); OPEX humano alto pero conocido."
    }),
    ("riesgo",          "⑤", "Riesgo",          "Estabilidad vendor; madurez tech; single-source; vendor lock-in; refacciones críticas.", {
        0:   "Vendor en transición crítica; tech prototipo <2 años; single-source partes críticas.",
        25:  "Vendor inestable; tech <3 años; single-source componentes.",
        50:  "Vendor estable; tech madura 5+ años; 2-3 fuentes componentes.",
        75:  "Vendor blue-chip; tech establecida 10+ años; 3+ fuentes.",
        100: "Tech estándar industria; arquitectura agnóstica; cero lock-in."
    }),
    ("sustentabilidad", "⑥", "Sustentabilidad", "Material reciclable; reducción de material vs status quo; DIM weight; energía/unidad.", {
        0:   "Material no reciclable; sobre-empaque >30%; energía alta.",
        25:  "Material parcialmente reciclable; sin reducción DIM weight.",
        50:  "Material mayormente reciclable; reducción DIM <10%.",
        75:  "Material reciclable + reducción 20-30% consumo + DIM 20-25%.",
        100: "Mono-material reciclable + reducción >50% + tracking carbono."
    }),
]

CRITERIA_KEYS = [c[0] for c in CRITERIA]
CRITERIA_LABELS = [c[2] for c in CRITERIA]
CRITERIA_SHORT = ["Fit", "Sop", "Tal", "Eco", "Rie", "Sus"]

DEFAULT_WEIGHTS_RATIONALE = (
    "Fit + Económica como co-masters (25% c/u): decisión técnico-financiera principal. "
    "Soporte MX (20%) refleja peso del sustento operativo post-instalación. "
    "Riesgo (15%) calibra vendor instability y lock-in. "
    "Talento default 10% asume tech madura — subir a 20%+ si requiere integración custom alta. "
    "Sustentabilidad (5%) como tie-breaker (relevante ESG, no determinante core)."
)

VENDORS = {
    "Packsize X-Series (auto-box)": {
        "color": "#FFE600",
        "evidence": {
            "fit_operativo":   (85, "Cajas erected 100×100×30 mm hasta 600×400×400 mm; throughput X4 ~600 cajas/hr, X5 ~900, X6 ~1200; APIs REST + Z-Fold mono-material.",
                                [("packsize.com/x-series", "https://www.packsize.com/product/x-series")]),
            "soporte_mx":      (90, "packsize.mx activo; case Autopartes Tracto verificable; Expo Pack GDL 2025 + MX 2026. Sin dirección oficina MX específica identificada → no llega a 100.",
                                [("packsize.mx", "https://www.packsize.mx/"), ("case Autopartes Tracto", "https://www.packsize.com/case-study/autopartes-tracto-de-mexico")]),
            "talento":         (70, "Universidades cercanas Tepotzotlán/Cuautitlán: UPIITA-IPN, ESIME Azcapotzalco, Tec CCM/Edo. Mex.; capacitación vendor ~2 sem. Programa formal vendor no confirmado.",
                                [("UPIITA-IPN", "https://www.upiita.ipn.mx/oferta-educativa/mecatronica")]),
            "economica":       (75, "CAPEX USD 250-700K + consumibles Z-Fold ~USD 0.10-0.20/caja; modelo PIM bundled máquina + material + soporte; payback industria 1-2 años (Modula 2025).",
                                [("Modula warehouse ROI", "https://modula.us/blog/automated-warehouse-systems-improve-space-labor-roi/")]),
            "riesgo":          (70, "Vendor privado estable; adquirió Sparck mayo 2025 → transition risk ~12 meses post-cierre. Z-Fold corrugado: múltiples proveedores (no single-source).",
                                [("Press release adquisición", "https://www.packsize.com/press-release/packsize-to-acquire-sparck-technologies")]),
            "sustentabilidad": (90, "Auto-box reduce material 20-30% vs caja fija; Z-Fold mono-material reciclable; reducción DIM weight 25-30% (UPS/FedEx divisor 5000-6000).",
                                []),
        },
    },
    "Sparck CMC Pakto (auto-wrap)": {
        "color": "#FCD400",
        "evidence": {
            "fit_operativo":   (80, "CartonWrap envuelve items rígidos con cartón fanfold; throughput 600-1000 wraps/hr; CVP Impack 500/hr; menos right-sizing dimensional vs erected box pero mayor velocidad rígidos.",
                                [("Sparck Technologies", "https://sparcktechnologies.com/")]),
            "soporte_mx":      (70, "Canal MX en transición a Packsize MX post-merger mayo 2025; SLA legacy europeo incierto durante integración 12 meses; eventual heredo de canal Packsize MX.",
                                [("Adquisición Packsize-Sparck", "https://interactanalysis.com/insight/packsize-to-acquire-sparck-technologies/")]),
            "talento":         (60, "Mismo pool universitario Packsize; capacitación 3-5 sem (tech wrap más reciente); programa certificación en consolidación post-merger.",
                                []),
            "economica":       (70, "CAPEX USD 200-500K; payback 1-2 años; OPEX consumibles cartón fanfold ~USD 0.08-0.15/wrap; competidor directo a Packsize categoría auto-pack.",
                                []),
            "riesgo":          (55, "Transición corporativa activa Packsize-Sparck; soporte y refacciones inciertos 12 meses post-merger; recomendación: cláusula contractual SLA bridging.",
                                [("Interact Analysis merger coverage", "https://interactanalysis.com/insight/packsize-to-acquire-sparck-technologies/")]),
            "sustentabilidad": (88, "Wrap a medida reduce material 15-25%; fanfold mono-material reciclable; DIM weight reducido 20-25%.",
                                []),
        },
    },
    "Sealed Air AutoBag 850HB (bagging)": {
        "color": "#2D3277",
        "evidence": {
            "fit_operativo":   (75, "AutoBag 850HB hybrid paper+poly hasta 700 órdenes/hr; fit soft goods, apparel, items pequeños; menos right-sizing dimensional (bolsa, no caja).",
                                [("PR AutoBag 850HB launch", "https://www.prnewswire.com/news-releases/sealed-air-launches-autobag-brand-850hb-hybrid-bagging-machine-for-paper-and-poly-mailers-302556908.html")]),
            "soporte_mx":      (95, "DOS plantas MX: Toluca (Edo. Mex.) + Monterrey (NL). Entidad legal Sealed Air de México Operations S de RL de CV (Tlalnepantla). Distribuyen a todo MX + Centroamérica. Toluca a ~40 km del corredor Tepotzotlán/Cuautitlán. Cadena consumibles 25+ años.",
                                [("Sealed Air LA", "https://www.sealedair.com/la/company/our-company/who-we-are"), ("Cobertura T21 MX", "https://t21.com.mx/automatizacion-punto-clave-en-la-operatividad-de-sealed-air/")]),
            "talento":         (80, "Tech bagging madura 30+ años; capacitación operador 1-2 sem; certificación Sealed Air estándar; pool abundante (operador packaging estándar) + presencia industrial MX.",
                                []),
            "economica":       (85, "CAPEX AutoBag 850HB ~USD 150-400K (más bajo que auto-box); OPEX consumibles paper+poly ~USD 0.05-0.12/bolsa; payback 1-2 años; consumibles manufacturados localmente MX.",
                                [("Modula warehouse ROI", "https://modula.us/blog/automated-warehouse-systems-improve-space-labor-roi/")]),
            "riesgo":          (85, "Vendor blue-chip público NYSE:SEE — 65+ años en mercado; tech bagging madura 30+ años; 2 plantas MX = redundancia geográfica; sin transiciones corporativas LATAM-críticas recientes.",
                                []),
            "sustentabilidad": (65, "AutoBag 850HB permite paper mailers (reciclable) + poly (menos); reducción material moderada vs caja fija; residual no siempre mono-stream reciclable; menor performance ESG vs auto-box cartón.",
                                []),
        },
    },
    "Quadient CVP Impack (auto-box)": {
        "color": "#3483FA",
        "evidence": {
            "fit_operativo":   (78, "CVP Impack auto-box similar a Packsize X-Series; throughput 600-1000 boxes/hr; integraciones WMS APIs estándar; OPC UA nativo no confirmado en docs públicos.",
                                [("Quadient CVP", "https://www.materialhandling247.com/product/cvp_500_automated_packing_system/")]),
            "soporte_mx":      (50, "Sin oficina MX identificada en búsqueda; canal global vía quadient.com; partner LATAM autorizado por confirmar vía RFI directo. Recomendación: short-list solo tras validar canal soporte.",
                                [("quadient.com contact", "https://www.quadient.com/en/contact-us")]),
            "talento":         (60, "Capacitación vendor ~3 sem; tech madura pero presencia LATAM débil reduce density de técnicos certificados localmente; pipeline universitario no específico.",
                                []),
            "economica":       (72, "CAPEX USD 300-800K; payback 1-2 años; consumibles propietarios (lock-in moderado).",
                                []),
            "riesgo":          (75, "Vendor estable público Euronext:QDT; tech madura; pero consolidación Packsize+Sparck mayo 2025 lo deja como independiente solo → riesgo presión competitiva 24 meses.",
                                []),
            "sustentabilidad": (85, "Auto-box right-sizing reduce material 20-25%; cartón fanfold reciclable estándar; performance similar a Packsize.",
                                []),
        },
    },
    "Status quo manual + Smurfit": {
        "color": "#8A8AA0",
        "evidence": {
            "fit_operativo":   (50, "Operador humano cubre cualquier forma de SKU (cero limitación técnica); throughput limitado por velocidad humana ~30-60 paquetes/hr; sin integración WMS automatizada.",
                                []),
            "soporte_mx":      (100, "Smurfit Westrock múltiples plantas MX; operador humano = incumbent con cero ramp-up; soporte HR establecido. NOTA: score 100 NO endorsement — refleja incumbencia. AHP evalúa si DESPLAZAR baseline, no si baseline es buena.",
                                [("Crownpack/Smurfit MX", "https://crownpack.com/store/mexico-city/")]),
            "talento":         (100, "Operador manual estándar = pool más abundante MX (INEGI ENOE); capacitación días no semanas; sin curva técnica. Mismo caveat: 100 = incumbencia, no superioridad.",
                                []),
            "economica":       (60, "CAPEX cero (sin inversión); OPEX alto: labor ~USD 7.5K-12K MXN/mes/operador cargado + over-pack desperdicia material 30-50% + DIM weight elevado eleva costo transporte UPS/FedEx.",
                                []),
            "riesgo":          (85, "Tech estándar industria; cero vendor lock-in; partes commodity. Pero costo de oportunidad de NO automatizar crece exponencialmente con volumen — riesgo competitivo, no operativo.",
                                []),
            "sustentabilidad": (40, "Over-pack 30-50% más material vs caja a medida; DIM weight elevado aumenta huella transporte; cartón reciclable pero desperdicio dimensional es la mayor carga ambiental.",
                                []),
        },
    },
}


SLIDER_IDS = ["ahp-w1", "ahp-w2", "ahp-w3", "ahp-w4", "ahp-w5", "ahp-w6"]


# ============================================================
#  DOM helpers
# ============================================================

def _val(slider_id: str) -> float:
    el = document.querySelector(f"#{slider_id}")
    return float(el.value) if el else 0.0


def _set_text(elem_id: str, text: str):
    el = document.querySelector(f"#{elem_id}")
    if el:
        el.textContent = text


def _el(tag, *, text=None, classes=None, style=None, href=None, target=None):
    e = document.createElement(tag)
    if text is not None:
        e.textContent = text
    if classes:
        e.className = classes
    if style:
        e.setAttribute("style", style)
    if href is not None:
        e.setAttribute("href", href)
    if target is not None:
        e.setAttribute("target", target)
    return e


def _clear(el):
    while el.firstChild:
        el.removeChild(el.firstChild)


# ============================================================
#  Render rúbrica panel (one-time render of criteria + anchors)
# ============================================================

def render_rubric_panel():
    target = document.querySelector("#ahp-rubric-panel")
    if not target:
        return
    _clear(target)
    grid = _el("div", style="display:grid;grid-template-columns:repeat(auto-fit,minmax(330px,1fr));gap:0.8rem;")
    for key, num, label, definition, anchors in CRITERIA:
        card = _el("div", style="background:#FFFFFF;border:1px solid #D8D8E3;border-radius:8px;padding:0.7rem 0.85rem;border-left:3px solid #FFE600;")
        head = _el("div", style="display:flex;align-items:baseline;gap:0.4rem;margin-bottom:0.25em;")
        head.appendChild(_el("span", text=num, style="color:#2D3277;font-weight:700;font-size:1.05rem;"))
        head.appendChild(_el("strong", text=label, style="color:#2D3277;font-size:0.95rem;"))
        card.appendChild(head)
        card.appendChild(_el("div", text=definition, style="font-size:0.78rem;color:#4A4A6A;margin-bottom:0.4em;line-height:1.35;"))
        anchors_list = _el("div", style="display:grid;grid-template-columns:32px 1fr;gap:0.2em 0.5em;font-size:0.74rem;line-height:1.3;")
        for score in [0, 25, 50, 75, 100]:
            anchors_list.appendChild(_el("span", text=str(score), style="font-family:JetBrains Mono,monospace;color:#2D3277;font-weight:600;text-align:right;"))
            anchors_list.appendChild(_el("span", text=anchors[score], style="color:#1A1A2E;"))
        card.appendChild(anchors_list)
        grid.appendChild(card)
    target.appendChild(grid)


# ============================================================
#  Compute + render chart
# ============================================================

def normalize(weights):
    total = sum(weights) or 1.0
    return [w / total for w in weights]


def compute_rankings(weights_norm):
    out = []
    for name, vdata in VENDORS.items():
        scores = [vdata["evidence"][k][0] for k in CRITERIA_KEYS]
        total = sum(w * s for w, s in zip(weights_norm, scores))
        out.append((name, total, scores, vdata["color"], vdata["evidence"]))
    out.sort(key=lambda x: -x[1])
    return out


def render_chart(ranked, weights_norm):
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
    weight_str = " · ".join(
        f"{CRITERIA_SHORT[i]} {weights_norm[i]:.0%}" for i in range(len(CRITERIA))
    )
    layout = {
        "title": {
            "text": f"<b>Ranking AHP — pesos normalizados</b><br><sub>{weight_str}</sub>",
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


# ============================================================
#  Render ranking with evidence per vendor per criterion (inline)
# ============================================================

def render_ranking_with_evidence(ranked):
    target = document.querySelector("#ahp-ranking")
    if not target:
        return
    _clear(target)
    for rank, (name, total, scores, color, evidence) in enumerate(ranked, start=1):
        block = _el("div", style=f"background:#FFFFFF;border:1px solid #D8D8E3;border-left:4px solid {color};border-radius:0 8px 8px 0;padding:0.9rem 1.1rem;margin:0.6rem 0;box-shadow:0 1px 3px rgba(45,50,119,0.05);")
        header = _el("div", style="display:flex;justify-content:space-between;align-items:baseline;gap:1rem;margin-bottom:0.5em;padding-bottom:0.4em;border-bottom:1px solid #EFEFF5;")
        title_box = _el("div")
        title_box.appendChild(_el("strong", text=f"#{rank} {name}", style="font-size:1rem;color:#1F2562;"))
        if rank == 1:
            tag = _el("span", text=" winner", classes="brand-tag brand-tag--yellow", style="font-size:0.65rem;margin-left:0.5em;")
            title_box.appendChild(tag)
        score_span = _el("span", text=f"{total:.1f}", style="font-family:JetBrains Mono,monospace;font-weight:700;color:#2D3277;font-size:1.05rem;")
        header.appendChild(title_box)
        header.appendChild(score_span)
        block.appendChild(header)

        # Evidence grid: 6 criteria
        evidence_grid = _el("div", style="display:grid;grid-template-columns:1fr;gap:0.45em;")
        for key, num, label, _def, _anchors in CRITERIA:
            score_v, ev_text, ev_sources = evidence[key]
            row = _el("div", style="font-size:0.82rem;line-height:1.45;padding:0.3em 0;border-bottom:1px dashed #EFEFF5;")
            row_head = _el("div", style="display:flex;justify-content:space-between;align-items:baseline;gap:0.5em;margin-bottom:0.15em;")
            row_head.appendChild(_el("span", text=f"{num} {label}", style="color:#2D3277;font-weight:600;font-size:0.78rem;"))
            row_head.appendChild(_el("span", text=str(score_v), style="font-family:JetBrains Mono,monospace;color:#2D3277;font-weight:700;font-size:0.85rem;"))
            row.appendChild(row_head)
            row.appendChild(_el("div", text=ev_text, style="color:#1A1A2E;font-size:0.78rem;"))
            if ev_sources:
                src_line = _el("div", style="margin-top:0.2em;font-size:0.72rem;color:#8A8AA0;")
                src_line.appendChild(_el("span", text="Fuentes: "))
                for i, (label_s, url_s) in enumerate(ev_sources):
                    a = _el("a", text=label_s, href=url_s, target="_blank", style="color:#3483FA;text-decoration:underline;text-decoration-thickness:1px;")
                    src_line.appendChild(a)
                    if i < len(ev_sources) - 1:
                        src_line.appendChild(_el("span", text=" · "))
                row.appendChild(src_line)
            evidence_grid.appendChild(row)
        block.appendChild(evidence_grid)
        target.appendChild(block)

    # Status quo note (rendered at the end)
    note = _el("div", style="margin-top:1rem;padding:0.7rem 1rem;background:#FFF7B8;border-left:3px solid #FCD400;border-radius:0 6px 6px 0;font-size:0.82rem;color:#1A1A2E;line-height:1.45;")
    note.appendChild(_el("strong", text="Sobre los scores del Status quo: ", style="color:#1F2562;"))
    note.appendChild(_el("span", text="los 100 en Soporte MX y Talento NO son endorsement del proceso manual — reflejan la ventaja de incumbencia (cero ramp-up, pool laboral abundante). El AHP evalúa si vale la pena DESPLAZAR la baseline, no si la baseline es buena en términos absolutos."))
    target.appendChild(note)


# ============================================================
#  Main render (called on every slider change)
# ============================================================

def render(*_args):
    raw = [_val(s) for s in SLIDER_IDS]
    for i, s in enumerate(SLIDER_IDS):
        _set_text(f"{s}-val", f"{int(raw[i])}%")

    norm = normalize(raw)
    ranked = compute_rankings(norm)
    render_chart(ranked, norm)
    render_ranking_with_evidence(ranked)


# ============================================================
#  Init
# ============================================================

render_rubric_panel()

_proxy = create_proxy(render)
for sid in SLIDER_IDS:
    el = document.querySelector(f"#{sid}")
    if el:
        el.addEventListener("input", _proxy)

render()

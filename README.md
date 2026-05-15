# Business Case · Innovación &amp; Robótica · Mercado Envíos MX

Work Sample por Andrés Islas Bravo · mayo 2026

## Qué es esto

Un deck HTML interactivo con widgets Python embebidos (PyScript) en lugar del PowerPoint tradicional. Los evaluadores pueden manipular supuestos en vivo y ver cómo cambian las decisiones en cada caso.

## Cómo abrirlo

**Opción 1 — URL pública (recomendado):** abrir el link público enviado por email. Carga directa desde Vercel.

**Opción 2 — Local (offline):** descomprimir el ZIP en cualquier carpeta, abrir `index.html` en Chrome o Firefox (versiones 90+ recomendadas). Si la primera carga tarda 5–10 segundos es por el bootstrap inicial de PyScript; las navegaciones posteriores son inmediatas.

**Opción 3 — PDF de respaldo:** en `exports/` hay versiones PDF del deck y de las motivacionales por si los widgets interactivos no cargan en el entorno del evaluador.

## Navegación

- `←` `→` mover entre slides
- `T` abrir índice
- `P` imprimir / guardar PDF
- `H` ir a la portada
- Click sobre items del índice para saltar directo

## Estructura

```
deliverable/
├── index.html              ← deck principal (9 slides single-page)
├── README.md               ← este archivo
├── css/                    ← tema institucional MELI + layout + print
├── js/                     ← navegación slide-to-slide
├── python/                 ← 4 widgets PyScript (uno por caso interactivo)
├── data/                   ← JSON de research (proveedores, universidades, benchmarks)
├── appendix/               ← páginas adicionales (roadmap 18m bonus, supuestos, conceptos, proveedores)
└── exports/                ← PDFs de respaldo
```

## Tecnología

- **HTML / CSS / vanilla JavaScript** — sin frameworks pesados.
- **PyScript** (Python 3 vía Pyodide) — cálculos y widgets ejecutándose en el navegador.
- **Plotly.js** — gráficas interactivas.
- **Inter + JetBrains Mono** — tipografía via Google Fonts (con fallback de sistema).
- **Colores institucionales MELI** — amarillo #FFE600 + azul #2D3277.

## Contenido del deck

| # | Slide | Widget Python |
|---|---|---|
| 00 | Portada | — |
| 01 | Marco del director (5 lentes) | — |
| 02 | Caso 1 · Empaquetado | ✓ AHP scorecard editable |
| 03 | Caso 2 · Shelves-to-Person | ✓ Bottleneck simulator |
| 04 | Caso 3 · Picking robótico | ✓ Productivity sim + tornado |
| 05 | Caso 4 · Plan data-driven | — (SQL skeletons) |
| 06 | Caso 5 · Roadmap Gantt | ✓ CPM + Monte Carlo |
| 07 | Motivacionales (resumen) | — |
| 99 | Cierre + appendix | — |

## Honestidad sobre la investigación

Cinco research agents fueron lanzados originalmente para investigar 20+ proveedores en MX en paralelo, pero los permisos de WebSearch/WebFetch en su entorno fueron denegados. La investigación se realizó alternativamente desde la sesión principal (que sí tenía esos permisos), priorizando los 5–7 proveedores más relevantes por categoría con fuentes verificables (todos los `source_link` en los JSON resuelven a páginas reales).

Los datos no verificables están declarados en `appendix/supuestos.html` con razonamiento y nivel de confianza.

## Contacto

- **Andrés Islas Bravo** — andresislas2107@gmail.com
- Para preguntas sobre el deliverable: responder al email.

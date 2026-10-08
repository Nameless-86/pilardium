---
tags: [project, app]
aliases: [Pilardium, Lombardía app]
---

Pilardium is a web app for a milanesa business. It joins cost modelling and stock control.

The backend uses FastAPI. The frontend is one HTML file in `static/index.html`. The tabs are Resumen, Insumos, Milanesas, Recetas, and Registro.

The app reads recipes from [[cost-workbook]]. It keeps stock in [[stock-state]]. Prices flow through [[insumo-price-sync]].

Run it with `./run.sh`. The first version came from the session on [[2026-10-06]].

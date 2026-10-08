---
tags: [excel, data]
aliases: [Costos.xlsx, Estructura de Costos]
---

The workbook `data/Costos.xlsx` holds the cost model. It is the source of truth for the [[pilardium-app]].

Sheets (see also [[recipe-management]]): Costos Fijos, Capital Aportado, Costos Generales, Recetas Pollo, Recetas Vaca, Recetas Cerdo, Precios y Márgenes, Punto de Equilibrio.

Each recipe sheet has five blocks of 29 rows. Each block has 16 ingredient rows. Quantities are for 1 kg of raw meat.

The app writes with openpyxl. It sets `fullCalcOnLoad`, so Excel recalculates formulas on open. The app makes a backup before each write.

Costs reach the workbook through [[insumo-price-sync]].

---
tags: [decision, pricing]
aliases: [price sync]
---

Decision: one insumo has one price. This price lives in [[stock-state]].

When the price changes, the app writes the cost into column D of each linked recipe row in [[cost-workbook]]. It converts units, for example $6000 per kg becomes 6 per g.

If the units do not match, the app leaves the cell alone and marks the row. A cucharada counts as 15 ml.

The [[pilardium-app]] runs this sync after each purchase with "update price", after a recipe save, and after an insumo edit.

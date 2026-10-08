---
tags: [data, stock]
aliases: [state.json]
---

`data/state.json` holds the stock of the [[pilardium-app]]. It has the insumos, the milanesa quantities, the name links, and the movement log.

Each movement stores its deltas. Undo reverses the deltas. Undo fails if the stock is already used.

The first load creates 28 insumos from the ingredient names in [[cost-workbook]]. Some spellings merge, for example "Pechuga" and "Pechuga de pollo".

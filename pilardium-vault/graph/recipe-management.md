---
tags: [decision, ui]
aliases: [recipe catalog]
---

The Recetas tab of the [[pilardium-app]] manages which recipes exist. The Milanesas tab only manages stock.

A recipe id is the meat name and the start row of its block, for example `pollo-150`. A rename keeps the id, so the stock stays linked in [[stock-state]].

Create: the app reuses an empty block, or adds a block at the end of the sheet in [[cost-workbook]]. Remove: the app turns the block into an empty `Indefinido` block. It refuses when the recipe has stock.

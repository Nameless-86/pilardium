# Pilardium

Pilardium joins two tools in one web app: the cost workbook and the stock control.

## How it works

The workbook `data/Costos.xlsx` is the source of truth for recipes, prices, fixed costs, and capital.
The app reads the workbook on each request. If you edit the workbook by hand, the app shows the change.

The app keeps the stock in `data/state.json`. This file holds the insumo quantities, the milanesa quantities, and the movement log.

An insumo has one price. When the price changes, the app writes the new cost into each recipe row that uses the insumo. The app converts units, for example kg to g.

## Run

1. Run `./run.sh`.
2. Open `http://127.0.0.1:8000`.

To use the app from a phone on the same network, run `HOST=0.0.0.0 ./run.sh`. The app has no login. Use it only on a trusted network.

To use another workbook, set `PILARDIUM_XLSX=/path/to/file.xlsx`.

## Tabs

- **Resumen**: fixed costs, break-even point, sales, capital, and alerts. You can edit the fixed costs here.
- **Insumos**: stock of raw material. You can buy, adjust, and edit each insumo.
- **Milanesas**: finished stock only. Choose the meat (Pollo, Vaca, Cerdo), then the recipe, to produce or sell.
- **Recetas**: the list of all recipes. Create a new recipe, edit one, or remove one. The app saves each change in the workbook.
- **Registro**: movement log with an undo button.

## Rules for the workbook

- Close Excel before you save from the app. Excel locks the file, and the app shows an error.
- The app makes a backup in `data/backups/` before it writes. It keeps the last 40 backups.
- The app does not calculate formulas in the file. Excel recalculates them when you open the file.
- A recipe block has 16 ingredient rows.
- A new recipe uses an empty block (named `Indefinido`) if one exists. If not, the app adds a block at the end of the sheet and a row in `Precios y Márgenes`.
- Remove a recipe to turn its block back into an empty `Indefinido` block. The app refuses if the recipe has stock.
# pilardium

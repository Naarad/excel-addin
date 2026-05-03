# Thoughtrons Excel Add-in — V1.0

A Microsoft 365 Excel add-in that adds a **Thoughtrons** ribbon tab with a block library, classification stamps, and quick-insert buttons.

```
Excel ribbon
└── Thoughtrons
    ├── Block Library    ┊ [Big button] — opens a side pane with all 11 blocks
    ├── Quick Insert     ┊ Brand Header · KPI Strip · RAG Dashboard · Risk Block · Action Items
    └── Classification   ┊ UNCLASSIFIED · INTERNAL · RESTRICTED · CONFIDENTIAL
```

**11 blocks across 3 categories:**

- **Headers** — Brand header (eyebrow + title + subtitle) · KPI strip (4 big-number cells)
- **Tables** — RAG dashboard · Risk block · Action items · RACI matrix · Pipeline block
- **Classification** — UNCLASSIFIED / INTERNAL / RESTRICTED / CONFIDENTIAL row at top of sheet

Each table block ships with: dark-fill mono header row, RAG/status dropdowns where appropriate, score and weighted-value formulas, brand-consistent typography.

## Architecture

Blocks are inserted programmatically at the active cell via the Excel JS API:

- `src/blocks.js` — defines each block as a JS function that uses `worksheet.getRange()`, applies brand styling, adds dropdowns, and writes formulas.
- `src/commands.js` — wires 9 ribbon buttons (5 quick-insert + 4 classification) to block functions.
- `src/taskpane.html` — the side pane library.
- `manifest.xml` — declares the ribbon and resources.

To add a new block: add it to `BLOCKS` in `blocks.js`, then add it to `LIBRARY` in `taskpane.html`.

## Deployment

Same as Word and PowerPoint. Find/replace `https://thoughtrons.github.io/excel-addin` with your hosting URL, upload manifest via M365 admin centre. See `Brand_System_v1/README_DEPLOYMENT.md`.

## Files

```
excel_addin/
├── manifest.xml
├── README.md
├── src/
│   ├── commands.html / commands.js
│   ├── blocks.js
│   └── taskpane.html
└── assets/icons/   ← 18 PNGs
```

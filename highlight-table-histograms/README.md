# Highlight Table + Marginal Histograms — Tableau Viz Extension

A custom chart **type** that renders inside a Tableau worksheet: a
color-encoded **highlight table** (rows × columns) with two marginal
histograms — **column totals** across the top and **row totals** down the
right. Everything visual is editable live from a gear-driven **settings
modal**.

> **Requires Tableau 2024.2+** (Viz Extensions API). Runs entirely in the
> worksheet — no servers, no API keys.

![layout](preview.html)

## Encodings (the Marks-card tiles)

| Tile | Role | Example |
|---|---|---|
| **Row** | matrix rows + right histogram | Category, Sub-Category |
| **Column** | matrix columns + top histogram | Region |
| **Value** | cell color + both totals | Sales (a single measure) |

**Row and Column accept several pills.** Drop e.g. *Category* then
*Sub-Category* on **Row** and the headers render **nested** (the outer field is
merged/grouped over its inner values), exactly like a native Tableau highlight
table. A row/column is the unique combination of its fields' values.

The extension pivots the worksheet's summary data into a matrix, colors each
cell by its value (single-hue sequential scale), and sums each row/column for
the histograms.

## Quick start (local, ~3 min)

1. **Serve the repo root on port 1234** (the folder that contains `lib/`, *not*
   this subfolder — the extension loads the shared API from `../lib`):

   ```bash
   # from the repo root (the clone of tableau-viz-extensions)
   python -m http.server 1234
   # or:  npx http-server -p 1234 -c-1
   ```

   This extension is then live at
   `http://localhost:1234/highlight-table-histograms/index.html` (the `<url>` in
   `manifest.trex`).

2. **Open a worksheet** (Superstore works well). Put at least one mark in play.

3. **Add the extension.** Marks card → mark-type dropdown → **Add Extension** →
   **Access Local Extensions** → choose `manifest.trex` from this folder.

4. **Drop fields** on the three tiles: a dimension on **Row**, a dimension on
   **Column**, and a measure on **Value**.

5. **Configure** with the **⚙ gear** (top-right) **or the "Format Extension"
   button** on the Marks card: colors, sizes, labels, histograms, number format.
   Changes apply instantly and persist with the workbook. (Both open the same
   dialog — the Format Extension button is wired via the Extensions API
   `configure` callback.)

> **Reload after each code edit:** mark-type dropdown → the extension's menu →
> **Reload**.

## Preview without Tableau

Open **`preview.html`** in a browser (or serve it). It stubs the Extensions API
with sample *Year × Month × Sales* data so you can iterate on the chart and the
config modal with no Tableau needed. `index.html` is what Tableau actually
loads.

## Interactivity (native tooltips + selection)

When **Tableau tooltips + selection** is on (General section, default on), every
mark is wired to the Extensions API:

- **Hover a cell** → the **native Tableau tooltip** appears
  (`Worksheet.hoverTupleAsync`), anchored at the cursor.
- **Hover a histogram bar** → a small **custom tooltip** shows that row/column's
  **aggregate total** (e.g. `Region: West` / `Sales: 739,814`). Bars are
  aggregates of many marks, which the per-mark native tooltip can't represent,
  so a custom one is used.
- **Click** → selects the underlying marks (`Worksheet.selectTuplesAsync`),
  fires any dashboard selection/filter actions, and **dims the unselected
  marks** in the chart so the selection is visible. Click the same mark again
  to **deselect** (clears the selection and the dimming). **Ctrl/Cmd-click**
  adds/removes from the selection instead of replacing it; clicking empty space
  clears it.

Cells map to a single mark; a **top bar** selects every mark in that column and
a **right bar** every mark in that row — which **cross-highlights** the whole
column/row (its cells stay bright, everything else dims). The **row and column
labels** are interactive too: clicking a label selects/cross-highlights its
whole row/column (re-click to deselect), and hovering shows the row/column
total. (Marks ↔ tuple ids are 1-based summary rows.)

The selection also stays in sync with the worksheet: it's re-read via
`getSelectedMarksAsync` on data changes and on `MarkSelectionChanged`, so
highlights driven from **other sheets in a dashboard** are reflected here too.

## What the config modal controls

- **General** — title, axis label, background, font, tooltips + selection.
- **Color scale** — low/high colors, reverse, emphasis (gamma).
- **Matrix cells** — gap, corner radius, in-cell value labels + size.
- **Labels** — show/hide + size for row/column labels, row-label width.
- **Top histogram** — show/hide, height, value labels.
- **Right histogram** — show/hide, width, value labels.
- **Histograms (shared)** — color bars by intensity, gap from matrix.
- **Number format** — prefix/suffix, decimals, unit (Auto / K / M / B / None),
  thousands separator.

Config is saved to the workbook (Extensions settings) with a `localStorage`
fallback. **Reset** restores defaults.

## Hosting on GitHub Pages

This extension is published at
**https://github.com/difemaro/tableau-viz-extensions**.

1. Push the repo and enable **Settings → Pages → Source: Deploy from a branch →
   `main` / root**. The site builds at
   `https://difemaro.github.io/tableau-viz-extensions/`.
2. The extension is then live at
   `https://difemaro.github.io/tableau-viz-extensions/highlight-table-histograms/index.html`
   — already set as the `<url>` in **`manifest.hosted.trex`**.
3. In Tableau, add the extension with **`manifest.hosted.trex`** (the HTTPS
   build), not `manifest.trex` (which stays on `localhost` for local dev).

> The shared API library lives at the **repo root** (`/lib`), and every
> extension's `index.html` loads it via `../lib/…`. On Pages that resolves to
> `https://difemaro.github.io/tableau-viz-extensions/lib/…`, so it just works —
> as long as you keep the repo's folder layout (don't move the extension out
> from under the root that holds `lib/`).

## Files

```
tableau-viz-extensions/            # repo root (serve this for local dev)
├── lib/tableau.extensions.1.latest.js   # SHARED Extensions API (one copy for all extensions)
└── highlight-table-histograms/
    ├── index.html            # #viz host + config-modal markup (loads ../lib/…)
    ├── preview.html          # standalone browser preview with mock data
    ├── manifest.trex         # LOCAL — encodings (Row/Column/Value), localhost URL
    ├── manifest.hosted.trex  # HOSTED — same, with the GitHub Pages URL
    └── src/
        ├── viz.js            # connect → encodings → pivot → render + config modal
        └── style.css         # grid layout, cells, histograms, modal
```

The Extensions API is **not** copied into this folder — `index.html` loads it
from the shared `../lib/`, so you serve the **repo root** (the folder that
contains `lib/`).

The render logic lives in the **BUILD ZONE** in `src/viz.js`
(`render(model)` and its `buildMatrixGrid` / `buildTopHist` / `buildRightHist`
helpers).

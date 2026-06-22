# Stream Graph — Tableau Viz Extension

A **streamgraph** (a.k.a. ThemeRiver): a stacked area chart whose layers flow
around a wiggling central baseline, so each series reads as an organic "stream"
that swells and shrinks over an ordered axis. Smooth curves, categorical color,
and right-edge series labels.

Requires **Tableau 2024.2+** (API 1.12). Everything runs locally in the
worksheet — no servers, no API keys.

![preview](./preview.html)

## Encodings (Marks card tiles)

| Tile | Role | Pills | Drives |
|---|---|---|---|
| **Axis** | dimension — **discrete or continuous** | 1 | the horizontal axis (one column per value — use an ordered field like month; a continuous/green date pill works too) |
| **Series** | discrete dimension | 1 | one colored stream per value |
| **Value** | continuous measure | 1 | each stream's thickness |

A continuous (green) date pill on **Axis** is supported; the columns are sorted
chronologically. Discrete pills keep Tableau's own sort order.

Each summary row is one **Axis × Series** cell; the measure is pivoted into a
series × axis matrix and stacked.

## Run it locally

The 2 MB Extensions API library is shared once at the workspace root
(`Viz_Extensions/lib/`), so serve the **workspace root**, not this folder:

```bash
# from C:\Claude_Code\Viz_Extensions
python -m http.server 1234
```

Then in Tableau: **Add Extension → Access Local Extensions →** pick
`stream-graph/manifest.trex` (it points at
`http://localhost:1234/stream-graph/index.html`).

Build a tiny test sheet first (e.g. a date dimension on **Axis**, a category on
**Series**, a measure on **Value**) so there's data to read.

## Iterate without Tableau

Open `preview.html` directly in a browser (or headless Chrome) — it mocks the
Tableau API with sample Fruit × Month production data so you can work on the
chart and the config modal with zero Tableau round-trips.

```bash
chrome --headless=new --screenshot=out.png --window-size=600,560 file:///…/stream-graph/preview.html
```

## Customizing

Open the **⚙ gear** (top-right) or the **Format Extension** button on the Marks
card. Everything is driven by one `config` object and a declarative schema, so
every option applies live and persists into the workbook. Highlights:

- **General** — title, subtitle, caption (footer), background, text/muted colors,
  font, tooltips + selection on/off.
- **Stream shape** — baseline offset (**Wiggle** streamgraph / **Silhouette**
  centered / **Zero** stacked / **Expand** 100%), stacking order (inside-out /
  total ↑↓ / reverse / data order), curve (smooth / linear / step), curviness.
- **Bands** — opacity, separator stroke (color + width).
- **Color** — palette (Cool / Tableau 10 / Bold / Pastel / Warm / Mono), or
  sequential (by series total), or a single color. **Colors are tied to each
  series' identity, not its position**, so a category keeps its color when you
  filter the view (no more "Furniture turns blue when filtered"). **Per-series
  overrides:** double-click a legend item to pick its color, or use the **Series
  colors** section of the config modal (each row has a ↺ reset-to-palette). The
  assignment is persisted with the workbook.
- **Axis & grid** — x-axis labels with **offset (staggered) labels** to fix
  overlap (alternates labels onto a 2nd row with connector ticks), collision-aware
  thinning + a post-render measurement pass that guarantees no two labels touch,
  vertical gridlines, optional zero baseline.
- **Legend** — optional swatch+name legend, positioned **top** or **bottom**,
  colored to match the streams. Handy when several streams end at similar values
  and the right-edge labels get crowded. **Interactive:** click a legend item to
  select/highlight that whole series (dimming the rest and filtering the
  dashboard); Ctrl/Cmd-click adds to the selection, click it again to clear. The
  legend dims in sync with the chart selection.
- **Series labels** — right-edge labels colored to match. When several streams
  end at close values the labels are **nudged apart vertically** (each keeps a
  thin leader line back to its stream) instead of overlapping or being hidden;
  only streams thinner than a small threshold at their end are skipped.

## Interactivity

Hover and click resolve to the **single (series, axis) cell under the cursor** —
e.g. one fruit-month — by snapping to the nearest axis column inside the hovered
stream:

- **Hover** → Tableau's **native tooltip** for that one mark (`hoverTupleAsync`),
  so it respects the worksheet's own field formatting. (A whole stream is an
  aggregate of many marks and can't be described by the per-mark tooltip — but a
  single cell can, which is why the interaction is per-cell.)
- **Click** → selects that one fruit-month mark (`selectTuplesAsync`), so when the
  viz drives a dashboard it **filters at fruit-month granularity, not the whole
  series**. The hovered band stays lit and the rest dim. Ctrl/Cmd-click adds to
  the selection; click empty space clears it.
- **Legend item** → single-click selects that *whole* series (all its cells)
  through the same path, so the legend doubles as a series filter/highlighter;
  **double-click** opens a color picker for that series.
- Selection re-syncs on data change and on selections driven from other sheets,
  mapping each selected mark back to its exact cell via the series/axis values.

## Publishing on GitHub Pages

Use `stream_graph_hosted.trex` (its `<url>` points at the Pages path). Make sure
the shared `lib/` is committed at the repo root — `index.html` loads `../lib/…`,
which on Pages resolves to `https://<user>.github.io/<repo>/lib/…`. Tableau
requires HTTPS for hosted extensions (`*.github.io` qualifies).

## Files

```
stream-graph/
  index.html                loads ../lib + src, render target #viz + modal markup
  preview.html              mocked Tableau API for local iteration
  src/viz.js                connect → encodings → data → SVG render + modal + persistence
  src/style.css             page chrome, theme vars, empty state, tooltip, gear + modal
  manifest.trex             LOCAL dev manifest (localhost:1234)
  stream_graph_hosted.trex  HOSTED manifest (GitHub Pages URL)
```

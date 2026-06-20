# Dendrogram Bars — Tableau Viz Extension

A fan-out bar chart: every category bar is tied back to a single **root node**
(the grand total) by a smooth bundled curve, like a one-level dendrogram. Bars
are sorted, colored from a categorical palette, and labelled with name + value.

Requires **Tableau 2024.2+** (API 1.12). Everything runs locally in the
worksheet — no servers, no API keys.

![preview](./preview.html)

## Encodings (Marks card tiles)

| Tile | Role | Pills | Drives |
|---|---|---|---|
| **Category** | discrete dimension | 1 | one bar per value |
| **Value** | continuous measure | 1 | bar length + the root total |

## Run it locally

The 2 MB Extensions API library is shared once at the workspace root
(`Viz_Extensions/lib/`), so serve the **workspace root**, not this folder:

```bash
# from C:\Claude_Code\Viz_Extensions
python -m http.server 1234
```

Then in Tableau: **Add Extension → Access Local Extensions →** pick
`dendrogram-bars/manifest.trex` (it points at
`http://localhost:1234/dendrogram-bars/index.html`).

Build a tiny test sheet first (Superstore: Sub-Category + Sales) so there's data
to read.

## Iterate without Tableau

Open `preview.html` directly in a browser (or headless Chrome) — it mocks the
Tableau API with sample Sub-Category × Sales data so you can work on the chart
and the config modal with zero Tableau round-trips.

```bash
chrome --headless --screenshot=out.png --window-size=720,560 file:///…/dendrogram-bars/preview.html
```

## Customizing

Open the **⚙ gear** (top-right) or the **Format Extension** button on the Marks
card. Everything is driven by one `config` object and a declarative schema, so
every option applies live and persists into the workbook. Highlights:

- **Theme** — background, text/muted colors, font (dark by default).
- **Root total** — caption, prefix (`$`), decimals; the total is auto-abbreviated
  (e.g. `$2.30M`).
- **Bars** — sort (value ↓/↑ or data order), row height, thickness, corner
  radius (pill), min width, opacity.
- **Fan curves** — show/hide, curviness, width, opacity, color (match bar or
  flat).
- **Palette** — Pastel / Tableau 10 / Bold / Mono (single-hue ramp).
- **Labels** — show category names / values, size, label-column width.
- **Number format** — prefix/suffix, decimals, unit (none/auto/K/M/B), thousands
  separator.
- **Axis** — optional bottom value axis with a configurable tick count.

## Interactivity

- **Hover** a bar → Tableau's native tooltip (`hoverTupleAsync`).
- **Click** a bar → selects that mark (`selectTuplesAsync`), dimming the rest and
  driving cross-highlight / dashboard actions. Ctrl/Cmd-click adds to the
  selection; click empty space clears it.
- Selection re-syncs on data change and on selections driven from other sheets.

## Publishing on GitHub Pages

Use `dendrogram_bars_hosted.trex` (its `<url>` points at the Pages path). Make
sure the shared `lib/` is committed at the repo root — `index.html` loads
`../lib/…`, which on Pages resolves to `https://<user>.github.io/<repo>/lib/…`.
Tableau requires HTTPS for hosted extensions (`*.github.io` qualifies).

## Files

```
dendrogram-bars/
  index.html                  loads ../lib + src, render target #viz + modal markup
  preview.html                mocked Tableau API for local iteration
  src/viz.js                  connect → encodings → data → SVG render + modal + persistence
  src/style.css               page chrome, theme vars, empty state, gear + modal
  manifest.trex               LOCAL dev manifest (localhost:1234)
  dendrogram_bars_hosted.trex HOSTED manifest (GitHub Pages URL)
```

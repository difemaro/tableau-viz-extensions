# Rounded Bars — Tableau Viz Extension

Pill-shaped **bars** (horizontal) or **columns** (vertical), one per category, where
each bar's length/height encodes the measure. The signature is the **color**:
instead of a flat fill, every bar is painted with a two-stop **linear gradient**
along (or across) its length — uniform, value-driven, a categorical palette, or a
single color, with optional **per-bar color overrides**.

Requires **Tableau 2024.2+** (API 1.12). Everything runs locally in the
worksheet — no servers, no API keys.

![preview](./preview.html)

## Encodings (Marks card tiles)

| Tile | Role | Pills | Drives |
|---|---|---|---|
| **Category** | dimension (discrete or continuous) | up to 5 | one rounded bar per **unique combination** of the dropped fields |
| **Value** | continuous measure | 1 | each bar's length/height |

Drop several dimensions on **Category** to get one bar per combination, displayed
with **nested, merged labels** like the highlight-table: the outer dimension is
shown once and spans its child bars, the innermost dimension labels each bar.
Bars are ordered hierarchically (groups stay contiguous, ordered by aggregate in
the sort direction; bars within a group by value). Works in both orientations —
horizontal nests as label columns on the left, vertical as merged rows below.
Groups are separated by a configurable **gap + divider line** (see the *Groups*
section of the modal) so it's clear where each one starts and ends. This nests
for **any number of pills**, matching the convention of a nested table: an
outer-level boundary gets a larger gap + a stronger, full-width line; a
deeper-level boundary gets a smaller gap + a fainter line that **indents to start
at that level's label column** (so a Ship-Mode break starts at the Ship-Mode
column, a Category break spans the full width).

## Run it locally

The 2 MB Extensions API library is shared once at the workspace root
(`Viz_Extensions/lib/`), so serve the **workspace root**, not this folder:

```bash
# from C:\Claude_Code\Viz_Extensions
python -m http.server 1234
```

Then in Tableau: **Add Extension → Access Local Extensions →** pick
`rounded-bars/manifest.trex` (it points at
`http://localhost:1234/rounded-bars/index.html`).

## Iterate without Tableau

Open `preview.html` directly in a browser (or headless Chrome) — it mocks the
Tableau API with sample Fund × AUM data so you can work on the chart and the
config modal with zero Tableau round-trips.

```bash
chrome --headless=new --screenshot=out.png --window-size=900,420 file:///…/rounded-bars/preview.html
```

## Customizing

Open the **⚙ gear** (top-right) or the **Format Extension** button on the Marks
card. Everything is driven by one `config` object and a declarative schema, so
every option applies live and persists into the workbook. Highlights:

- **General** — title, subtitle, background, text/muted colors, font, tooltips +
  selection. (Defaults to a dark theme — the gradient pops on dark.)
- **Layout** — orientation (horizontal bars / vertical columns), sort, fit-to-pane
  vs. fixed bar slot (then the pane scrolls), bar thickness, corner radius
  (0 = square … 1 = full pill), min length, and an optional **zero baseline line**.
- **Negative values** — bars are measured from a **zero baseline**: positive bars
  grow one way from zero, negatives the other, with the value label on the bar's
  outer end and a faint zero line (shown only when the data spans both signs).
  All-positive data is unaffected (zero sits at the edge, like a plain bar chart).
- **Color** — the gradient system:
  - **Gradient (uniform)** — every bar uses the same start→end gradient.
  - **Gradient by value** — the gradient brightens with the bar's value (each bar
    distinct). *Value-encoding, so per-bar overrides don't apply here.*
  - **Divergent by value** — each bar a **solid** color interpolated from a
    low→high pair by its value (low/high colors configurable). *Value-encoding.*
  - **Palette (by category)** — categorical solid, colored by category identity
    (stable across re-sort/filter).
  - **Single color.**
  - Gradient **start/end** colors and **direction** (along / across / diagonal).
- **Per-bar overrides** — **double-click any bar** to pick its color, or use the
  **Bar colors** section of the modal (each row has a ↺ reset). Available in every
  mode except *Gradient by value*. Persisted with the workbook.
- **Track** — optional faint full-length track behind each bar.
- **Labels** — category labels (left gutter for bars; under columns choose
  **Rotated**, **Offset rows** (staggered onto two rows so dense labels don't
  overlap), or **Flat**; long names get an ellipsis), value labels (auto
  inside/outside for bars; rotated inside columns), sizes, and placement. The
  **Category gutter** slider works in both orientations — left-gutter width for
  bars, label-band height for columns. In columns the band also **auto-fits** the
  longest rotated label so it doesn't truncate; the slider then caps it.
- **Number format** — prefix/suffix, decimals, unit (none/auto/K/M/B), thousands
  separator.

## Interactivity

- **Hover** a bar → Tableau's native tooltip (`hoverTupleAsync`).
- **Click** a bar → selects that mark (`selectTuplesAsync`), dimming the rest and
  driving cross-highlight / dashboard filters. Ctrl/Cmd-click adds; click empty
  space clears. **Drag** a marquee to select a region.
- **Double-click** a bar → color picker (per-bar override), in the modes that
  allow overrides.
- Selection re-syncs on data change and on selections driven from other sheets.

## Publishing on GitHub Pages

Use `rounded_bars_hosted.trex` (its `<url>` points at the Pages path). Make sure
the shared `lib/` is committed at the repo root — `index.html` loads `../lib/…`,
which on Pages resolves to `https://<user>.github.io/<repo>/lib/…`. Tableau
requires HTTPS for hosted extensions (`*.github.io` qualifies).

## Files

```
rounded-bars/
  index.html                loads ../lib + src, render target #viz + modal markup
  preview.html              mocked Tableau API for local iteration
  src/viz.js                connect → encodings → data → SVG render + modal + persistence
  src/style.css             page chrome, theme vars, empty state, gear + modal
  manifest.trex             LOCAL dev manifest (localhost:1234)
  rounded_bars_hosted.trex  HOSTED manifest (GitHub Pages URL)
```

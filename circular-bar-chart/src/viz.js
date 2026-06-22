'use strict';

/* =============================================================
   CIRCULAR BAR CHART — Tableau Viz Extension
   -------------------------------------------------------------
   One CONCENTRIC RING per category. Each ring has a faint
   full-circle track; a colored arc on top encodes the measure
   as ARC LENGTH (value / max of the sweep). Rings are stacked
   from the outside in (largest value outermost by default), with
   a name + value legend in the top-left. Everything visual is
   driven by `config` and editable live from the gear → modal.

   This differs from the sibling "radial bar chart" (which uses
   angular WEDGES whose RADIUS encodes value). Here every bar is a
   ring and ANGLE encodes value.

   The four moves (see CLAUDE.md):
     1. CONNECT   — initializeAsync(), grab the worksheet
     2. ENCODINGS — Category / Value tiles (manifest.trex)
     3. DATA      — read summary data, aggregate per category
     4. DRAW      — render(model) into #viz; redraw on data change

   Rendered as one inline <svg> (polar geometry is far easier as
   SVG arc paths than DOM/CSS). The plumbing matches the DOM
   charts; only the BUILD ZONE differs.
   ============================================================= */

(function () {
  const host = document.getElementById('viz');
  // Key for the config blob stored in the workbook (tableau.extensions.settings).
  // Keep STABLE so a workbook saved by an earlier build still loads its config.
  // Workbook settings are the single source of truth (no localStorage cache —
  // it would leak the last-used config across instances on one origin and break
  // "fresh viz = defaults"; we never paint before init resolves).
  const STORAGE_KEY = 'circularbar-config-v1';

  const SVG_NS = 'http://www.w3.org/2000/svg';
  const TAU = Math.PI * 2;

  // Categorical palette (Tableau 10-ish) for colorMode === 'palette'.
  const PALETTE = ['#4e79a7', '#f28e2b', '#e15759', '#76b7b2', '#59a14f',
                   '#edc948', '#b07aa1', '#ff9da7', '#9c755f', '#bab0ac'];

  /* ---------- config ---------- */
  const DEFAULT_CONFIG = {
    // identity
    title: '',
    bgColor: '#ffffff',
    fontFamily: 'system',

    // interactivity — native Tableau tooltips on hover + mark selection on click
    enableTooltips: true,

    // color (sequential by value is the cross-extension default — see CLAUDE.md)
    colorMode: 'sequential',  // palette | single | sequential
    barColor: '#1f6fff',      // used when colorMode === 'single'
    lowColor: '#dbe6f3',      // sequential scale endpoints (also drives "by value")
    highColor: '#1f4e79',
    reverseScale: false,

    // geometry
    startAngleDeg: 0,         // 0 = top; arcs grow clockwise
    sweepDeg: 270,           // arc the MAX value fills (360 = full ring)
    innerRadiusPct: 0.16,    // center hole as a fraction of the outer radius
    ringPadding: 0.4,        // gap between rings as a fraction of each ring slot
    roundedCaps: true,       // rounded arc ends

    // track behind each bar
    showTrack: true,
    trackColor: '#e9eaee',

    // ordering of the rings (outer = first)
    sortOrder: 'desc',       // desc | asc | none

    // category + value labels
    showLegend: true,         // master toggle for the name + value labels
    labelPlacement: 'bar-start', // 'bar-start' (beside each bar's start) | 'corner' (top-left legend)

    // value labels drawn at each arc's end
    showValues: false,
    labelSize: 12,

    // number formatting (values + tooltips)
    numPrefix: '',
    numSuffix: '',
    numDecimals: 2,
    numUnit: 'auto',          // auto | K | M | B | none
    numThousands: true,
  };

  const FONT_STACKS = {
    system: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
    serif: 'Georgia, "Times New Roman", serif',
    mono: 'ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace',
  };

  // Start from defaults; authoritative workbook settings load once
  // initializeAsync() resolves (settings API isn't ready before that, and
  // render() runs only after — so there's no pre-init paint to cache for).
  let config = { ...DEFAULT_CONFIG };
  let lastModel = null;

  // 1 — CONNECT ------------------------------------------------
  // Registering a `configure` callback makes Tableau show the "Format
  // Extension" button on the Marks card; clicking it opens our settings modal.
  tableau.extensions.initializeAsync({ configure: openConfigModal }).then(
    () => {
      config = loadConfig();

      const ws = tableau.extensions.worksheetContent.worksheet;
      ws.addEventListener(tableau.TableauEventType.SummaryDataChanged, update);

      // Reflect selection/highlight changes (incl. those driven from other
      // sheets in a dashboard) without rebuilding the whole chart.
      if (tableau.TableauEventType.MarkSelectionChanged) {
        ws.addEventListener(tableau.TableauEventType.MarkSelectionChanged, async () => {
          if (!lastModel) return;
          await syncSelectionFromWorksheet(lastModel);
          applySelectionStyles();
        });
      }

      if (window.ResizeObserver) {
        new ResizeObserver(() => lastModel && render(lastModel)).observe(host);
      }

      wireModal();
      wireInteractivity();
      update(); // first paint
    },
    (err) => showMessage('Could not initialize: ' + escapeHtml(String(err)))
  );

  // 2 + 3 — ENCODINGS + DATA → aggregate per category ----------
  async function update() {
    const ws = tableau.extensions.worksheetContent.worksheet;

    let fields = {};
    try {
      fields = await getEncodedFields(ws);
    } catch (e) {
      console.error('reading encodings failed:', e);
    }

    if (!fields.category || !fields.category.length ||
        !fields.value || !fields.value.length) {
      lastModel = null;
      showEmptyState(fields);
      return;
    }

    let table = { columns: [], data: [] };
    try {
      table = await readSummary(ws);
    } catch (e) {
      console.error('reading summary data failed:', e);
    }

    lastModel = buildModel(ws.name, fields, table);
    await syncSelectionFromWorksheet(lastModel);
    render(lastModel);
  }

  const SEP = String.fromCharCode(0);  // key separator (won't appear in a field value)
  function joinKey(values) { return values.map(String).join(SEP); }

  // Aggregate the flat summary rows into one item per category. Category can
  // hold SEVERAL pills → an item is the unique combination of its values.
  function buildModel(worksheet, fields, table) {
    const cols = table.columns;
    const catFields = fields.category;       // [name, ...] (outer -> inner)
    const valueField = fields.value[0];      // a single measure
    const catIdxs = catFields.map((n) => findColumn(cols, n));
    const valIdx = findValueColumn(cols, valueField);

    const pos = new Map();
    const items = [];           // [{ key, parts, label, value, tupleIds }]
    const tupleKeys = {};       // tupleId → rawKey (to match selected marks)

    table.data.forEach((row, idx) => {
      const tupleId = idx + 1;  // tuple ids are 1-based
      const parts = catIdxs.map((i) => cellText(row[i]));
      const key = parts.join(SEP);
      const rawKey = joinKey(catIdxs.map((i) => (row[i] ? row[i].value : '')));
      const raw = row[valIdx] ? Number(row[valIdx].value) : NaN;
      const v = Number.isFinite(raw) ? raw : null;

      let it = pos.get(key);
      if (!it) {
        it = { key, parts, label: parts.join(' / '), value: null, tupleIds: [] };
        pos.set(key, it);
        items.push(it);
      }
      if (v != null) {
        it.value = (it.value || 0) + v;
        it.tupleIds.push(tupleId);
        tupleKeys[tupleId] = rawKey;
      }
    });

    let min = Infinity, max = -Infinity;
    items.forEach((it) => {
      if (it.value != null) { min = Math.min(min, it.value); max = Math.max(max, it.value); }
    });
    if (!Number.isFinite(min)) { min = 0; max = 0; }

    return { worksheet, catFields, valueField, items, min, max, tupleKeys };
  }

  // Items in ring order (outer first). Nulls sink to the end either way.
  function sortedItems(model) {
    const arr = model.items.slice();
    if (config.sortOrder === 'desc') {
      arr.sort((a, b) => (b.value == null ? -Infinity : b.value) - (a.value == null ? -Infinity : a.value));
    } else if (config.sortOrder === 'asc') {
      arr.sort((a, b) => (a.value == null ? Infinity : a.value) - (b.value == null ? Infinity : b.value));
    }
    return arr;
  }

  /* ===========================================================
     ▼▼▼  BUILD ZONE — render the chart from the model  ▼▼▼
     =========================================================== */
  function render(model) {
    if (!model) return;
    applyTheme();
    host.innerHTML = '';

    if (!model.items.length) {
      showMessage('No data to display yet.');
      return;
    }

    const wrap = el('div', 'cb');
    if (config.title) wrap.appendChild(el('div', 'cb-title', config.title));
    const stage = el('div', 'cb-stage');
    wrap.appendChild(stage);
    host.appendChild(wrap);

    // Measure the stage *after* it's in the DOM so we get real pixels.
    drawSvg(model, stage);

    // Re-apply selection dimming/highlight (the DOM was just rebuilt).
    applySelectionStyles();
  }

  function drawSvg(model, stage) {
    const items = sortedItems(model);
    const n = items.length;
    const W = Math.max(40, stage.clientWidth || 600);
    const H = Math.max(40, stage.clientHeight || 400);
    // Keep the rings CIRCULAR: size to the smaller axis and center (a circular
    // chart must stay aspect-locked — stretching width≠height makes ellipses).
    const cx = W / 2, cy = H / 2;

    const svg = svgEl('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, class: 'cb-svg' });

    const pad = config.roundedCaps ? 10 : 6;     // headroom for round caps + labels
    const outerR = Math.max(8, Math.min(W, H) / 2 - pad);
    const innerR = outerR * clamp(config.innerRadiusPct, 0, 0.85);
    const avail = Math.max(2, outerR - innerR);
    const slot = avail / n;                       // radial space per ring (band + gap)
    const thickness = Math.max(1.5, slot * (1 - clamp(config.ringPadding, 0, 0.95)));

    const start = deg2rad(config.startAngleDeg);
    const maxSweep = deg2rad(clamp(config.sweepDeg, 30, 360));
    const fullRing = config.sweepDeg >= 359.9;
    const cap = config.roundedCaps ? 'round' : 'butt';

    items.forEach((it, i) => {
      const r = outerR - slot * i - slot / 2;     // i = 0 is the outermost ring
      const frac = model.max > 0 && it.value != null ? clamp(it.value / model.max, 0, 1) : 0;
      const a1 = start + frac * maxSweep;
      const color = colorFor(i, n, it.value, model);

      const g = svgEl('g', { class: 'cb-ring' });

      // faint full-(or sweep-)length track behind the bar
      if (config.showTrack) {
        const track = svgEl('path', { d: ringArcPath(cx, cy, r, start, start + maxSweep), class: 'cb-track' });
        track.setAttribute('fill', 'none');
        track.style.stroke = config.trackColor;
        track.style.strokeWidth = thickness;
        track.setAttribute('stroke-linecap', fullRing ? 'butt' : cap);
        g.appendChild(track);
      }

      // the value arc
      if (it.value != null && frac > 0) {
        const full = frac >= 0.9999 && fullRing;
        const bar = svgEl('path', { d: ringArcPath(cx, cy, r, start, a1), class: 'cb-bar' });
        bar.setAttribute('fill', 'none');
        bar.style.stroke = color;
        bar.style.strokeWidth = thickness;
        bar.setAttribute('stroke-linecap', full ? 'butt' : cap);
        g.appendChild(bar);
      }

      // name + value label beside the bar's START point (sits in the open gap
      // left of the start when the sweep is < 360°). Appended INTO the ring's
      // interactive <g> so it hovers/selects/dims together with the arc.
      if (config.showLegend && config.labelPlacement === 'bar-start') {
        const [sx, sy] = pt(cx, cy, r, start);
        const onLeft = sx <= cx + 0.5;            // start near the top → label to the left
        const off = thickness / 2 + 6;            // clear the (possibly rounded) start cap
        const ink = contrastInk(config.bgColor);
        const tx = svgEl('text', {
          x: sx + (onLeft ? -off : off), y: sy,
          'text-anchor': onLeft ? 'end' : 'start', 'dominant-baseline': 'middle',
          class: 'cb-startlabel',
        });
        tx.style.fontSize = config.labelSize + 'px';
        const nm = svgEl('tspan', { class: 'cb-sl-name' });
        nm.textContent = it.label;
        nm.style.fill = ink;
        tx.appendChild(nm);
        if (it.value != null) {
          const vs = svgEl('tspan', { class: 'cb-sl-val', dx: 6 });
          vs.textContent = formatNumber(it.value);
          vs.style.fill = ink;
          tx.appendChild(vs);
        }
        g.appendChild(tx);
      }

      // optional value label at the arc's end
      if (config.showValues && it.value != null && frac > 0.04) {
        const [vx, vy] = pt(cx, cy, r, a1);
        const txt = svgEl('text', {
          x: vx, y: vy, 'text-anchor': 'middle', 'dominant-baseline': 'middle', class: 'cb-value',
        });
        txt.textContent = formatNumber(it.value);
        txt.style.fontSize = Math.max(8, config.labelSize - 2) + 'px';
        txt.style.fill = contrastInk(config.bgColor);
        txt.style.stroke = config.bgColor;   // halo so it stays legible on the arc
        txt.style.strokeWidth = '3px';
        txt.style.paintOrder = 'stroke';
        g.appendChild(txt);
      }

      // Tag as an interactive mark (hover tooltip + click/marquee selection).
      if (config.enableTooltips && it.tupleIds.length) {
        g.classList.add('cb-interactive');
        g._tupleIds = it.tupleIds;
        g._tip = { fields: model.catFields, parts: it.parts, valueField: model.valueField, value: it.value };
      }

      svg.appendChild(g);
    });

    stage.appendChild(svg);

    // Fit the ACTUALLY-DRAWN content (arcs + bar-start labels) to the pane so a
    // partial sweep (e.g. 90° / 270°) doesn't waste the empty quadrant(s) — a
    // 90° fan would otherwise sit tiny in one corner. We scale the content's
    // bounding box UNIFORMLY (aspect-locked → circles never become ellipses) and
    // center it via a viewBox + "meet". getBBox() ignores stroke width, so pad
    // by half a ring's thickness (plus a little for the rounded caps).
    try {
      const bb = svg.getBBox();
      if (bb.width > 0 && bb.height > 0) {
        const fitPad = thickness / 2 + 4;
        const vbW = bb.width + 2 * fitPad;
        const vbH = bb.height + 2 * fitPad;
        svg.setAttribute('viewBox', `${bb.x - fitPad} ${bb.y - fitPad} ${vbW} ${vbH}`);
        svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
        // The viewBox zooms the WHOLE drawing, so a partial sweep (which zooms
        // more) would render larger text than a near-full one. Counter-scale the
        // font size by the fit factor so labels stay a CONSTANT on-screen px
        // size across any sweep, while the arcs still scale to fill. (Arc
        // thickness scaling is intended; only text is held constant.)
        const scale = Math.min(W / vbW, H / vbH);
        if (scale > 0 && isFinite(scale)) {
          svg.querySelectorAll('text').forEach((t) => {
            const px = parseFloat(t.style.fontSize) || config.labelSize;
            t.style.fontSize = (px / scale) + 'px';
          });
        }
      }
    } catch (e) { /* getBBox unavailable — keep the 0 0 W H viewBox */ }

    // top-left name + value legend (overlay), in ring order
    if (config.showLegend && config.labelPlacement === 'corner') buildLegend(model, items, stage);
  }

  // Two aligned columns (name right-aligned, value left-aligned), overlaid in
  // the top-left like the reference. Rows are display:contents so the spans line
  // up in the parent grid; each row carries its tuple ids so it selects/dims in
  // lockstep with its ring.
  function buildLegend(model, items, stage) {
    const leg = el('div', 'cb-legend');
    leg.style.color = contrastInk(config.bgColor);
    leg.style.fontSize = config.labelSize + 'px';
    items.forEach((it) => {
      const row = el('div', 'cb-leg-row');
      row.appendChild(el('span', 'cb-leg-name', it.label));
      row.appendChild(el('span', 'cb-leg-val', formatNumber(it.value)));
      if (config.enableTooltips && it.tupleIds.length) {
        row.classList.add('cb-interactive');
        row._tupleIds = it.tupleIds;
        row._tip = { fields: model.catFields, parts: it.parts, valueField: model.valueField, value: it.value };
      }
      leg.appendChild(row);
    });
    stage.appendChild(leg);
  }
  /* ▲▲▲  END BUILD ZONE  ▲▲▲ */

  /* ---------- polar / SVG geometry ---------- */
  function svgEl(tag, attrs) {
    const node = document.createElementNS(SVG_NS, tag);
    if (attrs) for (const k in attrs) node.setAttribute(k, attrs[k]);
    return node;
  }
  function deg2rad(d) { return (d * Math.PI) / 180; }
  function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }
  // Point on a circle. theta = 0 is straight up; increases CLOCKWISE.
  function pt(cx, cy, r, theta) {
    return [cx + r * Math.sin(theta), cy - r * Math.cos(theta)];
  }
  // An open arc stroke (no fill) along a0→a1 at radius r. A near-complete sweep
  // is drawn as two half-arcs (a single SVG arc can't represent a full 360°).
  function ringArcPath(cx, cy, r, a0, a1) {
    let sweep = a1 - a0;
    if (sweep < 0) sweep = 0;
    if (sweep >= TAU - 1e-3) {
      const [sx, sy] = pt(cx, cy, r, a0);
      const [hx, hy] = pt(cx, cy, r, a0 + Math.PI);
      return `M ${sx} ${sy} A ${r} ${r} 0 1 1 ${hx} ${hy} A ${r} ${r} 0 1 1 ${sx} ${sy}`;
    }
    const large = sweep > Math.PI ? 1 : 0;
    const [x1, y1] = pt(cx, cy, r, a0);
    const [x2, y2] = pt(cx, cy, r, a1);
    return `M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2}`;
  }

  /* ---------- color + number helpers ---------- */
  function colorFor(i, n, value, model) {
    if (config.colorMode === 'single') return config.barColor;
    if (config.colorMode === 'sequential') return scaleColor(norm(value, model.min, model.max));
    return PALETTE[i % PALETTE.length];
  }
  function norm(v, min, max) {
    if (v == null || max <= min) return 0;
    return clamp((v - min) / (max - min), 0, 1);
  }
  function scaleColor(t) {
    const tt = config.reverseScale ? 1 - t : t;
    const a = hexToRgb(config.lowColor);
    const b = hexToRgb(config.highColor);
    const c = a.map((ch, i) => Math.round(ch + (b[i] - ch) * tt));
    return `rgb(${c[0]}, ${c[1]}, ${c[2]})`;
  }
  function hexToRgb(hex) {
    const h = String(hex).replace('#', '');
    const f = h.length === 3 ? h.split('').map((x) => x + x).join('') : h;
    const num = parseInt(f, 16);
    return [(num >> 16) & 255, (num >> 8) & 255, num & 255];
  }
  function contrastInk(color) {
    let rgb;
    if (String(color).startsWith('rgb')) rgb = color.match(/\d+/g).map(Number);
    else rgb = hexToRgb(color);
    const lum = (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255;
    return lum > 0.6 ? '#1c2330' : '#ffffff';
  }

  const UNIT_DIV = { K: 1e3, M: 1e6, B: 1e9 };
  function formatNumber(v) {
    if (v == null || !Number.isFinite(v)) return '';
    let n = v, unit = '';
    if (config.numUnit === 'auto') {
      const abs = Math.abs(n);
      if (abs >= 1e9) { n = n / 1e9; unit = 'B'; }
      else if (abs >= 1e6) { n = n / 1e6; unit = 'M'; }
      else if (abs >= 1e3) { n = n / 1e3; unit = 'K'; }
    } else if (UNIT_DIV[config.numUnit]) {
      n = n / UNIT_DIV[config.numUnit];
      unit = config.numUnit;
    }
    let str = n.toFixed(config.numDecimals);
    if (config.numThousands) str = addThousands(str);
    return config.numPrefix + str + unit + config.numSuffix;
  }
  // Full (non-abbreviated) number for tooltips.
  function formatFull(v) {
    if (v == null || !Number.isFinite(v)) return '';
    let str = Number(v).toFixed(config.numDecimals);
    if (config.numThousands) str = addThousands(str);
    return config.numPrefix + str + config.numSuffix;
  }
  function addThousands(str) {
    const parts = str.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts.join('.');
  }

  /* =========================================================
     INTERACTIVITY — native Tableau tooltips + mark selection
     Each interactive element (ring <g> or legend row) carries the
     tuple ids it represents. Hover → hoverTupleAsync (single mark)
     or a custom DOM tip (aggregate). Click / marquee →
     selectTuplesAsync. Listeners are delegated off #viz.
     ========================================================= */
  const SELECT_SIMPLE = (tableau.SelectOptions && tableau.SelectOptions.Simple) || 'select-options-simple';

  let selectedTuples = new Set();   // drives worksheet selection + our dimming
  let hoverEl = null;
  let hoverAt = 0;

  const MARQUEE_THRESHOLD = 4;
  let dragStart = null, dragCur = null;
  let marqueeActive = false, marqueeAdditive = false, suppressClick = false;

  function getWorksheet() {
    return tableau.extensions.worksheetContent &&
           tableau.extensions.worksheetContent.worksheet;
  }

  function hoverTuple(tupleId, ev) {
    const ws = getWorksheet();
    if (!ws || !ws.hoverTupleAsync) return;
    const tip = ev ? { tooltipAnchorPoint: { x: ev.clientX, y: ev.clientY } } : undefined;
    try { ws.hoverTupleAsync(tupleId, tip).catch(() => {}); } catch (e) { /* ignore */ }
  }

  function pushSelection() {
    const ws = getWorksheet();
    if (ws && ws.selectTuplesAsync) {
      try { ws.selectTuplesAsync([...selectedTuples], SELECT_SIMPLE).catch(() => {}); } catch (e) { /* ignore */ }
    }
    applySelectionStyles();
  }

  function wireInteractivity() {
    host.addEventListener('mousemove', (ev) => {
      if (!config.enableTooltips || marqueeActive) return;
      const t = ev.target.closest && ev.target.closest('.cb-interactive');
      if (!t) { clearHoverState(); return; }

      // A ring that aggregates >1 mark can't be described by the native
      // per-mark tooltip → custom DOM tip. A single-mark ring uses native.
      if (t._tupleIds.length > 1) {
        if (hoverEl !== null) hoverTuple(0);
        hoverEl = null;
        showTip(t._tip, ev);
      } else {
        hideTip();
        const now = (window.performance && performance.now()) || Date.now();
        if (t !== hoverEl || now - hoverAt > 50) {
          hoverEl = t;
          hoverAt = now;
          hoverTuple(t._tupleIds[0], ev);
        }
      }
    });
    host.addEventListener('mouseleave', clearHoverState);
    host.addEventListener('click', (ev) => {
      if (!config.enableTooltips) return;
      if (suppressClick) { suppressClick = false; return; }
      const t = ev.target.closest && ev.target.closest('.cb-interactive');
      toggleSelection(t && t._tupleIds ? t._tupleIds : null, ev.ctrlKey || ev.metaKey);
    });

    host.addEventListener('mousedown', (ev) => {
      if (!config.enableTooltips || ev.button !== 0) return;
      suppressClick = false;
      dragStart = { x: ev.clientX, y: ev.clientY };
      dragCur = { x: ev.clientX, y: ev.clientY };
      marqueeAdditive = ev.ctrlKey || ev.metaKey;
      marqueeActive = false;
      document.addEventListener('mousemove', onMarqueeMove, true);
      document.addEventListener('mouseup', onMarqueeUp, true);
    });
  }

  function onMarqueeMove(ev) {
    if (!dragStart) return;
    dragCur = { x: ev.clientX, y: ev.clientY };
    if (!marqueeActive) {
      if (Math.abs(dragCur.x - dragStart.x) <= MARQUEE_THRESHOLD &&
          Math.abs(dragCur.y - dragStart.y) <= MARQUEE_THRESHOLD) return;
      marqueeActive = true;
      clearHoverState();
      document.body.style.userSelect = 'none';
      ensureMarquee().style.display = 'block';
    }
    ev.preventDefault();
    drawMarquee();
  }

  function onMarqueeUp() {
    document.removeEventListener('mousemove', onMarqueeMove, true);
    document.removeEventListener('mouseup', onMarqueeUp, true);
    if (marqueeActive) {
      applyMarquee(marqueeRect(), marqueeAdditive);
      hideMarquee();
      document.body.style.userSelect = '';
      suppressClick = true;
      marqueeActive = false;
    }
    dragStart = dragCur = null;
  }

  function marqueeRect() {
    return {
      left: Math.min(dragStart.x, dragCur.x), top: Math.min(dragStart.y, dragCur.y),
      right: Math.max(dragStart.x, dragCur.x), bottom: Math.max(dragStart.y, dragCur.y),
    };
  }
  // Select every bar whose drawn arc overlaps the marquee. Hit-test the visible
  // `.cb-bar` arc (not the <g>'s bounding box, which also spans the track +
  // label), then read tuple ids off its interactive ancestor.
  function applyMarquee(rect, additive) {
    const hits = new Set();
    host.querySelectorAll('.cb-bar').forEach((node) => {
      const r = node.getBoundingClientRect();
      const overlaps = !(r.right < rect.left || r.left > rect.right ||
                         r.bottom < rect.top || r.top > rect.bottom);
      if (!overlaps) return;
      const mark = node.closest('.cb-interactive');
      if (mark && mark._tupleIds) mark._tupleIds.forEach((id) => hits.add(id));
    });
    if (!additive) selectedTuples = new Set();
    hits.forEach((id) => selectedTuples.add(id));
    pushSelection();
  }

  let marqueeEl = null;
  function ensureMarquee() {
    if (!marqueeEl) {
      marqueeEl = el('div', 'cb-marquee');
      marqueeEl.style.display = 'none';
      document.body.appendChild(marqueeEl);
    }
    return marqueeEl;
  }
  function drawMarquee() {
    const r = marqueeRect();
    const m = ensureMarquee();
    m.style.left = r.left + 'px';
    m.style.top = r.top + 'px';
    m.style.width = (r.right - r.left) + 'px';
    m.style.height = (r.bottom - r.top) + 'px';
  }
  function hideMarquee() { if (marqueeEl) marqueeEl.style.display = 'none'; }

  // Click selection, mirroring native Tableau behaviour.
  function toggleSelection(ids, additive) {
    if (!ids || !ids.length) {
      if (!additive && selectedTuples.size) { selectedTuples.clear(); pushSelection(); }
      return;
    }
    const allSelected = ids.every((id) => selectedTuples.has(id));
    if (allSelected) {
      if (additive) ids.forEach((id) => selectedTuples.delete(id));
      else if (selectedTuples.size === ids.length) selectedTuples.clear();
      else selectedTuples = new Set(ids);
    } else {
      if (!additive) selectedTuples.clear();
      ids.forEach((id) => selectedTuples.add(id));
    }
    pushSelection();
  }

  // Dim marks that aren't part of the selection; full opacity when nothing is
  // selected. A mark counts as selected when all its tuples are selected. Both
  // the ring and its legend row carry the same ids, so they dim together.
  function applySelectionStyles() {
    const active = selectedTuples.size > 0;
    host.querySelectorAll('.cb-interactive').forEach((node) => {
      if (!active) { node.classList.remove('cb-dimmed', 'cb-selected'); return; }
      const sel = node._tupleIds.every((id) => selectedTuples.has(id));
      node.classList.toggle('cb-selected', sel);
      node.classList.toggle('cb-dimmed', !sel);
    });
  }

  // Re-derive the local selection from the worksheet's actual selected marks
  // (keeps us correct after filters shift tuple ids, and reflects highlights
  // driven from other sheets in a dashboard).
  async function syncSelectionFromWorksheet(model) {
    const ws = getWorksheet();
    if (!ws || !ws.getSelectedMarksAsync || !model || !model.tupleKeys) return;
    try {
      const result = await ws.getSelectedMarksAsync();
      const tables = (result && result.data) || [];
      const fieldNames = model.catFields;
      const keys = new Set();
      tables.forEach((tbl) => {
        const cols = tbl.columns || [];
        const idxs = fieldNames.map((n) => cols.findIndex((c) => c.fieldName === n));
        if (idxs.some((i) => i < 0)) return;
        (tbl.data || []).forEach((row) => keys.add(joinKey(idxs.map((i) => row[i].value))));
      });
      const next = new Set();
      Object.keys(model.tupleKeys).forEach((tid) => {
        if (keys.has(model.tupleKeys[tid])) next.add(Number(tid));
      });
      selectedTuples = next;
    } catch (e) { /* leave selection as-is on failure */ }
  }

  function clearHoverState() {
    if (hoverEl) { hoverTuple(0); hoverEl = null; }
    hideTip();
  }

  /* ---------- custom tooltip (aggregate rings) ---------- */
  let tipEl = null;
  function ensureTip() {
    if (!tipEl) {
      tipEl = el('div', 'cb-tip');
      tipEl.style.display = 'none';
      document.body.appendChild(tipEl);
    }
    return tipEl;
  }
  function showTip(info, ev) {
    const tip = ensureTip();
    let html = '';
    info.fields.forEach((field, k) => {
      html += '<div class="cb-tip-row"><span class="cb-tip-k">' + escapeHtml(field) +
        '</span><span class="cb-tip-v">' + escapeHtml(info.parts[k]) + '</span></div>';
    });
    html += '<div class="cb-tip-row"><span class="cb-tip-k">' + escapeHtml(info.valueField) +
      '</span><span class="cb-tip-v">' + escapeHtml(formatFull(info.value)) + '</span></div>';
    tip.innerHTML = html;
    tip.style.display = 'block';
    positionTip(ev);
  }
  function positionTip(ev) {
    if (!tipEl) return;
    const pad = 14;
    const r = tipEl.getBoundingClientRect();
    let x = ev.clientX + pad, y = ev.clientY + pad;
    if (x + r.width > window.innerWidth) x = ev.clientX - r.width - pad;
    if (y + r.height > window.innerHeight) y = ev.clientY - r.height - pad;
    tipEl.style.left = Math.max(0, x) + 'px';
    tipEl.style.top = Math.max(0, y) + 'px';
  }
  function hideTip() { if (tipEl) tipEl.style.display = 'none'; }

  /* ---------- plumbing (encodings + summary data) ---------- */
  async function getEncodedFields(ws) {
    const spec = await ws.getVisualSpecificationAsync();
    const marks = spec.marksSpecifications[spec.activeMarksSpecificationIndex];
    // Each tile can hold several pills → collect ALL fields per encoding id.
    const fields = {};
    for (const enc of marks.encodings) {
      if (enc.field) (fields[enc.id] || (fields[enc.id] = [])).push(enc.field.name);
    }
    return fields;
  }

  async function readSummary(ws) {
    // ignoreSelection: true is essential — otherwise a restored selection makes
    // the reader return only the selected marks, so the chart renders a subset.
    const reader = await ws.getSummaryDataReaderAsync(undefined, { ignoreSelection: true });
    let columns = [], data = [];
    for (let p = 0; p < reader.pageCount; p++) {
      const page = await reader.getPageAsync(p);
      columns = page.columns;
      data = data.concat(page.data);
    }
    if (reader.releaseAsync) { try { await reader.releaseAsync(); } catch (e) { /* ignore */ } }
    return { columns, data };
  }

  function findColumn(cols, fieldName) {
    let i = cols.findIndex((c) => c.fieldName === fieldName);
    if (i < 0) i = cols.findIndex((c) => c.fieldName && c.fieldName.indexOf(fieldName) >= 0);
    return i < 0 ? 0 : i;
  }
  function findValueColumn(cols, fieldName) {
    let i = cols.findIndex((c) => c.fieldName === fieldName);
    if (i < 0) i = cols.findIndex((c) => c.fieldName && c.fieldName.indexOf(fieldName) >= 0);
    if (i < 0) i = cols.findIndex((c) => isNumericCol(c));
    return i < 0 ? cols.length - 1 : i;
  }
  function isNumericCol(c) {
    const t = (c.dataType || '').toLowerCase();
    return t === 'float' || t === 'integer' || t === 'int';
  }
  function cellText(cell) {
    if (!cell) return '';
    return cell.formattedValue != null ? cell.formattedValue : String(cell.value);
  }

  /* ---------- theme ---------- */
  function applyTheme() {
    document.body.style.background = config.bgColor;
    document.documentElement.style.setProperty(
      '--cb-font', FONT_STACKS[config.fontFamily] || FONT_STACKS.system);
    document.documentElement.style.setProperty('--cb-label-size', config.labelSize + 'px');
  }

  /* =========================================================
     CONFIG MODAL — declarative schema (one entry per control)
     ========================================================= */
  const SCHEMA = [
    { section: 'General', fields: [
      { key: 'title', label: 'Chart title', type: 'text' },
      { key: 'bgColor', label: 'Background', type: 'color' },
      { key: 'fontFamily', label: 'Font', type: 'select',
        options: [['system', 'System'], ['serif', 'Serif'], ['mono', 'Monospace']] },
      { key: 'enableTooltips', label: 'Tableau tooltips + selection', type: 'checkbox' },
    ]},
    { section: 'Color', fields: [
      { key: 'colorMode', label: 'Color mode', type: 'select',
        options: [['palette', 'Categorical palette'], ['single', 'Single color'], ['sequential', 'Sequential (by value)']] },
      { key: 'barColor', label: 'Single color', type: 'color' },
      { key: 'lowColor', label: 'Sequential low', type: 'color' },
      { key: 'highColor', label: 'Sequential high', type: 'color' },
      { key: 'reverseScale', label: 'Reverse sequential', type: 'checkbox' },
    ]},
    { section: 'Geometry', fields: [
      { key: 'startAngleDeg', label: 'Start angle (deg)', type: 'range', min: 0, max: 360, step: 5 },
      { key: 'sweepDeg', label: 'Max arc / sweep (deg)', type: 'range', min: 90, max: 360, step: 5 },
      { key: 'innerRadiusPct', label: 'Center hole (fraction)', type: 'range', min: 0, max: 0.85, step: 0.02 },
      { key: 'ringPadding', label: 'Gap between rings', type: 'range', min: 0, max: 0.9, step: 0.05 },
      { key: 'roundedCaps', label: 'Rounded arc ends', type: 'checkbox' },
    ]},
    { section: 'Track & order', fields: [
      { key: 'showTrack', label: 'Show track behind bars', type: 'checkbox' },
      { key: 'trackColor', label: 'Track color', type: 'color' },
      { key: 'sortOrder', label: 'Sort rings', type: 'select',
        options: [['desc', 'Largest outer'], ['asc', 'Smallest outer'], ['none', 'Data order']] },
    ]},
    { section: 'Labels', fields: [
      { key: 'showLegend', label: 'Show category + value', type: 'checkbox' },
      { key: 'labelPlacement', label: 'Label placement', type: 'select',
        options: [['bar-start', 'Beside bar start'], ['corner', 'Top-left legend']] },
      { key: 'showValues', label: 'Show values on arcs', type: 'checkbox' },
      { key: 'labelSize', label: 'Label size (px)', type: 'range', min: 8, max: 20, step: 1 },
    ]},
    { section: 'Number format', fields: [
      { key: 'numPrefix', label: 'Prefix', type: 'text' },
      { key: 'numSuffix', label: 'Suffix', type: 'text' },
      { key: 'numDecimals', label: 'Decimals', type: 'range', min: 0, max: 4, step: 1 },
      { key: 'numUnit', label: 'Unit', type: 'select',
        options: [['auto', 'Auto'], ['K', 'Thousands (K)'], ['M', 'Millions (M)'], ['B', 'Billions (B)'], ['none', 'None']] },
      { key: 'numThousands', label: 'Thousands separator', type: 'checkbox' },
    ]},
  ];

  function openConfigModal() {
    const overlay = document.getElementById('cfg-overlay');
    if (overlay) overlay.hidden = false;
  }

  function wireModal() {
    const overlay = document.getElementById('cfg-overlay');
    const openBtn = document.getElementById('cfg-open');
    const doneBtn = document.getElementById('cfg-done');
    const resetBtn = document.getElementById('cfg-reset');

    buildModalBody();

    // Closing forces an immediate workbook save (cancels the debounce).
    const close = () => { overlay.hidden = true; clearTimeout(saveTimer); flushWorkbookSettings(); };

    openBtn.addEventListener('click', openConfigModal);
    doneBtn.addEventListener('click', close);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !overlay.hidden) close(); });
    resetBtn.addEventListener('click', () => {
      config = { ...DEFAULT_CONFIG };
      saveConfig();
      buildModalBody();
      if (lastModel) render(lastModel);
    });
  }

  function buildModalBody() {
    const body = document.getElementById('cfg-body');
    body.innerHTML = '';
    SCHEMA.forEach((sec) => {
      const wrap = el('div', 'cfg-section');
      wrap.appendChild(el('h3', 'cfg-section-title', sec.section));
      sec.fields.forEach((f) => wrap.appendChild(buildField(f)));
      body.appendChild(wrap);
    });
  }

  function buildField(f) {
    const row = el('label', 'cfg-field');
    row.appendChild(el('span', 'cfg-label', f.label));
    let input;
    const val = config[f.key];

    if (f.type === 'checkbox') {
      input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = !!val;
      input.addEventListener('change', () => commit(f.key, input.checked));
      row.classList.add('cfg-check');
    } else if (f.type === 'color') {
      input = document.createElement('input');
      input.type = 'color';
      input.value = val;
      input.addEventListener('input', () => commit(f.key, input.value));
    } else if (f.type === 'select') {
      input = document.createElement('select');
      f.options.forEach(([v, label]) => {
        const o = document.createElement('option');
        o.value = v; o.textContent = label;
        if (v === val) o.selected = true;
        input.appendChild(o);
      });
      input.addEventListener('change', () => commit(f.key, input.value));
    } else if (f.type === 'range') {
      const box = el('span', 'cfg-range-box');
      input = document.createElement('input');
      input.type = 'range';
      input.min = f.min; input.max = f.max; input.step = f.step;
      input.value = val;
      const out = el('span', 'cfg-range-out', String(val));
      input.addEventListener('input', () => {
        const num = parseFloat(input.value);
        out.textContent = String(num);
        commit(f.key, num);
      });
      box.appendChild(input); box.appendChild(out);
      row.appendChild(box);
      return row;
    } else { // text
      input = document.createElement('input');
      input.type = 'text';
      input.value = val == null ? '' : val;
      input.addEventListener('input', () => commit(f.key, input.value));
    }
    row.appendChild(input);
    return row;
  }

  function commit(key, value) {
    config[key] = value;
    saveConfig();
    if (lastModel) render(lastModel);
  }

  /* ---------- config persistence (workbook settings) ----------
     tableau.extensions.settings is saved INTO the workbook: it survives
     save/close/reopen and travels with the .twb/.twbx. It is the single
     source of truth. saveAsync() must be serialized (one save in flight) and
     debounced (a slider drag shouldn't fire dozens of saves). ------------- */
  function loadConfig() {
    let saved = {};
    try {
      const store = tableau.extensions.settings;
      const raw = store && store.get(STORAGE_KEY);
      if (raw) saved = JSON.parse(raw);
    } catch (e) {
      console.error('reading workbook settings failed:', e);
    }
    return { ...DEFAULT_CONFIG, ...saved };
  }

  let saveTimer = null;
  let saveInFlight = false;
  let savePending = false;

  function saveConfig() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(flushWorkbookSettings, 350);
  }

  function flushWorkbookSettings() {
    const store = tableau.extensions.settings;
    if (!store) return;
    try {
      store.set(STORAGE_KEY, JSON.stringify(config));
    } catch (e) {
      console.error('settings.set failed:', e);
      return;
    }
    if (saveInFlight) { savePending = true; return; }
    saveInFlight = true;
    store.saveAsync().then(
      () => { saveInFlight = false; if (savePending) { savePending = false; flushWorkbookSettings(); } },
      (err) => {
        saveInFlight = false;
        console.error('settings.saveAsync failed:', err);
        if (savePending) { savePending = false; flushWorkbookSettings(); }
      }
    );
  }

  /* ---------- tiny DOM helpers + states ---------- */
  function el(tag, cls, text) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }
  function showMessage(html) {
    applyTheme();
    host.innerHTML = '<div class="cb-empty"><p>' + html + '</p></div>';
  }

  // Rich "get started" card shown before both tiles are filled. Styled with a
  // FIXED palette (never the configurable theme vars) so a fresh viz always
  // reads correctly regardless of config.
  function showEmptyState(fields) {
    applyTheme();
    fields = fields || {};
    const chip = (id, label, kind) => {
      const got = fields[id] && fields[id].length;
      return '<span class="cb-chip ' + kind + (got ? ' filled' : '') + '">' +
        label + (got ? ': ' + escapeHtml(fields[id].join(', ')) : '') + '</span>';
    };
    host.innerHTML =
      '<div class="cb-empty"><div class="cb-empty-card">' +
        '<div class="cb-empty-head">' +
          '<div class="cb-empty-thumb">' + buildThumb() + '</div>' +
          '<div class="cb-empty-meta">' +
            '<div class="cb-empty-title">Circular Bar Chart</div>' +
            '<div class="cb-empty-desc">Concentric rings — one per category — where each colored ' +
            'arc’s length encodes the measure against a faint full-circle track. Customizable ' +
            'from the &#9881; gear or the <b>Format Extension</b> button.</div>' +
          '</div>' +
        '</div>' +
        '<div class="cb-empty-guide">' +
          '<div class="cb-empty-guide-title">Get started</div>' +
          '<div class="cb-empty-guide-row">Drag one or more <b>dimensions</b> onto ' +
            chip('category', 'Category', 'dim') + ' and one <b>measure</b> onto ' +
            chip('value', 'Value', 'measure') + '.' +
          '</div>' +
        '</div>' +
      '</div></div>';
  }

  // Inline SVG thumbnail: a mini concentric circular bar chart (fixed colors).
  function buildThumb() {
    const cx = 42, cy = 42, r0 = 9, rmax = 35;
    const fr = [0.92, 0.66, 0.5, 0.74, 0.34];
    const blues = ['#1f6fff', '#4a8bff', '#7aa9ff', '#3f7fff', '#9cc0ff'];
    const n = fr.length;
    const slot = (rmax - r0) / n;
    const th = slot * 0.62;
    let s = '';
    for (let i = 0; i < n; i++) {
      const r = rmax - slot * i - slot / 2;
      s += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#e9eaee" stroke-width="${th}"/>`;
      const d = ringArcPath(cx, cy, r, 0, fr[i] * TAU * 0.92);
      s += `<path d="${d}" fill="none" stroke="${blues[i]}" stroke-width="${th}" stroke-linecap="round"/>`;
    }
    return `<svg viewBox="0 0 84 84" width="84" height="84" xmlns="${SVG_NS}" role="img" aria-label="Circular bar chart preview">${s}</svg>`;
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"]/g, (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }
})();

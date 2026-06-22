'use strict';

/* =============================================================
   RADIAL BAR CHART — Tableau Viz Extension
   -------------------------------------------------------------
   One bar per category, drawn as an annular sector (wedge) that
   grows OUTWARD from an inner radius. Bar length encodes the
   measure; bars are laid out around a circle (full or partial
   sweep). Everything visual is driven by `config` and editable
   live from the gear → settings modal.

   The four moves (see CLAUDE.md):
     1. CONNECT   — initializeAsync(), grab the worksheet
     2. ENCODINGS — Category / Value tiles (manifest.trex)
     3. DATA      — read summary data, aggregate per category
     4. DRAW      — render(model) into #viz; redraw on data change

   This one renders as a single inline <svg> (polar geometry is
   far easier as SVG paths than as DOM/CSS). The plumbing is the
   same as the DOM charts; only the BUILD ZONE differs.
   ============================================================= */

(function () {
  const host = document.getElementById('viz');
  // Key for the config blob stored in the workbook (tableau.extensions.settings).
  // Keep this STABLE so a workbook saved by an earlier build still loads its
  // config. Workbook settings are the single source of truth (no localStorage
  // cache — it would leak the last-used config across instances on one origin
  // and break "fresh viz = defaults"; we never paint before init resolves).
  const STORAGE_KEY = 'radialbar-config-v1';

  const SVG_NS = 'http://www.w3.org/2000/svg';

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
    barColor: '#4e79a7',      // used when colorMode === 'single'
    lowColor: '#dbe6f3',      // sequential scale endpoints (also drives "by value")
    highColor: '#1f4e79',
    reverseScale: false,
    // Per-category color overrides { categoryKey: '#hex' } — double-click a bar,
    // or use the modal's "Series colors" section. Always win (palette mode only).
    seriesColors: {},
    // First-seen order of category keys (append-only, persisted). PALETTE colors
    // are assigned by a category's position HERE, not its index in the current
    // (filtered) bars — so a category keeps its color across filtering (see
    // CLAUDE.md "Color categorical marks by series IDENTITY").
    seriesOrder: [],

    // geometry
    innerRadiusPct: 0.25,     // inner hole as a fraction of the outer radius
    startAngleDeg: 0,         // 0 = top; increases clockwise
    sweepDeg: 360,            // total arc the bars span (360 = full ring)
    gapDeg: 2,                // angular gap between adjacent bars
    showTrack: true,          // faint full-length track behind each bar

    // comparison grid + center total
    showGrid: true,           // concentric value rings behind the bars
    gridTicks: 4,             // number of rings (evenly spaced from 0 to max)
    gridLabels: false,        // value label at each ring
    showTotal: true,          // sum of all values shown in the inner circle

    // labels
    showLabels: true,         // category labels around the outside
    showValues: false,        // measure value drawn on each bar
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

  // Start from defaults; the authoritative workbook settings are loaded once
  // initializeAsync() resolves (the settings API isn't ready before that, and
  // render() runs only after that point — so there is no pre-init paint to
  // cache for).
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

      // Re-render on size changes. We listen THREE ways and coalesce through
      // rAF: ResizeObserver (pane/object resizes), window 'resize' (belt-and-
      // braces), and a polling backstop (Tableau often sizes the iframe to its
      // final dimensions AFTER first paint, and in some hosts neither RO nor
      // 'resize' fires for that — leaving the chart stuck at its initial size).
      if (window.ResizeObserver) new ResizeObserver(scheduleRender).observe(host);
      window.addEventListener('resize', scheduleRender);
      startSizeWatch();

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

  // Coalesce resize-driven re-renders into a single rAF so a burst of resize
  // events (drag-resizing a pane) produces one redraw per frame.
  let renderRAF = 0;
  function scheduleRender() {
    if (!lastModel || renderRAF) return;
    renderRAF = requestAnimationFrame(() => { renderRAF = 0; render(lastModel); });
  }
  // Poll the host size as a backstop for resize events that don't fire. Renders
  // directly (not via the rAF-coalesced scheduleRender) so it works even where
  // rAF is starved; the 250ms interval is its own throttle.
  let watchedW = 0, watchedH = 0;
  function startSizeWatch() {
    setInterval(() => {
      const w = host.clientWidth + 0.001 * (window.innerWidth || 0);
      const h = host.clientHeight;
      if (w === watchedW && h === watchedH) return;
      watchedW = w; watchedH = h;
      if (lastModel) render(lastModel);
    }, 250);
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
    // lock in stable palette slots per category identity (palette mode only)
    if (config.colorMode === 'palette') ensureSeriesOrder(model);

    const wrap = el('div', 'rb');
    if (config.title) wrap.appendChild(el('div', 'rb-title', config.title));
    const stage = el('div', 'rb-stage');
    wrap.appendChild(stage);
    host.appendChild(wrap);

    // Measure the stage *after* it's in the DOM so we get real pixels.
    drawSvg(model, stage);

    // Re-apply selection dimming/highlight (the DOM was just rebuilt).
    applySelectionStyles();
  }

  function drawSvg(model, stage) {
    const items = model.items;
    const W = Math.max(40, stage.clientWidth || 600);
    const H = Math.max(40, stage.clientHeight || 400);
    const cx = W / 2, cy = H / 2;

    const svg = svgEl('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, class: 'rb-svg' });

    // Reserve room for outside labels; bars fill the rest.
    const labelPad = config.showLabels ? Math.max(36, config.labelSize * 3) : 10;
    const outerR = Math.max(8, Math.min(W, H) / 2 - labelPad);
    const innerR = outerR * clamp(config.innerRadiusPct, 0, 0.85);

    const n = items.length;
    const start = deg2rad(config.startAngleDeg);
    const sweep = deg2rad(clamp(config.sweepDeg, 10, 360));
    const gap = deg2rad(Math.max(0, config.gapDeg));
    // Full ring: a gap after every bar (incl. wrap). Partial arc: gaps between only.
    const fullRing = config.sweepDeg >= 359.9;
    const totalGap = gap * (fullRing ? n : Math.max(0, n - 1));
    const slice = Math.max(0.001, (sweep - totalGap) / n);

    const ticks = Math.max(2, Math.round(config.gridTicks));
    const gridInk = contrastInk(config.bgColor);

    // Comparison grid: concentric rings at evenly-spaced value levels, drawn
    // BEHIND the bars so each bar's end can be read against them. Spans the
    // same sweep as the bars (a full circle when the sweep is 360°).
    if (config.showGrid && model.max > 0) {
      for (let i = 1; i <= ticks; i++) {
        const rr = innerR + (i / ticks) * (outerR - innerR);
        const ring = fullRing
          ? svgEl('circle', { cx, cy, r: rr, class: 'rb-grid' })
          : svgEl('path', { d: arcStroke(cx, cy, rr, start, start + sweep), class: 'rb-grid' });
        ring.style.stroke = gridInk;
        ring.style.strokeOpacity = '0.18';
        svg.appendChild(ring);
      }
    }

    let a = start;
    items.forEach((it, i) => {
      const a0 = a;
      const a1 = a + slice;
      const mid = (a0 + a1) / 2;
      const t = model.max > 0 && it.value != null ? clamp(it.value / model.max, 0, 1) : 0;
      const r = innerR + t * (outerR - innerR);
      const color = colorFor(it, model);

      const g = svgEl('g', { class: 'rb-bar-g' });

      if (config.showTrack) {
        g.appendChild(svgEl('path', { d: annulus(cx, cy, innerR, outerR, a0, a1), class: 'rb-track' }));
      }
      if (it.value != null) {
        g.appendChild(svgEl('path', { d: annulus(cx, cy, innerR, r, a0, a1), fill: color, class: 'rb-bar' }));
      }

      // Category label, horizontal, just outside the ring; anchor by side.
      if (config.showLabels) {
        const [lx, ly] = pt(cx, cy, outerR + 8, mid);
        const s = Math.sin(mid);
        const anchor = s > 0.08 ? 'start' : s < -0.08 ? 'end' : 'middle';
        const txt = svgEl('text', {
          x: lx, y: ly, 'text-anchor': anchor, 'dominant-baseline': 'middle', class: 'rb-label',
        });
        txt.textContent = it.label;
        txt.style.fontSize = config.labelSize + 'px';
        g.appendChild(txt);
      }

      // Value label, drawn near the bar's outer end (sits on the bar).
      if (config.showValues && it.value != null) {
        const [vx, vy] = pt(cx, cy, Math.max(innerR + 10, r - 12), mid);
        const txt = svgEl('text', {
          x: vx, y: vy, 'text-anchor': 'middle', 'dominant-baseline': 'middle', class: 'rb-value',
        });
        txt.textContent = formatNumber(it.value);
        txt.style.fontSize = Math.max(8, config.labelSize - 1) + 'px';
        txt.style.fill = contrastInk(color);
        g.appendChild(txt);
      }

      // Tag as an interactive mark (hover tooltip + click/marquee selection).
      // Tuple ids back the worksheet selection; the tip payload drives our
      // custom tooltip when a wedge aggregates more than one mark.
      if (config.enableTooltips && it.tupleIds.length) {
        g.classList.add('rb-interactive');
        g._tupleIds = it.tupleIds;
        g._key = it.key;          // identity for the per-series color override
        g._curColor = color;      // seed value for the color picker
        g._tip = { fields: model.catFields, parts: it.parts, valueField: model.valueField, value: it.value };
      }

      svg.appendChild(g);
      a = a1 + gap;
    });

    // Grid value labels, drawn ON TOP of the bars (with a background-colored
    // halo so they stay legible) along the sweep's start radial.
    if (config.showGrid && config.gridLabels && model.max > 0) {
      for (let i = 1; i <= ticks; i++) {
        const rr = innerR + (i / ticks) * (outerR - innerR);
        const [gx, gy] = pt(cx, cy, rr, start);
        const txt = svgEl('text', {
          x: gx, y: gy, 'text-anchor': 'middle', 'dominant-baseline': 'middle', class: 'rb-grid-label',
        });
        txt.textContent = formatNumber((i / ticks) * model.max);
        txt.style.fontSize = Math.max(8, config.labelSize - 2) + 'px';
        txt.style.fill = gridInk;
        txt.style.stroke = config.bgColor;   // halo
        txt.style.strokeWidth = '3px';
        txt.style.paintOrder = 'stroke';
        svg.appendChild(txt);
      }
    }

    // Center total: the summed measure in the inner hole, with the measure
    // field name as a caption. Colored for contrast against the background.
    if (config.showTotal) {
      const total = items.reduce((s, it) => (it.value != null ? s + it.value : s), 0);
      const ink = contrastInk(config.bgColor);
      const hasCap = !!model.valueField;
      const numSize = Math.max(11, Math.min(innerR * 0.5, 30));
      const capSize = Math.max(8, Math.min(innerR * 0.24, 13));
      const num = svgEl('text', {
        x: cx, y: hasCap ? cy - capSize * 0.55 : cy,
        'text-anchor': 'middle', 'dominant-baseline': 'middle', class: 'rb-total-num',
      });
      num.textContent = formatNumber(total);
      num.style.fontSize = numSize + 'px';
      num.style.fill = ink;
      svg.appendChild(num);
      if (hasCap) {
        const cap = svgEl('text', {
          x: cx, y: cy + numSize * 0.5,
          'text-anchor': 'middle', 'dominant-baseline': 'middle', class: 'rb-total-cap',
        });
        cap.textContent = model.valueField;
        cap.style.fontSize = capSize + 'px';
        svg.appendChild(cap);
      }
    }

    stage.appendChild(svg);
  }
  /* ▲▲▲  END BUILD ZONE  ▲▲▲ */

  /* ---------- polar / SVG geometry ---------- */
  function svgEl(tag, attrs) {
    const n = document.createElementNS(SVG_NS, tag);
    if (attrs) for (const k in attrs) n.setAttribute(k, attrs[k]);
    return n;
  }
  function deg2rad(d) { return (d * Math.PI) / 180; }
  function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }
  // Point on a circle. theta = 0 is straight up; increases CLOCKWISE.
  function pt(cx, cy, r, theta) {
    return [cx + r * Math.sin(theta), cy - r * Math.cos(theta)];
  }
  // Annular sector (wedge with a hole) from angle a0→a1, radii r0..r1.
  // Outer arc drawn clockwise (sweep 1), inner arc back (sweep 0). When the
  // inner radius is ~0 it degenerates to a pie slice from the center.
  function annulus(cx, cy, r0, r1, a0, a1) {
    const large = (a1 - a0) > Math.PI ? 1 : 0;
    const [x1, y1] = pt(cx, cy, r1, a0);
    const [x2, y2] = pt(cx, cy, r1, a1);
    if (r0 <= 0.01) {
      return `M ${cx} ${cy} L ${x1} ${y1} A ${r1} ${r1} 0 ${large} 1 ${x2} ${y2} Z`;
    }
    const [x3, y3] = pt(cx, cy, r0, a1);
    const [x4, y4] = pt(cx, cy, r0, a0);
    return `M ${x4} ${y4} L ${x1} ${y1} A ${r1} ${r1} 0 ${large} 1 ${x2} ${y2} ` +
           `L ${x3} ${y3} A ${r0} ${r0} 0 ${large} 0 ${x4} ${y4} Z`;
  }
  // A single open arc stroke (no fill) along a0→a1 at radius r — used for the
  // comparison grid rings on partial sweeps.
  function arcStroke(cx, cy, r, a0, a1) {
    const large = (a1 - a0) > Math.PI ? 1 : 0;
    const [x1, y1] = pt(cx, cy, r, a0);
    const [x2, y2] = pt(cx, cy, r, a1);
    return `M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2}`;
  }

  /* ---------- color + number helpers ----------
     A category's color is resolved by IDENTITY, not by bar position:
       1. an explicit per-category override (config.seriesColors[key]) wins;
       2. else by colorMode — single / sequential(by value) / palette;
       3. palette is indexed by the category's slot in the persisted first-seen
          order, so filtering never moves a category's hue. */
  function ensureSeriesOrder(model) {
    const order = config.seriesOrder || (config.seriesOrder = []);
    let changed = false;
    model.items.forEach((it) => { if (order.indexOf(it.key) < 0) { order.push(it.key); changed = true; } });
    if (changed) saveConfig();
  }
  function stableIndex(key) {
    const order = config.seriesOrder || [];
    const i = order.indexOf(key);
    return i < 0 ? 0 : i;
  }
  function colorFor(it, model) {
    // Per-category overrides + identity-stable palette slots apply ONLY in the
    // categorical palette mode. In sequential mode color encodes the value (a
    // continuous scale — an override would punch a hole in the gradient); in
    // single mode every bar is one color. Both ignore seriesColors/seriesOrder.
    if (config.colorMode === 'single') return config.barColor;
    if (config.colorMode === 'sequential') return scaleColor(norm(it.value, model.min, model.max));
    const ov = config.seriesColors && config.seriesColors[it.key];
    if (ov) return ov;                                  // explicit override wins
    return PALETTE[stableIndex(it.key) % PALETTE.length];
  }
  // Coerce any color we produce (hex or "rgb(r, g, b)") to "#rrggbb" for an
  // <input type=color>.
  function anyToHex(c) {
    if (!c) return '#888888';
    const s = String(c);
    if (s[0] === '#') return s.length === 4 ? '#' + s.slice(1).split('').map((x) => x + x).join('') : s;
    const m = s.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
    if (m) return '#' + [1, 2, 3].map((i) => (+m[i]).toString(16).padStart(2, '0')).join('');
    return '#888888';
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
    const n = parseInt(f, 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
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
     Each interactive <g> carries the tuple ids it represents.
     Hover → hoverTupleAsync (single mark) or a custom DOM tip
     (aggregate). Click / marquee → selectTuplesAsync. Listeners
     are delegated off #viz and attached once.
     ========================================================= */
  const SELECT_SIMPLE = (tableau.SelectOptions && tableau.SelectOptions.Simple) || 'select-options-simple';

  let selectedTuples = new Set();   // drives worksheet selection + our dimming
  let hoverEl = null;
  let hoverAt = 0;
  let clickTimer = null;            // single/double-click disambiguation

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

  // Open the OS color picker for a category, store the chosen color as a
  // persisted override, and re-render. (A hidden <input type=color> is the
  // simplest cross-host picker.)
  function pickSeriesColor(key, currentColor) {
    const inp = document.createElement('input');
    inp.type = 'color';
    inp.value = anyToHex(currentColor);
    inp.style.cssText = 'position:fixed;left:-9999px;top:0';
    document.body.appendChild(inp);
    inp.addEventListener('input', () => {
      config.seriesColors[key] = inp.value;
      saveConfig();
      if (lastModel) render(lastModel);
    });
    inp.addEventListener('change', () => { inp.remove(); });
    inp.click();
  }

  function wireInteractivity() {
    host.addEventListener('mousemove', (ev) => {
      if (!config.enableTooltips || marqueeActive) return;
      const t = ev.target.closest && ev.target.closest('.rb-interactive');
      if (!t) { clearHoverState(); return; }

      // A wedge that aggregates >1 mark can't be described by the native
      // per-mark tooltip → custom DOM tip. A single-mark wedge uses native.
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
      const t = ev.target.closest && ev.target.closest('.rb-interactive');
      if (t) {
        // Delay the select so a double-click (color pick) can cancel it and
        // doesn't also toggle the selection.
        if (clickTimer) return;
        const ids = t._tupleIds, additive = ev.ctrlKey || ev.metaKey;
        clickTimer = setTimeout(() => { clickTimer = null; toggleSelection(ids, additive); }, 220);
      } else {
        toggleSelection(null, ev.ctrlKey || ev.metaKey);   // empty space → clear now
      }
    });
    // Double-click a bar → pick a persisted color for that category. Only
    // meaningful in palette mode (sequential/single ignore per-category
    // overrides) — otherwise let the normal click-select stand.
    host.addEventListener('dblclick', (ev) => {
      if (!config.enableTooltips || config.colorMode !== 'palette') return;
      const t = ev.target.closest && ev.target.closest('.rb-interactive');
      if (!t || !t._key) return;
      if (clickTimer) { clearTimeout(clickTimer); clickTimer = null; }
      ev.preventDefault();
      pickSeriesColor(t._key, t._curColor);
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
  // Select every bar whose drawn wedge overlaps the marquee. We hit-test the
  // visible `.rb-bar` path (not the <g>'s bounding box, which also spans the
  // label), then read tuple ids off its interactive ancestor.
  function applyMarquee(rect, additive) {
    const hits = new Set();
    host.querySelectorAll('.rb-bar').forEach((node) => {
      const r = node.getBoundingClientRect();
      const overlaps = !(r.right < rect.left || r.left > rect.right ||
                         r.bottom < rect.top || r.top > rect.bottom);
      if (!overlaps) return;
      const mark = node.closest('.rb-interactive');
      if (mark && mark._tupleIds) mark._tupleIds.forEach((id) => hits.add(id));
    });
    if (!additive) selectedTuples = new Set();
    hits.forEach((id) => selectedTuples.add(id));
    pushSelection();
  }

  let marqueeEl = null;
  function ensureMarquee() {
    if (!marqueeEl) {
      marqueeEl = el('div', 'rb-marquee');
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

  // Dim bars that aren't part of the selection; full opacity when nothing is
  // selected. A mark counts as selected when all its tuples are selected.
  function applySelectionStyles() {
    const active = selectedTuples.size > 0;
    host.querySelectorAll('.rb-interactive').forEach((node) => {
      if (!active) { node.classList.remove('rb-dimmed', 'rb-selected'); return; }
      const sel = node._tupleIds.every((id) => selectedTuples.has(id));
      node.classList.toggle('rb-selected', sel);
      node.classList.toggle('rb-dimmed', !sel);
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

  /* ---------- custom tooltip (aggregate wedges) ---------- */
  let tipEl = null;
  function ensureTip() {
    if (!tipEl) {
      tipEl = el('div', 'rb-tip');
      tipEl.style.display = 'none';
      document.body.appendChild(tipEl);
    }
    return tipEl;
  }
  function showTip(info, ev) {
    const tip = ensureTip();
    let html = '';
    info.fields.forEach((field, k) => {
      html += '<div class="rb-tip-row"><span class="rb-tip-k">' + escapeHtml(field) +
        '</span><span class="rb-tip-v">' + escapeHtml(info.parts[k]) + '</span></div>';
    });
    html += '<div class="rb-tip-row"><span class="rb-tip-k">' + escapeHtml(info.valueField) +
      '</span><span class="rb-tip-v">' + escapeHtml(formatFull(info.value)) + '</span></div>';
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
      '--rb-font', FONT_STACKS[config.fontFamily] || FONT_STACKS.system);
    document.documentElement.style.setProperty('--rb-label-size', config.labelSize + 'px');
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
      { key: 'innerRadiusPct', label: 'Inner hole (fraction)', type: 'range', min: 0, max: 0.85, step: 0.05 },
      { key: 'startAngleDeg', label: 'Start angle (deg)', type: 'range', min: 0, max: 360, step: 5 },
      { key: 'sweepDeg', label: 'Sweep (deg)', type: 'range', min: 90, max: 360, step: 5 },
      { key: 'gapDeg', label: 'Gap between bars (deg)', type: 'range', min: 0, max: 20, step: 1 },
      { key: 'showTrack', label: 'Show track behind bars', type: 'checkbox' },
    ]},
    { section: 'Grid & total', fields: [
      { key: 'showGrid', label: 'Show comparison grid', type: 'checkbox' },
      { key: 'gridTicks', label: 'Grid rings', type: 'range', min: 2, max: 8, step: 1 },
      { key: 'gridLabels', label: 'Grid value labels', type: 'checkbox' },
      { key: 'showTotal', label: 'Show total in center', type: 'checkbox' },
    ]},
    { section: 'Labels', fields: [
      { key: 'showLabels', label: 'Show category labels', type: 'checkbox' },
      { key: 'showValues', label: 'Show values on bars', type: 'checkbox' },
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
    if (!overlay) return;
    buildModalBody();          // rebuild so the dynamic "Series colors" list
    overlay.hidden = false;    // reflects whatever categories are in the view
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
      // fresh copies of the mutable collections (don't alias DEFAULT_CONFIG).
      config = { ...DEFAULT_CONFIG, seriesColors: {}, seriesOrder: [] };
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
    const colorsSec = buildSeriesColorsSection();
    if (colorsSec) body.appendChild(colorsSec);
  }
  // Dynamic section: one color swatch per category currently in the view, plus a
  // ↺ reset-to-palette. Built from lastModel (the schema is static; categories
  // are data-driven), so it only appears once fields are dropped.
  function buildSeriesColorsSection() {
    // Per-category overrides only apply in the categorical palette mode, so the
    // section would do nothing in sequential/single — hide it there.
    if (config.colorMode !== 'palette') return null;
    if (!lastModel || !lastModel.items.length) return null;
    const wrap = el('div', 'cfg-section');
    wrap.appendChild(el('h3', 'cfg-section-title', 'Series colors'));
    lastModel.items.forEach((it) => {
      const row = el('label', 'cfg-field');
      row.appendChild(el('span', 'cfg-label', it.label));
      const box = el('span', 'cfg-color-box');
      const inp = document.createElement('input');
      inp.type = 'color';
      inp.value = anyToHex(colorFor(it, lastModel));
      inp.addEventListener('input', () => {
        config.seriesColors[it.key] = inp.value;
        saveConfig();
        if (lastModel) render(lastModel);
      });
      const reset = el('button', 'cfg-mini', '↺');
      reset.type = 'button';
      reset.title = 'Reset to palette / scale';
      reset.addEventListener('click', (e) => {
        e.preventDefault();
        delete config.seriesColors[it.key];
        saveConfig();
        inp.value = anyToHex(colorFor(it, lastModel));
        if (lastModel) render(lastModel);
      });
      box.appendChild(inp);
      box.appendChild(reset);
      row.appendChild(box);
      wrap.appendChild(row);
    });
    return wrap;
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
    const cfg = { ...DEFAULT_CONFIG, ...saved };
    // Deep-copy the mutable collections so we never mutate DEFAULT_CONFIG's
    // shared objects (which would leak across instances and break Reset).
    cfg.seriesColors = { ...(saved.seriesColors || {}) };
    cfg.seriesOrder = [...(saved.seriesOrder || [])];
    return cfg;
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
    host.innerHTML = '<div class="rb-empty"><p>' + html + '</p></div>';
  }

  // Rich "get started" card shown before both tiles are filled. Styled with a
  // FIXED palette (never the configurable theme vars) so a fresh viz always
  // reads correctly regardless of config.
  function showEmptyState(fields) {
    applyTheme();
    fields = fields || {};
    const chip = (id, label, kind) => {
      const got = fields[id] && fields[id].length;
      return '<span class="rb-chip ' + kind + (got ? ' filled' : '') + '">' +
        label + (got ? ': ' + escapeHtml(fields[id].join(', ')) : '') + '</span>';
    };
    host.innerHTML =
      '<div class="rb-empty"><div class="rb-empty-card">' +
        '<div class="rb-empty-head">' +
          '<div class="rb-empty-thumb">' + buildThumb() + '</div>' +
          '<div class="rb-empty-meta">' +
            '<div class="rb-empty-title">Radial Bar Chart</div>' +
            '<div class="rb-empty-desc">Bars arranged around a circle, growing outward from a ' +
            'center hole — length encodes the measure. Customizable from the &#9881; gear ' +
            'or the <b>Format Extension</b> button.</div>' +
          '</div>' +
        '</div>' +
        '<div class="rb-empty-guide">' +
          '<div class="rb-empty-guide-title">Get started</div>' +
          '<div class="rb-empty-guide-row">Drag one or more <b>dimensions</b> onto ' +
            chip('category', 'Category', 'dim') + ' and one <b>measure</b> onto ' +
            chip('value', 'Value', 'measure') + '.' +
          '</div>' +
        '</div>' +
      '</div></div>';
  }

  // Inline SVG thumbnail: a mini radial bar chart (fixed colors).
  function buildThumb() {
    const cx = 42, cy = 42, r0 = 11, rmax = 34, n = 7;
    const tops = [0.5, 0.82, 0.35, 1, 0.6, 0.74, 0.45];
    let s = '';
    for (let i = 0; i < n; i++) {
      const a0 = (i / n) * 2 * Math.PI + 0.06;
      const a1 = ((i + 1) / n) * 2 * Math.PI - 0.06;
      const r = r0 + tops[i] * (rmax - r0);
      s += `<path d="${annulus(cx, cy, r0, r, a0, a1)}" fill="${PALETTE[i % PALETTE.length]}"/>`;
    }
    return `<svg viewBox="0 0 84 84" width="84" height="84" xmlns="${SVG_NS}" role="img" aria-label="Radial bar chart preview">${s}</svg>`;
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"]/g, (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }
})();

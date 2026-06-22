'use strict';

/* =============================================================
   ROUNDED BARS — Tableau Viz Extension
   -------------------------------------------------------------
   Pill-shaped bars (horizontal) or columns (vertical), one per
   category, where each bar's length/height encodes the measure.
   The signature is the COLOR: instead of a flat fill each bar is
   painted with a two-stop LINEAR GRADIENT along (or across) its
   length — uniform, value-driven, a categorical palette, or a
   single color, with optional per-bar color overrides.

   The four moves (see CLAUDE.md):
     1. CONNECT   — initializeAsync(), grab the worksheet
     2. ENCODINGS — Category (dimension) + Value (measure)
     3. DATA      — read summary data → one item per category
     4. DRAW      — render(model) into #viz as an SVG; redraw on change
   ============================================================= */

(function () {
  const host = document.getElementById('viz');
  const STORAGE_KEY = 'roundedbars-config-v1';
  const SVGNS = 'http://www.w3.org/2000/svg';

  const PALETTES = {
    cool:      ['#3d7bf4', '#c13bd6', '#5a5a9e', '#27a0c4', '#8e44ad', '#4cc1a3', '#7b8794', '#6b5bd0', '#d081e0', '#b9bcc4'],
    tableau10: ['#4e79a7', '#f28e2b', '#e15759', '#76b7b2', '#59a14f', '#edc948', '#b07aa1', '#ff9da7', '#9c755f', '#bab0ac'],
    bold:      ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'],
    pastel:    ['#a8d0f0', '#a5d6a7', '#ef9a9a', '#ce93d8', '#ffcc80', '#80cbc4', '#f8bbd0', '#fff176', '#b0bec5', '#c5e1a5'],
    warm:      ['#e15759', '#f28e2b', '#edc948', '#b07aa1', '#ff9da7', '#9c755f', '#d4a35a', '#c44e52', '#dd8452', '#937860'],
  };

  /* ---------- config ---------- */
  const DEFAULT_CONFIG = {
    // identity / theme (light by default; switch bgColor to a dark color to make
    // the gradient bars pop on dark)
    title: '',
    subtitle: '',
    bgColor: '#ffffff',
    textColor: '#1c2330',
    mutedColor: '#6b7280',
    fontFamily: 'system',
    enableTooltips: true,

    // layout
    orientation: 'horizontal',  // horizontal (bars) | vertical (columns)
    sort: 'desc',               // desc | asc | none
    fitToPane: true,            // fit all bars to the cross dimension (no scroll)
    barSlot: 34,                // px per bar when NOT fitting (then #viz scrolls)
    barThickness: 0.62,         // bar thickness as a fraction of its slot
    cornerRadius: 1,            // 0..1 of half-thickness (1 = full pill)
    barMinLength: 6,            // keep tiny values visible

    // COLOR — the signature gradient system
    colorMode: 'gradient',      // gradient | gradient-value | divergent | palette | single
    gradientStart: '#173a5e',   // gradient low end (dark)
    gradientEnd: '#9fd4e6',     // gradient high end (light)
    gradientDirection: 'along', // along | across | diagonal
    divLowColor: '#e7e8f1',     // divergent: low-value (solid) color
    divHighColor: '#2f2f63',    // divergent: high-value (solid) color
    singleColor: '#5a8fc7',
    palette: 'cool',            // cool | tableau10 | bold | pastel | warm
    // per-bar color overrides { categoryKey: '#hex' } + first-seen order — used
    // ONLY in the categorical palette mode (identity-stable hues + overrides);
    // gradient / single define color globally and ignore these. (See CLAUDE.md.)
    seriesColors: {},
    seriesOrder: [],

    // track behind each bar
    showTrack: false,
    trackColor: '#1b2230',

    // labels
    showCategoryLabels: true,
    categoryLabelSize: 13,
    labelWidth: 200,            // left gutter for category names (horizontal)
    categoryLabelMode: 'rotate', // vertical category labels: rotate | stagger | flat
    showValueLabels: true,
    valueLabelSize: 12,
    valueLabelPlacement: 'auto', // auto (inside/outside) | outside | inside

    // group demarcation (only when several Category pills → nested groups):
    // extra space + a divider line between top-level groups.
    groupGap: 14,
    groupSeparator: true,
    groupSeparatorColor: '#9b9ea2',

    // value-number format
    numPrefix: '',
    numSuffix: '',
    numDecimals: 2,
    numUnit: 'auto',            // none | auto | K | M | B
    numThousands: true,
  };

  const FONT_STACKS = {
    system: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
    serif: 'Georgia, "Times New Roman", serif',
    mono: 'ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace',
  };

  let config = { ...DEFAULT_CONFIG };
  let lastModel = null;

  // 1 — CONNECT ------------------------------------------------
  tableau.extensions.initializeAsync({ configure: openConfigModal }).then(
    () => {
      config = loadConfig();
      const ws = tableau.extensions.worksheetContent.worksheet;
      ws.addEventListener(tableau.TableauEventType.SummaryDataChanged, update);
      if (tableau.TableauEventType.MarkSelectionChanged) {
        ws.addEventListener(tableau.TableauEventType.MarkSelectionChanged, async () => {
          if (!lastModel) return;
          await syncSelectionFromWorksheet(lastModel);
          applySelectionStyles();
        });
      }
      // resize: ResizeObserver + window resize + size-poll backstop, rAF-coalesced
      if (window.ResizeObserver) new ResizeObserver(scheduleRender).observe(host);
      window.addEventListener('resize', scheduleRender);
      startSizeWatch();

      wireModal();
      wireInteractivity();
      update();
    },
    (err) => showMessage('Could not initialize: ' + escapeHtml(String(err)))
  );

  // 2 + 3 — ENCODINGS + DATA → one item per category -----------
  async function update() {
    const ws = tableau.extensions.worksheetContent.worksheet;
    let fields = {};
    try { fields = await getEncodedFields(ws); }
    catch (e) { console.error('reading encodings failed:', e); }

    if (!fields.category || !fields.category.length || !fields.value || !fields.value.length) {
      lastModel = null;
      showEmptyState(fields);
      return;
    }
    let table = { columns: [], data: [] };
    try { table = await readSummary(ws); }
    catch (e) { console.error('reading summary data failed:', e); }

    lastModel = buildModel(ws.name, fields, table);
    await syncSelectionFromWorksheet(lastModel);
    render(lastModel);
  }

  // Category can hold SEVERAL pills → one bar per unique COMBINATION of their
  // values (rows are aggregated by a composite key). SEP is a NUL built at
  // runtime (never a literal control byte in source — see CLAUDE.md).
  const SEP = String.fromCharCode(0);
  function joinKey(values) { return values.map(String).join(SEP); }

  function buildModel(worksheet, fields, table) {
    const cols = table.columns;
    const catFields = fields.category;          // [outer, …, inner]
    const valField = fields.value[0];
    const catIdxs = catFields.map((n) => findColumn(cols, n));
    const valIdx = findValueColumn(cols, valField);

    const pos = new Map();        // composite key → item
    const items = [];
    const tupleKeys = {};         // tupleId → composite RAW key (for selection sync)
    table.data.forEach((row, idx) => {
      const tupleId = idx + 1;    // tuple ids are 1-based
      const parts = catIdxs.map((i) => cellText(row[i]));
      const key = parts.join(SEP);
      const rawKey = joinKey(catIdxs.map((i) => (row[i] ? row[i].value : '')));
      const raw = row[valIdx] ? Number(row[valIdx].value) : NaN;
      const v = Number.isFinite(raw) ? raw : null;

      let it = pos.get(key);
      if (!it) {
        it = { key, rawKey, parts, label: parts.join(' / '), value: null, tupleIds: [] };
        pos.set(key, it);
        items.push(it);
      }
      if (v != null) { it.value = (it.value || 0) + v; it.tupleIds.push(tupleId); tupleKeys[tupleId] = rawKey; }
    });

    let min = Infinity, max = -Infinity;
    items.forEach((it) => { if (it.value != null) { min = Math.min(min, it.value); max = Math.max(max, it.value); } });
    if (!Number.isFinite(min)) { min = 0; max = 0; }
    return { worksheet, catFields, valField, items, min, max, tupleKeys };
  }

  // Order items for display. With ONE category field this is a flat value sort.
  // With SEVERAL, it's HIERARCHICAL — items sharing an outer-level prefix stay
  // contiguous (so the nested/merged labels form solid groups), groups ordered
  // by their aggregate in the sort direction, leaves by value within a group.
  function sortedItems(model) {
    const items = model.items.slice();
    const levels = (model.catFields || []).length;
    if (levels <= 1) {
      if (config.sort === 'desc') items.sort((a, b) => (b.value == null ? -Infinity : b.value) - (a.value == null ? -Infinity : a.value));
      else if (config.sort === 'asc') items.sort((a, b) => (a.value == null ? Infinity : a.value) - (b.value == null ? Infinity : b.value));
      return items;
    }
    // precompute each prefix's group sum + first-seen order
    const sum = new Map(), seen = new Map();
    let c = 0;
    model.items.forEach((it) => {
      for (let L = 1; L <= levels; L++) {
        const pre = it.parts.slice(0, L).join(SEP);
        sum.set(pre, (sum.get(pre) || 0) + (it.value || 0));
        if (!seen.has(pre)) seen.set(pre, c++);
      }
    });
    const dir = config.sort === 'asc' ? 1 : -1;
    items.sort((a, b) => {
      for (let L = 1; L <= levels; L++) {
        const pa = a.parts.slice(0, L).join(SEP), pb = b.parts.slice(0, L).join(SEP);
        if (pa === pb) continue;                       // same group at this level → go deeper
        if (config.sort === 'none') return seen.get(pa) - seen.get(pb);
        const d = (sum.get(pa) - sum.get(pb)) * dir;    // order groups by aggregate
        return d !== 0 ? d : seen.get(pa) - seen.get(pb);
      }
      return 0;
    });
    return items;
  }

  // Boundaries between LEAF groups: any index where the PARENT prefix (every
  // category level except the innermost, per-bar one) changes. Returns a Map of
  // boundary index → the highest (outermost) level that changed. Works for ANY
  // number of pills: with 3 pills (Category > Sub-Category > inner) a change in
  // Sub-Category separates groups, and a change in Category separates them more
  // strongly. With 1 pill there are no parent levels → no boundaries.
  function groupBoundaries(items, nLevels) {
    const m = new Map();
    for (let i = 1; i < items.length; i++) {
      let hl = -1;
      for (let L = 0; L < nLevels - 1; L++) {
        if (items[i].parts[L] !== items[i - 1].parts[L]) { hl = L; break; }
      }
      if (hl >= 0) m.set(i, hl);
    }
    return m;
  }
  // outer boundaries (smaller hl) get a much larger gap + a far stronger line, so
  // a sub-group never reads like a top-level category. The OUTERMOST boundary
  // (hl 0) gets a full-strength line; inner ones a faint hairline (the gap does
  // most of the separating there).
  function gapWeight(hl, nLevels) { return 1 + (nLevels - 2 - hl) * 1.3; }
  function sepOpacity(hl) { return hl === 0 ? 1 : Math.max(0.16, 0.3 - (hl - 1) * 0.07); }
  function sepWidth(hl) { return hl === 0 ? 1.4 : 1; }

  /* ---------- resize plumbing ---------- */
  let renderRAF = 0;
  function scheduleRender() {
    if (!lastModel || renderRAF) return;
    renderRAF = requestAnimationFrame(() => { renderRAF = 0; render(lastModel); });
  }
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
  const HOST_PAD = 32;
  function measureWidth(wrap) {
    const vpW = (document.documentElement && document.documentElement.clientWidth) || window.innerWidth || 0;
    const wrapW = wrap.clientWidth || 0;
    const vpAvail = Math.max(0, vpW - HOST_PAD);
    const W = wrapW >= vpAvail * 0.75 ? wrapW : vpAvail;
    return Math.max(10, Math.round(W)) || 800;
  }

  /* ===========================================================
     ▼▼▼  BUILD ZONE — render the chart from the model  ▼▼▼
     =========================================================== */
  function render(model) {
    if (!model) return;
    applyTheme();
    host.innerHTML = '';

    const items = sortedItems(model);
    const n = items.length;
    if (!n) { showMessage('No data to display yet.'); return; }
    if (isCategoricalPalette()) ensureSeriesOrder(model);

    const wrap = el('div', 'rb');
    let titleEl = null, subtitleEl = null;
    if (config.title) { titleEl = el('div', 'rb-title', config.title); wrap.appendChild(titleEl); }
    if (config.subtitle) { subtitleEl = el('div', 'rb-subtitle', config.subtitle); wrap.appendChild(subtitleEl); }
    host.appendChild(wrap);

    const W = measureWidth(wrap);
    const availH = (wrap.clientHeight || ((host.clientHeight || 460) - 28));
    const headH = (titleEl ? titleEl.offsetHeight : 0) + (subtitleEl ? subtitleEl.offsetHeight : 0);
    const paneH = Math.max(60, availH - headH);

    const maxVal = Math.max(1e-9, ...items.map((it) => it.value || 0));
    const svg = svgEl('svg', { class: 'rb-svg' });
    const defs = svgEl('defs', {});
    svg.appendChild(defs);

    // Attach the (empty) SVG NOW so getComputedTextLength / getBBox work while we
    // build + measure labels inside renderHorizontal/Vertical (an unattached SVG
    // measures 0, which would put every value label outside and never truncate).
    wrap.appendChild(svg);
    if (config.orientation === 'vertical') renderVertical(model, items, n, maxVal, W, paneH, svg, defs);
    else renderHorizontal(model, items, n, maxVal, W, paneH, svg, defs);

    applySelectionStyles();
  }

  // Trim an SVG <text> to a pixel width, adding an ellipsis (binary search).
  function fitText(textEl, fullText, maxW) {
    textEl.textContent = fullText;
    let w = 0;
    try { w = textEl.getComputedTextLength(); } catch (e) { return; }
    if (w <= maxW || maxW <= 0) return;
    let lo = 0, hi = fullText.length;
    while (lo < hi) {
      const mid = Math.ceil((lo + hi) / 2);
      textEl.textContent = fullText.slice(0, mid) + '…';
      let ww = 0; try { ww = textEl.getComputedTextLength(); } catch (e) { ww = 0; }
      if (ww <= maxW) lo = mid; else hi = mid - 1;
    }
    textEl.textContent = lo > 0 ? fullText.slice(0, lo) + '…' : '…';
  }

  // Widest of a list of strings at a given font size (px). Uses a temp <text> in
  // the live SVG so getComputedTextLength works.
  function maxLabelWidth(svg, labels, fontSize) {
    const t = svgEl('text', { 'font-size': fontSize, x: -9999, y: -9999 });
    svg.appendChild(t);
    let max = 0;
    for (const s of labels) {
      t.textContent = s == null ? '' : s;
      let w = 0; try { w = t.getComputedTextLength(); } catch (e) { w = 0; }
      if (w > max) max = w;
    }
    t.remove();
    return max;
  }

  // ---- horizontal bars (rows) ----
  function renderHorizontal(model, items, n, maxVal, W, paneH, svg, defs) {
    const catGutter = config.showCategoryLabels ? config.labelWidth : 8;
    const valGutter = config.showValueLabels ? Math.max(46, config.valueLabelSize * 4) : 8;
    const pad = { top: 6, bottom: 6, left: catGutter + 8, right: valGutter + 8 };

    const contentH = n * config.barSlot + pad.top + pad.bottom;
    const svgH = config.fitToPane ? paneH : Math.max(paneH, contentH);
    svg.setAttribute('width', W);
    svg.setAttribute('height', svgH);

    const plotTop = pad.top, plotBottom = svgH - pad.bottom;
    const plotLeft = pad.left, plotRight = W - pad.right;
    const barMax = Math.max(8, plotRight - plotLeft);

    // nested category labels: columns outer→inner across the left gutter.
    const nLevels = (model.catFields || []).length;
    const colW = catGutter / Math.max(1, nLevels);

    // group boundaries → gap between leaf groups (shrinks the band); outer
    // boundaries get a bigger gap.
    const bmap = nLevels > 1 ? groupBoundaries(items, nLevels) : new Map();
    let totalGap = 0; bmap.forEach((hl) => { totalGap += config.groupGap * gapWeight(hl, nLevels); });
    const band = (plotBottom - plotTop - totalGap) / n;
    const rowYs = [];
    { let yc = plotTop; for (let i = 0; i < n; i++) { if (bmap.has(i)) yc += config.groupGap * gapWeight(bmap.get(i), nLevels); rowYs.push(yc + band / 2); yc += band; } }

    const valueLabels = [];
    items.forEach((it, i) => {
      const cy = rowYs[i];
      const th = Math.max(2, Math.min(band - 2, band * config.barThickness));
      const len = it.value == null ? 0 : Math.max(config.barMinLength, (it.value / maxVal) * barMax);
      const fill = barFill(it, i, model, defs, true);

      const g = svgEl('g', { class: 'rb-item rb-interactive' });
      g._tupleIds = it.tupleIds;   // all tuples of this category combination
      g._key = it.rawKey;
      g._curColor = repColor(it, i, model);
      svg.appendChild(g);   // attach now so getComputedTextLength works for fitText

      if (config.showTrack) {
        g.appendChild(svgEl('rect', { class: 'rb-track', x: plotLeft, y: cy - th / 2, width: barMax, height: th,
          rx: rxFor(th, barMax), fill: config.trackColor }));
      }
      if (it.value != null) {
        g.appendChild(svgEl('rect', { class: 'rb-bar', x: plotLeft, y: cy - th / 2, width: len, height: th,
          rx: rxFor(th, len), fill: fill }));
      }
      if (config.showCategoryLabels) {
        if (nLevels <= 1) {
          // single dimension → one right-aligned label per row (image-7 style)
          const t = svgEl('text', { class: 'rb-cat', x: plotLeft - 10, y: cy, 'text-anchor': 'end',
            'dominant-baseline': 'middle', 'font-size': config.categoryLabelSize, fill: config.textColor });
          g.appendChild(t);
          fitText(t, it.label, plotLeft - 14);
        } else {
          // several dimensions → the INNERMOST level is per-row, left-aligned in
          // its column (outer levels are merged after the loop).
          const L = nLevels - 1;
          const t = svgEl('text', { class: 'rb-cat', x: L * colW + 2, y: cy, 'text-anchor': 'start',
            'dominant-baseline': 'middle', 'font-size': config.categoryLabelSize, fill: config.textColor });
          g.appendChild(t);
          fitText(t, it.parts[L], colW - 6);
        }
      }
      if (config.showValueLabels && it.value != null) {
        const t = svgEl('text', { class: 'rb-val', y: cy, 'dominant-baseline': 'middle', 'font-size': config.valueLabelSize });
        t.textContent = formatNumber(it.value);
        g.appendChild(t);
        valueLabels.push({ t, end: plotLeft + len, tipColor: g._curColor });
      }
    });

    // group separator lines (in the gaps); outer boundaries are stronger.
    if (nLevels > 1 && config.groupSeparator && bmap.size) {
      const sep = svgEl('g', { class: 'rb-group-sep' });
      bmap.forEach((hl, i) => {
        const y = (rowYs[i - 1] + rowYs[i]) / 2;
        // indent the line to where the changed level's label column starts:
        // outermost (hl 0) → full width; deeper levels start further right.
        const x1 = hl * colW + 2;
        sep.appendChild(svgEl('line', { x1, y1: y, x2: W - 2, y2: y,
          stroke: config.groupSeparatorColor, 'stroke-width': sepWidth(hl), 'stroke-opacity': sepOpacity(hl) }));
      });
      svg.appendChild(sep);
    }

    // merged labels for the OUTER levels (one per group, centered across its
    // rows) → the nested highlight-table look.
    if (config.showCategoryLabels && nLevels > 1) drawMergedLabelsH(svg, items, rowYs, nLevels, colW);

    // place value labels: outside the bar end, or inside (contrast) when they'd
    // overflow the pane — measured now that the SVG is live.
    placeValueLabelsH(valueLabels, W);
  }

  // Merged outer-level category labels: for each level L (0…nLevels-2), group
  // consecutive rows sharing the parts[0..L] prefix and draw the label once,
  // vertically centered across the group. Static (non-interactive) headers.
  function drawMergedLabelsH(svg, items, rowYs, nLevels, colW) {
    const grp = svgEl('g', { class: 'rb-cat-merged' });
    svg.appendChild(grp);   // attach first so fitText can measure
    for (let L = 0; L < nLevels - 1; L++) {
      let start = 0;
      for (let i = 1; i <= items.length; i++) {
        const prev = items[i - 1].parts.slice(0, L + 1).join(SEP);
        const cur = i < items.length ? items[i].parts.slice(0, L + 1).join(SEP) : null;
        if (cur !== prev) {
          const yc = (rowYs[start] + rowYs[i - 1]) / 2;
          const t = svgEl('text', { class: 'rb-cat', x: L * colW + 2, y: yc, 'text-anchor': 'start',
            'dominant-baseline': 'middle', 'font-size': config.categoryLabelSize,
            fill: config.textColor, 'font-weight': L === 0 ? 600 : 400 });
          grp.appendChild(t);
          fitText(t, items[start].parts[L], colW - 6);
          start = i;
        }
      }
    }
  }

  function placeValueLabelsH(labels, W) {
    labels.forEach((L) => {
      // tentatively outside
      L.t.setAttribute('x', L.end + 8);
      L.t.setAttribute('text-anchor', 'start');
      L.t.setAttribute('fill', config.textColor);
      let w = 0;
      try { w = L.t.getBBox().width; } catch (e) { w = 0; }
      const force = config.valueLabelPlacement;
      const overflow = (L.end + 8 + w) > (W - 2);
      const inside = force === 'inside' || (force !== 'outside' && overflow);
      if (inside) {
        L.t.setAttribute('x', L.end - 8);
        L.t.setAttribute('text-anchor', 'end');
        L.t.setAttribute('fill', contrastInk(L.tipColor));
      }
    });
  }

  // ---- vertical columns ----
  function renderVertical(model, items, n, maxVal, W, paneH, svg, defs) {
    const catMode = config.categoryLabelMode;
    const nLevels = (model.catFields || []).length;
    const rowH = config.categoryLabelSize + 9;
    // INNERMOST level uses the chosen mode; each OUTER level adds a merged row.
    // The inner band AUTO-FITS the longest label (a rotated label needs height =
    // its text width), so labels don't truncate by default. The Category-gutter
    // slider caps it, so you can still constrain the label area in vertical mode.
    let innerH;
    if (!config.showCategoryLabels) {
      innerH = 0;
    } else if (catMode === 'rotate') {
      const need = maxLabelWidth(svg, items.map((it) => it.parts[nLevels - 1]), config.categoryLabelSize) + 16;
      innerH = Math.max(24, Math.min(config.labelWidth, need));
    } else if (catMode === 'stagger') {
      innerH = Math.min(config.labelWidth, 48);
    } else {
      innerH = Math.min(config.labelWidth, 28);
    }
    const catGutter = config.showCategoryLabels ? innerH + Math.max(0, nLevels - 1) * rowH : 8;
    // value labels are rotated & usually INSIDE the column → small top gutter,
    // unless forced 'outside' (then reserve room for the upward label).
    const valGutter = config.showValueLabels ? (config.valueLabelPlacement === 'outside' ? 72 : 16) : 6;
    const pad = { top: valGutter + 4, bottom: catGutter + 4, left: 8, right: 8 };

    const contentW = n * config.barSlot + pad.left + pad.right;
    const svgW = config.fitToPane ? W : Math.max(W, contentW);
    const svgH = paneH;
    svg.setAttribute('width', svgW);
    svg.setAttribute('height', svgH);

    const plotLeft = pad.left, plotRight = svgW - pad.right;
    const plotTop = pad.top, plotBottom = svgH - pad.bottom;
    const barMax = Math.max(8, plotBottom - plotTop);

    // group boundaries → gap between leaf groups (shrinks the band); outer
    // boundaries get a bigger gap.
    const bmap = nLevels > 1 ? groupBoundaries(items, nLevels) : new Map();
    let totalGap = 0; bmap.forEach((hl) => { totalGap += config.groupGap * gapWeight(hl, nLevels); });
    const band = (plotRight - plotLeft - totalGap) / n;
    const colXs = [];
    { let xc = plotLeft; for (let i = 0; i < n; i++) { if (bmap.has(i)) xc += config.groupGap * gapWeight(bmap.get(i), nLevels); colXs.push(xc + band / 2); xc += band; } }

    items.forEach((it, i) => {
      const cx = colXs[i];
      const th = Math.max(2, Math.min(band - 2, band * config.barThickness));
      const h = it.value == null ? 0 : Math.max(config.barMinLength, (it.value / maxVal) * barMax);
      const fill = barFill(it, i, model, defs, false);

      const g = svgEl('g', { class: 'rb-item rb-interactive' });
      g._tupleIds = it.tupleIds;   // all tuples of this category combination
      g._key = it.rawKey;
      g._curColor = repColor(it, i, model);
      svg.appendChild(g);   // attach now so getComputedTextLength works for fitText

      if (config.showTrack) {
        g.appendChild(svgEl('rect', { class: 'rb-track', x: cx - th / 2, y: plotTop, width: th, height: barMax,
          rx: rxFor(th, barMax), fill: config.trackColor }));
      }
      if (it.value != null) {
        g.appendChild(svgEl('rect', { class: 'rb-bar', x: cx - th / 2, y: plotBottom - h, width: th, height: h,
          rx: rxFor(th, h), fill: fill }));
      }
      if (config.showValueLabels && it.value != null) {
        // rotated -90° (reads bottom→top) to avoid colliding across narrow
        // columns; inside near the top for tall bars, just above for short ones.
        const t = svgEl('text', { class: 'rb-val', 'font-size': config.valueLabelSize });
        t.textContent = formatNumber(it.value);
        g.appendChild(t);
        let len = 0; try { len = t.getComputedTextLength(); } catch (e) { len = 0; }
        const topY = plotBottom - h;
        const place = config.valueLabelPlacement;
        const inside = place === 'inside' || (place !== 'outside' && (h - 14) >= len);
        const py = inside ? topY + 8 : topY - 8;
        t.setAttribute('x', cx); t.setAttribute('y', py);
        t.setAttribute('text-anchor', inside ? 'end' : 'start');
        t.setAttribute('dominant-baseline', 'central');
        t.setAttribute('transform', `rotate(-90 ${cx} ${py})`);
        t.setAttribute('fill', inside ? contrastInk(g._curColor) : config.textColor);
      }
      if (config.showCategoryLabels) {
        // innermost level per column (uses the chosen mode); outer levels merged
        // below, after the loop.
        const innerLabel = it.parts[nLevels - 1];
        const t = svgEl('text', { class: 'rb-cat', 'font-size': config.categoryLabelSize, fill: config.textColor });
        if (catMode === 'rotate') {
          const ly = plotBottom + 12;
          t.setAttribute('x', cx); t.setAttribute('y', ly);
          t.setAttribute('text-anchor', 'end');
          t.setAttribute('dominant-baseline', 'central');
          t.setAttribute('transform', `rotate(-90 ${cx} ${ly})`);
          g.appendChild(t);
          fitText(t, innerLabel, innerH - 14);      // rotated → inner-row budget
        } else if (catMode === 'stagger') {
          // offset every other label onto a 2nd row (with a connector tick) so
          // horizontal labels stop colliding — doubles the per-row room.
          const lowered = i % 2 === 1;
          const ly = plotBottom + 14 + (lowered ? 16 : 0);
          if (lowered) {
            g.appendChild(svgEl('line', { x1: cx, y1: plotBottom + 3, x2: cx, y2: ly - 10,
              stroke: config.mutedColor, 'stroke-width': 1, 'stroke-opacity': 0.3 }));
          }
          t.setAttribute('x', cx); t.setAttribute('y', ly);
          t.setAttribute('text-anchor', 'middle');
          t.setAttribute('dominant-baseline', 'hanging');
          g.appendChild(t);
          fitText(t, innerLabel, 2 * band - 8);      // two rows → ~double the room
        } else {                                     // flat (single horizontal row)
          t.setAttribute('x', cx); t.setAttribute('y', plotBottom + 14);
          t.setAttribute('text-anchor', 'middle');
          t.setAttribute('dominant-baseline', 'hanging');
          g.appendChild(t);
          fitText(t, innerLabel, band - 4);
        }
      }
    });

    // group separator lines (vertical, in the gaps); outer boundaries stronger.
    if (nLevels > 1 && config.groupSeparator && bmap.size) {
      const sep = svgEl('g', { class: 'rb-group-sep' });
      bmap.forEach((hl, i) => {
        const x = (colXs[i - 1] + colXs[i]) / 2;
        // extend down through the bars + inner labels, then only as far as the
        // changed level's merged row (outermost reaches the lowest row).
        const y2 = plotBottom + innerH + (nLevels - 1 - hl) * rowH;
        sep.appendChild(svgEl('line', { x1: x, y1: plotTop, x2: x, y2,
          stroke: config.groupSeparatorColor, 'stroke-width': sepWidth(hl), 'stroke-opacity': sepOpacity(hl) }));
      });
      svg.appendChild(sep);
    }

    // merged outer-level labels in rows below the inner labels (nested look).
    if (config.showCategoryLabels && nLevels > 1) {
      drawMergedLabelsV(svg, items, colXs, band, plotBottom, innerH, rowH, nLevels);
    }
  }

  // Merged outer-level labels for vertical columns: each outer level is a row
  // below the inner labels; consecutive columns sharing a prefix are grouped and
  // the label is drawn once, centered across the group's columns (gap-aware via
  // colXs).
  function drawMergedLabelsV(svg, items, colXs, band, plotBottom, innerH, rowH, nLevels) {
    const grp = svgEl('g', { class: 'rb-cat-merged' });
    svg.appendChild(grp);
    for (let L = 0; L < nLevels - 1; L++) {
      const yc = plotBottom + innerH + (nLevels - 2 - L) * rowH + rowH * 0.5;
      let start = 0;
      for (let i = 1; i <= items.length; i++) {
        const prev = items[i - 1].parts.slice(0, L + 1).join(SEP);
        const cur = i < items.length ? items[i].parts.slice(0, L + 1).join(SEP) : null;
        if (cur !== prev) {
          const xc = (colXs[start] + colXs[i - 1]) / 2;
          const runW = (colXs[i - 1] - colXs[start]) + band;
          const t = svgEl('text', { class: 'rb-cat', x: xc, y: yc, 'text-anchor': 'middle',
            'dominant-baseline': 'central', 'font-size': config.categoryLabelSize,
            fill: config.textColor, 'font-weight': L === 0 ? 600 : 400 });
          grp.appendChild(t);
          fitText(t, items[start].parts[L], runW - 8);
          start = i;
        }
      }
    }
  }

  function rxFor(th, len) {
    return Math.max(0, Math.min((th / 2) * clamp(config.cornerRadius, 0, 1), len / 2, th / 2));
  }

  /* ---------- color / gradient ---------- */
  function isCategoricalPalette() { return config.colorMode === 'palette'; }
  // Per-bar color overrides are a categorical-palette feature only — in the
  // gradient / single modes the color is defined globally, so we hide the
  // overrides (the "Bar colors" modal section, the double-click picker, and the
  // override lookup) outside palette mode.
  function overridesAllowed() { return config.colorMode === 'palette'; }

  function ensureSeriesOrder(model) {
    const order = config.seriesOrder || (config.seriesOrder = []);
    let changed = false;
    model.items.forEach((it) => { if (order.indexOf(it.rawKey) < 0) { order.push(it.rawKey); changed = true; } });
    if (changed) saveConfig();
  }
  function stableIndex(key) {
    const order = config.seriesOrder || [];
    const i = order.indexOf(key);
    return i < 0 ? 0 : i;
  }

  // Resolve a bar's FILL. Returns a solid color OR a `url(#id)` referencing a
  // gradient appended to `defs`. `horiz` orients the gradient along the bar.
  function barFill(it, i, model, defs, horiz) {
    const ov = overridesAllowed() && config.seriesColors && config.seriesColors[it.rawKey];
    if (ov) return ov;
    if (config.colorMode === 'single') return config.singleColor;
    if (config.colorMode === 'divergent') return divColor(it, model);   // solid, by value
    if (config.colorMode === 'palette') {
      const pal = PALETTES[config.palette] || PALETTES.cool;
      return pal[stableIndex(it.rawKey) % pal.length];
    }
    // gradient / gradient-value → a per-bar <linearGradient>
    let c0 = config.gradientStart, c1 = config.gradientEnd;
    if (config.colorMode === 'gradient-value') {
      const t = (model.max <= model.min || it.value == null) ? 1 : (it.value - model.min) / (model.max - model.min);
      c1 = mix(config.gradientStart, config.gradientEnd, Math.max(0.06, t));  // end brightens with value
    }
    const id = 'rb-g-' + i;
    defs.appendChild(linearGradient(id, c0, c1, gradCoords(horiz)));
    return 'url(#' + id + ')';
  }
  // representative solid color of a bar (for the contrast of inside labels and
  // the per-bar color-picker seed).
  function repColor(it, i, model) {
    const ov = overridesAllowed() && config.seriesColors && config.seriesColors[it.rawKey];
    if (ov) return ov;
    if (config.colorMode === 'single') return config.singleColor;
    if (config.colorMode === 'divergent') return divColor(it, model);
    if (config.colorMode === 'palette') {
      const pal = PALETTES[config.palette] || PALETTES.cool;
      return pal[stableIndex(it.rawKey) % pal.length];
    }
    if (config.colorMode === 'gradient-value') {
      const t = (model.max <= model.min || it.value == null) ? 1 : (it.value - model.min) / (model.max - model.min);
      return mix(config.gradientStart, config.gradientEnd, Math.max(0.06, t));
    }
    return config.gradientEnd;
  }
  // Solid divergent color: interpolate divLow→divHigh by the bar's value.
  function divColor(it, model) {
    const t = (model.max <= model.min || it.value == null) ? 0.5 : (it.value - model.min) / (model.max - model.min);
    return mix(config.divLowColor, config.divHighColor, Math.max(0, Math.min(1, t)));
  }
  function gradCoords(horiz) {
    const d = config.gradientDirection;
    if (d === 'diagonal') return horiz ? [0, 1, 1, 0] : [0, 1, 1, 0];
    if (d === 'across') return horiz ? [0, 0, 0, 1] : [0, 0, 1, 0];
    // along
    return horiz ? [0, 0, 1, 0] : [0, 1, 0, 0];
  }
  function linearGradient(id, c0, c1, coords) {
    const g = svgEl('linearGradient', { id, gradientUnits: 'objectBoundingBox',
      x1: coords[0], y1: coords[1], x2: coords[2], y2: coords[3] });
    g.appendChild(svgEl('stop', { offset: '0', 'stop-color': c0 }));
    g.appendChild(svgEl('stop', { offset: '1', 'stop-color': c1 }));
    return g;
  }
  function mix(hexA, hexB, t) {
    const a = hexToRgb(hexA), b = hexToRgb(hexB);
    const c = a.map((ch, i) => Math.round(ch + (b[i] - ch) * t));
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
    return lum > 0.6 ? '#10151c' : '#ffffff';
  }
  function anyToHex(c) {
    if (!c) return '#888888';
    const s = String(c);
    if (s[0] === '#') return s.length === 4 ? '#' + s.slice(1).split('').map((x) => x + x).join('') : s;
    const m = s.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
    if (m) return '#' + [1, 2, 3].map((k) => (+m[k]).toString(16).padStart(2, '0')).join('');
    return '#888888';
  }
  function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }

  /* ---------- number formatting ---------- */
  const UNIT_DIV = { K: 1e3, M: 1e6, B: 1e9 };
  function abbreviate(n) {
    const abs = Math.abs(n);
    if (abs >= 1e9) return [n / 1e9, 'B'];
    if (abs >= 1e6) return [n / 1e6, 'M'];
    if (abs >= 1e3) return [n / 1e3, 'K'];
    return [n, ''];
  }
  function withThousands(str) {
    const parts = str.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts.join('.');
  }
  function formatNumber(v) {
    if (v == null || !Number.isFinite(v)) return '';
    let n = v, unit = '';
    if (config.numUnit === 'auto') { [n, unit] = abbreviate(n); }
    else if (UNIT_DIV[config.numUnit]) { n = n / UNIT_DIV[config.numUnit]; unit = config.numUnit; }
    let str = n.toFixed(config.numDecimals);
    if (config.numThousands) str = withThousands(str);
    return config.numPrefix + str + unit + config.numSuffix;
  }

  /* =========================================================
     INTERACTIVITY — native tooltips + selection + marquee +
     double-click per-bar color
     ========================================================= */
  const SELECT_SIMPLE = (tableau.SelectOptions && tableau.SelectOptions.Simple) || 'select-options-simple';
  let selectedTuples = new Set();
  let hoverEl = null, hoverAt = 0;
  let clickTimer = null;
  const MARQUEE_THRESHOLD = 4;
  let dragStart = null, dragCur = null;
  let marqueeActive = false, marqueeAdditive = false, suppressClick = false;

  function getWorksheet() {
    return tableau.extensions.worksheetContent && tableau.extensions.worksheetContent.worksheet;
  }
  function hoverTuple(tupleId, ev) {
    const ws = getWorksheet();
    if (!ws || !ws.hoverTupleAsync) return;
    const tip = ev ? { tooltipAnchorPoint: { x: ev.clientX, y: ev.clientY } } : undefined;
    try { ws.hoverTupleAsync(tupleId, tip).catch(() => {}); } catch (e) { /* ignore */ }
  }
  function clearHoverState() { if (hoverEl) { hoverTuple(0); hoverEl = null; } }
  function pushSelection() {
    const ws = getWorksheet();
    if (ws && ws.selectTuplesAsync) {
      try { ws.selectTuplesAsync([...selectedTuples], SELECT_SIMPLE).catch(() => {}); } catch (e) { /* ignore */ }
    }
    applySelectionStyles();
  }
  function pickBarColor(key, currentColor) {
    const inp = document.createElement('input');
    inp.type = 'color';
    inp.value = anyToHex(currentColor);
    inp.style.cssText = 'position:fixed;left:-9999px;top:0';
    document.body.appendChild(inp);
    inp.addEventListener('input', () => { config.seriesColors[key] = inp.value; saveConfig(); if (lastModel) render(lastModel); });
    inp.addEventListener('change', () => { inp.remove(); });
    inp.click();
  }

  function wireInteractivity() {
    host.addEventListener('mousemove', (ev) => {
      if (!config.enableTooltips || marqueeActive) return;
      const t = ev.target.closest && ev.target.closest('.rb-interactive');
      if (!t) { clearHoverState(); return; }
      const now = (window.performance && performance.now()) || Date.now();
      if (t !== hoverEl || now - hoverAt > 50) { hoverEl = t; hoverAt = now; hoverTuple(t._tupleIds[0], ev); }
    });
    host.addEventListener('mouseleave', clearHoverState);
    host.addEventListener('click', (ev) => {
      if (!config.enableTooltips) return;
      if (suppressClick) { suppressClick = false; return; }
      const t = ev.target.closest && ev.target.closest('.rb-interactive');
      if (t && overridesAllowed()) {
        // delay so a double-click (color pick) can cancel the select
        if (clickTimer) return;
        const ids = t._tupleIds, additive = ev.ctrlKey || ev.metaKey;
        clickTimer = setTimeout(() => { clickTimer = null; toggleSelection(ids, additive); }, 220);
      } else {
        toggleSelection(t && t._tupleIds ? t._tupleIds : null, ev.ctrlKey || ev.metaKey);
      }
    });
    host.addEventListener('dblclick', (ev) => {
      if (!config.enableTooltips || !overridesAllowed()) return;
      const t = ev.target.closest && ev.target.closest('.rb-interactive');
      if (!t || !t._key) return;
      if (clickTimer) { clearTimeout(clickTimer); clickTimer = null; }
      ev.preventDefault();
      pickBarColor(t._key, t._curColor);
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
      if (Math.abs(dragCur.x - dragStart.x) <= MARQUEE_THRESHOLD && Math.abs(dragCur.y - dragStart.y) <= MARQUEE_THRESHOLD) return;
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
    return { left: Math.min(dragStart.x, dragCur.x), top: Math.min(dragStart.y, dragCur.y),
      right: Math.max(dragStart.x, dragCur.x), bottom: Math.max(dragStart.y, dragCur.y) };
  }
  function applyMarquee(rect, additive) {
    const hits = new Set();
    host.querySelectorAll('.rb-bar').forEach((node) => {
      const r = node.getBoundingClientRect();
      const overlaps = !(r.right < rect.left || r.left > rect.right || r.bottom < rect.top || r.top > rect.bottom);
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
    if (!marqueeEl) { marqueeEl = el('div', 'rb-marquee'); marqueeEl.style.display = 'none'; document.body.appendChild(marqueeEl); }
    return marqueeEl;
  }
  function drawMarquee() {
    const r = marqueeRect(), m = ensureMarquee();
    m.style.left = r.left + 'px'; m.style.top = r.top + 'px';
    m.style.width = (r.right - r.left) + 'px'; m.style.height = (r.bottom - r.top) + 'px';
  }
  function hideMarquee() { if (marqueeEl) marqueeEl.style.display = 'none'; }

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
  function applySelectionStyles() {
    const active = selectedTuples.size > 0;
    host.querySelectorAll('.rb-interactive').forEach((node) => {
      if (!active) { node.classList.remove('rb-dimmed', 'rb-selected'); return; }
      const sel = node._tupleIds.every((id) => selectedTuples.has(id));
      node.classList.toggle('rb-selected', sel);
      node.classList.toggle('rb-dimmed', !sel);
    });
  }
  async function syncSelectionFromWorksheet(model) {
    const ws = getWorksheet();
    if (!ws || !ws.getSelectedMarksAsync || !model) return;
    try {
      const result = await ws.getSelectedMarksAsync();
      const tables = (result && result.data) || [];
      // Build the same composite RAW key from EVERY category field, then map
      // back to our tuple ids via model.tupleKeys (collision-free).
      const keys = new Set();
      tables.forEach((tbl) => {
        const cols = tbl.columns || [];
        const idxs = model.catFields.map((n) => cols.findIndex((c) => c.fieldName === n));
        if (idxs.some((i) => i < 0)) return;
        (tbl.data || []).forEach((row) => keys.add(joinKey(idxs.map((i) => row[i].value))));
      });
      const next = new Set();
      Object.keys(model.tupleKeys).forEach((tid) => { if (keys.has(model.tupleKeys[tid])) next.add(Number(tid)); });
      selectedTuples = next;
    } catch (e) { /* leave as-is */ }
  }

  /* ---------- plumbing (encodings + summary data) ---------- */
  async function getEncodedFields(ws) {
    const spec = await ws.getVisualSpecificationAsync();
    const marks = spec.marksSpecifications[spec.activeMarksSpecificationIndex];
    const fields = {};
    for (const enc of marks.encodings) {
      if (enc.field) (fields[enc.id] || (fields[enc.id] = [])).push(enc.field.name);
    }
    return fields;
  }
  async function readSummary(ws) {
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
    document.documentElement.style.setProperty('--rb-font', FONT_STACKS[config.fontFamily] || FONT_STACKS.system);
    document.documentElement.style.setProperty('--rb-ink', config.textColor);
    document.documentElement.style.setProperty('--rb-muted', config.mutedColor);
    document.documentElement.style.setProperty('--rb-bg', config.bgColor);
  }

  /* =========================================================
     CONFIG MODAL (declarative schema)
     ========================================================= */
  const SCHEMA = [
    { section: 'General', fields: [
      { key: 'title', label: 'Chart title', type: 'text' },
      { key: 'subtitle', label: 'Subtitle', type: 'text' },
      { key: 'bgColor', label: 'Background', type: 'color' },
      { key: 'textColor', label: 'Text color', type: 'color' },
      { key: 'mutedColor', label: 'Muted color', type: 'color' },
      { key: 'fontFamily', label: 'Font', type: 'select',
        options: [['system', 'System'], ['serif', 'Serif'], ['mono', 'Monospace']] },
      { key: 'enableTooltips', label: 'Tooltips + selection', type: 'checkbox' },
    ]},
    { section: 'Layout', fields: [
      { key: 'orientation', label: 'Orientation', type: 'select',
        options: [['horizontal', 'Horizontal bars'], ['vertical', 'Vertical columns']] },
      { key: 'sort', label: 'Sort', type: 'select',
        options: [['desc', 'Value ↓'], ['asc', 'Value ↑'], ['none', 'Data order']] },
      { key: 'fitToPane', label: 'Fit all bars to pane', type: 'checkbox' },
      { key: 'barSlot', label: 'Bar slot (px) — when not fitting', type: 'range', min: 14, max: 80, step: 2 },
      { key: 'barThickness', label: 'Bar thickness', type: 'range', min: 0.15, max: 0.98, step: 0.02 },
      { key: 'cornerRadius', label: 'Corner radius (pill)', type: 'range', min: 0, max: 1, step: 0.05 },
      { key: 'barMinLength', label: 'Min bar length (px)', type: 'range', min: 0, max: 30, step: 1 },
    ]},
    { section: 'Color', fields: [
      { key: 'colorMode', label: 'Color mode', type: 'select',
        options: [['gradient', 'Gradient (uniform)'], ['gradient-value', 'Gradient by value'], ['divergent', 'Divergent (by value)'], ['palette', 'Palette (by category)'], ['single', 'Single color']] },
      { key: 'gradientStart', label: 'Gradient start', type: 'color' },
      { key: 'gradientEnd', label: 'Gradient end', type: 'color' },
      { key: 'gradientDirection', label: 'Gradient direction', type: 'select',
        options: [['along', 'Along bar'], ['across', 'Across bar'], ['diagonal', 'Diagonal']] },
      { key: 'divLowColor', label: 'Divergent low (value)', type: 'color' },
      { key: 'divHighColor', label: 'Divergent high (value)', type: 'color' },
      { key: 'palette', label: 'Palette', type: 'select',
        options: [['cool', 'Cool'], ['tableau10', 'Tableau 10'], ['bold', 'Bold'], ['pastel', 'Pastel'], ['warm', 'Warm']] },
      { key: 'singleColor', label: 'Single color', type: 'color' },
    ]},
    { section: 'Track', fields: [
      { key: 'showTrack', label: 'Show track behind bars', type: 'checkbox' },
      { key: 'trackColor', label: 'Track color', type: 'color' },
    ]},
    { section: 'Labels', fields: [
      { key: 'showCategoryLabels', label: 'Show category labels', type: 'checkbox' },
      { key: 'categoryLabelSize', label: 'Category size (px)', type: 'range', min: 8, max: 22, step: 1 },
      { key: 'labelWidth', label: 'Category gutter (px)', type: 'range', min: 60, max: 360, step: 10 },
      { key: 'categoryLabelMode', label: 'Category labels (vertical)', type: 'select',
        options: [['rotate', 'Rotated'], ['stagger', 'Offset rows'], ['flat', 'Flat']] },
      { key: 'showValueLabels', label: 'Show value labels', type: 'checkbox' },
      { key: 'valueLabelSize', label: 'Value size (px)', type: 'range', min: 8, max: 22, step: 1 },
      { key: 'valueLabelPlacement', label: 'Value placement', type: 'select',
        options: [['auto', 'Auto'], ['outside', 'Outside'], ['inside', 'Inside']] },
    ]},
    { section: 'Groups (multi-pill)', fields: [
      { key: 'groupGap', label: 'Gap between groups (px)', type: 'range', min: 0, max: 48, step: 2 },
      { key: 'groupSeparator', label: 'Separator line', type: 'checkbox' },
      { key: 'groupSeparatorColor', label: 'Separator color', type: 'color' },
    ]},
    { section: 'Number format', fields: [
      { key: 'numPrefix', label: 'Prefix', type: 'text' },
      { key: 'numSuffix', label: 'Suffix', type: 'text' },
      { key: 'numDecimals', label: 'Decimals', type: 'range', min: 0, max: 4, step: 1 },
      { key: 'numUnit', label: 'Unit', type: 'select',
        options: [['none', 'None'], ['auto', 'Auto'], ['K', 'Thousands (K)'], ['M', 'Millions (M)'], ['B', 'Billions (B)']] },
      { key: 'numThousands', label: 'Thousands separator', type: 'checkbox' },
    ]},
  ];

  function openConfigModal() {
    const overlay = document.getElementById('cfg-overlay');
    if (!overlay) return;
    buildModalBody();
    overlay.hidden = false;
  }
  function wireModal() {
    const overlay = document.getElementById('cfg-overlay');
    const openBtn = document.getElementById('cfg-open');
    const doneBtn = document.getElementById('cfg-done');
    const resetBtn = document.getElementById('cfg-reset');
    buildModalBody();
    const close = () => { overlay.hidden = true; clearTimeout(saveTimer); flushWorkbookSettings(); };
    openBtn.addEventListener('click', openConfigModal);
    doneBtn.addEventListener('click', close);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !overlay.hidden) close(); });
    resetBtn.addEventListener('click', () => {
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
    const colorsSec = buildBarColorsSection();
    if (colorsSec) body.appendChild(colorsSec);
  }
  // Dynamic per-bar color overrides — only in the categorical palette mode
  // (gradient / single define color globally, so the section is hidden there).
  function buildBarColorsSection() {
    if (!overridesAllowed() || !lastModel || !lastModel.items.length) return null;
    const items = sortedItems(lastModel);
    const wrap = el('div', 'cfg-section');
    wrap.appendChild(el('h3', 'cfg-section-title', 'Bar colors'));
    items.forEach((it, i) => {
      const row = el('label', 'cfg-field');
      row.appendChild(el('span', 'cfg-label', it.label));
      const box = el('span', 'cfg-color-box');
      const inp = document.createElement('input');
      inp.type = 'color';
      inp.value = anyToHex(repColor(it, i, lastModel));
      inp.addEventListener('input', () => { config.seriesColors[it.rawKey] = inp.value; saveConfig(); if (lastModel) render(lastModel); });
      const reset = el('button', 'cfg-mini', '↺');
      reset.type = 'button'; reset.title = 'Reset to mode color';
      reset.addEventListener('click', (e) => {
        e.preventDefault();
        delete config.seriesColors[it.rawKey]; saveConfig();
        inp.value = anyToHex(repColor(it, i, lastModel));
        if (lastModel) render(lastModel);
      });
      box.appendChild(inp); box.appendChild(reset);
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
      input.type = 'checkbox'; input.checked = !!val;
      input.addEventListener('change', () => commit(f.key, input.checked));
      row.classList.add('cfg-check');
    } else if (f.type === 'color') {
      input = document.createElement('input');
      input.type = 'color'; input.value = val;
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
      input.type = 'range'; input.min = f.min; input.max = f.max; input.step = f.step; input.value = val;
      const out = el('span', 'cfg-range-out', String(val));
      input.addEventListener('input', () => { const num = parseFloat(input.value); out.textContent = String(num); commit(f.key, num); });
      box.appendChild(input); box.appendChild(out);
      row.appendChild(box);
      return row;
    } else {
      input = document.createElement('input');
      input.type = 'text'; input.value = val == null ? '' : val;
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

  /* ---------- config persistence (workbook settings only) ---------- */
  function loadConfig() {
    let saved = {};
    try {
      const store = tableau.extensions.settings;
      const raw = store && store.get(STORAGE_KEY);
      if (raw) saved = JSON.parse(raw);
    } catch (e) { console.error('reading workbook settings failed:', e); }
    const cfg = { ...DEFAULT_CONFIG, ...saved };
    cfg.seriesColors = { ...(saved.seriesColors || {}) };
    cfg.seriesOrder = [...(saved.seriesOrder || [])];
    return cfg;
  }
  let saveTimer = null, saveInFlight = false, savePending = false;
  function saveConfig() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(flushWorkbookSettings, 350);
  }
  function flushWorkbookSettings() {
    const store = tableau.extensions.settings;
    if (!store) return;
    try { store.set(STORAGE_KEY, JSON.stringify(config)); }
    catch (e) { console.error('settings.set failed:', e); return; }
    if (saveInFlight) { savePending = true; return; }
    saveInFlight = true;
    store.saveAsync().then(
      () => { saveInFlight = false; if (savePending) { savePending = false; flushWorkbookSettings(); } },
      (err) => { saveInFlight = false; console.error('settings.saveAsync failed:', err);
        if (savePending) { savePending = false; flushWorkbookSettings(); } }
    );
  }

  /* ---------- tiny DOM/SVG helpers + states ---------- */
  function el(tag, cls, text) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }
  function svgEl(tag, attrs) {
    const n = document.createElementNS(SVGNS, tag);
    if (attrs) for (const k in attrs) n.setAttribute(k, attrs[k]);
    return n;
  }
  function showMessage(html) {
    applyTheme();
    host.innerHTML = '<div class="rb-empty"><p>' + html + '</p></div>';
  }
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
            '<div class="rb-empty-title">Rounded Bars</div>' +
            '<div class="rb-empty-desc">Pill-shaped bars or columns, one per category, each painted with a ' +
            'linear <b>gradient</b> along its length. Switch orientation, gradient, palette, or set a color per ' +
            'bar — all from the &#9881; gear or the <b>Format Extension</b> button.</div>' +
          '</div>' +
        '</div>' +
        '<div class="rb-empty-guide">' +
          '<div class="rb-empty-guide-title">Get started</div>' +
          '<div class="rb-empty-guide-row">Drag one or more <b>dimensions</b> onto ' +
            chip('category', 'Category', 'dim') + ' and a <b>measure</b> onto ' +
            chip('value', 'Value', 'measure') + '.' +
          '</div>' +
        '</div>' +
      '</div></div>';
  }
  // Inline SVG thumbnail: a few rounded gradient bars.
  function buildThumb() {
    const w = [70, 54, 40, 30, 20], cols = ['#173a5e', '#9fd4e6'];
    let bars = '<defs><linearGradient id="rbt" x1="0" y1="0" x2="1" y2="0">' +
      `<stop offset="0" stop-color="${cols[0]}"/><stop offset="1" stop-color="${cols[1]}"/></linearGradient></defs>`;
    w.forEach((bw, i) => { const y = 8 + i * 14; bars += `<rect x="6" y="${y}" width="${bw}" height="9" rx="4.5" fill="url(#rbt)"/>`; });
    return `<svg viewBox="0 0 84 84" width="84" height="84" xmlns="${SVGNS}" role="img" aria-label="Rounded bars preview">${bars}</svg>`;
  }
  function escapeHtml(s) {
    return String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }
})();

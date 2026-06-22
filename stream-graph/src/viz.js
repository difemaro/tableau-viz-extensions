'use strict';

/* =============================================================
   STREAM GRAPH — Tableau Viz Extension
   -------------------------------------------------------------
   A streamgraph (a.k.a. ThemeRiver): a stacked area chart whose
   layers flow around a wiggling central baseline, so each series
   reads as an organic "stream" thickening and thinning over an
   ordered axis (e.g. months). Smooth curves, categorical color,
   right-edge series labels.

   The four moves (see CLAUDE.md):
     1. CONNECT   — initializeAsync(), grab the worksheet
     2. ENCODINGS — Axis (x dimension) + Series (color dimension) + Value (measure)
     3. DATA      — read summary data → pivot into a series×axis matrix
     4. DRAW      — render(model) into #viz as an SVG; redraw on change
   ============================================================= */

(function () {
  const host = document.getElementById('viz');
  const STORAGE_KEY = 'streamgraph-config-v1';
  const SVGNS = 'http://www.w3.org/2000/svg';

  /* ---------- config ---------- */
  const DEFAULT_CONFIG = {
    // identity / theme (light by default)
    title: '',
    subtitle: '',
    caption: '',
    bgColor: '#ffffff',
    fontFamily: 'system',
    textColor: '#1c2330',
    mutedColor: '#6b7280',

    // streamgraph shape
    offset: 'wiggle',        // wiggle (classic streamgraph) | silhouette (centered) | zero (stacked) | expand (100%)
    order: 'inside-out',     // inside-out | ascending | descending | reverse | none (data order)
    curve: 'smooth',         // smooth (spline) | linear | step
    curveTension: 1,         // 0 = straight, 1 = full Catmull-Rom curviness (smooth only)

    // bands
    layerOpacity: 1,
    showStroke: true,        // thin separator stroke between layers
    strokeColor: '#ffffff',
    strokeWidth: 1,

    // color
    colorMode: 'palette',    // palette (categorical) | sequential (by total) | single
    palette: 'cool',         // cool | tableau10 | bold | pastel | warm | mono
    monoBaseColor: '#5a8fc7',
    singleColor: '#5a8fc7',
    seqLowColor: '#d6dee8',  // sequential — low total
    seqHighColor: '#3b6ea5', // sequential — high total
    // Per-series color overrides { seriesKey: '#hex' } — double-click a legend
    // item or use the modal's "Series colors" section. Always win over the mode.
    seriesColors: {},
    // First-seen order of series keys (append-only, persisted). Palette colors
    // are assigned by a series' position HERE, not its index in the current
    // (possibly filtered) data — so a category keeps its color when you filter.
    seriesOrder: [],

    // x axis & gridlines
    showXAxis: true,
    staggerLabels: false,    // offset every other label onto a 2nd row (no overlap)
    minLabelSpacing: 48,     // px; thin x labels so they don't collide
    showGridlines: true,
    gridColor: '#e6e8ec',
    showZeroLine: false,     // faint baseline through the stream

    // legend
    showLegend: true,
    legendPosition: 'top',   // top | bottom

    // series labels (right edge)
    showSeriesLabels: false,
    labelSize: 13,
    labelWidth: 92,          // reserved right-hand gutter for labels
    labelMinThickness: 6,    // skip a label only when the band end is thinner than
                             // this (px); close/thin streams are de-collided, not hidden

    // interactivity — native Tableau tooltip on hover + per-cell selection
    enableTooltips: true,
  };

  const FONT_STACKS = {
    system: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
    serif: 'Georgia, "Times New Roman", serif',
    mono: 'ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace',
  };

  const PALETTES = {
    cool:      ['#3d7bf4','#c13bd6','#5a5a9e','#b9bcc4','#27a0c4','#8e44ad','#4cc1a3','#7b8794','#6b5bd0','#d081e0'],
    tableau10: ['#4e79a7','#f28e2b','#e15759','#76b7b2','#59a14f','#edc948','#b07aa1','#ff9da7','#9c755f','#bab0ac'],
    bold:      ['#1f77b4','#ff7f0e','#2ca02c','#d62728','#9467bd','#8c564b','#e377c2','#7f7f7f','#bcbd22','#17becf'],
    pastel:    ['#a8d0f0','#a5d6a7','#ef9a9a','#ce93d8','#ffcc80','#80cbc4','#f8bbd0','#fff176','#b0bec5','#c5e1a5'],
    warm:      ['#e15759','#f28e2b','#edc948','#b07aa1','#ff9da7','#9c755f','#d4a35a','#c44e52','#dd8452','#937860'],
  };

  let config = { ...DEFAULT_CONFIG };
  let lastModel = null;
  // geometry from the last render — used by the hover tooltip to locate the
  // nearest axis column under the cursor.
  let geom = { svg: null, xs: [], axisValues: [] };

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

      // Re-render (sized to the new host width) on any size change — same
      // belt-and-braces approach as the other SVG extensions: ResizeObserver +
      // window resize + a polling backstop, all coalesced through rAF.
      if (window.ResizeObserver) new ResizeObserver(scheduleRender).observe(host);
      window.addEventListener('resize', scheduleRender);
      startSizeWatch();

      wireModal();
      wireInteractivity();
      update();
    },
    (err) => showMessage('Could not initialize: ' + escapeHtml(String(err)))
  );

  // 2 + 3 — ENCODINGS + DATA → series×axis matrix --------------
  async function update() {
    const ws = tableau.extensions.worksheetContent.worksheet;

    let fields = {};
    try { fields = await getEncodedFields(ws); }
    catch (e) { console.error('reading encodings failed:', e); }

    if (!fields.axis || !fields.axis.length ||
        !fields.series || !fields.series.length ||
        !fields.value || !fields.value.length) {
      lastModel = null;
      showEmptyState(fields);
      return;
    }

    let table = { columns: [], data: [] };
    try { table = await readSummary(ws); }
    catch (e) { console.error('reading summary data failed:', e); }

    lastModel = buildModel(ws.name, fields, table);
    if (lastModel.axisValues.length < 2) {
      showMessage('A stream graph needs at least 2 axis values. Add a dimension with more points to the <b>Axis</b> tile.');
      lastModel = null;
      return;
    }
    await syncSelectionFromWorksheet(lastModel);
    render(lastModel);
  }

  // Pivot the flat summary rows into a series×axis value matrix. Each summary
  // row is one (axis, series) cell with the aggregated measure; tuple ids are
  // 1-based row numbers, collected per series for selection.
  function buildModel(worksheet, fields, table) {
    const cols = table.columns;
    const axisField = fields.axis[0];
    const seriesField = fields.series[0];
    const valField = fields.value[0];
    const axisIdx = findColumn(cols, axisField);
    const serIdx = findColumn(cols, seriesField);
    const valIdx = findValueColumn(cols, valField);

    const axisIndex = new Map(), axisValues = [];
    const seriesIndex = new Map(), seriesList = [];
    const cells = new Map();      // "si|ai" -> { value, tupleId }

    table.data.forEach((row, i) => {
      const aRaw = row[axisIdx] ? String(row[axisIdx].value) : '';
      const aVal = row[axisIdx] ? row[axisIdx].value : aRaw;  // native value (Date/number) for sorting
      const aLab = cellText(row[axisIdx]);
      const sRaw = row[serIdx] ? String(row[serIdx].value) : '';
      const sLab = cellText(row[serIdx]);
      const raw = row[valIdx] ? Number(row[valIdx].value) : NaN;
      const value = Number.isFinite(raw) ? raw : 0;

      if (!axisIndex.has(aRaw)) { axisIndex.set(aRaw, axisValues.length); axisValues.push({ raw: aRaw, label: aLab, nativeValue: aVal }); }
      if (!seriesIndex.has(sRaw)) { seriesIndex.set(sRaw, seriesList.length); seriesList.push({ raw: sRaw, label: sLab, tupleIds: [], total: 0 }); }
      const ai = axisIndex.get(aRaw), si = seriesIndex.get(sRaw);
      const tupleId = i + 1;
      cells.set(si + '|' + ai, { value, tupleId });
      seriesList[si].tupleIds.push(tupleId);
      seriesList[si].total += value;
    });

    // Order the axis. For a CONTINUOUS axis (a date or numeric pill) sort by the
    // real value so the streams flow left→right in chronological / numeric order
    // regardless of the row order the data reader returns. For a discrete pill we
    // keep Tableau's own order (e.g. Jan…Dec for a month-name dimension).
    const keyOf = axisSortKey((cols[axisIdx] && cols[axisIdx].dataType || '').toLowerCase());
    let order = axisValues.map((_, i) => i);
    if (keyOf) {
      const keys = axisValues.map((a) => keyOf(a.nativeValue, a.raw));
      if (keys.every((k) => k != null)) order = order.slice().sort((a, b) => keys[a] - keys[b]);
    }

    // Apply the order: rebuild axisValues + axisIndex, then materialize the dense
    // matrix / tuple lookup (missing combos → 0) in the final column order.
    const orderedAxis = order.map((oldAi) => axisValues[oldAi]);
    const finalAxisIndex = new Map();
    orderedAxis.forEach((a, newAi) => finalAxisIndex.set(a.raw, newAi));

    const m = orderedAxis.length, n = seriesList.length;
    const matrix = seriesList.map(() => new Array(m).fill(0));
    const tupleAt = seriesList.map(() => new Array(m).fill(0));
    for (let si = 0; si < n; si++) {
      for (let newAi = 0; newAi < m; newAi++) {
        const c = cells.get(si + '|' + order[newAi]);
        if (c) { matrix[si][newAi] = c.value; tupleAt[si][newAi] = c.tupleId; }
      }
    }

    return { worksheet, axisField, seriesField, valField, axisValues: orderedAxis,
             seriesList, matrix, tupleAt, axisIndex: finalAxisIndex, seriesIndex };
  }

  // Returns a (nativeValue, rawString) → number key for SORTABLE axis types
  // (dates, integers, floats), or null to leave a discrete axis in data order.
  function axisSortKey(dataType) {
    if (dataType.indexOf('date') >= 0) {
      return (v, raw) => { const d = Date.parse(v != null ? v : raw); return Number.isNaN(d) ? null : d; };
    }
    if (dataType === 'int' || dataType === 'integer' || dataType === 'float') {
      return (v, raw) => { const num = Number(v != null ? v : raw); return Number.isFinite(num) ? num : null; };
    }
    return null;
  }

  /* ---------- stacking order & layout ---------- */
  // Return series indices in the chosen stacking order (bottom → top).
  function stackingOrder(model) {
    const n = model.seriesList.length;
    const idx = Array.from({ length: n }, (_, i) => i);
    const totals = model.seriesList.map((s) => s.total);
    if (config.order === 'none') return idx;
    if (config.order === 'reverse') return idx.slice().reverse();
    if (config.order === 'ascending') return idx.slice().sort((a, b) => totals[a] - totals[b]);
    if (config.order === 'descending') return idx.slice().sort((a, b) => totals[b] - totals[a]);
    // inside-out (default): largest layers in the middle — the classic
    // streamgraph order (d3 stackOrderInsideOut).
    const asc = idx.slice().sort((a, b) => totals[a] - totals[b]);
    let top = 0, bottom = 0;
    const tops = [], bottoms = [];
    asc.forEach((j) => {
      if (top < bottom) { top += totals[j]; tops.push(j); }
      else { bottom += totals[j]; bottoms.push(j); }
    });
    return bottoms.reverse().concat(tops);
  }

  // Build per-layer {lower, upper} boundaries (in value space) for the chosen
  // offset mode. `ordered` is the list of series indices, bottom → top.
  function computeLayers(model, ordered) {
    const m = model.axisValues.length;
    let vals = ordered.map((si) => model.matrix[si]);
    let base;

    if (config.offset === 'expand') {
      vals = vals.map((a) => a.slice());
      for (let i = 0; i < m; i++) {
        let sum = 0; for (let j = 0; j < vals.length; j++) sum += vals[j][i] || 0;
        if (sum) for (let j = 0; j < vals.length; j++) vals[j][i] = (vals[j][i] || 0) / sum;
      }
      base = new Array(m).fill(0);
    } else if (config.offset === 'silhouette') {
      base = new Array(m);
      for (let i = 0; i < m; i++) {
        let sum = 0; for (let j = 0; j < vals.length; j++) sum += vals[j][i] || 0;
        base[i] = -sum / 2;
      }
    } else if (config.offset === 'wiggle') {
      base = wiggleBaseline(vals, m);
    } else { // zero
      base = new Array(m).fill(0);
    }

    const layers = ordered.map(() => ({ lower: new Array(m), upper: new Array(m) }));
    const cum = base.slice();
    for (let j = 0; j < vals.length; j++) {
      for (let i = 0; i < m; i++) {
        layers[j].lower[i] = cum[i];
        cum[i] += vals[j][i] || 0;
        layers[j].upper[i] = cum[i];
      }
    }
    return layers;
  }

  // Minimized-wiggle baseline (Byron & Wattenberg, == d3 stackOffsetWiggle):
  // the bottom of the lowest layer at each column, chosen to minimize the
  // overall "wiggle" so the streams stay as horizontal as possible.
  function wiggleBaseline(vals, m) {
    const n = vals.length;
    const base = new Array(m).fill(0);
    let y = 0;
    for (let i = 1; i < m; i++) {
      let s1 = 0, s2 = 0;
      for (let j = 0; j < n; j++) {
        const sij0 = vals[j][i] || 0, sij1 = vals[j][i - 1] || 0;
        let s3 = (sij0 - sij1) / 2;
        for (let k = 0; k < j; k++) s3 += (vals[k][i] || 0) - (vals[k][i - 1] || 0);
        s1 += sij0; s2 += s3 * sij0;
      }
      if (s1) y -= s2 / s1;
      base[i] = y;
    }
    return base;
  }

  // Coalesce resize-driven re-renders into a single rAF.
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
  const HOST_PAD_X = 32;
  function measureWidth(wrap) {
    const vpW = (document.documentElement && document.documentElement.clientWidth) || window.innerWidth || 0;
    const wrapW = wrap.clientWidth || 0;
    const vpAvail = Math.max(0, vpW - HOST_PAD_X);
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

    const m = model.axisValues.length;
    const n = model.seriesList.length;
    if (m < 2 || !n) { showMessage('No data to display yet.'); return; }

    // color domain (series totals) — shared by the bands and the legend.
    const totals = model.seriesList.map((s) => s.total);
    const loT = Math.min(...totals), hiT = Math.max(...totals);
    // lock in stable palette slots per series identity (categorical palette only)
    if (isCategoricalPalette()) ensureSeriesOrder(model);
    const colorOf = (si) => seriesColorFor(model, si, totals, loT, hiT);

    const wrap = el('div', 'sg');
    let titleEl = null, subtitleEl = null;
    if (config.title) { titleEl = el('div', 'sg-title', config.title); wrap.appendChild(titleEl); }
    if (config.subtitle) { subtitleEl = el('div', 'sg-subtitle', config.subtitle); wrap.appendChild(subtitleEl); }
    // Optional legend. Built up front so its height is reserved out of the SVG
    // area; a 'top' legend sits above the chart now, a 'bottom' one is moved
    // below the SVG after it's drawn (it's attached here only to be measured).
    const legendEl = config.showLegend ? buildLegend(model, colorOf) : null;
    if (legendEl) wrap.appendChild(legendEl);
    // Append now so we can measure the laid-out content box (host.clientWidth
    // is unreliable in some Tableau hosts).
    host.appendChild(wrap);
    const captionEl = config.caption ? el('div', 'sg-caption', config.caption) : null;

    // wrap is height:100% (flex column), so wrap.clientHeight is the FULL pane —
    // subtract only the actual title/subtitle/legend/caption element heights to
    // get the height left for the SVG (measuring wrap.clientHeight as "head"
    // would eat it all and flatten the chart).
    const W = measureWidth(wrap);
    const availH = (wrap.clientHeight || ((host.clientHeight || 460) - 28));
    const headH = (titleEl ? titleEl.offsetHeight : 0) + (subtitleEl ? subtitleEl.offsetHeight : 0);
    const legendH = legendEl ? legendEl.offsetHeight : 0;
    const captionH = config.caption ? 22 : 0;
    const svgH = Math.max(80, availH - headH - legendH - captionH);

    // Staggered (offset) labels add a second row, so reserve extra bottom room.
    const stagger = config.showXAxis && config.staggerLabels;
    const pad = {
      top: 10,
      right: 12 + (config.showSeriesLabels ? config.labelWidth : 0),
      bottom: config.showXAxis ? (stagger ? 44 : 28) : 8,
      left: 12,
    };
    const plotLeft = pad.left;
    const plotW = Math.max(10, W - pad.left - pad.right);
    const plotTop = pad.top;
    const plotH = Math.max(10, svgH - pad.top - pad.bottom);

    // x positions for each axis column.
    const xs = [];
    for (let i = 0; i < m; i++) xs.push(plotLeft + (m === 1 ? 0 : (i / (m - 1)) * plotW));

    // stacking + boundaries
    const ordered = stackingOrder(model);
    const layers = computeLayers(model, ordered);

    // value → y. Larger value sits higher (smaller y). Fit the full stack.
    let vMin = Infinity, vMax = -Infinity;
    layers.forEach((L) => {
      for (let i = 0; i < m; i++) { vMin = Math.min(vMin, L.lower[i]); vMax = Math.max(vMax, L.upper[i]); }
    });
    if (!Number.isFinite(vMin) || vMin === vMax) { vMin = 0; vMax = vMax === vMin ? vMax + 1 : vMax; }
    const yScale = (v) => plotTop + (vMax - v) / (vMax - vMin) * plotH;

    const svg = svgEl('svg', { class: 'sg-svg', width: W, height: svgH });

    // which x columns get a label / gridline (thin to avoid collisions).
    // Staggering alternates labels across two rows, so each row only needs HALF
    // the horizontal room → allow labels twice as dense.
    const effSpacing = Math.max(12, config.minLabelSpacing / (stagger ? 2 : 1));
    const step = Math.max(1, Math.ceil(m / Math.max(1, Math.floor(plotW / effSpacing))));
    const showTick = (i) => i % step === 0 || i === m - 1;

    // --- gridlines (behind the streams) ---
    if (config.showGridlines) {
      const g = svgEl('g', { class: 'sg-grid' });
      for (let i = 0; i < m; i++) {
        if (!showTick(i)) continue;
        g.appendChild(svgEl('line', { x1: xs[i], y1: plotTop, x2: xs[i], y2: plotTop + plotH,
          stroke: config.gridColor, 'stroke-width': 1 }));
      }
      svg.appendChild(g);
    }
    if (config.showZeroLine) {
      const y0 = yScale(0);
      if (y0 >= plotTop && y0 <= plotTop + plotH) {
        svg.appendChild(svgEl('line', { x1: plotLeft, y1: y0, x2: plotLeft + plotW, y2: y0,
          stroke: config.mutedColor, 'stroke-width': 1, 'stroke-opacity': 0.4, 'stroke-dasharray': '3 3' }));
      }
    }

    // --- one band per series ---
    const endLabels = [];   // collected, then de-collided + drawn after the loop
    layers.forEach((L, j) => {
      const si = ordered[j];
      const s = model.seriesList[si];
      const color = colorOf(si);

      const topPts = xs.map((x, i) => [x, yScale(L.upper[i])]);
      const botPts = xs.map((x, i) => [x, yScale(L.lower[i])]);
      const d = bandPath(topPts, botPts);

      const g = svgEl('g', { class: 'sg-layer sg-interactive' });
      // per-axis tuple ids for THIS series, so hover/click resolve to a single
      // (series, axis) mark — e.g. a specific fruit-month — not the whole stream.
      g._tupleAt = model.tupleAt[si];

      const path = svgEl('path', { class: 'sg-band', d, fill: color, 'fill-opacity': config.layerOpacity });
      if (config.showStroke) { path.setAttribute('stroke', config.strokeColor); path.setAttribute('stroke-width', config.strokeWidth); }
      g.appendChild(path);

      // Collect the right-edge label candidate (drawn later). Keep the band-end
      // span so we can de-collide labels of thin/close-valued streams instead of
      // hiding them — the ideal y is the band's vertical center at the last col.
      if (config.showSeriesLabels) {
        const yTop = yScale(L.upper[m - 1]), yBot = yScale(L.lower[m - 1]);
        if (Math.abs(yBot - yTop) >= config.labelMinThickness) {
          endLabels.push({ g, color, label: s.label, idealY: (yTop + yBot) / 2 });
        }
      }

      svg.appendChild(g);
    });

    // Draw the right-edge series labels, nudged apart so close end-values don't
    // overlap (each keeps a thin leader back to its stream).
    if (endLabels.length) placeEndLabels(endLabels, xs[m - 1], plotTop, plotH);

    // --- x axis labels ---
    // With staggering ON, every other SHOWN label drops to a second row (with a
    // faint connector tick) so dense/long labels stop overlapping — the classic
    // "offset labels" trick. (A measurement pass below then guarantees no
    // same-row overlap remains, whatever the label widths.)
    const axisLabelItems = [];
    if (config.showXAxis) {
      const ax = svgEl('g', { class: 'sg-axis' });
      const baseY = plotTop + plotH + 16;
      const rowGap = 15;
      let shown = 0;
      for (let i = 0; i < m; i++) {
        if (!showTick(i)) continue;
        const row = stagger && (shown % 2 === 1) ? 1 : 0;
        const y = baseY + row * rowGap;
        const anchor = i === 0 ? 'start' : (i === m - 1 ? 'end' : 'middle');
        let lineEl = null;
        if (row === 1) {
          lineEl = svgEl('line', { x1: xs[i], y1: plotTop + plotH + 3, x2: xs[i], y2: y - 9,
            stroke: config.mutedColor, 'stroke-width': 1, 'stroke-opacity': 0.3 });
          ax.appendChild(lineEl);
        }
        const lab = svgEl('text', { x: xs[i], y, 'text-anchor': anchor, fill: config.mutedColor, 'font-size': 11 });
        lab.textContent = model.axisValues[i].label;
        ax.appendChild(lab);
        axisLabelItems.push({ el: lab, lineEl, row });
        shown++;
      }
      svg.appendChild(ax);
    }

    wrap.appendChild(svg);
    // Move a 'bottom' legend below the SVG (it was attached above only to be
    // measured); a 'top' legend stays where it is.
    if (legendEl && config.legendPosition === 'bottom') wrap.appendChild(legendEl);
    if (captionEl) wrap.appendChild(captionEl);

    // Post-render guarantee: drop any axis label still overlapping the previous
    // KEPT label on the same row (measured in px now that the SVG is live).
    // minLabelSpacing pre-thins for target density; this removes the leftovers
    // that a fixed spacing proxy can't catch with variable-width labels.
    if (axisLabelItems.length) pruneOverlappingLabels(axisLabelItems, 6);

    geom = { svg, xs };
    applySelectionStyles();
  }
  /* ▲▲▲  END BUILD ZONE  ▲▲▲ */

  // Build the HTML legend (swatch + series name per series), colored to match
  // the bands. Lives in the flex column above/below the SVG.
  function buildLegend(model, colorOf) {
    const box = el('div', 'sg-legend ' + (config.legendPosition === 'bottom' ? 'sg-legend-bottom' : 'sg-legend-top'));
    model.seriesList.forEach((s, si) => {
      const item = el('div', 'sg-legend-item');
      const sw = el('span', 'sg-legend-swatch');
      sw.style.background = colorOf(si);
      item.appendChild(sw);
      item.appendChild(el('span', 'sg-legend-text', s.label));
      // Interactive legend:
      //   • single-click  → select that whole series (same path as clicking a
      //     stream): highlights here + on the chart and filters the dashboard;
      //   • double-click  → pick a color for that series (persisted override).
      // Single/double are disambiguated with a short timer so a dbl-click does
      // NOT also toggle the selection. We stop propagation so the host's
      // empty-click "clear selection" doesn't fire. Always allow the color pick;
      // gate only the selection on the interactivity toggle.
      item._tupleIds = s.tupleIds;     // read by applySelectionStyles for dimming
      item.classList.add('sg-legend-click');
      // Double-click colour-pick only makes sense in palette mode (sequential/
      // single ignore per-series overrides); elsewhere the legend just selects.
      const canPick = config.colorMode === 'palette';
      item.title = canPick ? 'Click to select · double-click to set color' : 'Click to select';
      let clickTimer = null;
      item.addEventListener('click', (ev) => {
        ev.stopPropagation();
        if (!config.enableTooltips) return;
        const additive = ev.ctrlKey || ev.metaKey;
        if (!canPick) { toggleSelection(s.tupleIds, additive); return; }   // instant
        if (clickTimer) return;        // 2nd click of a dbl-click → ignore
        clickTimer = setTimeout(() => { clickTimer = null; toggleSelection(s.tupleIds, additive); }, 220);
      });
      item.addEventListener('dblclick', (ev) => {
        ev.preventDefault(); ev.stopPropagation();
        if (!canPick) return;
        if (clickTimer) { clearTimeout(clickTimer); clickTimer = null; }
        pickSeriesColor(s.raw, colorOf(si));
      });
      box.appendChild(item);
    });
    return box;
  }

  // Open the OS color picker for a series, then store the chosen color as a
  // persisted override and re-render. (A hidden <input type=color> is the
  // simplest cross-host picker.)
  function pickSeriesColor(rawKey, currentColor) {
    const inp = document.createElement('input');
    inp.type = 'color';
    inp.value = anyToHex(currentColor);
    inp.style.cssText = 'position:fixed;left:-9999px;top:0';
    document.body.appendChild(inp);
    inp.addEventListener('input', () => {
      config.seriesColors[rawKey] = inp.value;
      saveConfig();
      if (lastModel) render(lastModel);
    });
    inp.addEventListener('change', () => { inp.remove(); });
    inp.click();
  }

  // Place right-edge series labels at their bands' end-centers, but push apart
  // any that would overlap (close end values) so every label stays readable.
  // A label nudged away from its band gets a thin leader curve back to it.
  function placeEndLabels(items, edgeX, plotTop, plotH) {
    const lineH = config.labelSize + 5;
    const labelX = edgeX + 10;
    // greedy down-pack from the top, then shift the whole stack up if it spilled
    // past the bottom, then clamp the top — keeps labels inside the plot band.
    items.sort((a, b) => a.idealY - b.idealY);
    let prev = -Infinity;
    items.forEach((it) => { it.y = Math.max(it.idealY, prev + lineH); prev = it.y; });
    const overflow = prev - (plotTop + plotH);
    if (overflow > 0) items.forEach((it) => { it.y -= overflow; });
    const up = plotTop - items[0].y;
    if (up > 0) items.forEach((it) => { it.y += up; });
    items.forEach((it) => {
      if (Math.abs(it.y - it.idealY) > 1.5) {
        it.g.appendChild(svgEl('path', { class: 'sg-leader', fill: 'none', stroke: it.color,
          'stroke-width': 1, 'stroke-opacity': 0.5,
          d: `M ${edgeX + 1} ${it.idealY} C ${edgeX + 6} ${it.idealY}, ${labelX - 5} ${it.y}, ${labelX - 1} ${it.y}` }));
      }
      const text = svgEl('text', { class: 'sg-series-label', x: labelX, y: it.y,
        'dominant-baseline': 'middle', 'text-anchor': 'start',
        'font-size': config.labelSize, 'font-weight': 700, fill: it.color });
      text.textContent = it.label;
      it.g.appendChild(text);
    });
  }

  // Remove axis labels that still collide with the previous kept label on the
  // same row. Measured per row (0 = top, 1 = staggered lower) so stagger only
  // has to clear every-other label. `gap` is the min px between boxes.
  function pruneOverlappingLabels(items, gap) {
    const lastRight = {};
    items.forEach((it) => {
      let bb;
      try { bb = it.el.getBBox(); } catch (e) { return; }
      const left = bb.x, right = bb.x + bb.width;
      const prev = lastRight[it.row];
      if (prev != null && left < prev + gap) {
        it.el.remove();
        if (it.lineEl) it.lineEl.remove();
      } else {
        lastRight[it.row] = right;
      }
    });
  }

  // Build a closed smooth band: top boundary L→R, then bottom boundary R→L.
  function bandPath(topPts, botPts) {
    const top = curvePath(topPts);
    const botRev = botPts.slice().reverse();
    const bot = curvePath(botRev);
    const botL = 'L' + bot.slice(1); // turn the bottom's leading 'M' into a join
    return top + ' ' + botL + ' Z';
  }
  // Path string (starting with M) through points, per the curve mode.
  function curvePath(pts) {
    const n = pts.length;
    if (!n) return '';
    if (n === 1) return `M ${pts[0][0]} ${pts[0][1]}`;
    if (config.curve === 'linear') {
      return 'M ' + pts.map((p) => `${p[0]} ${p[1]}`).join(' L ');
    }
    if (config.curve === 'step') {
      let d = `M ${pts[0][0]} ${pts[0][1]}`;
      for (let i = 1; i < n; i++) d += ` L ${pts[i][0]} ${pts[i - 1][1]} L ${pts[i][0]} ${pts[i][1]}`;
      return d;
    }
    // smooth: Catmull-Rom → cubic Bézier (passes through every point).
    const k = config.curveTension;
    let d = `M ${pts[0][0]} ${pts[0][1]}`;
    for (let i = 0; i < n - 1; i++) {
      const p0 = pts[i - 1] || pts[i], p1 = pts[i], p2 = pts[i + 1], p3 = pts[i + 2] || p2;
      const c1x = p1[0] + (p2[0] - p0[0]) / 6 * k, c1y = p1[1] + (p2[1] - p0[1]) / 6 * k;
      const c2x = p2[0] - (p3[0] - p1[0]) / 6 * k, c2y = p2[1] - (p3[1] - p1[1]) / 6 * k;
      d += ` C ${c1x} ${c1y}, ${c2x} ${c2y}, ${p2[0]} ${p2[1]}`;
    }
    return d;
  }

  /* ---------- color helpers ----------
     A series' color is resolved by IDENTITY, not by its index in the current
     (possibly filtered) data:
       1. an explicit per-series override (config.seriesColors[key]) always wins;
       2. else by colorMode — single / sequential(by total) / palette;
       3. palette colors are indexed by the series' position in the persisted
          first-seen order (config.seriesOrder), so filtering to a subset keeps
          each category on the same swatch.
  */
  // Record any newly-seen series keys in the persisted first-seen order so their
  // palette slot is stable across filters/reloads. Saves only when it grows.
  function ensureSeriesOrder(model) {
    const order = config.seriesOrder || (config.seriesOrder = []);
    let changed = false;
    model.seriesList.forEach((s) => { if (order.indexOf(s.raw) < 0) { order.push(s.raw); changed = true; } });
    if (changed) saveConfig();
  }
  function stableIndex(rawKey) {
    const order = config.seriesOrder || [];
    const i = order.indexOf(rawKey);
    return i < 0 ? 0 : i;
  }
  // The categorical palette mode (NOT the mono ramp) is the only place identity-
  // stable slots + per-series overrides apply.
  function isCategoricalPalette() {
    return config.colorMode === 'palette' && config.palette !== 'mono';
  }
  // Per-series overrides + identity-stable palette slots apply ONLY in palette
  // mode. In sequential mode color encodes the series total on a continuous
  // low→high scale (an override would punch a hole in it); in single mode every
  // stream is one color. Both ignore seriesColors/seriesOrder.
  function seriesColorFor(model, si, totals, lo, hi) {
    const s = model.seriesList[si];
    if (config.colorMode === 'single') return config.singleColor;
    if (config.colorMode === 'sequential') {
      const t = (hi <= lo) ? 0.5 : (totals[si] - lo) / (hi - lo);
      return mix(config.seqLowColor, config.seqHighColor, Math.max(0, Math.min(1, t)));
    }
    // palette mode: an explicit per-series override always wins.
    const ov = config.seriesColors && config.seriesColors[s.raw];
    if (ov) return ov;
    const oi = stableIndex(s.raw);
    if (config.palette === 'mono') {
      const denom = Math.max(1, (config.seriesOrder || []).length - 1);
      return mix(config.monoBaseColor, '#ffffff', (oi / denom) * 0.7);   // rank ramp
    }
    const pal = PALETTES[config.palette] || PALETTES.cool;
    return pal[oi % pal.length];
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
  // Coerce any CSS color we produce (hex or "rgb(r, g, b)") to "#rrggbb" so an
  // <input type=color> can show it.
  function anyToHex(c) {
    if (!c) return '#888888';
    const s = String(c);
    if (s[0] === '#') return s.length === 4 ? '#' + s.slice(1).split('').map((x) => x + x).join('') : s;
    const m = s.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
    if (m) return '#' + [1, 2, 3].map((i) => (+m[i]).toString(16).padStart(2, '0')).join('');
    return '#888888';
  }

  /* =========================================================
     INTERACTIVITY — native Tableau tooltip + per-cell selection
     We resolve the cursor to the nearest axis column INSIDE the
     hovered band, so each interaction targets a single (series,
     axis) mark — e.g. one fruit-month. That single tuple is then:
       • hovered via hoverTupleAsync → Tableau's NATIVE tooltip
         (a single mark IS describable, unlike a whole stream), and
       • selected via selectTuplesAsync → so dashboard filtering is
         at fruit-month granularity, not the full series.
     ========================================================= */
  const SELECT_SIMPLE = (tableau.SelectOptions && tableau.SelectOptions.Simple) || 'select-options-simple';
  let selectedTuples = new Set();
  let hoverTupleId = 0, hoverAt = 0;

  function wireInteractivity() {
    host.addEventListener('mousemove', (ev) => {
      if (!config.enableTooltips) return;
      const t = ev.target.closest && ev.target.closest('.sg-interactive');
      const tid = t ? cellTupleAt(t, ev) : 0;
      if (!tid) { clearHover(); return; }
      const now = (window.performance && performance.now()) || Date.now();
      // refire when the cell changes, and periodically so the native tooltip
      // re-anchors as the cursor slides along a stream.
      if (tid !== hoverTupleId || now - hoverAt > 60) {
        hoverTuple(tid, ev);
        hoverTupleId = tid; hoverAt = now;
      }
    });
    host.addEventListener('mouseleave', clearHover);
    host.addEventListener('click', (ev) => {
      if (!config.enableTooltips) return;
      const t = ev.target.closest && ev.target.closest('.sg-interactive');
      const tid = t ? cellTupleAt(t, ev) : 0;
      toggleSelection(tid ? [tid] : null, ev.ctrlKey || ev.metaKey);
    });
  }

  // Tuple id of the single (series, axis) cell under the cursor within a band.
  function cellTupleAt(group, ev) {
    if (!group._tupleAt) return 0;
    const ai = nearestAxis(ev);
    return ai >= 0 ? (group._tupleAt[ai] || 0) : 0;
  }
  // Nearest axis column index to the cursor (svg coords == client offset since
  // the SVG is drawn 1:1 in pixels with no viewBox).
  function nearestAxis(ev) {
    if (!geom.svg || !geom.xs.length) return -1;
    const rect = geom.svg.getBoundingClientRect();
    const localX = ev.clientX - rect.left;
    let best = 0, bestD = Infinity;
    for (let i = 0; i < geom.xs.length; i++) {
      const dd = Math.abs(geom.xs[i] - localX);
      if (dd < bestD) { bestD = dd; best = i; }
    }
    return best;
  }

  function getWorksheet() {
    return tableau.extensions.worksheetContent && tableau.extensions.worksheetContent.worksheet;
  }
  // Show Tableau's native tooltip for one mark; pass 0 to clear it.
  function hoverTuple(tupleId, ev) {
    const ws = getWorksheet();
    if (!ws || !ws.hoverTupleAsync) return;
    const tip = ev ? { tooltipAnchorPoint: { x: ev.clientX, y: ev.clientY } } : undefined;
    try { ws.hoverTupleAsync(tupleId, tip).catch(() => {}); } catch (e) { /* ignore */ }
  }
  function clearHover() {
    if (hoverTupleId) { hoverTuple(0); hoverTupleId = 0; }
  }

  function pushSelection() {
    const ws = getWorksheet();
    if (ws && ws.selectTuplesAsync) {
      try { ws.selectTuplesAsync([...selectedTuples], SELECT_SIMPLE).catch(() => {}); } catch (e) { /* ignore */ }
    }
    applySelectionStyles();
  }
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
  // Dim a band when a selection is active and NONE of its cells are selected
  // (selection is now per cell, so "all selected" would almost never hold).
  function applySelectionStyles() {
    const active = selectedTuples.size > 0;
    host.querySelectorAll('.sg-interactive').forEach((node) => {
      if (!active) { node.classList.remove('sg-dimmed', 'sg-selected'); return; }
      const sel = (node._tupleAt || []).some((id) => id && selectedTuples.has(id));
      node.classList.toggle('sg-selected', sel);
      node.classList.toggle('sg-dimmed', !sel);
    });
    // keep the legend in sync: a series is "selected" if any of its tuples are.
    host.querySelectorAll('.sg-legend-item').forEach((node) => {
      if (!active || !node._tupleIds) { node.classList.remove('sg-dimmed', 'sg-selected'); return; }
      const sel = node._tupleIds.some((id) => selectedTuples.has(id));
      node.classList.toggle('sg-selected', sel);
      node.classList.toggle('sg-dimmed', !sel);
    });
  }
  // Re-sync the local selection set from the worksheet (filters shift tuple ids;
  // other sheets can drive the selection). Map each selected mark back to its
  // exact (series, axis) cell via the index maps + tupleAt — collision-free,
  // unlike matching by series alone. If a selection carries only the series
  // dimension (no axis), light up that whole stream as a fallback.
  async function syncSelectionFromWorksheet(model) {
    const ws = getWorksheet();
    if (!ws || !ws.getSelectedMarksAsync || !model) return;
    try {
      const result = await ws.getSelectedMarksAsync();
      const tables = (result && result.data) || [];
      const next = new Set();
      tables.forEach((tbl) => {
        const cols = tbl.columns || [];
        const sCol = cols.findIndex((c) => c.fieldName === model.seriesField);
        const aCol = cols.findIndex((c) => c.fieldName === model.axisField);
        if (sCol < 0) return;
        (tbl.data || []).forEach((row) => {
          const si = model.seriesIndex.get(String(row[sCol].value));
          if (si == null) return;
          if (aCol >= 0) {
            const ai = model.axisIndex.get(String(row[aCol].value));
            if (ai != null) { const id = model.tupleAt[si][ai]; if (id) next.add(id); }
          } else {
            model.seriesList[si].tupleIds.forEach((id) => next.add(id));
          }
        });
      });
      selectedTuples = next;
    } catch (e) { /* leave selection as-is */ }
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
    document.documentElement.style.setProperty('--sg-font', FONT_STACKS[config.fontFamily] || FONT_STACKS.system);
    document.documentElement.style.setProperty('--sg-ink', config.textColor);
    document.documentElement.style.setProperty('--sg-muted', config.mutedColor);
    document.documentElement.style.setProperty('--sg-bg', config.bgColor);
  }

  /* =========================================================
     CONFIG MODAL (declarative schema)
     ========================================================= */
  const SCHEMA = [
    { section: 'General', fields: [
      { key: 'title', label: 'Chart title', type: 'text' },
      { key: 'subtitle', label: 'Subtitle', type: 'text' },
      { key: 'caption', label: 'Caption (footer)', type: 'text' },
      { key: 'bgColor', label: 'Background', type: 'color' },
      { key: 'textColor', label: 'Text color', type: 'color' },
      { key: 'mutedColor', label: 'Muted color', type: 'color' },
      { key: 'fontFamily', label: 'Font', type: 'select',
        options: [['system', 'System'], ['serif', 'Serif'], ['mono', 'Monospace']] },
      { key: 'enableTooltips', label: 'Tooltips + selection', type: 'checkbox' },
    ]},
    { section: 'Stream shape', fields: [
      { key: 'offset', label: 'Baseline', type: 'select',
        options: [['wiggle', 'Wiggle (streamgraph)'], ['silhouette', 'Silhouette (centered)'], ['zero', 'Zero (stacked)'], ['expand', 'Expand (100%)']] },
      { key: 'order', label: 'Stacking order', type: 'select',
        options: [['inside-out', 'Inside-out'], ['ascending', 'Total ↑'], ['descending', 'Total ↓'], ['reverse', 'Reverse'], ['none', 'Data order']] },
      { key: 'curve', label: 'Curve', type: 'select',
        options: [['smooth', 'Smooth'], ['linear', 'Linear'], ['step', 'Step']] },
      { key: 'curveTension', label: 'Curviness (smooth)', type: 'range', min: 0, max: 1.2, step: 0.05 },
    ]},
    { section: 'Bands', fields: [
      { key: 'layerOpacity', label: 'Band opacity', type: 'range', min: 0.2, max: 1, step: 0.05 },
      { key: 'showStroke', label: 'Separator stroke', type: 'checkbox' },
      { key: 'strokeColor', label: 'Stroke color', type: 'color' },
      { key: 'strokeWidth', label: 'Stroke width (px)', type: 'range', min: 0, max: 4, step: 0.5 },
    ]},
    { section: 'Color', fields: [
      { key: 'colorMode', label: 'Color mode', type: 'select',
        options: [['palette', 'Palette (by series)'], ['sequential', 'Sequential (by total)'], ['single', 'Single color']] },
      { key: 'palette', label: 'Palette', type: 'select',
        options: [['cool', 'Cool'], ['tableau10', 'Tableau 10'], ['bold', 'Bold'], ['pastel', 'Pastel'], ['warm', 'Warm'], ['mono', 'Mono (single hue)']] },
      { key: 'monoBaseColor', label: 'Mono base color', type: 'color' },
      { key: 'singleColor', label: 'Single color', type: 'color' },
      { key: 'seqLowColor', label: 'Sequential — low', type: 'color' },
      { key: 'seqHighColor', label: 'Sequential — high', type: 'color' },
    ]},
    { section: 'Axis & grid', fields: [
      { key: 'showXAxis', label: 'Show x-axis labels', type: 'checkbox' },
      { key: 'staggerLabels', label: 'Offset labels (stagger rows)', type: 'checkbox' },
      { key: 'minLabelSpacing', label: 'Min label spacing (px)', type: 'range', min: 20, max: 140, step: 4 },
      { key: 'showGridlines', label: 'Vertical gridlines', type: 'checkbox' },
      { key: 'gridColor', label: 'Gridline color', type: 'color' },
      { key: 'showZeroLine', label: 'Baseline (zero) line', type: 'checkbox' },
    ]},
    { section: 'Legend', fields: [
      { key: 'showLegend', label: 'Show legend', type: 'checkbox' },
      { key: 'legendPosition', label: 'Position', type: 'select',
        options: [['top', 'Top'], ['bottom', 'Bottom']] },
    ]},
    { section: 'Series labels', fields: [
      { key: 'showSeriesLabels', label: 'Show series labels (right edge)', type: 'checkbox' },
      { key: 'labelSize', label: 'Label size (px)', type: 'range', min: 8, max: 22, step: 1 },
      { key: 'labelWidth', label: 'Label gutter (px)', type: 'range', min: 40, max: 200, step: 4 },
      { key: 'labelMinThickness', label: 'Min band thickness to label (px)', type: 'range', min: 0, max: 40, step: 1 },
    ]},
  ];

  function openConfigModal() {
    const overlay = document.getElementById('cfg-overlay');
    if (!overlay) return;
    buildModalBody();   // rebuild so the dynamic "Series colors" list reflects
    overlay.hidden = false;   // whatever series are currently in the view
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
  // Dynamic section: one color swatch per series currently in the view, plus a
  // reset-to-palette button. Built from lastModel (the schema is static, the
  // series are data-driven), so it only appears once fields are dropped.
  function buildSeriesColorsSection() {
    // Per-series overrides only apply in palette mode — hide the section
    // (would do nothing) in sequential/single.
    if (config.colorMode !== 'palette') return null;
    if (!lastModel || !lastModel.seriesList.length) return null;
    const totals = lastModel.seriesList.map((s) => s.total);
    const loT = Math.min(...totals), hiT = Math.max(...totals);
    const wrap = el('div', 'cfg-section');
    wrap.appendChild(el('h3', 'cfg-section-title', 'Series colors'));
    lastModel.seriesList.forEach((s, si) => {
      const row = el('label', 'cfg-field');
      row.appendChild(el('span', 'cfg-label', s.label));
      const box = el('span', 'cfg-color-box');
      const inp = document.createElement('input');
      inp.type = 'color';
      inp.value = anyToHex(seriesColorFor(lastModel, si, totals, loT, hiT));
      inp.addEventListener('input', () => {
        config.seriesColors[s.raw] = inp.value;
        saveConfig();
        if (lastModel) render(lastModel);
      });
      const reset = el('button', 'cfg-mini', '↺');
      reset.type = 'button';
      reset.title = 'Reset to palette';
      reset.addEventListener('click', (e) => {
        e.preventDefault();
        delete config.seriesColors[s.raw];
        saveConfig();
        inp.value = anyToHex(seriesColorFor(lastModel, si, totals, loT, hiT));
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
    } else {
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

  /* ---------- config persistence (workbook settings only) ---------- */
  function loadConfig() {
    let saved = {};
    try {
      const store = tableau.extensions.settings;
      const raw = store && store.get(STORAGE_KEY);
      if (raw) saved = JSON.parse(raw);
    } catch (e) { console.error('reading workbook settings failed:', e); }
    const cfg = { ...DEFAULT_CONFIG, ...saved };
    // Deep-copy the mutable collections so we never mutate DEFAULT_CONFIG's
    // shared objects (which would leak across instances and break Reset).
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

  /* ---------- tiny DOM/SVG helpers ---------- */
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
    host.innerHTML = '<div class="sg-empty"><p>' + html + '</p></div>';
  }
  function showEmptyState(fields) {
    applyTheme();
    fields = fields || {};
    const chip = (id, label, kind) => {
      const got = fields[id] && fields[id].length;
      return '<span class="sg-chip ' + kind + (got ? ' filled' : '') + '">' +
        label + (got ? ': ' + escapeHtml(fields[id].join(', ')) : '') + '</span>';
    };
    host.innerHTML =
      '<div class="sg-empty"><div class="sg-empty-card">' +
        '<div class="sg-empty-head">' +
          '<div class="sg-empty-thumb">' + buildThumb() + '</div>' +
          '<div class="sg-empty-meta">' +
            '<div class="sg-empty-title">Stream Graph</div>' +
            '<div class="sg-empty-desc">A streamgraph: stacked areas flowing around a wiggling baseline, so each ' +
            'series reads as a stream that swells and shrinks over time. Smooth, colored, and labelled — fully ' +
            'customizable from the &#9881; gear or the <b>Format Extension</b> button.</div>' +
          '</div>' +
        '</div>' +
        '<div class="sg-empty-guide">' +
          '<div class="sg-empty-guide-title">Get started</div>' +
          '<div class="sg-empty-guide-row">Drag an ordered <b>dimension</b> onto ' +
            chip('axis', 'Axis', 'dim') + ', a <b>dimension</b> onto ' +
            chip('series', 'Series', 'dim') + ', and a <b>measure</b> onto ' +
            chip('value', 'Value', 'measure') + '.' +
          '</div>' +
        '</div>' +
      '</div></div>';
  }
  // Inline SVG thumbnail: a tiny three-layer stream.
  function buildThumb() {
    const cols = ['#3d7bf4', '#c13bd6', '#b9bcc4'];
    const layers = [
      'M0 50 C 18 46, 30 40, 48 44 S 78 54, 96 50 L 96 64 C 78 60, 60 70, 48 66 S 18 58, 0 62 Z',
      'M0 38 C 18 30, 34 22, 48 30 S 80 42, 96 36 L 96 50 C 80 56, 66 44, 48 44 S 18 46, 0 50 Z',
      'M0 26 C 20 18, 32 10, 48 16 S 78 28, 96 22 L 96 36 C 78 42, 64 30, 48 30 S 18 30, 0 38 Z',
    ];
    let svg = '';
    layers.forEach((d, i) => { svg += `<path d="${d}" fill="${cols[i]}" fill-opacity="0.92"/>`; });
    return `<svg viewBox="0 0 96 84" width="84" height="74" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Stream graph preview">${svg}</svg>`;
  }
  function escapeHtml(s) {
    return String(s).replace(/[&<>"]/g, (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }
})();

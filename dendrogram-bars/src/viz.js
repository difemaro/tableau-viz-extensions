'use strict';

/* =============================================================
   DENDROGRAM BARS — Tableau Viz Extension
   -------------------------------------------------------------
   A "fan-out" bar chart: every category bar is tied back to a
   single root node (the grand total) by a smooth bundled curve,
   like a one-level dendrogram. Bars are sorted, colored from a
   categorical palette, and labelled with name + value.

   The four moves (see CLAUDE.md):
     1. CONNECT   — initializeAsync(), grab the worksheet
     2. ENCODINGS — Category (dimension) + Value (measure)
     3. DATA      — read summary data → one item per category
     4. DRAW      — render(model) into #viz as an SVG; redraw on change
   ============================================================= */

(function () {
  const host = document.getElementById('viz');
  const STORAGE_KEY = 'dendrobars-config-v1';
  const SVGNS = 'http://www.w3.org/2000/svg';

  /* ---------- config ---------- */
  const DEFAULT_CONFIG = {
    // identity / theme (light by default)
    title: '',
    bgColor: '#ffffff',
    fontFamily: 'system',
    textColor: '#1c2330',
    mutedColor: '#6b7280',

    // root (grand-total) node
    showRoot: true,
    rootLabel: 'Total',
    rootCaptionSize: 12,
    rootValueSize: 26,
    totalPrefix: '$',
    totalDecimals: 2,        // total is abbreviated (auto unit), e.g. $2.30M

    // bars
    sort: 'desc',            // desc | asc | none
    fitHeight: true,         // fit all rows to the pane height (no scroll);
                             //   off → use rowHeight px/row and scroll on overflow
    rowHeight: 28,           // px per category (used only when fitHeight is off)
    barThickness: 0.55,      // bar height as a fraction of its row band
    barRadius: 20,           // corner radius (high = pill); clamped to barH/2
    barMinWidth: 4,          // keep tiny values visible
    barOpacity: 1,

    // fan curves
    showCurves: true,
    curveTension: 0.55,      // 0 = straight, 1 = very S-shaped
    curveWidth: 1.5,
    curveOpacity: 0.55,
    curveColorMode: 'flat',  // match (bar color) | flat
    curveFlatColor: '#000000',

    // color
    colorMode: 'divergent',  // palette (by category) | single | divergent (by value)
    singleColor: '#5a8fc7',  // used when colorMode = single
    divLowColor: '#d6dee8',  // low-value color  (light grey-blue)
    divHighColor: '#3b6ea5', // high-value color (blue)
    palette: 'pastel',       // pastel | tableau10 | bold | mono  (colorMode = palette)
    monoBaseColor: '#5a8fc7',

    // labels (right column)
    showCategoryNames: true,
    showValueLabels: true,
    labelSize: 13,
    labelWidth: 170,         // reserved width for the right-hand label column

    // value-number format (for the right-hand labels)
    numPrefix: '',
    numSuffix: '',
    numDecimals: 0,
    numUnit: 'none',         // none | auto | K | M | B
    numThousands: true,

    // optional value axis along the bottom
    showAxis: false,
    axisTicks: 6,

    // interactivity — native Tableau tooltips on hover + selection on click
    enableTooltips: true,
  };

  const FONT_STACKS = {
    system: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
    serif: 'Georgia, "Times New Roman", serif',
    mono: 'ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace',
  };

  const PALETTES = {
    pastel:    ['#bdbdbd','#a5d6a7','#ef9a9a','#ce93d8','#ffcc80','#ef9a9a','#90caf9','#c5b358',
                '#81c784','#a8d0f0','#80cbc4','#9e9e9e','#f8bbd0','#ffb74d','#fff176','#e57373','#4db6ac'],
    tableau10: ['#4e79a7','#f28e2b','#e15759','#76b7b2','#59a14f','#edc948','#b07aa1','#ff9da7','#9c755f','#bab0ac'],
    bold:      ['#1f77b4','#ff7f0e','#2ca02c','#d62728','#9467bd','#8c564b','#e377c2','#7f7f7f','#bcbd22','#17becf'],
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

      // Re-render (sized to the new host width) on any size change. We listen
      // BOTH ways and coalesce through rAF: ResizeObserver catches pane/object
      // resizes, window 'resize' is a belt-and-braces fallback for hosts where
      // RO under-fires. Without a re-render the SVG would keep its old pixel
      // width and only visually scale, distorting the chart.
      if (window.ResizeObserver) {
        new ResizeObserver(scheduleRender).observe(host);
      }
      window.addEventListener('resize', scheduleRender);
      // Backstop: poll the host size. Tableau often sizes the extension iframe
      // to its final dimensions AFTER first paint, and in some hosts neither
      // ResizeObserver nor window 'resize' fires for that — leaving the chart
      // stuck at its initial (narrow) width. Polling guarantees we notice the
      // change and re-render to fill the real width. Cheap (no work unless the
      // measured size actually changed).
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

    if (!fields.category || !fields.category.length ||
        !fields.value || !fields.value.length) {
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

  function buildModel(worksheet, fields, table) {
    const cols = table.columns;
    const catField = fields.category[0];
    const valField = fields.value[0];
    const catIdx = findColumn(cols, catField);
    const valIdx = findValueColumn(cols, valField);

    const items = [];
    let total = 0;
    table.data.forEach((row, idx) => {
      const label = cellText(row[catIdx]);
      const raw = row[valIdx] ? Number(row[valIdx].value) : NaN;
      const value = Number.isFinite(raw) ? raw : null;
      const rawKey = row[catIdx] ? String(row[catIdx].value) : '';
      if (value != null) total += value;
      // tuple ids are 1-based summary-data rows
      items.push({ label, value, tupleId: idx + 1, rawKey });
    });

    return { worksheet, catField, valField, items, total };
  }

  function sortedItems(model) {
    const items = model.items.slice();
    if (config.sort === 'desc') items.sort((a, b) => (b.value || -Infinity) - (a.value || -Infinity));
    else if (config.sort === 'asc') items.sort((a, b) => (a.value || Infinity) - (b.value || Infinity));
    return items;
  }

  // Coalesce resize-driven re-renders into a single rAF so a burst of resize
  // events (drag-resizing a pane) produces one redraw per frame, each sized to
  // the current host width.
  let renderRAF = 0;
  function scheduleRender() {
    if (!lastModel || renderRAF) return;
    renderRAF = requestAnimationFrame(() => { renderRAF = 0; render(lastModel); });
  }

  // Poll the host size as a backstop for resize events that don't fire (see
  // init). Re-renders only when the measured width/height actually changes.
  // Renders DIRECTLY (not via the rAF-coalesced scheduleRender) so the backstop
  // works even where rAF callbacks are starved; the 250ms interval is its own
  // throttle.
  let watchedW = 0, watchedH = 0;
  function startSizeWatch() {
    setInterval(() => {
      // Track the viewport too — in hosts where host.clientWidth is stuck, the
      // viewport width is what actually changes when the pane is resized.
      const w = host.clientWidth + 0.001 * (window.innerWidth || 0);
      const h = host.clientHeight;
      if (w === watchedW && h === watchedH) return;
      watchedW = w; watchedH = h;
      if (lastModel) render(lastModel);
    }, 250);
  }

  // Decide the drawing width. Prefer the laid-out wrapper's content width; if
  // that's clearly under the viewport (container hasn't stretched / host
  // reported a bogus small width), fall back to the viewport minus our padding.
  const HOST_PAD_X = 32;   // #viz horizontal padding (16 left + 16 right)
  function measureWidth(wrap) {
    const vpW = (document.documentElement && document.documentElement.clientWidth) ||
                window.innerWidth || 0;
    const wrapW = wrap.clientWidth || 0;
    const vpAvail = Math.max(0, vpW - HOST_PAD_X);
    // Use the wrapper width unless it's well under the viewport (≈ not stretched).
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

    const wrap = el('div', 'db');
    let titleEl = null;
    if (config.title) { titleEl = el('div', 'db-title', config.title); wrap.appendChild(titleEl); }
    // Append the wrapper NOW (before building the SVG) so we can measure the
    // actually laid-out content box. host.clientWidth/Height have proven
    // unreliable in some Tableau hosts; the stretched wrapper is accurate.
    host.appendChild(wrap);

    // --- geometry --------------------------------------------------
    const W = measureWidth(wrap);
    // Available height for the chart = wrapper content box minus the title.
    const availH = wrap.clientHeight || ((host.clientHeight || 500) - 28);
    const containerH = Math.max(40, availH - (titleEl ? titleEl.offsetHeight : 0));
    const pad = {
      top: 16,
      right: 14,
      bottom: config.showAxis ? 34 : 14,
      left: config.showRoot ? 130 : 24,
    };

    // Height behaviour:
    //  • fitHeight (default): the SVG is exactly the pane height, so every row
    //    is visible and the rows compress to fit — no scrollbar.
    //  • off: rows keep `rowHeight` px; when that's taller than the pane the
    //    SVG grows and #viz scrolls vertically.
    const contentH = n * config.rowHeight + pad.top + pad.bottom;
    const svgH = config.fitHeight ? containerH : Math.max(containerH, contentH);
    const plotTop = pad.top, plotBottom = svgH - pad.bottom;
    const plotH = plotBottom - plotTop;
    const band = plotH / n;

    const rootX = pad.left;
    const barStartX = rootX + Math.min(170, (W - rootX) * 0.28);
    const labelW = (config.showCategoryNames || config.showValueLabels) ? config.labelWidth : 0;
    const maxBarW = Math.max(10, W - barStartX - labelW - pad.right);
    const rootY = plotTop + plotH / 2;

    const maxVal = Math.max(1, ...items.map((it) => it.value || 0));
    const scaleW = (v) => (v == null ? 0 : Math.max(config.barMinWidth, (v / maxVal) * maxBarW));
    // True value range (no floor) for the divergent color scale.
    const vals = items.map((it) => it.value).filter((v) => v != null);
    const loVal = vals.length ? Math.min(...vals) : 0;
    const hiVal = vals.length ? Math.max(...vals) : 1;

    // NO viewBox / preserveAspectRatio: we draw in pixels (1 user unit = 1px),
    // so the chart is never scaled. A viewBox + "meet" would uniformly scale
    // the whole drawing DOWN (shrinking width too, leaving empty space on the
    // right) whenever the flex container is shorter than svgH — which is what
    // made the chart look like it wasn't filling the width in a short pane.
    const svg = svgEl('svg', { class: 'db-svg', width: W, height: svgH });

    // --- optional bottom axis -------------------------------------
    if (config.showAxis) {
      const axis = svgEl('g', { class: 'db-axis' });
      const y = plotBottom + 6;
      axis.appendChild(svgEl('line', { x1: barStartX, y1: y, x2: barStartX + maxBarW, y2: y,
        stroke: config.mutedColor, 'stroke-width': 1, 'stroke-opacity': 0.5 }));
      const ticks = Math.max(2, config.axisTicks | 0);
      for (let t = 0; t <= ticks; t++) {
        const frac = t / ticks;
        const x = barStartX + frac * maxBarW;
        axis.appendChild(svgEl('line', { x1: x, y1: y, x2: x, y2: y + 4,
          stroke: config.mutedColor, 'stroke-width': 1, 'stroke-opacity': 0.6 }));
        const lab = svgEl('text', { x, y: y + 16, 'text-anchor': 'middle',
          fill: config.mutedColor, 'font-size': 10 });
        lab.textContent = formatNumber(frac * maxVal);
        axis.appendChild(lab);
      }
      svg.appendChild(axis);
    }

    // --- one group per category (curve + bar + label) -------------
    items.forEach((it, i) => {
      const y = plotTop + band * (i + 0.5);
      const barH = Math.max(2, Math.min(band - 2, band * config.barThickness));
      const color = colorFor(i, n, it.value, loVal, hiVal);

      const g = svgEl('g', { class: 'db-row db-interactive' });
      g._tupleIds = [it.tupleId];
      g._tip = { name: it.label, field: model.catField, valueField: model.valField, value: it.value };

      // fan curve from the root to the bar's left end
      if (config.showCurves) {
        const dx = (barStartX - rootX) * config.curveTension;
        const d = `M ${rootX} ${rootY} C ${rootX + dx} ${rootY}, ${barStartX - dx} ${y}, ${barStartX} ${y}`;
        const path = svgEl('path', { class: 'db-curve', d, fill: 'none',
          stroke: config.curveColorMode === 'flat' ? config.curveFlatColor : color,
          'stroke-width': config.curveWidth, 'stroke-opacity': config.curveOpacity,
          'stroke-linecap': 'round' });
        g.appendChild(path);
      }

      // the bar (rounded pill)
      const w = scaleW(it.value);
      const rect = svgEl('rect', { class: 'db-bar', x: barStartX, y: y - barH / 2,
        width: w, height: barH, rx: Math.min(config.barRadius, barH / 2),
        fill: color, 'fill-opacity': config.barOpacity });
      g.appendChild(rect);

      // right-hand label: "Name  value" (right-aligned in the label column)
      if (config.showCategoryNames || config.showValueLabels) {
        const tx = W - pad.right;
        const text = svgEl('text', { class: 'db-label', x: tx, y, 'text-anchor': 'end',
          'dominant-baseline': 'middle', 'font-size': config.labelSize });
        // Text colors mirror the root node: the value (number) is the prominent
        // text (textColor) — same as the total value — and the category name is
        // the muted text — same as the total's caption.
        if (config.showCategoryNames) {
          const nameT = svgEl('tspan', { fill: config.mutedColor });
          nameT.textContent = it.label;
          text.appendChild(nameT);
        }
        if (config.showValueLabels && it.value != null) {
          const valT = svgEl('tspan', { fill: config.textColor, dx: config.showCategoryNames ? 8 : 0 });
          valT.textContent = formatNumber(it.value);
          text.appendChild(valT);
        }
        g.appendChild(text);
      }

      svg.appendChild(g);
    });

    // --- root (grand-total) node ----------------------------------
    if (config.showRoot) {
      const rg = svgEl('g', { class: 'db-root' });
      rg.appendChild(svgEl('circle', { cx: rootX, cy: rootY, r: 3.5, fill: config.textColor }));
      const lx = rootX - 14;
      if (config.rootLabel) {
        const cap = svgEl('text', { x: lx, y: rootY - 14, 'text-anchor': 'end',
          fill: config.mutedColor, 'font-size': config.rootCaptionSize,
          'letter-spacing': '0.06em' });
        cap.textContent = config.rootLabel.toUpperCase();
        rg.appendChild(cap);
      }
      const tot = svgEl('text', { x: lx, y: rootY + config.rootValueSize * 0.55,
        'text-anchor': 'end', fill: config.textColor, 'font-size': config.rootValueSize,
        'font-weight': 700 });
      tot.textContent = formatTotal(model.total);
      rg.appendChild(tot);
      svg.appendChild(rg);
    }

    wrap.appendChild(svg);

    applySelectionStyles();
  }
  /* ▲▲▲  END BUILD ZONE  ▲▲▲ */

  /* ---------- color helpers ----------
     A bar's color depends on the chosen colorMode:
       • single    — one fixed color for every bar
       • divergent — interpolate low→high color by the bar's value
       • palette   — categorical (pastel/tableau10/bold), or a single-hue ramp
                     for 'mono'. This is the original behaviour.
     `value`, `loVal`, `hiVal` are only needed for the divergent mode. */
  function colorFor(i, n, value, loVal, hiVal) {
    if (config.colorMode === 'single') return config.singleColor;
    if (config.colorMode === 'divergent') {
      const t = (value == null || hiVal <= loVal) ? 0.5 : (value - loVal) / (hiVal - loVal);
      return mix(config.divLowColor, config.divHighColor, Math.max(0, Math.min(1, t)));
    }
    if (config.palette === 'mono') {
      const t = n <= 1 ? 0 : (i / (n - 1)) * 0.7;   // base → lighter
      return mix(config.monoBaseColor, '#ffffff', t);
    }
    const pal = PALETTES[config.palette] || PALETTES.pastel;
    return pal[i % pal.length];
  }
  function mix(hexA, hexB, t) {
    const a = hexToRgb(hexA), b = hexToRgb(hexB);
    const c = a.map((ch, i) => Math.round(ch + (b[i] - ch) * t));
    return `rgb(${c[0]}, ${c[1]}, ${c[2]})`;
  }
  function hexToRgb(hex) {
    const h = String(hex).replace('#', '');
    const f = h.length === 3 ? h.split('').map((x) => x + x).join('') : h;
    const n = parseInt(f, 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }

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
  // The root total is always shown abbreviated (e.g. $2.30M).
  function formatTotal(v) {
    if (v == null || !Number.isFinite(v)) return '';
    const [n, unit] = abbreviate(v);
    return config.totalPrefix + withThousands(n.toFixed(config.totalDecimals)) + unit;
  }

  /* =========================================================
     INTERACTIVITY — native tooltips + mark selection
     ========================================================= */
  const SELECT_SIMPLE = (tableau.SelectOptions && tableau.SelectOptions.Simple) || 'select-options-simple';

  let selectedTuples = new Set();
  let hoverEl = null, hoverAt = 0;

  // Marquee (rubber-band) drag-select state. A press-and-drag in the viz draws a
  // rectangle and selects every bar it covers; a press that doesn't move past
  // the threshold stays a click. (See CLAUDE.md "Marquee (drag-to-select)".)
  const MARQUEE_THRESHOLD = 4;   // px of movement before a press becomes a drag
  let dragStart = null, dragCur = null;
  let marqueeActive = false, marqueeAdditive = false, suppressClick = false;

  function wireInteractivity() {
    host.addEventListener('mousemove', (ev) => {
      if (!config.enableTooltips) return;
      if (marqueeActive) return;   // dragging a marquee → no hover tooltips
      const t = ev.target.closest && ev.target.closest('.db-interactive');
      if (!t) { clearHoverState(); return; }
      const now = (window.performance && performance.now()) || Date.now();
      if (t !== hoverEl || now - hoverAt > 50) {
        hoverEl = t; hoverAt = now;
        hoverTuple(t._tupleIds[0], ev);
      }
    });
    host.addEventListener('mouseleave', clearHoverState);
    host.addEventListener('click', (ev) => {
      if (!config.enableTooltips) return;
      if (suppressClick) { suppressClick = false; return; }  // came from a drag
      const t = ev.target.closest && ev.target.closest('.db-interactive');
      toggleSelection(t && t._tupleIds ? t._tupleIds : null, ev.ctrlKey || ev.metaKey);
    });

    // Press-and-drag → marquee select. Track move/up on document so the drag
    // keeps working when the pointer leaves the viz.
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
      const dx = Math.abs(dragCur.x - dragStart.x), dy = Math.abs(dragCur.y - dragStart.y);
      if (dx <= MARQUEE_THRESHOLD && dy <= MARQUEE_THRESHOLD) return;
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
  // Select every bar the rectangle overlaps. We hit-test the visible bar (not
  // the whole row group, whose box also spans the curve + label) and read the
  // tuple ids off its row ancestor.
  function applyMarquee(rect, additive) {
    const hits = new Set();
    host.querySelectorAll('.db-bar').forEach((bar) => {
      const r = bar.getBoundingClientRect();
      const overlaps = !(r.right < rect.left || r.left > rect.right ||
                         r.bottom < rect.top || r.top > rect.bottom);
      if (!overlaps) return;
      const row = bar.closest('.db-row');
      if (row && row._tupleIds) row._tupleIds.forEach((id) => hits.add(id));
    });
    if (!additive) selectedTuples = new Set();
    hits.forEach((id) => selectedTuples.add(id));
    pushSelection();
  }
  let marqueeEl = null;
  function ensureMarquee() {
    if (!marqueeEl) {
      marqueeEl = el('div', 'db-marquee');
      marqueeEl.style.display = 'none';
      document.body.appendChild(marqueeEl);
    }
    return marqueeEl;
  }
  function drawMarquee() {
    const r = marqueeRect(), m = ensureMarquee();
    m.style.left = r.left + 'px';
    m.style.top = r.top + 'px';
    m.style.width = (r.right - r.left) + 'px';
    m.style.height = (r.bottom - r.top) + 'px';
  }
  function hideMarquee() { if (marqueeEl) marqueeEl.style.display = 'none'; }

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
  function clearHoverState() {
    if (hoverEl) { hoverTuple(0); hoverEl = null; }
  }

  function pushSelection() {
    const ws = getWorksheet();
    if (ws && ws.selectTuplesAsync) {
      try { ws.selectTuplesAsync([...selectedTuples], SELECT_SIMPLE).catch(() => {}); }
      catch (e) { /* ignore */ }
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

  function applySelectionStyles() {
    const active = selectedTuples.size > 0;
    host.querySelectorAll('.db-interactive').forEach((node) => {
      if (!active) { node.classList.remove('db-dimmed', 'db-selected'); return; }
      const sel = node._tupleIds.every((id) => selectedTuples.has(id));
      node.classList.toggle('db-selected', sel);
      node.classList.toggle('db-dimmed', !sel);
    });
  }

  async function syncSelectionFromWorksheet(model) {
    const ws = getWorksheet();
    if (!ws || !ws.getSelectedMarksAsync || !model) return;
    try {
      const result = await ws.getSelectedMarksAsync();
      const tables = (result && result.data) || [];
      const keys = new Set();
      tables.forEach((tbl) => {
        const cols = tbl.columns || [];
        const ci = cols.findIndex((c) => c.fieldName === model.catField);
        if (ci < 0) return;
        (tbl.data || []).forEach((row) => keys.add(String(row[ci].value)));
      });
      const next = new Set();
      model.items.forEach((it) => { if (keys.has(it.rawKey)) next.add(it.tupleId); });
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
    document.documentElement.style.setProperty('--db-font', FONT_STACKS[config.fontFamily] || FONT_STACKS.system);
    document.documentElement.style.setProperty('--db-ink', config.textColor);
    document.documentElement.style.setProperty('--db-muted', config.mutedColor);
    document.documentElement.style.setProperty('--db-bg', config.bgColor);
  }

  /* =========================================================
     CONFIG MODAL (declarative schema)
     ========================================================= */
  const SCHEMA = [
    { section: 'General', fields: [
      { key: 'title', label: 'Chart title', type: 'text' },
      { key: 'bgColor', label: 'Background', type: 'color' },
      { key: 'textColor', label: 'Text color', type: 'color' },
      { key: 'mutedColor', label: 'Muted / value color', type: 'color' },
      { key: 'fontFamily', label: 'Font', type: 'select',
        options: [['system', 'System'], ['serif', 'Serif'], ['mono', 'Monospace']] },
      { key: 'enableTooltips', label: 'Tableau tooltips + selection', type: 'checkbox' },
    ]},
    { section: 'Root (grand total)', fields: [
      { key: 'showRoot', label: 'Show root total', type: 'checkbox' },
      { key: 'rootLabel', label: 'Caption', type: 'text' },
      { key: 'totalPrefix', label: 'Total prefix', type: 'text' },
      { key: 'totalDecimals', label: 'Total decimals', type: 'range', min: 0, max: 4, step: 1 },
      { key: 'rootCaptionSize', label: 'Caption size (px)', type: 'range', min: 8, max: 22, step: 1 },
      { key: 'rootValueSize', label: 'Total size (px)', type: 'range', min: 12, max: 48, step: 1 },
    ]},
    { section: 'Bars', fields: [
      { key: 'sort', label: 'Sort', type: 'select',
        options: [['desc', 'Value ↓'], ['asc', 'Value ↑'], ['none', 'Data order']] },
      { key: 'fitHeight', label: 'Fit to window height', type: 'checkbox' },
      { key: 'rowHeight', label: 'Row height (px) — when not fitting', type: 'range', min: 14, max: 64, step: 2 },
      { key: 'barThickness', label: 'Bar thickness', type: 'range', min: 0.15, max: 0.95, step: 0.05 },
      { key: 'barRadius', label: 'Corner radius (px)', type: 'range', min: 0, max: 32, step: 1 },
      { key: 'barMinWidth', label: 'Min bar width (px)', type: 'range', min: 0, max: 30, step: 1 },
      { key: 'barOpacity', label: 'Bar opacity', type: 'range', min: 0.2, max: 1, step: 0.05 },
    ]},
    { section: 'Fan curves', fields: [
      { key: 'showCurves', label: 'Show curves', type: 'checkbox' },
      { key: 'curveTension', label: 'Curviness', type: 'range', min: 0, max: 1, step: 0.05 },
      { key: 'curveWidth', label: 'Curve width (px)', type: 'range', min: 0.5, max: 6, step: 0.5 },
      { key: 'curveOpacity', label: 'Curve opacity', type: 'range', min: 0.05, max: 1, step: 0.05 },
      { key: 'curveColorMode', label: 'Curve color', type: 'select',
        options: [['match', 'Match bar'], ['flat', 'Flat color']] },
      { key: 'curveFlatColor', label: 'Flat curve color', type: 'color' },
    ]},
    { section: 'Color', fields: [
      { key: 'colorMode', label: 'Color mode', type: 'select',
        options: [['palette', 'Palette (by category)'], ['single', 'Single color'], ['divergent', 'Divergent (by value)']] },
      { key: 'singleColor', label: 'Single color', type: 'color' },
      { key: 'divLowColor', label: 'Divergent — low', type: 'color' },
      { key: 'divHighColor', label: 'Divergent — high', type: 'color' },
      { key: 'palette', label: 'Palette', type: 'select',
        options: [['pastel', 'Pastel'], ['tableau10', 'Tableau 10'], ['bold', 'Bold'], ['mono', 'Mono (single hue)']] },
      { key: 'monoBaseColor', label: 'Mono base color', type: 'color' },
    ]},
    { section: 'Labels', fields: [
      { key: 'showCategoryNames', label: 'Show category names', type: 'checkbox' },
      { key: 'showValueLabels', label: 'Show values', type: 'checkbox' },
      { key: 'labelSize', label: 'Label size (px)', type: 'range', min: 8, max: 22, step: 1 },
      { key: 'labelWidth', label: 'Label column width (px)', type: 'range', min: 60, max: 320, step: 10 },
    ]},
    { section: 'Number format (values)', fields: [
      { key: 'numPrefix', label: 'Prefix', type: 'text' },
      { key: 'numSuffix', label: 'Suffix', type: 'text' },
      { key: 'numDecimals', label: 'Decimals', type: 'range', min: 0, max: 4, step: 1 },
      { key: 'numUnit', label: 'Unit', type: 'select',
        options: [['none', 'None'], ['auto', 'Auto'], ['K', 'Thousands (K)'], ['M', 'Millions (M)'], ['B', 'Billions (B)']] },
      { key: 'numThousands', label: 'Thousands separator', type: 'checkbox' },
    ]},
    { section: 'Axis', fields: [
      { key: 'showAxis', label: 'Show value axis', type: 'checkbox' },
      { key: 'axisTicks', label: 'Tick count', type: 'range', min: 2, max: 12, step: 1 },
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

  /* ---------- config persistence ----------
     Config lives ONLY in the workbook settings — saved per viz instance, and
     travelling with the .twb/.twbx. We intentionally do NOT cache in
     localStorage: that cache is shared by every instance on this origin, so a
     brand-new viz would inherit the last-used colors instead of starting from
     the defaults. A fresh viz therefore has empty settings → pure DEFAULT_CONFIG.
     (render() only runs after loadConfig() resolves, so there's no first-paint
     flash that a cache would have prevented anyway.) */
  function loadConfig() {
    let saved = {};
    try {
      const store = tableau.extensions.settings;
      const raw = store && store.get(STORAGE_KEY);
      if (raw) saved = JSON.parse(raw);
    } catch (e) { console.error('reading workbook settings failed:', e); }
    return { ...DEFAULT_CONFIG, ...saved };
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
    host.innerHTML = '<div class="db-empty"><p>' + html + '</p></div>';
  }
  function showEmptyState(fields) {
    applyTheme();
    fields = fields || {};
    const chip = (id, label, kind) => {
      const got = fields[id] && fields[id].length;
      return '<span class="db-chip ' + kind + (got ? ' filled' : '') + '">' +
        label + (got ? ': ' + escapeHtml(fields[id].join(', ')) : '') + '</span>';
    };
    host.innerHTML =
      '<div class="db-empty"><div class="db-empty-card">' +
        '<div class="db-empty-head">' +
          '<div class="db-empty-thumb">' + buildThumb() + '</div>' +
          '<div class="db-empty-meta">' +
            '<div class="db-empty-title">Dendrogram Bars</div>' +
            '<div class="db-empty-desc">A fan-out bar chart: every category bar curves back to a ' +
            'single root total. Sorted, colored, and labelled — fully customizable from the &#9881; gear ' +
            'or the <b>Format Extension</b> button.</div>' +
          '</div>' +
        '</div>' +
        '<div class="db-empty-guide">' +
          '<div class="db-empty-guide-title">Get started</div>' +
          '<div class="db-empty-guide-row">Drag a <b>dimension</b> onto ' +
            chip('category', 'Category', 'dim') + ' and a <b>measure</b> onto ' +
            chip('value', 'Value', 'measure') + '.' +
          '</div>' +
        '</div>' +
      '</div></div>';
  }
  // Inline SVG thumbnail: a mini fan of curves into bars.
  function buildThumb() {
    const W = 96, H = 84, rx = 8, ry = H / 2, bx = 40, n = 6;
    const cols = ['#a5d6a7','#ef9a9a','#ce93d8','#ffcc80','#90caf9','#80cbc4'];
    let svg = '';
    for (let i = 0; i < n; i++) {
      const y = 12 + (i + 0.5) * ((H - 20) / n);
      const dx = (bx - rx) * 0.6;
      svg += `<path d="M ${rx} ${ry} C ${rx + dx} ${ry}, ${bx - dx} ${y}, ${bx} ${y}" fill="none" stroke="${cols[i]}" stroke-width="1.2" stroke-opacity="0.7"/>`;
      const w = 18 + (n - i) * 5;
      svg += `<rect x="${bx}" y="${y - 4}" width="${w}" height="8" rx="4" fill="${cols[i]}"/>`;
    }
    svg += `<circle cx="${rx}" cy="${ry}" r="2.5" fill="#e8e9ee"/>`;
    return `<svg viewBox="0 0 ${W} ${H}" width="84" height="74" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Dendrogram bars preview">${svg}</svg>`;
  }
  function escapeHtml(s) {
    return String(s).replace(/[&<>"]/g, (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }
})();

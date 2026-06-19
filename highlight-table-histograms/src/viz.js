'use strict';

/* =============================================================
   HIGHLIGHT TABLE + MARGINAL HISTOGRAMS — Tableau Viz Extension
   -------------------------------------------------------------
   A color-encoded matrix (rows x columns) with:
     • a vertical histogram of COLUMN totals across the top, and
     • a horizontal histogram of ROW totals down the right side.
   Everything visual is driven by `config` and editable live from
   the gear → settings modal.

   The four moves (see CLAUDE.md):
     1. CONNECT   — initializeAsync(), grab the worksheet
     2. ENCODINGS — Row / Column / Value tiles (manifest.trex)
     3. DATA      — read summary data, pivot into a matrix
     4. DRAW      — render(model) into #viz; redraw on data change
   ============================================================= */

(function () {
  const host = document.getElementById('viz');
  // Key for the local-cache copy of the config (workbook settings are the
  // source of truth and are isolated per extension instance). If you run other
  // extensions from this same localhost origin, give each a distinct key.
  const STORAGE_KEY = 'hth-config-v1';

  /* ---------- config ---------- */
  const DEFAULT_CONFIG = {
    // identity
    title: '',
    axisLabel: '',            // small label above the row labels (e.g. "Sales")
    bgColor: '#ffffff',
    fontFamily: 'system',

    // interactivity — native Tableau tooltips on hover + mark selection on click
    enableTooltips: true,

    // color scale (single-hue sequential)
    lowColor: '#ededf3',
    highColor: '#34346b',
    reverseScale: false,
    scalePower: 1,            // gamma: >1 emphasises high values, <1 the low end

    // matrix cells
    cellGap: 3,
    cellRadius: 0,
    showCellLabels: false,
    cellLabelSize: 11,

    // row / column labels
    showRowLabels: true,
    showColLabels: true,
    rowLabelWidth: 70,
    colLabelHeight: 22,
    labelSize: 12,

    // top histogram (column totals)
    showTopHist: true,
    topHistHeight: 200,
    showTopLabels: true,

    // right histogram (row totals)
    showRightHist: true,
    rightHistWidth: 300,
    showRightLabels: true,

    // shared histogram options
    colorHistograms: true,    // color bars by total intensity (vs. a flat high color)
    histGap: 16,              // space between the matrix and each histogram
    histBarRadius: 2,         // bar corner radius: 0 = square, higher = rounded

    // number formatting (totals + optional cell labels)
    numPrefix: '',
    numSuffix: '',
    numDecimals: 0,
    numUnit: 'auto',          // auto | K | M | B | none
    numThousands: true,
  };

  const FONT_STACKS = {
    system: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
    serif: 'Georgia, "Times New Roman", serif',
    mono: 'ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace',
  };

  // Start from defaults + the same-machine cache so the very first paint has
  // something sensible; the authoritative workbook settings are loaded once
  // initializeAsync() resolves (the settings API isn't ready before that).
  let config = { ...DEFAULT_CONFIG, ...loadLocalCache() };
  let lastModel = null;

  // 1 — CONNECT ------------------------------------------------
  // Registering a `configure` callback makes Tableau show the "Format
  // Extension" button on the Marks card; clicking it opens our settings modal.
  // (Same dialog as the in-viz ⚙ gear.)
  tableau.extensions.initializeAsync({ configure: openConfigModal }).then(
    () => {
      // Now that the API is initialized, read the settings stored in the
      // workbook (these survive save/close/reopen and travel with the file).
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

  // 2 + 3 — ENCODINGS + DATA → pivot into a matrix model -------
  async function update() {
    const ws = tableau.extensions.worksheetContent.worksheet;

    let fields = {};
    try {
      fields = await getEncodedFields(ws);
    } catch (e) {
      console.error('reading encodings failed:', e);
    }

    if (!fields.row || !fields.row.length ||
        !fields.column || !fields.column.length ||
        !fields.value || !fields.value.length) {
      lastModel = null;
      showMessage(
        'Drop at least one field on each tile in the Marks card: ' +
        '<b>Row</b>, <b>Column</b>, and <b>Value</b>. ' +
        'Row and Column accept several pills (nested headers).'
      );
      return;
    }

    let table = { columns: [], data: [] };
    try {
      table = await readSummary(ws);
    } catch (e) {
      console.error('reading summary data failed:', e);
    }

    lastModel = buildMatrix(ws.name, fields, table);
    // Re-sync selection to the (possibly changed) data before drawing.
    await syncSelectionFromWorksheet(lastModel);
    render(lastModel);
  }

  const SEP = String.fromCharCode(0);  // key separator (won't appear inside a field value)
  function joinKey(values) { return values.map(String).join(SEP); }

  // Pivot the flat summary rows into a matrix. Row and Column can each hold
  // SEVERAL fields → a row/column is the unique combination of its fields'
  // values; we keep the per-level parts for nested headers.
  function buildMatrix(worksheet, fields, table) {
    const cols = table.columns;
    const rowFields = fields.row;          // [name, ...] (outer -> inner)
    const colFields = fields.column;
    const valueField = fields.value[0];    // a single measure
    const rowIdxs = rowFields.map((n) => findColumn(cols, n));
    const colIdxs = colFields.map((n) => findColumn(cols, n));
    const valIdx = findValueColumn(cols, valueField);

    const rowKeys = [], colKeys = [];      // composite keys (identity)
    const rowParts = [], colParts = [];    // per-level value arrays (display)
    const rowPos = new Map(), colPos = new Map();
    const matrix = [];          // matrix[r][c] = number | null
    // Tuple ids (1-based summary-data row index) backing each visual element,
    // for hoverTupleAsync / selectTuplesAsync.
    const cellTuples = [];      // cellTuples[r][c] = [tupleId, …]
    const rowTupleIds = [];     // rowTupleIds[r]   = [tupleId, …]  (whole row)
    const colTupleIds = [];     // colTupleIds[c]   = [tupleId, …]  (whole column)
    const tupleKeys = {};       // tupleId → "rawRow\u0000rawCol" (to match selected marks)

    table.data.forEach((dataRow, idx) => {
      const tupleId = idx + 1;  // tuple ids are 1-based
      const rParts = rowIdxs.map((i) => cellText(dataRow[i]));
      const cParts = colIdxs.map((i) => cellText(dataRow[i]));
      const rk = rParts.join(SEP);
      const ck = cParts.join(SEP);
      // Stable mark key across ALL row+col fields (raw values), to match the
      // marks returned by getSelectedMarksAsync.
      const rawKey = joinKey([...rowIdxs, ...colIdxs].map((i) => dataRow[i] ? dataRow[i].value : ''));
      const raw = dataRow[valIdx] ? Number(dataRow[valIdx].value) : NaN;
      const v = Number.isFinite(raw) ? raw : null;

      if (!rowPos.has(rk)) { rowPos.set(rk, rowKeys.length); rowKeys.push(rk); rowParts.push(rParts); }
      if (!colPos.has(ck)) { colPos.set(ck, colKeys.length); colKeys.push(ck); colParts.push(cParts); }
      const r = rowPos.get(rk), c = colPos.get(ck);

      if (!matrix[r]) matrix[r] = [];
      if (!cellTuples[r]) cellTuples[r] = [];
      if (v != null) {
        matrix[r][c] = (matrix[r][c] || 0) + v;
        (cellTuples[r][c] || (cellTuples[r][c] = [])).push(tupleId);
        (rowTupleIds[r] || (rowTupleIds[r] = [])).push(tupleId);
        (colTupleIds[c] || (colTupleIds[c] = [])).push(tupleId);
        tupleKeys[tupleId] = rawKey;
      }
    });

    // Fill ragged rows and compute marginal totals.
    let cellMin = Infinity, cellMax = -Infinity;
    const rowTotals = rowKeys.map(() => null);
    const colTotals = colKeys.map(() => null);

    for (let r = 0; r < rowKeys.length; r++) {
      if (!matrix[r]) matrix[r] = [];
      for (let c = 0; c < colKeys.length; c++) {
        const v = matrix[r][c];
        if (v == null) { matrix[r][c] = null; continue; }
        cellMin = Math.min(cellMin, v);
        cellMax = Math.max(cellMax, v);
        rowTotals[r] = (rowTotals[r] || 0) + v;
        colTotals[c] = (colTotals[c] || 0) + v;
      }
    }
    if (!Number.isFinite(cellMin)) { cellMin = 0; cellMax = 0; }

    return {
      worksheet, fields, rowFields, colFields, valueField,
      rowKeys, colKeys, rowParts, colParts, matrix,
      rowTotals, colTotals, cellMin, cellMax,
      cellTuples, rowTupleIds, colTupleIds, tupleKeys,
    };
  }

  // Group consecutive rows/columns that share the same parts[0..level] prefix
  // (for merged, nested headers). Returns [{ start, end, label }] runs.
  function levelGroups(partsArr, level) {
    const groups = [];
    const keyAt = (i) => partsArr[i].slice(0, level + 1).join(SEP);
    let start = 0;
    for (let i = 1; i <= partsArr.length; i++) {
      if (i === partsArr.length || keyAt(i) !== keyAt(start)) {
        groups.push({ start, end: i - 1, label: partsArr[start][level] });
        start = i;
      }
    }
    return groups;
  }
  // Union of tuple ids / sum of totals across a contiguous range [s, e].
  function rangeTuples(arr, s, e) {
    const out = [];
    for (let i = s; i <= e; i++) if (arr[i]) out.push.apply(out, arr[i]);
    return out;
  }
  function rangeTotal(arr, s, e) {
    let t = null;
    for (let i = s; i <= e; i++) if (arr[i] != null) t = (t || 0) + arr[i];
    return t;
  }
  // Marginal-tooltip payload for a header/bar: one line per field down to
  // `level`, then the measure total. Returns null when there's nothing to show.
  function buildHeaderTip(fields, parts, level, valueField, total) {
    if (total == null) return null;
    const dims = [];
    for (let k = 0; k <= level; k++) dims.push({ field: fields[k], label: parts[k] });
    return { dims, valueField, total };
  }

  /* ===========================================================
     ▼▼▼  BUILD ZONE — render the chart from the model  ▼▼▼
     =========================================================== */
  function render(model) {
    if (!model) return;
    applyTheme();
    host.innerHTML = '';

    const { rowKeys, colKeys } = model;
    const nRows = rowKeys.length, nCols = colKeys.length;
    if (!nRows || !nCols) {
      showMessage('No data to display yet.');
      return;
    }

    const wrap = el('div', 'hth');
    if (config.title) {
      wrap.appendChild(el('div', 'hth-title', config.title));
    }

    // Row / Column can each be several fields → that many header LEVELS.
    const rowFields = model.rowFields, colFields = model.colFields;
    const Rn = config.showRowLabels ? rowFields.length : 0;  // row header levels
    const Cn = config.showColLabels ? colFields.length : 0;  // column header levels
    const T  = config.showTopHist ? 1 : 0;
    const matrixCol = Rn + 1;            // 1-based grid line where the matrix starts
    const matrixRow = T + Cn + 1;

    // One CSS grid; every region is placed explicitly so the row-label levels,
    // matrix and right histogram share row tracks (and the top histogram,
    // column-label levels and matrix share column tracks) — pixel-aligned.
    const grid = el('div', 'hth-grid');
    grid.style.setProperty('--cell-gap', config.cellGap + 'px');
    grid.style.setProperty('--hist-gap', config.histGap + 'px');

    const colTracks = [];
    for (let i = 0; i < Rn; i++) colTracks.push(config.rowLabelWidth + 'px');
    colTracks.push('minmax(0, 1fr)');
    if (config.showRightHist) colTracks.push(config.rightHistWidth + 'px');
    grid.style.gridTemplateColumns = colTracks.join(' ');

    const rowTracks = [];
    if (config.showTopHist) rowTracks.push(config.topHistHeight + 'px');
    for (let j = 0; j < Cn; j++) rowTracks.push(config.colLabelHeight + 'px');
    rowTracks.push('minmax(0, 1fr)');
    grid.style.gridTemplateRows = rowTracks.join(' ');

    const place = (node, r1, r2, c1, c2) => {
      node.style.gridRow = r1 + ' / ' + r2;
      node.style.gridColumn = c1 + ' / ' + c2;
      grid.appendChild(node);
    };

    // corner (top-left), holds the optional axis label
    if (Rn > 0 && (T > 0 || Cn > 0)) {
      const corner = el('div', 'hth-corner');
      if (config.axisLabel) corner.appendChild(el('span', 'hth-axis', config.axisLabel));
      place(corner, 1, matrixRow, 1, matrixCol);
    }

    // top histogram (column totals)
    if (config.showTopHist) {
      const th = buildTopHist(model);
      th.style.marginBottom = config.histGap + 'px';   // gap from the matrix
      place(th, 1, 2, matrixCol, matrixCol + 1);
    }

    // column-label levels (nested, merged) — one grid row per field
    if (config.showColLabels) {
      colFields.forEach((field, j) => {
        const lvl = el('div', 'hth-collabels');
        lvl.style.gridTemplateColumns = `repeat(${nCols}, minmax(0, 1fr))`;
        levelGroups(model.colParts, j).forEach((g) => {
          const cls = 'hth-collabel' + (j < colFields.length - 1 ? ' parent' : '');
          const lab = el('div', cls, g.label);
          lab.style.gridColumn = (g.start + 1) + ' / ' + (g.end + 2);
          markInteractive(lab, rangeTuples(model.colTupleIds, g.start, g.end),
            buildHeaderTip(colFields, model.colParts[g.start], j, model.valueField,
              rangeTotal(model.colTotals, g.start, g.end)));
          lvl.appendChild(lab);
        });
        place(lvl, T + 1 + j, T + 2 + j, matrixCol, matrixCol + 1);
      });
    }

    // row-label levels (nested, merged) — one grid column per field
    if (config.showRowLabels) {
      rowFields.forEach((field, i) => {
        const lvl = el('div', 'hth-rowlabels');
        lvl.style.gridTemplateRows = `repeat(${nRows}, minmax(0, 1fr))`;
        levelGroups(model.rowParts, i).forEach((g) => {
          const cls = 'hth-rowlabel' + (i < rowFields.length - 1 ? ' parent' : '');
          const lab = el('div', cls, g.label);
          lab.style.gridRow = (g.start + 1) + ' / ' + (g.end + 2);
          markInteractive(lab, rangeTuples(model.rowTupleIds, g.start, g.end),
            buildHeaderTip(rowFields, model.rowParts[g.start], i, model.valueField,
              rangeTotal(model.rowTotals, g.start, g.end)));
          lvl.appendChild(lab);
        });
        place(lvl, matrixRow, matrixRow + 1, 1 + i, 2 + i);
      });
    }

    // the matrix
    place(buildMatrixGrid(model), matrixRow, matrixRow + 1, matrixCol, matrixCol + 1);

    // right histogram (row totals)
    if (config.showRightHist) {
      const rh = buildRightHist(model);
      rh.style.marginLeft = config.histGap + 'px';     // gap from the matrix
      place(rh, matrixRow, matrixRow + 1, matrixCol + 1, matrixCol + 2);
    }

    wrap.appendChild(grid);
    host.appendChild(wrap);

    // Now that bars have real pixel widths, decide whether each right-hist
    // label fits inside its bar or must sit just outside it.
    layoutRightLabels();

    // Re-apply selection dimming/highlight (the DOM was just rebuilt).
    applySelectionStyles();
  }

  // Place each row-total label inside its bar when it fits, otherwise just to
  // the right of the bar (so short bars don't have overflowing/overlapping
  // labels). Re-runs on every render (incl. resize) since it measures pixels.
  function layoutRightLabels() {
    const PAD = 8;
    const outsideInk = contrastInk(config.bgColor);
    host.querySelectorAll('.hth-righthist-row').forEach((row) => {
      const bar = row.querySelector('.hth-bar-h');
      const lbl = row.querySelector('.hth-bar-label-h');
      if (!bar || !lbl) return;
      const barW = bar.getBoundingClientRect().width;
      const labelW = lbl.getBoundingClientRect().width;
      if (labelW + PAD * 2 <= barW) {
        lbl.style.left = (barW - labelW - PAD) + 'px';
        lbl.style.color = contrastInk(lbl.dataset.bg || config.highColor);
      } else {
        lbl.style.left = (barW + 6) + 'px';
        lbl.style.color = outsideInk;
      }
    });
  }

  function buildMatrixGrid(model) {
    const { matrix, rowKeys, colKeys, rowParts, colParts, cellMin, cellMax, cellTuples } = model;
    const m = el('div', 'hth-matrix');
    m.style.gridTemplateColumns = `repeat(${colKeys.length}, minmax(0, 1fr))`;
    m.style.gridTemplateRows = `repeat(${rowKeys.length}, minmax(0, 1fr))`;

    for (let r = 0; r < rowKeys.length; r++) {
      for (let c = 0; c < colKeys.length; c++) {
        const v = matrix[r][c];
        const cell = el('div', 'hth-cell');
        cell.style.borderRadius = config.cellRadius + 'px';
        if (v == null) {
          cell.classList.add('empty');
        } else {
          const t = norm(v, cellMin, cellMax);
          const bg = scaleColor(t);
          cell.style.background = bg;
          // Only use the browser's native `title` tooltip as a fallback when
          // interactivity is off — otherwise it double-renders on top of the
          // (richer) Tableau tooltip driven by hoverTupleAsync.
          if (!config.enableTooltips) {
            cell.title = `${rowParts[r].join(' / ')} · ${colParts[c].join(' / ')}: ${formatNumber(v)}`;
          }
          markInteractive(cell, cellTuples[r] && cellTuples[r][c]);
          if (config.showCellLabels) {
            const lbl = el('span', 'hth-cell-label', formatNumber(v));
            lbl.style.color = contrastInk(bg);
            lbl.style.fontSize = config.cellLabelSize + 'px';
            cell.appendChild(lbl);
          }
        }
        m.appendChild(cell);
      }
    }
    return m;
  }

  function buildTopHist(model) {
    const { colKeys, colTotals, colTupleIds, colParts, colFields, valueField } = model;
    const max = Math.max(0, ...colTotals.filter((v) => v != null));
    const tmin = Math.min(...colTotals.filter((v) => v != null));
    const tmax = Math.max(...colTotals.filter((v) => v != null));

    const box = el('div', 'hth-tophist');
    box.style.gridTemplateColumns = `repeat(${colKeys.length}, minmax(0, 1fr))`;
    const ceiling = config.showTopLabels ? 88 : 100; // leave room for the value label
    colTotals.forEach((tot, ci) => {
      const colCell = el('div', 'hth-tophist-col');
      // A column bar represents every mark in that column → custom tooltip
      // shows the column total (the native per-mark tooltip can't).
      markInteractive(colCell, colTupleIds[ci],
        buildHeaderTip(colFields, colParts[ci], colFields.length - 1, valueField, tot));
      const bar = el('div', 'hth-bar-v');
      const h = max > 0 && tot != null ? (tot / max) * ceiling : 0;
      bar.style.height = h + '%';
      bar.style.borderRadius = `${config.histBarRadius}px ${config.histBarRadius}px 0 0`;
      bar.style.background = config.colorHistograms
        ? scaleColor(norm(tot, tmin, tmax))
        : config.highColor;
      if (config.showTopLabels && tot != null) {
        const lbl = el('div', 'hth-bar-label', formatNumber(tot));
        lbl.style.fontSize = config.labelSize + 'px';
        colCell.appendChild(lbl);
      }
      colCell.appendChild(bar);
      box.appendChild(colCell);
    });
    return box;
  }

  function buildRightHist(model) {
    const { rowKeys, rowTotals, rowTupleIds, rowParts, rowFields, valueField } = model;
    const max = Math.max(0, ...rowTotals.filter((v) => v != null));
    const tmin = Math.min(...rowTotals.filter((v) => v != null));
    const tmax = Math.max(...rowTotals.filter((v) => v != null));

    const box = el('div', 'hth-righthist');
    box.style.gridTemplateRows = `repeat(${rowKeys.length}, minmax(0, 1fr))`;
    rowTotals.forEach((tot, ri) => {
      const rowCell = el('div', 'hth-righthist-row');
      // A row bar represents every mark in that row → custom tooltip shows
      // the row total (the native per-mark tooltip can't).
      markInteractive(rowCell, rowTupleIds[ri],
        buildHeaderTip(rowFields, rowParts[ri], rowFields.length - 1, valueField, tot));
      const bar = el('div', 'hth-bar-h');
      const w = max > 0 && tot != null ? (tot / max) * 100 : 0;
      bar.style.width = w + '%';
      bar.style.borderRadius = `0 ${config.histBarRadius}px ${config.histBarRadius}px 0`;
      const bg = config.colorHistograms ? scaleColor(norm(tot, tmin, tmax)) : config.highColor;
      bar.style.background = bg;
      rowCell.appendChild(bar);
      if (config.showRightLabels && tot != null) {
        // Label is a sibling of the bar (not a child): layoutRightLabels()
        // positions it inside or outside the bar once widths are known.
        const lbl = el('span', 'hth-bar-label-h', formatNumber(tot));
        lbl.style.fontSize = config.labelSize + 'px';
        lbl.dataset.bg = bg;
        rowCell.appendChild(lbl);
      }
      box.appendChild(rowCell);
    });
    return box;
  }
  /* ▲▲▲  END BUILD ZONE  ▲▲▲ */

  /* ---------- color + number helpers ---------- */
  function norm(v, min, max) {
    if (v == null || max <= min) return 0;
    let t = (v - min) / (max - min);
    t = Math.max(0, Math.min(1, t));
    return Math.pow(t, config.scalePower);
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
    if (color.startsWith('rgb')) {
      rgb = color.match(/\d+/g).map(Number);
    } else {
      rgb = hexToRgb(color);
    }
    // Relative luminance (sRGB approximation).
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
    const suffix = unit + config.numSuffix;
    let str = n.toFixed(config.numDecimals);
    if (config.numThousands) {
      const parts = str.split('.');
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      str = parts.join('.');
    }
    return config.numPrefix + str + suffix;
  }

  /* =========================================================
     INTERACTIVITY — native Tableau tooltips + mark selection
     ---------------------------------------------------------
     Each interactive element carries the tuple ids (1-based
     summary-data rows) it represents. Hover → hoverTupleAsync
     shows Tableau's own tooltip; click → selectTuplesAsync
     selects the marks (driving highlight + dashboard actions).
     Listeners are delegated off #viz and attached once.
     ========================================================= */
  const SELECT_SIMPLE = (tableau.SelectOptions && tableau.SelectOptions.Simple) || 'select-options-simple';
  const SELECT_TOGGLE = (tableau.SelectOptions && tableau.SelectOptions.Toggle) || 'select-options-toggle';

  // Tag an element as a mark: stash its tuple ids + show a pointer cursor.
  // `marginal` (optional) marks an aggregate element (a histogram bar): such
  // elements get a custom tooltip showing the correct total, since the native
  // hoverTupleAsync tooltip can only describe a single mark.
  function markInteractive(node, tupleIds, marginal) {
    if (!config.enableTooltips || !tupleIds || !tupleIds.length) return;
    node.classList.add('hth-interactive');
    node._tupleIds = tupleIds;
    if (marginal) node._marginal = marginal;
  }

  function getWorksheet() {
    return tableau.extensions.worksheetContent &&
           tableau.extensions.worksheetContent.worksheet;
  }

  function hoverTuple(tupleId, ev) {
    const ws = getWorksheet();
    if (!ws || !ws.hoverTupleAsync) return;
    const tooltip = ev ? { tooltipAnchorPoint: { x: ev.clientX, y: ev.clientY } } : undefined;
    try { ws.hoverTupleAsync(tupleId, tooltip).catch(() => {}); } catch (e) { /* ignore */ }
  }

  // Push the current local selection to Tableau. We always send the full set
  // with SelectOptions.Simple (computing add/remove/toggle ourselves) — an
  // empty array clears the worksheet selection. Mirrors Tableau's own Sankey.
  function pushSelection() {
    const ws = getWorksheet();
    if (!ws || !ws.selectTuplesAsync) return;
    try {
      ws.selectTuplesAsync([...selectedTuples], SELECT_SIMPLE).catch(() => {});
    } catch (e) { /* ignore */ }
    applySelectionStyles();
  }

  // The set of currently selected tuple ids (drives both the worksheet
  // selection and our own dim/highlight styling).
  let selectedTuples = new Set();
  let hoverEl = null;     // element currently driving the tooltip
  let hoverAt = 0;        // last hover update timestamp (throttle)

  function wireInteractivity() {
    host.addEventListener('mousemove', (ev) => {
      if (!config.enableTooltips) return;
      const t = ev.target.closest && ev.target.closest('.hth-interactive');
      if (!t) { clearHoverState(); return; }

      if (t._marginal) {
        // Histogram bar → custom aggregate tooltip; never the per-mark one.
        if (hoverEl !== null) hoverTuple(0);   // drop any native cell tooltip
        hoverEl = null;
        showMarginalTip(t._marginal, ev);
      } else {
        // Matrix cell → native Tableau tooltip (single mark, correct).
        hideMarginalTip();
        const now = (window.performance && performance.now()) || Date.now();
        if (t !== hoverEl || now - hoverAt > 50) {  // throttle position updates
          hoverEl = t;
          hoverAt = now;
          hoverTuple(t._tupleIds[0], ev);
        }
      }
    });
    host.addEventListener('mouseleave', clearHoverState);
    host.addEventListener('click', (ev) => {
      if (!config.enableTooltips) return;
      const t = ev.target.closest && ev.target.closest('.hth-interactive');
      const ids = t && t._tupleIds ? t._tupleIds : null;
      toggleSelection(ids, ev.ctrlKey || ev.metaKey);
    });
  }

  // Update the local selection on click, mirroring native Tableau behaviour:
  //   • click a mark → select it (replacing the previous selection)
  //   • click the only selected mark again → deselect (clear)
  //   • Ctrl/Cmd-click → add/remove that mark from the selection
  //   • click empty space → clear (unless Ctrl/Cmd held)
  function toggleSelection(ids, additive) {
    if (!ids || !ids.length) {
      if (!additive && selectedTuples.size) { selectedTuples.clear(); pushSelection(); }
      return;
    }
    const allSelected = ids.every((id) => selectedTuples.has(id));
    if (allSelected) {
      if (additive) {
        ids.forEach((id) => selectedTuples.delete(id));        // remove from set
      } else if (selectedTuples.size === ids.length) {
        selectedTuples.clear();                                // only this → deselect
      } else {
        selectedTuples = new Set(ids);                         // collapse to this
      }
    } else {
      if (!additive) selectedTuples.clear();
      ids.forEach((id) => selectedTuples.add(id));             // add to set
    }
    pushSelection();
  }

  // Dim marks that aren't part of the selection; full opacity when nothing is
  // selected. A mark counts as selected when all its tuples are selected, which
  // gives a natural cross-highlight (clicking a row bar lights up that row).
  function applySelectionStyles() {
    const active = selectedTuples.size > 0;
    host.querySelectorAll('.hth-interactive').forEach((node) => {
      if (!active) { node.classList.remove('hth-dimmed', 'hth-selected'); return; }
      const sel = node._tupleIds.every((id) => selectedTuples.has(id));
      node.classList.toggle('hth-selected', sel);
      node.classList.toggle('hth-dimmed', !sel);
    });
  }

  // Re-derive the local selection from the worksheet's actual selected marks
  // (keeps us correct after filters change tuple ids, and reflects highlights
  // driven from other sheets in a dashboard).
  async function syncSelectionFromWorksheet(model) {
    const ws = getWorksheet();
    if (!ws || !ws.getSelectedMarksAsync || !model || !model.tupleKeys) return;
    try {
      const result = await ws.getSelectedMarksAsync();
      const tables = (result && result.data) || [];
      const fieldNames = [...model.rowFields, ...model.colFields];
      const keys = new Set();
      tables.forEach((tbl) => {
        const cols = tbl.columns || [];
        const idxs = fieldNames.map((n) => cols.findIndex((c) => c.fieldName === n));
        if (idxs.some((i) => i < 0)) return;
        (tbl.data || []).forEach((row) => {
          keys.add(joinKey(idxs.map((i) => row[i].value)));
        });
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
    hideMarginalTip();
  }

  /* ---------- custom tooltip for histogram (aggregate) bars ---------- */
  let tipEl = null;
  function ensureTip() {
    if (!tipEl) {
      tipEl = el('div', 'hth-tip');
      tipEl.style.display = 'none';
      document.body.appendChild(tipEl);
    }
    return tipEl;
  }
  function showMarginalTip(info, ev) {
    const tip = ensureTip();
    let html = '';
    info.dims.forEach((d) => {
      html += '<div class="hth-tip-row"><span class="hth-tip-k">' + escapeHtml(d.field) +
        '</span><span class="hth-tip-v">' + escapeHtml(d.label) + '</span></div>';
    });
    html += '<div class="hth-tip-row"><span class="hth-tip-k">' + escapeHtml(info.valueField) +
      '</span><span class="hth-tip-v">' + escapeHtml(formatFull(info.total)) + '</span></div>';
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
  function hideMarginalTip() {
    if (tipEl) tipEl.style.display = 'none';
  }

  // Full (non-abbreviated) number for tooltips, with thousands separators.
  function formatFull(v) {
    if (v == null || !Number.isFinite(v)) return '';
    let str = Number(v).toFixed(config.numDecimals);
    if (config.numThousands) {
      const parts = str.split('.');
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      str = parts.join('.');
    }
    return config.numPrefix + str + config.numSuffix;
  }

  /* ---------- plumbing (encodings + summary data) ---------- */
  async function getEncodedFields(ws) {
    const spec = await ws.getVisualSpecificationAsync();
    const marks = spec.marksSpecifications[spec.activeMarksSpecificationIndex];
    // Each tile can hold several pills → collect ALL fields per encoding id,
    // in card order. (Multiple encodings share the same id, one per field.)
    const fields = {};
    for (const enc of marks.encodings) {
      if (enc.field) (fields[enc.id] || (fields[enc.id] = [])).push(enc.field.name);
    }
    return fields;
  }

  async function readSummary(ws) {
    // ignoreSelection: true is essential — otherwise, when marks are selected
    // (e.g. a selection restored on workbook reopen), the reader returns ONLY
    // the selected marks, so the chart renders just that subset (one cell).
    const reader = await ws.getSummaryDataReaderAsync(undefined, { ignoreSelection: true });
    let columns = [];
    let data = [];
    for (let p = 0; p < reader.pageCount; p++) {
      const page = await reader.getPageAsync(p);
      columns = page.columns;
      data = data.concat(page.data);
    }
    if (reader.releaseAsync) {
      try { await reader.releaseAsync(); } catch (e) { /* ignore */ }
    }
    return { columns, data };
  }

  // Match a summary column to an encoding field by name; fall back to first
  // column whose name contains the field name (handles "SUM(Sales)" etc.).
  function findColumn(cols, fieldName) {
    let i = cols.findIndex((c) => c.fieldName === fieldName);
    if (i < 0) i = cols.findIndex((c) => c.fieldName && c.fieldName.indexOf(fieldName) >= 0);
    return i < 0 ? 0 : i;
  }

  function findValueColumn(cols, fieldName) {
    let i = cols.findIndex((c) => c.fieldName === fieldName);
    if (i < 0) i = cols.findIndex((c) => c.fieldName && c.fieldName.indexOf(fieldName) >= 0);
    if (i < 0) i = cols.findIndex((c) => isNumericCol(c)); // any measure
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
      '--hth-font', FONT_STACKS[config.fontFamily] || FONT_STACKS.system
    );
    document.documentElement.style.setProperty('--hth-label-size', config.labelSize + 'px');
  }

  /* =========================================================
     CONFIG MODAL
     ========================================================= */
  // Declarative schema: each control reads/writes one config key.
  const SCHEMA = [
    { section: 'General', fields: [
      { key: 'title', label: 'Chart title', type: 'text' },
      { key: 'axisLabel', label: 'Axis label (top-left)', type: 'text' },
      { key: 'bgColor', label: 'Background', type: 'color' },
      { key: 'fontFamily', label: 'Font', type: 'select',
        options: [['system', 'System'], ['serif', 'Serif'], ['mono', 'Monospace']] },
      { key: 'enableTooltips', label: 'Tableau tooltips + selection', type: 'checkbox' },
    ]},
    { section: 'Color scale', fields: [
      { key: 'lowColor', label: 'Low color', type: 'color' },
      { key: 'highColor', label: 'High color', type: 'color' },
      { key: 'reverseScale', label: 'Reverse scale', type: 'checkbox' },
      { key: 'scalePower', label: 'Emphasis (gamma)', type: 'range', min: 0.3, max: 3, step: 0.1 },
    ]},
    { section: 'Matrix cells', fields: [
      { key: 'cellGap', label: 'Cell gap (px)', type: 'range', min: 0, max: 12, step: 1 },
      { key: 'cellRadius', label: 'Corner radius (px)', type: 'range', min: 0, max: 16, step: 1 },
      { key: 'showCellLabels', label: 'Show value in cells', type: 'checkbox' },
      { key: 'cellLabelSize', label: 'Cell label size (px)', type: 'range', min: 7, max: 20, step: 1 },
    ]},
    { section: 'Labels', fields: [
      { key: 'showRowLabels', label: 'Show row labels', type: 'checkbox' },
      { key: 'showColLabels', label: 'Show column labels', type: 'checkbox' },
      { key: 'rowLabelWidth', label: 'Row label width (px)', type: 'range', min: 30, max: 200, step: 5 },
      { key: 'colLabelHeight', label: 'Column label height (px)', type: 'range', min: 14, max: 60, step: 2 },
      { key: 'labelSize', label: 'Label font size (px)', type: 'range', min: 8, max: 20, step: 1 },
    ]},
    { section: 'Top histogram (column totals)', fields: [
      { key: 'showTopHist', label: 'Show', type: 'checkbox' },
      { key: 'topHistHeight', label: 'Height (px)', type: 'range', min: 60, max: 400, step: 10 },
      { key: 'showTopLabels', label: 'Show value labels', type: 'checkbox' },
    ]},
    { section: 'Right histogram (row totals)', fields: [
      { key: 'showRightHist', label: 'Show', type: 'checkbox' },
      { key: 'rightHistWidth', label: 'Width (px)', type: 'range', min: 80, max: 500, step: 10 },
      { key: 'showRightLabels', label: 'Show value labels', type: 'checkbox' },
    ]},
    { section: 'Histograms (shared)', fields: [
      { key: 'colorHistograms', label: 'Color bars by intensity', type: 'checkbox' },
      { key: 'histBarRadius', label: 'Bar corners (0 = square)', type: 'range', min: 0, max: 16, step: 1 },
      { key: 'histGap', label: 'Gap from matrix (px)', type: 'range', min: 0, max: 48, step: 2 },
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

  // Open the settings modal. Used by both the in-viz ⚙ gear and Tableau's
  // "Format Extension" button (registered as the `configure` callback).
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

    // Closing the modal forces an immediate workbook save (cancels the
    // debounce) so settings are committed the moment the user is done — even
    // if they save the workbook right away.
    const close = () => { overlay.hidden = true; clearTimeout(saveTimer); flushWorkbookSettings(); };

    openBtn.addEventListener('click', openConfigModal);
    doneBtn.addEventListener('click', close);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !overlay.hidden) close();
    });
    resetBtn.addEventListener('click', () => {
      config = { ...DEFAULT_CONFIG };
      saveConfig(config);
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
    saveConfig(config);
    if (lastModel) render(lastModel);
  }

  /* ---------- config persistence (workbook settings + localStorage) ----------
     Two stores:
       • tableau.extensions.settings — saved INTO the workbook. Survives
         save/close/reopen and travels with the .twb/.twbx. This is the one
         that matters; it requires initializeAsync() to have resolved first.
       • localStorage — a same-machine cache for instant first paint only.
     saveAsync() must be serialized: Tableau allows only one save in flight,
     and a second call while one is pending rejects. We also debounce so a
     slider drag doesn't fire dozens of saves. ---------------------------- */
  function loadConfig() {
    let saved = {};
    try {
      const store = tableau.extensions.settings;
      const raw = store && store.get(STORAGE_KEY);
      if (raw) saved = JSON.parse(raw);
    } catch (e) {
      console.error('reading workbook settings failed:', e);
    }
    if (!Object.keys(saved).length) saved = loadLocalCache();
    return { ...DEFAULT_CONFIG, ...saved };
  }

  function loadLocalCache() {
    try {
      const ls = window.localStorage.getItem(STORAGE_KEY);
      if (ls) return JSON.parse(ls);
    } catch (e) { /* ignore */ }
    return {};
  }

  let saveTimer = null;   // debounce handle
  let saveInFlight = false;
  let savePending = false;

  function saveConfig(cfg) {
    // Same-machine cache: write immediately so a quick reload restores state
    // even before the (debounced) workbook save lands.
    try { window.localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg)); } catch (e) { /* ignore */ }

    // Workbook settings: debounce, then persist. The actual values are read
    // from the live `config` at flush time, so we always store the latest.
    clearTimeout(saveTimer);
    saveTimer = setTimeout(flushWorkbookSettings, 350);
  }

  function flushWorkbookSettings() {
    const store = tableau.extensions.settings;
    if (!store) return; // settings API not available yet
    try {
      store.set(STORAGE_KEY, JSON.stringify(config));
    } catch (e) {
      console.error('settings.set failed:', e);
      return;
    }
    if (saveInFlight) { savePending = true; return; } // serialize saveAsync
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

  /* ---------- tiny DOM helpers ---------- */
  function el(tag, cls, text) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }
  function showMessage(html) {
    applyTheme();
    host.innerHTML = '<div class="hth-empty"><p>' + html + '</p></div>';
  }
  function escapeHtml(s) {
    return String(s).replace(/[&<>"]/g, (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }
})();

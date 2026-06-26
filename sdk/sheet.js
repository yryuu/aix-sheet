/**
 * aix-sheet SDK (ES Module)
 *
 *   import { Sheet } from 'aix-sheet';
 *   import { Sheet } from './sdk/sheet.js';
 */

// ============================================================
// A1 notation helpers
// ============================================================
function colToIdx(label) {
  let c = 0;
  for (const ch of label.toUpperCase()) {
    if (ch < 'A' || ch > 'Z') throw err('INVALID_REF', `列ラベルが不正: "${label}"`);
    c = c * 26 + ch.charCodeAt(0) - 64;
  }
  return c - 1;
}
function idxToCol(c) {
  let s = '';
  c++;
  while (c > 0) { c--; s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26); }
  return s;
}
function parseRef(ref) {
  if (typeof ref !== 'string') throw err('INVALID_REF', `セル参照は文字列で指定 (例: "A1"). 受け取った値: ${JSON.stringify(ref)}`);
  const m = ref.match(/^([A-Za-z]+)(\d+)$/);
  if (!m) throw err('INVALID_REF', `セル参照の形式が不正: "${ref}". A1 形式で指定してください (例: "A1", "B12", "AA3")`);
  return { col: colToIdx(m[1]), row: parseInt(m[2]) - 1, label: m[1].toUpperCase() + m[2] };
}
function parseRange(ref) {
  if (typeof ref !== 'string') throw err('INVALID_REF', `範囲は文字列で指定 (例: "A1:B5")`);
  if (!ref.includes(':')) {
    const p = parseRef(ref);
    return { r1: p.row, r2: p.row, c1: p.col, c2: p.col };
  }
  const [a, b] = ref.split(':');
  const pa = parseRef(a), pb = parseRef(b);
  return {
    r1: Math.min(pa.row, pb.row), r2: Math.max(pa.row, pb.row),
    c1: Math.min(pa.col, pb.col), c2: Math.max(pa.col, pb.col)
  };
}
function makeRef(r, c) { return idxToCol(c) + (r + 1); }

// ============================================================
// Custom errors (LLM-friendly messages)
// ============================================================
function err(code, msg) {
  const e = new Error(`[${code}] ${msg}`);
  e.code = code;
  return e;
}

// ============================================================
// Style validation
// ============================================================
const STYLE_KEYS = new Set(['bold', 'italic', 'underline', 'color', 'bgColor',
                            'align', 'fontSize', 'fontFamily', 'border', 'numFmt']);
const ALIGN_VALUES = new Set(['left', 'center', 'right']);
const BORDER_STYLES = new Set(['thin', 'medium', 'thick', 'dotted', 'dashed', 'double']);
const BORDER_SIDES  = ['top', 'right', 'bottom', 'left'];

/**
 * Normalize a border spec to { top?, right?, bottom?, left? } where each side
 * is either undefined (no border) or { style, color }.
 * Accepted inputs:
 *   true / 'all'                 → all 4 sides, thin black
 *   false / null                 → no borders
 *   { top, right, bottom, left } → per-side; each value: true|false|{style,color}
 *   { all: ... }                 → shorthand for all 4 sides
 */
function normalizeBorder(input) {
  if (input === false || input === null) return {};
  const defaultSide = { style: 'thin', color: '#000000' };
  const normSide = (v) => {
    if (v === false || v === null || v === undefined) return undefined;
    if (v === true) return { ...defaultSide };
    if (typeof v === 'string') {
      if (!BORDER_STYLES.has(v)) throw err('INVALID_BORDER_STYLE', `border style は ${[...BORDER_STYLES].join('|')} のいずれか. 受け取った値: "${v}"`);
      return { style: v, color: '#000000' };
    }
    if (typeof v === 'object') {
      const style = v.style || 'thin';
      if (!BORDER_STYLES.has(style)) throw err('INVALID_BORDER_STYLE', `border style 不正: "${style}"`);
      return { style, color: v.color || '#000000' };
    }
    throw err('INVALID_BORDER', `border の値が不正: ${JSON.stringify(v)}`);
  };
  if (input === true || input === 'all') {
    return { top: { ...defaultSide }, right: { ...defaultSide }, bottom: { ...defaultSide }, left: { ...defaultSide } };
  }
  if (typeof input === 'object') {
    if (input.all !== undefined) {
      const side = normSide(input.all);
      return side ? { top: side, right: { ...side }, bottom: { ...side }, left: { ...side } } : {};
    }
    const out = {};
    for (const k of Object.keys(input)) {
      if (!BORDER_SIDES.includes(k)) throw err('INVALID_BORDER_SIDE', `border キーは top|right|bottom|left|all のみ. 受け取った値: "${k}"`);
      const s = normSide(input[k]);
      if (s) out[k] = s;
    }
    return out;
  }
  throw err('INVALID_BORDER', `border の値が不正: ${JSON.stringify(input)}`);
}

function validateStyle(s) {
  if (!s || typeof s !== 'object') throw err('INVALID_STYLE', `style はオブジェクトで指定. 例: { bold: true, color: "#ff0000" }`);
  for (const k of Object.keys(s)) {
    if (!STYLE_KEYS.has(k)) {
      throw err('INVALID_STYLE_KEY',
        `未知のスタイルキー: "${k}". 利用可能: ${[...STYLE_KEYS].join(', ')}`);
    }
  }
  if (s.align !== undefined && !ALIGN_VALUES.has(s.align)) {
    throw err('INVALID_ALIGN', `align は "left" | "center" | "right" のみ. 受け取った値: "${s.align}"`);
  }
  if (s.fontSize !== undefined && (typeof s.fontSize !== 'number' || s.fontSize <= 0)) {
    throw err('INVALID_FONT_SIZE', `fontSize は正の数値 (px). 受け取った値: ${s.fontSize}`);
  }
}

// ============================================================
// Date helpers
// ============================================================
// Excel "serial date": days since 1899-12-30 (local midnight).
// We don't model Excel's 1900 leap-year bug; for any date >= 1900-03-01 it's exact.
const _XL_EPOCH = new Date(1899, 11, 30);

function _isDate(v) { return v instanceof Date && !isNaN(v.getTime()); }

/** Convert a Date (local midnight) to Excel serial (number). */
function dateToSerial(d) {
  return Math.round((d - _XL_EPOCH) / 86400000);
}
/** Convert an Excel serial back to a Date (local midnight). */
function serialToDate(n) {
  return new Date(_XL_EPOCH.getTime() + n * 86400000);
}
/** Parse a "YYYY-MM-DD" or ISO date string into a local-midnight Date. */
function _parseDateString(s) {
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(+m[1], +m[2] - 1, +m[3]);
}
/** Format a Date with an Excel-like numFmt (subset: yyyy/yy/mm/m/dd/d, case-insensitive). */
function _formatDate(d, fmt) {
  if (!_isDate(d)) return String(d ?? '');
  const y = d.getFullYear(), M = d.getMonth() + 1, D = d.getDate();
  const pad = n => String(n).padStart(2, '0');
  return fmt.replace(/yyyy|yy|mm|m|dd|d/gi, t => ({
    yyyy: String(y), yy: String(y).slice(-2),
    mm: pad(M), m: String(M),
    dd: pad(D), d: String(D)
  }[t.toLowerCase()]));
}
/** Format a number with an Excel-like numFmt (subset). */
function _formatNumber(n, fmt) {
  if (typeof n !== 'number' || !isFinite(n)) return String(n);
  // Percent
  if (fmt.includes('%')) {
    const decimals = (fmt.match(/0\.(0+)/) || ['', ''])[1].length;
    return (n * 100).toFixed(decimals) + '%';
  }
  // Decimal count from format
  const decM = fmt.match(/0\.(0+)/);
  const decimals = decM ? decM[1].length : 0;
  const useSep = fmt.includes('#,##0') || fmt.includes(',');
  const prefixM = fmt.match(/^([^#0]+)/);
  const prefix = prefixM ? prefixM[1] : '';
  const suffixM = fmt.match(/([^#0.,]+)$/);
  const suffix = suffixM ? suffixM[1] : '';
  let s = n.toFixed(decimals);
  if (useSep) {
    const [intPart, fracPart] = s.split('.');
    s = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (fracPart ? '.' + fracPart : '');
  }
  return prefix + s + suffix;
}

/** Display string for a cell value (applies numFmt if present). */
function formatCellValue(v, numFmt) {
  // If formula yielded a numeric serial but cell has a date numFmt, coerce back to Date.
  if (typeof v === 'number' && numFmt && /[ymd]/i.test(numFmt) && v > 0 && v < 200000) {
    v = serialToDate(v);
  }
  if (_isDate(v)) return numFmt ? _formatDate(v, numFmt) : _formatDate(v, 'yyyy/m/d');
  if (typeof v === 'number' && numFmt) return _formatNumber(v, numFmt);
  return v === null || v === undefined ? '' : String(v);
}

// ============================================================
// XLSX helpers (shared by Sheet.toXLSX / Workbook.toXLSX)
// ============================================================
/**
 * Load a SheetJS-compatible library. Prefers xlsx-js-style (community fork
 * that writes cell styles); falls back to plain xlsx if only that is present.
 * Browser: looks at window.XLSX (the lazy-loader in viewer puts xlsx-js-style there).
 */
async function _loadXLSX() {
  if (typeof window !== 'undefined' && window.XLSX) return window.XLSX;
  if (typeof process === 'undefined' || !process.versions?.node) return null;
  for (const name of ['xlsx-js-style', 'xlsx']) {
    try {
      const mod = await import(name);
      return mod.default ?? mod;
    } catch {}
  }
  return null;
}

const _BORDER_STYLE_MAP = {
  thin: 'thin', medium: 'medium', thick: 'thick',
  dotted: 'dotted', dashed: 'dashed', double: 'double'
};
function _toXlsxBorderSide(side) {
  if (!side) return undefined;
  return { style: _BORDER_STYLE_MAP[side.style] || 'thin',
           color: { rgb: (side.color || '#000000').replace('#', '') } };
}
function _writeAOA(sheet, usedRange) {
  // Build aoa with Date → Date (SheetJS handles when cellDates), but for safety we
  // also tag those cells explicitly after creation via _patchDateCells().
  const aoa = [];
  for (let r = 0; r < usedRange.rows; r++) {
    const row = [];
    for (let c = 0; c < usedRange.cols; c++) {
      const ref = makeRef(r, c);
      const cell = sheet.cells[ref];
      if (!cell) { row.push(null); continue; }
      if (cell.f) row.push({ f: cell.f.slice(1) });
      else if (_isDate(cell.v)) row.push(cell.v);
      else row.push(cell.v ?? null);
    }
    aoa.push(row);
  }
  return aoa;
}

function _patchDateCells(sheet, ws) {
  // SheetJS aoa_to_sheet writes Date as ISO string by default. Replace those with
  // numeric Excel serial + t:'n' + a numFmt (z) so Excel renders as date.
  for (const [ref, cell] of Object.entries(sheet.cells)) {
    if (!_isDate(cell.v) || !ws[ref]) continue;
    const fmt = cell.s?.numFmt || 'yyyy/m/d';
    ws[ref].t = 'n';
    ws[ref].v = dateToSerial(cell.v);
    ws[ref].z = fmt;
  }
}

function _applySheetMetaToWS(sheet, ws, usedRange) {
  // Column widths
  if (Object.keys(sheet.cols).length) {
    ws['!cols'] = [];
    for (let c = 0; c < usedRange.cols; c++) {
      const w = sheet.cols[idxToCol(c)];
      ws['!cols'].push(w ? { wpx: w } : {});
    }
  }
  // Row heights
  if (Object.keys(sheet.rows).length) {
    ws['!rows'] = [];
    for (let r = 0; r < usedRange.rows; r++) {
      const h = sheet.rows[r + 1];
      ws['!rows'].push(h ? { hpx: h } : {});
    }
  }
  // Cell styles
  for (const [ref, cell] of Object.entries(sheet.cells)) {
    if (!cell.s) continue;
    if (!ws[ref]) ws[ref] = { t: 's', v: '' };  // create stub so style attaches
    ws[ref].s = ws[ref].s || {};
    if (cell.s.bold)      ws[ref].s.font = { ...(ws[ref].s.font || {}), bold: true };
    if (cell.s.italic)    ws[ref].s.font = { ...(ws[ref].s.font || {}), italic: true };
    if (cell.s.underline) ws[ref].s.font = { ...(ws[ref].s.font || {}), underline: true };
    if (cell.s.color)     ws[ref].s.font = { ...(ws[ref].s.font || {}), color: { rgb: cell.s.color.replace('#','') } };
    if (cell.s.fontSize)  ws[ref].s.font = { ...(ws[ref].s.font || {}), sz: cell.s.fontSize };
    if (cell.s.fontFamily) ws[ref].s.font = { ...(ws[ref].s.font || {}), name: cell.s.fontFamily };
    if (cell.s.bgColor)   ws[ref].s.fill = { fgColor: { rgb: cell.s.bgColor.replace('#','') }, patternType: 'solid' };
    if (cell.s.align)     ws[ref].s.alignment = { horizontal: cell.s.align };
    if (cell.s.border) {
      const b = {};
      for (const side of ['top', 'right', 'bottom', 'left']) {
        const x = _toXlsxBorderSide(cell.s.border[side]);
        if (x) b[side] = x;
      }
      if (Object.keys(b).length) ws[ref].s.border = b;
    }
    if (cell.s.numFmt) ws[ref].z = cell.s.numFmt;
  }
  // Merges
  if (sheet.merges.length) {
    ws['!merges'] = sheet.merges.map(m => ({ s: { r: m.r1, c: m.c1 }, e: { r: m.r2, c: m.c2 } }));
  }
}

// ============================================================
// Sheet class
// ============================================================
class Sheet {
  constructor(name = 'Sheet1') {
    this.name = name;
    this.cells = {};        // { "A1": { v: ..., f: ..., s: {...} } }
    this.cols = {};         // { "A": 120 }  width in px
    this.rows = {};         // { 1: 24 }     height in px
    this.images = [];       // [{ id, src, anchor, offset, size }]
    this.merges = [];       // [{ r1, c1, r2, c2 }] 0-based inclusive
    this.cfs = [];          // [{ range: "G3:Z11", formula: "=...", style: {bgColor, color, bold, italic} }]
    this.version = '1.0';
    this._imgCounter = 0;
  }

  // ---- Static loaders ----

  /** Parse from a .aix.json string or object */
  static fromJSON(input) {
    const data = typeof input === 'string' ? JSON.parse(input) : input;
    if (!data || data.type !== 'aix-sheet') {
      throw err('INVALID_FORMAT', 'aix-sheet 形式ではありません (type !== "aix-sheet")');
    }
    const s = new Sheet(data.name || 'Sheet1');
    s.cells = {};
    for (const [k, c] of Object.entries(data.cells || {})) {
      if (c && c.t === 'd' && typeof c.v === 'string') {
        const d = _parseDateString(c.v);
        s.cells[k] = { ...c, v: d };
        delete s.cells[k].t;
      } else {
        s.cells[k] = c;
      }
    }
    s.cols  = data.cols  || {};
    s.rows  = data.rows  || {};
    s.images = Array.isArray(data.images) ? data.images : [];
    s._imgCounter = s.images.length;
    s.merges = Array.isArray(data.merges)
      ? data.merges.map(m => (typeof m === 'string' ? Sheet._parseRangeToMerge(m) : { r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 }))
      : [];
    s.cfs = Array.isArray(data.cfs) ? data.cfs.map(r => ({ ...r, style: { ...r.style } })) : [];
    s.version = data.version || '1.0';
    return s;
  }

  static _parseRangeToMerge(ref) {
    const r = parseRange(ref);
    return { r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c2 };
  }

  /** Load from a file path (Node.js only). Returns a Promise. */
  static async load(path) {
    if (typeof process === 'undefined' || !process.versions?.node) {
      throw err('NODE_ONLY', 'Sheet.load() は Node.js でのみ利用可能');
    }
    const fs = await import('node:fs/promises');
    return Sheet.fromJSON(await fs.readFile(path, 'utf8'));
  }

  // ---- Write ----

  /**
   * Write a value to a cell or range.
   *   write("A1", 42)
   *   write("A1", "=SUM(B1:B10)")
   *   write("A1:A5", [1,2,3,4,5])         // 1D → column-direction
   *   write("A1:C1", [1,2,3])              // 1D → row-direction
   *   write("A1:C2", [[1,2,3],[4,5,6]])    // 2D
   */
  write(ref, value) {
    const range = parseRange(ref);
    const rows = range.r2 - range.r1 + 1;
    const cols = range.c2 - range.c1 + 1;

    if (rows === 1 && cols === 1) {
      this._setCell(range.r1, range.c1, value);
      return this;
    }

    if (!Array.isArray(value)) {
      throw err('INVALID_VALUE', `範囲書き込みには配列が必要. 範囲: "${ref}", 受け取った値: ${typeof value}`);
    }

    if (Array.isArray(value[0])) {
      // 2D
      if (value.length !== rows || value[0].length !== cols) {
        throw err('SHAPE_MISMATCH',
          `配列の形 ${value.length}×${value[0].length} が範囲 ${rows}×${cols} と一致しません`);
      }
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          this._setCell(range.r1 + r, range.c1 + c, value[r][c]);
        }
      }
    } else {
      // 1D
      if (rows === 1) {
        if (value.length !== cols) throw err('SHAPE_MISMATCH', `配列長 ${value.length} が列数 ${cols} と一致しません`);
        for (let c = 0; c < cols; c++) this._setCell(range.r1, range.c1 + c, value[c]);
      } else if (cols === 1) {
        if (value.length !== rows) throw err('SHAPE_MISMATCH', `配列長 ${value.length} が行数 ${rows} と一致しません`);
        for (let r = 0; r < rows; r++) this._setCell(range.r1 + r, range.c1, value[r]);
      } else {
        throw err('SHAPE_MISMATCH', `2次元範囲 ${ref} には2次元配列を渡してください`);
      }
    }
    return this;
  }

  _setCell(r, c, v) {
    const ref = makeRef(r, c);
    if (v === null || v === undefined || v === '') {
      if (this.cells[ref]) { delete this.cells[ref].v; delete this.cells[ref].f;
        if (!this.cells[ref].s) delete this.cells[ref]; }
      return;
    }
    if (!this.cells[ref]) this.cells[ref] = {};
    if (typeof v === 'string' && v.startsWith('=')) {
      this.cells[ref].f = v;
      delete this.cells[ref].v;
    } else {
      this.cells[ref].v = v;
      delete this.cells[ref].f;
      // Auto-apply a default numFmt if storing a Date and the cell has none yet
      if (_isDate(v)) {
        if (!this.cells[ref].s) this.cells[ref].s = {};
        if (!this.cells[ref].s.numFmt) this.cells[ref].s.numFmt = 'yyyy/m/d';
      }
    }
  }

  // ---- Read ----

  /** Return raw cell record: { v, f, s }. Empty cell returns {}. */
  get(ref) {
    const p = parseRef(ref);
    return this.cells[p.label] ? { ...this.cells[p.label] } : {};
  }

  /** Computed value (evaluates formula). For range, returns 2D array. */
  value(ref) {
    if (ref.includes(':')) {
      const range = parseRange(ref);
      const out = [];
      for (let r = range.r1; r <= range.r2; r++) {
        const row = [];
        for (let c = range.c1; c <= range.c2; c++) row.push(this._eval(r, c));
        out.push(row);
      }
      return out;
    }
    const p = parseRef(ref);
    return this._eval(p.row, p.col);
  }

  /** Alias for value(). */
  read(ref) { return this.value(ref); }

  // ---- Style ----

  /**
   * Apply style to cell or range. Merges with existing style.
   *   style("A1", { bold: true })
   *   style("A1:D1", { bold: true, bgColor: "#eee" })
   */
  style(ref, opts) {
    validateStyle(opts);
    const range = parseRange(ref);
    const normalizedBorder = (opts.border !== undefined) ? normalizeBorder(opts.border) : undefined;
    for (let r = range.r1; r <= range.r2; r++) {
      for (let c = range.c1; c <= range.c2; c++) {
        const k = makeRef(r, c);
        if (!this.cells[k]) this.cells[k] = {};
        if (!this.cells[k].s) this.cells[k].s = {};
        const toApply = { ...opts };
        if (normalizedBorder !== undefined) {
          if (Object.keys(normalizedBorder).length === 0) {
            delete this.cells[k].s.border;
            delete toApply.border;
          } else {
            toApply.border = { ...(this.cells[k].s.border || {}), ...normalizedBorder };
          }
        }
        Object.assign(this.cells[k].s, toApply);
      }
    }
    return this;
  }

  // ---- Merge ----

  /**
   * Merge a range of cells. Only the top-left cell keeps its value;
   * other cells become "covered" and are hidden in the viewer.
   * Throws if the range overlaps an existing merge.
   *   merge("A1:D1")
   *   merge("B2:C3")
   */
  merge(ref) {
    const r = parseRange(ref);
    if (r.r1 === r.r2 && r.c1 === r.c2) {
      throw err('INVALID_MERGE', `merge には2セル以上の範囲が必要: "${ref}"`);
    }
    for (const m of this.merges) {
      const overlap = !(r.r2 < m.r1 || r.r1 > m.r2 || r.c2 < m.c1 || r.c1 > m.c2);
      if (overlap) throw err('MERGE_OVERLAP', `範囲 "${ref}" は既存のマージと重なります`);
    }
    this.merges.push({ r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c2 });
    return this;
  }

  /**
   * Remove a merge containing the given cell or matching the given range.
   *   unmerge("A1")     // remove merge that covers A1
   *   unmerge("A1:D1")  // remove merge that exactly matches
   */
  unmerge(ref) {
    const r = parseRange(ref);
    this.merges = this.merges.filter(m =>
      !(m.r1 === r.r1 && m.c1 === r.c1 && m.r2 === r.r2 && m.c2 === r.c2) &&
      !(r.r1 === r.r2 && r.c1 === r.c2 && r.r1 >= m.r1 && r.r1 <= m.r2 && r.c1 >= m.c1 && r.c1 <= m.c2)
    );
    return this;
  }

  // ---- Fill (Excel-style autofill) ----

  /**
   * Copy a source cell/range to a target range. Formulas have their relative
   * refs shifted; absolute refs ($A$1, $A1, A$1) stay locked. Dates increment
   * by one day per row when the source is a single cell. Style is copied.
   *
   *   sheet.fill("A1", "A2:A10")   // copy A1 down
   *   sheet.fill("B1", "B2:B5")    // formula B1=A1+10 → B2=A2+10, B3=A3+10, ...
   */
  fill(srcRef, targetRef) {
    const src = parseRange(srcRef);
    const tgt = parseRange(targetRef);
    const srcRows = src.r2 - src.r1 + 1;
    const srcCols = src.c2 - src.c1 + 1;
    const series = _detectSeries(this, src);
    const isVertical = srcCols === 1 && srcRows > 1;
    for (let r = tgt.r1; r <= tgt.r2; r++) {
      for (let c = tgt.c1; c <= tgt.c2; c++) {
        if (r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2) continue;
        const sr = src.r1 + ((r - src.r1) % srcRows + srcRows) % srcRows;
        const sc = src.c1 + ((c - src.c1) % srcCols + srcCols) % srcCols;
        const srcCell = this.cells[makeRef(sr, sc)];
        let v;
        if (series) {
          const axisIdx = isVertical ? (r - src.r1) : (c - src.c1);
          if (series.type === 'number') v = series.start + axisIdx * series.step;
          else v = serialToDate(dateToSerial(series.start) + axisIdx * series.step);
        } else if (srcCell?.f) {
          v = _shiftFormula(srcCell.f, r - sr, c - sc);
        } else if (srcCell && _isDate(srcCell.v) && srcRows * srcCols === 1) {
          const d = new Date(srcCell.v);
          d.setDate(d.getDate() + (r - sr));
          v = d;
        } else if (srcCell) {
          v = srcCell.v;
        } else continue;
        this._setCell(r, c, v);
        if (srcCell?.s) {
          const k = makeRef(r, c);
          if (!this.cells[k]) this.cells[k] = {};
          this.cells[k].s = { ...srcCell.s };
        }
      }
    }
    // Extend any CF rule that overlaps the source range, in the fill direction.
    // (We use "overlap" rather than "fully contains" so that a column drag across
    //  multiple per-row CF rules still extends each of them.)
    const verticalFill   = tgt.r2 > src.r2;
    const horizontalFill = tgt.c2 > src.c2;
    if (verticalFill || horizontalFill) {
      for (const rule of this.cfs) {
        const rr = parseRange(rule.range);
        const overlap = !(src.r2 < rr.r1 || src.r1 > rr.r2 || src.c2 < rr.c1 || src.c1 > rr.c2);
        if (!overlap) continue;
        const newR2 = verticalFill   ? Math.max(rr.r2, tgt.r2) : rr.r2;
        const newC2 = horizontalFill ? Math.max(rr.c2, tgt.c2) : rr.c2;
        rule.range = makeRef(rr.r1, rr.c1) + ':' + makeRef(newR2, newC2);
      }
    }
    return this;
  }

  /** Return the merge record covering (r,c) or null. */
  mergeAt(r, c) {
    return this.merges.find(m => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) || null;
  }

  // ---- Conditional formatting ----

  /**
   * Add a conditional-formatting rule. The formula is Excel-style (must start with =).
   * Relative refs (e.g. D3) shift with the cell being evaluated; absolute refs ($D$3, $D3, D$3)
   * lock the column and/or row. ROW() / COLUMN() return the 1-based row/col of the current cell.
   * Style overrides apply on top of the cell's own style.
   *
   *   sheet.cf("G3:Z11", {
   *     formula: "=AND(COLUMN()-7>=$D3, COLUMN()-7<=$E3)",
   *     style:   { bgColor: "#4472c4", color: "#ffffff", bold: true }
   *   });
   */
  cf(range, opts) {
    parseRange(range); // validate
    if (!opts || typeof opts.formula !== 'string' || !opts.formula.startsWith('=')) {
      throw err('INVALID_CF', `cf には { formula: "=...", style: {...} } を指定してください`);
    }
    if (!opts.style || typeof opts.style !== 'object') {
      throw err('INVALID_CF', `cf.style は { bgColor, color, bold, italic } のような style オブジェクトで指定してください`);
    }
    this.cfs.push({ range, formula: opts.formula, style: { ...opts.style } });
    return this;
  }

  /** Remove a CF rule by exact range match, or clear all if no range given. */
  clearCF(range) {
    if (range === undefined) this.cfs = [];
    else this.cfs = this.cfs.filter(c => c.range !== range);
    return this;
  }

  /** Evaluate a CF formula in the context of (row, col). Returns boolean. */
  _evalCF(formula, row, col, range) {
    try {
      const translated = _translateCFFormula(formula.slice(1).trim(), {
        row, col, rangeR1: range.r1, rangeC1: range.c1
      });
      const v = this._evalExpr(translated, new Set());
      return v === true || v === 'TRUE' || (typeof v === 'number' && v !== 0);
    } catch { return false; }
  }

  /** Return merged style overrides from any matching CF rules for (row, col). */
  cfStyleAt(row, col) {
    let merged = null;
    for (const rule of this.cfs) {
      const r = parseRange(rule.range);
      if (row < r.r1 || row > r.r2 || col < r.c1 || col > r.c2) continue;
      if (this._evalCF(rule.formula, row, col, r)) {
        merged = { ...(merged || {}), ...rule.style };
      }
    }
    return merged;
  }

  // ---- Clear / Delete ----

  /** Clear values & styles in a range. */
  clear(ref) {
    const range = parseRange(ref);
    for (let r = range.r1; r <= range.r2; r++) {
      for (let c = range.c1; c <= range.c2; c++) delete this.cells[makeRef(r, c)];
    }
    return this;
  }

  // ---- Images ----

  /**
   * Add an image to the sheet.
   *
   * @param {string} anchor - top-left cell ref like "B2"
   * @param {string} source - data URL ("data:image/png;base64,...") or HTTP(S) URL
   * @param {object} [options]
   * @param {{x:number,y:number}} [options.offset] - px offset within the anchor cell (default {x:0,y:0})
   * @param {{width:number,height:number}} [options.size] - display size in px (default {width:200,height:150})
   * @param {string} [options.id] - custom id; auto-generated if omitted
   * @returns {string} the image id
   */
  addImage(anchor, source, options = {}) {
    parseRef(anchor); // validate
    if (typeof source !== 'string' || source.length === 0) {
      throw err('INVALID_IMAGE_SRC', '画像ソースは data URL または http(s) URL の文字列で指定してください');
    }
    const id = options.id || `img_${++this._imgCounter}`;
    if (this.images.find(im => im.id === id)) {
      throw err('DUPLICATE_IMAGE_ID', `画像ID "${id}" は既に存在します`);
    }
    this.images.push({
      id,
      src: source,
      anchor,
      offset: { x: options.offset?.x ?? 0, y: options.offset?.y ?? 0 },
      size:   { width:  options.size?.width  ?? 200,
                height: options.size?.height ?? 150 }
    });
    return id;
  }

  /** Get an image by id. */
  getImage(id) { return this.images.find(im => im.id === id) || null; }

  /** Remove an image by id. */
  removeImage(id) {
    const i = this.images.findIndex(im => im.id === id);
    if (i >= 0) this.images.splice(i, 1);
    return this;
  }

  /** Move an image to a new anchor cell, optionally with offset. */
  moveImage(id, newAnchor, offset) {
    const im = this.getImage(id);
    if (!im) throw err('IMAGE_NOT_FOUND', `画像が見つかりません: ${id}`);
    parseRef(newAnchor);
    im.anchor = newAnchor;
    if (offset) im.offset = { x: offset.x ?? 0, y: offset.y ?? 0 };
    return this;
  }

  /** Resize an image. */
  resizeImage(id, size) {
    const im = this.getImage(id);
    if (!im) throw err('IMAGE_NOT_FOUND', `画像が見つかりません: ${id}`);
    if (size.width  > 0) im.size.width  = size.width;
    if (size.height > 0) im.size.height = size.height;
    return this;
  }

  // ---- Row / Column sizing ----

  /** Set column width in px. col can be "A" or 0-based index. */
  colWidth(col, px) {
    const label = typeof col === 'number' ? idxToCol(col) : col.toUpperCase();
    this.cols[label] = px;
    return this;
  }
  /** Set row height in px. row is 1-based. */
  rowHeight(row, px) { this.rows[row] = px; return this; }

  // ---- Insert / Delete rows & columns ----

  /** Insert `count` blank rows above the given 1-based row number. */
  insertRow(rowNum, count = 1) {
    const r0 = rowNum - 1;
    this._remapCells((r, c) => ({ r: r >= r0 ? r + count : r, c }));
    this._remapRows((n) => n >= rowNum ? n + count : n);
    this._remapMergesRow(r0, count, false);
    return this;
  }
  /** Delete `count` rows starting at the given 1-based row number. */
  deleteRow(rowNum, count = 1) {
    const r0 = rowNum - 1, r1 = r0 + count - 1;
    this._remapCells((r, c) => {
      if (r >= r0 && r <= r1) return null;
      if (r > r1) return { r: r - count, c };
      return { r, c };
    });
    this._remapRows((n) => {
      if (n >= rowNum && n < rowNum + count) return null;
      if (n >= rowNum + count) return n - count;
      return n;
    });
    this._remapMergesRow(r0, count, true);
    return this;
  }
  /** Insert `count` blank columns to the left of the given column label. */
  insertCol(colLabel, count = 1) {
    const c0 = colToIdx(typeof colLabel === 'number' ? idxToCol(colLabel) : colLabel);
    this._remapCells((r, c) => ({ r, c: c >= c0 ? c + count : c }));
    this._remapCols((label) => {
      const i = colToIdx(label);
      return i >= c0 ? idxToCol(i + count) : label;
    });
    this._remapMergesCol(c0, count, false);
    return this;
  }
  /** Delete `count` columns starting at the given column label. */
  deleteCol(colLabel, count = 1) {
    const c0 = colToIdx(typeof colLabel === 'number' ? idxToCol(colLabel) : colLabel);
    const c1 = c0 + count - 1;
    this._remapCells((r, c) => {
      if (c >= c0 && c <= c1) return null;
      if (c > c1) return { r, c: c - count };
      return { r, c };
    });
    this._remapCols((label) => {
      const i = colToIdx(label);
      if (i >= c0 && i <= c1) return null;
      if (i > c1) return idxToCol(i - count);
      return label;
    });
    this._remapMergesCol(c0, count, true);
    return this;
  }

  _remapMergesRow(r0, count, isDelete) {
    const next = [];
    for (const m of this.merges) {
      let { r1, r2, c1, c2 } = m;
      if (isDelete) {
        const delEnd = r0 + count - 1;
        // Drop merges entirely inside the deleted range
        if (r1 >= r0 && r2 <= delEnd) continue;
        if (r1 > delEnd) { r1 -= count; r2 -= count; }
        else if (r1 >= r0) { r1 = r0; r2 = Math.max(r0, r2 - count); }
        else if (r2 >= r0) { r2 = Math.max(r1, r2 - Math.min(count, r2 - r0 + 1)); }
      } else {
        if (r1 >= r0) { r1 += count; r2 += count; }
      }
      if (r1 < r2 || c1 < c2) next.push({ r1, c1, r2, c2 });
    }
    this.merges = next;
  }
  _remapMergesCol(c0, count, isDelete) {
    const next = [];
    for (const m of this.merges) {
      let { r1, r2, c1, c2 } = m;
      if (isDelete) {
        const delEnd = c0 + count - 1;
        if (c1 >= c0 && c2 <= delEnd) continue;
        if (c1 > delEnd) { c1 -= count; c2 -= count; }
        else if (c1 >= c0) { c1 = c0; c2 = Math.max(c0, c2 - count); }
        else if (c2 >= c0) { c2 = Math.max(c1, c2 - Math.min(count, c2 - c0 + 1)); }
      } else {
        if (c1 >= c0) { c1 += count; c2 += count; }
      }
      if (r1 < r2 || c1 < c2) next.push({ r1, c1, r2, c2 });
    }
    this.merges = next;
  }

  _remapCells(fn) {
    const next = {};
    for (const [ref, data] of Object.entries(this.cells)) {
      const p = parseRef(ref);
      const m = fn(p.row, p.col);
      if (m === null) continue;
      next[makeRef(m.r, m.c)] = data;
    }
    this.cells = next;
  }
  _remapRows(fn) {
    const next = {};
    for (const [k, v] of Object.entries(this.rows)) {
      const m = fn(Number(k));
      if (m !== null) next[m] = v;
    }
    this.rows = next;
  }
  _remapCols(fn) {
    const next = {};
    for (const [k, v] of Object.entries(this.cols)) {
      const m = fn(k);
      if (m !== null) next[m] = v;
    }
    this.cols = next;
  }

  // ---- Inspection (LLM-friendly) ----

  /** Bounding box of used cells: { rows, cols }. Empty sheet = {rows:0, cols:0}. */
  usedRange() {
    let maxR = -1, maxC = -1;
    for (const ref of Object.keys(this.cells)) {
      const p = parseRef(ref);
      if (p.row > maxR) maxR = p.row;
      if (p.col > maxC) maxC = p.col;
    }
    return { rows: maxR + 1, cols: maxC + 1 };
  }

  /** Render used range as a Markdown table (computed values). For LLM context. */
  /**
   * Render the sheet as Markdown for LLM context.
   *
   * Default output includes:
   *   1. data table of computed values
   *   2. **Formulas:** list of every cell with an `f` (so editors don't
   *      accidentally overwrite a derived cell)
   *   3. **Merges:** list of merged ranges (if any)
   *   4. **Conditional formatting:** list of CF rules (if any)
   *
   * Options:
   *   maxRows: cap the data table (default 50)
   *   meta:    false to omit all three meta sections (legacy data-only output)
   */
  toMarkdown(opts = {}) {
    const limit = opts.maxRows ?? 50;
    const u = this.usedRange();
    if (u.rows === 0) return '*(空のシート)*';
    const rows = Math.min(u.rows, limit);
    const header = ['', ...Array.from({ length: u.cols }, (_, c) => idxToCol(c))];
    const sep    = header.map(() => '---');
    const lines = ['| ' + header.join(' | ') + ' |', '| ' + sep.join(' | ') + ' |'];
    for (let r = 0; r < rows; r++) {
      const row = [String(r + 1)];
      for (let c = 0; c < u.cols; c++) row.push(String(this._eval(r, c)));
      lines.push('| ' + row.join(' | ') + ' |');
    }
    if (u.rows > limit) lines.push(`\n*...(${u.rows - limit} 行省略)*`);

    if (opts.meta !== false) {
      // Sort formula cells by row then column for stable, readable output
      const formulaCells = Object.entries(this.cells)
        .filter(([, c]) => c.f)
        .map(([ref, c]) => ({ ref, f: c.f, p: parseRef(ref) }))
        .sort((a, b) => a.p.row - b.p.row || a.p.col - b.p.col);
      if (formulaCells.length) {
        lines.push('', '**Formulas:**');
        for (const { ref, f } of formulaCells) lines.push(`- \`${ref}\` = \`${f}\``);
      }
      if (this.merges.length) {
        lines.push('', '**Merges:**');
        for (const m of this.merges) {
          lines.push(`- \`${makeRef(m.r1, m.c1)}:${makeRef(m.r2, m.c2)}\``);
        }
      }
      if (this.cfs.length) {
        const styleHint = (s) => Object.entries(s || {}).map(([k, v]) => `${k}=${v}`).join(', ');
        // When there's a manageable number of rules, list every one.
        // Otherwise group by style (most Gantt-style setups create many
        // per-cell rules that share the same look) and show a sample formula.
        const compact = this.cfs.length > 30;
        if (!compact) {
          lines.push('', '**Conditional formatting:**');
          for (const rule of this.cfs) {
            lines.push(`- \`${rule.range}\` when \`${rule.formula}\` → ${styleHint(rule.style)}`);
          }
        } else {
          lines.push('', `**Conditional formatting:** (${this.cfs.length} rules, grouped by style)`);
          const byStyle = new Map();
          for (const rule of this.cfs) {
            const key = JSON.stringify(rule.style || {});
            if (!byStyle.has(key)) byStyle.set(key, { style: rule.style, rules: [] });
            byStyle.get(key).rules.push(rule);
          }
          for (const { style, rules } of byStyle.values()) {
            let minR = Infinity, minC = Infinity, maxR = -Infinity, maxC = -Infinity;
            for (const r of rules) {
              const p = parseRange(r.range);
              minR = Math.min(minR, p.r1); maxR = Math.max(maxR, p.r2);
              minC = Math.min(minC, p.c1); maxC = Math.max(maxC, p.c2);
            }
            const bbox = `${makeRef(minR, minC)}:${makeRef(maxR, maxC)}`;
            lines.push(`- ${rules.length} rules over \`${bbox}\` → ${styleHint(style)}`);
            lines.push(`  sample: \`${rules[0].range}\` when \`${rules[0].formula}\``);
          }
        }
      }
    }
    return lines.join('\n');
  }

  // ---- Serialization ----

  toJSON() {
    // Serialize Date values to "YYYY-MM-DD" strings with t:"d" tag
    const cellsOut = {};
    for (const [k, c] of Object.entries(this.cells)) {
      if (_isDate(c.v)) {
        const d = c.v;
        const iso = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
        cellsOut[k] = { ...c, v: iso, t: 'd' };
      } else {
        cellsOut[k] = c;
      }
    }
    const out = {
      version: this.version,
      type: 'aix-sheet',
      name: this.name,
      cells: cellsOut,
      cols: this.cols,
      rows: this.rows
    };
    if (this.images.length) out.images = this.images;
    if (this.merges.length) out.merges = this.merges.map(m => `${makeRef(m.r1, m.c1)}:${makeRef(m.r2, m.c2)}`);
    if (this.cfs.length) out.cfs = this.cfs;
    return out;
  }

  /** Save to a .aix.json file (Node.js only). Returns a Promise. */
  async save(path) {
    if (typeof process === 'undefined' || !process.versions?.node) {
      throw err('NODE_ONLY', 'save() は Node.js のみ. ブラウザでは toJSON() で取得して任意の手段で保存してください');
    }
    const fs = await import('node:fs/promises');
    await fs.writeFile(path, JSON.stringify(this.toJSON(), null, 2), 'utf8');
    return this;
  }

  // ---- CSV ----

  toCSV() {
    const u = this.usedRange();
    const out = [];
    for (let r = 0; r < u.rows; r++) {
      const cells = [];
      for (let c = 0; c < u.cols; c++) {
        const v = String(this._eval(r, c));
        cells.push(/[,\n"]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v);
      }
      out.push(cells.join(','));
    }
    return out.join('\n');
  }

  static fromCSV(text, name = 'Sheet1') {
    const s = new Sheet(name);
    const lines = text.split(/\r?\n/).filter(l => l.length > 0);
    lines.forEach((line, r) => {
      const cells = [];
      let cur = '', inQ = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (inQ) {
          if (ch === '"' && line[i+1] === '"') { cur += '"'; i++; }
          else if (ch === '"') inQ = false;
          else cur += ch;
        } else {
          if (ch === '"') inQ = true;
          else if (ch === ',') { cells.push(cur); cur = ''; }
          else cur += ch;
        }
      }
      cells.push(cur);
      cells.forEach((v, c) => {
        if (v === '') return;
        const num = Number(v);
        s._setCell(r, c, !isNaN(num) && v.trim() !== '' ? num : v);
      });
    });
    return s;
  }

  // ---- XLSX export (requires SheetJS to be available) ----

  /**
   * Convert to .xlsx Blob (browser) or Buffer (Node).
   * Requires SheetJS:
   *   Browser: <script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
   *   Node:    npm install xlsx
   */
  async toXLSX() {
    const XLSX = await _loadXLSX();
    if (!XLSX) throw err('XLSX_MISSING', 'SheetJS が見つかりません. ブラウザでは xlsx-js-style を CDN ロード、Node では `npm install xlsx-js-style` (推奨) または `npm install xlsx` を実行してください');

    const u = this.usedRange();
    const aoa = _writeAOA(this, u);
    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
    _patchDateCells(this, ws);
    _applySheetMetaToWS(this, ws, u);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, this.name);
    const base = XLSX.write(wb, { type: typeof window !== 'undefined' ? 'array' : 'buffer', bookType: 'xlsx' });

    let out = base;
    if (this.images.length) {
      const { injectImages } = await import('./xlsx-images.js');
      out = await injectImages(out, [{ name: this.name, images: this.images }]);
    }
    if (this.cfs.length) {
      const { injectCF } = await import('./xlsx-cf.js');
      out = await injectCF(out, [{ name: this.name, cfs: this.cfs }]);
    }
    return out;
  }

  // ============================================================
  // Formula evaluation
  // ============================================================
  _eval(r, c, visited = new Set()) {
    const key = `${this.name}::${r},${c}`;
    if (visited.has(key)) return '#CIRC!';
    const ref = makeRef(r, c);
    const cell = this.cells[ref];
    if (!cell) return '';
    if (cell.f === undefined) return cell.v ?? '';

    const next = new Set(visited); next.add(key);
    try { return this._evalExpr(cell.f.slice(1).trim(), next); }
    catch (e) { return '#ERR!'; }
  }

  /** Resolve a Sheet by name (cross-sheet refs). Returns null if not found. */
  _sheetByName(name) {
    if (!name) return this;
    return this._workbook ? this._workbook.sheet(name) : null;
  }

  _evalExpr(expr, visited) {
    expr = expr.trim();

    const agg = expr.match(/^(SUM|AVERAGE|MAX|MIN|COUNT|COUNTA)\s*\((.+)\)$/i);
    if (agg) {
      const fn = agg[1].toUpperCase();
      const vals = this._collectArgs(agg[2], visited);
      const nums = vals.map(Number).filter(n => !isNaN(n));
      if (fn === 'SUM')     return nums.reduce((a,b)=>a+b, 0);
      if (fn === 'AVERAGE') return nums.length ? nums.reduce((a,b)=>a+b,0)/nums.length : 0;
      if (fn === 'MAX')     return nums.length ? Math.max(...nums) : 0;
      if (fn === 'MIN')     return nums.length ? Math.min(...nums) : 0;
      if (fn === 'COUNT')   return nums.length;
      if (fn === 'COUNTA')  return vals.filter(v => v !== '').length;
    }

    const ifM = expr.match(/^IF\s*\((.+)\)$/i);
    if (ifM) {
      const parts = splitArgs(ifM[1]);
      const cond = this._evalExpr(parts[0], visited);
      const t = cond && cond !== 0 && cond !== 'FALSE' && cond !== false;
      return this._evalExpr(t ? parts[1] : (parts[2] || '""'), visited);
    }

    const andM = expr.match(/^AND\s*\((.+)\)$/i);
    if (andM) {
      for (const a of splitArgs(andM[1])) {
        const v = this._evalExpr(a, visited);
        if (!v || v === 'FALSE' || v === false || v === 0) return false;
      }
      return true;
    }
    const orM = expr.match(/^OR\s*\((.+)\)$/i);
    if (orM) {
      for (const a of splitArgs(orM[1])) {
        const v = this._evalExpr(a, visited);
        if (v && v !== 'FALSE' && v !== false && v !== 0) return true;
      }
      return false;
    }
    const notM = expr.match(/^NOT\s*\((.+)\)$/i);
    if (notM) {
      const v = this._evalExpr(notM[1], visited);
      return !v || v === 'FALSE' || v === false || v === 0;
    }

    const wkM = expr.match(/^WEEKDAY\s*\((.+)\)$/i);
    if (wkM) {
      const args = splitArgs(wkM[1]);
      const v = this._evalExpr(args[0], visited);
      const type = args[1] ? Number(this._evalExpr(args[1], visited)) : 1;
      const serial = _isDate(v) ? dateToSerial(v) : Number(v);
      if (!isFinite(serial)) return '#ERR!';
      const js = serialToDate(serial).getDay(); // 0=Sun..6=Sat
      if (type === 2) return js === 0 ? 7 : js;
      if (type === 3) return js === 0 ? 6 : js - 1;
      return js + 1;
    }

    // DATE() is also injected into the JS-eval context below so it works
    // when nested inside AND/OR/comparisons (e.g. "D2<=DATE(2026,7,7)").

    // Allow optional $ on column and/or row for absolute refs
    const REF_RE = /(?:'([^']+)'!|([A-Za-z_][\w]*)!)?(\$?[A-Za-z]+\$?\d+)/g;
    const resolved = expr.replace(REF_RE, (m, qSheet, uSheet, ref) => {
      try {
        const sheetName = qSheet || uSheet;
        const target = this._sheetByName(sheetName);
        if (!target) return sheetName ? '"#REF!"' : m;
        const cleanRef = ref.replace(/\$/g, '');
        const p = parseRef(cleanRef);
        const v = target._eval(p.row, p.col, visited);
        if (_isDate(v)) return dateToSerial(v);
        if (typeof v === 'number') return v;
        if (v === '') return '0';
        if (typeof v === 'string') {
          // Coerce date-looking strings to Excel serials so comparisons like
          // D2<=DATE(2026,7,7) work even when the cell stores the date as a
          // plain string. Accept both "2026-07-01" and "2026/7/1" (the latter
          // is what numFmt 'yyyy/m/d' renders, and what users / AIs commonly
          // pass to write()).
          const dm = v.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
          if (dm) return dateToSerial(new Date(+dm[1], +dm[2] - 1, +dm[3]));
        }
        return isNaN(Number(v)) ? JSON.stringify(String(v)) : Number(v);
      } catch { return m; }
    });
    // Inject a handful of Excel functions into the JS eval scope so they work
    // anywhere in the expression, not just as the top-level form.
    const DATE  = (y, m, d) => dateToSerial(new Date(Number(y), Number(m) - 1, Number(d)));
    const TODAY = () => { const n = new Date(); return dateToSerial(new Date(n.getFullYear(), n.getMonth(), n.getDate())); };
    const NOW   = TODAY;
    return Function('DATE', 'TODAY', 'NOW',
      '"use strict";return(' + _xlOpsToJs(resolved) + ')')(DATE, TODAY, NOW);
  }

  _collectArgs(s, visited) {
    const out = [];
    for (const a of splitArgs(s)) {
      const t = a.trim();
      // Strip optional sheet prefix
      const qm = t.match(/^(?:'([^']+)'|([A-Za-z_][\w]*))!(.+)$/);
      const sheetName = qm ? (qm[1] || qm[2]) : null;
      const refPart = qm ? qm[3] : t;
      const target = this._sheetByName(sheetName);
      if (sheetName && !target) { out.push('#REF!'); continue; }
      const toScalar = v => _isDate(v) ? dateToSerial(v) : v;
      if (refPart.includes(':')) {
        const r = parseRange(refPart);
        for (let rr = r.r1; rr <= r.r2; rr++)
          for (let cc = r.c1; cc <= r.c2; cc++) out.push(toScalar(target._eval(rr, cc, visited)));
      } else {
        try { const p = parseRef(refPart); out.push(toScalar(target._eval(p.row, p.col, visited))); }
        catch { const n = Number(refPart); if (!isNaN(n)) out.push(n); }
      }
    }
    return out;
  }
}

/**
 * Translate a CF formula for the cell at (ctx.row, ctx.col), given the rule's
 * range top-left at (ctx.rangeR1, ctx.rangeC1).
 *   - ROW() / COLUMN() → literal numbers (1-based)
 *   - Relative refs (D3) shift by the cell's offset from the range top-left
 *   - Absolute refs ($D3, D$3, $D$3) lock the column and/or row
 *   - Comparison ops: '=' → '==', '<>' → '!='
 */
function _translateCFFormula(formula, ctx) {
  formula = formula
    .replace(/\bROW\s*\(\s*\)/gi, ctx.row + 1)
    .replace(/\bCOLUMN\s*\(\s*\)/gi, ctx.col + 1);
  formula = formula.replace(/(\$?)([A-Za-z]+)(\$?)(\d+)/g, (m, dc, colLabel, dr, rowStr) => {
    let refCol, refRow;
    try { refCol = colToIdx(colLabel); } catch { return m; }
    refRow = parseInt(rowStr) - 1;
    const newCol = dc ? refCol : refCol + (ctx.col - ctx.rangeC1);
    const newRow = dr ? refRow : refRow + (ctx.row - ctx.rangeR1);
    return idxToCol(newCol) + (newRow + 1);
  });
  return formula;
}

/**
 * Shift relative refs in a formula by (dr, dc). Absolute parts ($) stay locked.
 * Used by Sheet.fill().
 */
function _shiftFormula(formula, dr, dc) {
  return formula.replace(/(\$?)([A-Za-z]+)(\$?)(\d+)/g, (m, dollarC, colLabel, dollarR, rowStr) => {
    let refCol, refRow;
    try { refCol = colToIdx(colLabel); } catch { return m; }
    refRow = parseInt(rowStr) - 1;
    if (!dollarC) refCol += dc;
    if (!dollarR) refRow += dr;
    if (refCol < 0 || refRow < 0) return m;
    return (dollarC ? '$' : '') + idxToCol(refCol) + (dollarR ? '$' : '') + (refRow + 1);
  });
}

/**
 * Detect an arithmetic series (numbers or dates) in a 1-D source range.
 * Returns { type: 'number'|'date', start, step } or null.
 * Requires >=2 cells, all numbers OR all Dates, with constant difference.
 */
function _detectSeries(sheet, src) {
  const srcRows = src.r2 - src.r1 + 1;
  const srcCols = src.c2 - src.c1 + 1;
  if (!((srcCols === 1 && srcRows > 1) || (srcRows === 1 && srcCols > 1))) return null;
  const len = Math.max(srcRows, srcCols);
  const isVertical = srcCols === 1;
  const vals = [];
  for (let i = 0; i < len; i++) {
    const r = isVertical ? src.r1 + i : src.r1;
    const c = isVertical ? src.c1 : src.c1 + i;
    const cell = sheet.cells[makeRef(r, c)];
    if (!cell || cell.f !== undefined) return null;
    vals.push(cell.v);
  }
  // Treat numeric strings as numbers (mixed "1" + 2 still counts as a series)
  const nums = vals.map(v => {
    if (typeof v === 'number') return v;
    if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) return Number(v);
    return NaN;
  });
  if (nums.every(v => !isNaN(v))) {
    const step = nums[1] - nums[0];
    for (let i = 2; i < nums.length; i++) {
      if (nums[i] - nums[i-1] !== step) return null;
    }
    return { type: 'number', start: nums[0], step };
  }
  if (vals.every(v => _isDate(v))) {
    const startSerial = dateToSerial(vals[0]);
    const step = dateToSerial(vals[1]) - startSerial;
    for (let i = 2; i < vals.length; i++) {
      if (dateToSerial(vals[i]) - dateToSerial(vals[i-1]) !== step) return null;
    }
    return { type: 'date', start: vals[0], step };
  }
  return null;
}

/** Replace Excel comparison ops with JS equivalents (= → ==, <> → !=). */
function _xlOpsToJs(expr) {
  // Order matters: '<>' first to avoid being broken into '<' + '>'.
  expr = expr.replace(/<>/g, '!=');
  // Replace single '=' that is not part of '==', '>=', '<=', '!=' with '=='.
  expr = expr.replace(/(^|[^=<>!])=(?!=)/g, '$1==');
  return expr;
}

function splitArgs(str) {
  const out = []; let depth = 0, cur = '', inStr = false;
  for (const ch of str) {
    if (ch === '"') { inStr = !inStr; cur += ch; continue; }
    if (inStr) { cur += ch; continue; }
    if (ch === '(') depth++; else if (ch === ')') depth--;
    if (ch === ',' && depth === 0) { out.push(cur.trim()); cur = ''; }
    else cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}

// ============================================================
// Workbook — container for multiple Sheets
// ============================================================
export class Workbook {
  constructor() {
    this.version = '1.0';
    this.sheets = [];            // Sheet[]
    this.activeIndex = 0;
  }

  /** Create from a .aix.json payload (workbook or single sheet). */
  static fromJSON(input) {
    const data = typeof input === 'string' ? JSON.parse(input) : input;
    if (!data) throw err('INVALID_FORMAT', '空のJSONです');
    const wb = new Workbook();

    if (data.type === 'aix-workbook') {
      wb.version = data.version || '1.0';
      wb.activeIndex = data.activeSheet ?? 0;
      for (const sd of (data.sheets || [])) {
        const s = Sheet.fromJSON({ ...sd, type: 'aix-sheet', version: '1.0' });
        wb._attach(s);
      }
      if (!wb.sheets.length) wb.addSheet('Sheet1');
      return wb;
    }

    if (data.type === 'aix-sheet') {
      const s = Sheet.fromJSON(data);
      wb._attach(s);
      return wb;
    }

    throw err('INVALID_FORMAT', `不明な type: "${data.type}" (期待: "aix-workbook" または "aix-sheet")`);
  }

  static async load(path) {
    if (typeof process === 'undefined' || !process.versions?.node) {
      throw err('NODE_ONLY', 'Workbook.load() は Node.js でのみ利用可能');
    }
    const fs = await import('node:fs/promises');
    return Workbook.fromJSON(await fs.readFile(path, 'utf8'));
  }

  /** Add a new empty sheet. Returns the Sheet instance. */
  addSheet(name) {
    name = name || this._uniqueName('Sheet');
    if (this.sheets.find(s => s.name === name)) {
      throw err('DUPLICATE_SHEET', `シート名 "${name}" は既に存在します`);
    }
    const s = new Sheet(name);
    this._attach(s);
    return s;
  }

  _attach(sheet) {
    sheet._workbook = this;
    this.sheets.push(sheet);
  }

  _uniqueName(prefix) {
    let n = this.sheets.length + 1;
    while (this.sheets.find(s => s.name === `${prefix}${n}`)) n++;
    return `${prefix}${n}`;
  }

  /** Get a sheet by name or index. */
  sheet(nameOrIndex) {
    if (typeof nameOrIndex === 'number') return this.sheets[nameOrIndex] || null;
    return this.sheets.find(s => s.name === nameOrIndex) || null;
  }

  /** Currently active sheet (used by viewer). */
  get active() { return this.sheets[this.activeIndex] || null; }
  set active(nameOrIndexOrSheet) {
    let i;
    if (typeof nameOrIndexOrSheet === 'number') i = nameOrIndexOrSheet;
    else if (typeof nameOrIndexOrSheet === 'string') i = this.sheets.findIndex(s => s.name === nameOrIndexOrSheet);
    else i = this.sheets.indexOf(nameOrIndexOrSheet);
    if (i < 0 || i >= this.sheets.length) throw err('SHEET_NOT_FOUND', `シートが見つかりません`);
    this.activeIndex = i;
  }

  removeSheet(nameOrIndex) {
    const i = typeof nameOrIndex === 'number'
      ? nameOrIndex
      : this.sheets.findIndex(s => s.name === nameOrIndex);
    if (i < 0 || i >= this.sheets.length) throw err('SHEET_NOT_FOUND', `シートが見つかりません: ${nameOrIndex}`);
    if (this.sheets.length === 1) throw err('LAST_SHEET', '最後のシートは削除できません');
    this.sheets[i]._workbook = null;
    this.sheets.splice(i, 1);
    if (this.activeIndex >= this.sheets.length) this.activeIndex = this.sheets.length - 1;
    return this;
  }

  renameSheet(oldName, newName) {
    const s = this.sheet(oldName);
    if (!s) throw err('SHEET_NOT_FOUND', `シートが見つかりません: ${oldName}`);
    if (this.sheets.find(x => x !== s && x.name === newName)) {
      throw err('DUPLICATE_SHEET', `シート名 "${newName}" は既に存在します`);
    }
    s.name = newName;
    return this;
  }

  moveSheet(fromIndex, toIndex) {
    const [s] = this.sheets.splice(fromIndex, 1);
    this.sheets.splice(toIndex, 0, s);
    return this;
  }

  toJSON() {
    return {
      version: this.version,
      type: 'aix-workbook',
      activeSheet: this.activeIndex,
      sheets: this.sheets.map(s => {
        const j = s.toJSON();
        delete j.type;
        delete j.version;
        return j;
      })
    };
  }

  /**
   * Workbook-level markdown — by default returns a compact table of contents
   * (one row per sheet, with size + formula/merge/CF counts). Use this first
   * when reading an unknown workbook, then call `wb.sheet(name).toMarkdown()`
   * on the specific sheet you want to drill into.
   *
   *   wb.toMarkdown()                        // TOC
   *   wb.toMarkdown('Sales')                 // shorthand for wb.sheet('Sales').toMarkdown()
   *   wb.toMarkdown({ all: true })           // every sheet concatenated (heavy)
   */
  toMarkdown(arg) {
    // Single-sheet shortcut
    if (typeof arg === 'string') {
      const s = this.sheet(arg);
      if (!s) throw err('SHEET_NOT_FOUND', `シートが見つかりません: ${arg}`);
      return s.toMarkdown();
    }
    const opts = arg && typeof arg === 'object' ? arg : {};
    if (opts.all) {
      return this.sheets
        .map(s => `# ${s.name}\n\n${s.toMarkdown(opts)}`)
        .join('\n\n');
    }
    // Default: TOC
    const lines = [`# Workbook (${this.sheets.length} sheets)`, ''];
    lines.push('| Sheet | Range | Formulas | Merges | CF rules |');
    lines.push('| --- | --- | --- | --- | --- |');
    for (const s of this.sheets) {
      const u = s.usedRange();
      const range = u.rows === 0 ? '*(empty)*' : `${makeRef(0, 0)}:${makeRef(u.rows - 1, u.cols - 1)}`;
      const fc = Object.values(s.cells).filter(c => c.f).length;
      lines.push(`| ${s.name} | ${range} | ${fc} | ${s.merges.length} | ${s.cfs.length} |`);
    }
    lines.push('');
    lines.push('_Call `sheet.toMarkdown()` (or `wb.toMarkdown("SheetName")`) for the actual content of a sheet._');
    return lines.join('\n');
  }

  async save(path) {
    if (typeof process === 'undefined' || !process.versions?.node) {
      throw err('NODE_ONLY', 'save() は Node.js のみ');
    }
    const fs = await import('node:fs/promises');
    await fs.writeFile(path, JSON.stringify(this.toJSON(), null, 2), 'utf8');
    return this;
  }

  /**
   * Export as XLSX with all sheets. Requires SheetJS for the base file,
   * and JSZip when any sheet contains images (manual zip injection).
   */
  async toXLSX() {
    const XLSX = await _loadXLSX();
    if (!XLSX) throw err('XLSX_MISSING', 'SheetJS が見つかりません');
    const wb = XLSX.utils.book_new();
    for (const s of this.sheets) {
      const u = s.usedRange();
      const aoa = _writeAOA(s, u);
      const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
      _patchDateCells(s, ws);
      _applySheetMetaToWS(s, ws, u);
      XLSX.utils.book_append_sheet(wb, ws, s.name);
    }
    const base = XLSX.write(wb, { type: typeof window !== 'undefined' ? 'array' : 'buffer', bookType: 'xlsx' });

    let out = base;
    if (this.sheets.some(s => s.images && s.images.length > 0)) {
      const { injectImages } = await import('./xlsx-images.js');
      out = await injectImages(out, this.sheets.map(s => ({ name: s.name, images: s.images })));
    }
    if (this.sheets.some(s => s.cfs && s.cfs.length > 0)) {
      const { injectCF } = await import('./xlsx-cf.js');
      out = await injectCF(out, this.sheets.map(s => ({ name: s.name, cfs: s.cfs })));
    }
    return out;
  }
}

export { Sheet, parseRef, parseRange, makeRef, idxToCol, colToIdx,
         formatCellValue, dateToSerial, serialToDate };
export default Sheet;

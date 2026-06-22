/**
 * aix-sheet Viewer (vanilla JS, framework-agnostic)
 *
 *   import { SheetView } from 'aix-sheet/viewer';
 *   import 'aix-sheet/viewer/styles.css';
 *
 *   const view = new SheetView('#app', { sheet, rows: 50, cols: 26 });
 *   view.on('change', ({ ref, value }) => { ... });
 *   view.destroy();
 */
import { Sheet, Workbook, parseRef, parseRange, makeRef, idxToCol, colToIdx, formatCellValue } from '../sdk/sheet.js';

const FUNCTIONS = [
  { name: 'SUM',     desc: '合計' }, { name: 'AVERAGE', desc: '平均' },
  { name: 'MAX',     desc: '最大値' }, { name: 'MIN', desc: '最小値' },
  { name: 'COUNT',   desc: '数値の個数' }, { name: 'COUNTA', desc: '空でない個数' },
  { name: 'IF',      desc: '条件分岐' }, { name: 'AND', desc: '論理積' },
  { name: 'OR',      desc: '論理和' }, { name: 'NOT', desc: '論理否定' },
  { name: 'ROUND',   desc: '四捨五入' }, { name: 'FLOOR', desc: '切り捨て' },
  { name: 'CEILING', desc: '切り上げ' }, { name: 'ABS', desc: '絶対値' },
  { name: 'SQRT',    desc: '平方根' }, { name: 'POWER', desc: '累乗' },
  { name: 'MOD',     desc: '余り' }, { name: 'LEN', desc: '文字数' },
  { name: 'LEFT',    desc: '左から取得' }, { name: 'RIGHT', desc: '右から取得' },
  { name: 'MID',     desc: '中間取得' }, { name: 'UPPER', desc: '大文字' },
  { name: 'LOWER',   desc: '小文字' }, { name: 'TRIM', desc: 'スペース除去' },
  { name: 'CONCAT',  desc: '文字列結合' }
];

/** Simple event emitter. */
class EventBus {
  constructor() { this._h = new Map(); }
  on(ev, fn)  { if (!this._h.has(ev)) this._h.set(ev, new Set()); this._h.get(ev).add(fn); return this; }
  off(ev, fn) { this._h.get(ev)?.delete(fn); return this; }
  emit(ev, payload) { this._h.get(ev)?.forEach(fn => { try { fn(payload); } catch(e){ console.error(e); } }); }
}

export class SheetView extends EventBus {
  /**
   * @param {string|HTMLElement} container CSS selector or DOM element
   * @param {object} options
   * @param {Workbook} [options.workbook] multi-sheet workbook (preferred)
   * @param {Sheet} [options.sheet] single-sheet shortcut (wrapped in a Workbook)
   * @param {number} [options.rows=50]
   * @param {number} [options.cols=26]
   * @param {boolean} [options.readOnly=false]
   * @param {boolean} [options.toolbar=true]
   * @param {boolean} [options.formulaBar=true]
   * @param {boolean} [options.statusBar=true]
   * @param {boolean} [options.tabs=true]
   */
  constructor(container, options = {}) {
    super();
    this.container = typeof container === 'string' ? document.querySelector(container) : container;
    if (!this.container) throw new Error('[SheetView] container not found');

    // Resolve workbook + active sheet
    if (options.workbook) {
      this.workbook = options.workbook;
    } else if (options.sheet) {
      this.workbook = new Workbook();
      this.workbook._attach(options.sheet);
    } else {
      this.workbook = new Workbook();
      this.workbook.addSheet('Sheet1');
    }
    this.sheet = this.workbook.active;

    this.numRows = options.rows ?? 50;
    this.numCols = options.cols ?? 26;
    this.readOnly = !!options.readOnly;
    this.showToolbar    = options.toolbar    !== false;
    this.showFormulaBar = options.formulaBar !== false;
    this.showStatusBar  = options.statusBar  !== false;
    this.showTabs       = options.tabs       !== false;

    const u = this.sheet.usedRange();
    if (u.rows > this.numRows) this.numRows = u.rows + 10;
    if (u.cols > this.numCols) this.numCols = u.cols + 5;

    // UI state
    this.activeRow = 0; this.activeCol = 0;
    this.selStart = { r: 0, c: 0 }; this.selEnd = { r: 0, c: 0 };
    this.isEditing = false;
    this.isSelectingRange = false;
    this.currentEditor = null;
    this.colWidths = new Array(this.numCols).fill(100);
    this.undoStack = [];
    this.clipboard = null;
    this.fpick = { active: false, anchor: null, end: null, input: null, insStart: 0, insEnd: 0 };
    this.ac = { items: [], index: -1, inputEl: null, wordStart: 0 };

    // Pull col widths from sheet
    for (const [label, w] of Object.entries(this.sheet.cols)) {
      const ci = colToIdx(label);
      if (ci < this.numCols) this.colWidths[ci] = w;
    }

    this._handlers = { keydown: null, mouseup: null, docClick: null };
    this._build();
  }

  // ============================================================
  // Public API
  // ============================================================

  /** Re-render the current sheet (call after external mutations). */
  refresh() { this._renderAllCells(); this._renderImages(); this._updateSelection(); this._updateFormulaBar(); }

  /** Replace the underlying Sheet (wraps it in a new single-sheet Workbook). */
  setSheet(sheet) {
    const wb = new Workbook();
    wb._attach(sheet);
    this.setWorkbook(wb);
  }

  /** Get the underlying Sheet instance. */
  getSheet() { return this.sheet; }

  /** Switch active sheet by name or index. */
  switchSheet(nameOrIndex) {
    if (this.isEditing) this._commitEdit();
    this.workbook.active = nameOrIndex;
    this.sheet = this.workbook.active;
    // reset per-sheet UI state
    this.activeRow = 0; this.activeCol = 0;
    this.selStart = { r: 0, c: 0 }; this.selEnd = { r: 0, c: 0 };
    this.undoStack.length = 0;
    // resize grid if needed
    const u = this.sheet.usedRange();
    if (u.rows > this.numRows) this.numRows = u.rows + 10;
    if (u.cols > this.numCols) this.numCols = u.cols + 5;
    // restore col widths from this sheet
    this.colWidths = new Array(this.numCols).fill(100);
    for (const [label, w] of Object.entries(this.sheet.cols)) {
      const ci = colToIdx(label);
      if (ci < this.numCols) this.colWidths[ci] = w;
    }
    this._buildTable();
    this._renderTabs();
    this._updateSelection(); this._updateFormulaBar();
    this.emit('sheet-change', { name: this.sheet.name, index: this.workbook.activeIndex });
  }

  /** Add a new sheet and switch to it. */
  addSheet(name) {
    const s = this.workbook.addSheet(name);
    this.switchSheet(s.name);
    return s;
  }

  /** Tear down the viewer and remove from DOM. */
  destroy() {
    document.removeEventListener('keydown', this._handlers.keydown);
    document.removeEventListener('mouseup', this._handlers.mouseup);
    document.removeEventListener('click', this._handlers.docClick);
    this.root.remove();
    this.contextMenu?.remove();
    this.acEl?.remove();
    this.container.classList.remove('aix-host');
  }

  /** Programmatically focus a cell. */
  focusCell(ref) {
    const p = parseRef(ref);
    this.activeRow = p.row; this.activeCol = p.col;
    this.selStart = { r: p.row, c: p.col }; this.selEnd = { r: p.row, c: p.col };
    this._updateSelection(); this._updateFormulaBar();
  }

  // ============================================================
  // Build DOM
  // ============================================================
  _build() {
    this.container.classList.add('aix-host');
    this.root = document.createElement('div');
    this.root.className = 'aix-root';
    this.container.appendChild(this.root);

    if (this.showToolbar)    this._buildToolbar();
    if (this.showFormulaBar) this._buildFormulaBar();

    this.sheetContainer = document.createElement('div');
    this.sheetContainer.className = 'aix-sheet-container';
    this.table = document.createElement('table');
    this.table.className = 'aix-spreadsheet';
    this.sheetContainer.appendChild(this.table);

    // Image overlay layer sits on top of the table inside the scrollable container.
    this.imageLayer = document.createElement('div');
    this.imageLayer.className = 'aix-image-layer';
    this.sheetContainer.appendChild(this.imageLayer);

    // Formula-reference highlight layer (colored outlines around refs while editing).
    this.refBoxLayer = document.createElement('div');
    this.refBoxLayer.className = 'aix-fref-layer';
    this.sheetContainer.appendChild(this.refBoxLayer);

    this.root.appendChild(this.sheetContainer);
    this.selectedImageId = null;

    if (this.showTabs)      this._buildTabs();
    if (this.showStatusBar) this._buildStatusBar();

    this._buildContextMenu();
    this._buildAutocomplete();

    this._buildTable();
    this._attachGlobalHandlers();
    this._updateSelection();
    this._updateFormulaBar();
  }

  _buildToolbar() {
    const tb = document.createElement('div');
    tb.className = 'aix-toolbar';
    tb.innerHTML = `
      <span class="aix-title">📊 ${this.sheet.name}</span>
      <div class="aix-sep"></div>
      <button data-act="add-row">+ 行</button>
      <button data-act="del-row">- 行</button>
      <button data-act="add-col">+ 列</button>
      <button data-act="del-col">- 列</button>
      <div class="aix-sep"></div>
      <select data-act="font-family">
        <option value="Segoe UI,Meiryo,sans-serif">Segoe UI</option>
        <option value="Arial,sans-serif">Arial</option>
        <option value="Consolas,monospace">Consolas</option>
        <option value="メイリオ,Meiryo,sans-serif">メイリオ</option>
      </select>
      <input data-act="font-size" type="number" value="13" min="8" max="72" style="width:42px">
      <button data-act="bold"><b>B</b></button>
      <button data-act="italic"><i>I</i></button>
      <button data-act="underline"><u>U</u></button>
      <div class="aix-sep"></div>
      <button data-act="align-left">≡</button>
      <button data-act="align-center">≡</button>
      <button data-act="align-right">≡</button>
      <div class="aix-sep"></div>
      <label>A <input data-act="color" type="color" value="#000000"></label>
      <label>BG <input data-act="bg" type="color" value="#ffffff"></label>
      <div class="aix-sep"></div>
      <span class="aix-split">
        <button data-act="merge-center" title="セルを結合して中央揃え">⬚ 結合して中央</button>
        <button data-act="merge-menu" class="aix-split-arrow" title="結合メニュー">▼</button>
      </span>
      <select data-act="border" title="罫線">
        <option value="">罫線…</option>
        <option value="all">全て</option>
        <option value="outside">外枠</option>
        <option value="top">上</option>
        <option value="right">右</option>
        <option value="bottom">下</option>
        <option value="left">左</option>
        <option value="none">なし</option>
      </select>
      <select data-act="numfmt" title="表示形式">
        <option value="">書式…</option>
        <option value="__clear">標準</option>
        <option value="0">数値 (整数)</option>
        <option value="0.00">数値 (.00)</option>
        <option value="#,##0">数値 (3桁区切り)</option>
        <option value="yyyy/m/d">日付 (2026/6/1)</option>
        <option value="m/d">日付 (6/1)</option>
        <option value="yyyy-mm-dd">日付 (2026-06-01)</option>
        <option value="yyyy年m月d日">日付 (2026年6月1日)</option>
        <option value="¥#,##0">通貨 (¥)</option>
        <option value="0.0%">パーセント</option>
      </select>
      <button data-act="cf" title="条件付き書式">⚡ 条件付き書式</button>
      <div class="aix-sep"></div>
      <button data-act="add-image">🖼 画像</button>
      <div class="aix-sep"></div>
      <button data-act="save-aix">.aix.json 保存</button>
      <button data-act="load-aix">.aix.json 読込</button>
      <button data-act="save-csv">CSV保存</button>
      <button data-act="save-xlsx" title="Excel 形式 (.xlsx) で保存">📊 Excel 保存</button>
      <input type="file" data-act="file-aix" accept=".json" hidden>
      <input type="file" data-act="file-csv" accept=".csv" hidden>
      <input type="file" data-act="file-image" accept="image/*" hidden>
    `;
    if (this.readOnly) tb.querySelectorAll('button, input, select').forEach(el => el.disabled = true);
    this.root.appendChild(tb);
    this.toolbar = tb;

    tb.addEventListener('click', (e) => {
      const act = e.target.closest('[data-act]')?.dataset.act;
      if (!act) return;
      this._onToolbar(act, e);
    });
    tb.addEventListener('change', (e) => {
      const act = e.target.dataset.act;
      if (act === 'font-family')  this._applyFormat('fontFamily', e.target.value);
      if (act === 'font-size')    this._applyFormat('fontSize', Number(e.target.value));
      if (act === 'color')        this._applyFormat('color', e.target.value);
      if (act === 'bg')           this._applyFormat('bgColor', e.target.value);
      if (act === 'border')       { this._applyBorder(e.target.value); e.target.value = ''; }
      if (act === 'numfmt')       { this._applyNumFmt(e.target.value); e.target.value = ''; }
      if (act === 'file-aix')     this._handleFileLoad(e.target.files[0], 'aix');
      if (act === 'file-csv')     this._handleFileLoad(e.target.files[0], 'csv');
      if (act === 'file-image')   this._handleImageFile(e.target.files[0]);
    });
  }

  _onToolbar(act, e) {
    if (act === 'add-row')      this._insertRow(this.activeRow + 1, 1);
    else if (act === 'del-row') this._deleteRow();
    else if (act === 'add-col') this._insertCol(this.activeCol + 1, 1);
    else if (act === 'del-col') this._deleteCol();
    else if (act === 'bold')    this._toggleStyle('bold');
    else if (act === 'italic')  this._toggleStyle('italic');
    else if (act === 'underline') this._toggleStyle('underline');
    else if (act === 'align-left')   this._applyFormat('align', 'left');
    else if (act === 'align-center') this._applyFormat('align', 'center');
    else if (act === 'align-right')  this._applyFormat('align', 'right');
    else if (act === 'save-aix') this._downloadAIX();
    else if (act === 'save-csv') this._downloadCSV();
    else if (act === 'save-xlsx') this._downloadXLSX();
    else if (act === 'load-aix') this.toolbar.querySelector('[data-act="file-aix"]').click();
    else if (act === 'add-image') this.toolbar.querySelector('[data-act="file-image"]').click();
    else if (act === 'merge-center') this._mergeAndCenter();
    else if (act === 'merge-menu')   this._openMergeMenu(e);
    else if (act === 'cf')       this._openCFDialog();
  }

  // ---- Merge / Border (toolbar handlers) ----

  // Excel: "セルを結合して中央揃え"
  _mergeAndCenter() {
    if (this.readOnly) return;
    const sel = this._currentSelectionBox();
    // If active cell is in an existing merge, toggle off
    const existing = this.sheet.mergeAt(this.activeRow, this.activeCol);
    if (existing) {
      const ref = makeRef(existing.r1, existing.c1) + ':' + makeRef(existing.r2, existing.c2);
      this.sheet.unmerge(ref);
    } else {
      if (sel.r1 === sel.r2 && sel.c1 === sel.c2) return;
      const ref = makeRef(sel.r1, sel.c1) + ':' + makeRef(sel.r2, sel.c2);
      try { this.sheet.merge(ref); }
      catch (err) { alert(err.message); return; }
      this.sheet.style(makeRef(sel.r1, sel.c1), { align: 'center' });
    }
    this._buildTable();
    this.emit('change', { type: 'merge' });
  }

  // Excel: "セルを結合" (without centering)
  _mergeCellsOnly() {
    if (this.readOnly) return;
    const sel = this._currentSelectionBox();
    if (sel.r1 === sel.r2 && sel.c1 === sel.c2) return;
    const ref = makeRef(sel.r1, sel.c1) + ':' + makeRef(sel.r2, sel.c2);
    try { this.sheet.merge(ref); }
    catch (err) { alert(err.message); return; }
    this._buildTable();
    this.emit('change', { type: 'merge' });
  }

  // Excel: "横方向に結合" (each row in selection merged independently)
  _mergeAcross() {
    if (this.readOnly) return;
    const sel = this._currentSelectionBox();
    if (sel.c1 === sel.c2) return;
    for (let r = sel.r1; r <= sel.r2; r++) {
      const ref = makeRef(r, sel.c1) + ':' + makeRef(r, sel.c2);
      try { this.sheet.merge(ref); } catch (err) { /* skip rows that overlap */ }
    }
    this._buildTable();
    this.emit('change', { type: 'merge' });
  }

  // Excel: "セル結合の解除"
  _unmergeCells() {
    if (this.readOnly) return;
    const sel = this._currentSelectionBox();
    // Remove every merge that overlaps the selection
    const remaining = [];
    for (const m of this.sheet.merges) {
      const overlap = !(sel.r2 < m.r1 || sel.r1 > m.r2 || sel.c2 < m.c1 || sel.c1 > m.c2);
      if (!overlap) remaining.push(m);
    }
    this.sheet.merges = remaining;
    this._buildTable();
    this.emit('change', { type: 'unmerge' });
  }

  _currentSelectionBox() {
    return {
      r1: Math.min(this.selStart.r, this.selEnd.r),
      r2: Math.max(this.selStart.r, this.selEnd.r),
      c1: Math.min(this.selStart.c, this.selEnd.c),
      c2: Math.max(this.selStart.c, this.selEnd.c),
    };
  }

  _openMergeMenu(e) {
    if (this._mergeMenu) { this._closeMergeMenu(); return; }
    const anchor = e.target.closest('button');
    if (!anchor) return;
    const rect = anchor.getBoundingClientRect();
    const menu = document.createElement('div');
    menu.className = 'aix-popup-menu';
    menu.style.top  = (rect.bottom + window.scrollY + 2) + 'px';
    menu.style.left = (rect.left + window.scrollX) + 'px';
    menu.innerHTML = `
      <button type="button" data-merge-act="across">横方向に結合</button>
      <button type="button" data-merge-act="merge">セルを結合</button>
      <button type="button" data-merge-act="unmerge">セル結合の解除</button>
    `;
    document.body.appendChild(menu);
    this._mergeMenu = menu;
    menu.addEventListener('click', (ev) => {
      const act = ev.target.dataset.mergeAct;
      this._closeMergeMenu();
      if (act === 'across')  this._mergeAcross();
      if (act === 'merge')   this._mergeCellsOnly();
      if (act === 'unmerge') this._unmergeCells();
    });
    // Close on outside click
    setTimeout(() => {
      const closer = (ev) => {
        if (this._mergeMenu && !this._mergeMenu.contains(ev.target)) {
          this._closeMergeMenu();
          document.removeEventListener('mousedown', closer);
        }
      };
      document.addEventListener('mousedown', closer);
    }, 0);
  }

  _closeMergeMenu() {
    if (this._mergeMenu) { this._mergeMenu.remove(); this._mergeMenu = null; }
  }

  _applyBorder(kind) {
    if (this.readOnly || !kind) return;
    const r1 = Math.min(this.selStart.r, this.selEnd.r);
    const r2 = Math.max(this.selStart.r, this.selEnd.r);
    const c1 = Math.min(this.selStart.c, this.selEnd.c);
    const c2 = Math.max(this.selStart.c, this.selEnd.c);
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++) this._saveUndo(r, c);
    if (kind === 'none') {
      const ref = makeRef(r1, c1) + ':' + makeRef(r2, c2);
      this.sheet.style(ref, { border: false });
    } else if (kind === 'all') {
      const ref = makeRef(r1, c1) + ':' + makeRef(r2, c2);
      this.sheet.style(ref, { border: true });
    } else if (kind === 'outside') {
      // Top row, bottom row, left col, right col only on the boundary side
      for (let c = c1; c <= c2; c++) {
        this.sheet.style(makeRef(r1, c), { border: { top: true } });
        this.sheet.style(makeRef(r2, c), { border: { bottom: true } });
      }
      for (let r = r1; r <= r2; r++) {
        this.sheet.style(makeRef(r, c1), { border: { left: true } });
        this.sheet.style(makeRef(r, c2), { border: { right: true } });
      }
    } else {
      // Single side: top|right|bottom|left
      const ref = makeRef(r1, c1) + ':' + makeRef(r2, c2);
      this.sheet.style(ref, { border: { [kind]: true } });
    }
    this._renderAllCells();
    this.emit('change', { type: 'border' });
  }

  _applyNumFmt(fmt) {
    if (this.readOnly || !fmt) return;
    const r1 = Math.min(this.selStart.r, this.selEnd.r);
    const r2 = Math.max(this.selStart.r, this.selEnd.r);
    const c1 = Math.min(this.selStart.c, this.selEnd.c);
    const c2 = Math.max(this.selStart.c, this.selEnd.c);
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++) this._saveUndo(r, c);
    const ref = makeRef(r1, c1) + ':' + makeRef(r2, c2);
    if (fmt === '__clear') {
      // Manually clear numFmt from each cell
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
        const key = makeRef(r, c);
        if (this.sheet.cells[key]?.s?.numFmt) delete this.sheet.cells[key].s.numFmt;
      }
    } else {
      this.sheet.style(ref, { numFmt: fmt });
    }
    this._renderAllCells();
    this.emit('change', { type: 'numfmt' });
  }

  // ---- Conditional-formatting dialog ----

  _openCFDialog() {
    if (this.readOnly) return;
    if (this._cfDialog) this._closeCFDialog();

    const r1 = Math.min(this.selStart.r, this.selEnd.r);
    const r2 = Math.max(this.selStart.r, this.selEnd.r);
    const c1 = Math.min(this.selStart.c, this.selEnd.c);
    const c2 = Math.max(this.selStart.c, this.selEnd.c);
    const defaultRange = makeRef(r1, c1) + (r1 === r2 && c1 === c2 ? '' : ':' + makeRef(r2, c2));

    const overlay = document.createElement('div');
    overlay.className = 'aix-modal-overlay';
    overlay.innerHTML = `
      <div class="aix-modal">
        <div class="aix-modal-header">
          <span>⚡ 条件付き書式</span>
          <button class="aix-modal-close" type="button">&times;</button>
        </div>
        <div class="aix-modal-body">
          <div class="aix-cf-section">
            <h4>既存のルール</h4>
            <div class="aix-cf-list"></div>
          </div>
          <div class="aix-cf-section">
            <h4>ルールを追加</h4>
            <div class="aix-cf-row"><label>範囲</label>
              <input class="aix-cf-range" value="${defaultRange}" placeholder="A1:B5"></div>
            <div class="aix-cf-row"><label>数式</label>
              <input class="aix-cf-formula" placeholder="=A1>100"></div>
            <div class="aix-cf-row aix-cf-style-row">
              <label><input type="checkbox" class="aix-cf-bold"> 太字</label>
              <label><input type="checkbox" class="aix-cf-italic"> 斜体</label>
              <label>文字 <input type="color" class="aix-cf-color" value="#000000"></label>
              <label>背景 <input type="color" class="aix-cf-bg" value="#ffeb9c"></label>
            </div>
            <div class="aix-cf-row aix-cf-preset-row">
              <span>プリセット:</span>
              <button type="button" data-preset="positive">プラス値を緑</button>
              <button type="button" data-preset="negative">マイナス値を赤</button>
              <button type="button" data-preset="weekend">週末をオレンジ</button>
            </div>
            <div class="aix-cf-row">
              <button class="aix-cf-add" type="button">追加</button>
            </div>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    this._cfDialog = overlay;

    const renderList = () => {
      const list = overlay.querySelector('.aix-cf-list');
      list.innerHTML = '';
      if (!this.sheet.cfs.length) {
        list.innerHTML = '<div class="aix-cf-empty">ルールはまだありません</div>';
        return;
      }
      this.sheet.cfs.forEach((rule, i) => {
        const item = document.createElement('div');
        item.className = 'aix-cf-item';
        const swatch = rule.style.bgColor || '#ddd';
        item.innerHTML = `
          <span class="aix-cf-swatch" style="background:${swatch}"></span>
          <span class="aix-cf-range-label">${rule.range}</span>
          <code class="aix-cf-formula-label">${rule.formula}</code>
          <button class="aix-cf-del" type="button" data-i="${i}">削除</button>
        `;
        list.appendChild(item);
      });
      list.querySelectorAll('.aix-cf-del').forEach(b => b.onclick = () => {
        const i = Number(b.dataset.i);
        this.sheet.cfs.splice(i, 1);
        this._renderAllCells();
        renderList();
        this.emit('change', { type: 'cf-delete' });
      });
    };
    renderList();

    overlay.querySelector('.aix-modal-close').onclick = () => this._closeCFDialog();
    overlay.addEventListener('click', (e) => { if (e.target === overlay) this._closeCFDialog(); });

    const formulaInput = overlay.querySelector('.aix-cf-formula');
    overlay.querySelectorAll('[data-preset]').forEach(btn => btn.onclick = () => {
      const p = btn.dataset.preset;
      const firstCell = (overlay.querySelector('.aix-cf-range').value.split(':')[0] || 'A1');
      if (p === 'positive') {
        formulaInput.value = `=${firstCell}>0`;
        overlay.querySelector('.aix-cf-bg').value = '#c6efce';
        overlay.querySelector('.aix-cf-color').value = '#006100';
      } else if (p === 'negative') {
        formulaInput.value = `=${firstCell}<0`;
        overlay.querySelector('.aix-cf-bg').value = '#ffc7ce';
        overlay.querySelector('.aix-cf-color').value = '#9c0006';
      } else if (p === 'weekend') {
        formulaInput.value = `=OR(WEEKDAY(${firstCell})=1, WEEKDAY(${firstCell})=7)`;
        overlay.querySelector('.aix-cf-bg').value = '#fce4d6';
        overlay.querySelector('.aix-cf-color').value = '#000000';
      }
    });

    overlay.querySelector('.aix-cf-add').onclick = () => {
      const range   = overlay.querySelector('.aix-cf-range').value.trim();
      const formula = overlay.querySelector('.aix-cf-formula').value.trim();
      if (!range || !formula) { alert('範囲と数式を入力してください'); return; }
      const fm = formula.startsWith('=') ? formula : '=' + formula;
      const style = {};
      if (overlay.querySelector('.aix-cf-bold').checked)   style.bold = true;
      if (overlay.querySelector('.aix-cf-italic').checked) style.italic = true;
      const bg = overlay.querySelector('.aix-cf-bg').value;
      const fg = overlay.querySelector('.aix-cf-color').value;
      if (bg && bg !== '#ffffff') style.bgColor = bg;
      if (fg && fg !== '#000000') style.color = fg;
      try { this.sheet.cf(range, { formula: fm, style }); }
      catch (e) { alert(e.message); return; }
      this._renderAllCells();
      renderList();
      overlay.querySelector('.aix-cf-formula').value = '';
      this.emit('change', { type: 'cf-add' });
    };
  }

  _closeCFDialog() {
    if (this._cfDialog) {
      this._cfDialog.remove();
      this._cfDialog = null;
    }
  }

  _buildFormulaBar() {
    const fb = document.createElement('div');
    fb.className = 'aix-formula-bar';
    fb.innerHTML = `
      <input class="aix-cell-ref" readonly value="A1">
      <span style="color:#888; font-size:11px;">fx</span>
      <input class="aix-formula-input" type="text" placeholder="値または数式 (例: =SUM(A1:A10))">
    `;
    this.root.appendChild(fb);
    this.cellRefEl     = fb.querySelector('.aix-cell-ref');
    this.formulaInput  = fb.querySelector('.aix-formula-input');

    this.formulaInput.addEventListener('input',   () => this._onFormulaBarInput());
    this.formulaInput.addEventListener('keydown', (e) => this._onFormulaBarKeyDown(e));
  }

  _buildTabs() {
    const bar = document.createElement('div');
    bar.className = 'aix-tabs';
    this.root.appendChild(bar);
    this.tabsBar = bar;
    this._renderTabs();
  }

  _renderTabs() {
    if (!this.tabsBar) return;
    this.tabsBar.innerHTML = '';
    this.workbook.sheets.forEach((s, i) => {
      const tab = document.createElement('div');
      tab.className = 'aix-tab' + (i === this.workbook.activeIndex ? ' aix-tab-active' : '');
      tab.textContent = s.name;
      tab.title = s.name;
      tab.addEventListener('click', () => { if (i !== this.workbook.activeIndex) this.switchSheet(i); });
      tab.addEventListener('dblclick', () => this._renameTab(i));
      tab.addEventListener('contextmenu', (e) => this._tabContextMenu(e, i));
      this.tabsBar.appendChild(tab);
    });
    if (!this.readOnly) {
      const add = document.createElement('div');
      add.className = 'aix-tab-add';
      add.textContent = '+';
      add.title = 'シートを追加';
      add.addEventListener('click', () => this.addSheet());
      this.tabsBar.appendChild(add);
    }
    // Update toolbar title to match active sheet
    const titleEl = this.toolbar?.querySelector('.aix-title');
    if (titleEl) titleEl.textContent = '📊 ' + this.sheet.name;
  }

  _renameTab(i) {
    if (this.readOnly) return;
    const old = this.workbook.sheets[i].name;
    const name = prompt('シート名を変更:', old);
    if (!name || name === old) return;
    try {
      this.workbook.renameSheet(old, name);
      this._renderTabs();
    } catch (e) { alert(e.message); }
  }

  _tabContextMenu(e, i) {
    e.preventDefault();
    if (this.readOnly) return;
    if (this.workbook.sheets.length === 1) return;
    if (confirm(`シート "${this.workbook.sheets[i].name}" を削除しますか？`)) {
      this.workbook.removeSheet(i);
      this.sheet = this.workbook.active;
      this._buildTable(); this._renderTabs(); this._updateSelection(); this._updateFormulaBar();
    }
  }

  _buildStatusBar() {
    const sb = document.createElement('div');
    sb.className = 'aix-status-bar';
    sb.innerHTML = `<span class="aix-status-cell">A1</span><span class="aix-status-info">準備完了</span><span class="aix-status-calc"></span>`;
    this.root.appendChild(sb);
    this.statusCell = sb.querySelector('.aix-status-cell');
    this.statusInfo = sb.querySelector('.aix-status-info');
    this.statusCalc = sb.querySelector('.aix-status-calc');
  }

  _buildContextMenu() {
    const m = document.createElement('div');
    m.className = 'aix-context-menu';
    m.innerHTML = `
      <div data-ctx="ins-row-above">上に行を挿入</div>
      <div data-ctx="ins-row-below">下に行を挿入</div>
      <div class="aix-sep"></div>
      <div data-ctx="ins-col-left">左に列を挿入</div>
      <div data-ctx="ins-col-right">右に列を挿入</div>
      <div class="aix-sep"></div>
      <div data-ctx="del-row">行を削除</div>
      <div data-ctx="del-col">列を削除</div>
      <div class="aix-sep"></div>
      <div data-ctx="clear">内容をクリア</div>
    `;
    document.body.appendChild(m);
    this.contextMenu = m;
    m.addEventListener('click', (e) => {
      const act = e.target.dataset.ctx;
      if (!act) return;
      this._onContextMenu(act);
      this._hideContextMenu();
    });
  }

  _onContextMenu(act) {
    const r = this.ctxRow, c = this.ctxCol;
    if (act === 'ins-row-above') this._insertRow(r + 1, 1);
    else if (act === 'ins-row-below') this._insertRow(r + 2, 1);
    else if (act === 'ins-col-left')  this._insertCol(c + 1, 1);
    else if (act === 'ins-col-right') this._insertCol(c + 2, 1);
    else if (act === 'del-row') this._deleteRow();
    else if (act === 'del-col') this._deleteCol();
    else if (act === 'clear')   this._clearSelection();
  }

  _buildAutocomplete() {
    const ac = document.createElement('div');
    ac.className = 'aix-autocomplete';
    document.body.appendChild(ac);
    this.acEl = ac;
  }

  _attachGlobalHandlers() {
    this._handlers.keydown = (e) => this._onKeyDown(e);
    this._handlers.mouseup = () => this._onMouseUp();
    this._handlers.docClick = () => { this._hideContextMenu(); this._hideAutocomplete(); };
    document.addEventListener('keydown', this._handlers.keydown);
    document.addEventListener('mouseup', this._handlers.mouseup);
    document.addEventListener('click', this._handlers.docClick);
  }

  // ============================================================
  // Table rendering
  // ============================================================
  _buildTable() {
    this.table.innerHTML = '';
    this._cellMap = new Map();   // (r,c) → TD, keyed by r*(numCols+1)+c
    const thead = this.table.createTHead();
    const hrow = thead.insertRow();
    const corner = document.createElement('th');
    corner.className = 'aix-corner';
    corner.onclick = () => this._selectAll();
    hrow.appendChild(corner);

    for (let c = 0; c < this.numCols; c++) {
      const th = document.createElement('th');
      th.className = 'aix-col-header';
      th.dataset.col = c;
      th.style.width = this.colWidths[c] + 'px';
      th.style.minWidth = this.colWidths[c] + 'px';
      th.textContent = idxToCol(c);
      const handle = document.createElement('div');
      handle.className = 'aix-col-resize-handle';
      handle.addEventListener('mousedown', (e) => this._startColResize(e, c));
      th.appendChild(handle);
      th.addEventListener('click', (e) => this._selectColumn(c, e));
      hrow.appendChild(th);
    }

    const tbody = this.table.createTBody();
    for (let r = 0; r < this.numRows; r++) {
      const row = tbody.insertRow();
      const rh = document.createElement('th');
      rh.className = 'aix-row-header';
      rh.textContent = r + 1;
      rh.addEventListener('click', (e) => this._selectRow(r, e));
      row.appendChild(rh);

      for (let c = 0; c < this.numCols; c++) {
        const td = row.insertCell();
        td.dataset.row = r; td.dataset.col = c;
        this._cellMap.set(r * (this.numCols + 1) + c, td);
        this._renderCell(td, r, c);
        td.addEventListener('mousedown', (e) => this._onCellMouseDown(e, r, c));
        td.addEventListener('mouseover',  (e) => this._onCellMouseOver(e, r, c));
        td.addEventListener('dblclick',   () => this._startEdit(r, c));
        td.addEventListener('contextmenu',(e) => this._showContextMenu(e, r, c));
      }
    }
    this._applyMerges();
    this._updateSelection();
    this._renderImages();
  }

  _applyMerges() {
    if (!this.sheet.merges) return;
    // `_buildTable` always rebuilds from scratch, so there's no prior merge
    // state to undo here — we just set rowSpan/colSpan on each anchor and
    // remove the covered TDs entirely from the DOM (proper HTML way to merge).
    for (const m of this.sheet.merges) {
      const anchor = this._getCell(m.r1, m.c1);
      if (!anchor) continue;
      anchor.rowSpan = m.r2 - m.r1 + 1;
      anchor.colSpan = m.c2 - m.c1 + 1;
      anchor.classList.add('aix-merge-anchor');
      for (let r = m.r1; r <= m.r2; r++) {
        for (let c = m.c1; c <= m.c2; c++) {
          if (r === m.r1 && c === m.c1) continue;
          const td = this._getCell(r, c);
          if (td) {
            td.remove();
            this._cellMap.delete(r * (this.numCols + 1) + c);
          }
        }
      }
    }
  }

  _getCell(r, c) {
    if (r < 0 || r >= this.numRows || c < 0 || c >= this.numCols) return null;
    return this._cellMap?.get(r * (this.numCols + 1) + c) ?? null;
  }

  _renderCell(td, r, c) {
    const ref = makeRef(r, c);
    const cell = this.sheet.cells[ref];
    const v = this.sheet.value(ref);
    const base = cell?.s || {};
    td.textContent = formatCellValue(v, base.numFmt);
    const cfOverride = this.sheet.cfs?.length ? this.sheet.cfStyleAt(r, c) : null;
    this._applyStyleToEl(td, cfOverride ? { ...base, ...cfOverride } : base);
  }

  _renderAllCells() {
    for (let r = 0; r < this.numRows; r++)
      for (let c = 0; c < this.numCols; c++) {
        const td = this._getCell(r, c);
        if (td && !this.isEditing) this._renderCell(td, r, c);
      }
  }

  // ============================================================
  // Image overlay
  // ============================================================
  _renderImages() {
    if (!this.imageLayer) return;
    this.imageLayer.innerHTML = '';
    for (const img of this.sheet.images) {
      const el = this._createImageEl(img);
      if (el) this.imageLayer.appendChild(el);
    }
  }

  _createImageEl(img) {
    const p = parseRef(img.anchor);
    const td = this._getCell(p.row, p.col);
    if (!td) return null;
    // Position relative to the table (which is offsetParent of td after offsetParent chain)
    const left = td.offsetLeft + (img.offset?.x ?? 0);
    const top  = td.offsetTop  + (img.offset?.y ?? 0);

    const wrap = document.createElement('div');
    wrap.className = 'aix-image' + (this.selectedImageId === img.id ? ' aix-image-selected' : '');
    wrap.dataset.id = img.id;
    wrap.style.left   = left + 'px';
    wrap.style.top    = top  + 'px';
    wrap.style.width  = img.size.width  + 'px';
    wrap.style.height = img.size.height + 'px';

    const im = document.createElement('img');
    im.src = img.src;
    im.draggable = false;
    wrap.appendChild(im);

    // Resize handle
    const handle = document.createElement('div');
    handle.className = 'aix-image-resize';
    wrap.appendChild(handle);

    wrap.addEventListener('mousedown', (e) => this._onImageMouseDown(e, img.id));
    handle.addEventListener('mousedown', (e) => { e.stopPropagation(); this._startImageResize(e, img.id); });
    return wrap;
  }

  _onImageMouseDown(e, id) {
    if (e.button !== 0) return;
    e.stopPropagation();
    this.selectedImageId = id;
    this._renderImages();
    if (this.readOnly) return;
    this._startImageDrag(e, id);
  }

  _startImageDrag(e, id) {
    const img = this.sheet.getImage(id); if (!img) return;
    const wrap = this.imageLayer.querySelector(`[data-id="${id}"]`);
    const startX = e.clientX, startY = e.clientY;
    const startLeft = parseFloat(wrap.style.left), startTop = parseFloat(wrap.style.top);
    const move = (ev) => {
      wrap.style.left = (startLeft + ev.clientX - startX) + 'px';
      wrap.style.top  = (startTop  + ev.clientY - startY) + 'px';
    };
    const up = (ev) => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
      // Compute new anchor + offset
      const finalLeft = parseFloat(wrap.style.left), finalTop = parseFloat(wrap.style.top);
      const { anchor, offset } = this._coordToAnchor(finalLeft, finalTop);
      img.anchor = anchor;
      img.offset = offset;
      this._renderImages();
      this.emit('image-change', { id });
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  }

  _startImageResize(e, id) {
    const img = this.sheet.getImage(id); if (!img) return;
    const startX = e.clientX, startY = e.clientY;
    const startW = img.size.width, startH = img.size.height;
    const wrap = this.imageLayer.querySelector(`[data-id="${id}"]`);
    const move = (ev) => {
      const w = Math.max(20, startW + (ev.clientX - startX));
      const h = Math.max(20, startH + (ev.clientY - startY));
      wrap.style.width = w + 'px'; wrap.style.height = h + 'px';
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
      img.size = { width: parseFloat(wrap.style.width), height: parseFloat(wrap.style.height) };
      this._renderImages();
      this.emit('image-change', { id });
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  }

  /** Convert pixel coords (relative to table top-left) to nearest cell anchor + sub-cell offset. */
  _coordToAnchor(x, y) {
    let row = 0, col = 0, offY = 0, offX = 0;
    for (let r = 0; r < this.numRows; r++) {
      const td = this._getCell(r, 0); if (!td) continue;
      if (td.offsetTop + td.offsetHeight > y) { row = r; offY = y - td.offsetTop; break; }
      if (r === this.numRows - 1) { row = r; offY = y - td.offsetTop; }
    }
    for (let c = 0; c < this.numCols; c++) {
      const td = this._getCell(row, c); if (!td) continue;
      if (td.offsetLeft + td.offsetWidth > x) { col = c; offX = x - td.offsetLeft; break; }
      if (c === this.numCols - 1) { col = c; offX = x - td.offsetLeft; }
    }
    return { anchor: makeRef(row, col), offset: { x: Math.max(0, offX), y: Math.max(0, offY) } };
  }

  _deleteSelectedImage() {
    if (!this.selectedImageId) return false;
    this.sheet.removeImage(this.selectedImageId);
    this.selectedImageId = null;
    this._renderImages();
    return true;
  }

  _handleImageFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const dataUrl = e.target.result;
      // Default to natural image size up to 300px max, anchored at active cell
      const tempImg = new Image();
      tempImg.onload = () => {
        let w = tempImg.naturalWidth, h = tempImg.naturalHeight;
        const max = 300;
        if (w > max) { h *= max / w; w = max; }
        if (h > max) { w *= max / h; h = max; }
        const anchor = makeRef(this.activeRow, this.activeCol);
        this.sheet.addImage(anchor, dataUrl, { size: { width: Math.round(w), height: Math.round(h) } });
        this._renderImages();
        this.emit('image-change', { id: this.sheet.images[this.sheet.images.length - 1].id });
      };
      tempImg.src = dataUrl;
    };
    reader.readAsDataURL(file);
  }

  _applyStyleToEl(el, s) {
    el.style.fontWeight     = s.bold      ? 'bold'      : '';
    el.style.fontStyle      = s.italic    ? 'italic'    : '';
    el.style.textDecoration = s.underline ? 'underline' : '';
    el.style.color          = s.color     || '';
    el.style.backgroundColor= s.bgColor   || '';
    el.style.textAlign      = s.align     || '';
    el.style.fontSize       = s.fontSize  ? s.fontSize + 'px' : '';
    el.style.fontFamily     = s.fontFamily|| '';
    // Border (per-side: { style, color })
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const cssKey = 'border' + side[0].toUpperCase() + side.slice(1);
      const b = s.border?.[side];
      if (b) {
        const w = b.style === 'thick' ? '2px' : b.style === 'medium' ? '1.5px' : '1px';
        const cssStyle = (b.style === 'dashed' || b.style === 'dotted' || b.style === 'double') ? b.style : 'solid';
        el.style[cssKey] = `${w} ${cssStyle} ${b.color || '#000'}`;
      } else {
        el.style[cssKey] = '';
      }
    }
  }

  // ============================================================
  // Selection
  // ============================================================
  _onCellMouseDown(e, r, c) {
    if (e.button !== 0) return;

    // Save formula bar state before focus changes
    const formulaBarFocused = document.activeElement === this.formulaInput;
    const formulaBarPos = formulaBarFocused ? this.formulaInput.selectionStart : null;
    const formulaBarVal = this.formulaInput.value;

    // Formula bar in edit mode and cursor not at end → range pick mode
    if (formulaBarFocused && formulaBarVal.startsWith('=') && formulaBarPos < formulaBarVal.length) {
      e.preventDefault();
      this.fpick.active = true;
      this.fpick.anchor = { r, c };
      this.fpick.end = { r, c };
      this.fpick.input = this.formulaInput;
      this.fpick.insStart = formulaBarPos;
      this.fpick.insEnd = this.formulaInput.selectionEnd;
      this._applyFpick();
      return;
    }
    // Formula bar cursor at end or not focused → plain navigation

    if (this.isEditing && this.currentEditor) {
      const v = this.currentEditor.input.value;
      const curPos = this.currentEditor.input.selectionStart;
      if (v.startsWith('=') && curPos < v.length) {
        // Cursor not at end → range pick mode
        e.preventDefault();
        this.fpick.active = true;
        this.fpick.anchor = { r, c };
        this.fpick.end    = { r, c };
        this.fpick.input  = this.currentEditor.input;
        this.fpick.insStart = curPos;
        this.fpick.insEnd   = this.currentEditor.input.selectionEnd;
        this._applyFpick();
        return;
      }
      // Cursor at end → proceed with normal navigation
    }

    this._hideAutocomplete();
    if (this.isEditing) this._commitEdit();
    if (this.selectedImageId) { this.selectedImageId = null; this._renderImages(); }

    this.activeRow = r; this.activeCol = c;
    if (e.shiftKey) { this.selEnd = { r, c }; }
    else { this.selStart = { r, c }; this.selEnd = { r, c }; this.isSelectingRange = true; }
    this._updateSelection();
    this._updateFormulaBar();
    e.preventDefault();
    this.emit('select', { range: this._currentSelectionRef() });
  }

  _onCellMouseOver(e, r, c) {
    if (this.fpick.active) { this.fpick.end = { r, c }; this._applyFpick(); return; }
    if (!this.isSelectingRange) return;
    this.selEnd = { r, c };
    this._updateSelection();
  }

  _onMouseUp() {
    if (this.fpick.active) {
      this.fpick.active = false;
      if (this.fpick.input) {
        this.fpick.input.focus();
        this.fpick.input.setSelectionRange(this.fpick.insEnd, this.fpick.insEnd);
      }
      return;
    }
    this.isSelectingRange = false;
  }

  _applyFpick() {
    const { anchor, end, input, insStart } = this.fpick;
    if (!input || !anchor) return;
    const r1 = Math.min(anchor.r, end.r), r2 = Math.max(anchor.r, end.r);
    const c1 = Math.min(anchor.c, end.c), c2 = Math.max(anchor.c, end.c);
    const ref = (r1 === r2 && c1 === c2) ? makeRef(r1, c1) : `${makeRef(r1, c1)}:${makeRef(r2, c2)}`;
    const v = input.value;
    const newV = v.slice(0, insStart) + ref + v.slice(this.fpick.insEnd);
    input.value = newV;
    this.fpick.insEnd = insStart + ref.length;
    if (this.formulaInput) this.formulaInput.value = newV;

    this._clearFpickHighlight();
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++)
        this._getCell(r, c)?.classList.add('aix-fpick');
  }

  _clearFpickHighlight() {
    this.table.querySelectorAll('td.aix-fpick').forEach(td => td.classList.remove('aix-fpick'));
  }

  // ---- Formula reference highlighting (Excel-style colored outlines) ----

  _clearFormulaRefHighlight() {
    if (this.refBoxLayer) this.refBoxLayer.innerHTML = '';
  }

  _highlightFormulaRefs(value) {
    this._clearFormulaRefHighlight();
    if (!value || !value.startsWith('=')) return;
    const palette = ['#1f77b4', '#2ca02c', '#d62728', '#9467bd', '#ff7f0e', '#17becf', '#bcbd22'];
    const body = value.slice(1);
    const RANGE_RE = /(?:'([^']+)'!|([A-Za-z_][\w]*)!)?(\$?[A-Za-z]+\$?\d+)(?::(\$?[A-Za-z]+\$?\d+))?/g;
    const seen = new Set();
    let i = 0;
    let m;
    while ((m = RANGE_RE.exec(body))) {
      const sheetName = m[1] || m[2];
      if (sheetName && sheetName !== this.sheet.name) continue;
      const ref1 = m[3].replace(/\$/g, '');
      const ref2 = m[4] ? m[4].replace(/\$/g, '') : null;
      const key = ref2 ? `${ref1}:${ref2}` : ref1;
      if (seen.has(key)) continue;
      seen.add(key);
      const color = palette[i % palette.length];
      i++;
      try {
        const range = ref2 ? parseRange(`${ref1}:${ref2}`) : parseRange(ref1);
        this._drawRefBox(range, color);
      } catch {}
    }
  }

  _drawRefBox(range, color) {
    const tl = this._getCell(range.r1, range.c1);
    const br = this._getCell(range.r2, range.c2);
    if (!tl || !br) return;
    const top    = tl.offsetTop;
    const left   = tl.offsetLeft;
    const width  = br.offsetLeft + br.offsetWidth  - left;
    const height = br.offsetTop  + br.offsetHeight - top;
    const box = document.createElement('div');
    box.className = 'aix-fref-box';
    box.style.top    = top + 'px';
    box.style.left   = left + 'px';
    box.style.width  = width + 'px';
    box.style.height = height + 'px';
    box.style.borderColor = color;
    this.refBoxLayer.appendChild(box);
  }

  _updateSelection() {
    this.table.querySelectorAll('td').forEach(td => td.classList.remove('aix-in-selection', 'aix-active-cell'));
    this.table.querySelectorAll('.aix-col-header, .aix-row-header').forEach(h => h.classList.remove('aix-selected'));

    const r1 = Math.min(this.selStart.r, this.selEnd.r), r2 = Math.max(this.selStart.r, this.selEnd.r);
    const c1 = Math.min(this.selStart.c, this.selEnd.c), c2 = Math.max(this.selStart.c, this.selEnd.c);
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
      const td = this._getCell(r, c);
      if (td) td.classList.add(r === this.activeRow && c === this.activeCol ? 'aix-active-cell' : 'aix-in-selection');
    }
    const active = this._getCell(this.activeRow, this.activeCol);
    if (active) { active.classList.remove('aix-in-selection'); active.classList.add('aix-active-cell'); }

    const thead = this.table.tHead, tbody = this.table.tBodies[0];
    for (let c = c1; c <= c2; c++) thead?.rows[0]?.children[c + 1]?.classList.add('aix-selected');
    for (let r = r1; r <= r2; r++) tbody?.rows[r]?.children[0]?.classList.add('aix-selected');

    if (this.statusCell)
      this.statusCell.textContent = (r1 === r2 && c1 === c2) ? makeRef(r1, c1) : `${makeRef(r1,c1)}:${makeRef(r2,c2)}`;
    this._updateStatusCalc(r1, r2, c1, c2);
    this._updateFillHandle(r2, c2);
  }

  // ---- Fill handle (Excel-style autofill) ----

  _updateFillHandle(r2, c2) {
    this.table.querySelectorAll('.aix-fill-handle').forEach(el => el.remove());
    if (this.readOnly) return;
    const td = this._getCell(r2, c2);
    if (!td || td.style.display === 'none') return;
    const handle = document.createElement('div');
    handle.className = 'aix-fill-handle';
    handle.addEventListener('mousedown', (e) => this._startFillDrag(e));
    td.style.position = 'relative';
    td.appendChild(handle);
  }

  _startFillDrag(e) {
    e.preventDefault();
    e.stopPropagation();
    this._fillSrc = {
      r1: Math.min(this.selStart.r, this.selEnd.r),
      r2: Math.max(this.selStart.r, this.selEnd.r),
      c1: Math.min(this.selStart.c, this.selEnd.c),
      c2: Math.max(this.selStart.c, this.selEnd.c),
    };
    this._fillTarget = null;
    const move = (ev) => this._onFillMove(ev);
    const up   = () => {
      this._endFillDrag();
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  }

  _onFillMove(e) {
    const td = e.target.closest('td');
    if (!td) return;
    const r = Number(td.dataset.row), c = Number(td.dataset.col);
    if (isNaN(r) || isNaN(c)) return;
    const dRow = r > this._fillSrc.r2 ? r - this._fillSrc.r2 : 0;
    const dCol = c > this._fillSrc.c2 ? c - this._fillSrc.c2 : 0;
    let target;
    if (dRow === 0 && dCol === 0) { this._fillTarget = null; this._renderFillPreview(null); return; }
    if (dRow >= dCol) {
      target = { r1: this._fillSrc.r1, r2: Math.max(this._fillSrc.r2, r),
                 c1: this._fillSrc.c1, c2: this._fillSrc.c2 };
    } else {
      target = { r1: this._fillSrc.r1, r2: this._fillSrc.r2,
                 c1: this._fillSrc.c1, c2: Math.max(this._fillSrc.c2, c) };
    }
    this._fillTarget = target;
    this._renderFillPreview(target);
  }

  _renderFillPreview(target) {
    this.table.querySelectorAll('.aix-fill-preview').forEach(t => t.classList.remove('aix-fill-preview'));
    if (!target) return;
    for (let r = target.r1; r <= target.r2; r++) {
      for (let c = target.c1; c <= target.c2; c++) {
        if (r >= this._fillSrc.r1 && r <= this._fillSrc.r2 &&
            c >= this._fillSrc.c1 && c <= this._fillSrc.c2) continue;
        const td = this._getCell(r, c);
        if (td) td.classList.add('aix-fill-preview');
      }
    }
  }

  _endFillDrag() {
    this.table.querySelectorAll('.aix-fill-preview').forEach(t => t.classList.remove('aix-fill-preview'));
    if (!this._fillTarget) { this._fillSrc = null; return; }
    const t = this._fillTarget;
    const srcRef = makeRef(this._fillSrc.r1, this._fillSrc.c1) + ':' + makeRef(this._fillSrc.r2, this._fillSrc.c2);
    const tgtRef = makeRef(t.r1, t.c1) + ':' + makeRef(t.r2, t.c2);
    // Save undo for every target cell that will be overwritten
    for (let r = t.r1; r <= t.r2; r++) for (let c = t.c1; c <= t.c2; c++) this._saveUndo(r, c);
    this.sheet.fill(srcRef, tgtRef);
    this.selStart = { r: t.r1, c: t.c1 };
    this.selEnd   = { r: t.r2, c: t.c2 };
    this._renderAllCells();
    this._updateSelection();
    this.emit('change', { type: 'fill', src: srcRef, target: tgtRef });
    this._fillSrc = null;
    this._fillTarget = null;
  }

  _updateStatusCalc(r1, r2, c1, c2) {
    if (!this.statusCalc) return;
    let nums = [], count = 0;
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
      const v = this.sheet.value(makeRef(r, c));
      if (v !== '' && !isNaN(Number(v))) nums.push(Number(v));
      if (v !== '') count++;
    }
    if (nums.length > 1) {
      const sum = nums.reduce((a, b) => a + b, 0);
      this.statusCalc.textContent = `合計: ${sum.toLocaleString()}  平均: ${(sum/nums.length).toLocaleString()}  個数: ${count}`;
    } else this.statusCalc.textContent = '';
  }

  _currentSelectionRef() {
    const r1 = Math.min(this.selStart.r, this.selEnd.r), r2 = Math.max(this.selStart.r, this.selEnd.r);
    const c1 = Math.min(this.selStart.c, this.selEnd.c), c2 = Math.max(this.selStart.c, this.selEnd.c);
    return (r1 === r2 && c1 === c2) ? makeRef(r1, c1) : `${makeRef(r1,c1)}:${makeRef(r2,c2)}`;
  }

  _selectAll() {
    this.selStart = { r: 0, c: 0 };
    this.selEnd   = { r: this.numRows - 1, c: this.numCols - 1 };
    this.activeRow = 0; this.activeCol = 0;
    this._updateSelection();
  }
  _selectRow(r, e) {
    if (e.shiftKey) { this.selStart.c = 0; this.selEnd = { r, c: this.numCols - 1 }; }
    else { this.selStart = { r, c: 0 }; this.selEnd = { r, c: this.numCols - 1 }; }
    this.activeRow = r; this.activeCol = 0;
    this._updateSelection(); this._updateFormulaBar();
  }
  _selectColumn(c, e) {
    if (e.shiftKey) { this.selStart.r = 0; this.selEnd = { r: this.numRows - 1, c }; }
    else { this.selStart = { r: 0, c }; this.selEnd = { r: this.numRows - 1, c }; }
    this.activeRow = 0; this.activeCol = c;
    this._updateSelection(); this._updateFormulaBar();
  }

  // ============================================================
  // Keyboard
  // ============================================================
  _onKeyDown(e) {
    if (!this.root.contains(document.activeElement) && document.activeElement !== document.body) return;
    if (document.activeElement === this.formulaInput) return;
    if (this.isEditing) return;
    if (this.readOnly && e.key !== 'c' && !(e.ctrlKey || e.metaKey)) return;

    const k = e.key, ctrl = e.ctrlKey || e.metaKey;
    if (ctrl && k === 'b') { e.preventDefault(); this._toggleStyle('bold'); return; }
    if (ctrl && k === 'i') { e.preventDefault(); this._toggleStyle('italic'); return; }
    if (ctrl && k === 'u') { e.preventDefault(); this._toggleStyle('underline'); return; }
    if (ctrl && k === 'c') { e.preventDefault(); this._copySelection(); return; }
    if (ctrl && k === 'v') { e.preventDefault(); this._pasteSelection(); return; }
    if (ctrl && k === 'z') { e.preventDefault(); this._undo(); return; }
    if (ctrl && k === 'a') { e.preventDefault(); this._selectAll(); return; }

    const arrows = ['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Tab','Enter'];
    if (arrows.includes(k)) {
      e.preventDefault();
      const dr = k === 'ArrowDown' || k === 'Enter' ? 1 : k === 'ArrowUp' ? -1 : 0;
      const dc = k === 'ArrowRight'|| k === 'Tab'   ? 1 : k === 'ArrowLeft' ? -1 : 0;
      this._moveActive(dr, dc, e.shiftKey);
      return;
    }
    if (k === 'Delete' || k === 'Backspace') {
      e.preventDefault();
      if (this.selectedImageId && this._deleteSelectedImage()) return;
      this._clearSelection();
      return;
    }
    if (k === 'F2') { e.preventDefault(); this._startEdit(this.activeRow, this.activeCol); return; }
    if (k.length === 1 && !ctrl) this._startEdit(this.activeRow, this.activeCol, k);
  }

  _moveActive(dr, dc, shift) {
    const nr = Math.max(0, Math.min(this.numRows - 1, this.activeRow + dr));
    const nc = Math.max(0, Math.min(this.numCols - 1, this.activeCol + dc));
    this.activeRow = nr; this.activeCol = nc;
    if (shift) { this.selEnd = { r: nr, c: nc }; }
    else { this.selStart = { r: nr, c: nc }; this.selEnd = { r: nr, c: nc }; }
    this._updateSelection(); this._updateFormulaBar();
    this._getCell(nr, nc)?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  // ============================================================
  // Editing
  // ============================================================
  _startEdit(r, c, initial = null) {
    if (this.readOnly) return;
    if (this.isEditing) this._commitEdit();
    this.isEditing = true;
    this.activeRow = r; this.activeCol = c;
    this.selStart = { r, c }; this.selEnd = { r, c };
    this._updateSelection();

    const td = this._getCell(r, c); if (!td) return;
    const ref = makeRef(r, c);
    const cell = this.sheet.cells[ref];
    const raw = cell?.f ?? (cell?.v !== undefined ? formatCellValue(cell.v, cell?.s?.numFmt) : '');

    const input = document.createElement('input');
    input.className = 'aix-cell-editor';
    input.value = initial !== null ? initial : raw;
    const s = cell?.s || {};
    input.style.fontWeight = s.bold ? 'bold' : '';
    input.style.fontStyle  = s.italic ? 'italic' : '';
    input.style.fontSize   = s.fontSize ? s.fontSize + 'px' : '';
    input.style.textAlign  = s.align || '';

    td.textContent = '';
    td.appendChild(input);
    input.focus();
    // Place cursor at end (matches Excel double-click / F2 behavior)
    input.setSelectionRange(input.value.length, input.value.length);

    input.addEventListener('keydown', (e) => this._onEditorKeyDown(e));
    input.addEventListener('input', () => {
      if (this.formulaInput) this.formulaInput.value = input.value;
      this._checkAutocomplete(input);
      this._clearFpickHighlight();
      this._highlightFormulaRefs(input.value);
    });
    this.currentEditor = { input, r, c };
    if (this.formulaInput) this.formulaInput.value = input.value;
    this._checkAutocomplete(input);
    this._highlightFormulaRefs(input.value);
    this.emit('edit-start', { ref });
  }

  _onEditorKeyDown(e) {
    if (this.ac.items.length > 0) {
      if (e.key === 'ArrowDown') { e.preventDefault(); this._acMove(1); return; }
      if (e.key === 'ArrowUp')   { e.preventDefault(); this._acMove(-1); return; }
      if ((e.key === 'Tab' || e.key === 'Enter') && this.ac.index >= 0) {
        e.preventDefault(); this._acConfirm(); return;
      }
      if (e.key === 'Escape') { e.preventDefault(); this._hideAutocomplete(); return; }
    }
    if (e.key === 'Enter') { e.preventDefault(); this._clearFpickHighlight(); this._commitEdit(); this._moveActive(1, 0, false); }
    else if (e.key === 'Tab') { e.preventDefault(); this._clearFpickHighlight(); this._commitEdit(); this._moveActive(0, 1, false); }
    else if (e.key === 'Escape') { this._clearFpickHighlight(); this._hideAutocomplete(); this._cancelEdit(); }
  }

  _commitEdit() {
    if (!this.isEditing || !this.currentEditor) return;
    const { input, r, c } = this.currentEditor;
    const ref = makeRef(r, c);
    const oldRaw = this.sheet.cells[ref] ? (this.sheet.cells[ref].f ?? this.sheet.cells[ref].v) : '';
    this._saveUndo(r, c);

    const raw = input.value;
    if (raw === '') this.sheet.clear(ref);
    else if (raw.startsWith('=')) this.sheet.write(ref, raw);
    else {
      // Try to coerce: number → number, date-looking → Date, else string
      const dateMatch = raw.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/) || raw.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
      if (dateMatch) {
        const today = new Date();
        const y = dateMatch.length === 4 ? +dateMatch[1] : today.getFullYear();
        const mo = dateMatch.length === 4 ? +dateMatch[2] : +dateMatch[1];
        const d = dateMatch.length === 4 ? +dateMatch[3] : +dateMatch[2];
        this.sheet.write(ref, new Date(y, mo - 1, d));
      } else if (!isNaN(Number(raw)) && raw.trim() !== '') {
        this.sheet.write(ref, Number(raw));
      } else {
        this.sheet.write(ref, raw);
      }
    }

    this.isEditing = false;
    this.currentEditor = null;
    this._hideAutocomplete();
    this._clearFormulaRefHighlight();
    this._renderAllCells();
    this._updateFormulaBar();
    this.emit('change', { ref, oldValue: oldRaw, newValue: raw });
    this.emit('edit-end', { ref });
  }

  _cancelEdit() {
    if (!this.isEditing || !this.currentEditor) return;
    const { r, c } = this.currentEditor;
    this.isEditing = false;
    this.currentEditor = null;
    const td = this._getCell(r, c);
    if (td) this._renderCell(td, r, c);
    this.emit('edit-end', { ref: makeRef(r, c) });
  }

  // ============================================================
  // Formula bar
  // ============================================================
  _updateFormulaBar() {
    if (!this.formulaInput) return;
    this.cellRefEl.value = makeRef(this.activeRow, this.activeCol);
    const cell = this.sheet.cells[makeRef(this.activeRow, this.activeCol)];
    this.formulaInput.value = cell ? (cell.f ?? (cell.v !== undefined ? String(cell.v) : '')) : '';
  }
  _onFormulaBarInput() {
    this._checkAutocomplete(this.formulaInput);
    if (this.isEditing && this.currentEditor) this.currentEditor.input.value = this.formulaInput.value;
    this._highlightFormulaRefs(this.formulaInput.value);
  }
  _onFormulaBarKeyDown(e) {
    if (this.ac.items.length > 0) {
      if (e.key === 'ArrowDown') { e.preventDefault(); this._acMove(1); return; }
      if (e.key === 'ArrowUp')   { e.preventDefault(); this._acMove(-1); return; }
      if ((e.key === 'Tab' || e.key === 'Enter') && this.ac.index >= 0) { e.preventDefault(); this._acConfirm(); return; }
      if (e.key === 'Escape') { e.preventDefault(); this._hideAutocomplete(); return; }
    }
    if (e.key === 'Enter') {
      e.preventDefault();
      this._hideAutocomplete();
      const ref = makeRef(this.activeRow, this.activeCol);
      const old = this.sheet.cells[ref];
      this._saveUndo(this.activeRow, this.activeCol);
      const v = this.formulaInput.value;
      if (v === '') this.sheet.clear(ref);
      else this.sheet.write(ref, v.startsWith('=') ? v : (isNaN(Number(v)) || v.trim() === '' ? v : Number(v)));
      this._renderAllCells();
      this.formulaInput.blur();
      this._moveActive(1, 0, false);
      this.emit('change', { ref, oldValue: old, newValue: v });
    } else if (e.key === 'Escape') { this._hideAutocomplete(); this._updateFormulaBar(); this.formulaInput.blur(); }
  }

  // ============================================================
  // Autocomplete
  // ============================================================
  _checkAutocomplete(inputEl) {
    const val = inputEl.value, pos = inputEl.selectionStart;
    if (!val.startsWith('=')) { this._hideAutocomplete(); return; }
    const before = val.slice(0, pos);
    const m = before.match(/(?:^|[^A-Za-z])([A-Za-z]{1,})$/);
    if (!m) { this._hideAutocomplete(); return; }
    const word = m[1].toUpperCase();
    const wordStart = before.lastIndexOf(m[1]);
    const matches = FUNCTIONS.filter(f => f.name.startsWith(word) && f.name !== word);
    if (matches.length === 0) { this._hideAutocomplete(); return; }
    this.ac = { items: matches, index: -1, inputEl, wordStart };

    let ref = inputEl;
    if (this.currentEditor) {
      const td = this._getCell(this.currentEditor.r, this.currentEditor.c);
      if (td) ref = td;
    }
    const rect = ref.getBoundingClientRect();
    this.acEl.style.left = rect.left + 'px';
    this.acEl.style.top  = rect.bottom + 'px';
    this.acEl.style.display = 'block';
    this.acEl.innerHTML = matches.map((f, i) =>
      `<div class="aix-ac-item" data-idx="${i}"><span class="aix-ac-name">${f.name}</span><span class="aix-ac-desc">${f.desc}</span></div>`
    ).join('');
    this.acEl.querySelectorAll('.aix-ac-item').forEach((el, i) => {
      el.addEventListener('mousedown', () => { this.ac.index = i; this._acConfirm(); });
    });
  }
  _hideAutocomplete() { this.ac.items = []; this.ac.index = -1; this.acEl.style.display = 'none'; }
  _acMove(dir) {
    this.ac.index = Math.max(-1, Math.min(this.ac.items.length - 1, this.ac.index + dir));
    this.acEl.querySelectorAll('.aix-ac-item').forEach((el, i) => {
      el.classList.toggle('aix-ac-active', i === this.ac.index);
      if (i === this.ac.index) el.scrollIntoView({ block: 'nearest' });
    });
  }
  _acConfirm() {
    if (this.ac.index < 0) return;
    const fn = this.ac.items[this.ac.index];
    const inp = this.ac.inputEl;
    const val = inp.value, pos = inp.selectionStart;
    const newV = val.slice(0, this.ac.wordStart) + fn.name + '(' + val.slice(pos);
    inp.value = newV;
    const np = this.ac.wordStart + fn.name.length + 1;
    inp.setSelectionRange(np, np);
    inp.focus();
    if (this.formulaInput) this.formulaInput.value = newV;
    if (this.currentEditor) this.currentEditor.input.value = newV;
    this._hideAutocomplete();
  }

  // ============================================================
  // Formatting / Undo / Copy / Paste / Clear
  // ============================================================
  _applyFormat(prop, val) {
    if (this.readOnly) return;
    const ref = this._currentSelectionRef();
    this.sheet.style(ref, { [prop]: val });
    this._renderAllCells();
  }
  _toggleStyle(prop) {
    const cur = this.sheet.cells[makeRef(this.activeRow, this.activeCol)]?.s?.[prop];
    this._applyFormat(prop, !cur);
  }
  _saveUndo(r, c) {
    const ref = makeRef(r, c);
    const cell = this.sheet.cells[ref];
    this.undoStack.push({ ref, cell: cell ? JSON.parse(JSON.stringify(cell)) : null });
    if (this.undoStack.length > 100) this.undoStack.shift();
  }
  _undo() {
    if (!this.undoStack.length) return;
    const { ref, cell } = this.undoStack.pop();
    if (cell === null) delete this.sheet.cells[ref];
    else this.sheet.cells[ref] = cell;
    this._renderAllCells();
    this._updateFormulaBar();
  }
  _copySelection() {
    const r1 = Math.min(this.selStart.r, this.selEnd.r), r2 = Math.max(this.selStart.r, this.selEnd.r);
    const c1 = Math.min(this.selStart.c, this.selEnd.c), c2 = Math.max(this.selStart.c, this.selEnd.c);
    this.clipboard = { r1, c1, data: [] };
    for (let r = r1; r <= r2; r++) {
      const row = [];
      for (let c = c1; c <= c2; c++) {
        const cell = this.sheet.cells[makeRef(r, c)];
        row.push(cell ? JSON.parse(JSON.stringify(cell)) : null);
      }
      this.clipboard.data.push(row);
    }
    if (this.statusInfo) {
      this.statusInfo.textContent = '📋 コピー完了';
      setTimeout(() => this.statusInfo.textContent = '準備完了', 1500);
    }
  }
  _pasteSelection() {
    if (!this.clipboard) return;
    const { data } = this.clipboard;
    for (let r = 0; r < data.length; r++) for (let c = 0; c < data[r].length; c++) {
      const nr = this.activeRow + r, nc = this.activeCol + c;
      if (nr >= this.numRows || nc >= this.numCols) continue;
      const ref = makeRef(nr, nc);
      this._saveUndo(nr, nc);
      if (data[r][c] === null) delete this.sheet.cells[ref];
      else this.sheet.cells[ref] = JSON.parse(JSON.stringify(data[r][c]));
    }
    this._renderAllCells();
  }
  _clearSelection() {
    if (this.readOnly) return;
    const r1 = Math.min(this.selStart.r, this.selEnd.r), r2 = Math.max(this.selStart.r, this.selEnd.r);
    const c1 = Math.min(this.selStart.c, this.selEnd.c), c2 = Math.max(this.selStart.c, this.selEnd.c);
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
      const ref = makeRef(r, c);
      this._saveUndo(r, c);
      delete this.sheet.cells[ref];
    }
    this._renderAllCells();
  }

  // ============================================================
  // Row / Column operations
  // ============================================================
  _insertRow(rowNum, count = 1) {
    if (this.readOnly) return;
    this.sheet.insertRow(rowNum, count);
    this.numRows += count;
    this._buildTable();
  }
  _deleteRow() {
    if (this.readOnly) return;
    const r1 = Math.min(this.selStart.r, this.selEnd.r), r2 = Math.max(this.selStart.r, this.selEnd.r);
    if (this.numRows - (r2 - r1 + 1) < 1) return;
    this.sheet.deleteRow(r1 + 1, r2 - r1 + 1);
    this.numRows -= r2 - r1 + 1;
    this.activeRow = Math.min(this.activeRow, this.numRows - 1);
    this.selStart = { r: this.activeRow, c: this.activeCol }; this.selEnd = { ...this.selStart };
    this._buildTable(); this._updateFormulaBar();
  }
  _insertCol(colIdx, count = 1) {
    if (this.readOnly) return;
    this.sheet.insertCol(idxToCol(colIdx - 1), count);
    for (let i = 0; i < count; i++) this.colWidths.splice(colIdx - 1, 0, 100);
    this.numCols += count;
    this._buildTable();
  }
  _deleteCol() {
    if (this.readOnly) return;
    const c1 = Math.min(this.selStart.c, this.selEnd.c), c2 = Math.max(this.selStart.c, this.selEnd.c);
    if (this.numCols - (c2 - c1 + 1) < 1) return;
    this.sheet.deleteCol(idxToCol(c1), c2 - c1 + 1);
    this.colWidths.splice(c1, c2 - c1 + 1);
    this.numCols -= c2 - c1 + 1;
    this.activeCol = Math.min(this.activeCol, this.numCols - 1);
    this.selStart = { r: this.activeRow, c: this.activeCol }; this.selEnd = { ...this.selStart };
    this._buildTable(); this._updateFormulaBar();
  }

  // ============================================================
  // Context menu
  // ============================================================
  _showContextMenu(e, r, c) {
    e.preventDefault();
    this.ctxRow = r; this.ctxCol = c;
    const inSel = r >= Math.min(this.selStart.r, this.selEnd.r) && r <= Math.max(this.selStart.r, this.selEnd.r);
    if (!inSel) { this.activeRow = r; this.activeCol = c; this.selStart = { r, c }; this.selEnd = { r, c }; this._updateSelection(); }
    this.contextMenu.style.display = 'block';
    this.contextMenu.style.left = e.clientX + 'px';
    this.contextMenu.style.top  = e.clientY + 'px';
  }
  _hideContextMenu() { this.contextMenu.style.display = 'none'; }

  // ============================================================
  // Column resize
  // ============================================================
  _startColResize(e, c) {
    e.preventDefault(); e.stopPropagation();
    const start = { x: e.clientX, w: this.colWidths[c] };
    const move = (ev) => {
      this.colWidths[c] = Math.max(30, start.w + (ev.clientX - start.x));
      const w = this.colWidths[c] + 'px';
      const th = this.table.tHead?.rows[0]?.children[c + 1];
      if (th) { th.style.width = w; th.style.minWidth = w; }
      for (let r = 0; r < this.numRows; r++) {
        const td = this._getCell(r, c);
        if (td) { td.style.width = w; td.style.minWidth = w; }
      }
    };
    const up = () => {
      this.sheet.colWidth(idxToCol(c), this.colWidths[c]);
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  }

  // ============================================================
  // File I/O (browser)
  // ============================================================
  _downloadAIX() {
    const data = this.workbook.sheets.length > 1 ? this.workbook.toJSON() : this.sheet.toJSON();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const baseName = this.workbook.sheets.length > 1 ? 'workbook' : (this.sheet.name || 'sheet');
    this._download(blob, baseName + '.aix.json');
  }
  _downloadCSV() {
    const blob = new Blob(['﻿' + this.sheet.toCSV()], { type: 'text/csv;charset=utf-8' });
    this._download(blob, (this.sheet.name || 'sheet') + '.csv');
  }
  async _downloadXLSX() {
    const btn = this.toolbar?.querySelector('[data-act="save-xlsx"]');
    const restore = btn ? (() => { btn.disabled = false; btn.textContent = '📊 Excel 保存'; }) : null;
    if (btn) { btn.disabled = true; btn.textContent = '⏳ 準備中…'; }
    try {
      await this._ensureXLSXLibs();
      const buf = await (this.workbook.sheets.length > 1
        ? this.workbook.toXLSX()
        : this.sheet.toXLSX());
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const baseName = this.workbook.sheets.length > 1 ? 'workbook' : (this.sheet.name || 'sheet');
      this._download(blob, baseName + '.xlsx');
    } catch (e) {
      alert('Excel 出力に失敗: ' + (e?.message || e));
      console.error('[aix-sheet] xlsx export failed', e);
    } finally {
      restore?.();
    }
  }

  /** Lazy-load xlsx-js-style (community fork with style-write support) + JSZip from CDN. */
  _ensureXLSXLibs() {
    const load = (src) => new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src; s.async = true;
      s.onload = resolve;
      s.onerror = () => reject(new Error('CDN 読み込み失敗: ' + src));
      document.head.appendChild(s);
    });
    const tasks = [];
    if (typeof window.XLSX === 'undefined') {
      // xlsx-js-style: drop-in fork of SheetJS that actually writes cell styles
      // (bgColor / border / numFmt / font etc.). Original SheetJS Community
      // edition silently drops them on write.
      tasks.push(load('https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js'));
    }
    if (typeof window.JSZip === 'undefined') {
      tasks.push(load('https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js'));
    }
    return Promise.all(tasks);
  }
  _download(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name; a.click();
    URL.revokeObjectURL(url);
  }
  _handleFileLoad(file, kind) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        if (kind === 'csv') {
          this.setSheet(Sheet.fromCSV(e.target.result));
        } else {
          // Auto-detect workbook vs single sheet
          const parsed = JSON.parse(e.target.result);
          if (parsed.type === 'aix-workbook') this.setWorkbook(Workbook.fromJSON(parsed));
          else this.setSheet(Sheet.fromJSON(parsed));
        }
        if (this.statusInfo) { this.statusInfo.textContent = `📂 ${kind.toUpperCase()} 読込完了`; setTimeout(() => this.statusInfo.textContent = '準備完了', 1500); }
      } catch (err) { alert('読み込みエラー: ' + err.message); }
    };
    reader.readAsText(file, 'UTF-8');
  }

  /** Replace the underlying Workbook. */
  setWorkbook(wb) {
    this.workbook = wb;
    this.sheet = wb.active;
    const u = this.sheet.usedRange();
    if (u.rows > this.numRows) this.numRows = u.rows + 10;
    if (u.cols > this.numCols) this.numCols = u.cols + 5;
    this.colWidths = new Array(this.numCols).fill(100);
    for (const [label, w] of Object.entries(this.sheet.cols)) {
      const ci = colToIdx(label);
      if (ci < this.numCols) this.colWidths[ci] = w;
    }
    this._buildTable(); this._renderTabs(); this._updateSelection(); this._updateFormulaBar();
  }

  /** Get the underlying Workbook. */
  getWorkbook() { return this.workbook; }
}

export default SheetView;

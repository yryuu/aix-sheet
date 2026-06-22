/**
 * aix-sheet SDK — TypeScript definitions
 * AI-friendly spreadsheet manipulation.
 *
 * Cell references use A1 notation: "A1", "B12", "AA3".
 * Ranges use colon notation: "A1:C5".
 */

export type BorderStyleName = 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'double';

export interface BorderSide {
  style: BorderStyleName;
  color?: string;
}

/**
 * Border specification.
 * - true / 'all'                          → all 4 sides, thin black
 * - false / null                          → no borders
 * - { top, right, bottom, left }          → per-side; each: true|false|'thin'|...|{style,color}
 * - { all: ... }                          → shorthand for all 4 sides
 */
export type BorderSpec =
  | boolean
  | 'all'
  | null
  | Partial<Record<'top' | 'right' | 'bottom' | 'left' | 'all',
        boolean | BorderStyleName | BorderSide>>;

export interface CellStyle {
  /** Bold text */
  bold?: boolean;
  /** Italic text */
  italic?: boolean;
  /** Underline text */
  underline?: boolean;
  /** Text color (CSS color string, e.g. "#ff0000" or "red") */
  color?: string;
  /** Background color (CSS color string) */
  bgColor?: string;
  /** Horizontal alignment */
  align?: 'left' | 'center' | 'right';
  /** Font size in pixels */
  fontSize?: number;
  /** Font family name */
  fontFamily?: string;
  /** Cell borders. Accepts shortcuts (true, 'all') or per-side objects. */
  border?: BorderSpec;
}

/** Merged-range descriptor (stored internally as 0-based indices). */
export interface MergeRange {
  r1: number; c1: number; r2: number; c2: number;
}

/** Conditional formatting rule. */
export interface CFRule {
  /** A1 range, e.g. "G3:Z11". Rule applies to each cell in this range. */
  range: string;
  /**
   * Excel-style formula starting with "=". Relative refs (D3) shift per cell;
   * absolute refs ($D$3, $D3, D$3) lock the column and/or row.
   * ROW() / COLUMN() resolve to the 1-based row/col of the cell being evaluated.
   * Comparison ops use Excel syntax: '=' (equal), '<>' (not equal).
   */
  formula: string;
  /** Style overrides applied when the formula evaluates to truthy. */
  style: Pick<CellStyle, 'bgColor' | 'color' | 'bold' | 'italic'>;
}

export interface CellRecord {
  /** Raw value (when not a formula) */
  v?: string | number | boolean;
  /** Formula string starting with "=" */
  f?: string;
  /** Style applied to this cell */
  s?: CellStyle;
}

export interface ImageRecord {
  id: string;
  /** data URL ("data:image/png;base64,...") or http(s) URL */
  src: string;
  /** anchor cell, e.g. "B2" */
  anchor: string;
  offset: { x: number; y: number };
  size:   { width: number; height: number };
}

export interface AddImageOptions {
  id?: string;
  offset?: { x?: number; y?: number };
  size?:   { width?: number; height?: number };
}

export interface AixSheetData {
  version: string;
  type: 'aix-sheet';
  name: string;
  cells: Record<string, CellRecord>;
  cols: Record<string, number>;
  rows: Record<number, number>;
  images?: ImageRecord[];
  /** Merged ranges, serialized as A1 strings ("A1:D1"). */
  merges?: string[];
  /** Conditional-formatting rules. */
  cfs?: CFRule[];
}

export interface AixWorkbookData {
  version: string;
  type: 'aix-workbook';
  activeSheet: number;
  sheets: Array<Omit<AixSheetData, 'type' | 'version'>>;
}

export type CellValue = string | number | boolean | null;

export class Sheet {
  name: string;
  cells: Record<string, CellRecord>;
  cols: Record<string, number>;
  rows: Record<number, number>;
  images: ImageRecord[];
  merges: MergeRange[];
  cfs: CFRule[];
  version: string;

  constructor(name?: string);

  /** Parse from a .aix.json string or parsed object. */
  static fromJSON(input: string | AixSheetData): Sheet;

  /** Load from a .aix.json file path. Node.js only. */
  static load(path: string): Sheet;

  /** Build a Sheet from CSV text. */
  static fromCSV(text: string, name?: string): Sheet;

  /**
   * Write a value, formula, or array to a cell or range.
   *
   *   write("A1", 42)
   *   write("A1", "Hello")
   *   write("A1", "=SUM(B1:B10)")              // formula
   *   write("A1:A5", [1,2,3,4,5])               // 1D → column
   *   write("A1:C1", [1,2,3])                   // 1D → row
   *   write("A1:C2", [[1,2,3],[4,5,6]])         // 2D
   */
  write(ref: string, value: CellValue | CellValue[] | CellValue[][]): this;

  /** Returns the raw cell record { v, f, s }. Empty cell returns {}. */
  get(ref: string): CellRecord;

  /** Returns the computed value (evaluates formulas). Range returns 2D array. */
  value(ref: string): CellValue | CellValue[][];

  /** Alias for value(). */
  read(ref: string): CellValue | CellValue[][];

  /** Apply style to a cell or range. Merges with existing style. */
  style(ref: string, opts: CellStyle): this;

  /** Clear values and styles in a range. */
  clear(ref: string): this;

  /** Set column width in pixels. col is letter ("A") or 0-based index. */
  colWidth(col: string | number, px: number): this;

  /** Set row height in pixels. row is 1-based. */
  rowHeight(row: number, px: number): this;

  /** Insert `count` blank rows above the given 1-based row number. */
  insertRow(rowNum: number, count?: number): this;
  /** Delete `count` rows starting at the given 1-based row number. */
  deleteRow(rowNum: number, count?: number): this;
  /** Insert `count` blank columns to the left of the given column label. */
  insertCol(colLabel: string | number, count?: number): this;
  /** Delete `count` columns starting at the given column label. */
  deleteCol(colLabel: string | number, count?: number): this;

  /**
   * Embed an image at a cell anchor.
   * @param anchor - top-left cell ref like "B2"
   * @param source - data URL or http(s) URL
   * @returns the image id
   */
  addImage(anchor: string, source: string, options?: AddImageOptions): string;
  getImage(id: string): ImageRecord | null;
  removeImage(id: string): this;
  moveImage(id: string, newAnchor: string, offset?: { x?: number; y?: number }): this;
  resizeImage(id: string, size: { width?: number; height?: number }): this;

  /**
   * Merge a range of cells. The top-left cell keeps its value; covered cells
   * are hidden in the viewer and excluded from xlsx output.
   * Throws if the range overlaps an existing merge.
   */
  merge(ref: string): this;
  /** Remove a merge by exact range match, or any merge containing a cell ref. */
  unmerge(ref: string): this;
  /** Find the merge containing (row, col), or null. */
  mergeAt(row: number, col: number): MergeRange | null;

  /**
   * Add a conditional-formatting rule. The viewer re-evaluates rules whenever
   * cells change; xlsx export emits matching <conditionalFormatting> so Excel
   * does the same on open.
   */
  cf(range: string, opts: { formula: string; style: CFRule['style'] }): this;
  /** Remove a CF rule by exact range, or clear all if no range passed. */
  clearCF(range?: string): this;
  /** Returns the merged style overrides from all matching CF rules, or null. */
  cfStyleAt(row: number, col: number): CFRule['style'] | null;

  /** Bounding box of used cells. */
  usedRange(): { rows: number; cols: number };

  /** Render the sheet as a Markdown table (for LLM context). */
  toMarkdown(opts?: { maxRows?: number }): string;

  /** Serialize to plain object (the .aix.json format). */
  toJSON(): AixSheetData;

  /** Write the sheet to a .aix.json file. Node.js only. */
  save(path: string): this;

  /** Export to CSV string. */
  toCSV(): string;

  /** Export to .xlsx. Requires SheetJS to be loaded. Returns ArrayBuffer (browser) or Buffer (Node). */
  toXLSX(): ArrayBuffer | Buffer;
}

/**
 * Container for multiple Sheets.
 * Cross-sheet formula references use `SheetName!A1` notation,
 * or `'Sheet Name'!A1` when the name contains spaces.
 */
export class Workbook {
  version: string;
  sheets: Sheet[];
  activeIndex: number;
  /** Currently active sheet. */
  active: Sheet;

  constructor();

  /** Parse from `.aix.json`. Accepts both workbook and single-sheet formats. */
  static fromJSON(input: string | AixWorkbookData | AixSheetData): Workbook;

  /** Load from a `.aix.json` file. Node.js only. */
  static load(path: string): Promise<Workbook>;

  /** Add a new empty sheet. Auto-generates a unique name if omitted. */
  addSheet(name?: string): Sheet;

  /** Get a sheet by name or 0-based index. Returns null if missing. */
  sheet(nameOrIndex: string | number): Sheet | null;

  removeSheet(nameOrIndex: string | number): this;
  renameSheet(oldName: string, newName: string): this;
  moveSheet(fromIndex: number, toIndex: number): this;

  toJSON(): AixWorkbookData;
  save(path: string): Promise<this>;
  toXLSX(): Promise<ArrayBuffer | Buffer>;
}

/** Helpers */
export function parseRef(ref: string): { row: number; col: number; label: string };
export function parseRange(ref: string): { r1: number; r2: number; c1: number; c2: number };
export function makeRef(row: number, col: number): string;
export function idxToCol(c: number): string;
export function colToIdx(label: string): number;

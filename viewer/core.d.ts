import type { Sheet, Workbook, AixSheetData, AixWorkbookData } from '../sdk/sheet';

export interface SheetViewOptions {
  /** Multi-sheet workbook (preferred). */
  workbook?: Workbook;
  /** Single-sheet shortcut. Wrapped in an implicit Workbook. */
  sheet?: Sheet;
  /** Visible row count (default: 50). */
  rows?: number;
  /** Visible column count (default: 26). */
  cols?: number;
  /** If true, disables editing. */
  readOnly?: boolean;
  /** Show toolbar (default: true). */
  toolbar?: boolean;
  /** Show formula bar (default: true). */
  formulaBar?: boolean;
  /** Show status bar (default: true). */
  statusBar?: boolean;
  /** Show sheet tabs (default: true). */
  tabs?: boolean;
}

/** Fired on a cell value/formula edit (commit). */
export interface CellChangeEvent {
  ref: string;
  oldValue: unknown;
  newValue: unknown;
}
/**
 * Fired on structural / formatting operations performed via the toolbar
 * or programmatic ops. The `type` discriminates which operation.
 *   - 'merge' / 'unmerge'   — cell merge changes
 *   - 'border'              — border style applied to a range
 *   - 'numfmt'              — number format applied
 *   - 'cf-add' / 'cf-delete'— conditional-formatting rule mutated
 *   - 'fill'                — autofill drag completed (carries src / target)
 */
export interface OperationChangeEvent {
  type: 'merge' | 'unmerge' | 'border' | 'numfmt' | 'cf-add' | 'cf-delete' | 'fill';
  src?: string;
  target?: string;
}
export type ChangeEvent = CellChangeEvent | OperationChangeEvent;

export interface SelectEvent { range: string }
export interface EditEvent   { ref: string }
export interface SheetChangeEvent { name: string; index: number }
export interface ImageChangeEvent { id: string }

export type SheetViewEvent =
  | 'change' | 'select' | 'edit-start' | 'edit-end' | 'sheet-change' | 'image-change';

export class SheetView {
  constructor(container: string | HTMLElement, options?: SheetViewOptions);

  refresh(): void;
  setSheet(sheet: Sheet): void;
  setWorkbook(workbook: Workbook): void;
  getSheet(): Sheet;
  getWorkbook(): Workbook;
  switchSheet(nameOrIndex: string | number): void;
  addSheet(name?: string): Sheet;
  focusCell(ref: string): void;
  /** Latest .aix.json as a plain object. */
  toJSON(): AixWorkbookData | AixSheetData;
  /** Same as toJSON() but stringified. */
  toAixJson(pretty?: boolean): string;
  /** xlsx bytes; requires xlsx-js-style + jszip. */
  toXLSX(): Promise<ArrayBuffer | Buffer>;
  destroy(): void;

  on(event: 'change',       handler: (e: ChangeEvent)      => void): this;
  on(event: 'select',       handler: (e: SelectEvent)      => void): this;
  on(event: 'edit-start',   handler: (e: EditEvent)        => void): this;
  on(event: 'edit-end',     handler: (e: EditEvent)        => void): this;
  on(event: 'sheet-change', handler: (e: SheetChangeEvent) => void): this;
  on(event: 'image-change', handler: (e: ImageChangeEvent) => void): this;
  off(event: SheetViewEvent, handler: (...args: any[]) => void): this;
}

export default SheetView;

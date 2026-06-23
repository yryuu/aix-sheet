import type { CSSProperties } from 'react';
import type { Sheet, Workbook, AixSheetData, AixWorkbookData } from '../sdk/sheet';
import type { SheetView, SheetViewOptions, ChangeEvent, SelectEvent, EditEvent, SheetChangeEvent, ImageChangeEvent } from './core';

export interface SheetViewReactProps extends Omit<SheetViewOptions, 'sheet' | 'workbook'> {
  /** Multi-sheet workbook. Changing identity remounts the view. */
  workbook?: Workbook;
  /** Single-sheet shortcut. */
  sheet?: Sheet;
  onChange?:      (e: ChangeEvent)      => void;
  onSelect?:      (e: SelectEvent)      => void;
  onEditStart?:   (e: EditEvent)        => void;
  onEditEnd?:     (e: EditEvent)        => void;
  onSheetChange?: (e: SheetChangeEvent) => void;
  onImageChange?: (e: ImageChangeEvent) => void;
  style?: CSSProperties;
  className?: string;
}

export interface SheetViewReactHandle {
  refresh: () => void;
  getSheet: () => Sheet | undefined;
  getWorkbook: () => Workbook | undefined;
  setSheet: (s: Sheet) => void;
  setWorkbook: (w: Workbook) => void;
  switchSheet: (nameOrIndex: string | number) => void;
  addSheet: (name?: string) => Sheet | undefined;
  focusCell: (ref: string) => void;
  /** Latest .aix.json as a plain object (workbook for multi-sheet, sheet otherwise). */
  toJSON: () => AixWorkbookData | AixSheetData | undefined;
  /** Same as toJSON() but stringified. */
  toAixJson: (pretty?: boolean) => string | undefined;
  /** xlsx bytes (ArrayBuffer in browser, Buffer in Node). */
  toXLSX: () => Promise<ArrayBuffer | Buffer> | undefined;
  view: () => SheetView | null;
}

export const SheetViewReact: React.ForwardRefExoticComponent<
  SheetViewReactProps & React.RefAttributes<SheetViewReactHandle>
>;

export default SheetViewReact;

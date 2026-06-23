/**
 * React wrapper for SheetView.
 *
 *   import { SheetViewReact } from 'aix-sheet/react';
 *   import 'aix-sheet/viewer/styles.css';
 *
 *   <SheetViewReact
 *     sheet={mySheet}
 *     onChange={({ ref, newValue }) => ...}
 *     style={{ height: 500 }}
 *   />
 *
 * Requires React 17+ as a peer dependency.
 */
import React, { useRef, useEffect, useImperativeHandle, forwardRef } from 'react';
import { SheetView } from './core.js';

export const SheetViewReact = forwardRef(function SheetViewReact(props, ref) {
  const {
    workbook, sheet, rows, cols, readOnly,
    toolbar, formulaBar, statusBar, tabs,
    onChange, onSelect, onEditStart, onEditEnd, onSheetChange, onImageChange,
    style, className, ...rest
  } = props;

  const hostRef = useRef(null);
  const viewRef = useRef(null);

  useEffect(() => {
    const view = new SheetView(hostRef.current, {
      workbook, sheet, rows, cols, readOnly,
      toolbar, formulaBar, statusBar, tabs
    });
    viewRef.current = view;
    if (onChange)      view.on('change', onChange);
    if (onSelect)      view.on('select', onSelect);
    if (onEditStart)   view.on('edit-start', onEditStart);
    if (onEditEnd)     view.on('edit-end', onEditEnd);
    if (onSheetChange) view.on('sheet-change', onSheetChange);
    if (onImageChange) view.on('image-change', onImageChange);
    return () => view.destroy();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [workbook, sheet]);

  useImperativeHandle(ref, () => ({
    refresh:     () => viewRef.current?.refresh(),
    getSheet:    () => viewRef.current?.getSheet(),
    getWorkbook: () => viewRef.current?.getWorkbook(),
    setSheet:    (s) => viewRef.current?.setSheet(s),
    setWorkbook: (w) => viewRef.current?.setWorkbook(w),
    switchSheet: (n) => viewRef.current?.switchSheet(n),
    addSheet:    (n) => viewRef.current?.addSheet(n),
    focusCell:   (r) => viewRef.current?.focusCell(r),
    /** Latest .aix.json as a plain object (workbook or single sheet). */
    toJSON:      () => viewRef.current?.toJSON(),
    /** Same as toJSON() but stringified. */
    toAixJson:   (pretty) => viewRef.current?.toAixJson(pretty),
    /** Promise<ArrayBuffer|Buffer> — xlsx bytes. */
    toXLSX:      () => viewRef.current?.toXLSX(),
    view:        () => viewRef.current
  }), []);

  return React.createElement('div', {
    ref: hostRef,
    className: ['aix-react-host', className].filter(Boolean).join(' '),
    style: { width: '100%', height: '100%', ...(style || {}) },
    ...rest
  });
});

export default SheetViewReact;

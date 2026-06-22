# aix-sheet

AI-friendly spreadsheet SDK + embeddable viewer UI.

- **SDK** — A small JavaScript API for building spreadsheets, designed so that
  LLMs can drive it correctly on the first try. A1 notation everywhere, strict
  enums, helpful error messages.
- **Workbook** — Multi-sheet container with cross-sheet formula references
  (`Sheet1!A1` / `'Sheet Name'!A1`).
- **Viewer** — A framework-agnostic spreadsheet UI you can mount into any DOM
  element. Toolbar, formula bar, function autocomplete, range-pick drag,
  Excel-style autofill, formula-ref highlights, conditional-formatting dialog,
  merge & center split button, sheet tabs, undo, CSV/JSON/XLSX I/O.
- **React** — A `<SheetViewReact />` wrapper for React apps.
- **File format** — Plain JSON (`.aix.json`). Round-trips through `.xlsx` with
  styles, conditional formatting, merges and Date numFmts preserved.

---

## Install

This package is distributed via GitHub (not yet on the npm registry):

```bash
npm install github:yryuu/aix-sheet
```

Or in `package.json`:

```json
"dependencies": {
  "aix-sheet": "github:yryuu/aix-sheet"
}
```

For `.xlsx` export, you also need the style-writing fork of SheetJS plus JSZip
(both are `optionalDependencies` — only required when you actually call `toXLSX()`):

```bash
npm install xlsx-js-style jszip
```

Or clone and run the demo locally:

```bash
git clone https://github.com/yryuu/aix-sheet.git
cd aix-sheet
npm install
npm run serve      # demo at http://localhost:8080
```

---

## Quick start

### Vanilla JS (multi-sheet workbook + viewer)

```html
<link rel="stylesheet" href="node_modules/aix-sheet/viewer/styles.css">
<div id="app" style="height: 500px"></div>

<script type="module">
  import { Workbook } from 'aix-sheet';
  import { SheetView } from 'aix-sheet/viewer';

  const wb = new Workbook();
  const sales = wb.addSheet('Sales');
  sales.write('A1:B1', ['商品', '売上']);
  sales.style('A1:B1', { bold: true, bgColor: '#217346', color: '#fff' });
  sales.write('A2:A4', ['りんご', 'みかん', 'ぶどう']);
  sales.write('B2:B4', [12000, 18500, 9800]);
  sales.write('B5', '=SUM(B2:B4)');

  const summary = wb.addSheet('Summary');
  summary.write('A1', '総売上');
  summary.write('B1', '=SUM(Sales!B2:B4)');     // cross-sheet reference

  const view = new SheetView('#app', { workbook: wb });
  view.on('change', e => console.log(e));
</script>
```

### React

```jsx
import { Sheet } from 'aix-sheet';
import { SheetViewReact } from 'aix-sheet/react';
import 'aix-sheet/viewer/styles.css';

function App() {
  const sheet = useMemo(() => {
    const s = new Sheet();
    s.write('A1', 'Hello');
    return s;
  }, []);
  return (
    <SheetViewReact
      sheet={sheet}
      onChange={({ ref, newValue }) => console.log(ref, newValue)}
      style={{ height: 500 }}
    />
  );
}
```

### Node.js — generate an xlsx programmatically

```js
import { Workbook } from 'aix-sheet';
import { writeFile } from 'node:fs/promises';

const wb = new Workbook();
const s = wb.addSheet('Report');
s.write('A1:C1', ['日付', '商品', '売上']);
s.style('A1:C1', { bold: true, bgColor: '#217346', color: '#fff', border: true });
s.write('A2:A4', [new Date(2026,5,17), new Date(2026,5,18), new Date(2026,5,19)]);
s.write('B2:B4', ['りんご', 'みかん', 'ぶどう']);
s.write('C2:C4', [12000, 18500, 9800]);
s.write('B5', '合計');
s.write('C5', '=SUM(C2:C4)');
s.style('A5:C5', { bold: true, border: { top: 'medium' } });

await writeFile('report.xlsx', await wb.toXLSX());
```

---

## Package exports

| Specifier | What you get |
|---|---|
| `aix-sheet` | `{ Sheet, Workbook }` — SDK only, framework-free |
| `aix-sheet/viewer` | `{ SheetView }` — vanilla viewer (mount into a DOM node) |
| `aix-sheet/viewer/styles.css` | the viewer's CSS |
| `aix-sheet/react` | `{ SheetViewReact }` — React component wrapper |
| `aix-sheet/llms.txt` | AI/LLM usage guide (paste into your system prompt) |

---

## SDK reference

Full reference: [`sdk/llms.txt`](sdk/llms.txt) — also intended as an LLM system-prompt snippet.

```js
new Sheet(name?)
new Workbook()                          // multi-sheet container

// Read / write
sheet.write(ref, value)                 // "A1" | "A1:C5"  + value or array
sheet.value(ref) / sheet.read(ref)      // computed value
sheet.get(ref)                          // raw { v, f, s }

// Styling (merged with existing style)
sheet.style(ref, {
  bold, italic, underline,
  color, bgColor,
  align: 'left' | 'center' | 'right',
  fontSize, fontFamily,
  border: true | 'all' | false | { top, right, bottom, left, all },
  numFmt: 'yyyy/m/d' | '#,##0' | '0.0%' | '¥#,##0' | ...,
})

// Merge
sheet.merge('A1:D1')
sheet.unmerge('A1:D1')
sheet.mergeAt(row, col)

// Conditional formatting (Excel-compatible; re-evaluates on open)
sheet.cf('G3:Z3', {
  formula: '=AND(COLUMN()-7>=$D3, COLUMN()-7<=$E3)',
  style:   { bgColor: '#4472c4', color: '#fff', bold: true }
})
sheet.clearCF(rangeOrAll?)
sheet.cfStyleAt(row, col)               // resolve current overrides

// Autofill (Excel-style; relative refs shift, absolute $ refs stay)
sheet.fill('A1:A2', 'A3:A10')           // 1,2 → 3,4,5,...   (series detected)
sheet.fill('B1', 'B2:B10')              // copy + shift refs / +1 day for Date

// Structure
sheet.insertRow(rowNum, count?)
sheet.deleteRow(rowNum, count?)
sheet.insertCol(colLabel, count?)
sheet.deleteCol(colLabel, count?)
sheet.colWidth(col, px)
sheet.rowHeight(row, px)
sheet.clear(ref)

// Images
sheet.addImage('B2', dataUrl, { size, offset })
sheet.removeImage(id) / moveImage / resizeImage

// I/O
sheet.toJSON() / sheet.save(path)       // .aix.json
sheet.toCSV()  / Sheet.fromCSV(text)
await sheet.toXLSX()                    // requires xlsx-js-style (+ jszip)
await wb.toXLSX()                       // all sheets

// LLM context helper
sheet.toMarkdown({ maxRows? })
```

### Supported formula functions

`SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNT`, `COUNTA`, `IF`, `AND`, `OR`, `NOT`,
`WEEKDAY`. Plus `+ - * / % **`, comparisons (`= <> > < >= <=`), absolute refs
(`$A$1`, `$A1`, `A$1`), `ROW()`, `COLUMN()`. Anything beyond that is still
written to `.xlsx` verbatim and re-evaluated by Excel on open.

---

## Viewer reference

```js
const view = new SheetView(container, {
  workbook,            // OR pass a single `sheet`
  rows: 50,            // visible rows
  cols: 26,            // visible cols
  readOnly: false,
  toolbar: true,
  formulaBar: true,
  statusBar: true,
  tabs: true,
});

view.on('change',     ({ ref, oldValue, newValue, type? }) => {})
view.on('select',     ({ range }) => {})
view.on('edit-start', ({ ref }) => {})
view.on('edit-end',   ({ ref }) => {})
view.on('sheet-change', ({ name }) => {})
view.on('image-change', ({ id }) => {})

view.refresh()
view.setSheet(newSheet) / view.setWorkbook(wb)
view.focusCell('B12')
view.destroy()
```

### Toolbar features

- 行/列 追加・削除
- フォント (family / size), 太字 / 斜体 / 下線, 文字色 / 背景色, 左右中央
- **マージ split button**: 結合して中央 / 横方向に結合 / セルを結合 / 解除
- **罫線** dropdown: 全て / 外枠 / 上下左右 / なし (style はそれぞれ thin)
- **書式** dropdown: 標準 / 数値 (3桁区切り, 小数) / 日付 (yyyy/m/d, M/d など) / 通貨 / パーセント
- **⚡ 条件付き書式 ダイアログ**: 範囲, 数式, bg/fg color, 太字/斜体, プリセット (プラス値を緑 / マイナス値を赤 / 週末をオレンジ)
- 🖼 画像 (data URL anchor + drag/resize)
- .aix.json 保存 / 読込, CSV 保存, **📊 Excel 保存** (CDN から xlsx-js-style + JSZip 自動ロード)

### Cell-editing UX (Excel-like)

- ダブルクリックで編集、`F2` で編集、cursor は末尾に置かれる
- 編集中に他のセルをクリック: cursor が末尾なら単に移動、末尾でなければ参照を挿入 (range-pick mode)
- 編集中の数式の参照セルが viewer 上で**カラフルな点線枠**で囲まれる
- **オートフィル**: 選択範囲の右下の緑ハンドルを縦/横にドラッグ
  - 単一セル: 値・スタイル・式 (相対参照シフト) をコピー / Date は1日刻みで進める
  - 複数セル: 等差数列 / 等差日付を検出して外挿 (`1, 2` → `3, 4, 5...`)
  - 範囲内に CF rule があれば fill 方向に rule range を自動拡張
- `Cmd/Ctrl + Z` で undo

---

## File format (`.aix.json`)

```json
{
  "version": "1.0",
  "type": "aix-sheet",
  "name": "Sheet1",
  "cells": {
    "A1": { "v": "Hello", "s": { "bold": true } },
    "B1": { "f": "=SUM(A1:A10)" },
    "C1": { "v": "2026-06-01", "t": "d", "s": { "numFmt": "yyyy/m/d" } }
  },
  "cols": { "A": 120 },
  "rows": { "1": 24 },
  "merges": ["A1:D1"],
  "cfs": [
    { "range": "G3:Z3", "formula": "=$D3>0", "style": { "bgColor": "#4472c4" } }
  ],
  "images": []
}
```

`cells[ref].v` = literal · `.f` = formula · `.s` = style · `.t: "d"` = Date.

A workbook file uses `"type": "aix-workbook"` with a `sheets[]` array of the same shape (minus `type`/`version`).

You **may edit this JSON directly** — the loader will accept hand-edited files.

---

## Why a custom format?

LLMs struggle with the complexity of existing Excel libraries (deeply nested
APIs, version drift, missing docs). `aix-sheet` keeps the SDK surface tiny,
the API names obvious, and the file format human-readable JSON. AI agents can
either call the SDK or edit the JSON directly. Final XLSX conversion is
deferred to a known-good library (`xlsx-js-style` for the styled output).

See [`sdk/llms.txt`](sdk/llms.txt) — drop it into your LLM's system prompt.

---

## Development

```bash
git clone https://github.com/yryuu/aix-sheet.git
cd aix-sheet
npm install                          # pulls xlsx-js-style + jszip for tests
npm run serve                        # http://localhost:8080 — demo with Gantt
node examples/basic.js
node examples/formulas.js
node examples/workbook.js
node examples/with-image.js
```

The demo (`index.html`) shows: cross-sheet refs, images, merge + border,
**weekly Gantt chart** with dates driven by a Config sheet and bars rendered
via conditional formatting.

## License

MIT

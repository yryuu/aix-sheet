/**
 * xlsx-panes.js — inject frozen-pane definitions into a SheetJS-generated xlsx.
 *
 * Excel's "ウィンドウ枠の固定" (freeze panes). Adds a `<pane>` element inside
 * each worksheet's `<sheetView>` so the first N rows / M columns stay put
 * while the rest of the sheet scrolls.
 */

import { makeRef } from './sheet.js';

async function getJSZip() {
  if (typeof window !== 'undefined' && window.JSZip) return window.JSZip;
  if (typeof process !== 'undefined' && process.versions?.node) {
    try { const mod = await import('jszip'); return mod.default ?? mod; } catch {}
  }
  return null;
}

function buildPaneXml(rows, cols) {
  const r = Math.max(0, rows | 0);
  const c = Math.max(0, cols | 0);
  if (r === 0 && c === 0) return '';
  const topLeft = makeRef(r, c);
  let activePane;
  if (r > 0 && c > 0) activePane = 'bottomRight';
  else if (r > 0)     activePane = 'bottomLeft';
  else                activePane = 'topRight';
  const attrs = [];
  if (c > 0) attrs.push(`xSplit="${c}"`);
  if (r > 0) attrs.push(`ySplit="${r}"`);
  attrs.push(`topLeftCell="${topLeft}"`);
  attrs.push(`activePane="${activePane}"`);
  attrs.push(`state="frozen"`);
  return `<pane ${attrs.join(' ')}/><selection pane="${activePane}"/>`;
}

/**
 * @param {ArrayBuffer|Uint8Array|Buffer} xlsxData
 * @param {Array<{name:string, frozenPane: {rows:number, cols:number} | null}>} perSheet
 */
export async function injectPanes(xlsxData, perSheet) {
  if (!perSheet.some(s => s.frozenPane)) return xlsxData;
  const JSZip = await getJSZip();
  if (!JSZip) throw new Error('[JSZIP_MISSING] JSZip が必要です');
  const zip = await JSZip.loadAsync(xlsxData);

  for (let i = 0; i < perSheet.length; i++) {
    const { frozenPane } = perSheet[i];
    if (!frozenPane) continue;
    const pane = buildPaneXml(frozenPane.rows, frozenPane.cols);
    if (!pane) continue;
    const path = `xl/worksheets/sheet${i + 1}.xml`;
    const f = zip.file(path);
    if (!f) continue;
    let xml = await f.async('string');
    // Replace <sheetView .../> (self-closing) or inject before </sheetView>
    if (/<sheetView[^>]*\/>/.test(xml)) {
      xml = xml.replace(/<sheetView([^>]*)\/>/, `<sheetView$1>${pane}</sheetView>`);
    } else if (/<sheetView[^>]*>/.test(xml)) {
      xml = xml.replace(/(<sheetView[^>]*>)/, `$1${pane}`);
    } else {
      // No <sheetViews> at all — add one before <sheetData>
      const block = `<sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>`;
      xml = xml.replace('<sheetData', `${block}<sheetData`);
    }
    zip.file(path, xml);
  }

  return zip.generateAsync({
    type: typeof window !== 'undefined' ? 'arraybuffer' : 'nodebuffer'
  });
}

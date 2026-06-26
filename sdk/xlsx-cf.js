/**
 * xlsx-cf.js — inject conditional formatting into a SheetJS-generated xlsx.
 *
 * Adds <conditionalFormatting> rules to each worksheet and a corresponding
 * <dxfs> entry to xl/styles.xml. Excel re-evaluates the rules on open, so
 * editing referenced cells updates the formatting automatically.
 *
 * Requires JSZip (same as xlsx-images.js).
 */

async function getJSZip() {
  if (typeof window !== 'undefined' && window.JSZip) return window.JSZip;
  if (typeof process !== 'undefined' && process.versions?.node) {
    try { const mod = await import('jszip'); return mod.default ?? mod; } catch {}
  }
  return null;
}

function xmlEscape(s) {
  return String(s).replace(/[<>&"']/g, ch =>
    ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' }[ch]));
}

/** Normalize a CSS-ish hex color (#fff, #abc123, #FF217346…) to ARGB (8 hex). */
function toARGB(hex) {
  let h = String(hex || '').replace('#', '').toUpperCase();
  if (h.length === 3) h = h.split('').map(c => c + c).join('');
  if (h.length === 4) {
    const [r, g, b, a] = h.split('').map(c => c + c);
    h = a + r + g + b;
  }
  if (h.length === 6) return 'FF' + h;
  if (h.length === 8) return h;
  return 'FF000000';
}

/** Build a <dxf> element from a style object (subset: bgColor, color, bold, italic). */
function buildDxf(style) {
  const parts = [];
  const fontParts = [];
  if (style.bold) fontParts.push('<b/>');
  if (style.italic) fontParts.push('<i/>');
  if (style.color) fontParts.push(`<color rgb="${toARGB(style.color)}"/>`);
  if (fontParts.length) parts.push(`<font>${fontParts.join('')}</font>`);
  if (style.bgColor) {
    parts.push(`<fill><patternFill patternType="solid"><bgColor rgb="${toARGB(style.bgColor)}"/></patternFill></fill>`);
  }
  return `<dxf>${parts.join('')}</dxf>`;
}

/**
 * Inject conditional-formatting rules and matching dxfs into an xlsx.
 * @param {ArrayBuffer|Uint8Array|Buffer} xlsxData
 * @param {Array<{name:string, cfs:Array<{range:string,formula:string,style:object}>}>} perSheet
 * @returns {Promise<ArrayBuffer|Buffer>}
 */
export async function injectCF(xlsxData, perSheet) {
  const JSZip = await getJSZip();
  if (!JSZip) {
    throw new Error('[JSZIP_MISSING] JSZip が必要です. ブラウザでは https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js を、Node では `npm install jszip` を実行してください');
  }

  // Skip if no sheet has CFs
  if (!perSheet.some(s => s.cfs && s.cfs.length)) return xlsxData;

  const zip = await JSZip.loadAsync(xlsxData);

  // Build a single dxfs list (shared across sheets). Each CF rule -> one dxf.
  const dxfs = [];
  const sheetCfXml = []; // [{ sheetIndex, xml }]

  for (let s = 0; s < perSheet.length; s++) {
    const { cfs } = perSheet[s];
    if (!cfs || !cfs.length) continue;
    const lines = [];
    for (const rule of cfs) {
      const dxfId = dxfs.length;
      dxfs.push(buildDxf(rule.style || {}));
      // Excel formula: strip leading '='
      const formula = (rule.formula || '').replace(/^=/, '');
      lines.push(
        `<conditionalFormatting sqref="${xmlEscape(rule.range)}">` +
        `<cfRule type="expression" dxfId="${dxfId}" priority="${dxfId + 1}">` +
        `<formula>${xmlEscape(formula)}</formula>` +
        `</cfRule></conditionalFormatting>`
      );
    }
    sheetCfXml.push({ sheetIndex: s, xml: lines.join('') });
  }

  // 1) Patch styles.xml: append <dxfs>
  const stylesPath = 'xl/styles.xml';
  let styles = await zip.file(stylesPath).async('string');
  if (/<dxfs[^>]*\/?>/.test(styles)) {
    // Existing dxfs section — replace it (we own it for this workbook)
    styles = styles.replace(/<dxfs[\s\S]*?<\/dxfs>|<dxfs[^>]*\/>/,
      `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>`);
  } else {
    // Insert before </styleSheet>. Per schema dxfs comes after cellStyles, before tableStyles.
    styles = styles.replace('</styleSheet>',
      `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs></styleSheet>`);
  }
  zip.file(stylesPath, styles);

  // 2) Patch each sheet xml: insert <conditionalFormatting> blocks.
  // Position: after </mergeCells> if present, else after </sheetData>; must come BEFORE
  // <ignoredErrors>, <drawing>, <pageMargins>, etc. The simplest reliable point is to
  // insert it right before any of those known later elements, or before </worksheet>.
  for (const { sheetIndex, xml } of sheetCfXml) {
    const n = sheetIndex + 1;
    const path = `xl/worksheets/sheet${n}.xml`;
    const f = zip.file(path);
    if (!f) continue;
    let sheetXml = await f.async('string');

    // Try insertion points in priority order (schema order):
    const LATER_TAGS = ['<dataValidations', '<hyperlinks', '<printOptions',
      '<pageMargins', '<pageSetup', '<headerFooter', '<rowBreaks', '<colBreaks',
      '<customProperties', '<cellWatches', '<ignoredErrors', '<smartTags', '<drawing',
      '<legacyDrawing', '<legacyDrawingHF', '<picture', '<oleObjects', '<controls',
      '<webPublishItems', '<tableParts', '<extLst'];
    let inserted = false;
    for (const tag of LATER_TAGS) {
      const idx = sheetXml.indexOf(tag);
      if (idx !== -1) {
        sheetXml = sheetXml.slice(0, idx) + xml + sheetXml.slice(idx);
        inserted = true;
        break;
      }
    }
    if (!inserted) {
      // Insert right before </worksheet>
      sheetXml = sheetXml.replace('</worksheet>', xml + '</worksheet>');
    }
    zip.file(path, sheetXml);
  }

  return zip.generateAsync({
    type: typeof window !== 'undefined' ? 'arraybuffer' : 'nodebuffer'
  });
}

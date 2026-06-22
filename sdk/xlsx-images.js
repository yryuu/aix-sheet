/**
 * xlsx-images.js — inject images into a SheetJS-generated xlsx (zip+xml).
 *
 * Requires JSZip:
 *   Browser: <script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
 *   Node:    npm install jszip
 *
 * Reference: ECMA-376 SpreadsheetML / DrawingML (xdr:oneCellAnchor + xdr:pic)
 */
import { parseRef } from './sheet.js';

const PX_TO_EMU = 9525; // 1 pixel ≈ 9525 EMU at 96 DPI

async function getJSZip() {
  if (typeof window !== 'undefined' && window.JSZip) return window.JSZip;
  if (typeof process !== 'undefined' && process.versions?.node) {
    try {
      const mod = await import('jszip');
      return mod.default ?? mod;
    } catch {}
  }
  return null;
}

function decodeDataURL(src) {
  const m = src.match(/^data:image\/([a-zA-Z]+);base64,(.+)$/);
  if (!m) return null;
  let ext = m[1].toLowerCase();
  if (ext === 'jpeg') ext = 'jpg';
  if (!['png', 'jpg', 'gif'].includes(ext)) ext = 'png';
  const b64 = m[2];
  let binary;
  if (typeof Buffer !== 'undefined') {
    binary = Buffer.from(b64, 'base64');
  } else {
    const bin = atob(b64);
    binary = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) binary[i] = bin.charCodeAt(i);
  }
  return { ext, binary };
}

function xmlEscape(s) {
  return String(s).replace(/[<>&"']/g, ch =>
    ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' }[ch]));
}

function buildDrawingXml(images) {
  const anchors = images.map((img, i) => {
    const p = parseRef(img.anchor);
    const cx = Math.round(img.size.width  * PX_TO_EMU);
    const cy = Math.round(img.size.height * PX_TO_EMU);
    const offX = Math.round((img.offset?.x || 0) * PX_TO_EMU);
    const offY = Math.round((img.offset?.y || 0) * PX_TO_EMU);
    const id = i + 2; // start at 2 (id=1 sometimes reserved)
    return `  <xdr:oneCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>${p.col}</xdr:col>
      <xdr:colOff>${offX}</xdr:colOff>
      <xdr:row>${p.row}</xdr:row>
      <xdr:rowOff>${offY}</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="${cx}" cy="${cy}"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="${id}" name="Picture ${i + 1}" descr="${xmlEscape(img.id)}"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId${i + 1}"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${cx}" cy="${cy}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>`;
  }).join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${anchors}
</xdr:wsDr>`;
}

function buildDrawingRels(images) {
  const rels = images.map((img, i) =>
    `  <Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${img._mediaName}"/>`
  ).join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels}
</Relationships>`;
}

/**
 * Inject images into a SheetJS-generated xlsx ArrayBuffer/Buffer.
 * @param {ArrayBuffer|Uint8Array|Buffer} xlsxData - the raw xlsx bytes
 * @param {Array<{name: string, images: Array}>} perSheet - { name, images } per sheet (1:1 to workbook sheets)
 * @returns {Promise<ArrayBuffer|Buffer>}
 */
export async function injectImages(xlsxData, perSheet) {
  const JSZip = await getJSZip();
  if (!JSZip) {
    throw new Error('[JSZIP_MISSING] JSZip が必要です. ブラウザでは https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js を、Node では `npm install jszip` を実行してください');
  }

  const zip = await JSZip.loadAsync(xlsxData);

  let mediaCounter = 0;
  const extsUsed = new Set();
  const drawingsToAdd = []; // {sheetIndex, images}

  for (let s = 0; s < perSheet.length; s++) {
    const entry = perSheet[s];
    if (!entry.images || entry.images.length === 0) continue;

    const processed = [];
    for (const img of entry.images) {
      const dec = decodeDataURL(img.src);
      if (!dec) continue; // skip non-data URLs (cannot embed remote URLs)
      mediaCounter++;
      const mediaName = `image${mediaCounter}.${dec.ext}`;
      zip.file(`xl/media/${mediaName}`, dec.binary, { binary: true });
      extsUsed.add(dec.ext);
      processed.push({ ...img, _mediaName: mediaName });
    }
    if (!processed.length) continue;

    drawingsToAdd.push({ sheetIndex: s, images: processed });
  }

  // Write drawings + update sheet xml + sheet rels
  for (const d of drawingsToAdd) {
    const n = d.sheetIndex + 1;
    zip.file(`xl/drawings/drawing${n}.xml`, buildDrawingXml(d.images));
    zip.file(`xl/drawings/_rels/drawing${n}.xml.rels`, buildDrawingRels(d.images));

    // Patch worksheet
    const sheetXmlPath = `xl/worksheets/sheet${n}.xml`;
    let sheetXml = await zip.file(sheetXmlPath).async('string');
    if (!sheetXml.includes('<drawing ')) {
      // Insert <drawing r:id="rId_drawing"/> before </worksheet>
      sheetXml = sheetXml.replace(
        '</worksheet>',
        '<drawing r:id="rIdDrawing1"/></worksheet>'
      );
      // Ensure xmlns:r namespace is declared on <worksheet>
      if (!/xmlns:r=/.test(sheetXml)) {
        sheetXml = sheetXml.replace(
          /<worksheet([^>]*)>/,
          '<worksheet$1 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        );
      }
      zip.file(sheetXmlPath, sheetXml);
    }

    // Patch / create sheet rels
    const relsPath = `xl/worksheets/_rels/sheet${n}.xml.rels`;
    const existing = zip.file(relsPath);
    if (existing) {
      let rels = await existing.async('string');
      if (!rels.includes(`Target="../drawings/drawing${n}.xml"`)) {
        rels = rels.replace(
          '</Relationships>',
          `<Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${n}.xml"/></Relationships>`
        );
      }
      zip.file(relsPath, rels);
    } else {
      zip.file(relsPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${n}.xml"/>
</Relationships>`);
    }
  }

  // Update [Content_Types].xml
  let ct = await zip.file('[Content_Types].xml').async('string');
  const mimeMap = { png: 'image/png', jpg: 'image/jpeg', gif: 'image/gif' };
  for (const ext of extsUsed) {
    if (!ct.includes(`Extension="${ext}"`)) {
      ct = ct.replace('</Types>', `<Default Extension="${ext}" ContentType="${mimeMap[ext]}"/></Types>`);
    }
  }
  for (const d of drawingsToAdd) {
    const n = d.sheetIndex + 1;
    const part = `/xl/drawings/drawing${n}.xml`;
    if (!ct.includes(`PartName="${part}"`)) {
      ct = ct.replace('</Types>',
        `<Override PartName="${part}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);
    }
  }
  zip.file('[Content_Types].xml', ct);

  return zip.generateAsync({
    type: typeof window !== 'undefined' ? 'arraybuffer' : 'nodebuffer'
  });
}

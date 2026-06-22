/**
 * with-image.js — image embedded in spreadsheet, exported to .aix.json and .xlsx.
 * Run: node examples/with-image.js   (requires `npm install xlsx jszip`)
 */
import { Workbook } from '../sdk/sheet.js';
import { writeFile } from 'node:fs/promises';
import zlib from 'node:zlib';

const wb = new Workbook();
const s = wb.addSheet('Report');
s.write('A1', 'タイトル');
s.style('A1', { bold: true, fontSize: 18 });
s.write('A3:B3', ['項目', '値']);
s.style('A3:B3', { bold: true, bgColor: '#217346', color: '#ffffff' });
s.write('A4:A6', ['売上', '費用', '利益']);
s.write('B4:B6', [10000, 6000, '=B4-B5']);

// Build a real solid-color PNG (width x height, RGB triple) using Node's zlib.
function makePng(width, height, [r, g, b]) {
  const u32 = (n) => Buffer.from([n>>>24&0xff, n>>>16&0xff, n>>>8&0xff, n&0xff]);
  const crcTable = [];
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crcTable[i] = c >>> 0;
  }
  const crc = (buf) => {
    let c = 0xffffffff;
    for (const x of buf) c = (crcTable[(c ^ x) & 0xff] ^ (c >>> 8)) >>> 0;
    return c ^ 0xffffffff;
  };
  const chunk = (type, data) => {
    const t = Buffer.from(type, 'ascii');
    return Buffer.concat([u32(data.length), t, data, u32(crc(Buffer.concat([t, data])))]);
  };

  // IHDR: width, height, bit-depth=8, color-type=2(RGB), compression=0, filter=0, interlace=0
  const ihdr = Buffer.concat([u32(width), u32(height), Buffer.from([8, 2, 0, 0, 0])]);

  // Raw image data: per scanline: 1 filter byte (0) + width*3 bytes of RGB
  const row = Buffer.alloc(width * 3);
  for (let x = 0; x < width; x++) { row[x*3] = r; row[x*3+1] = g; row[x*3+2] = b; }
  const raw = Buffer.alloc(height * (1 + width * 3));
  for (let y = 0; y < height; y++) {
    raw[y * (1 + width * 3)] = 0;
    row.copy(raw, y * (1 + width * 3) + 1);
  }
  const idat = zlib.deflateSync(raw);

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', idat),
    chunk('IEND', Buffer.alloc(0))
  ]);
}

const redPng = makePng(100, 60, [220, 50, 50]);
const dataUrl = 'data:image/png;base64,' + redPng.toString('base64');

s.addImage('D2', dataUrl, { size: { width: 200, height: 120 } });

console.log('Images:', s.images);

// Save .aix.json
await wb.save('examples/with-image.aix.json');
console.log('→ examples/with-image.aix.json');

// Save .xlsx (requires xlsx + jszip)
try {
  const buf = await wb.toXLSX();
  await writeFile('examples/with-image.xlsx', buf);
  console.log('→ examples/with-image.xlsx (Excelで開いて画像が表示されることを確認してください)');
} catch (e) {
  console.error('xlsx export failed:', e.message);
}

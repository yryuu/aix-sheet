/**
 * basic.js — minimal example.
 * Run: node examples/basic.js
 */
import { Sheet } from '../sdk/sheet.js';

const s = new Sheet('Sales');

// Header row
s.write('A1:C1', ['商品', '価格', '在庫']);
s.style('A1:C1', { bold: true, bgColor: '#217346', color: '#ffffff', align: 'center' });

// Data
s.write('A2:A4', ['りんご', 'みかん', 'ぶどう']);
s.write('B2:B4', [120, 80, 350]);
s.write('C2:C4', [50, 100, 20]);

// Totals
s.write('A5', '合計');
s.write('B5', '=SUM(B2:B4)');
s.write('C5', '=SUM(C2:C4)');
s.style('A5:C5', { bold: true, bgColor: '#e8f4ee' });

// Column widths
s.colWidth('A', 120).colWidth('B', 80).colWidth('C', 80);

// Inspect
console.log(s.toMarkdown());
console.log('\nB5 =', s.value('B5'));

// Save
await s.save('examples/sales.aix.json');
console.log('\n→ examples/sales.aix.json に保存しました');

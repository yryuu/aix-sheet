/**
 * workbook.js — multi-sheet workbook example.
 * Run: node examples/workbook.js
 */
import { Workbook } from '../sdk/sheet.js';

const wb = new Workbook();

// Sheet 1: input data
const input = wb.addSheet('Input');
input.write('A1:B1', ['日付', '売上']);
input.style('A1:B1', { bold: true, bgColor: '#217346', color: '#ffffff' });
input.write('A2:A4', ['2026-06-17', '2026-06-18', '2026-06-19']);
input.write('B2:B4', [12000, 18500, 9800]);

// Sheet 2: summary referencing Sheet 1
const summary = wb.addSheet('Summary');
summary.write('A1', '指標');
summary.write('B1', '値');
summary.style('A1:B1', { bold: true });
summary.write('A2', '合計');         summary.write('B2', '=SUM(Input!B2:B4)');
summary.write('A3', '平均');         summary.write('B3', '=AVERAGE(Input!B2:B4)');
summary.write('A4', '最大');         summary.write('B4', '=MAX(Input!B2:B4)');
summary.write('A5', '前日比 (3/2)');  summary.write('B5', '=Input!B4-Input!B3');

// Activate the Summary sheet by default
wb.active = 'Summary';

console.log('=== Input ===\n' + input.toMarkdown());
console.log('\n=== Summary ===\n' + summary.toMarkdown());

console.log('\n指標:');
console.log('  合計:', summary.value('B2'));
console.log('  平均:', summary.value('B3'));
console.log('  最大:', summary.value('B4'));
console.log('  前日比:', summary.value('B5'));

await wb.save('examples/sales-book.aix.json');
console.log('\n→ examples/sales-book.aix.json に保存しました');

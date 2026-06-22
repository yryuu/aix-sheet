/**
 * formulas.js — demonstrate formulas, IF, conditional styling.
 * Run: node examples/formulas.js
 */
import { Sheet } from '../sdk/sheet.js';

const s = new Sheet('Scores');

s.write('A1:D1', ['名前', '点数', '判定', '偏差']);
s.style('A1:D1', { bold: true });

const names  = ['田中', '佐藤', '鈴木', '山田', '伊藤'];
const scores = [78, 92, 45, 67, 88];

s.write(`A2:A${names.length + 1}`,  names);
s.write(`B2:B${scores.length + 1}`, scores);

// IF formula: 60点以上で「合格」
for (let r = 2; r <= scores.length + 1; r++) {
  s.write(`C${r}`, `=IF(B${r}>=60,"合格","不合格")`);
}

// Deviation from average
s.write('F1', '平均');
s.write('F2', `=AVERAGE(B2:B${scores.length + 1})`);
for (let r = 2; r <= scores.length + 1; r++) {
  s.write(`D${r}`, `=B${r}-F2`);
}

// Style failing rows
scores.forEach((score, i) => {
  const r = i + 2;
  if (score < 60) s.style(`A${r}:D${r}`, { color: '#cc0000' });
});

console.log(s.toMarkdown());
console.log('\n平均:', s.value('F2'));
console.log('鈴木の判定:', s.value('C4'));

await s.save('examples/scores.aix.json');

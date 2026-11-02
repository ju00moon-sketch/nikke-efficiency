import fs from 'node:fs';
import assert from 'node:assert/strict';
import { corrections } from './corrections.mjs';
import { CellError, Workbook, parseFormula } from './engine.mjs';
import { expandCells, formatValue, inputFor, isSupportedFormat, pickupConversion, rowTier, ruleReferences, tierAt, timelineEntries } from './view-model.mjs';

const root = new URL('./snapshot/', import.meta.url);
const manifest = JSON.parse(fs.readFileSync(new URL('manifest.json', root), 'utf8'));
const sheets = manifest.sheets.map(entry => JSON.parse(fs.readFileSync(new URL(entry.file, root), 'utf8')));
const workbook = new Workbook(sheets, corrections);
const sheet = name => sheets.find(item => item.name === name);
const visible = sheets.filter(item => manifest.sheets.find(entry => entry.name === item.name).visible);

assert.equal(manifest.visibleSheets, 19);
assert.equal(manifest.sheets.filter(s => s.visible).reduce((n, s) => n + s.rowCount, 0), 2124);
assert.equal(manifest.sheets.filter(s => s.visible).reduce((n, s) => n + s.sourceCellCount, 0), 21041);
assert.equal(manifest.sheets.filter(s => s.visible).reduce((n, s) => n + s.cellCount, 0), 21022);
assert.equal(manifest.sheets.find(s => s.name === '패키지효율').hiddenRowCount, 273);

let parsed = 0;
for (const item of sheets) for (const [, , cells] of item.rows) for (const [, , formula] of cells) {
  if (formula) { parseFormula(formula); parsed++; }
}
console.log(`Parsed ${parsed} formulas`);

let compared = 0, comparedText = 0;
for (const item of visible) {
  for (const [, , cells] of item.rows) for (const [address, cached, formula] of cells) {
    if (!formula) continue;
    const calculated = workbook.value(item.name, address);
    if (typeof cached === 'number') {
      assert.equal(typeof calculated, 'number', `${item.name}!${address} must stay numeric`);
      assert.ok(Math.abs(calculated - cached) <= 1e-5 * Math.max(1, Math.abs(cached)), `${item.name}!${address}: ${calculated} != ${cached}`);
      compared++;
    } else {
      assert.equal(calculated, cached, `${item.name}!${address}`);
      comparedText++;
    }
  }
}
assert.equal(compared, 7988);
assert.equal(comparedText, 21);
assert.equal(workbook.errors.size, 0);
console.log(`Matched ${compared} numeric and ${comparedText} text formulas against source values`);

function close(sheetName, address, expected, tolerance = 1e-6) {
  const value = Number(workbook.value(sheetName, address));
  assert.ok(Math.abs(value - expected) <= tolerance, `${sheetName}!${address}: ${value} != ${expected}`);
}

close('패키지효율', 'G7', 2.2);
close('패키지효율', 'I7', 1.200533333, 1e-8);
close('픽업효율', 'G6', 10);
close('픽업효율', 'H6', 600);
close('뽑기효율', 'G6', 10);
close('뽑기효율', 'I6', 5);
close('AFFECTION TRAIL', 'G8', 20);
close('재화가치기준', 'E7', 5);

const base = Number(workbook.value('패키지효율', 'I7'));
workbook.set('재화가치기준', 'D8', 300);
const adjusted = Number(workbook.value('패키지효율', 'I7'));
assert.ok(adjusted > base, 'Value edit must recompute package efficiency');
assert.equal(Number(workbook.value('패키지효율', 'J7')), 3300);
workbook.reset();
close('패키지효율', 'I7', 1.200533333, 1e-8);

// Edits that break a formula must show the spreadsheet error, never the stale snapshot value.
workbook.set('패키지효율', 'F7', 0);
assert.equal(workbook.value('패키지효율', 'I7'), 0, 'IFERROR must fall back when dividing by zero');
workbook.reset();
workbook.set('패키지효율', 'M7', 0);
workbook.set('패키지효율', 'N7', 0);
assert.equal(workbook.value('패키지효율', 'G7'), 0);
assert.ok(workbook.value('패키지효율', 'H7') instanceof CellError, 'Price per pull must become an error when there are no pulls');
assert.equal(workbook.value('패키지효율', 'H7').code, '#DIV/0!');
workbook.reset();
workbook.set('재화가치기준', 'G14', 0);
assert.equal(workbook.value('재화가치기준', 'H20').code, '#N/A', 'Unknown outpost level must not reuse the snapshot value');
assert.equal(workbook.value('재화가치기준', 'H24').code, '#N/A', 'Errors must propagate to dependent cells');
workbook.reset();
workbook.set('패키지효율', 'B110', true);
workbook.set('패키지효율', 'B110', false);
assert.equal(workbook.overrides.size, 0, 'Restoring the snapshot value is not counted as an edit');
assert.equal(workbook.errors.size, 0);

// Corrected source formulas: identical at the default inputs (checked above), right once inputs change.
assert.equal(workbook.corrections.size, 8);
const uncorrected = new Workbook(sheets);
uncorrected.set('스텝업', 'C7', 2);
assert.ok(Math.abs(uncorrected.value('스텝업', 'H7') - 0.6059259259) < 1e-9, 'The source formula ignores the repeat count');
workbook.set('스텝업', 'C7', 2);
assert.equal(workbook.value('스텝업', 'F7'), 67500);
close('스텝업', 'H7', (8750 + 32150 * 2) / (7500 + 30000 * 2), 1e-12);
workbook.reset();
workbook.set('재화가치기준', 'D7', 0);
assert.equal(workbook.value('스텝업', 'J9'), 28500, 'STEP 11 free jewels use the free jewel value');
assert.equal(workbook.value('패키지효율', 'J206'), 28500);
close('스텝업', 'G9', 32500 / 45000, 1e-12);
workbook.reset();
for (const [cell, option, value] of [['C13', '커스텀모듈 2', 400], ['C14', '커스텀모듈 2', 400], ['C22', '커스텀모듈 5', 1000]]) {
  const target = `D${cell.slice(1)}`;
  uncorrected.set('커스텀', cell, option);
  assert.equal(uncorrected.value('커스텀', target).code, '#N/A', `The source lookup for ${target} skips its first dropdown option`);
  workbook.set('커스텀', cell, option);
  assert.equal(workbook.value('커스텀', target), value);
  workbook.reset();
}

// Pickup conversion stays separate from the source's pull count, which adds ordinary tickets.
const productRow = name => sheet('패키지효율').rows.find(([, hidden, cells]) => !hidden && cells.some(cell => cell[0].startsWith('D') && cell[1] === name))[0];
const pickup = name => pickupConversion(workbook, productRow(name));
assert.deepEqual(pickup('4주년 패키지Ⅰ'), { pulls: 10, cost: 3000 });
close('패키지효율', `G${productRow('4주년 패키지Ⅰ')}`, 20);
assert.equal(pickup('4주년 패키지Ⅱ').pulls, 24);
assert.ok(Math.abs(pickup('4주년 패키지Ⅲ').pulls - 36.6666666667) < 1e-9);
assert.deepEqual(pickup('특수 모집 패키지Ⅰ'), { pulls: 20, cost: 1500 });
assert.deepEqual(pickup('신입지휘관 출석 패키지'), { pulls: 10, cost: 600 });
close('패키지효율', `G${productRow('신입지휘관 출석 패키지')}`, 20);
assert.ok(Math.abs(pickup('(월간) 30일 보급품').cost - 675.6756756757) < 1e-9);
assert.equal(pickup('신입지휘관 패키지Ⅰ'), null, 'Ordinary tickets alone are not pickup pulls');

// Number formats copied from the source cells.
assert.equal(formatValue(681.8181818, '[$₩-412]#,##0'), '₩682');
assert.equal(formatValue(1.200533333, '0.00%'), '120.05%');
assert.equal(formatValue(0.6, '0%'), '60%');
assert.equal(formatValue(2.2, '#,##0.0'), '2.2');
assert.equal(formatValue(15000, '#,##0.00'), '15,000.00');
assert.equal(formatValue(1884.433333, '0.0'), '1884.4');
assert.equal(formatValue(new CellError('#DIV/0!'), '0.00%'), '#DIV/0!');
for (const item of visible) for (const code of item.formats ?? []) assert.ok(isSupportedFormat(code), `${item.name}: unsupported format ${code}`);

// Inputs come from the source's dropdown validations, checkboxes and value column.
let selects = 0;
for (const item of visible) for (const rule of item.inputs ?? []) {
  for (const address of expandCells(rule.cells)) {
    const input = inputFor(workbook, item, address);
    assert.equal(input?.kind, 'select', `${item.name}!${address} must be a dropdown`);
    assert.ok(input.options.length > 0, `${item.name}!${address} has no options`);
    const current = workbook.value(item.name, address);
    assert.ok(input.options.some(option => option.value === current), `${item.name}!${address}: ${current} is not an option`);
    selects++;
  }
}
assert.equal(selects, 698);
assert.deepEqual(inputFor(workbook, sheet('재화가치기준'), 'H14').options.map(option => option.value), [0.1, 0.2, 0.3, 0.4, 0.5, 0.6]);
assert.equal(inputFor(workbook, sheet('재화가치기준'), 'D8').kind, 'number');
assert.equal(inputFor(workbook, sheet('재화가치기준'), 'E7'), null, 'Formula cells are never inputs');
assert.equal(inputFor(workbook, sheet('패키지효율'), 'B110').kind, 'checkbox');
assert.equal(inputFor(workbook, sheet('패키지효율'), 'I2').kind, 'number', 'Highlight thresholds are editable');
assert.equal(inputFor(workbook, sheet('AFFECTION TRAIL'), 'P3'), null);
assert.ok(inputFor(workbook, sheet('커스텀'), 'C7').options.some(option => option.value === '스킬매뉴얼Ⅰ 120'));

// Highlight rules copied from the source's conditional formats.
for (const item of visible) for (const rule of item.rules ?? []) {
  for (const reference of ruleReferences(rule)) assert.ok(reference.startsWith('$'), `${item.name}: ${rule.formula} has a relative column`);
}
assert.equal(rowTier(workbook, sheet('패키지효율'), productRow('특수 모집 패키지Ⅰ')), 'best');
assert.equal(rowTier(workbook, sheet('패키지효율'), productRow('4주년 패키지Ⅰ')), 'off', 'Unchecked rows are dimmed first');
assert.equal(rowTier(workbook, sheet('패키지효율'), productRow('특수 모집 패키지Ⅱ')), null);
workbook.set('패키지효율', `B${productRow('4주년 패키지Ⅱ')}`, true);
assert.equal(rowTier(workbook, sheet('패키지효율'), productRow('4주년 패키지Ⅱ')), 'good');
workbook.set('패키지효율', 'I2', 0.9);
assert.equal(rowTier(workbook, sheet('패키지효율'), productRow('4주년 패키지Ⅱ')), 'best', 'Threshold edits must re-evaluate highlights');
workbook.reset();
const tiers = { best: 0, good: 0, off: 0 };
for (const [row, hidden] of sheet('패키지효율').rows) if (row >= 6 && !hidden) {
  const tier = rowTier(workbook, sheet('패키지효율'), row);
  if (tier) tiers[tier]++;
}
assert.deepEqual(tiers, { best: 18, good: 6, off: 9 });
assert.equal(tierAt(workbook, sheet('AFFECTION TRAIL'), 'U6'), 'done');

// The horizontal schedule becomes dated entries using the source's merged cells.
const entries = timelineEntries(workbook, sheet('이벤트일정'));
assert.equal(entries.length, 1879);
assert.deepEqual(entries.find(entry => entry.address === 'I3'), { row: 3, address: 'I3', start: '2022-11-10', end: '2022-11-23', days: 14, category: '', value: 'NO CALLER ID' });
assert.equal(entries.find(entry => entry.address === 'W5').category, '이벤트');
assert.equal(entries.find(entry => entry.address === 'W5').end, '2022-12-07');

console.log(`Snapshot checks passed; recalculated ${base.toFixed(6)} -> ${adjusted.toFixed(6)}; ${selects} dropdown cells; ${entries.length} schedule entries`);

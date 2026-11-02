import { CellError, columnLetters, columnNumber, parseFormula } from './engine.mjs';

const ADDRESS = /^\$?([A-Z]+)\$?(\d+)$/;

function position(address) {
  const match = address.match(ADDRESS);
  if (!match) throw new Error(`Bad cell address: ${address}`);
  return [columnNumber(match[1]), Number(match[2])];
}

function areas(sqref) {
  return sqref.split(' ').map(part => {
    const [first, last = first] = part.split(':');
    const [c1, r1] = position(first), [c2, r2] = position(last);
    return { c1, r1, c2, r2 };
  });
}

export function expandCells(sqref) {
  const cells = [];
  for (const { c1, r1, c2, r2 } of areas(sqref)) {
    for (let row = r1; row <= r2; row++) for (let column = c1; column <= c2; column++) cells.push(`${columnLetters(column)}${row}`);
  }
  return cells;
}

// Number formats used by the source cells: optional currency, grouping, fixed decimals, percent.
const FORMAT = /^(?:\[\$([^\]-]*)(?:-[0-9A-F]+)?\])?(#,##)?0(?:\.(0+))?(%)?(?:_.)?$/;
const general = new Intl.NumberFormat('ko-KR', { maximumFractionDigits: 6 });
const fixed = new Map();

export function isSupportedFormat(code) {
  return code === 'General' || FORMAT.test(code);
}

export function cellFormat(sheet, cell) {
  return cell?.[4] == null ? undefined : sheet.formats[cell[4]];
}

export function formatValue(value, code) {
  if (value instanceof CellError) return value.code;
  if (value === true) return 'TRUE';
  if (value === false) return 'FALSE';
  if (typeof value !== 'number') return String(value ?? '');
  const match = code ? FORMAT.exec(code) : null;
  if (!match) return general.format(value);
  const [, currency = '', grouping, decimals = '', percent] = match;
  const key = `${decimals.length}${grouping ? ',' : ''}`;
  if (!fixed.has(key)) {
    fixed.set(key, new Intl.NumberFormat('ko-KR', { minimumFractionDigits: decimals.length, maximumFractionDigits: decimals.length, useGrouping: Boolean(grouping) }));
  }
  const text = fixed.get(key).format(percent ? value * 100 : value);
  if (percent) return `${text}%`;
  return text.startsWith('-') ? `-${currency}${text.slice(1)}` : `${currency}${text}`;
}

function references(node, list = []) {
  if (node.type === 'ref') list.push(node.address);
  else if (node.type === 'unary') references(node.value, list);
  else if (node.type === 'binary') { references(node.left, list); references(node.right, list); }
  else if (node.type === 'call') for (const arg of node.args) references(arg, list);
  return list;
}

const compiledRules = new WeakMap();

function compile(rule) {
  let entry = compiledRules.get(rule);
  if (!entry) {
    const refs = references(parseFormula(rule.formula));
    entry = { refs, areas: areas(rule.cells), fixedColumns: refs.every(ref => ref.startsWith('$')) };
    compiledRules.set(rule, entry);
  }
  return entry;
}

export function ruleReferences(rule) {
  return compile(rule).refs;
}

// Conditional formats apply in priority order and the first matching rule wins, as in Google Sheets.
export function tierAt(workbook, sheet, address) {
  const [column, row] = position(address);
  for (const rule of sheet.rules ?? []) {
    const { areas: ranges, fixedColumns } = compile(rule);
    if (!ranges.some(area => column >= area.c1 && column <= area.c2 && row >= area.r1 && row <= area.r2)) continue;
    const origin = ranges[0];
    const result = workbook.evaluateFormula(sheet.name, rule.formula, row - origin.r1, fixedColumns ? 0 : column - origin.c1);
    if (result === true || (typeof result === 'number' && result !== 0)) return rule.tier;
  }
  return null;
}

// A row's tier is the tier of its leftmost highlighted cell, which belongs to the sheet's main table.
export function rowTier(workbook, sheet, row) {
  let first = Infinity;
  for (const rule of sheet.rules ?? []) {
    for (const area of compile(rule).areas) if (row >= area.r1 && row <= area.r2) first = Math.min(first, area.c1);
  }
  return first === Infinity ? null : tierAt(workbook, sheet, `${columnLetters(first)}${row}`);
}

function parseOption(text) {
  const percent = text.match(/^(-?\d+(?:\.\d+)?)%$/);
  if (percent) return { value: Number(percent[1]) / 100, label: text };
  const number = Number(text);
  return { value: text.trim() !== '' && Number.isFinite(number) ? number : text, label: text };
}

function rangeOptions(workbook, source) {
  const match = source.match(/^(?:'((?:[^']|'')+)'|([^!']+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/);
  if (!match) throw new Error(`Unsupported dropdown source: ${source}`);
  const sheetName = match[1] ? match[1].replaceAll("''", "'") : match[2];
  const options = [];
  for (const address of expandCells(`${match[3]}${match[4]}:${match[5]}${match[6]}`)) {
    if (!workbook.raw(sheetName, address)) continue;
    const value = workbook.value(sheetName, address);
    if (value === '' || value instanceof CellError) continue;
    options.push({ value, label: formatValue(value) });
  }
  return options;
}

const inputMaps = new WeakMap();

function inputMap(workbook, sheet) {
  let maps = inputMaps.get(workbook);
  if (!maps) inputMaps.set(workbook, maps = new Map());
  let map = maps.get(sheet.name);
  if (map) return map;
  map = new Map();
  const plain = address => {
    const cell = workbook.raw(sheet.name, address);
    return cell && !cell[2] ? cell : null;
  };
  for (const [, , cells] of sheet.rows) for (const [address, , formula, type] of cells) {
    if (type === 'b' && !formula) map.set(address, { kind: 'checkbox' });
  }
  // The value column of the value basis tab is the source's documented place to enter your own values.
  if (sheet.name === '재화가치기준') for (let row = 6; row <= 72; row++) {
    const cell = plain(`D${row}`);
    if (typeof cell?.[1] === 'number') map.set(cell[0], { kind: 'number' });
  }
  // Fixed cells compared by highlight rules (e.g. the recommended efficiency) are editable thresholds.
  for (const rule of sheet.rules ?? []) for (const reference of ruleReferences(rule)) {
    if (!/^\$[A-Z]+\$\d+$/.test(reference)) continue;
    const cell = plain(reference.replaceAll('$', ''));
    if (typeof cell?.[1] === 'number') map.set(cell[0], { kind: 'number' });
  }
  for (const rule of sheet.inputs ?? []) {
    const options = rule.options ? rule.options.map(parseOption) : rangeOptions(workbook, rule.source);
    for (const address of expandCells(rule.cells)) if (plain(address)) map.set(address, { kind: 'select', options });
  }
  maps.set(sheet.name, map);
  return map;
}

export function inputFor(workbook, sheet, address) {
  return inputMap(workbook, sheet).get(address) ?? null;
}

// Pickup pulls count jewels (300 per pull) and advanced recruit tickets only; ordinary tickets stay separate.
export function pickupConversion(workbook, row) {
  const value = column => Number(workbook.value('패키지효율', `${column}${row}`));
  const price = value('F'), pulls = value('M') / 300 + value('O');
  if (!Number.isFinite(pulls) || pulls <= 0 || !Number.isFinite(price)) return null;
  return { pulls, cost: price / pulls };
}

// Column C of the schedule is 2022-11-04; the day numbers in row 2 confirm one column per day.
const TIMELINE_START = Date.UTC(2022, 10, 4);
const day = column => new Date(TIMELINE_START + (column - 3) * 86400000).toISOString().slice(0, 10);

export function timelineEntries(workbook, sheet) {
  const { spans = {}, groups = [] } = sheet.timeline ?? {};
  const entries = [];
  for (const [row, , cells] of sheet.rows) {
    if (row < 3) continue;
    const category = (groups.find(([first, last]) => row >= first && row <= last)?.[2] ?? '').replace(/\s*\n\s*/g, ' ');
    for (const [address] of cells) {
      const [column] = position(address);
      if (column < 3) continue;
      const end = spans[address] ? position(spans[address])[0] : column;
      entries.push({ row, address, start: day(column), end: day(end), days: end - column + 1, category, value: workbook.value(sheet.name, address) });
    }
  }
  return entries;
}

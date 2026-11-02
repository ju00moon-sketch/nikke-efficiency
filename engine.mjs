const REF = /^\$?[A-Z]{1,4}\$?\d+$/;

export class CellError {
  constructor(code) { this.code = code; }
  toString() { return this.code; }
}

export function columnNumber(letters) {
  let number = 0;
  for (const letter of letters) number = number * 26 + letter.charCodeAt(0) - 64;
  return number;
}

export function columnLetters(number) {
  let result = '';
  while (number) {
    number--;
    result = String.fromCharCode(65 + number % 26) + result;
    number = Math.floor(number / 26);
  }
  return result;
}

function location(address) {
  const match = address.replaceAll('$', '').match(/^([A-Z]+)(\d+)$/);
  if (!match) throw new Error(`Bad cell address: ${address}`);
  return [columnNumber(match[1]), Number(match[2])];
}

export function shiftReference(address, rows, columns) {
  const match = address.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/);
  if (!match) throw new Error(`Bad cell address: ${address}`);
  const [, fixedColumn, column, fixedRow, row] = match;
  return `${fixedColumn}${fixedColumn ? column : columnLetters(columnNumber(column) + columns)}${fixedRow}${fixedRow ? row : Number(row) + rows}`;
}

function shiftNode(node, rows, columns) {
  if (node.type === 'ref') return { ...node, address: shiftReference(node.address, rows, columns) };
  if (node.type === 'unary') return { ...node, value: shiftNode(node.value, rows, columns) };
  if (node.type === 'binary') return { ...node, left: shiftNode(node.left, rows, columns), right: shiftNode(node.right, rows, columns) };
  if (node.type === 'call') return { ...node, args: node.args.map(arg => shiftNode(arg, rows, columns)) };
  return node;
}

function tokenize(formula) {
  const tokens = [];
  let i = 0;
  while (i < formula.length) {
    const tail = formula.slice(i);
    if (/^\s/.test(tail)) { i++; continue; }
    if (tail[0] === '"') {
      let value = '';
      i++;
      while (i < formula.length) {
        if (formula[i] === '"') {
          if (formula[i + 1] === '"') { value += '"'; i += 2; continue; }
          i++;
          break;
        }
        value += formula[i++];
      }
      tokens.push({ type: 'string', value });
      continue;
    }
    if (tail[0] === "'") {
      let value = '';
      i++;
      while (i < formula.length) {
        if (formula[i] === "'") {
          if (formula[i + 1] === "'") { value += "'"; i += 2; continue; }
          i++;
          break;
        }
        value += formula[i++];
      }
      tokens.push({ type: 'sheet', value });
      continue;
    }
    const number = tail.match(/^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/);
    if (number) { tokens.push({ type: 'number', value: Number(number[0]) }); i += number[0].length; continue; }
    const word = tail.match(/^\$?[A-Za-z_][A-Za-z0-9_.$]*/);
    if (word) { tokens.push({ type: REF.test(word[0]) ? 'ref' : 'word', value: word[0] }); i += word[0].length; continue; }
    const op = tail.match(/^(?:>=|<=|<>|[+*/=<>:(),&!%-])/);
    if (op) { tokens.push({ type: op[0], value: op[0] }); i += op[0].length; continue; }
    throw new Error(`Unsupported formula token: ${tail.slice(0, 20)}`);
  }
  tokens.push({ type: 'end' });
  return tokens;
}

const priority = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '%': 4, ':': 5 };

export function parseFormula(formula) {
  const tokens = tokenize(formula.replace(/^=/, ''));
  let index = 0;
  const peek = () => tokens[index];
  const take = () => tokens[index++];
  const expect = type => { if (peek().type !== type) throw new Error(`Expected ${type}, found ${peek().type}`); take(); };
  function prefix() {
    const token = take();
    if (token.type === 'number' || token.type === 'string') return { type: token.type, value: token.value };
    if (token.type === '+' || token.type === '-') return { type: 'unary', op: token.type, value: expression(5) };
    if (token.type === '(') { const inner = expression(0); expect(')'); return inner; }
    if (token.type === 'sheet' || token.type === 'word') {
      if (peek().type === '!') { take(); const ref = take(); if (ref.type !== 'ref') throw new Error('Expected sheet cell reference'); return { type: 'ref', sheet: token.value, address: ref.value }; }
      if (token.type === 'word' && peek().type === '(') {
        take(); const args = [];
        while (peek().type !== ')') { args.push(expression(0)); if (peek().type !== ',') break; take(); }
        expect(')');
        return { type: 'call', name: token.value.toUpperCase(), args };
      }
      if (token.value.toUpperCase() === 'TRUE' || token.value.toUpperCase() === 'FALSE') return { type: 'boolean', value: token.value.toUpperCase() === 'TRUE' };
    }
    if (token.type === 'ref') return { type: 'ref', address: token.value };
    throw new Error(`Unexpected formula token: ${token.type} ${token.value || ''}`);
  }
  function expression(minimum) {
    let left = prefix();
    while (priority[peek().type] != null && priority[peek().type] >= minimum) {
      const op = take().type;
      const right = expression(priority[op] + 1);
      left = { type: 'binary', op, left, right };
    }
    return left;
  }
  const result = expression(0);
  if (peek().type !== 'end') throw new Error(`Trailing formula token: ${peek().type}`);
  return result;
}

function numeric(value) {
  if (value instanceof CellError) throw value;
  if (value == null || value === '') return 0;
  if (typeof value === 'boolean') return value ? 1 : 0;
  const number = Number(value);
  if (!Number.isFinite(number)) throw new CellError('#VALUE!');
  return number;
}

function comparable(a, b) {
  if (a instanceof CellError) throw a;
  if (b instanceof CellError) throw b;
  if (typeof a === 'number' || typeof b === 'number') {
    const an = Number(a), bn = Number(b);
    if (Number.isFinite(an) && Number.isFinite(bn)) return [an, bn];
  }
  return [String(a).toLowerCase(), String(b).toLowerCase()];
}

function text(value) {
  if (value instanceof CellError) throw value;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

export class Workbook {
  constructor(sheets, corrections = []) {
    this.sheets = new Map(sheets.map(sheet => [sheet.name, sheet]));
    this.cells = new Map();
    for (const sheet of sheets) {
      const map = new Map();
      for (const [, , rowCells] of sheet.rows) for (const cell of rowCells) map.set(cell[0], cell);
      this.cells.set(sheet.name, map);
    }
    // Corrected formulas replace the snapshot formula for calculation only; the snapshot rows stay untouched.
    this.corrections = new Map();
    for (const correction of corrections) {
      const cell = this.raw(correction.sheet, correction.cell);
      if (cell?.[2] !== correction.source) throw new Error(`Correction no longer matches ${correction.sheet}!${correction.cell}`);
      this.cells.get(correction.sheet).set(correction.cell, [cell[0], cell[1], correction.formula, ...cell.slice(3)]);
      this.corrections.set(`${correction.sheet}!${correction.cell}`, correction);
    }
    this.overrides = new Map();
    this.cache = new Map();
    this.ast = new Map();
    this.errors = new Map();
  }
  set(sheet, address, value) {
    const key = `${sheet}!${address}`;
    const cell = this.raw(sheet, address);
    // Setting a plain cell back to its snapshot value removes the override instead of storing a no-op edit.
    if (cell && !cell[2] && Object.is(value, cell[3] === 'b' ? Boolean(Number(cell[1])) : cell[1])) this.overrides.delete(key);
    else this.overrides.set(key, value);
    this.cache.clear();
    this.errors.clear();
  }
  reset() {
    this.overrides.clear();
    this.cache.clear();
    this.errors.clear();
  }
  raw(sheet, address) { return this.cells.get(sheet)?.get(address); }
  parse(formula) {
    let ast = this.ast.get(formula);
    if (!ast) { ast = parseFormula(formula); this.ast.set(formula, ast); }
    return ast;
  }
  value(sheet, address, stack = new Set()) {
    address = address.replaceAll('$', '');
    const key = `${sheet}!${address}`;
    if (this.overrides.has(key)) return this.overrides.get(key);
    if (this.cache.has(key)) return this.cache.get(key);
    const cell = this.raw(sheet, address);
    if (!cell) return 0;
    const [ , cached, formula, type] = cell;
    if (!formula) return type === 'b' ? Boolean(Number(cached)) : cached;
    if (stack.has(key)) {
      this.errors.set(key, 'Circular reference');
      return new CellError('#REF!');
    }
    stack.add(key);
    let result;
    try {
      result = this.evaluate(this.parse(formula), sheet, stack);
      if (result == null || Array.isArray(result)) result = new CellError('#VALUE!');
      else if (typeof result === 'number' && !Number.isFinite(result)) result = new CellError('#NUM!');
    } catch (error) {
      if (!(error instanceof CellError)) this.errors.set(key, error.message);
      result = error instanceof CellError ? error : new CellError('#ERROR!');
    } finally { stack.delete(key); }
    this.cache.set(key, result);
    return result;
  }
  // Evaluates a formula written for one cell as if it were copied `rows`/`columns` away (conditional formats).
  evaluateFormula(sheet, formula, rows = 0, columns = 0) {
    const key = `${sheet}\u0000${formula}\u0000${rows}\u0000${columns}`;
    if (this.cache.has(key)) return this.cache.get(key);
    let result;
    try {
      result = this.evaluate(shiftNode(this.parse(formula), rows, columns), sheet, new Set());
    } catch (error) {
      if (!(error instanceof CellError)) throw error;
      result = error;
    }
    this.cache.set(key, result);
    return result;
  }
  evaluate(node, sheet, stack) {
    if (node.type === 'number' || node.type === 'string' || node.type === 'boolean') return node.value;
    if (node.type === 'ref') {
      const value = this.value(node.sheet || sheet, node.address, stack);
      if (value instanceof CellError) throw value;
      return value;
    }
    if (node.type === 'unary') return (node.op === '-' ? -1 : 1) * numeric(this.evaluate(node.value, sheet, stack));
    if (node.type === 'binary') {
      if (node.op === ':') {
        if (node.left.type !== 'ref' || node.right.type !== 'ref') throw new Error('Invalid range');
        const target = node.left.sheet || node.right.sheet || sheet;
        const [firstCol, firstRow] = location(node.left.address);
        const [lastCol, lastRow] = location(node.right.address);
        const rows = [];
        for (let row = firstRow; row <= lastRow; row++) {
          const cells = [];
          for (let col = firstCol; col <= lastCol; col++) cells.push(this.value(target, `${columnLetters(col)}${row}`, stack));
          rows.push(cells);
        }
        return rows;
      }
      const left = this.evaluate(node.left, sheet, stack), right = this.evaluate(node.right, sheet, stack);
      if (node.op === '&') return text(left) + text(right);
      if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) {
        const [a, b] = comparable(left, right);
        return { '=': a === b, '<>': a !== b, '<': a < b, '>': a > b, '<=': a <= b, '>=': a >= b }[node.op];
      }
      const a = numeric(left), b = numeric(right);
      if (node.op === '+') return a + b;
      if (node.op === '-') return a - b;
      if (node.op === '*') return a * b;
      if (node.op === '/') { if (b === 0) throw new CellError('#DIV/0!'); return a / b; }
      if (node.op === '%') return a % b;
    }
    if (node.type === 'call') {
      const arg = index => this.evaluate(node.args[index], sheet, stack);
      const flat = () => node.args.flatMap((_, index) => { const v = arg(index); return Array.isArray(v) ? v.flat() : [v]; });
      // Like Sheets, SUM/MAX skip text and logical values inside ranges but not errors.
      const numbers = () => node.args.flatMap((_, index) => {
        const v = arg(index);
        if (!Array.isArray(v)) return [numeric(v)];
        return v.flat().filter(item => { if (item instanceof CellError) throw item; return typeof item === 'number'; });
      });
      if (node.name === 'IF') return arg(0) ? arg(1) : node.args.length > 2 ? arg(2) : false;
      if (node.name === 'IFERROR') {
        try { return arg(0); } catch (error) { if (error instanceof CellError) return arg(1); throw error; }
      }
      if (node.name === 'AND' || node.name === 'OR') {
        const values = flat().map(item => typeof item === 'boolean' ? item : numeric(item) !== 0);
        return node.name === 'AND' ? values.every(Boolean) : values.some(Boolean);
      }
      if (node.name === 'MAX') { const values = numbers(); return values.length ? Math.max(...values) : 0; }
      if (node.name === 'SUM') return numbers().reduce((sum, item) => sum + item, 0);
      if (node.name === 'ROUND') { const p = 10 ** numeric(arg(1)); return Math.round(numeric(arg(0)) * p) / p; }
      if (node.name === 'COUNTIF') {
        const criterion = arg(1);
        return arg(0).flat().filter(v => !(v instanceof CellError) && comparable(v, criterion).every((x, i, a) => x === a[0])).length;
      }
      if (node.name === 'VLOOKUP') {
        const needle = arg(0), matrix = arg(1), column = numeric(arg(2)) - 1;
        const found = matrix.find(row => { if (row[0] instanceof CellError) return false; const [a, b] = comparable(row[0], needle); return a === b; });
        if (!found) throw new CellError('#N/A');
        if (column < 0 || column >= found.length) throw new CellError('#REF!');
        if (found[column] instanceof CellError) throw found[column];
        return found[column];
      }
      throw new Error(`Unsupported function: ${node.name}`);
    }
    throw new Error(`Unsupported formula node: ${node.type}`);
  }
}

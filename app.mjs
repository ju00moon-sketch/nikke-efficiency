import { corrections } from './corrections.mjs';
import { CellError, Workbook, columnNumber } from './engine.mjs';
import { cellFormat, formatValue, inputFor, pickupConversion, rowTier, tierAt, timelineEntries } from './view-model.mjs';

const $ = selector => document.querySelector(selector);
const STORAGE_KEY = 'nikke-values-v1';
const THEME_KEY = 'nikke-theme-v1';
const PAGE = 150;
const COMPARISON_PAGE = 20;
const groups = [
  ['기준 · 모집', ['재화가치기준', '패키지효율', '픽업효율', 'AFFECTION TRAIL', '뽑기효율']],
  ['재화 · 패키지', ['코더', '스킬칩', '모듈', '유료쥬얼', '스텝업', '커스텀', '월정액', '코스튬뽑기(4주년)']],
  ['상점 · 기록', ['일반상점', '고철상점', '교환상점', '주년이벤트', '이벤트일정', 'CD-Keys']]
];
const descriptions = {
  '재화가치기준': '전초기지 레벨, 택틱 아카데미 효과와 재화별 가치를 바꾸면 모든 탭의 효율이 다시 계산됩니다.',
  '패키지효율': '원본의 전체 재화 환산 효율과 별도 픽업 단가를 함께 비교합니다. 체크한 상품만 누적 합계에 들어갑니다.',
  '픽업효율': '쥬얼과 고급 모집권을 기준으로 한 픽업 모집 비교입니다.',
  'AFFECTION TRAIL': '구매 선택과 반복 횟수에 따른 애정의 자취 보상을 확인합니다.',
  '뽑기효율': '일반 모집권을 포함한 전체 모집 기준 비교입니다.',
  '커스텀': '구성품 칸에서 원본과 같은 선택지를 고르면 패키지 효율이 다시 계산됩니다.',
  '월정액': '성장 보급상자에서 받을 품목을 고르면 30일 기준 효율이 다시 계산됩니다.',
  '이벤트일정': '원본의 가로 타임라인을 시작일과 종료일이 있는 목록으로 펼쳤습니다.',
  'CD-Keys': '원본에 기록된 교환 코드입니다. 현재 사용 가능 여부는 확인되지 않았습니다.'
};
const LABEL_IN_B = ['재화가치기준', '커스텀', '월정액', '코스튬뽑기(4주년)', '일반상점', '교환상점', 'CD-Keys'];
const TIERS = { best: '추천', good: '가성비', off: '미선택', done: '달성' };
const TIER_HELP = { best: '추천 효율 이상', good: '가성비 효율 이상', off: '체크 해제', done: '보상 기준 도달' };
const PICKUP = {
  PICKUP_PULLS: { label: '픽업 환산 횟수*', format: '#,##0.0' },
  PICKUP_COST: { label: '픽업 1회 단가*', format: '[$₩-412]#,##0' }
};
const SETTINGS = [
  ['G14', '전초기지 레벨'],
  ['H14', '택틱 아카데미 효과: 크레디트'],
  ['I14', '택틱 아카데미 효과: 배틀데이터셋'],
  ['J14', '택틱 아카데미 효과: 코어더스트'],
  ['E6', '재화가치 1당 가격(원)']
];
const EVENT_SORTS = [['', '원본 순서'], ['start', '시작일'], ['end', '종료일'], ['days', '기간'], ['category', '구분'], ['text', '내용']];
const CATEGORIES = [
  ['all', '전체', () => true],
  ['monthly', '월간·30일', name => /월간|30일|월정액/.test(name)],
  ['step', '스텝업', name => /스텝업|STEP/i.test(name)],
  ['new', '신입·복귀', name => /신입|복귀/.test(name)],
  ['anniversary', '주년', name => /주년|ANNIVERSARY|4TH/i.test(name)],
  ['popup', '팝업', name => /팝업/.test(name)]
];

let manifest, workbook, current = '패키지효율', activePage = 'efficiency', category = 'all', search = '', sortColumn = '', direction = 1, showHidden = true, limit = COMPARISON_PAGE;
let visibleRows = [];
const openComparisonRows = new Set();
const comparisonSlots = [null, null];
let valueListOpen = null;
const stepupOpenYears = new Set();
let stepupYearsInitialized = false;
const isComparison = () => activePage === 'efficiency' || activePage === 'compare';

function text(tag, value, className = '') {
  const el = document.createElement(tag);
  el.textContent = value;
  if (className) el.className = className;
  return el;
}
const sheetInfo = name => manifest.sheets.find(item => item.name === name);
const sheetData = name => workbook.sheets.get(name);
const colFromAddress = address => address.match(/^[A-Z]+/)[0];
const scrollBehavior = () => window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth';

function savedOverrides() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}'); } catch { return {}; }
}
function saveOverrides() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(Object.fromEntries(workbook.overrides))); } catch { /* private mode: edits last for this visit only */ }
}

function setTheme(theme) {
  document.documentElement.dataset.theme = theme;
  $('#theme-toggle').textContent = theme === 'dark' ? '라이트 모드' : '다크 모드';
  $('#theme-toggle').setAttribute('aria-label', theme === 'dark' ? '라이트 모드 켜기' : '다크 모드 켜기');
  document.querySelector('meta[name="theme-color"]').content = theme === 'dark' ? '#131820' : '#f5f6f8';
  try { localStorage.setItem(THEME_KEY, theme); } catch { /* temporary theme */ }
}

function renderTabs() {
  const picker = $('#sheet-picker');
  picker.replaceChildren();
  for (const [group, names] of groups) {
    const options = document.createElement('optgroup');
    options.label = group;
    for (const name of names) options.append(new Option(name, name));
    picker.append(options);
  }
  picker.value = current;
}

function chooseTab(name) {
  current = name;
  activePage = 'source';
  search = '';
  sortColumn = name === '이벤트일정' ? 'start' : '';
  direction = name === '이벤트일정' ? -1 : 1;
  $('#search').value = '';
  renderTabs();
  renderSheet();
  $('#workspace').scrollIntoView({ behavior: scrollBehavior(), block: 'start' });
}

function choosePage(page) {
  activePage = page;
  if (page === 'efficiency' || page === 'compare') current = '패키지효율';
  if (page === 'values') current = '재화가치기준';
  search = '';
  sortColumn = page === 'source' && current === '이벤트일정' ? 'start' : '';
  direction = page === 'source' && current === '이벤트일정' ? -1 : 1;
  $('#search').value = '';
  renderTabs();
  renderSheet();
  $('#workspace').scrollIntoView({ behavior: scrollBehavior(), block: 'start' });
}

function columnsFor(sheet) {
  const columns = new Set();
  for (const [, , cells] of sheet.rows) for (const cell of cells) columns.add(colFromAddress(cell[0]));
  const sorted = [...columns].sort((a, b) => columnNumber(a) - columnNumber(b));
  return sheet.name === '패키지효율' ? [...sorted, 'PICKUP_PULLS', 'PICKUP_COST'] : sorted;
}

function sourceHeader(sheet, column) {
  const header = sheet.rows.find(row => row[0] === 5)?.[2].find(cell => colFromAddress(cell[0]) === column);
  const label = header?.[1];
  return label && typeof label === 'string' && label.length < 20 ? `${column} · ${label}` : column;
}
const columnLabel = (sheet, column) => PICKUP[column]?.label ?? sourceHeader(sheet, column);

function valueAt(name, row, column) {
  if (PICKUP[column]) {
    if (row[0] < 6) return null;
    const pickup = pickupConversion(workbook, row[0]);
    return pickup && (column === 'PICKUP_PULLS' ? pickup.pulls : pickup.cost);
  }
  const cell = row[2].find(item => colFromAddress(item[0]) === column);
  return cell ? workbook.value(name, cell[0]) : null;
}

const isBlank = value => value == null || value === '' || value instanceof CellError;
function compareValues(a, b) {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  return String(a).localeCompare(String(b), 'ko', { numeric: true });
}
// Blank and error cells stay at the end in both directions.
function sortBy(read) {
  return (a, b) => {
    const x = read(a), y = read(b);
    if (isBlank(x) || isBlank(y)) return isBlank(x) === isBlank(y) ? 0 : isBlank(x) ? 1 : -1;
    return compareValues(x, y) * direction;
  };
}

function rowText(name, row) {
  const sheet = sheetData(name);
  return [row[0], ...row[2].map(cell => {
    const value = workbook.value(name, cell[0]);
    return value instanceof CellError ? value.code : `${formatValue(value, cellFormat(sheet, cell))} ${value}`;
  })].join(' ').toLocaleLowerCase();
}

function standardRows(sheet) {
  let rows = sheet.rows.filter(row => showHidden || !row[1]);
  if (search) rows = rows.filter(row => rowText(sheet.name, row).includes(search));
  if (!sortColumn) return rows;
  // Title and threshold rows above the data (rows 1 to 5) stay on top while sorting.
  const head = rows.filter(row => row[0] < 6);
  const order = sortBy(row => valueAt(sheet.name, row, sortColumn));
  const body = rows.filter(row => row[0] >= 6).sort((a, b) => order(a, b) || a[0] - b[0]);
  return [...head, ...body];
}

function control(sheet, cell, input, label) {
  const address = cell[0];
  const value = workbook.value(sheet.name, address);
  let element, options;
  if (input.kind === 'checkbox') {
    element = document.createElement('input');
    element.type = 'checkbox';
    element.checked = value === true;
  } else if (input.kind === 'select') {
    options = input.options.some(option => option.value === value) ? input.options : [{ value, label: formatValue(value, cellFormat(sheet, cell)) }, ...input.options];
    element = document.createElement('select');
    options.forEach((option, index) => element.add(new Option(option.label, String(index), false, option.value === value)));
  } else {
    element = document.createElement('input');
    element.type = 'number';
    element.step = 'any';
    element.inputMode = 'decimal';
    element.value = String(value);
  }
  element.dataset.cell = address;
  element.setAttribute('aria-label', label);
  element.addEventListener('change', () => {
    const area = element.closest('[data-area]')?.dataset.area;
    const card = element.closest('.package-card');
    if (card) openComparisonRows.add(Number(card.dataset.row));
    let next = null;
    if (input.kind === 'checkbox') next = element.checked;
    else if (input.kind === 'select') next = options[Number(element.value)].value;
    else if (element.value.trim() !== '' && Number.isFinite(Number(element.value))) next = Number(element.value);
    if (next !== null) {
      workbook.set(sheet.name, address, next);
      saveOverrides();
    }
    renderSheet(false);
    document.querySelector(`[data-area="${area}"] [data-cell="${address}"]`)?.focus({ preventScroll: true });
  });
  if (input.kind !== 'checkbox') return element;
  const hit = document.createElement('label');
  hit.className = 'check-hit';
  hit.append(element);
  return hit;
}

function tableHeader(columns, labels) {
  const thead = document.createElement('thead');
  const tr = document.createElement('tr');
  const corner = text('th', '행');
  corner.scope = 'col';
  tr.append(corner);
  for (const column of columns) {
    const th = text('th', labels.get(column));
    th.scope = 'col';
    th.title = PICKUP[column] ? '원본에 없는 보조 계산입니다. 표 아래 설명을 참고하세요.' : `원본 ${column}열`;
    tr.append(th);
  }
  thead.append(tr);
  return thead;
}

function renderCell(sheet, row, column, cell, rowLabel, heading, tier) {
  const td = document.createElement('td');
  if (column === (LABEL_IN_B.includes(sheet.name) ? 'B' : 'D')) td.classList.add('label-cell');
  if (PICKUP[column]) {
    const value = valueAt(sheet.name, row, column);
    if (value != null) td.append(formatValue(value, PICKUP[column].format));
    td.classList.add('numeric', 'derived');
    return td;
  }
  if (!cell) return td;
  const cellTier = tierAt(workbook, sheet, cell[0]);
  if (cellTier && cellTier !== tier) td.dataset.tier = cellTier;
  const input = inputFor(workbook, sheet, cell[0]);
  if (input) {
    td.classList.add('input-cell');
    td.append(control(sheet, cell, input, `${rowLabel} ${heading}`));
    return td;
  }
  const value = workbook.value(sheet.name, cell[0]);
  if (value instanceof CellError) {
    td.textContent = value.code;
    td.classList.add('cell-error');
    td.title = '바꾼 값에서는 원본 수식도 오류를 냅니다.';
    return td;
  }
  td.textContent = formatValue(value, cellFormat(sheet, cell));
  const correction = workbook.corrections.get(`${sheet.name}!${cell[0]}`);
  if (correction) {
    td.classList.add('corrected');
    td.title = correction.reason;
    td.append(text('small', '보정', 'cell-note'));
  }
  if (typeof value === 'number') {
    td.classList.add('numeric');
    if (heading.includes('효율')) td.classList.add('efficiency');
  }
  return td;
}

function renderRow(sheet, row, columns, labels) {
  const [number, folded, cells] = row;
  const tr = document.createElement('tr');
  const tier = rowTier(workbook, sheet, number);
  if (tier) tr.dataset.tier = tier;
  if (folded) tr.classList.add('folded');
  const head = document.createElement('th');
  head.scope = 'row';
  head.append(text('span', number));
  if (folded) head.append(text('small', '접힘', 'row-note'));
  if (tier) head.append(text('small', TIERS[tier], 'row-note tier-note'));
  tr.append(head);
  const byColumn = new Map(cells.map(cell => [colFromAddress(cell[0]), cell]));
  const labelCell = byColumn.get(LABEL_IN_B.includes(sheet.name) ? 'B' : 'D');
  const labelValue = labelCell && workbook.value(sheet.name, labelCell[0]);
  const rowLabel = typeof labelValue === 'string' ? labelValue.split('\n')[0] : `${number}행`;
  for (const column of columns) tr.append(renderCell(sheet, row, column, byColumn.get(column), rowLabel, labels.get(column), tier));
  return tr;
}

function legendText(sheet) {
  const parts = [];
  if (sheet.name === '패키지효율') parts.push('* 픽업 환산 횟수 = 쥬얼 ÷ 300 + 고급 모집권. 일반 모집권은 따로 세며, 원본 G·H열의 뽑기 횟수와 단가에는 일반 모집권이 더해져 있습니다.');
  const tiers = [...new Set((sheet.rules ?? []).map(rule => rule.tier))];
  if (tiers.length) parts.push(`행 표시는 원본의 조건부 서식 규칙을 따릅니다: ${tiers.map(tier => `${TIERS[tier]}(${TIER_HELP[tier]})`).join(', ')}. 기준 효율은 표 위쪽 행에서 바꿀 수 있습니다.`);
  if (sheet.rows.some(row => row[1])) parts.push('‘접힘’ 행은 원본에서 접혀 있던 행입니다.');
  const reasons = new Map();
  for (const item of workbook.corrections.values()) if (item.sheet === sheet.name) reasons.set(item.reason, [...(reasons.get(item.reason) ?? []), item.cell]);
  for (const [reason, cells] of reasons) parts.push(`‘보정’ ${cells.join(', ')}: ${reason}`);
  if (reasons.size) parts.push('보정한 칸도 기본값에서는 원본과 결과가 같습니다.');
  const editable = sheet.inputs?.length || sheet.rows.some(row => row[2].some(cell => cell[3] === 'b' && !cell[2]));
  if (!parts.length && editable) parts.push('입력칸과 체크박스로 바꾼 값은 이 브라우저에만 저장됩니다.');
  return parts.join(' ');
}

function showTable(table, unit, noun) {
  const area = $('#table-area');
  const pageSize = isComparison() ? COMPARISON_PAGE : PAGE;
  if (visibleRows.length) {
    area.replaceChildren(table);
    if (limit < visibleRows.length) {
      const more = text('button', `다음 ${Math.min(pageSize, visibleRows.length - limit)}${unit} 보기 ↓`, 'more-button tool-button');
      more.type = 'button';
      more.addEventListener('click', () => { limit += pageSize; renderSheet(false); });
      area.append(more);
    }
  } else {
    area.replaceChildren(text('p', search ? `검색어 ‘${search}’에 맞는 ${noun}이 없습니다.` : `표시할 ${noun}이 없습니다.`, 'empty-state'));
  }
  const total = visibleRows.length.toLocaleString('ko-KR');
  $('#result-count').textContent = `${total}${unit} 중 ${Math.min(limit, visibleRows.length).toLocaleString('ko-KR')}${unit} 표시`;
}

function packageValue(sheet, row, column) {
  const value = valueAt(sheet.name, row, column);
  if (value == null || value === '') return '—';
  if (value instanceof CellError) return value.code;
  const cell = row[2].find(item => colFromAddress(item[0]) === column);
  return formatValue(value, PICKUP[column]?.format ?? cellFormat(sheet, cell));
}

function packageMetric(label, value, className = '') {
  const metric = document.createElement('div');
  metric.className = `package-metric ${className}`;
  metric.append(text('span', label), text('strong', value));
  return metric;
}

function packageCard(sheet, row, cumulativePrice = null) {
  const [number, folded, cells] = row;
  const byColumn = new Map(cells.map(cell => [colFromAddress(cell[0]), cell]));
  const name = String(valueAt(sheet.name, row, 'D'));
  const tier = rowTier(workbook, sheet, number);
  const card = document.createElement('article');
  card.className = 'package-card';
  card.dataset.area = 'comparison';
  card.dataset.row = number;
  card.setAttribute('role', 'listitem');
  if (tier) card.dataset.tier = tier;
  const header = document.createElement('div');
  header.className = 'package-card-head';
  const title = document.createElement('div');
  title.className = 'package-title';
  const badges = text('div', `${number}행${folded ? ' · 원본 접힘' : ''}${tier ? ` · ${TIERS[tier]}` : ''}`, 'package-meta');
  title.append(badges, text('h3', name));
  header.append(title, packageMetric('효율', packageValue(sheet, row, 'I'), 'package-efficiency'));
  const metrics = document.createElement('div');
  metrics.className = 'package-metrics';
  metrics.append(packageMetric('가격', packageValue(sheet, row, 'F')), packageMetric('픽업 1회 단가', packageValue(sheet, row, 'PICKUP_COST')));
  const composition = packageValue(sheet, row, 'E');
  if (composition !== '—') card.append(header, metrics, text('p', composition, 'package-composition'));
  else card.append(header, metrics);
  if (cumulativePrice != null) card.append(text('p', `이 단계까지 누적 결제 ${formatValue(cumulativePrice, '[$₩-412]#,##0')} · 원본 스텝업 탭 기준`, 'stepup-cumulative'));
  const compare = document.createElement('label');
  compare.className = 'compare-picker';
  const check = document.createElement('input');
  check.type = 'checkbox';
  check.checked = comparisonSlots.includes(number);
  check.disabled = !check.checked && comparisonSlots.every(Boolean);
  check.setAttribute('aria-label', `${name} 비교 대상`);
  check.addEventListener('change', () => {
    if (check.checked) comparisonSlots[comparisonSlots.indexOf(null)] = number;
    else comparisonSlots[comparisonSlots.indexOf(number)] = null;
    renderSheet(false);
  });
  compare.append(check, text('span', check.checked ? '비교 대상에 담김' : '비교 대상에 담기'));
  card.append(compare);
  const details = document.createElement('details');
  details.className = 'package-detail';
  details.open = openComparisonRows.has(number);
  details.append(text('summary', '구매 설정과 재화 상세'));
  details.addEventListener('toggle', () => {
    if (!details.isConnected) return;
    if (details.open) openComparisonRows.add(number);
    else openComparisonRows.delete(number);
  });
  const detailBody = document.createElement('div');
  detailBody.className = 'package-detail-body';
  for (const [column, label] of [['B', '구매 합계에 포함'], ['C', '반복 구매']]) {
    const cell = byColumn.get(column), input = cell && inputFor(workbook, sheet, cell[0]);
    if (!input) continue;
    const field = document.createElement('div');
    field.className = 'package-setting';
    field.append(text('span', label), control(sheet, cell, input, `${name} ${label}`));
    detailBody.append(field);
  }
  const detailMetrics = document.createElement('div');
  detailMetrics.className = 'package-detail-metrics';
  for (const [column, label] of [['G', '원본 뽑기횟수'], ['PICKUP_PULLS', '픽업 환산 횟수'], ['M', '쥬얼'], ['N', '일반 모집권'], ['O', '고급 모집권']]) {
    if (column !== 'PICKUP_PULLS' && !byColumn.has(column)) continue;
    detailMetrics.append(packageMetric(label, packageValue(sheet, row, column)));
  }
  detailBody.append(detailMetrics);
  details.append(detailBody);
  card.append(details);
  return card;
}

function productRows(sheet) {
  return sheet.rows.filter(row => row[0] >= 6 && typeof valueAt(sheet.name, row, 'D') === 'string' && typeof valueAt(sheet.name, row, 'F') === 'number');
}

function renderCategoryFilters(sheet) {
  const panel = $('#category-filter');
  panel.replaceChildren(text('span', '상품명 기준 분류', 'filter-label'));
  const rows = productRows(sheet);
  for (const [key, label, matches] of CATEGORIES) {
    const count = rows.filter(row => matches(valueAt(sheet.name, row, 'D'))).length;
    const button = text('button', `${label} ${count}`, 'category-chip');
    button.type = 'button';
    button.setAttribute('aria-pressed', String(category === key));
    button.addEventListener('click', () => { category = key; if (key === 'step') { sortColumn = ''; direction = 1; } renderSheet(); });
    panel.append(button);
  }
}

function stepNumber(sheet, row) {
  return Number(String(valueAt(sheet.name, row, 'D') ?? '').match(/STEP\s*(\d+)/i)?.[1] ?? 0);
}

function stepupReferenceRows(sheet) {
  const reference = sheetData('스텝업');
  const matches = new Map();
  for (const row of reference.rows) {
    const step = stepNumber(reference, row);
    if (!step) continue;
    const price = valueAt(reference.name, row, 'E');
    const efficiency = valueAt(reference.name, row, 'G');
    const found = sheet.rows.find(candidate => stepNumber(sheet, candidate) === step
      && valueAt(sheet.name, candidate, 'F') === price
      && Math.abs(valueAt(sheet.name, candidate, 'I') - efficiency) < 1e-7);
    if (found) matches.set(found[0], valueAt(reference.name, row, 'F'));
  }
  return matches;
}

function renderStepupPanel() {
  const panel = $('#stepup-panel');
  panel.replaceChildren();
  if (!isComparison() || category !== 'step') return;
  const periods = timelineEntries(workbook, sheetData('이벤트일정'))
    .filter(item => item.category === '패키지' && /^스텝업 패키지/.test(String(item.value)))
    .sort((a, b) => b.start.localeCompare(a.start));
  const heading = text('h3', `원본 일정의 스텝업 패키지 기간 · ${periods.length}건`);
  panel.append(heading, text('p', '아래 날짜는 이벤트일정 탭에 기록된 시작일과 마지막 날짜입니다. 각 STEP 상품 행과 기간의 연결은 원본에 없어 임의로 배정하지 않았습니다. 현재 판매 여부도 확인되지 않았습니다.', 'stepup-explanation'));
  if (!stepupYearsInitialized && periods.length) {
    stepupOpenYears.add(Number(periods[0].start.slice(0, 4)));
    stepupYearsInitialized = true;
  }
  const years = new Map();
  for (const period of periods) {
    const year = Number(period.start.slice(0, 4));
    if (!years.has(year)) years.set(year, []);
    years.get(year).push(period);
  }
  for (const [year, entries] of years) {
    const details = document.createElement('details');
    details.className = 'stepup-year';
    details.open = stepupOpenYears.has(year);
    details.addEventListener('toggle', () => {
      if (!details.isConnected) return;
      if (details.open) stepupOpenYears.add(year); else stepupOpenYears.delete(year);
    });
    details.append(text('summary', `${year}년 · ${entries.length}건`));
    const list = document.createElement('ul');
    list.className = 'stepup-periods';
    for (const item of entries) {
      const entry = document.createElement('li');
      entry.append(text('strong', `${item.start} ~ ${item.end}`), text('span', `${item.value} · 이벤트일정 ${item.address}`));
      list.append(entry);
    }
    details.append(list);
    panel.append(details);
  }
}

function renderComparisonPanel(sheet) {
  const panel = $('#compare-panel');
  panel.replaceChildren(text('p', '같은 재화 가치 기준으로 두 상품을 비교합니다. 아래 상품 목록에서 비교 대상을 고르거나 선택창을 사용하세요.', 'compare-help'));
  const rows = productRows(sheet);
  const choices = document.createElement('div');
  choices.className = 'compare-choices';
  for (let slot = 0; slot < 2; slot++) {
    const field = document.createElement('label');
    field.append(text('span', `비교 상품 ${slot + 1}`));
    const select = document.createElement('select');
    select.setAttribute('aria-label', `비교 상품 ${slot + 1}`);
    select.add(new Option('상품을 선택하세요', ''));
    for (const row of rows) select.add(new Option(`${valueAt(sheet.name, row, 'D')} · ${row[0]}행`, String(row[0])));
    select.value = comparisonSlots[slot] == null ? '' : String(comparisonSlots[slot]);
    select.addEventListener('change', () => {
      const next = select.value ? Number(select.value) : null;
      const other = 1 - slot;
      if (next != null && comparisonSlots[other] === next) comparisonSlots[other] = null;
      comparisonSlots[slot] = next;
      renderSheet(false);
    });
    field.append(select);
    choices.append(field);
  }
  panel.append(choices);
  const results = document.createElement('div');
  results.className = 'compare-results';
  for (let slot = 0; slot < 2; slot++) {
    const row = rows.find(item => item[0] === comparisonSlots[slot]);
    const result = document.createElement('div');
    result.className = 'compare-result';
    if (!row) result.append(text('p', `비교 상품 ${slot + 1}을 선택하세요.`, 'compare-empty'));
    else {
      result.append(text('small', `비교 상품 ${slot + 1} · 원본 ${row[0]}행`), text('h3', valueAt(sheet.name, row, 'D')));
      const metrics = document.createElement('div');
      metrics.className = 'compare-result-metrics';
      for (const [column, label] of [['F', '가격'], ['I', '전체 재화 효율'], ['PICKUP_PULLS', '픽업 환산 횟수'], ['PICKUP_COST', '픽업 1회 단가']]) metrics.append(packageMetric(label, packageValue(sheet, row, column)));
      result.append(metrics, text('p', packageValue(sheet, row, 'E'), 'compare-composition'));
    }
    results.append(result);
  }
  panel.append(results, text('p', '픽업 환산은 쥬얼 ÷ 300 + 고급 모집권으로 계산하며 일반 모집권은 합산하지 않습니다. 판매 여부는 확인되지 않았습니다.', 'compare-footnote'));
}

function renderPackageComparison(sheet) {
  const sort = $('#sort-column');
  const isStepup = category === 'step';
  sort.replaceChildren(...(isStepup ? [['', '묶음별 STEP 순서']] : [['', '원본 순서'], ['D', '상품명'], ['F', '가격'], ['I', '효율'], ['PICKUP_COST', '픽업 1회 단가']]).map(([value, label]) => new Option(label, value)));
  sort.value = sortColumn;
  sort.disabled = isStepup;
  $('#sort-direction').disabled = isStepup;
  const matchesCategory = CATEGORIES.find(([key]) => key === category)?.[2] ?? CATEGORIES[0][2];
  visibleRows = standardRows(sheet).filter(row => row[0] >= 6 && typeof valueAt(sheet.name, row, 'D') === 'string' && typeof valueAt(sheet.name, row, 'F') === 'number' && matchesCategory(valueAt(sheet.name, row, 'D')));
  if (isStepup) {
    const reference = stepupReferenceRows(sheet);
    const byStep = (a, b) => stepNumber(sheet, a) - stepNumber(sheet, b) || a[0] - b[0];
    const currentRows = visibleRows.filter(row => reference.has(row[0])).sort(byStep);
    const otherRows = visibleRows.filter(row => !reference.has(row[0])).sort(byStep);
    visibleRows = [...currentRows, ...otherRows];
    const groups = document.createElement('div');
    groups.className = 'stepup-product-groups';
    for (const [title, note, rows] of [
      ['스텝업 탭에 정리된 구성', '원본 스텝업 탭과 단계·가격·효율이 일치하는 5개 상품입니다. 각 카드의 가격은 해당 STEP 단일 가격, 누적 결제는 그 단계까지의 합계입니다. 판매기간 연결은 확인되지 않았습니다.', currentRows],
      ['패키지효율 탭의 다른 구성', '같은 STEP 이름이라도 구성품과 원본 행이 다릅니다. 판매기간은 확인되지 않아 날짜별 일정에 배정하지 않았습니다.', otherRows]
    ]) {
      if (!rows.length) continue;
      const section = document.createElement('section');
      section.className = 'stepup-product-group';
      section.append(text('h3', `${title} · ${rows.length}개`), text('p', note, 'stepup-group-note'));
      const cards = document.createElement('div');
      cards.className = 'package-list';
      cards.setAttribute('role', 'list');
      for (const row of rows) cards.append(packageCard(sheet, row, reference.get(row[0])));
      section.append(cards);
      groups.append(section);
    }
    showTable(groups, '개', '상품');
    $('#legend').textContent = legendText(sheet);
    return;
  }
  const list = document.createElement('div');
  list.className = 'package-list';
  list.setAttribute('role', 'list');
  list.setAttribute('aria-label', '패키지 상품 비교');
  for (const row of visibleRows.slice(0, limit)) list.append(packageCard(sheet, row));
  showTable(list, '개', '상품');
  $('#legend').textContent = legendText(sheet);
}

function renderStandard(sheet) {
  if (isComparison()) return renderPackageComparison(sheet);
  $('#sort-column').disabled = false;
  $('#sort-direction').disabled = false;
  const columns = columnsFor(sheet);
  const labels = new Map(columns.map(column => [column, columnLabel(sheet, column)]));
  const sort = $('#sort-column');
  sort.replaceChildren(new Option('원본 순서', ''), ...columns.map(column => new Option(labels.get(column), column)));
  sort.value = sortColumn;
  visibleRows = standardRows(sheet);
  const table = document.createElement('table');
  table.className = 'data-table';
  table.dataset.area = 'table';
  table.setAttribute('aria-label', `${sheet.name} 원본 표`);
  table.append(tableHeader(columns, labels));
  const tbody = document.createElement('tbody');
  for (const row of visibleRows.slice(0, limit)) tbody.append(renderRow(sheet, row, columns, labels));
  table.append(tbody);
  showTable(table, '행', '행');
  $('#legend').textContent = legendText(sheet);
}

function renderEvents(sheet) {
  const sort = $('#sort-column');
  sort.replaceChildren(...EVENT_SORTS.map(([value, label]) => new Option(label, value)));
  sort.value = sortColumn;
  visibleRows = timelineEntries(workbook, sheet)
    .map(entry => ({ ...entry, text: formatValue(entry.value) }))
    .filter(item => !search || `${item.start} ${item.end} ${item.category} ${item.text} ${item.address}`.toLocaleLowerCase().includes(search));
  if (sortColumn) {
    const order = sortBy(item => item[sortColumn]);
    visibleRows.sort((a, b) => order(a, b) || a.start.localeCompare(b.start) || a.row - b.row);
  }
  const table = document.createElement('table');
  table.className = 'data-table event-table';
  table.setAttribute('aria-label', '이벤트 일정 목록');
  const thead = document.createElement('thead');
  const top = document.createElement('tr');
  for (const label of ['시작일', '종료일', '기간', '구분', '내용', '원본 셀']) {
    const th = text('th', label);
    th.scope = 'col';
    top.append(th);
  }
  thead.append(top);
  table.append(thead);
  const tbody = document.createElement('tbody');
  for (const item of visibleRows.slice(0, limit)) {
    const row = document.createElement('tr');
    row.append(text('td', item.start, 'date-cell'), text('td', item.end, 'date-cell'), text('td', `${item.days}일`, 'numeric'), text('td', item.category), text('td', item.text, 'event-name'), text('td', item.address));
    tbody.append(row);
  }
  table.append(tbody);
  showTable(table, '개', '기록');
  $('#legend').textContent = '원본의 날짜 가로축(2022.11.04~2026.12.31)에서 병합된 칸의 처음과 끝을 시작일과 종료일로 옮겼습니다. 과거와 예정 일정이 함께 들어 있는 스냅샷입니다.';
}

function valueItem(sheet, address, label) {
  const cell = workbook.raw(sheet.name, address);
  const input = cell && inputFor(workbook, sheet, address);
  if (!input) return null;
  const item = document.createElement('div');
  item.className = 'value-item';
  item.append(text('span', label), control(sheet, cell, input, label));
  return item;
}

function renderValueTools() {
  const target = $('#value-tools');
  target.replaceChildren();
  if (activePage !== 'values') return;
  const sheet = sheetData(current);
  const panel = document.createElement('section');
  panel.className = 'value-panel';
  panel.dataset.area = 'panel';
  panel.setAttribute('aria-label', '내 기준값');
  const settings = document.createElement('div');
  settings.className = 'value-grid';
  settings.append(...SETTINGS.map(([address, label]) => valueItem(sheet, address, label)).filter(Boolean));
  const values = document.createElement('div');
  values.className = 'value-grid';
  for (let row = 6; row <= 72; row++) {
    const name = workbook.value(current, `B${row}`);
    const item = typeof name === 'string' && valueItem(sheet, `D${row}`, name);
    if (item) values.append(item);
  }
  const list = document.createElement('details');
  list.className = 'value-list';
  list.open = valueListOpen ?? window.matchMedia('(min-width: 861px)').matches;
  list.addEventListener('toggle', () => { if (list.isConnected) valueListOpen = list.open; });
  list.append(text('summary', `재화별 가치 ${values.childElementCount}개`), values);
  panel.append(text('h3', '내 기준값'), text('p', '값을 바꾸면 모든 탭의 계산에 바로 반영되고 이 브라우저에만 저장됩니다.'), settings, list);
  target.append(panel);
}

function renderOverrideNotice() {
  const notice = $('#override-notice');
  const count = workbook.overrides.size;
  notice.hidden = count === 0;
  if (!count) return;
  const reset = text('button', '원본 값으로 되돌리기', 'tool-button');
  reset.type = 'button';
  reset.addEventListener('click', () => {
    workbook.reset();
    saveOverrides();
    renderSheet(false);
    $('#section-title').focus();
  });
  notice.replaceChildren(text('span', `이 브라우저에서 바꾼 값 ${count}개가 계산에 반영되어 있어 원본 스냅샷과 다를 수 있습니다.`), reset);
}

function renderSheet(resetLimit = true) {
  if (resetLimit) limit = isComparison() ? COMPARISON_PAGE : PAGE;
  const sheet = sheetData(current), info = sheetInfo(current);
  for (const button of document.querySelectorAll('.primary-nav [data-page]')) {
    if (button.dataset.page === activePage) button.setAttribute('aria-current', 'page');
    else button.removeAttribute('aria-current');
  }
  const compareCount = comparisonSlots.filter(Boolean).length;
  $('#compare-count').hidden = compareCount === 0;
  $('#compare-count').textContent = compareCount;
  $('#source-tools').hidden = activePage !== 'source';
  $('#compare-panel').hidden = activePage !== 'compare';
  $('#toolbar').hidden = activePage === 'values';
  $('#category-filter').hidden = !isComparison();
  $('#stepup-panel').hidden = !isComparison() || category !== 'step';
  $('#list-heading').hidden = activePage === 'values';
  $('#legend-details').hidden = activePage === 'values';
  $('#table-area').hidden = activePage === 'values';
  $('#table-footer').hidden = activePage === 'values';
  $('#table-area').classList.toggle('comparison-area', isComparison());
  $('#section-title').textContent = { efficiency: '패키지 효율표', compare: '두 상품 비교', values: '재화 가치 설정', source: current }[activePage];
  $('#section-description').textContent = {
    efficiency: '상품명, 가격과 전체 재화 효율을 한눈에 보고 구성품과 구매 설정을 펼쳐 볼 수 있습니다.',
    compare: '두 상품의 가격, 효율과 픽업 환산을 같은 기준으로 나란히 확인합니다.',
    values: '재화 가치를 바꾸면 효율표와 비교 결과가 즉시 다시 계산됩니다.',
    source: descriptions[current] || '공개 원본 탭의 행과 열, 계산 결과를 그대로 볼 수 있습니다.'
  }[activePage];
  $('#section-kicker').textContent = activePage === 'source' ? groups.find(([, names]) => names.includes(current))[0] : '2026.11.02 원본 기준';
  $('#sheet-count').textContent = `${info.rowCount.toLocaleString('ko-KR')}행 · ${info.cellCount.toLocaleString('ko-KR')}셀`;
  $('#sort-direction').textContent = direction === 1 ? '오름차순 ↑' : '내림차순 ↓';
  $('#hidden-wrap').hidden = !info.hiddenRowCount;
  $('#list-note').textContent = isComparison() ? '판매 여부 미확인 · 상품명 기준 분류' : '원본 스냅샷 · 자동 갱신 없음';
  renderOverrideNotice();
  renderValueTools();
  if (activePage === 'values') return;
  if (isComparison()) renderCategoryFilters(sheet);
  renderStepupPanel();
  if (activePage === 'compare') renderComparisonPanel(sheet);
  if (current === '이벤트일정') renderEvents(sheet); else renderStandard(sheet);
}

function downloadCsv() {
  let headings, records;
  if (current === '이벤트일정') {
    headings = ['시작일', '종료일', '기간(일)', '구분', '내용', '원본 셀'];
    records = visibleRows.map(item => [item.start, item.end, item.days, item.category, item.text, item.address]);
  } else {
    const sheet = sheetData(current);
    const columns = columnsFor(sheet);
    headings = ['행', ...columns.map(column => columnLabel(sheet, column))];
    records = visibleRows.map(row => [row[0], ...columns.map(column => {
      const value = valueAt(current, row, column);
      return value instanceof CellError ? value.code : value ?? '';
    })]);
  }
  const escape = value => `"${String(value).replaceAll('"', '""')}"`;
  const csv = '\uFEFF' + [headings, ...records].map(row => row.map(escape).join(',')).join('\r\n');
  const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = `nikke-${current}-2026-11-02.csv`;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function start() {
  try { setTheme(localStorage.getItem(THEME_KEY) === 'dark' ? 'dark' : 'light'); } catch { setTheme('light'); }
  manifest = await fetch('./snapshot/manifest.json').then(response => { if (!response.ok) throw new Error('목록을 불러올 수 없습니다.'); return response.json(); });
  const sheets = await Promise.all(manifest.sheets.map(async info => {
    const response = await fetch(`./snapshot/${info.file}`);
    if (!response.ok) throw new Error(`${info.name} 자료를 불러올 수 없습니다.`);
    return response.json();
  }));
  workbook = new Workbook(sheets, corrections);
  for (const [key, value] of Object.entries(savedOverrides())) {
    const split = key.lastIndexOf('!');
    const name = key.slice(0, split), address = key.slice(split + 1);
    const sheet = workbook.sheets.get(name);
    if (sheet && inputFor(workbook, sheet, address)) workbook.set(name, address, value);
  }
  const visible = manifest.sheets.filter(info => info.visible);
  $('#stat-tabs').textContent = visible.length;
  $('#stat-rows').textContent = visible.reduce((n, item) => n + item.rowCount, 0).toLocaleString('ko-KR');
  $('#stat-cells').textContent = visible.reduce((n, item) => n + item.sourceCellCount, 0).toLocaleString('ko-KR');
  renderTabs();
  renderSheet();
  document.querySelectorAll('.primary-nav [data-page]').forEach(button => button.addEventListener('click', () => choosePage(button.dataset.page)));
  $('#sheet-picker').addEventListener('change', event => chooseTab(event.target.value));
  $('#theme-toggle').addEventListener('click', () => setTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'));
  $('#search').addEventListener('input', event => { search = event.target.value.trim().toLocaleLowerCase(); renderSheet(); });
  $('#sort-column').addEventListener('change', event => { sortColumn = event.target.value; renderSheet(); });
  $('#sort-direction').addEventListener('click', () => { direction *= -1; renderSheet(); });
  $('#show-hidden').addEventListener('change', event => { showHidden = event.target.checked; renderSheet(); });
  $('#download-csv').addEventListener('click', downloadCsv);
}

start().catch(error => {
  $('#section-title').textContent = '자료를 불러오지 못했습니다';
  $('#table-area').replaceChildren(text('p', `${error.message} 페이지를 새로고침해 주세요.`, 'loading'));
});

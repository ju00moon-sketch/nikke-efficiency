import { products, pickupPulls, wonPerPickup, orderByPickupCost } from './data.mjs';

const won = value => `${new Intl.NumberFormat('ko-KR', { maximumFractionDigits: 0 }).format(value)}원`;
const number = value => new Intl.NumberFormat('ko-KR').format(value);
const pulls = value => `${new Intl.NumberFormat('ko-KR', { maximumFractionDigits: 1 }).format(value)}회`;

function node(tag, className, content) {
  const element = document.createElement(tag);
  if (className) element.className = className;
  if (content != null) element.textContent = content;
  return element;
}

function renderAnniversary() {
  const container = document.querySelector('#anniversary-cards');
  for (const product of products.filter(item => item.category === 'anniversary')) {
    const card = node('article', 'anniversary-card');
    const top = node('div', 'card-top');
    top.append(node('span', 'card-tag', product.name), node('strong', 'card-price', won(product.price)));
    const metric = node('div', 'card-metric');
    metric.append(node('strong', '', pulls(pickupPulls(product))), node('span', '', '픽업 환산'));
    const detail = node('p', 'card-detail', `픽업 1회당 ${won(wonPerPickup(product))} · 일반권 ${product.ordinary}장`);
    const extras = node('p', 'card-extras', product.extras);
    card.append(top, metric, detail, extras);
    container.append(card);
  }
}

function cell(label, content, className = '') {
  const td = node('td', className, content);
  td.dataset.label = label;
  return td;
}

function renderRows(filter) {
  const body = document.querySelector('#products-body');
  const visible = orderByPickupCost(products.filter(item => filter === 'all' || item.category === filter));
  body.replaceChildren();
  for (const product of visible) {
    const row = document.createElement('tr');
    const nameCell = cell('상품 / 수령 조건', '', 'name-cell');
    nameCell.append(node('strong', '', product.name), node('small', '', product.condition));
    const pickupCost = wonPerPickup(product);
    row.append(
      nameCell,
      cell('가격', won(product.price), 'numeric'),
      cell('쥬얼', number(product.gems), 'numeric'),
      cell('고급권', `${product.advanced}장`, 'numeric'),
      cell('픽업 환산', pulls(pickupPulls(product)), 'numeric emphasis'),
      cell('픽업 1회당', pickupCost === null ? '해당 없음' : won(pickupCost), 'numeric cost'),
      cell('일반권', `${product.ordinary}장`, 'numeric'),
    );
    body.append(row);
  }
  document.querySelector('#filter-status').textContent = `${visible.length}개 상품 표시`;
}

renderAnniversary();
document.querySelector('#all-count').textContent = products.length;
renderRows('all');
for (const button of document.querySelectorAll('.filter')) {
  button.addEventListener('click', () => {
    for (const item of document.querySelectorAll('.filter')) {
      const active = item === button;
      item.classList.toggle('active', active);
      item.setAttribute('aria-pressed', String(active));
    }
    renderRows(button.dataset.filter);
  });
}

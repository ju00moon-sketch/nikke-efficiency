import assert from 'node:assert/strict';
import test from 'node:test';
import { products, pickupPulls, wonPerPickup, orderByPickupCost } from './data.mjs';

const byId = Object.fromEntries(products.map(item => [item.id, item]));

test('4주년 패키지 픽업 환산에 일반 모집권을 포함하지 않는다', () => {
  assert.equal(pickupPulls(byId['anniversary-1']), 10);
  assert.equal(pickupPulls(byId['anniversary-2']), 24);
  assert.ok(Math.abs(pickupPulls(byId['anniversary-3']) - 36.6666666667) < 1e-9);
  assert.equal(byId['anniversary-1'].ordinary, 10);
  assert.equal(byId['anniversary-2'].ordinary, 20);
  assert.equal(byId['anniversary-3'].ordinary, 20);
});

test('비교 상품의 픽업 환산과 1회당 원가', () => {
  assert.equal(pickupPulls(byId['special-1']), 20);
  assert.equal(wonPerPickup(byId['special-1']), 1500);
  assert.equal(pickupPulls(byId.monthly), 11.1);
  assert.ok(Math.abs(wonPerPickup(byId.monthly) - 675.6756756757) < 1e-9);
  assert.ok(Math.abs(pickupPulls(byId['campaign-1']) - 33.3333333333) < 1e-9);
  assert.equal(wonPerPickup(byId['new-commander-1']), null);
});

test('픽업 효율 정렬에서 픽업 횟수가 없는 상품은 맨 뒤에 둔다', () => {
  const sorted = orderByPickupCost(products);
  assert.equal(sorted[0].id, 'new-commander-attendance');
  assert.equal(sorted.at(-1).id, 'new-commander-1');
  assert.ok(sorted.slice(0, -1).every((p, i, arr) =>
    i === 0 || wonPerPickup(arr[i - 1]) <= wonPerPickup(p)));
});

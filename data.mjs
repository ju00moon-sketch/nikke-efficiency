export const SOURCE_URL = 'https://docs.google.com/spreadsheets/d/1qUSwYnpDtWAD_aq_4L6dC6py2vM19WtLed2AGKSdIPk/edit?gid=1190480704#gid=1190480704';
export const RANKING_URL = 'https://docs.google.com/spreadsheets/d/1qUSwYnpDtWAD_aq_4L6dC6py2vM19WtLed2AGKSdIPk/edit?gid=755822179#gid=755822179';
export const CHECKED_DATE = '2026-11-02';

export const products = [
  {
    id: 'anniversary-1', name: '4주년 패키지 Ⅰ', category: 'anniversary',
    price: 30000, gems: 1500, advanced: 5, ordinary: 10,
    extras: '코어 더스트 1시간 상자 180개',
    condition: '4주년 기간 한정 · 판매 여부 게임 내 확인',
  },
  {
    id: 'anniversary-2', name: '4주년 패키지 Ⅱ', category: 'anniversary',
    price: 85000, gems: 4200, advanced: 10, ordinary: 20,
    extras: '스킬·버스트 선택 상자 각 30개, 코어·크레디트·배틀 데이터 1시간 상자 각 300개',
    condition: '4주년 기간 한정 · 판매 여부 게임 내 확인',
  },
  {
    id: 'anniversary-3', name: '4주년 패키지 Ⅲ', category: 'anniversary',
    price: 149000, gems: 8000, advanced: 10, ordinary: 20,
    extras: 'T9 기업 장비 선택 상자 4개, 모듈 50개, 커스텀 락키 200개',
    condition: '4주년 기간 한정 · 판매 여부 게임 내 확인',
  },
  {
    id: 'special-1', name: '특수 모집 패키지 Ⅰ', category: 'comparison',
    price: 30000, gems: 1500, advanced: 15, ordinary: 0,
    extras: '픽업 모집 중심 비교 상품',
    condition: '판매 여부 게임 내 확인',
  },
  {
    id: 'monthly', name: '30일 보급품', category: 'comparison',
    price: 7500, gems: 3330, advanced: 0, ordinary: 0,
    extras: '유료 쥬얼 330개 + 무료 쥬얼 3,000개',
    condition: '30일 동안 모두 수령한 경우 · 즉시 전량 지급 아님',
  },
  {
    id: 'campaign-1', name: '캠페인 패키지 Ⅰ', category: 'comparison',
    price: 30000, gems: 10000, advanced: 0, ordinary: 0,
    extras: '무료 쥬얼 10,000개',
    condition: '캠페인 진행 조건 충족 필요',
  },
  {
    id: 'campaign-2', name: '캠페인 패키지 Ⅱ', category: 'comparison',
    price: 30000, gems: 10000, advanced: 0, ordinary: 0,
    extras: '무료 쥬얼 10,000개',
    condition: '캠페인 진행 조건 충족 필요',
  },
  {
    id: 'new-commander-1', name: '신입 지휘관 패키지 Ⅰ', category: 'comparison',
    price: 1500, gems: 0, advanced: 0, ordinary: 10,
    extras: '일반 모집권 10장',
    condition: '신입 지휘관 대상 · 판매 여부 게임 내 확인',
  },
  {
    id: 'new-commander-attendance', name: '신입 지휘관 출석 패키지', category: 'comparison',
    price: 6000, gems: 0, advanced: 10, ordinary: 10,
    extras: '보급물자 Ⅰ 200개',
    condition: '신입 지휘관 대상 · 출석 수령 조건',
  },
];

export function pickupPulls(product) {
  return product.advanced + product.gems / 300;
}

export function wonPerPickup(product) {
  const pulls = pickupPulls(product);
  return pulls === 0 ? null : product.price / pulls;
}

export function orderByPickupCost(items) {
  return [...items].sort((a, b) =>
    (wonPerPickup(a) ?? Infinity) - (wonPerPickup(b) ?? Infinity) || a.price - b.price);
}

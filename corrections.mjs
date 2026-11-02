// Formula fixes applied on top of the unchanged snapshot. `source` must equal the snapshot formula,
// and every fix gives the snapshot's own result with the source's default inputs.
const stepUpCumulative = row => ({
  sheet: '스텝업',
  cell: `H${row}`,
  source: `IF(B${row}=TRUE,H${row - 1}*F${row - 1}/F${row}+G${row}*E${row}/F${row},H${row - 1})`,
  formula: `IF(B${row}=TRUE,H${row - 1}*F${row - 1}/F${row}+G${row}*E${row}*C${row}/F${row},H${row - 1})`,
  reason: '원본 누적효율 수식은 이번 구매의 가치에 반복구매 수량을 곱하지 않아 누적가격과 어긋납니다. 수량을 곱하도록 보정했습니다.'
});

const customLookup = (cell, source, formula) => ({
  sheet: '커스텀',
  cell,
  source,
  formula,
  reason: '원본 VLOOKUP 범위가 드롭다운 목록보다 한 행 아래에서 시작해, 목록의 첫 선택지를 고르면 #N/A가 됩니다. 조회 범위를 드롭다운 목록과 같게 맞췄습니다.'
});

export const corrections = [
  ...[7, 8, 9, 10].map(stepUpCumulative),
  {
    sheet: '스텝업',
    cell: 'J9',
    source: "('재화가치기준'!D7*2400+'재화가치기준'!D9*11)*'재화가치기준'!E6",
    formula: "('재화가치기준'!D6*2400+'재화가치기준'!D9*11)*'재화가치기준'!E6",
    reason: '원본은 STEP 11의 무료쥬얼 2,400개를 유료쥬얼 가치(D7)로 계산합니다. 같은 상품을 다루는 다른 다섯 탭처럼 무료쥬얼 가치(D6)를 쓰도록 보정했습니다.'
  },
  customLookup('D13', 'VLOOKUP(C13,B52:D57,3,FALSE)', 'VLOOKUP(C13,B51:D57,3,FALSE)'),
  customLookup('D14', 'VLOOKUP(C14,B61:D66,3,FALSE)', 'VLOOKUP(C14,B60:D66,3,FALSE)'),
  customLookup('D22', 'VLOOKUP(C22,B95:D98,3,FALSE)', 'VLOOKUP(C22,B94:D98,3,FALSE)')
];
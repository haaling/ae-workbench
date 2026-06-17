import {
  convertFreightToMovement,
  findFreightFallbackAmount,
  findFreightOrderNoInRow,
  normalizeOrderNo,
  normalizeMoney,
  normalizeCellValue,
  type RowData
} from '../calculatorDomain'

export function buildActualFreightRows(input: {
  freightRows: RowData[]
  freightOrderColumn: string
  effectiveFreightAmountCnyColumn: string
  effectiveFreightAmountUsdColumn: string
  orderIds: Set<string>
}): RowData[] {
  const {
    freightRows,
    freightOrderColumn,
    effectiveFreightAmountCnyColumn,
    effectiveFreightAmountUsdColumn,
    orderIds
  } = input

  return freightRows
    .map((row) => {
      const contextText = Object.values(row).map((value) => normalizeCellValue(value)).join(' ')
      const orderNo = findFreightOrderNoInRow(row, freightOrderColumn, orderIds)
      const cny = convertFreightToMovement(row[effectiveFreightAmountCnyColumn], contextText, effectiveFreightAmountCnyColumn)
      const usd = convertFreightToMovement(row[effectiveFreightAmountUsdColumn], contextText, effectiveFreightAmountUsdColumn)
      const fallback = Math.abs(cny) < 0.000001 && Math.abs(usd) < 0.000001
        ? convertFreightToMovement(findFreightFallbackAmount(row), contextText, 'fallback')
        : 0

      return {
        ...row,
        订单号: orderNo,
        计费金额合计CNY_标准化: normalizeMoney(cny || fallback),
        计费金额合计USD_标准化: normalizeMoney(usd)
      }
    })
    .filter((row) => {
      const id = normalizeOrderNo(row.订单号)
      return id && orderIds.has(id)
    })
}

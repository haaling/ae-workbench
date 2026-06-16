import { dedupeByKey, normalizeOrderNo, type RowData } from '../calculatorDomain'

export function buildOrderRows(ordersSheetRows: RowData[], ordersIdColumn: string): RowData[] {
  return dedupeByKey(ordersSheetRows, (row) => {
    const orderNo = normalizeOrderNo(row[ordersIdColumn])
    if (orderNo) {
      return `ORDER:${orderNo}`
    }
    return `RAW:${JSON.stringify(row)}`
  })
}

export function buildOrderIds(orderRows: RowData[], ordersIdColumn: string): Set<string> {
  return new Set(orderRows.map((row) => normalizeOrderNo(row[ordersIdColumn])).filter(Boolean))
}

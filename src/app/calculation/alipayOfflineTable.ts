import {
  normalizeCellValue,
  normalizeMoney,
  normalizeOrderNo,
  sortByOrderNo,
  type ExternalRecord,
  type RowData
} from '../calculatorDomain'
import type { OfflineFreightDiagnosticRow } from '../utils/appHelpers'

type OfflineFreightRecord = {
  订单号: string
  金额: number
  费用类别: string
  客户订单号: string
  备注: string
  来源文件: string
  来源Sheet: string
}

export function buildOfflineRowsByOrder(scopedOfflineRecords: OfflineFreightRecord[]): Map<string, OfflineFreightRecord[]> {
  const offlineRowsByOrder = new Map<string, OfflineFreightRecord[]>()
  scopedOfflineRecords.forEach((row) => {
    const current = offlineRowsByOrder.get(row.订单号) || []
    current.push(row)
    offlineRowsByOrder.set(row.订单号, current)
  })
  return offlineRowsByOrder
}

export function buildAlipayRowsByOrder(scopedAlipayRecords: ExternalRecord[]): Map<string, ExternalRecord[]> {
  const alipayRowsByOrder = new Map<string, ExternalRecord[]>()
  scopedAlipayRecords.forEach((row) => {
    const current = alipayRowsByOrder.get(row.订单号) || []
    current.push(row)
    alipayRowsByOrder.set(row.订单号, current)
  })
  return alipayRowsByOrder
}

export function buildUnmatchedAlipayRows(alipayRecords: ExternalRecord[], orderIds: Set<string>): RowData[] {
  return sortByOrderNo(
    alipayRecords
      .filter((row) => !row.订单号 || !orderIds.has(row.订单号))
      .map((row) => ({
        订单号: row.订单号,
        匹配状态: row.订单号 ? '支付宝有付款记录但订单表无对应订单' : '支付宝备注未识别订单号',
        店铺名: row.店铺名,
        订单号来源: row.订单号来源 || '未识别',
        支付宝交易单号: row.支付宝交易单号,
        支付宝交易单号和金额: row.支付宝交易单号 ? `${row.支付宝交易单号}/${Number(row.金额.toFixed(2))}` : `${Number(row.金额.toFixed(2))}`,
        交易对方: row.交易对方,
        商品说明: row.商品说明,
        金额: Number(row.金额.toFixed(2)),
        是否开发票: normalizeCellValue(row.是否开发票),
        '收/付款方式': row.收付款方式,
        交易订单号: row.交易订单号,
        商家订单号: row.商家订单号,
        备注: row.备注,
        来源文件: row.来源文件,
        来源Sheet: row.来源Sheet
      })),
    '订单号'
  )
}

export function buildAlipayMultiplicityRows(scopedAlipayRecords: ExternalRecord[]): RowData[] {
  const alipayRowsByOrder = buildAlipayRowsByOrder(scopedAlipayRecords)

  return [
    ...Array.from(alipayRowsByOrder.entries())
      .filter(([, rows]) => rows.length > 1)
      .map(([orderNo, rows]) => ({
        排查类型: '一订单对应多笔支付宝付款',
        订单号: orderNo,
        关联订单: orderNo,
        支付宝记录数: rows.length,
        店铺名: Array.from(new Set(rows.map((row) => normalizeCellValue(row.店铺名)).filter(Boolean))).join('、'),
        支付宝金额合计: normalizeMoney(rows.reduce((sum, row) => sum + row.金额, 0)),
        支付宝交易单号列表: rows.map((row) => normalizeCellValue(row.支付宝交易单号)).filter(Boolean).join('、'),
        支付宝备注列表: rows.map((row) => normalizeCellValue(row.备注)).filter(Boolean).join('；')
      })),
    ...Array.from(
      scopedAlipayRecords.reduce((map, row) => {
        const tradeKey =
          normalizeCellValue(row.支付宝交易单号) ||
          `${normalizeCellValue(row.来源文件)}|${normalizeCellValue(row.来源Sheet)}|${normalizeCellValue(row.备注)}|${row.金额}`
        const current = map.get(tradeKey) || []
        current.push(row)
        map.set(tradeKey, current)
        return map
      }, new Map<string, ExternalRecord[]>())
    )
      .filter(([, rows]) => new Set(rows.map((row) => row.订单号).filter(Boolean)).size > 1)
      .map(([tradeKey, rows]) => ({
        排查类型: '一笔支付宝付款对应多订单',
        订单号: '',
        关联订单: Array.from(new Set(rows.map((row) => row.订单号).filter(Boolean))).join('、'),
        支付宝记录数: rows.length,
        店铺名: Array.from(new Set(rows.map((row) => normalizeCellValue(row.店铺名)).filter(Boolean))).join('、'),
        支付宝金额合计: normalizeMoney(rows.reduce((sum, row) => sum + row.金额, 0)),
        支付宝交易单号列表: rows.map((row) => normalizeCellValue(row.支付宝交易单号)).filter(Boolean).join('、') || tradeKey,
        支付宝备注列表: rows.map((row) => normalizeCellValue(row.备注)).filter(Boolean).join('；')
      }))
  ].sort((left, right) =>
    normalizeCellValue(left.关联订单 || left.订单号).localeCompare(
      normalizeCellValue(right.关联订单 || right.订单号),
      'zh-Hans-CN',
      { numeric: true, sensitivity: 'base' }
    )
  )
}

export function buildOrdersWithoutAlipayRows(input: {
  orderRows: RowData[]
  ordersIdColumn: string
  effectiveOrdersStatusColumn: string
  effectiveOrdersTimeColumn: string
  alipayRowsByOrder: Map<string, ExternalRecord[]>
}): RowData[] {
  const { orderRows, ordersIdColumn, effectiveOrdersStatusColumn, effectiveOrdersTimeColumn, alipayRowsByOrder } = input

  return sortByOrderNo(
    orderRows
      .filter((row) => {
        const orderNo = normalizeOrderNo(row[ordersIdColumn])
        return orderNo && (alipayRowsByOrder.get(orderNo) || []).length === 0
      })
      .map((row) => ({
        订单号: normalizeOrderNo(row[ordersIdColumn]),
        订单状态: normalizeCellValue(row[effectiveOrdersStatusColumn]),
        订单时间: normalizeCellValue(row[effectiveOrdersTimeColumn]),
        支付宝付款记录数: 0,
        备注: '订单表有记录，但未匹配到支付宝付款记录'
      })),
    '订单号'
  )
}

export type { OfflineFreightRecord, OfflineFreightDiagnosticRow }

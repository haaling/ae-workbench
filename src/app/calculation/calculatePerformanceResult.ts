import type { ProcessResult, ResultSummary, RowData } from '../calculatorDomain'
import {
  DEFAULT_ORDER_ID_HINTS,
  DEFAULT_ALIPAY_AMOUNT_HINTS,
  DEFAULT_ALIPAY_REMARK_HINTS,
  buildCountMap,
  buildRefundRiskMap,
  buildRowsByOrderMap,
  convertFreightToMovement,
  isCompletedOrderStatus,
  isOrderIncomeSource,
  normalizeCellValue,
  normalizeMoney,
  normalizeOrderNo,
  prefixRow,
  sortByOrderNo,
  sortByRefundTypePriority,
  sortRefundTypesForDisplay,
  summarizeAmountByOrder,
  toFeeItemAmountColumn,
  toLogisticsSubItemAmountColumn,
  toNumericValue,
  toTypeAmountColumn,
  REFUND_BASE_AMOUNT_HINTS,
  hasKeyword
} from '../calculatorDomain'
import { buildExternalRecords, buildOfflineFreightRecordsWithDiagnostics } from '../utils/appHelpers'
import { buildActualFreightRows } from './freightTable'
import {
  buildAlipayMultiplicityRows,
  buildAlipayRowsByOrder,
  buildOfflineRowsByOrder,
  buildOrdersWithoutAlipayRows,
  buildUnmatchedAlipayRows,
  type OfflineFreightRecord
} from './alipayOfflineTable'
import { buildIncomeRefundEntries } from './incomeRefundTable'
import { buildOrderIds, buildOrderRows } from './ordersTable'
import type { IncomeUploadItem, MultiUploadItem } from './types'

export type CalculatePerformanceInput = {
  ordersSheetRows: RowData[]
  freightSheetRows: RowData[]
  incomeFiles: IncomeUploadItem[]
  refundFiles: IncomeUploadItem[]
  alipayFiles: MultiUploadItem[]
  offlineFiles: MultiUploadItem[]
  ordersIdColumn: string
  effectiveOrdersStatusColumn: string
  effectiveOrdersTimeColumn: string
  effectiveIncomeDetailAmountColumns: string[]
  effectiveRefundDetailAmountColumns: string[]
  effectiveRefundTypeColumn: string
  effectiveRefundProductNameColumn: string
  effectiveRefundSkuIdColumn: string
  freightOrderColumn: string
  effectiveFreightAmountCnyColumn: string
  effectiveFreightAmountUsdColumn: string
  effectiveFreightFulfillmentColumn: string
  effectiveFreightWaybillColumn: string
  shopName: string
  usdExchangeRate: string
}

export function calculatePerformanceResult(input: CalculatePerformanceInput): ProcessResult {
  const {
    ordersSheetRows,
    freightSheetRows,
    incomeFiles,
    refundFiles,
    alipayFiles,
    offlineFiles,
    ordersIdColumn,
    effectiveOrdersStatusColumn,
    effectiveOrdersTimeColumn,
    effectiveIncomeDetailAmountColumns,
    effectiveRefundDetailAmountColumns,
    effectiveRefundTypeColumn,
    effectiveRefundProductNameColumn,
    effectiveRefundSkuIdColumn,
    freightOrderColumn,
    effectiveFreightAmountCnyColumn,
    effectiveFreightAmountUsdColumn,
    effectiveFreightFulfillmentColumn,
    effectiveFreightWaybillColumn,
    shopName,
    usdExchangeRate
  } = input

  const DEBUG_CALC = true
  const isLogisticsExpenseType = (typeText: string): boolean => {
    const text = normalizeCellValue(typeText)
    if (!text) {
      return false
    }

    if (text.includes('支出>物流费用')) {
      return true
    }

    return /支出\s*(?:-|=)?\s*>\s*物流(?:费|费用)/.test(text)
  }

  const isExcludedFromLogisticsDeduction = (feeItemText: string): boolean => {
    const text = normalizeCellValue(feeItemText).toLowerCase()
    if (!text) {
      return false
    }

    return (
      text.includes('关税') ||
      text.includes('税费') ||
      text.includes('ddp') ||
      text.includes('tax') ||
      text.includes('tariff') ||
      text.includes('duty')
    )
  }

  const EXPECTED_INCOME_EXCLUDED_FEE_ITEMS = new Set([
    '物流上网超时处罚_费用项金额',
    '违背发货承诺处罚_费用项金额',
    '物流上网超时处罚',
    '违背发货承诺处罚'
  ])
  const orderRows = buildOrderRows(ordersSheetRows, ordersIdColumn)
  const orderIds = buildOrderIds(orderRows, ordersIdColumn)

  const alipayRecords = buildExternalRecords({
    files: alipayFiles,
    remarkHints: DEFAULT_ALIPAY_REMARK_HINTS,
    amountHints: DEFAULT_ALIPAY_AMOUNT_HINTS,
    mode: 'alipay',
    shopName,
    useParsedRemarkOrderNoOnly: true
  })
  const offlineFreightBuildResult = buildOfflineFreightRecordsWithDiagnostics({ files: offlineFiles })
  const offlineFreightRecords = offlineFreightBuildResult.records
  const offlineFreightDiagnosticRows = offlineFreightBuildResult.diagnostics

  const { incomeDetailRows, refundDetailRows, incomeEntries } = buildIncomeRefundEntries({
    incomeFiles,
    refundFiles,
    effectiveIncomeDetailAmountColumns,
    effectiveRefundDetailAmountColumns,
    effectiveRefundTypeColumn
  })

  const resolveRawOrderColumn = (headers: string[], preferred: string): string => {
    if (preferred && headers.includes(preferred)) {
      return preferred
    }

    return headers.find((header) => hasKeyword(header, DEFAULT_ORDER_ID_HINTS)) || preferred
  }

  const russiaOrderDetailOrderSet = new Set(
    incomeFiles.flatMap((item) => {
      const candidateSheets = item.file.sheets.filter(
        (sheet) => !sheet.headers.some((header) => normalizeCellValue(header) === '收支类型')
      )

      return candidateSheets.flatMap((sheet) => {
        const orderColumn = resolveRawOrderColumn(sheet.headers, item.orderColumn)
        return sheet.rows
          .map((row) => normalizeOrderNo(row[orderColumn]))
          .filter(Boolean)
      })
    })
  )

  if (DEBUG_CALC) {
    console.groupCollapsed('[Calc Debug] 上传识别与明细入库')
    console.log('incomeFiles:', incomeFiles.map((item) => ({
      fileName: item.file.fileName,
      selectedSheet: item.file.selectedSheetName,
      orderColumn: item.orderColumn,
      flowTypeColumn: item.flowTypeColumn || ''
    })))
    console.log('effectiveIncomeDetailAmountColumns:', effectiveIncomeDetailAmountColumns)
    console.log('effectiveRefundDetailAmountColumns:', effectiveRefundDetailAmountColumns)
    console.log('incomeDetailRows count:', incomeDetailRows.length)
    console.log('refundDetailRows count:', refundDetailRows.length)
    console.log('russiaOrderDetailOrderSet count:', russiaOrderDetailOrderSet.size)
    console.table(incomeDetailRows.slice(0, 20).map((row) => ({
      订单号: normalizeOrderNo(row.订单号),
      收支类型: normalizeCellValue(row.收支类型),
      变动金额: normalizeMoney(row.变动金额),
      来源: normalizeCellValue(row.来源),
      收支来源文件: normalizeCellValue(row.收支来源文件),
      收支来源Sheet: normalizeCellValue(row.收支来源Sheet)
    })))
    console.groupCollapsed('[Calc Debug] 线下运费解析诊断')
    console.log('offline rows total:', offlineFreightDiagnosticRows.length)
    console.log('offline rows included:', offlineFreightDiagnosticRows.filter((row) => row.状态 === '纳入计算').length)
    console.log('offline rows skipped:', offlineFreightDiagnosticRows.filter((row) => row.状态 === '已跳过').length)
    console.table(offlineFreightDiagnosticRows.slice(0, 40))
    console.groupEnd()
    console.groupEnd()
  }

  const isInvoiceMarkedYes = (value: unknown): boolean => {
    const text = normalizeCellValue(value).toLowerCase()
    return text === '是' || text === 'y' || text === 'yes' || text === 'true' || text.includes('是')
  }

  const scopedAlipayRecords = alipayRecords.filter((row) => orderIds.has(row.订单号))
  const scopedOfflineRecords = offlineFreightRecords.filter((row) => orderIds.has(row.订单号))
  const invoicedAlipayOrderSet = new Set(
    scopedAlipayRecords.filter((row) => isInvoiceMarkedYes(row.是否开发票)).map((row) => normalizeOrderNo(row.订单号))
  )

  const unmatchedAlipayRows = buildUnmatchedAlipayRows(alipayRecords, orderIds)

  const scopedIncomeDetailRows = incomeDetailRows.filter((row) => orderIds.has(normalizeOrderNo(row.订单号)))
  const scopedRefundDetailRows = refundDetailRows.filter((row) => orderIds.has(normalizeOrderNo(row.订单号)))
  const refundRawRows = refundFiles.flatMap((item) => {
    const selectedSheet = item.file.sheets.find((sheet) => sheet.name === item.file.selectedSheetName)
    if (!selectedSheet) {
      return [] as Array<{ 订单号: string; 商品名称: string; skuId: string }>
    }

    return selectedSheet.rows
      .map((row) => ({
        订单号: normalizeOrderNo(row[item.orderColumn]),
        商品名称: normalizeCellValue(row[effectiveRefundProductNameColumn]),
        skuId: normalizeCellValue(row[effectiveRefundSkuIdColumn])
      }))
      .filter((row) => row.订单号 && orderIds.has(row.订单号))
  })
  const refundRiskMap = buildRefundRiskMap(refundRawRows)

  const incomeOrderSet = new Set(scopedIncomeDetailRows.map((row) => normalizeOrderNo(row.订单号)).filter(Boolean))
  const refundOrderSet = new Set(scopedRefundDetailRows.map((row) => normalizeOrderNo(row.订单号)).filter(Boolean))
  const incomeSummaryMap = summarizeAmountByOrder(scopedIncomeDetailRows)
  const refundSummaryMap = summarizeAmountByOrder(scopedRefundDetailRows)

  const incomeOnlyOrdersRows = sortByOrderNo(
    Array.from(incomeOrderSet)
      .filter((orderNo) => !refundOrderSet.has(orderNo))
      .map((orderNo) => {
        const incomeSummary = incomeSummaryMap.get(orderNo)
        return {
          订单号: orderNo,
          差异方向: '仅订单明细有',
          订单明细条数: incomeSummary?.count || 0,
          订单明细金额合计: incomeSummary?.amount || 0,
          放退款明细条数: 0,
          放退款金额合计: 0,
          放退款类型: ''
        }
      }),
    '订单号'
  )

  const refundOnlyOrdersRows = sortByOrderNo(
    Array.from(refundOrderSet)
      .filter((orderNo) => !incomeOrderSet.has(orderNo))
      .map((orderNo) => {
        const refundSummary = refundSummaryMap.get(orderNo)
        return {
          订单号: orderNo,
          差异方向: '仅放退款明细有',
          订单明细条数: 0,
          订单明细金额合计: 0,
          放退款明细条数: refundSummary?.count || 0,
          放退款金额合计: refundSummary?.amount || 0,
          放退款类型: refundSummary ? sortRefundTypesForDisplay(Array.from(refundSummary.refundTypes)).join('、') : ''
        }
      }),
    '订单号'
  )

  const actualIncomeEntries = incomeEntries.filter((entry) => orderIds.has(normalizeOrderNo(entry.订单号)))
  const actualIncomeRows: RowData[] = []
  const incomeRowsByOrder = new Map<string, RowData[]>()

  actualIncomeEntries.forEach((entry) => {
    const orderNo = normalizeOrderNo(entry.订单号)
    const detailRow: RowData = {
      ...entry,
      订单号: orderNo,
      变动金额: normalizeMoney(entry.变动金额)
    }

    actualIncomeRows.push(detailRow)
    const current = incomeRowsByOrder.get(orderNo) || []
    current.push(detailRow)
    incomeRowsByOrder.set(orderNo, current)
  })

  const actualFreightRows = buildActualFreightRows({
    freightRows: freightSheetRows,
    freightOrderColumn,
    effectiveFreightAmountCnyColumn,
    effectiveFreightAmountUsdColumn,
    orderIds
  })

  const offlineRowsByOrder = buildOfflineRowsByOrder(scopedOfflineRecords as OfflineFreightRecord[])
  const alipayRowsByOrder = buildAlipayRowsByOrder(scopedAlipayRecords)

  const ordersWithoutAlipayRows = buildOrdersWithoutAlipayRows({
    orderRows,
    ordersIdColumn,
    effectiveOrdersStatusColumn,
    effectiveOrdersTimeColumn,
    alipayRowsByOrder
  })

  const alipayMultiplicityRows = buildAlipayMultiplicityRows(scopedAlipayRecords)

  const incomeCountMap = buildCountMap(actualIncomeRows, '订单号')
  const freightCountMap = buildCountMap(actualFreightRows, '订单号')
  const freightRowsByOrder = buildRowsByOrderMap(actualFreightRows, '订单号')

  const markedOrderRows: RowData[] = orderRows.map((row): RowData => {
    const orderNo = normalizeOrderNo(row[ordersIdColumn])
    const incomeCount = incomeCountMap.get(orderNo) || 0
    const freightCount = freightCountMap.get(orderNo) || 0
    const offlineCount = (offlineRowsByOrder.get(orderNo) || []).length
    const alipayCount = (alipayRowsByOrder.get(orderNo) || []).length

    return {
      ...row,
      收支记录状态: incomeCount > 0 ? '有记录' : '无记录',
      收支记录条数: incomeCount,
      运费记录条数: freightCount,
      线下发货记录条数: offlineCount,
      支付宝采购记录条数: alipayCount
    }
  })

  const integratedSummaryRows = sortByOrderNo(
    markedOrderRows.map((row): RowData => {
      const orderNo = normalizeOrderNo(row[ordersIdColumn])
      return {
        ...row,
        订单收支明细JSON: JSON.stringify(incomeRowsByOrder.get(orderNo) || []),
        运费明细JSON: JSON.stringify(freightRowsByOrder.get(orderNo) || [])
      }
    }),
    ordersIdColumn
  )

  const integratedDetailRows: RowData[] = []
  const performanceRows: RowData[] = []
  markedOrderRows.forEach((orderRow) => {
    const orderNo = normalizeOrderNo(orderRow[ordersIdColumn])
    const orderStatus = normalizeCellValue(orderRow[effectiveOrdersStatusColumn])
    const orderTime = normalizeCellValue(orderRow[effectiveOrdersTimeColumn])
    const incomeList = incomeRowsByOrder.get(orderNo) || []
    const freightList = freightRowsByOrder.get(orderNo) || []
    const offlineList = offlineRowsByOrder.get(orderNo) || []
    const alipayList = alipayRowsByOrder.get(orderNo) || []

    incomeList.forEach((incomeRow) => {
      const refundRiskText = refundRiskMap.get(orderNo) || ''
      performanceRows.push({
        订单号: orderNo,
        订单状态: orderStatus,
        订单时间: orderTime,
        收支类型: normalizeCellValue(incomeRow.收支类型),
        变动金额: normalizeMoney(incomeRow.变动金额),
        费用项: normalizeCellValue(incomeRow.费用项),
        物流履约单号: '',
        运单号: '',
        币种: 'CNY',
        来源: normalizeCellValue(incomeRow.来源),
        放退款类型: normalizeCellValue(incomeRow.放退款类型),
        放退款核查标记: refundRiskText,
        放退款行佣金合计: normalizeMoney(incomeRow.放退款行佣金合计),
        放退款税费候选金额: normalizeMoney(incomeRow.放退款税费候选金额),
        收支来源文件: normalizeCellValue(incomeRow.收支来源文件),
        收支来源Sheet: normalizeCellValue(incomeRow.收支来源Sheet),
        计费金额合计CNY: '',
        计费金额合计USD: ''
      })
    })

    freightList.forEach((freightRow) => {
      const movementCny = normalizeMoney(freightRow.计费金额合计CNY_标准化)
      const movementUsd = normalizeMoney(freightRow.计费金额合计USD_标准化)
      const hasCny = Math.abs(movementCny) > 0
      const hasUsd = Math.abs(movementUsd) > 0
      const selectedAmount = hasCny ? movementCny : movementUsd
      const billingType = selectedAmount >= 0 ? '物流赔付' : '物流费用'

      performanceRows.push({
        订单号: orderNo,
        订单状态: orderStatus,
        订单时间: orderTime,
        收支类型: '物流运费',
        变动金额: selectedAmount,
        费用项: billingType,
        物流履约单号: normalizeCellValue(freightRow[effectiveFreightFulfillmentColumn]),
        运单号: normalizeCellValue(freightRow[effectiveFreightWaybillColumn]),
        币种: hasCny ? 'CNY' : hasUsd ? 'USD' : '',
        来源: '运费表',
        放退款核查标记: '',
        计费金额合计CNY: movementCny,
        计费金额合计USD: movementUsd
      })
    })

    offlineList.forEach((offlineRow) => {
      const movement = convertFreightToMovement(offlineRow.金额, offlineRow.费用类别 || offlineRow.备注)
      performanceRows.push({
        订单号: orderNo,
        订单状态: orderStatus,
        订单时间: orderTime,
        收支类型: '线下运费',
        变动金额: movement,
        费用项: normalizeCellValue(offlineRow.费用类别) || '线下物流费用',
        物流履约单号: '',
        运单号: '',
        币种: 'CNY',
        来源: '线下发货记录',
        放退款核查标记: '',
        客户订单号: normalizeCellValue(offlineRow.客户订单号),
        收支来源文件: offlineRow.来源文件,
        收支来源Sheet: offlineRow.来源Sheet,
        计费金额合计CNY: movement,
        计费金额合计USD: ''
      })
    })

    alipayList.forEach((alipayRow) => {
      const movement = Number((-Math.abs(alipayRow.金额)).toFixed(2))
      performanceRows.push({
        订单号: orderNo,
        订单状态: orderStatus,
        订单时间: orderTime,
        收支类型: '采购支出',
        变动金额: movement,
        费用项: '采购付款',
        物流履约单号: '',
        运单号: '',
        币种: 'CNY',
        来源: '支付宝订单记录',
        放退款核查标记: '',
        是否开发票: normalizeCellValue(alipayRow.是否开发票),
        收支来源文件: alipayRow.来源文件,
        收支来源Sheet: alipayRow.来源Sheet,
        支付宝交易单号: alipayRow.支付宝交易单号,
        支付宝交易单号和金额: alipayRow.支付宝交易单号 ? `${alipayRow.支付宝交易单号}/${alipayRow.金额}` : `${alipayRow.金额}`,
        交易对方: alipayRow.交易对方,
        商品说明: alipayRow.商品说明,
        金额: alipayRow.金额,
        店铺名: alipayRow.店铺名,
        '收/付款方式': alipayRow.收付款方式,
        交易订单号: alipayRow.交易订单号,
        商家订单号: alipayRow.商家订单号,
        收支备注: alipayRow.备注,
        计费金额合计CNY: movement,
        计费金额合计USD: ''
      })
    })

    if (incomeList.length === 0 && freightList.length === 0 && offlineList.length === 0 && alipayList.length === 0) {
      integratedDetailRows.push({
        订单号: orderNo,
        数据来源: '仅订单表',
        ...prefixRow(orderRow, '订单表_')
      })
      return
    }

    incomeList.forEach((incomeRow) => {
      integratedDetailRows.push({
        订单号: orderNo,
        数据来源: normalizeCellValue(incomeRow.来源),
        ...prefixRow(orderRow, '订单表_'),
        ...prefixRow(incomeRow, '明细表_')
      })
    })

    freightList.forEach((freightRow) => {
      integratedDetailRows.push({
        订单号: orderNo,
        数据来源: '运费表',
        ...prefixRow(orderRow, '订单表_'),
        ...prefixRow(freightRow, '运费表_')
      })
    })

    offlineList.forEach((offlineRow) => {
      integratedDetailRows.push({
        订单号: orderNo,
        数据来源: '线下发货记录',
        ...prefixRow(orderRow, '订单表_'),
        ...prefixRow(offlineRow, '线下发货_')
      })
    })

    alipayList.forEach((alipayRow) => {
      integratedDetailRows.push({
        订单号: orderNo,
        数据来源: '支付宝订单记录',
        ...prefixRow(orderRow, '订单表_'),
        ...prefixRow(alipayRow, '支付宝_')
      })
    })
  })

  const sortedMarkedOrderRows = sortByOrderNo(markedOrderRows, ordersIdColumn)
  const ordersWithoutIncomeRows = sortedMarkedOrderRows.filter(
    (row) => normalizeCellValue(row.收支记录状态) === '无记录'
  )
  const sortedActualIncomeRows = sortByOrderNo(actualIncomeRows, '订单号')
  const sortedActualFreightRows = sortByOrderNo(actualFreightRows, '订单号')
  const sortedIntegratedDetailRows = sortByOrderNo(integratedDetailRows, '订单号')
  const sortedPerformanceRows = sortByOrderNo(performanceRows, '订单号')
  const uniqueOrderIds = Array.from(new Set(orderRows.map((row) => normalizeCellValue(row[ordersIdColumn])).filter(Boolean)))

  const allTypeColumns = Array.from(
    new Set(
      sortedPerformanceRows
        .filter((row) => isOrderIncomeSource(normalizeCellValue(row.来源)))
        .map((row) => toTypeAmountColumn(normalizeCellValue(row.收支类型)))
        .filter(Boolean)
    )
  )
  const allFeeItemColumns = Array.from(
    new Set(
      effectiveIncomeDetailAmountColumns
        .map((column) => toFeeItemAmountColumn(column))
        .concat(
          sortedPerformanceRows
            .filter((row) => isOrderIncomeSource(normalizeCellValue(row.来源)))
            .map((row) => toFeeItemAmountColumn(normalizeCellValue(row.费用项) || normalizeCellValue(row.收支类型)))
        )
        .filter(Boolean)
    )
  )
  const allLogisticsSubItemColumns = Array.from(
    new Set(
      sortedPerformanceRows.flatMap((row) => {
        const source = normalizeCellValue(row.来源)
        const type = normalizeCellValue(row.收支类型)
        const feeItem = normalizeCellValue(row.费用项) || type || '未分类'

        if (
          isOrderIncomeSource(source) &&
          isLogisticsExpenseType(type) &&
          !isExcludedFromLogisticsDeduction(feeItem)
        ) {
          return [toLogisticsSubItemAmountColumn('收支表', feeItem)]
        }
        if (source === '运费表' || type === '物流运费') {
          return [toLogisticsSubItemAmountColumn('金掌柜', feeItem)]
        }
        return []
      })
    )
  )
  const allOfflineCategoryColumns = Array.from(
    new Set(
      sortedPerformanceRows
        .filter((row) => normalizeCellValue(row.来源) === '线下发货记录')
        .map((row) => `线下运费_${normalizeCellValue(row.费用项) || '未分类线下费用'}`)
        .filter(Boolean)
    )
  )

  const aggregatedMap = new Map<string, RowData>()
  sortedPerformanceRows.forEach((row) => {
    const orderNo = normalizeOrderNo(row.订单号)
    if (!orderNo) {
      return
    }

    const amount = normalizeMoney(row.变动金额)
    const source = normalizeCellValue(row.来源)
    const type = normalizeCellValue(row.收支类型)
    const feeItem = normalizeCellValue(row.费用项) || type
    const typeColumn = toTypeAmountColumn(type)
    const feeItemColumn = toFeeItemAmountColumn(feeItem)
    const current = aggregatedMap.get(orderNo) || {
      订单号: orderNo,
      订单状态: normalizeCellValue(row.订单状态),
      是否放款: '否',
      订单时间: normalizeCellValue(row.订单时间),
      放退款类型: '',
      放退款核查标记: '',
      待结算金额合计: 0,
      备注: '',
      收入_收支明细表: 0,
      支出_收支明细表: 0,
      收支表_支出物流费用: 0,
      金掌柜物流费支出: 0,
      物流支出总和: 0,
      物流支出_收支表: 0,
      物流支出_金掌柜: 0,
      收支总和_不含物流费用: 0,
      差异_收支不含物流减退放款: 0,
      订单明细_净收支合计: 0,
      订单明细_放款金额合计: 0,
      订单明细_售中退款金额合计: 0,
      订单明细_售后退款金额合计: 0,
      订单明细_平台分账金额合计: 0,
      订单明细_平台分账退回金额合计: 0,
      订单明细_净放款基准金额: 0,
      订单明细_净放款口径收入: 0,
      金掌柜计费金额合计USD: 0,
      税费金额_金掌柜: 0,
      税费金额_收支表: 0,
      税费金额_放退款: 0,
      税费核对_USD汇率: 0,
      税费核对_按金掌柜USD折CNY: 0,
      税费核对_按收支表税费: 0,
      税费核对_基准来源: '',
      税费核对_按放退款税费候选金额: 0,
      税费核对_差异: 0,
      税费核对状态: '',
      收入_按订单明细: 0,
      预计可得_按收支及运费: 0,
      预计可得_按退放款及运费: 0,
      预计可得差异_收支减退放款: 0,
      预计可得校验状态: '一致',
      放退款_金额项合计: 0,
      放退款_其他费用合计: 0,
      放退款_佣金合计: 0,
      放退款_税费候选金额合计: 0,
      收入_按放退款明细: 0,
      收入差异_订单明细减放退款: 0,
      收入校验状态: '一致',
      线上运费: 0,
      线下运费: 0,
      采购费用: 0,
      __incomeStatementRowCount: 0,
      __incomeStatementNetAmount: 0,
      __incomeStatementLogisticsExpense: 0,
      __incomeStatementTaxFeeExcludedAmount: 0,
      __incomeRowCount: 0,
      __refundRowCount: 0,
      __expectedIncomeExcludedFeeAmount: 0,
      支付宝是否开发票: '',
      最终收入_未扣运费: 0,
      最终收入_扣运费: 0,
      收入合计: 0,
      支出合计: 0,
      总收支: 0
    }

    if (!normalizeCellValue(current.订单状态) && normalizeCellValue(row.订单状态)) {
      current.订单状态 = normalizeCellValue(row.订单状态)
    }
    if (!normalizeCellValue(current.订单时间) && normalizeCellValue(row.订单时间)) {
      current.订单时间 = normalizeCellValue(row.订单时间)
    }

    const riskText = refundRiskMap.get(orderNo) || ''
    if (riskText) {
      current.放退款核查标记 = riskText
    }

    const refundTypeText = normalizeCellValue(row.放退款类型)
    if (refundTypeText) {
      const existingTypes = normalizeCellValue(current.放退款类型)
        .split('、')
        .map((item) => item.trim())
        .filter(Boolean)
      if (!existingTypes.includes(refundTypeText)) {
        current.放退款类型 = sortRefundTypesForDisplay([...existingTypes, refundTypeText]).join('、')
      }
    }

    if (isOrderIncomeSource(source)) {
      current.__incomeRowCount = toNumericValue(current.__incomeRowCount) + 1
      if (source === '订单收支明细表') {
        current.__incomeStatementRowCount = toNumericValue(current.__incomeStatementRowCount) + 1
        current.__incomeStatementNetAmount = normalizeMoney(
          toNumericValue(current.__incomeStatementNetAmount) + amount
        )
      }

      current.订单明细_净收支合计 = normalizeMoney(toNumericValue(current.订单明细_净收支合计) + amount)

      if (amount >= 0) {
        current.收入_收支明细表 = normalizeMoney(toNumericValue(current.收入_收支明细表) + amount)
      } else {
        current.支出_收支明细表 = normalizeMoney(toNumericValue(current.支出_收支明细表) + Math.abs(amount))
      }

      if (isLogisticsExpenseType(type) && !isExcludedFromLogisticsDeduction(feeItem)) {
        const logisticsExpense = amount < 0 ? Math.abs(amount) : 0
        current.收支表_支出物流费用 = normalizeMoney(
          toNumericValue(current.收支表_支出物流费用) + logisticsExpense
        )

        if (source === '订单收支明细表') {
          current.__incomeStatementLogisticsExpense = normalizeMoney(
            toNumericValue(current.__incomeStatementLogisticsExpense) + logisticsExpense
          )
        }

        const logisticsSubItemColumn = toLogisticsSubItemAmountColumn('收支表', feeItem)
        current[logisticsSubItemColumn] = normalizeMoney(
          toNumericValue(current[logisticsSubItemColumn]) + logisticsExpense
        )
      }

      if (isLogisticsExpenseType(type) && isExcludedFromLogisticsDeduction(feeItem)) {
        const excludedTaxFeeExpense = amount < 0 ? Math.abs(amount) : 0
        current.__incomeStatementTaxFeeExcludedAmount = normalizeMoney(
          toNumericValue(current.__incomeStatementTaxFeeExcludedAmount) + excludedTaxFeeExpense
        )
      }

      if (type.includes('待结算金额')) {
        current.待结算金额合计 = normalizeMoney(toNumericValue(current.待结算金额合计) + Math.abs(amount))
      }
      if (type.includes('放款金额')) {
        current.订单明细_放款金额合计 = normalizeMoney(toNumericValue(current.订单明细_放款金额合计) + amount)
      }
      if (type.includes('售中退款金额')) {
        current.订单明细_售中退款金额合计 = normalizeMoney(
          toNumericValue(current.订单明细_售中退款金额合计) + Math.abs(amount)
        )
      }
      if (type.includes('售后退款金额')) {
        current.订单明细_售后退款金额合计 = normalizeMoney(
          toNumericValue(current.订单明细_售后退款金额合计) + Math.abs(amount)
        )
      }
      const isPlatformSplitReturnType =
        type.includes('平台分账退回金额') || type.includes('平台分账金额退回')
      if (isPlatformSplitReturnType) {
        current.订单明细_平台分账退回金额合计 = normalizeMoney(
          toNumericValue(current.订单明细_平台分账退回金额合计) + Math.abs(amount)
        )
      } else if (type.includes('平台分账金额')) {
        current.订单明细_平台分账金额合计 = normalizeMoney(
          toNumericValue(current.订单明细_平台分账金额合计) + Math.abs(amount)
        )
      }

      const isExpectedExcludedFeeItem =
        EXPECTED_INCOME_EXCLUDED_FEE_ITEMS.has(feeItem) ||
        EXPECTED_INCOME_EXCLUDED_FEE_ITEMS.has(type)
      if (isExpectedExcludedFeeItem) {
        current.__expectedIncomeExcludedFeeAmount = normalizeMoney(
          toNumericValue(current.__expectedIncomeExcludedFeeAmount) + amount
        )
      }
    }

    if (source === '放退款订单明细') {
      current.__refundRowCount = toNumericValue(current.__refundRowCount) + 1

      if (hasKeyword(type, REFUND_BASE_AMOUNT_HINTS)) {
        current.放退款_金额项合计 = normalizeMoney(toNumericValue(current.放退款_金额项合计) + amount)
        const tariffCandidate = normalizeMoney(row.放退款税费候选金额)
        if (tariffCandidate > 0.000001) {
          current.放退款_税费候选金额合计 = normalizeMoney(
            toNumericValue(current.放退款_税费候选金额合计) + tariffCandidate
          )
        }
      } else {
        current.放退款_其他费用合计 = normalizeMoney(
          toNumericValue(current.放退款_其他费用合计) + Math.abs(amount)
        )
        if (type.includes('佣金')) {
          current.放退款_佣金合计 = normalizeMoney(
            toNumericValue(current.放退款_佣金合计) + Math.abs(amount)
          )
        }
      }
    }

    if (source === '运费表' || type === '物流运费') {
      const freightUsdAmount = Math.abs(normalizeMoney(row.计费金额合计USD))
      current.金掌柜计费金额合计USD = normalizeMoney(
        toNumericValue(current.金掌柜计费金额合计USD) + freightUsdAmount
      )

      if (amount > 0) {
        current.金掌柜物流费支出 = normalizeMoney(toNumericValue(current.金掌柜物流费支出) + amount)
        current.物流支出_金掌柜 = normalizeMoney(toNumericValue(current.物流支出_金掌柜) + amount)

        const logisticsSubItemColumn = toLogisticsSubItemAmountColumn('金掌柜', feeItem)
        current[logisticsSubItemColumn] = normalizeMoney(
          toNumericValue(current[logisticsSubItemColumn]) + amount
        )
      }
    }

    if (source === '线下发货记录' || type === '线下运费') {
      current.线下运费 = normalizeMoney(toNumericValue(current.线下运费) + amount)
      const offlineCategoryColumn = `线下运费_${feeItem || '未分类线下费用'}`
      current[offlineCategoryColumn] = normalizeMoney(toNumericValue(current[offlineCategoryColumn]) + amount)
    }

    if (source === '支付宝订单记录' || type === '采购支出') {
      current.采购费用 = normalizeMoney(toNumericValue(current.采购费用) + Math.abs(amount))
      const invoiceFlag = normalizeCellValue(row.是否开发票)
      if (isInvoiceMarkedYes(invoiceFlag)) {
        current.支付宝是否开发票 = '是'
      } else if (!normalizeCellValue(current.支付宝是否开发票) && invoiceFlag) {
        current.支付宝是否开发票 = invoiceFlag
      }
    }

    current[typeColumn] = normalizeMoney(toNumericValue(current[typeColumn]) + amount)
    if (isOrderIncomeSource(source)) {
      current[feeItemColumn] = normalizeMoney(toNumericValue(current[feeItemColumn]) + amount)
    }

    const payoutAmount = normalizeMoney(current.订单明细_放款金额合计)
    const pendingSettlementAmount = normalizeMoney(current.待结算金额合计)
    const inSaleRefundAmount = normalizeMoney(current.订单明细_售中退款金额合计)
    const afterSaleRefundAmount = normalizeMoney(current.订单明细_售后退款金额合计)
    const platformSplitAmount = normalizeMoney(current.订单明细_平台分账金额合计)
    const platformSplitReturnAmount = normalizeMoney(current.订单明细_平台分账退回金额合计)
    const isPendingSettlementOrder =
      pendingSettlementAmount > 0.000001 && !isCompletedOrderStatus(normalizeCellValue(current.订单状态))
    const netPayoutBaseAmount = isPendingSettlementOrder ? pendingSettlementAmount : payoutAmount
    const netPayoutIncome = normalizeMoney(
      netPayoutBaseAmount -
        inSaleRefundAmount -
        afterSaleRefundAmount -
        platformSplitAmount +
        platformSplitReturnAmount
    )
    const incomeLogisticsExpense = normalizeMoney(current.收支表_支出物流费用)
    const hasNetPayoutInputs =
      Math.abs(netPayoutBaseAmount) > 0.000001 ||
      Math.abs(inSaleRefundAmount) > 0.000001 ||
      Math.abs(afterSaleRefundAmount) > 0.000001 ||
      Math.abs(platformSplitAmount) > 0.000001 ||
      Math.abs(platformSplitReturnAmount) > 0.000001
    const incomeFromOrderDetail = hasNetPayoutInputs
      ? netPayoutIncome
      : normalizeMoney(current.订单明细_净收支合计)
    const incomeFromIncomeStatement = normalizeMoney(
      toNumericValue(current.订单明细_净收支合计) +
      incomeLogisticsExpense +
      toNumericValue(current.__incomeStatementTaxFeeExcludedAmount)
    )
    current.订单明细_净放款基准金额 = netPayoutBaseAmount
    current.订单明细_净放款口径收入 = netPayoutIncome
    const incomeFromRefundDetail = normalizeMoney(
      toNumericValue(current.放退款_金额项合计) - toNumericValue(current.放退款_其他费用合计)
    )
    const incomeFromOrderDetailExcludingLogistics = incomeFromIncomeStatement
    const incomeDiff = normalizeMoney(incomeFromOrderDetailExcludingLogistics - incomeFromRefundDetail)
    const purchaseExpense = normalizeMoney(current.采购费用)
    const jzgFreightExpense = normalizeMoney(current.金掌柜物流费支出)
    const offlineFreightExpense = normalizeMoney(current.线下运费)
    const logisticsExpenseTotal = normalizeMoney(incomeLogisticsExpense + jzgFreightExpense + offlineFreightExpense)
    current.物流支出总和 = logisticsExpenseTotal
    const totalFreight = logisticsExpenseTotal
    const hasIncomeStatementRows = toNumericValue(current.__incomeStatementRowCount) > 0
    const hasRussiaOrderDetailRows =
      russiaOrderDetailOrderSet.has(orderNo) ||
      toNumericValue(current.__incomeRowCount) > toNumericValue(current.__incomeStatementRowCount)
    const isNonRussiaIncomeStatementOnlyOrder = hasIncomeStatementRows && !hasRussiaOrderDetailRows
    const hasRefundRows = toNumericValue(current.__refundRowCount) > 0
    const incomeFromIncomeStatementSupplement = normalizeMoney(
      toNumericValue(current.__incomeStatementNetAmount) +
      toNumericValue(current.__incomeStatementLogisticsExpense)
    )
    const hasIncomeStatementSupplement = Math.abs(incomeFromIncomeStatementSupplement) > 0.000001
    current.是否放款 = hasRefundRows ? '是' : '否'
    const incomeBasis =
      isNonRussiaIncomeStatementOnlyOrder
        ? 'income'
        : hasRussiaOrderDetailRows && hasRefundRows && hasIncomeStatementSupplement
          ? 'refund-plus-income-statement'
        : hasRefundRows
          ? 'refund'
          : 'order'

    const primaryIncomeBeforeFreight = incomeBasis === 'income'
      ? incomeFromIncomeStatement
      : incomeBasis === 'refund-plus-income-statement'
        ? normalizeMoney(incomeFromRefundDetail + incomeFromIncomeStatementSupplement)
      : incomeBasis === 'refund'
        ? incomeFromRefundDetail
        : incomeFromOrderDetail
    const expectedIncomeExcludedFeeAmount = incomeBasis === 'refund'
      ? 0
      : normalizeMoney(current.__expectedIncomeExcludedFeeAmount)
    const expectedIncomeBeforeFreight = normalizeMoney(
      primaryIncomeBeforeFreight - expectedIncomeExcludedFeeAmount
    )
    const expectedFromIncomeAndFreight = normalizeMoney(
      primaryIncomeBeforeFreight - logisticsExpenseTotal
    )
    const netAfterPurchaseAndFreight = normalizeMoney(expectedFromIncomeAndFreight - purchaseExpense)
    const expectedFromRefundAndFreight = normalizeMoney(
      incomeFromRefundDetail - incomeLogisticsExpense - jzgFreightExpense - offlineFreightExpense
    )
    const expectedIncomeDiff = normalizeMoney(expectedFromIncomeAndFreight - expectedFromRefundAndFreight)
    const usdRate = normalizeMoney(usdExchangeRate)
    const tariffByUsdInCny = normalizeMoney(toNumericValue(current.金掌柜计费金额合计USD) * usdRate)
    const tariffByUsdRaw = normalizeMoney(current.金掌柜计费金额合计USD)
    const tariffCandidateAmount = normalizeMoney(current.放退款_税费候选金额合计)
    const tariffByRefundCandidate = tariffCandidateAmount
    const incomeSideTaxFeeExcluded = normalizeMoney(current.__incomeStatementTaxFeeExcludedAmount)
    const tariffByIncomeStatement = incomeSideTaxFeeExcluded
    const hasIncomeSideTariff = Math.abs(tariffByIncomeStatement) > 0.000001
    const hasJzgTariff = Math.abs(tariffByUsdInCny) > 0.000001
    const hasJzgUsdRaw = Math.abs(tariffByUsdRaw) > 0.000001
    const tariffReferenceAmount = normalizeMoney(tariffByIncomeStatement + tariffByUsdInCny)
    const tariffReferenceSource = hasIncomeSideTariff && hasJzgTariff
      ? '收支表+金掌柜USD折算'
      : hasIncomeSideTariff
        ? '收支表'
        : hasJzgTariff
          ? '金掌柜USD折算'
          : ''
    const tariffDeductionForNet = normalizeMoney(tariffByIncomeStatement + tariffByUsdInCny)
    const refundTypeSummary = normalizeCellValue(current.放退款类型)
    const isDisputeRefundOrder = refundTypeSummary.includes('纠纷')
    const netAfterPurchaseFreightAndTariff = normalizeMoney(
      isDisputeRefundOrder ? netAfterPurchaseAndFreight : netAfterPurchaseAndFreight - tariffDeductionForNet
    )
    const tariffDiff = normalizeMoney(tariffByRefundCandidate - tariffReferenceAmount)
    const hasTariffInputs =
      Math.abs(tariffByRefundCandidate) > 0.000001 || Math.abs(tariffReferenceAmount) > 0.000001
    const tariffStatus = !hasTariffInputs
      ? ''
      : hasJzgUsdRaw && usdRate <= 0
        ? '缺少汇率'
        : Math.abs(tariffDiff) <= 1
          ? '基本对上'
          : '不一致'

    current.收入_按订单明细 = hasIncomeStatementRows ? incomeFromIncomeStatement : incomeFromOrderDetail
    current.收入_按放退款明细 = incomeFromRefundDetail
    current.收支总和_不含物流费用 = incomeFromOrderDetailExcludingLogistics
    current.差异_收支不含物流减退放款 = incomeDiff
    current.收入差异_订单明细减放退款 = incomeDiff
    current.收入校验状态 = Math.abs(incomeDiff) <= 0.01 ? '一致' : '不一致'
    current.预计可得_按收支及运费 = expectedFromIncomeAndFreight
    current.预计可得_按退放款及运费 = expectedFromRefundAndFreight
    current.预计可得差异_收支减退放款 = expectedIncomeDiff
    current.预计可得校验状态 = isNonRussiaIncomeStatementOnlyOrder
      ? '-'
      : Math.abs(expectedIncomeDiff) <= 0.01 ? '一致' : '不一致'
    current.最终收入_未扣运费 = expectedIncomeBeforeFreight
    current.最终收入_扣运费 = netAfterPurchaseFreightAndTariff
    current.税费金额_金掌柜 = tariffByUsdInCny
    current.税费金额_收支表 = tariffByIncomeStatement
    current.税费金额_放退款 = tariffByRefundCandidate
    current.税费核对_USD汇率 = usdRate
    current.税费核对_按金掌柜USD折CNY = tariffByUsdInCny
    current.税费核对_按收支表税费 = tariffByIncomeStatement
    current.税费核对_基准来源 = tariffReferenceSource
    current.税费核对_按放退款税费候选金额 = tariffByRefundCandidate
    current.税费核对_差异 = tariffDiff
    current.税费核对状态 = tariffStatus

    current.收入合计 = primaryIncomeBeforeFreight
    current.支出合计 = normalizeMoney(Math.max(totalFreight, 0) + purchaseExpense)
    current.总收支 = netAfterPurchaseFreightAndTariff

    aggregatedMap.set(orderNo, current)
  })

  const aggregatedRows = sortByRefundTypePriority(
    Array.from(aggregatedMap.values()).map((row) => {
      const next = { ...row }
      const pendingSettlement = normalizeMoney(next.待结算金额合计)
      const hasRefundRecord = Math.abs(toNumericValue(next.放退款_金额项合计)) > 0.000001
      const pendingRemark =
        pendingSettlement > 0 && !hasRefundRecord && isCompletedOrderStatus(normalizeCellValue(next.订单状态))
          ? '待结算金额在放退款明细无记录（交易已完成）'
          : ''
      const hasJzgFreight = Math.abs(toNumericValue(next.金掌柜物流费支出)) > 0.000001
      const noFreightRemark = !hasJzgFreight ? '金掌柜无物流费记录' : ''
      const mergedRemark = [pendingRemark, noFreightRemark].filter(Boolean).join('；')
      allTypeColumns.forEach((col) => {
        if (next[col] === undefined) {
          next[col] = 0
        }
      })
      allFeeItemColumns.forEach((col) => {
        if (next[col] === undefined) {
          next[col] = 0
        }
      })
      allLogisticsSubItemColumns.forEach((col) => {
        if (next[col] === undefined) {
          next[col] = 0
        }
      })
      allOfflineCategoryColumns.forEach((col) => {
        if (next[col] === undefined) {
          next[col] = 0
        }
      })

      return {
        订单号: normalizeOrderNo(next.订单号),
        订单时间: normalizeCellValue(next.订单时间),
        订单状态: normalizeCellValue(next.订单状态),
        是否放款: normalizeCellValue(next.是否放款) || '否',
        订单预计可得: normalizeMoney(next.最终收入_未扣运费),
        物流费用_支出表: normalizeMoney(next.收支表_支出物流费用),
        物流费用_金掌柜: normalizeMoney(next.金掌柜物流费支出),
        线下物流: normalizeMoney(next.线下运费),
        总物流费用: normalizeMoney(next.物流支出总和),
        采购费用: normalizeMoney(next.采购费用),
        支付宝是否开发票: normalizeCellValue(next.支付宝是否开发票),
        净利润: normalizeMoney(next.最终收入_扣运费),
        放退款类型: normalizeCellValue(next.放退款类型),
        放退款核查标记: normalizeCellValue(next.放退款核查标记),
        待结算金额合计: pendingSettlement,
        放退款_金额项合计: normalizeMoney(next.放退款_金额项合计),
        放退款_其他费用合计: normalizeMoney(next.放退款_其他费用合计),
        收入_按放退款明细: normalizeMoney(next.收入_按放退款明细),
        收入_按订单明细: normalizeMoney(next.收入_按订单明细),
        订单明细_放款金额合计: normalizeMoney(next.订单明细_放款金额合计),
        订单明细_售中退款金额合计: normalizeMoney(next.订单明细_售中退款金额合计),
        订单明细_售后退款金额合计: normalizeMoney(next.订单明细_售后退款金额合计),
        订单明细_平台分账金额合计: normalizeMoney(next.订单明细_平台分账金额合计),
        订单明细_平台分账退回金额合计: normalizeMoney(next.订单明细_平台分账退回金额合计),
        订单明细_净放款基准金额: normalizeMoney(next.订单明细_净放款基准金额),
        订单明细_净放款口径收入: normalizeMoney(next.订单明细_净放款口径收入),
        金掌柜计费金额合计USD: normalizeMoney(next.金掌柜计费金额合计USD),
        税费金额_金掌柜: normalizeMoney(next.税费金额_金掌柜),
        税费金额_收支表: normalizeMoney(next.税费金额_收支表),
        税费金额_放退款: normalizeMoney(next.税费金额_放退款),
        税费核对_USD汇率: normalizeMoney(next.税费核对_USD汇率),
        税费核对_基准来源: normalizeCellValue(next.税费核对_基准来源),
        税费核对_差异: normalizeMoney(next.税费核对_差异),
        税费核对状态: normalizeCellValue(next.税费核对状态),
        放退款_佣金合计: normalizeMoney(next.放退款_佣金合计),
        收入_收支明细表: normalizeMoney(next.收入_收支明细表),
        支出_收支明细表: normalizeMoney(next.支出_收支明细表),
        收支总和_不含物流费用: normalizeMoney(next.收支总和_不含物流费用),
        差异_收支不含物流减退放款: normalizeMoney(next.差异_收支不含物流减退放款),
        收入差异_订单明细减放退款: normalizeMoney(next.收入差异_订单明细减放退款),
        预计可得_按收支及运费: normalizeMoney(next.预计可得_按收支及运费),
        预计可得_按退放款及运费: normalizeMoney(next.预计可得_按退放款及运费),
        预计可得差异_收支减退放款: normalizeMoney(next.预计可得差异_收支减退放款),
        预计可得校验状态: normalizeCellValue(next.预计可得校验状态),
        备注: mergedRemark,
        ...Object.fromEntries(allOfflineCategoryColumns.map((col) => [col, normalizeMoney(next[col])])),
        ...Object.fromEntries(allLogisticsSubItemColumns.map((col) => [col, normalizeMoney(next[col])])),
        ...Object.fromEntries(allTypeColumns.map((col) => [col, normalizeMoney(next[col])])),
        ...Object.fromEntries(allFeeItemColumns.map((col) => [col, normalizeMoney(next[col])])),
      }
    }),
    '订单号',
    '放退款类型'
  )

  const incomeValidationRows = sortByOrderNo(
    aggregatedRows.map((row) => ({
      订单号: normalizeCellValue(row.订单号),
      收入_按订单明细: normalizeMoney(row.收入_按订单明细),
      收入_按放退款明细: normalizeMoney(row.收入_按放退款明细),
      收支总和_不含物流费用: normalizeMoney(row.收支总和_不含物流费用),
      差异_收支不含物流减退放款: normalizeMoney(row.差异_收支不含物流减退放款),
      收入差异_订单明细减放退款: normalizeMoney(row.收入差异_订单明细减放退款),
      放退款核查标记: normalizeCellValue(row.放退款核查标记),
      总物流费用: normalizeMoney(row.总物流费用),
      订单预计可得: normalizeMoney(row.订单预计可得),
      净利润: normalizeMoney(row.净利润)
    })),
    '订单号'
  )

  const ordersWithoutFreightRows = sortByOrderNo(
    aggregatedRows
      .filter((row) => Math.abs(toNumericValue(row.物流费用_金掌柜)) < 0.000001)
      .map((row) => ({
        订单号: normalizeCellValue(row.订单号),
        订单状态: normalizeCellValue(row.订单状态),
        备注: '金掌柜无物流费记录'
      })),
    '订单号'
  )

  const summary = aggregatedRows.reduce<ResultSummary>(
    (acc, row) => {
      const orderIncome = toNumericValue(row.收入_按订单明细)
      const logistics = normalizeMoney(toNumericValue(row.总物流费用))
      const purchaseExpense = normalizeMoney(toNumericValue(row.采购费用))
      const orderNo = normalizeOrderNo(row.订单号)
      const finalIncome = toNumericValue(row.净利润)
      const incomeBeforeFreight = toNumericValue(row.订单预计可得)
      const mismatch = normalizeCellValue(row.预计可得校验状态) === '不一致'
      const refundTypeText = normalizeCellValue(row.放退款类型)

      if (refundTypeText.includes('取消订单退款')) {
        acc.cancelRefundOrderCount += 1
      } else if (refundTypeText.includes('纠纷订单退款') || refundTypeText.includes('有纠纷订单退款')) {
        acc.disputeOrderCount += 1
      } else if (refundTypeText.includes('放款')) {
        acc.payoutOrderCount += 1
      } else {
        acc.otherOrderCount += 1
      }

      acc.totalIncome += orderIncome
      acc.totalExpense += logistics + purchaseExpense
      acc.logisticsAmount += logistics
      acc.nonLogisticsAmount += purchaseExpense
      if (invoicedAlipayOrderSet.has(orderNo)) {
        acc.totalInvoicedPurchaseAmount += purchaseExpense
      }
      acc.totalAmount += finalIncome
      acc.totalIncomeBeforeFreight += incomeBeforeFreight
      acc.totalIncomeAfterOnlineFreight += finalIncome
      if (mismatch) {
        acc.mismatchOrderCount += 1
      }

      return acc
    },
    {
      orderCount: uniqueOrderIds.length,
      totalAmount: 0,
      totalIncome: 0,
      totalExpense: 0,
      logisticsAmount: 0,
      nonLogisticsAmount: 0,
      rowCount: sortedPerformanceRows.length,
      mismatchOrderCount: 0,
      payoutOrderCount: 0,
      cancelRefundOrderCount: 0,
      disputeOrderCount: 0,
      otherOrderCount: 0,
      totalIncomeBeforeFreight: 0,
      totalIncomeAfterOnlineFreight: 0,
      totalInvoicedPurchaseAmount: 0
    } as ResultSummary
  )

  summary.totalAmount = Number(summary.totalAmount.toFixed(2))
  summary.totalIncome = Number(summary.totalIncome.toFixed(2))
  summary.totalExpense = Number(summary.totalExpense.toFixed(2))
  summary.logisticsAmount = Number(summary.logisticsAmount.toFixed(2))
  summary.nonLogisticsAmount = Number(summary.nonLogisticsAmount.toFixed(2))
  summary.totalIncomeBeforeFreight = Number(summary.totalIncomeBeforeFreight.toFixed(2))
  summary.totalIncomeAfterOnlineFreight = Number(summary.totalIncomeAfterOnlineFreight.toFixed(2))
  summary.totalInvoicedPurchaseAmount = Number(summary.totalInvoicedPurchaseAmount.toFixed(2))

  return {
    performanceRows: sortedPerformanceRows,
    aggregatedRows,
    dynamicTypeColumns: allTypeColumns,
    dynamicFeeItemColumns: allFeeItemColumns,
    dynamicTypeFeeItemColumns: [],
    dynamicLogisticsSubItemColumns: allLogisticsSubItemColumns,
    dynamicOfflineCategoryColumns: allOfflineCategoryColumns,
    incomeValidationRows,
    ordersWithoutIncomeRows,
    incomeOnlyOrdersRows,
    refundOnlyOrdersRows,
    ordersWithoutFreightRows,
    ordersWithoutAlipayRows,
    unmatchedAlipayRows,
    alipayMultiplicityRows,
    offlineFreightDiagnosticRows,
    summary,
    actualIncomeRows: sortedActualIncomeRows,
    actualFreightRows: sortedActualFreightRows,
    markedOrderRows: sortedMarkedOrderRows,
    integratedSummaryRows,
    integratedDetailRows: sortedIntegratedDetailRows
  }
}

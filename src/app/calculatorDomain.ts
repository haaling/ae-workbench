import * as XLSX from 'xlsx'

export type RowData = Record<string, string | number | boolean | null | undefined>

export type UploadedSheet = {
  name: string
  rows: RowData[]
  headers: string[]
}

export type UploadedFileState = {
  fileName: string
  sheets: UploadedSheet[]
  selectedSheetName: string
}

export type MultiUploadItem = {
  id: string
  file: UploadedFileState
}

export type IncomeUploadItem = {
  id: string
  file: UploadedFileState
  orderColumn: string
  flowTypeColumn?: string
}

export type TableKind = 'orders' | 'income' | 'refund' | 'freight' | 'alipay' | 'offline'

export type ProcessResult = {
  performanceRows: RowData[]
  aggregatedRows: RowData[]
  dynamicTypeColumns: string[]
  dynamicFeeItemColumns: string[]
  dynamicTypeFeeItemColumns: string[]
  dynamicLogisticsSubItemColumns: string[]
  dynamicOfflineCategoryColumns: string[]
  incomeValidationRows: RowData[]
  ordersWithoutIncomeRows: RowData[]
  incomeOnlyOrdersRows: RowData[]
  refundOnlyOrdersRows: RowData[]
  ordersWithoutFreightRows: RowData[]
  ordersWithoutAlipayRows: RowData[]
  unmatchedAlipayRows: RowData[]
  alipayMultiplicityRows: RowData[]
  offlineFreightDiagnosticRows: RowData[]
  summary: ResultSummary
  actualIncomeRows: RowData[]
  actualFreightRows: RowData[]
  markedOrderRows: RowData[]
  integratedSummaryRows: RowData[]
  integratedDetailRows: RowData[]
}

export type ResultSummary = {
  orderCount: number
  totalAmount: number
  totalIncome: number
  totalExpense: number
  logisticsAmount: number
  nonLogisticsAmount: number
  rowCount: number
  mismatchOrderCount: number
  payoutOrderCount: number
  cancelRefundOrderCount: number
  disputeOrderCount: number
  otherOrderCount: number
  totalIncomeBeforeFreight: number
  totalIncomeAfterOnlineFreight: number
  totalInvoicedPurchaseAmount: number
}

export type PersistedUploadBundle = {
  ordersFiles: MultiUploadItem[]
  incomeFiles: IncomeUploadItem[]
  refundFiles: IncomeUploadItem[]
  freightFiles: MultiUploadItem[]
  alipayFiles: MultiUploadItem[]
  offlineFiles: MultiUploadItem[]
  ordersIdColumn: string
  ordersStatusColumn: string
  ordersTimeColumn: string
  incomeDetailAmountColumns: string[]
  refundDetailAmountColumns: string[]
  refundTypeColumn: string
  refundProductNameColumn: string
  refundSkuIdColumn: string
  freightOrderColumn: string
  freightFulfillmentColumn: string
  freightWaybillColumn: string
  freightAmountCnyColumn: string
  freightAmountUsdColumn: string
  shopId: string
  shopName: string
  subsidiary: string
  usdExchangeRate: string
  workflowPeriod: string
}

export const UPLOAD_CACHE_STORAGE_KEY = 'performance_calculator_upload_cache_v1'

export const DEFAULT_ORDER_ID_HINTS = ['订单号', '订单编号', 'order', 'orderid', '订单']
export const DEFAULT_ORDER_STATUS_HINTS = ['订单状态', '状态', 'status']
export const DEFAULT_ORDER_TIME_HINTS = ['订单时间', '支付时间', '下单时间', '创建时间', '付款时间', 'time']
export const DEFAULT_FREIGHT_ID_HINTS = ['交易单号', '运单号', '交易号', '订单号', 'trade', 'waybill']
export const DEFAULT_FREIGHT_FULFILLMENT_HINTS = ['物流履约单号', '履约单号', '履约单', 'fulfillment']
export const DEFAULT_FREIGHT_WAYBILL_HINTS = ['运单号', '物流单号', 'waybill']
export const DEFAULT_FREIGHT_CNY_HINTS = ['计费金额合计cny', '计费金额cny', '金额cny', 'cny']
export const DEFAULT_FREIGHT_USD_HINTS = ['计费金额合计usd', '计费金额usd', '金额usd', 'usd']
export const DEFAULT_INCOME_DETAIL_EXCLUDE_HINTS = [
  '订单号',
  '订单编号',
  '支付时间',
  '发货时间',
  '确认收货时间',
  '成交金额',
  '收支记录状态',
  '记录状态',
  '收支来源',
  '匹配订单号'
]
export const DEFAULT_INCOME_DETAIL_AMOUNT_HINTS = [
  '收入金额',
  '支出金额',
  '收支金额',
  '金额',
  '计费金额',
  '佣金',
  '服务费',
  '营销',
  'cashback',
  '分账',
  '放款',
  '退款',
  '运费',
  '返利',
  '赔付',
  '手续费',
  '待结算金额',
  '售中退款金额',
  '放款金额',
  '售后退款金额'
]
export const DEFAULT_REFUND_TYPE_HINTS = ['放退款类型', '退款类型', '类型']
export const DEFAULT_INCOME_FLOW_TYPE_HINTS = ['收支类型', '收支方向', '资金方向', '流水类型']
export const DEFAULT_INCOME_FEE_ITEM_HINTS = ['费用项', '费用项目', '费用类型', '费用名称', '子项', '项目', '明细项']
export const DEFAULT_INCOME_MOVEMENT_HINTS = ['变动金额', '收支金额', '金额']
export const DEFAULT_REFUND_PRODUCT_NAME_HINTS = ['商品名称', '商品', '产品名称', '品名', 'product']
export const DEFAULT_REFUND_SKU_ID_HINTS = ['sku id', 'skuid', 'sku_id', 'sku', 'sku编号']
export const REFUND_BASE_AMOUNT_HINTS = ['放款金额', '退款金额', '放退款金额', '待结算金额']
export const DEFAULT_ALIPAY_REMARK_HINTS = ['备注', '订单备注', '商家备注', '说明', 'memo']
export const DEFAULT_ALIPAY_AMOUNT_HINTS = ['金额', '实付', '付款', '交易金额', '订单金额', '支出']
export const DEFAULT_ALIPAY_INVOICE_HINTS = ['是否开发票', '开票', '发票', '是否开票']
export const DEFAULT_ALIPAY_TRADE_NO_HINTS = ['支付宝交易号', '交易号', '交易流水号', 'trade no', '交易单号']
export const DEFAULT_ALIPAY_COUNTERPART_HINTS = ['交易对方', '对方账户', '对方姓名', '对方']
export const DEFAULT_ALIPAY_PRODUCT_HINTS = ['商品说明', '商品名称', '商品', '交易说明', '说明']
export const DEFAULT_ALIPAY_PAY_METHOD_HINTS = ['收/付款方式', '付款方式', '收款方式', '支付方式']
export const DEFAULT_ALIPAY_TRADE_ORDER_HINTS = ['交易订单号', '订单号', '业务订单号']
export const DEFAULT_ALIPAY_MERCHANT_ORDER_HINTS = ['商家订单号', '商户订单号', '外部订单号']
export const DEFAULT_OFFLINE_REMARK_HINTS = ['备注', '订单备注', '商家备注', '说明', 'memo']
export const DEFAULT_OFFLINE_AMOUNT_HINTS = ['运费', '金额', '费用', '实付', '付款', '支出']
export const DEFAULT_OFFLINE_ORDER_HINTS = ['客户订单号', '客户订单', '订单号', '订单编号', '外部订单号']
export const DEFAULT_OFFLINE_CATEGORY_HINTS = ['费用类别', '费用类型', '类别', '类型', '费用项', '费项']

export type ExternalRecord = {
  订单号: string
  金额: number
  备注: string
  是否开发票: string
  店铺名: string
  订单号来源: string
  支付宝交易单号: string
  交易对方: string
  商品说明: string
  收付款方式: string
  交易订单号: string
  商家订单号: string
  来源文件: string
  来源Sheet: string
}

export function hasKeyword(text: string, keywords: string[]): boolean {
  const value = normalizeCellValue(text)
  return keywords.some((keyword) => value.includes(keyword))
}

export function sortRefundTypesForDisplay(types: string[]): string[] {
  const deduped = Array.from(new Set(types.map((item) => normalizeCellValue(item)).filter(Boolean)))
  const isCancel = (text: string) => text.includes('放款前退款>取消订单退款')
  const isDispute = (text: string) => text.includes('有纠纷订单退款')

  const normal = deduped.filter((item) => !isDispute(item) && !isCancel(item))
  const dispute = deduped.filter((item) => isDispute(item))
  const cancel = deduped.filter((item) => isCancel(item))

  return [...normal, ...dispute, ...cancel]
}

export function normalizeCellValue(value: unknown): string {
  if (value === null || value === undefined) {
    return ''
  }
  return String(value).trim()
}

export function normalizeOrderNo(value: unknown): string {
  const text = normalizeCellValue(value)
  if (!text) {
    return ''
  }

  const sanitized = text.replace(/\s+/g, '').replace(/^['"]|['"]$/g, '')
  if (!sanitized) {
    return ''
  }

  if (/^\d+(?:\.0+)?$/.test(sanitized)) {
    return sanitized.replace(/\.0+$/, '')
  }

  const sciMatch = sanitized.match(/^([+-]?\d+(?:\.\d+)?)[eE]([+-]?\d+)$/)
  if (!sciMatch) {
    return sanitized
  }

  const mantissa = sciMatch[1].replace('+', '')
  const exponent = Number(sciMatch[2])
  if (!Number.isFinite(exponent)) {
    return sanitized
  }

  const sign = mantissa.startsWith('-') ? '-' : ''
  const unsigned = mantissa.replace(/^[-+]/, '')
  const [intPartRaw, fracPartRaw = ''] = unsigned.split('.')
  const digits = `${intPartRaw}${fracPartRaw}`.replace(/^0+/, '') || '0'
  const decimalPos = intPartRaw.length + exponent

  const expanded = decimalPos <= 0
    ? `0.${'0'.repeat(Math.abs(decimalPos))}${digits}`
    : decimalPos >= digits.length
      ? `${digits}${'0'.repeat(decimalPos - digits.length)}`
      : `${digits.slice(0, decimalPos)}.${digits.slice(decimalPos)}`

  const normalized = expanded
    .replace(/\.0+$/, '')
    .replace(/^(\d+)\.0+$/, '$1')
    .replace(/^0+(\d)/, '$1')

  return `${sign}${normalized}`
}

export function toNumericValue(value: unknown): number {
  const text = normalizeCellValue(value)
  if (!text) {
    return 0
  }

  const normalized = text.replace(/,/g, '').replace(/[^\d.-]/g, '')
  const parsed = Number(normalized)
  return Number.isFinite(parsed) ? parsed : 0
}

function looksLikeOrderLikeLongNumber(value: unknown): boolean {
  const text = normalizeCellValue(value).replace(/[,，\s]/g, '')
  return /^\d{14,}(?:\.0+)?$/.test(text)
}

function looksLikeOrderLikeSourceLabel(sourceLabel: unknown): boolean {
  const text = normalizeCellValue(sourceLabel).toLowerCase()
  if (!text) {
    return false
  }

  return [
    '订单号',
    '订单编号',
    '订单id',
    '交易单号',
    '交易订单号',
    '商家订单号',
    '客户订单号',
    '平台订单号',
    'order',
    'orderid',
    'trade',
    'waybill'
  ].some((hint) => text.includes(hint))
}

export function convertFreightToMovement(value: unknown, context: unknown = '', sourceLabel: unknown = ''): number {
  const amount = toNumericValue(value)
  if (!amount) {
    return 0
  }

  if (looksLikeOrderLikeLongNumber(value) || looksLikeOrderLikeSourceLabel(sourceLabel)) {
    return 0
  }

  const text = normalizeCellValue(context)
  const isRefund = /退费|退款|返还|赔付|补偿/.test(text)
  if (isRefund) {
    return Number((-Math.abs(amount)).toFixed(2))
  }

  return Number(Math.abs(amount).toFixed(2))
}

export function normalizeMoney(value: unknown): number {
  return Number(toNumericValue(value).toFixed(2))
}

export function toTypeAmountColumn(typeName: string): string {
  const text = normalizeCellValue(typeName) || '未分类'
  return `${text}_金额`
}

export function toFeeItemAmountColumn(feeItemName: string): string {
  const text = normalizeCellValue(feeItemName) || '未分类费用项'
  return `${text}_费用项金额`
}

export function toTypeFeeItemAmountColumn(typeName: string, feeItemName: string): string {
  const typeText = normalizeCellValue(typeName) || '未分类收支类型'
  const feeText = normalizeCellValue(feeItemName) || '未分类费用项'
  return `${typeText}+${feeText}_类型费用项金额`
}

export function toLogisticsSubItemAmountColumn(sourceLabel: '收支表' | '金掌柜', feeItemName: string): string {
  const feeText = normalizeCellValue(feeItemName) || '未分类'
  return `物流支出_${sourceLabel}_${feeText}`
}

export function isOrderIncomeSource(source: string): boolean {
  const text = normalizeCellValue(source)
  return text === '订单明细表' || text === '订单收支明细表' || text === '订单明细/收支明细表'
}

export function isExcludedIncomeDetailColumn(header: string): boolean {
  const key = normalizeCellValue(header).toLowerCase()
  if (!key) {
    return true
  }

  return DEFAULT_INCOME_DETAIL_EXCLUDE_HINTS.some((hint) => key.includes(hint.toLowerCase()))
}

export function isLikelyIncomeDetailAmountColumn(header: string): boolean {
  const key = normalizeCellValue(header).toLowerCase()
  if (!key) {
    return false
  }

  if (isExcludedIncomeDetailColumn(header)) {
    return false
  }

  if (DEFAULT_INCOME_DETAIL_AMOUNT_HINTS.some((hint) => key.includes(hint.toLowerCase()))) {
    return true
  }

  return key.includes('金额')
}

export function inferIncomeDetailAmountColumns(headers: string[]): string[] {
  return headers.filter((header) => isLikelyIncomeDetailAmountColumn(header))
}

export function toIncomeDetailMovement(columnName: string, value: unknown): number {
  const amount = toNumericValue(value)
  if (!amount) {
    return 0
  }

  const abs = Math.abs(amount)
  const key = normalizeCellValue(columnName).toLowerCase()

  if (key.includes('退回') || key.includes('退费') || key.includes('返还') || key.includes('赔付')) {
    return Number(abs.toFixed(2))
  }

  if (
    key.includes('佣金') ||
    key.includes('服务费') ||
    key.includes('营销') ||
    key.includes('cashback') ||
    key.includes('分账') ||
    key.includes('退款') ||
    key.includes('手续费') ||
    key.includes('费用')
  ) {
    return Number((-abs).toFixed(2))
  }

  return Number(amount.toFixed(2))
}

export function toIncomeDetailMovementByFlowType(columnName: string, value: unknown, flowType: unknown): number {
  const amount = toNumericValue(value)
  if (!amount) {
    return 0
  }

  const abs = Math.abs(amount)
  const flow = normalizeCellValue(flowType).toLowerCase()
  if (flow.includes('支出')) {
    return Number((-abs).toFixed(2))
  }
  if (flow.includes('收入')) {
    return Number(abs.toFixed(2))
  }

  return toIncomeDetailMovement(columnName, value)
}

export function toRefundDetailMovement(columnName: string, value: unknown, refundType: unknown): number {
  const amount = toNumericValue(value)
  if (!amount) {
    return 0
  }

  const abs = Math.abs(amount)
  const typeText = normalizeCellValue(refundType).toLowerCase()
  if (
    typeText.includes('客户取消') ||
    typeText.includes('买家取消') ||
    typeText.includes('取消订单')
  ) {
    return 0
  }

  if (typeText.includes('退款')) {
    return Number((-abs).toFixed(2))
  }
  if (typeText.includes('放款')) {
    return Number(abs.toFixed(2))
  }

  return toIncomeDetailMovement(columnName, value)
}

export function sortByOrderNo<T extends RowData>(rows: T[], key: string): T[] {
  const list = [...rows]
  list.sort((a, b) => {
    const left = normalizeCellValue(a[key])
    const right = normalizeCellValue(b[key])
    return left.localeCompare(right, 'zh-Hans-CN', { numeric: true, sensitivity: 'base' })
  })
  return list
}

export function sortByRefundTypePriority<T extends RowData>(rows: T[], orderKey: string, refundTypeKey: string): T[] {
  const list = [...rows]
  const rank = (refundType: string): number => {
    if (refundType.includes('取消订单退款')) {
      return 2
    }
    if (refundType.includes('纠纷订单退款') || refundType.includes('有纠纷订单退款')) {
      return 1
    }
    return 0
  }

  list.sort((a, b) => {
    const leftType = normalizeCellValue(a[refundTypeKey])
    const rightType = normalizeCellValue(b[refundTypeKey])
    const leftRank = rank(leftType)
    const rightRank = rank(rightType)

    if (leftRank !== rightRank) {
      return leftRank - rightRank
    }

    const leftOrder = normalizeCellValue(a[orderKey])
    const rightOrder = normalizeCellValue(b[orderKey])
    return leftOrder.localeCompare(rightOrder, 'zh-Hans-CN', { numeric: true, sensitivity: 'base' })
  })

  return list
}

export function isCompletedOrderStatus(status: string): boolean {
  const text = normalizeCellValue(status)
  if (!text) {
    return false
  }
  return ['交易完成', '交易成功', '已完成', '已签收', '完成'].some((keyword) => text.includes(keyword))
}

export function safeFileSegment(text: string, fallback: string): string {
  const value = normalizeCellValue(text).replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, '')
  return value || fallback
}

export function extractOrderNoFromRemark(remark: string): string {
  const text = normalizeCellValue(remark)
  if (!text) {
    return ''
  }

  // Alipay remark extraction uses strict 16-digit continuous order ids only.
  const candidates = Array.from(text.matchAll(/(?:^|\D)(\d{16})(?=\D|$)/g)).map((match) => match[1])
  return candidates.length > 0 ? candidates[candidates.length - 1] : ''
}

export function matchesShopRemarkPrefix(remark: string, currentShopName: string): boolean {
  const shop = normalizeCellValue(currentShopName).replace(/\s+/g, '').toLowerCase()
  if (!shop) {
    return true
  }

  const text = normalizeCellValue(remark).replace(/\s+/g, '').toLowerCase()
  if (!text) {
    return false
  }

  return (
    text.startsWith(shop) ||
    text.startsWith(`${shop}-`) ||
    text.startsWith(`${shop}—`) ||
    text.startsWith(`${shop}_`) ||
    text.includes(`${shop}-`) ||
    text.includes(`${shop}—`) ||
    text.includes(`${shop}_`)
  )
}

export function readUploadCache(): PersistedUploadBundle | null {
  if (typeof window === 'undefined') {
    return null
  }

  try {
    const raw = window.localStorage.getItem(UPLOAD_CACHE_STORAGE_KEY)
    if (!raw) {
      return null
    }
    const parsed = JSON.parse(raw) as PersistedUploadBundle
    return parsed
  } catch (error) {
    console.warn('读取本地上传缓存失败', error)
    return null
  }
}

export function writeUploadCache(payload: PersistedUploadBundle): { ok: boolean; reason?: 'quota' | 'unknown' } {
  if (typeof window === 'undefined') {
    return { ok: false, reason: 'unknown' }
  }

  try {
    window.localStorage.setItem(UPLOAD_CACHE_STORAGE_KEY, JSON.stringify(payload))
    return { ok: true }
  } catch (error) {
    const message = error instanceof Error ? error.message.toLowerCase() : ''
    const isQuota = message.includes('quota') || message.includes('exceeded')
    console.warn('写入本地上传缓存失败', error)
    return { ok: false, reason: isQuota ? 'quota' : 'unknown' }
  }
}

export function clearUploadCache() {
  if (typeof window === 'undefined') {
    return
  }
  window.localStorage.removeItem(UPLOAD_CACHE_STORAGE_KEY)
}

export function pickDefaultColumn(headers: string[], hints: string[]): string {
  if (headers.length === 0) {
    return ''
  }

  const loweredHeaders = headers.map((h) => h.toLowerCase())
  const hit = hints.find((hint) => loweredHeaders.some((header) => header.includes(hint.toLowerCase())))
  if (!hit) {
    return headers[0]
  }

  return headers[loweredHeaders.findIndex((header) => header.includes(hit.toLowerCase()))] || headers[0]
}

export function pickOptionalColumn(headers: string[], hints: string[]): string {
  if (headers.length === 0) {
    return ''
  }

  const loweredHeaders = headers.map((h) => h.toLowerCase())
  const hit = hints.find((hint) => loweredHeaders.some((header) => header.includes(hint.toLowerCase())))
  if (!hit) {
    return ''
  }

  return headers[loweredHeaders.findIndex((header) => header.includes(hit.toLowerCase()))] || ''
}

export function resolveColumnSelection(selected: string, headers: string[], hints: string[]): string {
  if (selected && headers.includes(selected)) {
    return selected
  }
  return pickDefaultColumn(headers, hints)
}

export function toUploadedState(fileName: string, workbook: XLSX.WorkBook): UploadedFileState {
  const sheets = workbook.SheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName]
    const rows = XLSX.utils.sheet_to_json<RowData>(worksheet, {
      defval: '',
      raw: false
    })

    const headers = Array.from(
      rows.reduce((set, row) => {
        Object.keys(row).forEach((key) => set.add(key))
        return set
      }, new Set<string>())
    )

    return {
      name: sheetName,
      rows,
      headers
    }
  })

  return {
    fileName,
    sheets,
    selectedSheetName: sheets[0]?.name || ''
  }
}

export function getSelectedSheet(file: UploadedFileState | null): UploadedSheet | null {
  if (!file) {
    return null
  }
  return file.sheets.find((sheet) => sheet.name === file.selectedSheetName) || file.sheets[0] || null
}

export function buildCountMap(rows: RowData[], key: string): Map<string, number> {
  const map = new Map<string, number>()
  rows.forEach((row) => {
    const id = normalizeOrderNo(row[key])
    if (!id) {
      return
    }
    map.set(id, (map.get(id) || 0) + 1)
  })
  return map
}

export function buildRowsByOrderMap(rows: RowData[], key: string): Map<string, RowData[]> {
  const map = new Map<string, RowData[]>()
  rows.forEach((row) => {
    const id = normalizeOrderNo(row[key])
    if (!id) {
      return
    }
    const current = map.get(id) || []
    current.push(row)
    map.set(id, current)
  })
  return map
}

export function prefixRow(row: RowData, prefix: string): RowData {
  return Object.fromEntries(Object.entries(row).map(([key, value]) => [`${prefix}${key}`, value]))
}

export function summarizeAmountByOrder(rows: RowData[]): Map<string, { count: number; amount: number; refundTypes: Set<string> }> {
  const map = new Map<string, { count: number; amount: number; refundTypes: Set<string> }>()
  rows.forEach((row) => {
    const orderNo = normalizeOrderNo(row.订单号)
    if (!orderNo) {
      return
    }

    const current = map.get(orderNo) || { count: 0, amount: 0, refundTypes: new Set<string>() }
    current.count += 1
    current.amount = normalizeMoney(current.amount + toNumericValue(row.变动金额))
    const refundType = normalizeCellValue(row.放退款类型)
    if (refundType) {
      current.refundTypes.add(refundType)
    }
    map.set(orderNo, current)
  })
  return map
}

export function buildRefundRiskMap(rows: Array<{ 订单号: string; 商品名称: string; skuId: string }>): Map<string, string> {
  const map = new Map<string, string>()
  const grouped = new Map<string, Array<{ 商品名称: string; skuId: string }>>()

  rows.forEach((row) => {
    const orderNo = normalizeCellValue(row.订单号)
    if (!orderNo) {
      return
    }
    const list = grouped.get(orderNo) || []
    list.push({
      商品名称: normalizeCellValue(row.商品名称),
      skuId: normalizeCellValue(row.skuId)
    })
    grouped.set(orderNo, list)
  })

  grouped.forEach((items, orderNo) => {
    const hasMultiRows = items.length >= 2
    const nonLogisticsRows = items.filter((item) => !/(物流费|运费|物流)/.test(item.商品名称))

    const skuCountMap = new Map<string, number>()
    nonLogisticsRows.forEach((item) => {
      if (!item.skuId) {
        return
      }
      skuCountMap.set(item.skuId, (skuCountMap.get(item.skuId) || 0) + 1)
    })

    const hasSameSkuInNonLogistics = Array.from(skuCountMap.values()).some((count) => count >= 2)

    if (hasMultiRows && nonLogisticsRows.length >= 2 && hasSameSkuInNonLogistics) {
      map.set(orderNo, '关税支出')
    }
  })

  return map
}

export function dedupeByKey<T>(rows: T[], toKey: (row: T) => string): T[] {
  const map = new Map<string, T>()
  rows.forEach((row) => {
    const key = toKey(row)
    if (!map.has(key)) {
      map.set(key, row)
    }
  })
  return Array.from(map.values())
}

export function validateDetailFiles(files: IncomeUploadItem[], label: string): string | null {
  for (const item of files) {
    const sheet = getSelectedSheet(item.file)
    if (!sheet) {
      return `${label}存在未选择工作表的文件：${item.file.fileName}`
    }

    const hasAnyContent = sheet.rows.some((row) =>
      Object.values(row).some((value) => normalizeCellValue(value))
    )
    if (!hasAnyContent) {
      continue
    }

    const hasOrderNo = sheet.rows.some((row) => normalizeCellValue(row[item.orderColumn]))
    if (!hasOrderNo) {
      return `${label}文件 ${item.file.fileName} 未识别到有效订单号。`
    }
  }
  return null
}

export function findFreightOrderNoInRow(row: RowData, preferredColumn: string, orderIds: Set<string>): string {
  const preferred = normalizeOrderNo(row[preferredColumn])
  if (preferred && orderIds.has(preferred)) {
    return preferred
  }

  for (const value of Object.values(row)) {
    const text = normalizeOrderNo(value)
    if (text && orderIds.has(text)) {
      return text
    }
  }

  return preferred
}

export function findFreightFallbackAmount(row: RowData): unknown {
  const entries = Object.entries(row)
  const likely = entries.find(([key, value]) => {
    const keyText = normalizeCellValue(key).toLowerCase()
    if (!(keyText.includes('金额') || keyText.includes('运费') || keyText.includes('费用') || keyText.includes('cny') || keyText.includes('usd'))) {
      return false
    }
    return Math.abs(toNumericValue(value)) > 0.000001
  })

  return likely?.[1] ?? ''
}

export function formatDateTime(value: Date): string {
  const yyyy = value.getFullYear()
  const mm = String(value.getMonth() + 1).padStart(2, '0')
  const dd = String(value.getDate()).padStart(2, '0')
  const hh = String(value.getHours()).padStart(2, '0')
  const min = String(value.getMinutes()).padStart(2, '0')
  const ss = String(value.getSeconds()).padStart(2, '0')
  return `${yyyy}-${mm}-${dd} ${hh}:${min}:${ss}`
}


import { useEffect, useMemo, useState, type ReactNode } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

type RowData = Record<string, string | number | boolean | null | undefined>

type UploadedSheet = {
  name: string
  rows: RowData[]
  headers: string[]
}

type UploadedFileState = {
  fileName: string
  sheets: UploadedSheet[]
  selectedSheetName: string
}

type MultiUploadItem = {
  id: string
  file: UploadedFileState
}

type IncomeUploadItem = {
  id: string
  file: UploadedFileState
  orderColumn: string
  flowTypeColumn?: string
}

type LedgerUploadItem = {
  id: string
  file: UploadedFileState
  orderColumn: string
  purchaseColumn: string
  freightColumn: string
}

type TableKind = 'orders' | 'income' | 'refund' | 'freight' | 'alipay' | 'offline' | 'ledger'

type ProcessResult = {
  performanceRows: RowData[]
  aggregatedRows: RowData[]
  dynamicTypeColumns: string[]
  dynamicFeeItemColumns: string[]
  dynamicTypeFeeItemColumns: string[]
  dynamicLogisticsSubItemColumns: string[]
  incomeValidationRows: RowData[]
  ordersWithoutIncomeRows: RowData[]
  incomeOnlyOrdersRows: RowData[]
  refundOnlyOrdersRows: RowData[]
  ordersWithoutFreightRows: RowData[]
  unmatchedAlipayRows: RowData[]
  summary: ResultSummary
  actualIncomeRows: RowData[]
  actualFreightRows: RowData[]
  markedOrderRows: RowData[]
  integratedSummaryRows: RowData[]
  integratedDetailRows: RowData[]
}

type ResultSummary = {
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
}

type PersistedUploadBundle = {
  ordersFiles: MultiUploadItem[]
  incomeFiles: IncomeUploadItem[]
  refundFiles: IncomeUploadItem[]
  freightFiles: MultiUploadItem[]
  alipayFiles: MultiUploadItem[]
  offlineFiles: MultiUploadItem[]
  ledgerFiles: LedgerUploadItem[]
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
}

const UPLOAD_CACHE_STORAGE_KEY = 'performance_calculator_upload_cache_v1'

const DEFAULT_ORDER_ID_HINTS = ['订单号', '订单编号', 'order', 'orderid', '订单']
const DEFAULT_ORDER_STATUS_HINTS = ['订单状态', '状态', 'status']
const DEFAULT_ORDER_TIME_HINTS = ['订单时间', '支付时间', '下单时间', '创建时间', '付款时间', 'time']
const DEFAULT_FREIGHT_ID_HINTS = ['交易单号', '运单号', '交易号', '订单号', 'trade', 'waybill']
const DEFAULT_FREIGHT_FULFILLMENT_HINTS = ['物流履约单号', '履约单号', '履约单', 'fulfillment']
const DEFAULT_FREIGHT_WAYBILL_HINTS = ['运单号', '物流单号', 'waybill']
const DEFAULT_FREIGHT_CNY_HINTS = ['计费金额合计cny', '计费金额cny', '金额cny', 'cny']
const DEFAULT_FREIGHT_USD_HINTS = ['计费金额合计usd', '计费金额usd', '金额usd', 'usd']
const DEFAULT_INCOME_DETAIL_EXCLUDE_HINTS = [
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
const DEFAULT_INCOME_DETAIL_AMOUNT_HINTS = [
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
  '补贴',
  '待结算金额',
  '售中退款金额',
  '放款金额',
  '售后退款金额'
]
const DEFAULT_REFUND_TYPE_HINTS = ['放退款类型', '退款类型', '类型']
const DEFAULT_INCOME_FLOW_TYPE_HINTS = ['收支类型', '收支方向', '资金方向', '流水类型']
const DEFAULT_INCOME_FEE_ITEM_HINTS = ['费用项', '费用项目', '费用类型', '费用名称', '子项', '项目', '明细项']
const DEFAULT_INCOME_MOVEMENT_HINTS = ['变动金额', '收支金额', '金额']
const DEFAULT_REFUND_PRODUCT_NAME_HINTS = ['商品名称', '商品', '产品名称', '品名', 'product']
const DEFAULT_REFUND_SKU_ID_HINTS = ['sku id', 'skuid', 'sku_id', 'sku', 'sku编号']
const REFUND_BASE_AMOUNT_HINTS = ['放款金额', '退款金额', '放退款金额', '待结算金额']
const DEFAULT_ALIPAY_REMARK_HINTS = ['备注', '订单备注', '商家备注', '说明', 'memo']
const DEFAULT_ALIPAY_AMOUNT_HINTS = ['金额', '实付', '付款', '交易金额', '订单金额', '支出']
const DEFAULT_ALIPAY_TRADE_NO_HINTS = ['支付宝交易号', '交易号', '交易流水号', 'trade no', '交易单号']
const DEFAULT_ALIPAY_COUNTERPART_HINTS = ['交易对方', '对方账户', '对方姓名', '对方']
const DEFAULT_ALIPAY_PRODUCT_HINTS = ['商品说明', '商品名称', '商品', '交易说明', '说明']
const DEFAULT_ALIPAY_PAY_METHOD_HINTS = ['收/付款方式', '付款方式', '收款方式', '支付方式']
const DEFAULT_ALIPAY_TRADE_ORDER_HINTS = ['交易订单号', '订单号', '业务订单号']
const DEFAULT_ALIPAY_MERCHANT_ORDER_HINTS = ['商家订单号', '商户订单号', '外部订单号']
const DEFAULT_OFFLINE_REMARK_HINTS = ['备注', '订单备注', '商家备注', '说明', 'memo']
const DEFAULT_OFFLINE_AMOUNT_HINTS = ['运费', '金额', '费用', '实付', '付款', '支出']
const DEFAULT_LEDGER_PURCHASE_HINTS = ['采购金额', '采购', '进货金额', '采购成本', '成本']
const DEFAULT_LEDGER_FREIGHT_HINTS = ['运费', '物流费', '快递费', '台账运费']
const LEDGER_MAPPING_STORAGE_KEY = 'performance_calculator_ledger_mappings_v1'

type ExternalRecord = {
  订单号: string
  金额: number
  备注: string
  支付宝交易单号: string
  交易对方: string
  商品说明: string
  收付款方式: string
  交易订单号: string
  商家订单号: string
  来源文件: string
  来源Sheet: string
}

type LedgerMappingProfile = {
  orderColumn: string
  purchaseColumn: string
  freightColumn: string
}

function hasKeyword(text: string, keywords: string[]): boolean {
  const value = normalizeCellValue(text)
  return keywords.some((keyword) => value.includes(keyword))
}

function sortRefundTypesForDisplay(types: string[]): string[] {
  const deduped = Array.from(new Set(types.map((item) => normalizeCellValue(item)).filter(Boolean)))
  const isCancel = (text: string) => text.includes('放款前退款>取消订单退款')
  const isDispute = (text: string) => text.includes('有纠纷订单退款')

  const normal = deduped.filter((item) => !isDispute(item) && !isCancel(item))
  const dispute = deduped.filter((item) => isDispute(item))
  const cancel = deduped.filter((item) => isCancel(item))

  return [...normal, ...dispute, ...cancel]
}

function normalizeCellValue(value: unknown): string {
  if (value === null || value === undefined) {
    return ''
  }
  return String(value).trim()
}

function normalizeOrderNo(value: unknown): string {
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

  let expanded = ''
  if (decimalPos <= 0) {
    expanded = `0.${'0'.repeat(Math.abs(decimalPos))}${digits}`
  } else if (decimalPos >= digits.length) {
    expanded = `${digits}${'0'.repeat(decimalPos - digits.length)}`
  } else {
    expanded = `${digits.slice(0, decimalPos)}.${digits.slice(decimalPos)}`
  }

  const normalized = expanded
    .replace(/\.0+$/, '')
    .replace(/^(\d+)\.0+$/, '$1')
    .replace(/^0+(\d)/, '$1')

  return `${sign}${normalized}`
}

function toNumericValue(value: unknown): number {
  const text = normalizeCellValue(value)
  if (!text) {
    return 0
  }

  const normalized = text.replace(/,/g, '').replace(/[^\d.-]/g, '')
  const parsed = Number(normalized)
  return Number.isFinite(parsed) ? parsed : 0
}

function convertFreightToMovement(value: unknown, context: unknown = ''): number {
  const amount = toNumericValue(value)
  if (!amount) {
    return 0
  }

  const text = normalizeCellValue(context)
  const isRefund = /退费|退款|返还|赔付|补偿/.test(text)
  if (isRefund) {
    return Number((-Math.abs(amount)).toFixed(2))
  }

  return Number(Math.abs(amount).toFixed(2))
}

function normalizeMoney(value: unknown): number {
  return Number(toNumericValue(value).toFixed(2))
}

function toTypeAmountColumn(typeName: string): string {
  const text = normalizeCellValue(typeName) || '未分类'
  return `${text}_金额`
}

function toFeeItemAmountColumn(feeItemName: string): string {
  const text = normalizeCellValue(feeItemName) || '未分类费用项'
  return `${text}_费用项金额`
}

function toTypeFeeItemAmountColumn(typeName: string, feeItemName: string): string {
  const typeText = normalizeCellValue(typeName) || '未分类收支类型'
  const feeText = normalizeCellValue(feeItemName) || '未分类费用项'
  return `${typeText}+${feeText}_类型费用项金额`
}

function toLogisticsSubItemAmountColumn(sourceLabel: '收支表' | '金掌柜', feeItemName: string): string {
  const feeText = normalizeCellValue(feeItemName) || '未分类'
  return `物流支出_${sourceLabel}_${feeText}`
}

function isOrderIncomeSource(source: string): boolean {
  const text = normalizeCellValue(source)
  return text === '订单明细表' || text === '订单收支明细表' || text === '订单明细/收支明细表'
}

function isExcludedIncomeDetailColumn(header: string): boolean {
  const key = normalizeCellValue(header).toLowerCase()
  if (!key) {
    return true
  }

  return DEFAULT_INCOME_DETAIL_EXCLUDE_HINTS.some((hint) => key.includes(hint.toLowerCase()))
}

function isLikelyIncomeDetailAmountColumn(header: string): boolean {
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

function inferIncomeDetailAmountColumns(headers: string[]): string[] {
  return headers.filter((header) => isLikelyIncomeDetailAmountColumn(header))
}

function toIncomeDetailMovement(columnName: string, value: unknown): number {
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

function toIncomeDetailMovementByFlowType(columnName: string, value: unknown, flowType: unknown): number {
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

function toRefundDetailMovement(columnName: string, value: unknown, refundType: unknown): number {
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

function sortByOrderNo<T extends RowData>(rows: T[], key: string): T[] {
  const list = [...rows]
  list.sort((a, b) => {
    const left = normalizeCellValue(a[key])
    const right = normalizeCellValue(b[key])
    return left.localeCompare(right, 'zh-Hans-CN', { numeric: true, sensitivity: 'base' })
  })
  return list
}

function sortByRefundTypePriority<T extends RowData>(rows: T[], orderKey: string, refundTypeKey: string): T[] {
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

function isCompletedOrderStatus(status: string): boolean {
  const text = normalizeCellValue(status)
  if (!text) {
    return false
  }
  return ['交易完成', '交易成功', '已完成', '已签收', '完成'].some((keyword) => text.includes(keyword))
}

function safeFileSegment(text: string, fallback: string): string {
  const value = normalizeCellValue(text).replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, '')
  return value || fallback
}

function extractOrderNoFromRemark(remark: string): string {
  const text = normalizeCellValue(remark)
  if (!text) {
    return ''
  }

  const dashMatch = text.match(/[-—_]\s*([A-Za-z0-9]+)\s*$/)
  if (dashMatch?.[1]) {
    return normalizeCellValue(dashMatch[1])
  }

  const parts = text.split(/[-—_]/).map((item) => normalizeCellValue(item)).filter(Boolean)
  if (parts.length >= 2) {
    return parts[parts.length - 1]
  }

  return ''
}

function matchesShopRemarkPrefix(remark: string, currentShopName: string): boolean {
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

function getLedgerProfileKey(fileName: string): string {
  const base = normalizeCellValue(fileName).replace(/\.[^.]+$/, '')
  if (!base) {
    return 'default'
  }
  const first = base.split(/[-_\s（(]/)[0]
  return normalizeCellValue(first) || base
}

function readLedgerMappings(): Record<string, LedgerMappingProfile> {
  try {
    const raw = globalThis.localStorage?.getItem(LEDGER_MAPPING_STORAGE_KEY)
    if (!raw) {
      return {}
    }
    const parsed = JSON.parse(raw)
    if (parsed && typeof parsed === 'object') {
      return parsed as Record<string, LedgerMappingProfile>
    }
  } catch (error) {
    console.warn('读取台账映射失败', error)
  }
  return {}
}

function readUploadCache(): PersistedUploadBundle | null {
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

function writeUploadCache(payload: PersistedUploadBundle): { ok: boolean; reason?: 'quota' | 'unknown' } {
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

function clearUploadCache() {
  if (typeof window === 'undefined') {
    return
  }
  window.localStorage.removeItem(UPLOAD_CACHE_STORAGE_KEY)
}

function writeLedgerMappings(next: Record<string, LedgerMappingProfile>) {
  try {
    globalThis.localStorage?.setItem(LEDGER_MAPPING_STORAGE_KEY, JSON.stringify(next))
  } catch (error) {
    console.warn('保存台账映射失败', error)
  }
}

function pickDefaultColumn(headers: string[], hints: string[]): string {
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

function pickOptionalColumn(headers: string[], hints: string[]): string {
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

function resolveColumnSelection(selected: string, headers: string[], hints: string[]): string {
  if (selected && headers.includes(selected)) {
    return selected
  }
  return pickDefaultColumn(headers, hints)
}

function toUploadedState(fileName: string, workbook: XLSX.WorkBook): UploadedFileState {
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

function getSelectedSheet(file: UploadedFileState | null): UploadedSheet | null {
  if (!file) {
    return null
  }
  return file.sheets.find((sheet) => sheet.name === file.selectedSheetName) || file.sheets[0] || null
}

function buildCountMap(rows: RowData[], key: string): Map<string, number> {
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

function buildRowsByOrderMap(rows: RowData[], key: string): Map<string, RowData[]> {
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

function prefixRow(row: RowData, prefix: string): RowData {
  return Object.fromEntries(Object.entries(row).map(([key, value]) => [`${prefix}${key}`, value]))
}

function summarizeAmountByOrder(rows: RowData[]): Map<string, { count: number; amount: number; refundTypes: Set<string> }> {
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

function buildRefundRiskMap(rows: Array<{ 订单号: string; 商品名称: string; skuId: string }>): Map<string, string> {
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
      map.set(orderNo, '放退款多条且非物流费行存在相同SKU(疑似含关税未剔除)')
    }
  })

  return map
}

function dedupeByKey<T>(rows: T[], toKey: (row: T) => string): T[] {
  const map = new Map<string, T>()
  rows.forEach((row) => {
    const key = toKey(row)
    if (!map.has(key)) {
      map.set(key, row)
    }
  })
  return Array.from(map.values())
}

function validateDetailFiles(files: IncomeUploadItem[], label: string): string | null {
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

function findFreightOrderNoInRow(row: RowData, preferredColumn: string, orderIds: Set<string>): string {
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

function findFreightFallbackAmount(row: RowData): unknown {
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

function formatDateTime(value: Date): string {
  const yyyy = value.getFullYear()
  const mm = String(value.getMonth() + 1).padStart(2, '0')
  const dd = String(value.getDate()).padStart(2, '0')
  const hh = String(value.getHours()).padStart(2, '0')
  const min = String(value.getMinutes()).padStart(2, '0')
  const ss = String(value.getSeconds()).padStart(2, '0')
  return `${yyyy}-${mm}-${dd} ${hh}:${min}:${ss}`
}

function App() {
  const [ordersFiles, setOrdersFiles] = useState<MultiUploadItem[]>([])
  const [incomeFiles, setIncomeFiles] = useState<IncomeUploadItem[]>([])
  const [refundFiles, setRefundFiles] = useState<IncomeUploadItem[]>([])
  const [freightFiles, setFreightFiles] = useState<MultiUploadItem[]>([])
  const [alipayFiles, setAlipayFiles] = useState<MultiUploadItem[]>([])
  const [offlineFiles, setOfflineFiles] = useState<MultiUploadItem[]>([])
  const [ledgerFiles, setLedgerFiles] = useState<LedgerUploadItem[]>([])

  const [ordersIdColumn, setOrdersIdColumn] = useState('')
  const [ordersStatusColumn, setOrdersStatusColumn] = useState('')
  const [ordersTimeColumn, setOrdersTimeColumn] = useState('')
  const [incomeDetailAmountColumns, setIncomeDetailAmountColumns] = useState<string[]>([])
  const [refundDetailAmountColumns, setRefundDetailAmountColumns] = useState<string[]>([])
  const [refundTypeColumn, setRefundTypeColumn] = useState('')
  const [refundProductNameColumn, setRefundProductNameColumn] = useState('')
  const [refundSkuIdColumn, setRefundSkuIdColumn] = useState('')
  const [freightOrderColumn, setFreightOrderColumn] = useState('')
  const [freightFulfillmentColumn, setFreightFulfillmentColumn] = useState('')
  const [freightWaybillColumn, setFreightWaybillColumn] = useState('')
  const [freightAmountCnyColumn, setFreightAmountCnyColumn] = useState('')
  const [freightAmountUsdColumn, setFreightAmountUsdColumn] = useState('')

  const [isProcessing, setIsProcessing] = useState(false)
  const [errorMessage, setErrorMessage] = useState('')
  const [result, setResult] = useState<ProcessResult | null>(null)
  const [lastCalculatedAt, setLastCalculatedAt] = useState('')
  const [calculationCount, setCalculationCount] = useState(0)
  const [shopId, setShopId] = useState('')
  const [shopName, setShopName] = useState('')
  const [subsidiary, setSubsidiary] = useState('')
  const [uploadCacheReady, setUploadCacheReady] = useState(false)

  useEffect(() => {
    const cache = readUploadCache()
    if (cache) {
      setOrdersFiles(Array.isArray(cache.ordersFiles) ? cache.ordersFiles : [])
      setIncomeFiles(Array.isArray(cache.incomeFiles) ? cache.incomeFiles : [])
      setRefundFiles(Array.isArray(cache.refundFiles) ? cache.refundFiles : [])
      setFreightFiles(Array.isArray(cache.freightFiles) ? cache.freightFiles : [])
      setAlipayFiles(Array.isArray(cache.alipayFiles) ? cache.alipayFiles : [])
      setOfflineFiles(Array.isArray(cache.offlineFiles) ? cache.offlineFiles : [])
      setLedgerFiles(Array.isArray(cache.ledgerFiles) ? cache.ledgerFiles : [])
      setOrdersIdColumn(normalizeCellValue(cache.ordersIdColumn))
      setOrdersStatusColumn(normalizeCellValue(cache.ordersStatusColumn))
      setOrdersTimeColumn(normalizeCellValue(cache.ordersTimeColumn))
      setIncomeDetailAmountColumns(Array.isArray(cache.incomeDetailAmountColumns) ? cache.incomeDetailAmountColumns : [])
      setRefundDetailAmountColumns(Array.isArray(cache.refundDetailAmountColumns) ? cache.refundDetailAmountColumns : [])
      setRefundTypeColumn(normalizeCellValue(cache.refundTypeColumn))
      setRefundProductNameColumn(normalizeCellValue(cache.refundProductNameColumn))
      setRefundSkuIdColumn(normalizeCellValue(cache.refundSkuIdColumn))
      setFreightOrderColumn(normalizeCellValue(cache.freightOrderColumn))
      setFreightFulfillmentColumn(normalizeCellValue(cache.freightFulfillmentColumn))
      setFreightWaybillColumn(normalizeCellValue(cache.freightWaybillColumn))
      setFreightAmountCnyColumn(normalizeCellValue(cache.freightAmountCnyColumn))
      setFreightAmountUsdColumn(normalizeCellValue(cache.freightAmountUsdColumn))
      setShopId(normalizeCellValue(cache.shopId))
      setShopName(normalizeCellValue(cache.shopName))
      setSubsidiary(normalizeCellValue(cache.subsidiary))
    }
    setUploadCacheReady(true)
  }, [])

  useEffect(() => {
    if (!uploadCacheReady) {
      return
    }

    const payload: PersistedUploadBundle = {
      ordersFiles,
      incomeFiles,
      refundFiles,
      freightFiles,
      alipayFiles,
      offlineFiles,
      ledgerFiles,
      ordersIdColumn,
      ordersStatusColumn,
      ordersTimeColumn,
      incomeDetailAmountColumns,
      refundDetailAmountColumns,
      refundTypeColumn,
      refundProductNameColumn,
      refundSkuIdColumn,
      freightOrderColumn,
      freightFulfillmentColumn,
      freightWaybillColumn,
      freightAmountCnyColumn,
      freightAmountUsdColumn,
      shopId,
      shopName,
      subsidiary
    }

    const writeResult = writeUploadCache(payload)
    if (!writeResult.ok && writeResult.reason === 'quota') {
      setErrorMessage('本地缓存空间不足，建议减少上传文件数量或先清空本地缓存。')
    }
  }, [
    uploadCacheReady,
    ordersFiles,
    incomeFiles,
    refundFiles,
    freightFiles,
    alipayFiles,
    offlineFiles,
    ledgerFiles,
    ordersIdColumn,
    ordersStatusColumn,
    ordersTimeColumn,
    incomeDetailAmountColumns,
    refundDetailAmountColumns,
    refundTypeColumn,
    refundProductNameColumn,
    refundSkuIdColumn,
    freightOrderColumn,
    freightFulfillmentColumn,
    freightWaybillColumn,
    freightAmountCnyColumn,
    freightAmountUsdColumn,
    shopId,
    shopName,
    subsidiary
  ])

  function clearLocalUploadCache() {
    clearUploadCache()
    setOrdersFiles([])
    setIncomeFiles([])
    setRefundFiles([])
    setFreightFiles([])
    setAlipayFiles([])
    setOfflineFiles([])
    setLedgerFiles([])
    setOrdersIdColumn('')
    setOrdersStatusColumn('')
    setOrdersTimeColumn('')
    setIncomeDetailAmountColumns([])
    setRefundDetailAmountColumns([])
    setRefundTypeColumn('')
    setRefundProductNameColumn('')
    setRefundSkuIdColumn('')
    setFreightOrderColumn('')
    setFreightFulfillmentColumn('')
    setFreightWaybillColumn('')
    setFreightAmountCnyColumn('')
    setFreightAmountUsdColumn('')
    setShopId('')
    setShopName('')
    setSubsidiary('')
    setResult(null)
    setLastCalculatedAt('')
    setCalculationCount(0)
    setErrorMessage('已清空本地缓存与当前上传数据。')
  }

  const ordersSheetRows = useMemo(
    () => ordersFiles.flatMap((item) => getSelectedSheet(item.file)?.rows || []),
    [ordersFiles]
  )
  const ordersHeaders = useMemo(
    () => Array.from(new Set(ordersFiles.flatMap((item) => getSelectedSheet(item.file)?.headers || []))),
    [ordersFiles]
  )
  const freightSheetRows = useMemo(
    () => freightFiles.flatMap((item) => getSelectedSheet(item.file)?.rows || []),
    [freightFiles]
  )
  const freightHeaders = useMemo(
    () => Array.from(new Set(freightFiles.flatMap((item) => getSelectedSheet(item.file)?.headers || []))),
    [freightFiles]
  )
  const incomeHeaders = useMemo(() => {
    return Array.from(new Set(incomeFiles.flatMap((item) => getSelectedSheet(item.file)?.headers || [])))
  }, [incomeFiles])
  const refundHeaders = useMemo(() => {
    return Array.from(new Set(refundFiles.flatMap((item) => getSelectedSheet(item.file)?.headers || [])))
  }, [refundFiles])
  const effectiveIncomeDetailAmountColumns = useMemo(() => {
    const valid = incomeDetailAmountColumns.filter(
      (column) => incomeHeaders.includes(column) && !isExcludedIncomeDetailColumn(column)
    )
    if (valid.length > 0) {
      return valid
    }
    return inferIncomeDetailAmountColumns(incomeHeaders)
  }, [incomeDetailAmountColumns, incomeHeaders])
  const effectiveRefundDetailAmountColumns = useMemo(() => {
    const valid = refundDetailAmountColumns.filter(
      (column) => refundHeaders.includes(column) && !isExcludedIncomeDetailColumn(column)
    )
    if (valid.length > 0) {
      return valid
    }
    return inferIncomeDetailAmountColumns(refundHeaders)
  }, [refundDetailAmountColumns, refundHeaders])
  const effectiveRefundTypeColumn = resolveColumnSelection(
    refundTypeColumn,
    refundHeaders,
    DEFAULT_REFUND_TYPE_HINTS
  )
  const effectiveRefundProductNameColumn = resolveColumnSelection(
    refundProductNameColumn,
    refundHeaders,
    DEFAULT_REFUND_PRODUCT_NAME_HINTS
  )
  const effectiveRefundSkuIdColumn = resolveColumnSelection(
    refundSkuIdColumn,
    refundHeaders,
    DEFAULT_REFUND_SKU_ID_HINTS
  )

  const effectiveOrdersStatusColumn = resolveColumnSelection(
    ordersStatusColumn,
    ordersHeaders,
    DEFAULT_ORDER_STATUS_HINTS
  )
  const effectiveOrdersTimeColumn = resolveColumnSelection(
    ordersTimeColumn,
    ordersHeaders,
    DEFAULT_ORDER_TIME_HINTS
  )
  const effectiveFreightFulfillmentColumn = resolveColumnSelection(
    freightFulfillmentColumn,
    freightHeaders,
    DEFAULT_FREIGHT_FULFILLMENT_HINTS
  )
  const effectiveFreightWaybillColumn = resolveColumnSelection(
    freightWaybillColumn,
    freightHeaders,
    DEFAULT_FREIGHT_WAYBILL_HINTS
  )
  const effectiveFreightAmountCnyColumn = resolveColumnSelection(
    freightAmountCnyColumn,
    freightHeaders,
    DEFAULT_FREIGHT_CNY_HINTS
  )
  const effectiveFreightAmountUsdColumn = resolveColumnSelection(
    freightAmountUsdColumn,
    freightHeaders,
    DEFAULT_FREIGHT_USD_HINTS
  )

  async function handleFileUpload(event: React.ChangeEvent<HTMLInputElement>, table: TableKind) {
    const files = Array.from(event.target.files || [])
    if (files.length === 0) {
      return
    }

    setErrorMessage('')

    try {
      for (const file of files) {
        const buffer = await file.arrayBuffer()
        const workbook = XLSX.read(buffer, { type: 'array' })
        const uploaded = toUploadedState(file.name, workbook)
        const defaultSheetHeaders = uploaded.sheets[0]?.headers || []

        if (table === 'orders') {
          setOrdersIdColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_ORDER_ID_HINTS))
          setOrdersStatusColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_ORDER_STATUS_HINTS))
          setOrdersTimeColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_ORDER_TIME_HINTS))
          setOrdersFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded }]
          })
          continue
        }

        if (table === 'income') {
          const orderColumn = pickDefaultColumn(defaultSheetHeaders, DEFAULT_ORDER_ID_HINTS)
          const flowTypeColumn = pickOptionalColumn(defaultSheetHeaders, DEFAULT_INCOME_FLOW_TYPE_HINTS)
          setIncomeDetailAmountColumns((prev) => {
            if (prev.length > 0) {
              return prev
            }
            return inferIncomeDetailAmountColumns(defaultSheetHeaders)
          })
          setIncomeFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded, orderColumn, flowTypeColumn }]
          })
          continue
        }

        if (table === 'refund') {
          const orderColumn = pickDefaultColumn(defaultSheetHeaders, DEFAULT_ORDER_ID_HINTS)
          setRefundTypeColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_REFUND_TYPE_HINTS))
          setRefundProductNameColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_REFUND_PRODUCT_NAME_HINTS))
          setRefundSkuIdColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_REFUND_SKU_ID_HINTS))
          setRefundDetailAmountColumns((prev) => {
            if (prev.length > 0) {
              return prev
            }
            return inferIncomeDetailAmountColumns(defaultSheetHeaders)
          })
          setRefundFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded, orderColumn }]
          })
          continue
        }

        if (table === 'freight') {
          setFreightOrderColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_FREIGHT_ID_HINTS))
          setFreightFulfillmentColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_FREIGHT_FULFILLMENT_HINTS))
          setFreightWaybillColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_FREIGHT_WAYBILL_HINTS))
          setFreightAmountCnyColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_FREIGHT_CNY_HINTS))
          setFreightAmountUsdColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_FREIGHT_USD_HINTS))
          setFreightFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded }]
          })
          continue
        }

        if (table === 'alipay') {
          setAlipayFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded }]
          })
          continue
        }

        if (table === 'offline') {
          setOfflineFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded }]
          })
          continue
        }

        if (table === 'ledger') {
          const mappings = readLedgerMappings()
          const profileKey = getLedgerProfileKey(file.name)
          const saved = mappings[profileKey]
          const orderColumn =
            saved?.orderColumn && defaultSheetHeaders.includes(saved.orderColumn)
              ? saved.orderColumn
              : pickDefaultColumn(defaultSheetHeaders, DEFAULT_ORDER_ID_HINTS)
          const purchaseColumn =
            saved?.purchaseColumn && defaultSheetHeaders.includes(saved.purchaseColumn)
              ? saved.purchaseColumn
              : pickDefaultColumn(defaultSheetHeaders, DEFAULT_LEDGER_PURCHASE_HINTS)
          const freightColumn =
            saved?.freightColumn && defaultSheetHeaders.includes(saved.freightColumn)
              ? saved.freightColumn
              : pickDefaultColumn(defaultSheetHeaders, DEFAULT_LEDGER_FREIGHT_HINTS)

          setLedgerFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded, orderColumn, purchaseColumn, freightColumn }]
          })
          continue
        }
      }

      event.target.value = ''
    } catch (error) {
      console.error(error)
      setErrorMessage(`读取文件失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  function removeOrdersFile(itemId: string) {
    setErrorMessage('')
    setOrdersFiles((prev) => prev.filter((item) => item.id !== itemId))
  }

  function removeFreightFile(itemId: string) {
    setErrorMessage('')
    setFreightFiles((prev) => prev.filter((item) => item.id !== itemId))
  }

  function removeSimpleUploadFile(itemId: string, table: 'alipay' | 'offline') {
    setErrorMessage('')
    if (table === 'alipay') {
      setAlipayFiles((prev) => prev.filter((item) => item.id !== itemId))
      return
    }
    setOfflineFiles((prev) => prev.filter((item) => item.id !== itemId))
  }

  function removeLedgerFile(itemId: string) {
    setErrorMessage('')
    setLedgerFiles((prev) => prev.filter((item) => item.id !== itemId))
  }

  function updateLedgerColumn(itemId: string, field: 'orderColumn' | 'purchaseColumn' | 'freightColumn', value: string) {
    setErrorMessage('')
    setLedgerFiles((prev) => {
      const next = prev.map((item) => (item.id === itemId ? { ...item, [field]: value } : item))
      const changed = next.find((item) => item.id === itemId)
      if (changed) {
        const mappings = readLedgerMappings()
        mappings[getLedgerProfileKey(changed.file.fileName)] = {
          orderColumn: changed.orderColumn,
          purchaseColumn: changed.purchaseColumn,
          freightColumn: changed.freightColumn
        }
        writeLedgerMappings(mappings)
      }
      return next
    })
  }

  function removeIncomeFile(itemId: string, table: 'income' | 'refund' = 'income') {
    setErrorMessage('')
    const setter = table === 'income' ? setIncomeFiles : setRefundFiles
    setter((prev) => prev.filter((item) => item.id !== itemId))
  }

  function buildExternalRecords(
    files: MultiUploadItem[],
    remarkHints: string[],
    amountHints: string[],
    mode: 'alipay' | 'offline'
  ): ExternalRecord[] {
    const records: ExternalRecord[] = []

    files.forEach((item) => {
      const selectedSheet = getSelectedSheet(item.file)
      if (!selectedSheet) {
        return
      }

      const remarkColumn = pickDefaultColumn(selectedSheet.headers, remarkHints)
      const amountColumn = pickDefaultColumn(selectedSheet.headers, amountHints)
      const orderColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ORDER_ID_HINTS)
      const tradeNoColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_TRADE_NO_HINTS)
      const counterpartColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_COUNTERPART_HINTS)
      const productColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_PRODUCT_HINTS)
      const payMethodColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_PAY_METHOD_HINTS)
      const tradeOrderColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_TRADE_ORDER_HINTS)
      const merchantOrderColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_MERCHANT_ORDER_HINTS)

      selectedSheet.rows.forEach((row) => {
        const remark = normalizeCellValue(row[remarkColumn])
        if (!matchesShopRemarkPrefix(remark, shopName)) {
          return
        }

        const orderNo = normalizeOrderNo(row[orderColumn]) || normalizeOrderNo(extractOrderNoFromRemark(remark))
        const amount = toNumericValue(row[amountColumn])
        const hasAnyValue = Object.values(row).some((value) => normalizeCellValue(value))

        if (!hasAnyValue) {
          return
        }
        if (!orderNo || Math.abs(amount) < 0.000001) {
          return
        }

        records.push({
          订单号: orderNo,
          金额: Number(amount.toFixed(2)),
          备注: remark,
          支付宝交易单号: mode === 'alipay' ? normalizeCellValue(row[tradeNoColumn]) : '',
          交易对方: mode === 'alipay' ? normalizeCellValue(row[counterpartColumn]) : '',
          商品说明: mode === 'alipay' ? normalizeCellValue(row[productColumn]) : '',
          收付款方式: mode === 'alipay' ? normalizeCellValue(row[payMethodColumn]) : '',
          交易订单号: mode === 'alipay' ? normalizeCellValue(row[tradeOrderColumn]) : '',
          商家订单号: mode === 'alipay' ? normalizeCellValue(row[merchantOrderColumn]) : '',
          来源文件: item.file.fileName,
          来源Sheet: selectedSheet.name
        })
      })
    })

    return records
  }

  function canProcess(): boolean {
    const incomeContentError = validateDetailFiles(incomeFiles, '订单明细/收支明细')
    const refundContentError = validateDetailFiles(refundFiles, '放退款订单明细')

    return Boolean(
      ordersFiles.length > 0 &&
      incomeFiles.length > 0 &&
      refundFiles.length > 0 &&
      !incomeContentError &&
      !refundContentError &&
      effectiveIncomeDetailAmountColumns.length > 0 &&
      effectiveRefundDetailAmountColumns.length > 0 &&
      ordersIdColumn &&
      effectiveOrdersStatusColumn
    )
  }

  function getProcessDisabledReason(): string {
    if (ordersFiles.length === 0) return '请先上传订单表文件。'
    if (incomeFiles.length === 0) return '请先上传订单明细/收支明细文件。'
    if (refundFiles.length === 0) return '请先上传放退款订单明细文件。'

    if (!ordersIdColumn) return '订单表缺少“订单号字段”，请检查订单表列名。'
    if (!effectiveOrdersStatusColumn) return '订单表缺少“订单状态字段”，请检查订单表列名。'

    if (effectiveIncomeDetailAmountColumns.length === 0) return '订单明细/收支明细未识别到金额字段。'
    if (effectiveRefundDetailAmountColumns.length === 0) return '放退款明细未识别到费用字段。'

    const incomeContentError = validateDetailFiles(incomeFiles, '订单明细/收支明细')
    if (incomeContentError) return incomeContentError
    const refundContentError = validateDetailFiles(refundFiles, '放退款订单明细')
    if (refundContentError) return refundContentError

    return ''
  }

  function runCalculation() {
    if (ordersFiles.length === 0 || incomeFiles.length === 0 || refundFiles.length === 0) {
      setErrorMessage('请先上传订单表、订单明细/收支明细、放退款明细三个区块的表。')
      return
    }

    if (
      !ordersIdColumn ||
      !effectiveOrdersStatusColumn ||
      effectiveIncomeDetailAmountColumns.length === 0 ||
      effectiveRefundDetailAmountColumns.length === 0
    ) {
      setErrorMessage('请先确认订单表、订单明细/收支明细、放退款明细的关键字段。')
      return
    }

    const incomeContentError = validateDetailFiles(incomeFiles, '订单明细/收支明细')
    if (incomeContentError) {
      setErrorMessage(incomeContentError)
      return
    }
    const refundContentError = validateDetailFiles(refundFiles, '放退款订单明细')
    if (refundContentError) {
      setErrorMessage(refundContentError)
      return
    }

    setIsProcessing(true)
    setErrorMessage('')

    try {
      const DEBUG_CALC = true
      const orderRows = dedupeByKey(ordersSheetRows, (row) => {
        const orderNo = normalizeOrderNo(row[ordersIdColumn])
        if (orderNo) {
          return `ORDER:${orderNo}`
        }
        return `RAW:${JSON.stringify(row)}`
      })
      const freightRows = freightSheetRows
      const alipayRecords = buildExternalRecords(alipayFiles, DEFAULT_ALIPAY_REMARK_HINTS, DEFAULT_ALIPAY_AMOUNT_HINTS, 'alipay')
      const offlineRecords = buildExternalRecords(offlineFiles, DEFAULT_OFFLINE_REMARK_HINTS, DEFAULT_OFFLINE_AMOUNT_HINTS, 'offline')
      const ledgerRecords = ledgerFiles.flatMap((item) => {
        const selectedSheet = getSelectedSheet(item.file)
        if (!selectedSheet) {
          return [] as Array<{ 订单号: string; 台账采购金额: number; 台账运费: number }>
        }

        return selectedSheet.rows
          .map((row) => {
            const contextText = Object.values(row).map((value) => normalizeCellValue(value)).join(' ')
            return {
              订单号: normalizeOrderNo(row[item.orderColumn]),
              台账采购金额: normalizeMoney(row[item.purchaseColumn]),
              台账运费: convertFreightToMovement(row[item.freightColumn], contextText)
            }
          })
          .filter((row) => row.订单号)
      })

      const buildDetailEntries = (
        files: IncomeUploadItem[],
        sourceLabel: string,
        selectedColumns: string[],
        typeColumn: string = ''
      ): RowData[] => {
        const rows: RowData[] = []

        files.forEach((item) => {
          const selectedSheet = getSelectedSheet(item.file)
          if (!selectedSheet) {
            return
          }

          const flowTypeColumn = item.flowTypeColumn && selectedSheet.headers.includes(item.flowTypeColumn)
            ? item.flowTypeColumn
            : pickOptionalColumn(selectedSheet.headers, DEFAULT_INCOME_FLOW_TYPE_HINTS)
          const feeItemColumn = pickOptionalColumn(selectedSheet.headers, DEFAULT_INCOME_FEE_ITEM_HINTS)
          const movementColumn =
            pickOptionalColumn(selectedSheet.headers, DEFAULT_INCOME_MOVEMENT_HINTS) ||
            pickOptionalColumn(selectedColumns, DEFAULT_INCOME_MOVEMENT_HINTS)
          const isIncomeMode = !typeColumn

          selectedSheet.rows.forEach((row) => {
            const orderNo = normalizeOrderNo(row[item.orderColumn])
            if (!orderNo) {
              return
            }

            const baseRow: RowData = {
              ...row,
              收支来源文件: item.file.fileName,
              收支来源Sheet: selectedSheet.name,
              收支匹配订单号: orderNo,
              收支匹配字段名: item.orderColumn,
              收支来源区块: sourceLabel
            }

            if (isIncomeMode) {
              const flowTypeValue = normalizeCellValue(flowTypeColumn ? row[flowTypeColumn] : row.收支类型)
              if (!flowTypeValue) {
                return
              }

              if (!movementColumn) {
                return
              }

              const feeItemValue = normalizeCellValue(feeItemColumn ? row[feeItemColumn] : '') || flowTypeValue || movementColumn
              const movementAmount = toIncomeDetailMovementByFlowType(movementColumn, row[movementColumn], flowTypeValue)
              if (Math.abs(movementAmount) < 0.000001) {
                return
              }

              rows.push({
                订单号: orderNo,
                收支类型: flowTypeValue || movementColumn,
                变动金额: movementAmount,
                费用项: feeItemValue,
                币种: 'CNY',
                来源: sourceLabel,
                放退款类型: '',
                收支来源文件: baseRow.收支来源文件,
                收支来源Sheet: baseRow.收支来源Sheet
              })
              return
            }

            selectedColumns.forEach((column) => {
              if (column === normalizeCellValue(baseRow.收支匹配字段名) || isExcludedIncomeDetailColumn(column)) {
                return
              }

              const movementAmount = typeColumn
                ? toRefundDetailMovement(column, baseRow[column], baseRow[typeColumn])
                : toIncomeDetailMovementByFlowType(
                  column,
                  baseRow[column],
                  flowTypeColumn ? baseRow[flowTypeColumn] : baseRow.收支类型
                )
              const refundTypeText = typeColumn ? normalizeCellValue(baseRow[typeColumn]) : ''
              const isCancelRefundType =
                refundTypeText.includes('客户取消') ||
                refundTypeText.includes('买家取消') ||
                refundTypeText.includes('取消订单')

              if (Math.abs(movementAmount) < 0.000001 && !isCancelRefundType) {
                return
              }

              rows.push({
                订单号: orderNo,
                收支类型: column,
                变动金额: movementAmount,
                费用项: column,
                币种: 'CNY',
                来源: sourceLabel,
                放退款类型: typeColumn ? normalizeCellValue(baseRow[typeColumn]) : '',
                收支来源文件: baseRow.收支来源文件,
                收支来源Sheet: baseRow.收支来源Sheet
              })
            })
          })
        })

        return rows
      }

      const incomeDetailRows = buildDetailEntries(incomeFiles, '订单明细表', effectiveIncomeDetailAmountColumns)
      const refundDetailRows = buildDetailEntries(
        refundFiles,
        '放退款订单明细',
        effectiveRefundDetailAmountColumns,
        effectiveRefundTypeColumn
      )
      const incomeEntries = [...incomeDetailRows, ...refundDetailRows]

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
        console.table(incomeDetailRows.slice(0, 20).map((row) => ({
          订单号: normalizeOrderNo(row.订单号),
          收支类型: normalizeCellValue(row.收支类型),
          变动金额: normalizeMoney(row.变动金额),
          来源: normalizeCellValue(row.来源),
          收支来源文件: normalizeCellValue(row.收支来源文件),
          收支来源Sheet: normalizeCellValue(row.收支来源Sheet)
        })))
        console.groupEnd()
      }

      const orderIds = new Set(
        orderRows
          .map((row) => normalizeOrderNo(row[ordersIdColumn]))
          .filter(Boolean)
      )

      const scopedAlipayRecords = alipayRecords.filter((row) => orderIds.has(row.订单号))
      const scopedOfflineRecords = offlineRecords.filter((row) => orderIds.has(row.订单号))
      const scopedLedgerRecords = ledgerRecords.filter((row) => orderIds.has(row.订单号))
      const unmatchedAlipayRows = sortByOrderNo(
        alipayRecords
          .filter((row) => !orderIds.has(row.订单号))
          .map((row) => ({
            订单号: row.订单号,
            支付宝交易单号: row.支付宝交易单号,
            支付宝交易单号和金额: row.支付宝交易单号 ? `${row.支付宝交易单号}/${Number(row.金额.toFixed(2))}` : `${Number(row.金额.toFixed(2))}`,
            交易对方: row.交易对方,
            商品说明: row.商品说明,
            金额: Number(row.金额.toFixed(2)),
            '收/付款方式': row.收付款方式,
            交易订单号: row.交易订单号,
            商家订单号: row.商家订单号,
            备注: row.备注,
            来源文件: row.来源文件,
            来源Sheet: row.来源Sheet,
            店铺过滤规则: normalizeCellValue(shopName) ? `按店铺前缀过滤：${normalizeCellValue(shopName)}` : '未按店铺名前缀过滤'
          })),
        '订单号'
      )

      const scopedIncomeDetailRows = incomeDetailRows.filter((row) => orderIds.has(normalizeOrderNo(row.订单号)))
      const scopedRefundDetailRows = refundDetailRows.filter((row) => orderIds.has(normalizeOrderNo(row.订单号)))
      const refundRawRows = refundFiles.flatMap((item) => {
        const selectedSheet = getSelectedSheet(item.file)
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

      const actualFreightRows = freightRows
        .map((row) => {
          const contextText = Object.values(row).map((value) => normalizeCellValue(value)).join(' ')
          const orderNo = findFreightOrderNoInRow(row, freightOrderColumn, orderIds)
          const cny = convertFreightToMovement(row[effectiveFreightAmountCnyColumn], contextText)
          const usd = convertFreightToMovement(row[effectiveFreightAmountUsdColumn], contextText)
          const fallback = Math.abs(cny) < 0.000001 && Math.abs(usd) < 0.000001
            ? convertFreightToMovement(findFreightFallbackAmount(row), contextText)
            : 0

          return {
            ...row,
            订单号: orderNo,
            计费金额合计CNY_标准化: cny || fallback,
            计费金额合计USD_标准化: usd
          }
        })
        .filter((row) => {
          const id = normalizeOrderNo(row.订单号)
          return id && orderIds.has(id)
        })

      const offlineRowsByOrder = new Map<string, ExternalRecord[]>()
      scopedOfflineRecords.forEach((row) => {
        const current = offlineRowsByOrder.get(row.订单号) || []
        current.push(row)
        offlineRowsByOrder.set(row.订单号, current)
      })

      const alipayRowsByOrder = new Map<string, ExternalRecord[]>()
      scopedAlipayRecords.forEach((row) => {
        const current = alipayRowsByOrder.get(row.订单号) || []
        current.push(row)
        alipayRowsByOrder.set(row.订单号, current)
      })

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
          const movement = convertFreightToMovement(offlineRow.金额, offlineRow.备注)
          performanceRows.push({
            订单号: orderNo,
            订单状态: orderStatus,
            订单时间: orderTime,
            收支类型: '线下运费',
            变动金额: movement,
            费用项: '线下物流费用',
            物流履约单号: '',
            运单号: '',
            币种: 'CNY',
            来源: '线下发货记录',
            放退款核查标记: '',
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
            收支来源文件: alipayRow.来源文件,
            收支来源Sheet: alipayRow.来源Sheet,
            支付宝交易单号: alipayRow.支付宝交易单号,
            支付宝交易单号和金额: alipayRow.支付宝交易单号 ? `${alipayRow.支付宝交易单号}/${alipayRow.金额}` : `${alipayRow.金额}`,
            交易对方: alipayRow.交易对方,
            商品说明: alipayRow.商品说明,
            金额: alipayRow.金额,
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

      if (DEBUG_CALC) {
        console.groupCollapsed('[Calc Debug] 业务明细入聚合前')
        console.log('sortedPerformanceRows count:', sortedPerformanceRows.length)
        console.table(sortedPerformanceRows.slice(0, 30).map((row) => ({
          订单号: normalizeOrderNo(row.订单号),
          来源: normalizeCellValue(row.来源),
          收支类型: normalizeCellValue(row.收支类型),
          变动金额: normalizeMoney(row.变动金额),
          放退款类型: normalizeCellValue(row.放退款类型)
        })))
        console.groupEnd()
      }

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
      const allTypeFeeItemColumns = Array.from(
        new Set(
          sortedPerformanceRows
            .filter((row) => isOrderIncomeSource(normalizeCellValue(row.来源)))
            .map((row) => toTypeFeeItemAmountColumn(normalizeCellValue(row.收支类型), normalizeCellValue(row.费用项)))
            .filter(Boolean)
        )
      )
      const allLogisticsSubItemColumns = Array.from(
        new Set(
          sortedPerformanceRows.flatMap((row) => {
            const source = normalizeCellValue(row.来源)
            const type = normalizeCellValue(row.收支类型)
            const feeItem = normalizeCellValue(row.费用项) || type || '未分类'

            if (isOrderIncomeSource(source) && type.includes('支出>物流费用')) {
              return [toLogisticsSubItemAmountColumn('收支表', feeItem)]
            }
            if (source === '运费表' || type === '物流运费') {
              return [toLogisticsSubItemAmountColumn('金掌柜', feeItem)]
            }
            return []
          })
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
        const typeColumn = toTypeAmountColumn(normalizeCellValue(row.收支类型))
        const feeItemColumn = toFeeItemAmountColumn(feeItem)
        const typeFeeItemColumn = toTypeFeeItemAmountColumn(type, feeItem)
        const current = aggregatedMap.get(orderNo) || {
          订单号: orderNo,
          订单状态: normalizeCellValue(row.订单状态),
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
          订单明细_平台分账金额合计: 0,
          收入_按订单明细: 0,
          预计可得_按收支及运费: 0,
          预计可得_按退放款及运费: 0,
          预计可得差异_收支减退放款: 0,
          预计可得校验状态: '一致',
          放退款_金额项合计: 0,
          放退款_其他费用合计: 0,
          收入_按放退款明细: 0,
          收入差异_订单明细减放退款: 0,
          收入校验状态: '一致',
          线上运费: 0,
          线下运费: 0,
          台账采购金额: 0,
          台账运费: 0,
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
          current.订单明细_净收支合计 = normalizeMoney(toNumericValue(current.订单明细_净收支合计) + amount)

          if (amount >= 0) {
            current.收入_收支明细表 = normalizeMoney(toNumericValue(current.收入_收支明细表) + amount)
          } else {
            current.支出_收支明细表 = normalizeMoney(toNumericValue(current.支出_收支明细表) + Math.abs(amount))
          }

          if (type.includes('支出>物流费用')) {
            const logisticsExpense = amount < 0 ? Math.abs(amount) : 0
            current.收支表_支出物流费用 = normalizeMoney(
              toNumericValue(current.收支表_支出物流费用) + logisticsExpense
            )
            current.物流支出_收支表 = normalizeMoney(toNumericValue(current.物流支出_收支表) + logisticsExpense)

            const logisticsSubItemColumn = toLogisticsSubItemAmountColumn('收支表', feeItem)
            current[logisticsSubItemColumn] = normalizeMoney(
              toNumericValue(current[logisticsSubItemColumn]) + logisticsExpense
            )
          }

          if (type.includes('待结算金额')) {
            current.待结算金额合计 = normalizeMoney(toNumericValue(current.待结算金额合计) + Math.abs(amount))
          }
          if (type.includes('放款金额')) {
            current.订单明细_放款金额合计 = normalizeMoney(toNumericValue(current.订单明细_放款金额合计) + amount)
          }
          if (type.includes('平台分账金额')) {
            current.订单明细_平台分账金额合计 = normalizeMoney(
              toNumericValue(current.订单明细_平台分账金额合计) + Math.abs(amount)
            )
          }
        }

        if (source === '放退款订单明细') {
          if (hasKeyword(type, REFUND_BASE_AMOUNT_HINTS)) {
            current.放退款_金额项合计 = normalizeMoney(toNumericValue(current.放退款_金额项合计) + amount)
          } else {
            current.放退款_其他费用合计 = normalizeMoney(
              toNumericValue(current.放退款_其他费用合计) + Math.abs(amount)
            )
          }
        }

        if (source === '运费表' || type === '物流运费') {
          current.线上运费 = normalizeMoney(toNumericValue(current.线上运费) + amount)
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
        }

        if (source === '支付宝订单记录' || type === '采购支出') {
          current.线下运费 = normalizeMoney(toNumericValue(current.线下运费) + amount)
        }

        current[typeColumn] = normalizeMoney(toNumericValue(current[typeColumn]) + amount)
        if (isOrderIncomeSource(source)) {
          current[feeItemColumn] = normalizeMoney(toNumericValue(current[feeItemColumn]) + amount)
          current[typeFeeItemColumn] = normalizeMoney(toNumericValue(current[typeFeeItemColumn]) + amount)
        }

        const legacyOrderDetailIncome = normalizeMoney(
          toNumericValue(current.订单明细_放款金额合计) - toNumericValue(current.订单明细_平台分账金额合计)
        )
        const hasLegacyOrderDetailColumns =
          Math.abs(toNumericValue(current.订单明细_放款金额合计)) > 0.000001 ||
          Math.abs(toNumericValue(current.订单明细_平台分账金额合计)) > 0.000001
        const incomeFromOrderDetail = hasLegacyOrderDetailColumns
          ? legacyOrderDetailIncome
          : normalizeMoney(current.订单明细_净收支合计)
        const incomeFromRefundDetail = normalizeMoney(
          toNumericValue(current.放退款_金额项合计) - toNumericValue(current.放退款_其他费用合计)
        )
        const incomeFromOrderDetailExcludingLogistics = normalizeMoney(
          incomeFromOrderDetail + toNumericValue(current.收支表_支出物流费用)
        )
        const incomeDiff = normalizeMoney(incomeFromOrderDetailExcludingLogistics - incomeFromRefundDetail)
        const totalFreight = normalizeMoney(toNumericValue(current.线上运费) + toNumericValue(current.线下运费))
        const incomeLogisticsExpense = normalizeMoney(current.收支表_支出物流费用)
        const jzgFreightExpense = normalizeMoney(current.金掌柜物流费支出)
        const logisticsExpenseTotal = normalizeMoney(incomeLogisticsExpense + jzgFreightExpense)
        const expectedFromIncomeAndFreight = normalizeMoney(
          incomeFromOrderDetailExcludingLogistics - logisticsExpenseTotal
        )
        const expectedFromRefundAndFreight = normalizeMoney(incomeFromRefundDetail - incomeLogisticsExpense - jzgFreightExpense)
        const expectedIncomeDiff = normalizeMoney(expectedFromIncomeAndFreight - expectedFromRefundAndFreight)

        current.收入_按订单明细 = incomeFromOrderDetail
        current.收入_按放退款明细 = incomeFromRefundDetail
        current.收支总和_不含物流费用 = incomeFromOrderDetailExcludingLogistics
        current.差异_收支不含物流减退放款 = incomeDiff
        current.收入差异_订单明细减放退款 = incomeDiff
        current.收入校验状态 = Math.abs(incomeDiff) <= 0.01 ? '一致' : '不一致'
        current.预计可得_按收支及运费 = expectedFromIncomeAndFreight
        current.预计可得_按退放款及运费 = expectedFromRefundAndFreight
        current.预计可得差异_收支减退放款 = expectedIncomeDiff
        current.预计可得校验状态 = Math.abs(expectedIncomeDiff) <= 0.01 ? '一致' : '不一致'
        current.物流支出总和 = logisticsExpenseTotal
        current.最终收入_未扣运费 = incomeFromOrderDetailExcludingLogistics
        current.最终收入_扣运费 = expectedFromIncomeAndFreight

        current.收入合计 = incomeFromOrderDetail
        current.支出合计 = normalizeMoney(Math.max(totalFreight, 0))
        current.总收支 = expectedFromIncomeAndFreight

        aggregatedMap.set(orderNo, current)
      })

      scopedLedgerRecords.forEach((row) => {
        const current = aggregatedMap.get(row.订单号)
        if (!current) {
          return
        }

        current.台账采购金额 = normalizeMoney(toNumericValue(current.台账采购金额) + row.台账采购金额)
        current.台账运费 = normalizeMoney(toNumericValue(current.台账运费) + row.台账运费)
        aggregatedMap.set(row.订单号, current)
      })

      if (DEBUG_CALC) {
        const aggregatedRawRows = Array.from(aggregatedMap.values())
        const zeroIncomeRows = aggregatedRawRows.filter(
          (row) => Math.abs(toNumericValue(row.收入_按订单明细)) < 0.000001
        )

        console.groupCollapsed('[Calc Debug] 聚合后关键字段（排查为0）')
        console.log('aggregatedRawRows count:', aggregatedRawRows.length)
        console.log('收入_按订单明细为0的订单数:', zeroIncomeRows.length)
        console.table(aggregatedRawRows.slice(0, 30).map((row) => ({
          订单号: normalizeCellValue(row.订单号),
          订单明细_净收支合计: normalizeMoney(row.订单明细_净收支合计),
          订单明细_放款金额合计: normalizeMoney(row.订单明细_放款金额合计),
          订单明细_平台分账金额合计: normalizeMoney(row.订单明细_平台分账金额合计),
          收入_按订单明细: normalizeMoney(row.收入_按订单明细),
          收入_按放退款明细: normalizeMoney(row.收入_按放退款明细),
          收入差异_订单明细减放退款: normalizeMoney(row.收入差异_订单明细减放退款)
        })))
        console.table(zeroIncomeRows.slice(0, 30).map((row) => ({
          订单号: normalizeCellValue(row.订单号),
          订单明细_净收支合计: normalizeMoney(row.订单明细_净收支合计),
          订单明细_放款金额合计: normalizeMoney(row.订单明细_放款金额合计),
          订单明细_平台分账金额合计: normalizeMoney(row.订单明细_平台分账金额合计),
          收入_按订单明细: normalizeMoney(row.收入_按订单明细),
          线上运费: normalizeMoney(row.线上运费),
          线下运费: normalizeMoney(row.线下运费)
        })))
        console.groupEnd()
      }

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
          const systemFreight = normalizeMoney(toNumericValue(next.线上运费) + toNumericValue(next.线下运费))
          const ledgerFreight = normalizeMoney(next.台账运费)
          const freightGap = normalizeMoney(ledgerFreight - systemFreight)
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
          allTypeFeeItemColumns.forEach((col) => {
            if (next[col] === undefined) {
              next[col] = 0
            }
          })
          allLogisticsSubItemColumns.forEach((col) => {
            if (next[col] === undefined) {
              next[col] = 0
            }
          })

          return {
            订单号: normalizeOrderNo(next.订单号),
            订单时间: normalizeCellValue(next.订单时间),
            订单状态: normalizeCellValue(next.订单状态),
            订单预计可得: normalizeMoney(next.最终收入_未扣运费),
            采购金额: 0,
            物流费用: normalizeMoney(next.收支表_支出物流费用),
            金掌柜物流费支出: normalizeMoney(next.金掌柜物流费支出),
            线下物流费用: 0,
            放退款类型: normalizeCellValue(next.放退款类型),
            放退款核查标记: normalizeCellValue(next.放退款核查标记),
            待结算金额合计: pendingSettlement,
            备注: mergedRemark,
            ...Object.fromEntries(allTypeFeeItemColumns.map((col) => [col, normalizeMoney(next[col])])),
            ...Object.fromEntries(allTypeColumns.map((col) => [col, normalizeMoney(next[col])])),
            ...Object.fromEntries(allFeeItemColumns.map((col) => [col, normalizeMoney(next[col])])),
            ...Object.fromEntries(allLogisticsSubItemColumns.map((col) => [col, normalizeMoney(next[col])])),
            订单预计可得_扣物流后: normalizeMoney(next.预计可得_按收支及运费),
            线上运费: normalizeMoney(next.线上运费),
            线下运费: normalizeMoney(next.线下运费),
            物流支出总和: normalizeMoney(next.物流支出总和),
            物流支出_收支表: normalizeMoney(next.物流支出_收支表),
            物流支出_金掌柜: normalizeMoney(next.物流支出_金掌柜),
            台账采购金额: normalizeMoney(next.台账采购金额),
            台账运费: ledgerFreight,
            运费差异_台账与系统: freightGap,
            最终收入_扣运费: normalizeMoney(next.最终收入_扣运费),
            收入_按订单明细: normalizeMoney(next.收入_按订单明细),
            收入_按放退款明细: normalizeMoney(next.收入_按放退款明细),
            收支总和_不含物流费用: normalizeMoney(next.收支总和_不含物流费用),
            差异_收支不含物流减退放款: normalizeMoney(next.差异_收支不含物流减退放款),
            收入差异_订单明细减放退款: normalizeMoney(next.收入差异_订单明细减放退款),
            收入校验状态: normalizeCellValue(next.收入校验状态),
            预计可得_按收支及运费: normalizeMoney(next.预计可得_按收支及运费),
            预计可得_按退放款及运费: normalizeMoney(next.预计可得_按退放款及运费),
            预计可得差异_收支减退放款: normalizeMoney(next.预计可得差异_收支减退放款),
            预计可得校验状态: normalizeCellValue(next.预计可得校验状态),
            收支表_支出物流费用: normalizeMoney(next.收支表_支出物流费用),
            收入_收支明细表: normalizeMoney(next.收入_收支明细表),
            支出_收支明细表: normalizeMoney(next.支出_收支明细表),
            订单明细_放款金额合计: normalizeMoney(next.订单明细_放款金额合计),
            订单明细_平台分账金额合计: normalizeMoney(next.订单明细_平台分账金额合计),
            放退款_金额项合计: normalizeMoney(next.放退款_金额项合计),
            放退款_其他费用合计: normalizeMoney(next.放退款_其他费用合计),
            收入合计: normalizeMoney(next.收入合计),
            支出合计: normalizeMoney(next.支出合计),
            总收支: normalizeMoney(next.总收支)
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
          收入校验状态: normalizeCellValue(row.收入校验状态),
          放退款核查标记: normalizeCellValue(row.放退款核查标记),
          线上运费: normalizeMoney(row.线上运费),
          订单预计可得: normalizeMoney(row.订单预计可得),
          最终收入_扣运费: normalizeMoney(row.最终收入_扣运费)
        })),
        '订单号'
      )

      const ordersWithoutFreightRows = sortByOrderNo(
        aggregatedRows
          .filter((row) => Math.abs(toNumericValue(row.金掌柜物流费支出)) < 0.000001)
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
          const logistics = normalizeMoney(toNumericValue(row.线上运费) + toNumericValue(row.线下运费))
          const finalIncome = toNumericValue(row.最终收入_扣运费)
          const incomeBeforeFreight = toNumericValue(row.订单预计可得)
          const mismatch = normalizeCellValue(row.收入校验状态) === '不一致'
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
          acc.totalExpense += logistics
          acc.logisticsAmount += logistics
          acc.nonLogisticsAmount += orderIncome
          acc.totalAmount += toNumericValue(row.总收支)
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
          totalIncomeAfterOnlineFreight: 0
        } as ResultSummary
      )

      summary.totalAmount = Number(summary.totalAmount.toFixed(2))
      summary.totalIncome = Number(summary.totalIncome.toFixed(2))
      summary.totalExpense = Number(summary.totalExpense.toFixed(2))
      summary.logisticsAmount = Number(summary.logisticsAmount.toFixed(2))
      summary.nonLogisticsAmount = Number(summary.nonLogisticsAmount.toFixed(2))
      summary.totalIncomeBeforeFreight = Number(summary.totalIncomeBeforeFreight.toFixed(2))
      summary.totalIncomeAfterOnlineFreight = Number(summary.totalIncomeAfterOnlineFreight.toFixed(2))

      setResult({
        performanceRows: sortedPerformanceRows,
        aggregatedRows,
        dynamicTypeColumns: allTypeColumns,
        dynamicFeeItemColumns: allFeeItemColumns,
        dynamicTypeFeeItemColumns: allTypeFeeItemColumns,
        dynamicLogisticsSubItemColumns: allLogisticsSubItemColumns,
        incomeValidationRows,
        ordersWithoutIncomeRows,
        incomeOnlyOrdersRows,
        refundOnlyOrdersRows,
        ordersWithoutFreightRows,
        unmatchedAlipayRows,
        summary,
        actualIncomeRows: sortedActualIncomeRows,
        actualFreightRows: sortedActualFreightRows,
        markedOrderRows: sortedMarkedOrderRows,
        integratedSummaryRows,
        integratedDetailRows: sortedIntegratedDetailRows
      })
      setLastCalculatedAt(formatDateTime(new Date()))
      setCalculationCount((prev) => prev + 1)
    } catch (error) {
      console.error(error)
      setErrorMessage(`计算失败：${error instanceof Error ? error.message : '未知错误'}`)
    } finally {
      setIsProcessing(false)
    }
  }

  function getExportFileName(reportName: string, dateText: string): string {
    const companyPart = safeFileSegment(subsidiary, '未分公司')
    const storePart = safeFileSegment(shopName || shopId, '未命名店铺')
    return `${companyPart}_${storePart}_${reportName}_${dateText}.xlsx`
  }

  function exportAggregatedWorkbook() {
    if (!result) {
      setErrorMessage('请先执行计算，再导出结果。')
      return
    }

    const workbook = XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.aggregatedRows),
      '订单聚合表'
    )

    const date = new Date()
    const dateText = `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`
    XLSX.writeFile(workbook, getExportFileName('订单聚合表', dateText))
  }

  function exportBusinessDetailWorkbook() {
    if (!result) {
      setErrorMessage('请先执行计算，再导出结果。')
      return
    }

    const workbook = XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.performanceRows),
      '业务明细表'
    )

    const date = new Date()
    const dateText = `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`
    XLSX.writeFile(workbook, getExportFileName('业务明细表', dateText))
  }

  function exportOtherSheetsWorkbook() {
    if (!result) {
      setErrorMessage('请先执行计算，再导出结果。')
      return
    }

    const workbook = XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.ordersWithoutIncomeRows),
      '订单表有但收支缺失'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.incomeOnlyOrdersRows),
      '仅订单明细有'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.refundOnlyOrdersRows),
      '仅放退款明细有'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.ordersWithoutFreightRows),
      '无线上线下运费订单'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.unmatchedAlipayRows),
      '支付宝无对应订单'
    )

    const date = new Date()
    const dateText = `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`
    XLSX.writeFile(workbook, getExportFileName('其他对账表', dateText))
  }

  function renderOrdersUploadCard() {
    return (
      <section className="upload-card">
        <h2 className="upload-title-row">
          <span>1) 订单表上传</span>
          <span className="upload-count-pill">已上传 {ordersFiles.length} 份</span>
        </h2>
        <p>选择结算月份的所有订单导出。</p>

        <label className="file-input-label">
          <span>批量上传订单表（可多选）</span>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            multiple
            onChange={(event) => void handleFileUpload(event, 'orders')}
          />
        </label>

        <div className="uploaded-files-list">
          {ordersFiles.length === 0 && (
            <div className="meta-row">
              <span className="meta-key">状态</span>
              <span className="meta-value">尚未上传订单文件</span>
            </div>
          )}

          {ordersFiles.map((item) => (
            <article key={item.id} className="income-file-item">
              <div className="income-file-head">
                <div className="file-name-row">
                  <strong>{item.file.fileName}</strong>
                </div>
                <button type="button" className="mini-danger" onClick={() => removeOrdersFile(item.id)}>
                  删除
                </button>
              </div>
            </article>
          ))}
        </div>
      </section>
    )
  }

  function renderDetailUploadCard(
    title: string,
    description: ReactNode,
    table: 'income' | 'refund',
    files: IncomeUploadItem[],
    emptyText: string
  ) {
    return (
      <section className="upload-card">
        <h2 className="upload-title-row">
          <span>{title}</span>
          <span className="upload-count-pill">已上传 {files.length} 份</span>
        </h2>
        <p>{description}</p>

        <label className="file-input-label">
          <span>批量上传（可多选）</span>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            multiple
            onChange={(event) => void handleFileUpload(event, table)}
          />
        </label>

        <div className="uploaded-files-list">
          {files.length === 0 && (
            <div className="meta-row">
              <span className="meta-key">状态</span>
              <span className="meta-value">{emptyText}</span>
            </div>
          )}

          {files.map((item) => (
            <article key={item.id} className="income-file-item">
              <div className="income-file-head">
                <div className="file-name-row">
                  <strong>{item.file.fileName}</strong>
                </div>
                <button type="button" className="mini-danger" onClick={() => removeIncomeFile(item.id, table)}>
                  删除
                </button>
              </div>
            </article>
          ))}
        </div>
      </section>
    )
  }

  function renderFreightUploadCard() {
    return (
      <section className="upload-card">
        <h2 className="upload-title-row">
          <span>2) 金掌柜运费表</span>
          <span className="upload-count-pill">已上传 {freightFiles.length} 份</span>
        </h2>
        <p>一次可上传多个运费文件，系统按交易单号字段匹配订单号并汇总。</p>

        <label className="file-input-label">
          <span>批量上传金掌柜运费表（可多选）</span>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            multiple
            onChange={(event) => void handleFileUpload(event, 'freight')}
          />
        </label>

        <div className="uploaded-files-list">
          {freightFiles.length === 0 && (
            <div className="meta-row">
              <span className="meta-key">状态</span>
              <span className="meta-value">尚未上传运费文件</span>
            </div>
          )}

          {freightFiles.map((item) => (
            <article key={item.id} className="income-file-item">
              <div className="income-file-head">
                <div className="file-name-row">
                  <strong>{item.file.fileName}</strong>
                </div>
                <button type="button" className="mini-danger" onClick={() => removeFreightFile(item.id)}>
                  删除
                </button>
              </div>
            </article>
          ))}
        </div>
      </section>
    )
  }

  function renderSimpleUploadCard(
    title: string,
    description: string,
    table: 'alipay' | 'offline',
    files: MultiUploadItem[],
    emptyText: string
  ) {
    return (
      <section className="upload-card">
        <h2 className="upload-title-row">
          <span>{title}</span>
          <span className="upload-count-pill">已上传 {files.length} 份</span>
        </h2>
        <p>{description}</p>

        <label className="file-input-label">
          <span>批量上传（可多选）</span>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            multiple
            onChange={(event) => void handleFileUpload(event, table)}
          />
        </label>

        <div className="uploaded-files-list">
          {files.length === 0 && (
            <div className="meta-row">
              <span className="meta-key">状态</span>
              <span className="meta-value">{emptyText}</span>
            </div>
          )}

          {files.map((item) => (
            <article key={item.id} className="income-file-item">
              <div className="income-file-head">
                <div className="file-name-row">
                  <strong>{item.file.fileName}</strong>
                </div>
                <button type="button" className="mini-danger" onClick={() => removeSimpleUploadFile(item.id, table)}>
                  删除
                </button>
              </div>
            </article>
          ))}
        </div>
      </section>
    )
  }

  function renderLedgerUploadCard() {
    return (
      <section className="upload-card">
        <h2 className="upload-title-row">
          <span>7) 业务员台账</span>
          <span className="upload-count-pill">已上传 {ledgerFiles.length} 份</span>
        </h2>
        <p>支持多文件。可分别指定订单号/采购金额/台账运费列名，系统会按业务员文件自动记忆对应关系。</p>

        <label className="file-input-label">
          <span>批量上传台账（可多选）</span>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            multiple
            onChange={(event) => void handleFileUpload(event, 'ledger')}
          />
        </label>

        <div className="uploaded-files-list">
          {ledgerFiles.length === 0 && (
            <div className="meta-row">
              <span className="meta-key">状态</span>
              <span className="meta-value">尚未上传业务员台账</span>
            </div>
          )}

          {ledgerFiles.map((item) => {
            const sheet = getSelectedSheet(item.file)
            const headers = sheet?.headers || []
            return (
              <article key={item.id} className="income-file-item ledger-file-item">
                <div className="income-file-head">
                  <div className="file-name-row">
                    <strong>{item.file.fileName}</strong>
                  </div>
                  <button type="button" className="mini-danger" onClick={() => removeLedgerFile(item.id)}>
                    删除
                  </button>
                </div>

                <div className="ledger-map-grid">
                  <div className="control-row">
                    <label>订单号列</label>
                    <select className="compact-select" value={item.orderColumn} onChange={(event) => updateLedgerColumn(item.id, 'orderColumn', event.target.value)}>
                      {headers.map((header) => (
                        <option key={header} value={header}>{header}</option>
                      ))}
                    </select>
                  </div>
                  <div className="control-row">
                    <label>采购金额列</label>
                    <select className="compact-select" value={item.purchaseColumn} onChange={(event) => updateLedgerColumn(item.id, 'purchaseColumn', event.target.value)}>
                      {headers.map((header) => (
                        <option key={header} value={header}>{header}</option>
                      ))}
                    </select>
                  </div>
                  <div className="control-row">
                    <label>台账运费列</label>
                    <select className="compact-select" value={item.freightColumn} onChange={(event) => updateLedgerColumn(item.id, 'freightColumn', event.target.value)}>
                      {headers.map((header) => (
                        <option key={header} value={header}>{header}</option>
                      ))}
                    </select>
                  </div>
                </div>
              </article>
            )
          })}
        </div>
      </section>
    )
  }

  return (
    <main className="page">
      <header className="page-header">
        <span className="badge">Cross-border Ops</span>
        <h1>业绩计算工作台</h1>
        <p>
          上传订单表、订单明细/收支明细表、放退款订单明细、线上运费、支付宝记录、线下发货记录后，系统会自动筛选并生成业绩与对账结果。
        </p>
        <div className="meta-row">
          <span className="meta-key">本地缓存</span>
          <span className="meta-value">已开启（上传后自动保存，刷新页面自动恢复）</span>
        </div>
        <div className="actions-row">
          <button type="button" className="ghost" onClick={clearLocalUploadCache}>
            清空本地缓存
          </button>
        </div>
      </header>

      <section className="upload-card store-meta-card">
        <h2>店铺信息</h2>
        <p>用于导出文件命名：所属分公司_店铺名_xxx表。</p>
        <div className="store-meta-grid">
          <div className="control-row">
            <label>店铺ID</label>
            <input
              className="compact-input"
              value={shopId}
              onChange={(event) => setShopId(event.target.value)}
              placeholder="例如：A12345"
            />
          </div>
          <div className="control-row">
            <label>店铺名</label>
            <input
              className="compact-input"
              value={shopName}
              onChange={(event) => setShopName(event.target.value)}
              placeholder="例如：DXM官方店"
            />
          </div>
          <div className="control-row">
            <label>所属分公司（可选）</label>
            <select
              className="compact-select"
              value={subsidiary}
              onChange={(event) => setSubsidiary(event.target.value)}
            >
              <option value="">未选择</option>
              <option value="华南分公司">华南分公司</option>
              <option value="华东分公司">华东分公司</option>
              <option value="华北分公司">华北分公司</option>
              <option value="华中分公司">华中分公司</option>
              <option value="西南分公司">西南分公司</option>
              <option value="海外事业部">海外事业部</option>
            </select>
          </div>
        </div>
      </section>

      <section className="upload-sections">
        {renderOrdersUploadCard()}
        {renderFreightUploadCard()}
        {renderDetailUploadCard(
          '3) 订单明细/收支明细表',
          <>
            请上传
            <span className="emphasis-russia">【俄罗斯】结算月至今的订单明细表；【非俄罗斯】结算月至今的订单收支明细表。</span>
          </>,
          'income',
          incomeFiles,
          '尚未上传订单明细/收支明细文件'
        )}
        {renderDetailUploadCard(
          '4) 放退款订单明细',
          <>
            请上传
            <span className="emphasis-russia">【俄罗斯】和【非俄罗斯】结算月至今</span>
            的放退款明细文件。
          </>,
          'refund',
          refundFiles,
          '尚未上传放退款订单明细文件'
        )}
        {renderSimpleUploadCard(
          '5) 支付宝订单记录（采购）',
          '支持多文件上传。备注按“店铺名-订单号”匹配，系统会自动摘取短线后的订单号。',
          'alipay',
          alipayFiles,
          '尚未上传支付宝订单记录'
        )}
        {renderSimpleUploadCard(
          '6) 线下发货订单记录',
          '支持多文件上传。备注按“店铺名-订单号”匹配，系统会自动摘取短线后的订单号。',
          'offline',
          offlineFiles,
          '尚未上传线下发货记录'
        )}
        {renderLedgerUploadCard()}
      </section>

      <section className="action-panel">
        <button type="button" onClick={runCalculation} disabled={!canProcess() || isProcessing}>
          {isProcessing ? '计算中...' : '计算业绩'}
        </button>
        <button type="button" className="ghost" onClick={exportAggregatedWorkbook} disabled={!result}>
          导出订单聚合表
        </button>
        <button type="button" className="ghost" onClick={exportBusinessDetailWorkbook} disabled={!result}>
          导出业务明细表
        </button>
        <button type="button" className="ghost" onClick={exportOtherSheetsWorkbook} disabled={!result}>
          导出其他对账Sheet
        </button>
        {result && lastCalculatedAt && (
          <div className="meta-row">
            <span className="meta-key">最近重新生成</span>
            <span className="meta-value">第 {calculationCount} 次（{lastCalculatedAt}）</span>
          </div>
        )}
        {!isProcessing && !canProcess() && (
          <div className="meta-row">
            <span className="meta-key">当前不可执行原因</span>
            <span className="meta-value">{getProcessDisabledReason()}</span>
          </div>
        )}
      </section>

      {errorMessage && <div className="error-box">{errorMessage}</div>}

      {result && (
        <section className="result-panel">
          <h2>计算结果预览</h2>

          <h3 className="overview-title">订单相关</h3>
          <div className="stats-grid">
            <article>
              <h3>订单总数</h3>
              <p>{result.summary.orderCount}</p>
            </article>
            <article>
              <h3>放款订单数</h3>
              <p>{result.summary.payoutOrderCount}</p>
            </article>
            <article>
              <h3>取消订单退款数</h3>
              <p>{result.summary.cancelRefundOrderCount}</p>
            </article>
            <article>
              <h3>纠纷订单数</h3>
              <p>{result.summary.disputeOrderCount}</p>
            </article>
            <article>
              <h3>其他</h3>
              <p>{result.summary.otherOrderCount}</p>
            </article>
          </div>

          <h3 className="overview-title">收支相关</h3>
          <div className="stats-grid">
            <article>
              <h3>订单预计可得</h3>
              <p>{result.summary.totalIncomeBeforeFreight}</p>
            </article>
            <article>
              <h3>所有收入_扣线上运费</h3>
              <p>{result.summary.totalIncomeAfterOnlineFreight}</p>
            </article>
          </div>

          <h3 className="table-title">订单聚合表（页面预览）</h3>
          {result.aggregatedRows.length > 0 && (
            <>
              <div className="meta-row">
                <span className="meta-key">已识别收支类型列</span>
                <span className="meta-value">
                  {result.dynamicTypeColumns.join('、') || '无'}
                </span>
              </div>
              <div className="meta-row">
                <span className="meta-key">已识别费用项列</span>
                <span className="meta-value">
                  {result.dynamicFeeItemColumns.join('、') || '无'}
                </span>
              </div>
              <div className="meta-row">
                <span className="meta-key">已识别收支类型+费用项列</span>
                <span className="meta-value">
                  {result.dynamicTypeFeeItemColumns.join('、') || '无'}
                </span>
              </div>
              <div className="meta-row">
                <span className="meta-key">组合列示例(前10)</span>
                <span className="meta-value">
                  {result.dynamicTypeFeeItemColumns.slice(0, 10).join('、') || '无'}
                </span>
              </div>
              <div className="meta-row">
                <span className="meta-key">物流子项扣费列</span>
                <span className="meta-value">
                  {result.dynamicLogisticsSubItemColumns.join('、') || '无'}
                </span>
              </div>
            </>
          )}
          <div className="detail-table-wrap">
            <table className="detail-table">
              <thead>
                <tr>
                  {Object.keys(result.aggregatedRows[0] || {}).map((col) => (
                    <th key={col}>{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {result.aggregatedRows.map((row, index) => (
                  <tr key={`${normalizeOrderNo(row.订单号)}_${index}`}>
                    {Object.keys(result.aggregatedRows[0] || {}).map((col) => (
                      <td key={col}>{normalizeCellValue(row[col])}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <p className="tip">
            页面仅展示订单聚合表相关结果；导出保持 3 个按钮：订单聚合表、业务明细表、其他对账Sheet。
          </p>
        </section>
      )}
    </main>
  )
}

export default App

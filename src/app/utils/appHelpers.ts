import type * as Domain from '../calculatorDomain'
import {
  DEFAULT_ALIPAY_INVOICE_HINTS,
  DEFAULT_ALIPAY_COUNTERPART_HINTS,
  DEFAULT_OFFLINE_CATEGORY_HINTS,
  DEFAULT_OFFLINE_AMOUNT_HINTS,
  DEFAULT_OFFLINE_ORDER_HINTS,
  DEFAULT_ALIPAY_MERCHANT_ORDER_HINTS,
  DEFAULT_ALIPAY_PAY_METHOD_HINTS,
  DEFAULT_ALIPAY_PRODUCT_HINTS,
  DEFAULT_ALIPAY_TRADE_NO_HINTS,
  DEFAULT_ALIPAY_TRADE_ORDER_HINTS,
  extractOrderNoFromRemark,
  normalizeCellValue,
  normalizeOrderNo,
  pickDefaultColumn,
  safeFileSegment,
  toNumericValue,
  validateDetailFiles,
  getSelectedSheet
} from '../calculatorDomain'

type MultiUploadItem = Domain.MultiUploadItem
type IncomeUploadItem = Domain.IncomeUploadItem
type ExternalRecord = Domain.ExternalRecord
type RowData = Domain.RowData

type ExternalRecordBuildInput = {
  files: MultiUploadItem[]
  remarkHints: string[]
  amountHints: string[]
  mode: 'alipay' | 'offline'
  shopName: string
  useParsedRemarkOrderNoOnly?: boolean
}

type OfflineFreightRecord = {
  订单号: string
  金额: number
  费用类别: string
  客户订单号: string
  备注: string
  来源文件: string
  来源Sheet: string
}

export type OfflineFreightDiagnosticRow = {
  来源文件: string
  来源Sheet: string
  行号: number
  客户订单号: string
  提取订单号: string
  费用类别: string
  金额原值: string
  金额标准化: number
  状态: '纳入计算' | '已跳过'
  原因: string
}

type ProcessValidationInput = {
  ordersFilesLength: number
  incomeFiles: IncomeUploadItem[]
  refundFiles: IncomeUploadItem[]
  effectiveIncomeDetailAmountColumnsLength: number
  effectiveRefundDetailAmountColumnsLength: number
  ordersIdColumn: string
  effectiveOrdersStatusColumn: string
}

function extractShopNameFromRemark(remark: string): string {
  const text = normalizeCellValue(remark)
  if (!text) {
    return ''
  }

  const pureOrder = normalizeOrderNo(text)
  if (pureOrder && pureOrder === text) {
    return ''
  }

  const parts = text
    .split(/[-—_]/)
    .map((item) => normalizeCellValue(item))
    .filter(Boolean)

  if (parts.length >= 2) {
    return parts[0]
  }

  return ''
}

function extractOrderNoFromCustomerOrder(value: string): string {
  const text = normalizeCellValue(value)
  if (!text) {
    return ''
  }

  // Prefer exact 16-digit chunks embedded in mixed text, e.g. aaa-1234759299222333J.
  const exact16DigitCandidates = Array.from(text.matchAll(/(?:^|\D)(\d{16})(?=\D|$)/g)).map((match) => match[1])
  if (exact16DigitCandidates.length > 0) {
    return exact16DigitCandidates[exact16DigitCandidates.length - 1]
  }

  const normalizedWhole = normalizeOrderNo(text)
  if (/^\d{10,}$/.test(normalizedWhole)) {
    return normalizedWhole
  }

  // Split by common separators and keep meaningful alphanumeric chunks.
  const parts = text
    .split(/[\s,，;；|｜/\\\-—_()（）【】\[\]]+/)
    .map((item) => normalizeOrderNo(item))
    .filter(Boolean)

  // Prefer pure numeric order ids (AliExpress order ids are typically long digits).
  const numericCandidates = parts.filter((item) => /^\d{10,}$/.test(item))
  if (numericCandidates.length > 0) {
    return numericCandidates[numericCandidates.length - 1]
  }

  // Fallback to long alphanumeric candidates and use the tail segment.
  const alphaNumCandidates = parts.filter((item) => /^[A-Za-z0-9]{10,}$/.test(item))
  if (alphaNumCandidates.length > 0) {
    return alphaNumCandidates[alphaNumCandidates.length - 1]
  }

  return ''
}

export function buildExternalRecords(input: ExternalRecordBuildInput): ExternalRecord[] {
  const { files, remarkHints, amountHints, mode, shopName, useParsedRemarkOrderNoOnly = false } = input
  const records: ExternalRecord[] = []

  files.forEach((item) => {
    const selectedSheet = getSelectedSheet(item.file)
    if (!selectedSheet) {
      return
    }

    const remarkColumn = pickDefaultColumn(selectedSheet.headers, remarkHints)
    const amountColumn = pickDefaultColumn(selectedSheet.headers, amountHints)
    const tradeNoColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_TRADE_NO_HINTS)
    const counterpartColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_COUNTERPART_HINTS)
    const productColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_PRODUCT_HINTS)
    const payMethodColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_PAY_METHOD_HINTS)
    const invoiceColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_INVOICE_HINTS)
    const tradeOrderColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_TRADE_ORDER_HINTS)
    const merchantOrderColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_ALIPAY_MERCHANT_ORDER_HINTS)

    selectedSheet.rows.forEach((row) => {
      const remark = normalizeCellValue(row[remarkColumn])
      const parsedRemarkOrderNo = normalizeOrderNo(extractOrderNoFromRemark(remark))
      if (!parsedRemarkOrderNo) {
        return
      }

      const orderNo = parsedRemarkOrderNo
      const parsedShopName = extractShopNameFromRemark(remark) || normalizeCellValue(shopName)
      const amount = toNumericValue(row[amountColumn])
      const hasAnyValue = Object.values(row).some((value) => normalizeCellValue(value))

      if (!hasAnyValue) {
        return
      }
      if (Math.abs(amount) < 0.000001) {
        return
      }

      records.push({
        订单号: orderNo,
        金额: Number(amount.toFixed(2)),
        备注: remark,
        是否开发票: mode === 'alipay' ? normalizeCellValue(row[invoiceColumn]) : '',
        店铺名: parsedShopName,
        订单号来源: useParsedRemarkOrderNoOnly ? '备注解析' : parsedRemarkOrderNo ? '备注解析' : '',
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

export function canProcessCalculation(input: ProcessValidationInput): boolean {
  const {
    ordersFilesLength,
    incomeFiles,
    refundFiles,
    effectiveIncomeDetailAmountColumnsLength,
    effectiveRefundDetailAmountColumnsLength,
    ordersIdColumn,
    effectiveOrdersStatusColumn
  } = input

  const incomeContentError = validateDetailFiles(incomeFiles, '订单明细/收支明细')
  const refundContentError = validateDetailFiles(refundFiles, '放退款订单明细')

  return Boolean(
    ordersFilesLength > 0 &&
      incomeFiles.length > 0 &&
      refundFiles.length > 0 &&
      !incomeContentError &&
      !refundContentError &&
      effectiveIncomeDetailAmountColumnsLength > 0 &&
      effectiveRefundDetailAmountColumnsLength > 0 &&
      ordersIdColumn &&
      effectiveOrdersStatusColumn
  )
}

export function buildOfflineFreightRecords(input: { files: MultiUploadItem[] }): OfflineFreightRecord[] {
  return buildOfflineFreightRecordsWithDiagnostics(input).records
}

export function buildOfflineFreightRecordsWithDiagnostics(input: { files: MultiUploadItem[] }): {
  records: OfflineFreightRecord[]
  diagnostics: OfflineFreightDiagnosticRow[]
} {
  const { files } = input
  const records: OfflineFreightRecord[] = []
  const diagnostics: OfflineFreightDiagnosticRow[] = []

  files.forEach((item) => {
    const selectedSheet = getSelectedSheet(item.file)
    if (!selectedSheet) {
      return
    }

    const orderColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_OFFLINE_ORDER_HINTS)
    const amountColumn = pickDefaultColumn(selectedSheet.headers, ['费用金额', ...DEFAULT_OFFLINE_AMOUNT_HINTS])
    const categoryColumn = pickDefaultColumn(selectedSheet.headers, DEFAULT_OFFLINE_CATEGORY_HINTS)
    const remarkColumn = pickDefaultColumn(selectedSheet.headers, ['备注', '说明', 'memo'])

    const keyOrder = selectedSheet.headers
    const rowsAsCells = selectedSheet.rows.map((row) => keyOrder.map((key) => normalizeCellValue(row[key])))

    const findColumnIndex = (cells: string[], hints: string[]): number => {
      const lowered = cells.map((cell) => normalizeCellValue(cell).toLowerCase())
      for (const hint of hints) {
        const idx = lowered.findIndex((cell) => cell.includes(hint.toLowerCase()))
        if (idx >= 0) {
          return idx
        }
      }
      return -1
    }

    let detectedHeaderRowIndex = -1
    let detectedOrderIdx = -1
    let detectedAmountIdx = -1
    let detectedCategoryIdx = -1
    let detectedRemarkIdx = -1

    for (let i = 0; i < rowsAsCells.length; i += 1) {
      const cells = rowsAsCells[i]
      const orderIdx = findColumnIndex(cells, DEFAULT_OFFLINE_ORDER_HINTS)
      const amountIdx = findColumnIndex(cells, ['费用金额', ...DEFAULT_OFFLINE_AMOUNT_HINTS])
      const categoryIdx = findColumnIndex(cells, DEFAULT_OFFLINE_CATEGORY_HINTS)
      const remarkIdx = findColumnIndex(cells, ['备注', '说明', 'memo'])

      const hitCount = [orderIdx, amountIdx, categoryIdx].filter((idx) => idx >= 0).length
      if (hitCount >= 2) {
        detectedHeaderRowIndex = i
        detectedOrderIdx = orderIdx
        detectedAmountIdx = amountIdx
        detectedCategoryIdx = categoryIdx
        detectedRemarkIdx = remarkIdx
        break
      }
    }

    const useDetectedHeader = detectedHeaderRowIndex >= 0 && detectedOrderIdx >= 0 && detectedAmountIdx >= 0

    diagnostics.push({
      来源文件: item.file.fileName,
      来源Sheet: selectedSheet.name,
      行号: 0,
      客户订单号: '',
      提取订单号: '',
      费用类别: '',
      金额原值: '',
      金额标准化: 0,
      状态: '已跳过',
      原因: useDetectedHeader
        ? `已启用自动表头识别，表头位于数据第${detectedHeaderRowIndex + 2}行`
        : `未识别到嵌入式表头，使用默认列映射（客户订单号列：${orderColumn || '未识别'}，金额列：${amountColumn || '未识别'}）`
    })

    const getValueByDetectedColumn = (cells: string[], detectedIdx: number, fallbackColumn: string, row: RowData): string => {
      if (useDetectedHeader && detectedIdx >= 0 && detectedIdx < cells.length) {
        return normalizeCellValue(cells[detectedIdx])
      }
      return normalizeCellValue(row[fallbackColumn])
    }

    selectedSheet.rows.forEach((row, index) => {
      const cells = rowsAsCells[index] || []

      if (useDetectedHeader && index <= detectedHeaderRowIndex) {
        diagnostics.push({
          来源文件: item.file.fileName,
          来源Sheet: selectedSheet.name,
          行号: index + 2,
          客户订单号: '',
          提取订单号: '',
          费用类别: '',
          金额原值: '',
          金额标准化: 0,
          状态: '已跳过',
          原因: '抬头/表头行'
        })
        return
      }

      const hasAnyValue = Object.values(row).some((value) => normalizeCellValue(value))
      const customerOrderText = getValueByDetectedColumn(cells, detectedOrderIdx, orderColumn, row)
      const orderNo = extractOrderNoFromCustomerOrder(customerOrderText)
      const amountRaw = getValueByDetectedColumn(cells, detectedAmountIdx, amountColumn, row)
      const amount = toNumericValue(amountRaw)
      const feeCategory = getValueByDetectedColumn(cells, detectedCategoryIdx, categoryColumn, row) || '未分类线下费用'
      const remarkText = getValueByDetectedColumn(cells, detectedRemarkIdx, remarkColumn, row)

      let reason = ''
      if (!hasAnyValue) {
        reason = '空行'
      } else if (!customerOrderText) {
        reason = `客户订单号列为空（当前识别列：${orderColumn || '未识别'}）`
      } else if (!orderNo) {
        reason = '客户订单号未提取到有效速卖通订单号'
      } else if (Math.abs(amount) < 0.000001) {
        reason = `金额为空或为0（当前识别列：${amountColumn || '未识别'}）`
      }

      diagnostics.push({
        来源文件: item.file.fileName,
        来源Sheet: selectedSheet.name,
        行号: index + 2,
        客户订单号: customerOrderText,
        提取订单号: orderNo,
        费用类别: feeCategory,
        金额原值: amountRaw,
        金额标准化: Number(amount.toFixed(2)),
        状态: reason ? '已跳过' : '纳入计算',
        原因: reason || '匹配成功'
      })

      if (reason) {
        return
      }

      records.push({
        订单号: orderNo,
        金额: Number(amount.toFixed(2)),
        费用类别: feeCategory,
        客户订单号: customerOrderText,
        备注: remarkText,
        来源文件: item.file.fileName,
        来源Sheet: selectedSheet.name
      })
    })
  })

  return { records, diagnostics }
}

export function getProcessDisabledReason(input: ProcessValidationInput): string {
  const {
    ordersFilesLength,
    incomeFiles,
    refundFiles,
    effectiveIncomeDetailAmountColumnsLength,
    effectiveRefundDetailAmountColumnsLength,
    ordersIdColumn,
    effectiveOrdersStatusColumn
  } = input

  if (ordersFilesLength === 0) return '请先上传订单表文件。'
  if (incomeFiles.length === 0) return '请先上传订单明细/收支明细文件。'
  if (refundFiles.length === 0) return '请先上传放退款订单明细文件。'

  if (!ordersIdColumn) return '订单表缺少“订单号字段”，请检查订单表列名。'
  if (!effectiveOrdersStatusColumn) return '订单表缺少“订单状态字段”，请检查订单表列名。'

  if (effectiveIncomeDetailAmountColumnsLength === 0) return '订单明细/收支明细未识别到金额字段。'
  if (effectiveRefundDetailAmountColumnsLength === 0) return '放退款明细未识别到费用字段。'

  const incomeContentError = validateDetailFiles(incomeFiles, '订单明细/收支明细')
  if (incomeContentError) return incomeContentError
  const refundContentError = validateDetailFiles(refundFiles, '放退款订单明细')
  if (refundContentError) return refundContentError

  return ''
}

export function toDateText(date: Date): string {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`
}

export function getExportFileName(reportName: string, dateText: string, meta: { shopName: string; shopId: string; subsidiary: string }): string {
  const companyPart = safeFileSegment(meta.subsidiary, '未分公司')
  const storePart = safeFileSegment(meta.shopName || meta.shopId, '未命名店铺')
  return `${companyPart}_${storePart}_${reportName}_${dateText}.xlsx`
}
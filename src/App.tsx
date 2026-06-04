import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'
import * as Domain from './app/calculatorDomain'
import {
  buildExternalRecords,
  buildOfflineFreightRecordsWithDiagnostics,
  canProcessCalculation,
  getExportFileName,
  getProcessDisabledReason as getProcessDisabledReasonText,
  toDateText
} from './app/utils/appHelpers'
import {
  ActionPanel,
  ResultPreviewPanel,
  StoreMetaSection
} from './app/components/AppSections'
import {
  DetailUploadCard as DetailUploadCardInternal,
  FreightUploadCard as FreightUploadCardInternal,
  OrdersUploadCard as OrdersUploadCardInternal,
  SimpleUploadCard as SimpleUploadCardInternal
} from './app/components/UploadCards'

type RowData = Domain.RowData
type MultiUploadItem = Domain.MultiUploadItem
type IncomeUploadItem = Domain.IncomeUploadItem
type TableKind = Domain.TableKind
type ProcessResult = Domain.ProcessResult
type PersistedUploadBundle = Domain.PersistedUploadBundle
type ExternalRecord = Domain.ExternalRecord
type ResultSummary = Domain.ResultSummary

const {
  DEFAULT_ORDER_ID_HINTS,
  DEFAULT_ORDER_STATUS_HINTS,
  DEFAULT_ORDER_TIME_HINTS,
  DEFAULT_FREIGHT_ID_HINTS,
  DEFAULT_FREIGHT_FULFILLMENT_HINTS,
  DEFAULT_FREIGHT_WAYBILL_HINTS,
  DEFAULT_FREIGHT_CNY_HINTS,
  DEFAULT_FREIGHT_USD_HINTS,
  DEFAULT_REFUND_TYPE_HINTS,
  DEFAULT_INCOME_FLOW_TYPE_HINTS,
  DEFAULT_INCOME_FEE_ITEM_HINTS,
  DEFAULT_INCOME_MOVEMENT_HINTS,
  DEFAULT_REFUND_PRODUCT_NAME_HINTS,
  DEFAULT_REFUND_SKU_ID_HINTS,
  REFUND_BASE_AMOUNT_HINTS,
  DEFAULT_ALIPAY_REMARK_HINTS,
  DEFAULT_ALIPAY_AMOUNT_HINTS,
  hasKeyword,
  sortRefundTypesForDisplay,
  normalizeCellValue,
  normalizeOrderNo,
  toNumericValue,
  convertFreightToMovement,
  normalizeMoney,
  toTypeAmountColumn,
  toFeeItemAmountColumn,
  toTypeFeeItemAmountColumn,
  toLogisticsSubItemAmountColumn,
  isOrderIncomeSource,
  isExcludedIncomeDetailColumn,
  inferIncomeDetailAmountColumns,
  toIncomeDetailMovementByFlowType,
  toRefundDetailMovement,
  sortByOrderNo,
  sortByRefundTypePriority,
  isCompletedOrderStatus,
  readUploadCache,
  writeUploadCache,
  clearUploadCache,
  pickDefaultColumn,
  pickOptionalColumn,
  resolveColumnSelection,
  toUploadedState,
  getSelectedSheet,
  buildCountMap,
  buildRowsByOrderMap,
  prefixRow,
  summarizeAmountByOrder,
  buildRefundRiskMap,
  dedupeByKey,
  validateDetailFiles,
  findFreightOrderNoInRow,
  findFreightFallbackAmount,
  formatDateTime
} = Domain

function App() {
  const [ordersFiles, setOrdersFiles] = useState<MultiUploadItem[]>([])
  const [incomeFiles, setIncomeFiles] = useState<IncomeUploadItem[]>([])
  const [refundFiles, setRefundFiles] = useState<IncomeUploadItem[]>([])
  const [freightFiles, setFreightFiles] = useState<MultiUploadItem[]>([])
  const [alipayFiles, setAlipayFiles] = useState<MultiUploadItem[]>([])
  const [offlineFiles, setOfflineFiles] = useState<MultiUploadItem[]>([])

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

  function scoreDecodedCsvText(text: string): number {
    const sample = text.slice(0, 4000)
    const headerHints = [
      '交易时间',
      '交易分类',
      '交易对方',
      '商品说明',
      '金额',
      '备注',
      '订单号',
      '订单状态',
      '收/付款方式',
      '交易订单号',
      '商家订单号'
    ]

    let score = 0
    headerHints.forEach((hint) => {
      if (sample.includes(hint)) {
        score += 8
      }
    })

    if (sample.includes('��')) {
      score -= 30
    }

    const mojibakeMatches = sample.match(/[æäåçéïð]/g)
    if (mojibakeMatches) {
      score -= mojibakeMatches.length
    }

    return score
  }

  function readWorkbookFromUpload(fileName: string, buffer: ArrayBuffer): XLSX.WorkBook {
    const isCsv = /\.csv$/i.test(fileName)
    if (!isCsv) {
      return XLSX.read(buffer, { type: 'array' })
    }

    const bytes = new Uint8Array(buffer)
    const hasUtf8Bom = bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf
    const candidateEncodings = hasUtf8Bom ? ['utf-8', 'gb18030', 'gbk'] : ['utf-8', 'gb18030', 'gbk']

    let bestText = ''
    let bestScore = Number.NEGATIVE_INFINITY

    candidateEncodings.forEach((encoding) => {
      try {
        const decoded = new TextDecoder(encoding as string).decode(buffer)
        const score = scoreDecodedCsvText(decoded)
        if (score > bestScore) {
          bestScore = score
          bestText = decoded
        }
      } catch {
        // Ignore unsupported decoders and keep trying the next candidate.
      }
    })

    return XLSX.read(bestText, { type: 'string' })
  }

  async function handleFileUpload(event: React.ChangeEvent<HTMLInputElement>, table: TableKind) {
    const files = Array.from(event.target.files || [])
    if (files.length === 0) {
      return
    }

    setErrorMessage('')

    try {
      for (const file of files) {
        const buffer = await file.arrayBuffer()
        const workbook = readWorkbookFromUpload(file.name, buffer)
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

  function removeIncomeFile(itemId: string, table: 'income' | 'refund' = 'income') {
    setErrorMessage('')
    const setter = table === 'income' ? setIncomeFiles : setRefundFiles
    setter((prev) => prev.filter((item) => item.id !== itemId))
  }

  function canProcess(): boolean {
    return canProcessCalculation({
      ordersFilesLength: ordersFiles.length,
      incomeFiles,
      refundFiles,
      effectiveIncomeDetailAmountColumnsLength: effectiveIncomeDetailAmountColumns.length,
      effectiveRefundDetailAmountColumnsLength: effectiveRefundDetailAmountColumns.length,
      ordersIdColumn,
      effectiveOrdersStatusColumn
    })
  }

  function getProcessDisabledReason(): string {
    return getProcessDisabledReasonText({
      ordersFilesLength: ordersFiles.length,
      incomeFiles,
      refundFiles,
      effectiveIncomeDetailAmountColumnsLength: effectiveIncomeDetailAmountColumns.length,
      effectiveRefundDetailAmountColumnsLength: effectiveRefundDetailAmountColumns.length,
      ordersIdColumn,
      effectiveOrdersStatusColumn
    })
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
              const canUseFlowMovementMode = Boolean(flowTypeValue && movementColumn)
              if (canUseFlowMovementMode) {
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

              // Fallback for order detail exports (e.g. Russia) that provide multiple amount columns
              // but do not include explicit flow type / movement columns.
              selectedColumns.forEach((column) => {
                if (column === normalizeCellValue(baseRow.收支匹配字段名) || isExcludedIncomeDetailColumn(column)) {
                  return
                }

                const movementAmount = toIncomeDetailMovementByFlowType(column, baseRow[column], flowTypeValue)
                if (Math.abs(movementAmount) < 0.000001) {
                  return
                }

                rows.push({
                  订单号: orderNo,
                  收支类型: column,
                  变动金额: movementAmount,
                  费用项: column,
                  币种: 'CNY',
                  来源: sourceLabel,
                  放退款类型: '',
                  收支来源文件: baseRow.收支来源文件,
                  收支来源Sheet: baseRow.收支来源Sheet
                })
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
        console.groupCollapsed('[Calc Debug] 线下运费解析诊断')
        console.log('offline rows total:', offlineFreightDiagnosticRows.length)
        console.log('offline rows included:', offlineFreightDiagnosticRows.filter((row) => row.状态 === '纳入计算').length)
        console.log('offline rows skipped:', offlineFreightDiagnosticRows.filter((row) => row.状态 === '已跳过').length)
        console.table(offlineFreightDiagnosticRows.slice(0, 40))
        console.groupEnd()
        console.groupEnd()
      }

      const orderIds = new Set(
        orderRows
          .map((row) => normalizeOrderNo(row[ordersIdColumn]))
          .filter(Boolean)
      )

      const scopedAlipayRecords = alipayRecords.filter((row) => orderIds.has(row.订单号))
      const scopedOfflineRecords = offlineFreightRecords.filter((row) => orderIds.has(row.订单号))
      const unmatchedAlipayRows = sortByOrderNo(
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
            '收/付款方式': row.收付款方式,
            交易订单号: row.交易订单号,
            商家订单号: row.商家订单号,
            备注: row.备注,
            来源文件: row.来源文件,
            来源Sheet: row.来源Sheet
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

      const offlineRowsByOrder = new Map<string, Array<{
        订单号: string
        金额: number
        费用类别: string
        客户订单号: string
        备注: string
        来源文件: string
        来源Sheet: string
      }>>()
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

      const ordersWithoutAlipayRows = sortByOrderNo(
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

      const alipayMultiplicityRows = [
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
          采购费用: 0,
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
        const purchaseExpense = normalizeMoney(current.采购费用)
        const incomeLogisticsExpense = normalizeMoney(current.收支表_支出物流费用)
        const jzgFreightExpense = normalizeMoney(current.金掌柜物流费支出)
        const offlineFreightExpense = normalizeMoney(current.线下运费)
        const logisticsExpenseTotal = normalizeMoney(incomeLogisticsExpense + jzgFreightExpense + offlineFreightExpense)
        current.物流支出总和 = logisticsExpenseTotal
        const totalFreight = logisticsExpenseTotal
        const expectedFromIncomeAndFreight = normalizeMoney(
          incomeFromOrderDetailExcludingLogistics - logisticsExpenseTotal
        )
        const netAfterPurchaseAndFreight = normalizeMoney(expectedFromIncomeAndFreight - purchaseExpense)
        const expectedFromRefundAndFreight = normalizeMoney(
          incomeFromRefundDetail - incomeLogisticsExpense - jzgFreightExpense - offlineFreightExpense
        )
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
        current.最终收入_未扣运费 = incomeFromOrderDetailExcludingLogistics
        current.最终收入_扣运费 = netAfterPurchaseAndFreight

        current.收入合计 = incomeFromOrderDetail
        current.支出合计 = normalizeMoney(Math.max(totalFreight, 0) + purchaseExpense)
        current.总收支 = netAfterPurchaseAndFreight

        aggregatedMap.set(orderNo, current)
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
          物流费用_支出表: normalizeMoney(row.收支表_支出物流费用),
          物流费用_金掌柜: normalizeMoney(row.金掌柜物流费支出),
          线下物流: normalizeMoney(row.线下运费),
          总物流费用: normalizeMoney(row.物流支出总和)
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
          allOfflineCategoryColumns.forEach((col) => {
            if (next[col] === undefined) {
              next[col] = 0
            }
          })

          return {
            订单号: normalizeOrderNo(next.订单号),
            订单时间: normalizeCellValue(next.订单时间),
            订单状态: normalizeCellValue(next.订单状态),
            订单预计可得: normalizeMoney(next.最终收入_未扣运费),
            物流费用_支出表: normalizeMoney(next.收支表_支出物流费用),
            物流费用_金掌柜: normalizeMoney(next.金掌柜物流费支出),
            线下物流: normalizeMoney(next.线下运费),
            总物流费用: normalizeMoney(next.物流支出总和),
            采购费用: normalizeMoney(next.采购费用),
            净利润: normalizeMoney(next.最终收入_扣运费),
            放退款类型: normalizeCellValue(next.放退款类型),
            放退款核查标记: normalizeCellValue(next.放退款核查标记),
            待结算金额合计: pendingSettlement,
            放退款_金额项合计: normalizeMoney(next.放退款_金额项合计),
            放退款_其他费用合计: normalizeMoney(next.放退款_其他费用合计),
            收入_按放退款明细: normalizeMoney(next.收入_按放退款明细),
            收入_按订单明细: normalizeMoney(next.收入_按订单明细),
            订单明细_放款金额合计: normalizeMoney(next.订单明细_放款金额合计),
            订单明细_平台分账金额合计: normalizeMoney(next.订单明细_平台分账金额合计),
            收入_收支明细表: normalizeMoney(next.收入_收支明细表),
            支出_收支明细表: normalizeMoney(next.支出_收支明细表),
            收支总和_不含物流费用: normalizeMoney(next.收支总和_不含物流费用),
            差异_收支不含物流减退放款: normalizeMoney(next.差异_收支不含物流减退放款),
            收入差异_订单明细减放退款: normalizeMoney(next.收入差异_订单明细减放退款),
            收入校验状态: normalizeCellValue(next.收入校验状态),
            预计可得_按收支及运费: normalizeMoney(next.预计可得_按收支及运费),
            预计可得_按退放款及运费: normalizeMoney(next.预计可得_按退放款及运费),
            预计可得差异_收支减退放款: normalizeMoney(next.预计可得差异_收支减退放款),
            预计可得校验状态: normalizeCellValue(next.预计可得校验状态),
            备注: mergedRemark,
            ...Object.fromEntries(allOfflineCategoryColumns.map((col) => [col, normalizeMoney(next[col])])),
            ...Object.fromEntries(allLogisticsSubItemColumns.map((col) => [col, normalizeMoney(next[col])])),
            ...Object.fromEntries(allTypeFeeItemColumns.map((col) => [col, normalizeMoney(next[col])])),
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
          收入校验状态: normalizeCellValue(row.收入校验状态),
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
          const finalIncome = toNumericValue(row.净利润)
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
          acc.totalExpense += logistics + purchaseExpense
          acc.logisticsAmount += logistics
          acc.nonLogisticsAmount += purchaseExpense
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

    const dateText = toDateText(new Date())
    XLSX.writeFile(workbook, getExportFileName('订单聚合表', dateText, { shopName, shopId, subsidiary }))
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

    const dateText = toDateText(new Date())
    XLSX.writeFile(workbook, getExportFileName('业务明细表', dateText, { shopName, shopId, subsidiary }))
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
      XLSX.utils.json_to_sheet(result.ordersWithoutAlipayRows),
      '订单有但支付宝缺失'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.unmatchedAlipayRows),
      '支付宝无对应订单'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.alipayMultiplicityRows),
      '支付宝匹配异常'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.offlineFreightDiagnosticRows),
      '线下运费解析诊断'
    )

    const dateText = toDateText(new Date())
    XLSX.writeFile(workbook, getExportFileName('其他对账表', dateText, { shopName, shopId, subsidiary }))
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

      <StoreMetaSection
        shopId={shopId}
        shopName={shopName}
        subsidiary={subsidiary}
        onShopIdChange={setShopId}
        onShopNameChange={setShopName}
        onSubsidiaryChange={setSubsidiary}
      />

      <section className="upload-sections">
        <OrdersUploadCardInternal
          files={ordersFiles}
          onFileUpload={handleFileUpload}
          onRemove={removeOrdersFile}
        />
        <FreightUploadCardInternal
          files={freightFiles}
          onFileUpload={handleFileUpload}
          onRemove={removeFreightFile}
        />
        <DetailUploadCardInternal
          title="3) 订单明细/收支明细表"
          description={
            <>
              请上传
              <span className="emphasis-russia">【俄罗斯】结算月至今的订单明细表；【非俄罗斯】结算月至今的订单收支明细表。</span>
            </>
          }
          table="income"
          files={incomeFiles}
          emptyText="尚未上传订单明细/收支明细文件"
          onFileUpload={handleFileUpload}
          onRemove={removeIncomeFile}
        />
        <DetailUploadCardInternal
          title="4) 放退款订单明细"
          description={
            <>
              请上传
              <span className="emphasis-russia">【俄罗斯】和【非俄罗斯】结算月至今</span>
              的放退款明细文件。
            </>
          }
          table="refund"
          files={refundFiles}
          emptyText="尚未上传放退款订单明细文件"
          onFileUpload={handleFileUpload}
          onRemove={removeIncomeFile}
        />
        <SimpleUploadCardInternal
          title="5) 支付宝订单记录（采购）"
          description="支持多文件上传。备注按“店铺名-订单号”匹配，系统会自动摘取短线后的订单号。"
          table="alipay"
          files={alipayFiles}
          emptyText="尚未上传支付宝订单记录"
          onFileUpload={handleFileUpload}
          onRemove={removeSimpleUploadFile}
        />
        <SimpleUploadCardInternal
          title="6) 线下发货订单记录"
          description="支持多文件上传。备注按“店铺名-订单号”匹配，系统会自动摘取短线后的订单号。"
          table="offline"
          files={offlineFiles}
          emptyText="尚未上传线下发货记录"
          onFileUpload={handleFileUpload}
          onRemove={removeSimpleUploadFile}
        />
      </section>

      <ActionPanel
        isProcessing={isProcessing}
        canProcess={canProcess()}
        result={result}
        lastCalculatedAt={lastCalculatedAt}
        calculationCount={calculationCount}
        processDisabledReason={getProcessDisabledReason()}
        onRunCalculation={runCalculation}
        onExportAggregated={exportAggregatedWorkbook}
        onExportBusinessDetail={exportBusinessDetailWorkbook}
        onExportOtherSheets={exportOtherSheetsWorkbook}
      />

      {errorMessage && <div className="error-box">{errorMessage}</div>}

      {result && <ResultPreviewPanel result={result} />}
    </main>
  )
}

export default App

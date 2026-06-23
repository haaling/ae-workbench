import {
  DEFAULT_INCOME_FEE_ITEM_HINTS,
  DEFAULT_INCOME_FLOW_TYPE_HINTS,
  DEFAULT_INCOME_MOVEMENT_HINTS,
  REFUND_BASE_AMOUNT_HINTS,
  getSelectedSheet,
  hasKeyword,
  isExcludedIncomeDetailColumn,
  isRefundTypeExcludedFromExpense,
  normalizeCellValue,
  normalizeMoney,
  normalizeOrderNo,
  pickOptionalColumn,
  toIncomeDetailMovementByFlowType,
  toNumericValue,
  toRefundDetailMovement,
  type RowData
} from '../calculatorDomain'
import type { IncomeUploadItem } from './types'

export function buildIncomeRefundEntries(input: {
  incomeFiles: IncomeUploadItem[]
  refundFiles: IncomeUploadItem[]
  effectiveIncomeDetailAmountColumns: string[]
  effectiveRefundDetailAmountColumns: string[]
  effectiveRefundTypeColumn: string
}): {
  incomeDetailRows: RowData[]
  refundDetailRows: RowData[]
  incomeEntries: RowData[]
} {
  const {
    incomeFiles,
    refundFiles,
    effectiveIncomeDetailAmountColumns,
    effectiveRefundDetailAmountColumns,
    effectiveRefundTypeColumn
  } = input

  const buildDetailEntries = (
    files: IncomeUploadItem[],
    sourceLabel: string,
    selectedColumns: string[],
    typeColumn = ''
  ): RowData[] => {
    const rows: RowData[] = []

    files.forEach((item) => {
      const selectedSheet = getSelectedSheet(item.file)
      if (!selectedSheet) {
        return
      }

      const isIncomeMode = !typeColumn
      const incomeStatementTypeColumn = selectedSheet.headers.find((header) => normalizeCellValue(header) === '收支类型') || ''
      const effectiveSourceLabel = isIncomeMode && incomeStatementTypeColumn ? '订单收支明细表' : sourceLabel
      const flowTypeColumn = incomeStatementTypeColumn || (
        item.flowTypeColumn && selectedSheet.headers.includes(item.flowTypeColumn)
          ? item.flowTypeColumn
          : pickOptionalColumn(selectedSheet.headers, DEFAULT_INCOME_FLOW_TYPE_HINTS)
      )
      const feeItemColumn = pickOptionalColumn(selectedSheet.headers, DEFAULT_INCOME_FEE_ITEM_HINTS)
      const movementColumn =
        pickOptionalColumn(selectedSheet.headers, DEFAULT_INCOME_MOVEMENT_HINTS) ||
        pickOptionalColumn(selectedColumns, DEFAULT_INCOME_MOVEMENT_HINTS)

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
          收支来源区块: effectiveSourceLabel
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
              来源: effectiveSourceLabel,
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
              来源: effectiveSourceLabel,
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

          const refundRowCommissionTotal = normalizeMoney(
            selectedColumns
              .filter((col) => normalizeCellValue(col).includes('佣金'))
              .reduce((sum, col) => sum + Math.abs(toNumericValue(baseRow[col])), 0)
          )

          const movementAmount = typeColumn
            ? toRefundDetailMovement(column, baseRow[column], baseRow[typeColumn])
            : toIncomeDetailMovementByFlowType(
              column,
              baseRow[column],
              flowTypeColumn ? baseRow[flowTypeColumn] : baseRow.收支类型
            )
          const refundTypeText = typeColumn ? normalizeCellValue(baseRow[typeColumn]) : ''
          const shouldKeepZeroRefundRow = isRefundTypeExcludedFromExpense(refundTypeText)

          if (Math.abs(movementAmount) < 0.000001 && !shouldKeepZeroRefundRow) {
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
            放退款行佣金合计: refundRowCommissionTotal,
            放退款税费候选金额:
              typeColumn && hasKeyword(column, REFUND_BASE_AMOUNT_HINTS) && refundRowCommissionTotal <= 0.01
                ? Math.abs(movementAmount)
                : 0,
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

  return {
    incomeDetailRows,
    refundDetailRows,
    incomeEntries: [...incomeDetailRows, ...refundDetailRows]
  }
}

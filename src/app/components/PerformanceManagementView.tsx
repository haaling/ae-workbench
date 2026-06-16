import { useMemo, useState } from 'react'
import { normalizeCellValue, normalizeMoney, toNumericValue, type RowData } from '../calculatorDomain'
import type { PerformanceWorkflowItem, WorkflowReviewer } from '../types/workflow'

const FIXED_WORKFLOW_STATUS_OPTIONS = ['draft', 'pushed', 'confirmed', 'archived'] as const

const PUSH_PREVIEW_COLUMNS = [
  { label: '订单号', aliases: ['订单号', '订单编号', '平台订单号', '订单id', '订单ID'] },
  { label: '订单时间', aliases: ['订单时间', '下单时间', '付款时间', '创建时间'] },
  { label: '订单状态', aliases: ['订单状态', '状态'] },
  { label: '订单预计可得', aliases: ['订单预计可得', '预计可得', '预计可得金额'] },
  { label: '总物流费用', aliases: ['总物流费用', '物流费用', '总物流费'] },
  { label: '采购费用', aliases: ['采购费用', '采购成本', '总采购费用'] },
  { label: '支付宝是否开发票', aliases: ['支付宝是否开发票', '是否开票', '支付宝开票'] },
  { label: '净利润', aliases: ['净利润', '总利润', '总收支'] }
] as const

type PerformanceManagementViewProps = {
  mode: 'push' | 'list'
  listTab: 'current' | 'history'
  uploadedPerformanceFileName: string
  workflowReviewers: WorkflowReviewer[]
  uploadedPerformanceRows: RowData[]
  performanceWorkflows: PerformanceWorkflowItem[]
  financeArchiveReadyWorkflows: PerformanceWorkflowItem[]
  workflowStatusLabelMap: Record<string, string>
  isWorkflowSubmitting: boolean
  isWorkflowListLoading: boolean
  onUploadCheckedPerformanceFile: (event: React.ChangeEvent<HTMLInputElement>) => void
  onCreateAndPushPerformanceWorkflow: () => void
  onRefreshWorkflowData: () => void
  onArchivePerformanceWorkflow: (workflowId: string) => void
}

export function PerformanceManagementView(props: PerformanceManagementViewProps) {
  const {
    mode,
    listTab,
    uploadedPerformanceFileName,
    workflowReviewers,
    uploadedPerformanceRows,
    performanceWorkflows,
    workflowStatusLabelMap,
    isWorkflowSubmitting,
    isWorkflowListLoading,
    onUploadCheckedPerformanceFile,
    onCreateAndPushPerformanceWorkflow,
    onRefreshWorkflowData,
    onArchivePerformanceWorkflow
  } = props

  const isPushMode = mode === 'push'
  const [currentStatusFilter, setCurrentStatusFilter] = useState('all')
  const [historyMonthFilter, setHistoryMonthFilter] = useState('')
  const [historyEmployeeFilter, setHistoryEmployeeFilter] = useState('')

  const now = new Date()
  const currentPeriod = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`

  const projectedUploadedPerformanceRows = useMemo(() => {
    return uploadedPerformanceRows.map((row) => {
      const source = row || {}
      const output: RowData = {}
      for (const column of PUSH_PREVIEW_COLUMNS) {
        const matchedKey = column.aliases.find((alias) => Object.prototype.hasOwnProperty.call(source, alias))
        output[column.label] = matchedKey ? normalizeCellValue(source[matchedKey]) : ''
      }
      return output
    })
  }, [uploadedPerformanceRows])

  const previewTotals = useMemo(() => {
    const isInvoiceMarkedYes = (value: unknown): boolean => {
      const text = normalizeCellValue(value).toLowerCase()
      return text === '是' || text === 'y' || text === 'yes' || text === 'true' || text.includes('是')
    }

    return projectedUploadedPerformanceRows.reduce<{ netProfitTotal: number; invoicedTotal: number }>(
      (acc, row) => {
        const netProfit = toNumericValue(row.净利润)
        const purchaseAmount = toNumericValue(row.采购费用)
        acc.netProfitTotal += netProfit
        if (isInvoiceMarkedYes(row.支付宝是否开发票)) {
          acc.invoicedTotal += purchaseAmount
        }
        return acc
      },
      { netProfitTotal: 0, invoicedTotal: 0 }
    )
  }, [projectedUploadedPerformanceRows])

  const allStatuses = useMemo(
    () => FIXED_WORKFLOW_STATUS_OPTIONS.filter((status) => workflowStatusLabelMap[status] || status),
    [workflowStatusLabelMap]
  )

  const currentMonthWorkflows = useMemo(
    () => performanceWorkflows.filter((item) => normalizeCellValue(item.period) === currentPeriod),
    [performanceWorkflows, currentPeriod]
  )

  const filteredCurrentMonthWorkflows = useMemo(
    () => currentMonthWorkflows.filter((item) => currentStatusFilter === 'all' || normalizeCellValue(item.status) === currentStatusFilter),
    [currentMonthWorkflows, currentStatusFilter]
  )

  const historySource = useMemo(
    () => performanceWorkflows.filter((item) => normalizeCellValue(item.period) !== currentPeriod),
    [performanceWorkflows, currentPeriod]
  )

  const historyEmployeeOptions = useMemo(
    () => Array.from(new Set([
      ...workflowReviewers.map((item) => normalizeCellValue(item.username)).filter(Boolean),
      ...historySource.map((item) => normalizeCellValue(item.submittedByName) || normalizeCellValue(item.assignedToUserName) || '未指定员工')
    ]))
      .sort((left, right) => left.localeCompare(right, 'zh-Hans-CN', { sensitivity: 'base' })),
    [workflowReviewers, historySource]
  )

  const filteredHistorySource = useMemo(
    () => historySource.filter((item) => {
      const period = normalizeCellValue(item.period) || '未指定月份'
      const employee = normalizeCellValue(item.submittedByName) || normalizeCellValue(item.assignedToUserName) || '未指定员工'
      const monthMatched = !historyMonthFilter || period === historyMonthFilter
      const employeeMatched = !historyEmployeeFilter || employee === historyEmployeeFilter
      return monthMatched && employeeMatched
    }),
    [historySource, historyMonthFilter, historyEmployeeFilter]
  )

  const historyByPeriodEmployee = useMemo(() => {
    const candidateKeys = ['总利润', '净利润', '总收支']

    const monthMap = new Map<string, Map<string, {
      employeeName: string
      workflowCount: number
      archivedCount: number
      confirmedCount: number
      storeSet: Set<string>
      rowCount: number
      totalProfit: number
    }>>()

    for (const item of filteredHistorySource) {
      const period = normalizeCellValue(item.period) || '未指定月份'
      const employeeName =
        normalizeCellValue(item.submittedByName) ||
        normalizeCellValue(item.assignedToUserName) ||
        '未指定员工'

      const periodMap = monthMap.get(period) || new Map()
      const current = periodMap.get(employeeName) || {
        employeeName,
        workflowCount: 0,
        archivedCount: 0,
        confirmedCount: 0,
        storeSet: new Set<string>(),
        rowCount: 0,
        totalProfit: 0
      }

      current.workflowCount += 1
      if (normalizeCellValue(item.status) === 'archived') {
        current.archivedCount += 1
      }
      if (normalizeCellValue(item.status) === 'confirmed') {
        current.confirmedCount += 1
      }
      const storeLabel = normalizeCellValue(item.storeLabel)
      if (storeLabel) {
        current.storeSet.add(storeLabel)
      }
      current.rowCount += Number(item.rowCountUploaded || item.rowCountCalculated || 0) || 0
      current.totalProfit = normalizeMoney(
        current.totalProfit +
          (item.uploadedRows || []).reduce((sum, row) => {
            const key = candidateKeys.find((candidate) => row[candidate] !== undefined)
            return key ? sum + toNumericValue(row[key]) : sum
          }, 0)
      )

      periodMap.set(employeeName, current)
      monthMap.set(period, periodMap)
    }

    return Array.from(monthMap.entries())
      .sort((left, right) => right[0].localeCompare(left[0], 'zh-Hans-CN', { numeric: true, sensitivity: 'base' }))
      .map(([period, employeeMap]) => {
        const rows = Array.from(employeeMap.values())
          .sort((left, right) => left.employeeName.localeCompare(right.employeeName, 'zh-Hans-CN', { sensitivity: 'base' }))
          .map((row) => ({
            employeeName: row.employeeName,
            workflowCount: row.workflowCount,
            archivedCount: row.archivedCount,
            confirmedCount: row.confirmedCount,
            storeCount: row.storeSet.size,
            rowCount: row.rowCount,
            totalProfit: normalizeMoney(row.totalProfit)
          }))

        return {
          period,
          rows,
          monthWorkflowCount: rows.reduce((sum, row) => sum + row.workflowCount, 0),
          monthStoreCount: rows.reduce((sum, row) => sum + row.storeCount, 0),
          monthTotalProfit: normalizeMoney(rows.reduce((sum, row) => sum + row.totalProfit, 0))
        }
      })
  }, [filteredHistorySource])

  return (
    <section className="upload-card">
      <h2>{isPushMode ? '绩效推送' : '绩效管理'}</h2>
      <p>
        {isPushMode
          ? '在计算页面上传核对后的绩效文件并推送给员工。'
          : '通过子菜单查看当月绩效和历史绩效，并按筛选条件快速定位数据。'}
      </p>
      {isPushMode && (
        <>
        <div className="control-row">
            <label>上传绩效文件（财务核对后）</label>
            <label className="file-input-label">
            <span>{uploadedPerformanceFileName ? `已选择：${uploadedPerformanceFileName}` : '选择 xlsx/csv 文件'}</span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={onUploadCheckedPerformanceFile} />
            </label>
        </div>

          <div className="actions-row">
            <button
              type="button"
              className="ghost"
              onClick={onCreateAndPushPerformanceWorkflow}
              disabled={isWorkflowSubmitting}
            >
              {isWorkflowSubmitting ? '提交中...' : '创建并推送核对'}
            </button>
          </div>

          {uploadedPerformanceRows.length > 0 && (
            <>
              <h3 className="table-title">绩效预览</h3>
              <div className="meta-row" style={{ marginBottom: 8 }}>
                <span className="meta-key">所有净利润总和：{normalizeMoney(previewTotals.netProfitTotal)}</span>
                <span className="meta-value">开票总额：{normalizeMoney(previewTotals.invoicedTotal)}</span>
              </div>
              <div className="detail-table-wrap" tabIndex={0}>
                <table className="detail-table">
                  <thead>
                    <tr>
                      {PUSH_PREVIEW_COLUMNS.map((column) => (
                        <th key={column.label}>{column.label}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {projectedUploadedPerformanceRows.slice(0, 200).map((row, index) => (
                      <tr key={`uploaded_performance_${index}`}>
                        {PUSH_PREVIEW_COLUMNS.map((column) => (
                          <td key={`${column.label}_${index}`}>{normalizeCellValue(row[column.label])}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <p className="tip">仅展示推送给员工的固定字段；汇总基于全部 {uploadedPerformanceRows.length} 行，表格展示前 200 行。</p>
            </>
          )}
        </>
      )}

      {!isPushMode && (
        <>
          <div className="actions-row">
            <button type="button" className="ghost" onClick={onRefreshWorkflowData} disabled={isWorkflowListLoading || isWorkflowSubmitting}>
              {isWorkflowListLoading ? '刷新中...' : '刷新数据'}
            </button>
          </div>
          {listTab === 'current' && (
            <>
              <h3 className="table-title">当月绩效（{currentPeriod}）</h3>
              {isWorkflowListLoading && <p className="hint-text">正在加载绩效数据，请稍候...</p>}
              <div className="admin-form-grid" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))' }}>
                <div className="control-row">
                  <label>状态筛选</label>
                  <select className="compact-select" value={currentStatusFilter} onChange={(event) => setCurrentStatusFilter(event.target.value)}>
                    <option value="all">全部状态</option>
                    {allStatuses.map((status) => (
                      <option key={status} value={status}>{workflowStatusLabelMap[status] || status}</option>
                    ))}
                  </select>
                </div>
              </div>
              <div className="admin-list">
                {filteredCurrentMonthWorkflows.length === 0 && <p className="hint-text">当月暂无匹配数据</p>}
                {filteredCurrentMonthWorkflows.map((item) => {
                  const status = normalizeCellValue(item.status)
                  const employeeName = normalizeCellValue(item.assignedToUserName) || normalizeCellValue(item.submittedByName) || '未指定员工'
                  return (
                    <div key={item.id} className="meta-row admin-list-item">
                      <span className="meta-key">员工：{employeeName} / 所属分公司：{item.subsidiaryLabel || '未分配分公司'} / 店铺名：{item.storeLabel}</span>
                      <span className="meta-value">状态：{workflowStatusLabelMap[status] || status} / 确认：{item.confirmedAt || '-'} / 归档：{item.archivedAt || '-'}</span>
                      {status === 'confirmed' && (
                        <button
                          type="button"
                          className="ghost"
                          onClick={() => onArchivePerformanceWorkflow(item.id)}
                          disabled={isWorkflowSubmitting}
                        >
                          落库存档
                        </button>
                      )}
                    </div>
                  )
                })}
              </div>
            </>
          )}

          {listTab === 'history' && (
            <>
              <h3 className="table-title">历史绩效（按月 / 按员工）</h3>
              {isWorkflowListLoading && <p className="hint-text">正在加载绩效数据，请稍候...</p>}
              <div className="admin-form-grid" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))' }}>
                <div className="control-row">
                  <label>月份筛选</label>
                  <input
                    type="month"
                    className="compact-input"
                    value={historyMonthFilter}
                    onChange={(event) => setHistoryMonthFilter(event.target.value)}
                    placeholder="全部月份"
                  />
                </div>
                <div className="control-row">
                  <label>员工筛选</label>
                  <select className="compact-select" value={historyEmployeeFilter} onChange={(event) => setHistoryEmployeeFilter(event.target.value)}>
                    <option value="">全部员工</option>
                    {historyEmployeeOptions.map((employee) => (
                      <option key={employee} value={employee}>{employee}</option>
                    ))}
                  </select>
                </div>
                <div className="control-row">
                  <label>筛选重置</label>
                  <button
                    type="button"
                    className="ghost"
                    onClick={() => {
                      setHistoryMonthFilter('')
                      setHistoryEmployeeFilter('')
                    }}
                    disabled={!historyMonthFilter && !historyEmployeeFilter}
                  >
                    清空筛选
                  </button>
                </div>
              </div>
              <div className="admin-list">
                {historyByPeriodEmployee.length === 0 && <p className="hint-text">暂无历史绩效</p>}
                {historyByPeriodEmployee.map((periodGroup) => (
                  <section key={periodGroup.period} className="admin-list-item">
                    <div className="meta-row">
                      <span className="meta-key">{periodGroup.period}</span>
                      <span className="meta-value">流程 {periodGroup.monthWorkflowCount} / 店铺覆盖 {periodGroup.monthStoreCount} / 总业绩 {periodGroup.monthTotalProfit}</span>
                    </div>
                    <div className="detail-table-wrap" tabIndex={0} style={{ marginTop: 8 }}>
                      <table className="detail-table">
                        <thead>
                          <tr>
                            <th>员工</th>
                            <th>流程数</th>
                            <th>已归档</th>
                            <th>待归档</th>
                            <th>店铺数</th>
                            <th>绩效行数</th>
                            <th>总业绩</th>
                          </tr>
                        </thead>
                        <tbody>
                          {periodGroup.rows.map((row) => (
                            <tr key={`${periodGroup.period}_${row.employeeName}`}>
                              <td>{row.employeeName}</td>
                              <td>{row.workflowCount}</td>
                              <td>{row.archivedCount}</td>
                              <td>{row.confirmedCount}</td>
                              <td>{row.storeCount}</td>
                              <td>{row.rowCount}</td>
                              <td>{row.totalProfit}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </section>
                ))}
              </div>
            </>
          )}
        </>
      )}
    </section>
  )
}

import { normalizeCellValue, normalizeMoney, toNumericValue } from '../calculatorDomain'
import type { PerformanceWorkflowItem } from '../types/workflow'

const WORKFLOW_DISPLAY_COLUMNS: Array<{ label: string; aliases: string[] }> = [
  { label: '订单号', aliases: ['订单号', '订单编号', '平台订单号', '订单id', '订单ID'] },
  { label: '订单时间', aliases: ['订单时间', '下单时间', '付款时间', '创建时间'] },
  { label: '订单状态', aliases: ['订单状态', '状态'] },
  { label: '订单预计可得', aliases: ['订单预计可得', '预计可得', '预计可得金额'] },
  { label: '总物流费用', aliases: ['总物流费用', '物流费用', '总物流费'] },
  { label: '采购费用', aliases: ['采购费用', '采购成本', '总采购费用'] },
  { label: '支付宝是否开发票', aliases: ['支付宝是否开发票', '是否开票', '支付宝开票'] },
  { label: '净利润', aliases: ['净利润', '总利润', '总收支'] }
]

function pickRowValue(row: Record<string, unknown>, aliases: string[]): string {
  for (const key of aliases) {
    if (row[key] !== undefined) {
      return normalizeCellValue(row[key])
    }
  }
  return ''
}

function toDisplayRows(rows: Array<Record<string, unknown>>): Array<Record<string, string>> {
  return rows.map((row) => {
    const output: Record<string, string> = {}
    for (const column of WORKFLOW_DISPLAY_COLUMNS) {
      output[column.label] = pickRowValue(row, column.aliases)
    }
    return output
  })
}

type MyPerformanceViewProps = {
  pendingWorkflows: PerformanceWorkflowItem[]
  historyWorkflows: PerformanceWorkflowItem[]
  workflowStatusLabelMap: Record<string, string>
  isSubmitting: boolean
  onConfirmWorkflow: (workflowId: string) => void
}

function calculateWorkflowTotalProfit(workflow: PerformanceWorkflowItem): number {
  const candidateKeys = ['总利润', '净利润', '总收支']
  return normalizeMoney(
    workflow.uploadedRows.reduce((sum, row) => {
      const key = candidateKeys.find((candidate) => row[candidate] !== undefined)
      if (!key) {
        return sum
      }
      return sum + toNumericValue(row[key])
    }, 0)
  )
}

function WorkflowBlock(props: {
  workflow: PerformanceWorkflowItem
  statusLabel: string
  showConfirm: boolean
  isSubmitting: boolean
  onConfirmWorkflow: (workflowId: string) => void
}) {
  const { workflow, statusLabel, showConfirm, isSubmitting, onConfirmWorkflow } = props
  const displayRows = toDisplayRows(workflow.uploadedRows as Array<Record<string, unknown>>)
  const expectedIncomeHeaderHint = '预计可得口径未扣除：物流上网超时处罚、违背发货承诺处罚；仅在净利润中扣除。'

  return (
    <section className="admin-card admin-card-full">
      <div className="meta-row">
        <span className="meta-key">店铺：{workflow.storeLabel}（{workflow.period} / {statusLabel}）</span>
        <span className="meta-value">总业绩（总利润）：{calculateWorkflowTotalProfit(workflow)}</span>
        {showConfirm ? (
          <button
            type="button"
            className="ghost"
            onClick={() => onConfirmWorkflow(workflow.id)}
            disabled={isSubmitting}
          >
            核对无误
          </button>
        ) : (
          <span className="meta-value">{workflow.archivedAt ? `归档：${workflow.archivedAt}` : `确认：${workflow.confirmedAt || '-'}`}</span>
        )}
      </div>

      {displayRows.length > 0 ? (
        <>
          <div className="detail-table-wrap" tabIndex={0}>
            <table className="detail-table">
              <thead>
                <tr>
                  {WORKFLOW_DISPLAY_COLUMNS.map((column) => (
                    <th key={column.label}>
                      <span className="table-header-label">{column.label}</span>
                      {column.label === '订单预计可得' && (
                        <span className="table-header-hint-wrap">
                          <span
                            className="table-header-hint"
                            aria-label={expectedIncomeHeaderHint}
                            tabIndex={0}
                          >
                            ?
                          </span>
                          <span className="table-header-tooltip" role="tooltip">
                            {expectedIncomeHeaderHint}
                          </span>
                        </span>
                      )}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {displayRows.slice(0, 200).map((row, index) => (
                  <tr key={`${workflow.id}_${index}`}>
                    {WORKFLOW_DISPLAY_COLUMNS.map((column) => (
                      <td key={`${workflow.id}_${column.label}_${index}`}>{normalizeCellValue(row[column.label])}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="tip">已展示前 200 行，总计 {displayRows.length} 行。</p>
        </>
      ) : (
        <p className="hint-text">该店铺暂无可展示明细行。</p>
      )}
    </section>
  )
}

export function MyPerformanceView(props: MyPerformanceViewProps) {
  const { pendingWorkflows, historyWorkflows, workflowStatusLabelMap, isSubmitting, onConfirmWorkflow } = props

  return (
    <section className="upload-card">
      <h2>我的业绩</h2>
      <p>上方为待复核绩效，下方为历史业绩记录。</p>

      <h3 className="table-title">待复核业绩表</h3>
      {pendingWorkflows.length === 0 && <p className="hint-text">暂无待核对绩效</p>}
      {pendingWorkflows.map((item) => (
        <WorkflowBlock
          key={item.id}
          workflow={item}
          statusLabel={workflowStatusLabelMap[item.status] || item.status}
          showConfirm
          isSubmitting={isSubmitting}
          onConfirmWorkflow={onConfirmWorkflow}
        />
      ))}

      <h3 className="table-title">历史业绩</h3>
      {historyWorkflows.length === 0 && <p className="hint-text">暂无历史业绩</p>}
      {historyWorkflows.map((item) => (
        <WorkflowBlock
          key={item.id}
          workflow={item}
          statusLabel={workflowStatusLabelMap[item.status] || item.status}
          showConfirm={false}
          isSubmitting={isSubmitting}
          onConfirmWorkflow={onConfirmWorkflow}
        />
      ))}
    </section>
  )
}

import type * as Domain from '../calculatorDomain'
import { normalizeCellValue, normalizeOrderNo } from '../calculatorDomain'

type ProcessResult = Domain.ProcessResult

type StoreMetaSectionProps = {
  shopId: string
  shopName: string
  subsidiary: string
  usdExchangeRate: string
  onShopIdChange: (value: string) => void
  onShopNameChange: (value: string) => void
  onSubsidiaryChange: (value: string) => void
  onUsdExchangeRateChange: (value: string) => void
}

export function StoreMetaSection(props: StoreMetaSectionProps) {
  const {
    shopId,
    shopName,
    subsidiary,
    usdExchangeRate,
    onShopIdChange,
    onShopNameChange,
    onSubsidiaryChange,
    onUsdExchangeRateChange
  } = props

  return (
    <section className="upload-card store-meta-card">
      <h2>店铺信息</h2>
      <p>用于导出文件命名：所属分公司_店铺名_xxx表。</p>
      <div className="store-meta-grid">
        <div className="control-row">
          <label>店铺ID</label>
          <input
            className="compact-input"
            value={shopId}
            onChange={(event) => onShopIdChange(event.target.value)}
            placeholder="例如：A12345"
          />
        </div>
        <div className="control-row">
          <label>店铺名</label>
          <input
            className="compact-input"
            value={shopName}
            onChange={(event) => onShopNameChange(event.target.value)}
            placeholder="例如：DXM官方店"
          />
        </div>
        <div className="control-row">
          <label>所属分公司（可选）</label>
          <select
            className="compact-select"
            value={subsidiary}
            onChange={(event) => onSubsidiaryChange(event.target.value)}
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
        <div className="control-row">
          <label>美元汇率（税费核对）</label>
          <input
            className="compact-input"
            value={usdExchangeRate}
            onChange={(event) => onUsdExchangeRateChange(event.target.value)}
            placeholder="例如：6.8"
          />
        </div>
      </div>
    </section>
  )
}

type ActionPanelProps = {
  isProcessing: boolean
  isUploadingToSaas: boolean
  canProcess: boolean
  result: ProcessResult | null
  lastCalculatedAt: string
  calculationCount: number
  processDisabledReason: string
  onRunCalculation: () => void
  onUploadToSaas: () => void
  onExportAggregated: () => void
  onExportBusinessDetail: () => void
  onExportOtherSheets: () => void
}

export function ActionPanel(props: ActionPanelProps) {
  const {
    isProcessing,
    isUploadingToSaas,
    canProcess,
    result,
    lastCalculatedAt,
    calculationCount,
    processDisabledReason,
    onRunCalculation,
    onUploadToSaas,
    onExportAggregated,
    onExportBusinessDetail,
    onExportOtherSheets
  } = props

  return (
    <section className="action-panel">
      <button type="button" onClick={onRunCalculation} disabled={!canProcess || isProcessing}>
        {isProcessing ? '计算中...' : '计算业绩'}
      </button>
      <button type="button" className="ghost" onClick={onExportAggregated} disabled={!result}>
        导出订单聚合表
      </button>
      <button type="button" className="ghost" onClick={onExportBusinessDetail} disabled={!result}>
        导出业务明细表
      </button>
      <button type="button" className="ghost" onClick={onExportOtherSheets} disabled={!result}>
        导出其他对账Sheet
      </button>
      <button type="button" className="ghost warn" onClick={onUploadToSaas} disabled={!result || isUploadingToSaas}>
        {isUploadingToSaas ? '上传中...' : '上传最终绩效到SaaS'}
      </button>
      {result && lastCalculatedAt && (
        <div className="meta-row">
          <span className="meta-key">最近重新生成</span>
          <span className="meta-value">第 {calculationCount} 次（{lastCalculatedAt}）</span>
        </div>
      )}
      {!isProcessing && !canProcess && (
        <div className="meta-row">
          <span className="meta-key">当前不可执行原因</span>
          <span className="meta-value">{processDisabledReason}</span>
        </div>
      )}
    </section>
  )
}

type ResultPreviewPanelProps = {
  result: ProcessResult
}

export function ResultPreviewPanel(props: ResultPreviewPanelProps) {
  const { result } = props

  const handleTableCopyShortcut = (event: React.KeyboardEvent<HTMLDivElement>) => {
    if (!(event.metaKey || event.ctrlKey) || event.key.toLowerCase() !== 'c') {
      return
    }

    const selectedText = globalThis.getSelection?.()?.toString() || ''
    if (!selectedText) {
      return
    }

    event.preventDefault()
    void globalThis.navigator?.clipboard?.writeText(selectedText)
  }

  return (
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
          <h3>净利润（扣运费+采购）</h3>
          <p>{result.summary.totalIncomeAfterOnlineFreight}</p>
        </article>
        <article>
          <h3>开票金额合计</h3>
          <p>{result.summary.totalInvoicedPurchaseAmount}</p>
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
          <div className="meta-row">
            <span className="meta-key">线下运费类别列</span>
            <span className="meta-value">
              {result.dynamicOfflineCategoryColumns.join('、') || '无'}
            </span>
          </div>
        </>
      )}
      <div className="detail-table-wrap" tabIndex={0} onKeyDown={handleTableCopyShortcut}>
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
  )
}
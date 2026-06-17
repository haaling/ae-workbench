import type { ReactNode } from 'react'
import * as XLSX from 'xlsx'
import type * as Domain from '../calculatorDomain'

type MultiUploadItem = Domain.MultiUploadItem
type IncomeUploadItem = Domain.IncomeUploadItem

type FileUploadHandler = (
  event: React.ChangeEvent<HTMLInputElement>,
  table: Domain.TableKind
) => Promise<void>

type OrdersUploadCardProps = {
  files: MultiUploadItem[]
  onFileUpload: FileUploadHandler
  onRemove: (itemId: string) => void
}

export function OrdersUploadCard(props: OrdersUploadCardProps) {
  const { files, onFileUpload, onRemove } = props

  return (
    <section className="upload-card">
      <h2 className="upload-title-row">
        <span>1) 订单表上传</span>
        <span className="upload-count-pill">已上传 {files.length} 份</span>
      </h2>
      <p className="upload-card-subtitle">选择结算月份的所有订单导出。</p>

      <label className="file-input-label">
        <span>批量上传订单表（可多选）</span>
        <input
          type="file"
          accept=".xlsx,.xls,.csv"
          multiple
          onChange={(event) => void onFileUpload(event, 'orders')}
        />
      </label>

      <div className="uploaded-files-list">
        {files.length === 0 && (
          <div className="meta-row upload-empty-state">
            <span className="meta-key">状态</span>
            <span className="meta-value">尚未上传订单文件</span>
          </div>
        )}

        {files.map((item) => (
          <article key={item.id} className="income-file-item">
            <div className="income-file-head">
              <div className="file-name-row">
                <strong>{item.file.fileName}</strong>
              </div>
              <button type="button" className="mini-danger" onClick={() => onRemove(item.id)}>
                删除
              </button>
            </div>
          </article>
        ))}
      </div>
    </section>
  )
}

type DetailUploadCardProps = {
  title: string
  description: ReactNode
  table: 'income' | 'refund'
  files: IncomeUploadItem[]
  emptyText: string
  onFileUpload: FileUploadHandler
  onRemove: (itemId: string, table: 'income' | 'refund') => void
}

export function DetailUploadCard(props: DetailUploadCardProps) {
  const { title, description, table, files, emptyText, onFileUpload, onRemove } = props

  return (
    <section className="upload-card">
      <h2 className="upload-title-row">
        <span>{title}</span>
        <span className="upload-count-pill">已上传 {files.length} 份</span>
      </h2>
      <p className="upload-card-subtitle">{description}</p>

      <label className="file-input-label">
        <span>批量上传（可多选）</span>
        <input
          type="file"
          accept=".xlsx,.xls,.csv"
          multiple
          onChange={(event) => void onFileUpload(event, table)}
        />
      </label>

      <div className="uploaded-files-list">
        {files.length === 0 && (
          <div className="meta-row upload-empty-state">
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
              <button type="button" className="mini-danger" onClick={() => onRemove(item.id, table)}>
                删除
              </button>
            </div>
          </article>
        ))}
      </div>
    </section>
  )
}

type FreightUploadCardProps = {
  files: MultiUploadItem[]
  onFileUpload: FileUploadHandler
  onRemove: (itemId: string) => void
}

export function FreightUploadCard(props: FreightUploadCardProps) {
  const { files, onFileUpload, onRemove } = props

  return (
    <section className="upload-card">
      <h2 className="upload-title-row">
        <span>2) 金掌柜运费表（可选）</span>
        <span className="upload-count-pill">已上传 {files.length} 份</span>
      </h2>
      <p className="upload-card-subtitle">可为空。一次可上传多个运费文件，系统按交易单号字段匹配订单号并汇总；未上传时按 0 处理。</p>

      <label className="file-input-label">
        <span>批量上传金掌柜运费表（可多选）</span>
        <input
          type="file"
          accept=".xlsx,.xls,.csv"
          multiple
          onChange={(event) => void onFileUpload(event, 'freight')}
        />
      </label>

      <div className="uploaded-files-list">
        {files.length === 0 && (
          <div className="meta-row upload-empty-state">
            <span className="meta-key">状态</span>
            <span className="meta-value">未上传运费文件（可不上传）</span>
          </div>
        )}

        {files.map((item) => (
          <article key={item.id} className="income-file-item">
            <div className="income-file-head">
              <div className="file-name-row">
                <strong>{item.file.fileName}</strong>
              </div>
              <button type="button" className="mini-danger" onClick={() => onRemove(item.id)}>
                删除
              </button>
            </div>
          </article>
        ))}
      </div>
    </section>
  )
}

type SimpleUploadCardProps = {
  title: string
  description: string
  table: 'alipay' | 'offline'
  files: MultiUploadItem[]
  emptyText: string
  onFileUpload: FileUploadHandler
  onRemove: (itemId: string, table: 'alipay' | 'offline') => void
}

export function SimpleUploadCard(props: SimpleUploadCardProps) {
  const { title, description, table, files, emptyText, onFileUpload, onRemove } = props

  const downloadTemplate = () => {
    const workbook = XLSX.utils.book_new()

    if (table === 'alipay') {
      const rows = [
        {
          金额: 120.5,
          备注: '示例店铺-1234567890123456'
        }
      ]
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), '支付宝采购模板')
      XLSX.writeFile(workbook, '支付宝订单记录_填写模板.xlsx')
      return
    }

    const rows = [
      {
        客户订单号: '示例店铺-1234567890123456',
        金额: 18.6
      }
    ]
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), '线下发货模板')
    XLSX.writeFile(workbook, '线下发货订单记录_填写模板.xlsx')
  }

  return (
    <section className="upload-card">
      <h2 className="upload-title-row">
        <span>{title}</span>
        <span className="upload-title-actions">
          <button type="button" className="upload-count-pill upload-pill-button" onClick={downloadTemplate}>
            下载填写模板
          </button>
          <span className="upload-count-pill">已上传 {files.length} 份</span>
        </span>
      </h2>
      <p className="upload-card-subtitle">{description}</p>

      <label className="file-input-label">
        <span>批量上传（可多选）</span>
        <input
          type="file"
          accept=".xlsx,.xls,.csv"
          multiple
          onChange={(event) => void onFileUpload(event, table)}
        />
      </label>

      <div className="uploaded-files-list">
        {files.length === 0 && (
          <div className="meta-row upload-empty-state">
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
              <button type="button" className="mini-danger" onClick={() => onRemove(item.id, table)}>
                删除
              </button>
            </div>
          </article>
        ))}
      </div>
    </section>
  )
}
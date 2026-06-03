import type { ReactNode } from 'react'
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
      <p>选择结算月份的所有订单导出。</p>

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
          <div className="meta-row">
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
      <p>{description}</p>

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
        <span>2) 金掌柜运费表</span>
        <span className="upload-count-pill">已上传 {files.length} 份</span>
      </h2>
      <p>一次可上传多个运费文件，系统按交易单号字段匹配订单号并汇总。</p>

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
          <div className="meta-row">
            <span className="meta-key">状态</span>
            <span className="meta-value">尚未上传运费文件</span>
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
          onChange={(event) => void onFileUpload(event, table)}
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
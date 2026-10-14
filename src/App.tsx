import { useMemo, useState } from 'react'
import { saveAs } from 'file-saver'
import * as XLSX from 'xlsx'
import './App.css'

type ProcessMode = 'all' | 'selected'

type SheetResult = {
  sheetName: string
  rowsFilled: number
  targetCols: number[]
  skipped: boolean
  reason?: string
}

type TitleState = {
  prefix: string
  usedColors: Set<string>
}

const COLOR_ABBRS = [
  'BK',
  'WH',
  'RD',
  'BU',
  'GN',
  'YL',
  'PK',
  'GY',
  'BN',
  'OR',
  'PP',
  'NV',
  'BG',
  'CY',
  'KH'
]

function normalizeHeaderText(text: unknown): string {
  return String(text ?? '')
    .replace(/[\s_*\-()（）]+/g, '')
    .trim()
    .toLowerCase()
}

function normalizeTitle(value: unknown): string {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim()
}

function randomBase36(length: number): string {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
  let result = ''
  for (let i = 0; i < length; i += 1) {
    result += chars[Math.floor(Math.random() * chars.length)]
  }
  return result
}

function pickColorAbbr(usedSet: Set<string>): string {
  const available = COLOR_ABBRS.filter((abbr) => !usedSet.has(abbr))
  const source = available.length > 0 ? available : COLOR_ABBRS
  return source[Math.floor(Math.random() * source.length)]
}

function isSkuLikeSecondHeader(normalized: string): boolean {
  if (!normalized) {
    return false
  }

  return ['sku属性', 'sku编码', 'skucode'].some((token) => normalized.includes(token))
}

function findSkuTargetColumns(rows: unknown[][]): number[] {
  const row1 = rows[0] ?? []
  const row2 = rows[1] ?? []
  const maxCols = Math.max(row1.length, row2.length)
  const result: number[] = []

  for (let c = 0; c < maxCols; c += 1) {
    const h1 = normalizeHeaderText(row1[c])
    const h2 = normalizeHeaderText(row2[c])
    const strict = h1 === 'sku信息' && h2 === 'sku属性'
    const tolerant = h1 === 'sku信息' && isSkuLikeSecondHeader(h2)
    if (strict || tolerant) {
      result.push(c)
    }
  }

  return result
}

function findTitleColumn(rows: unknown[][]): number {
  const row1 = rows[0] ?? []
  const row2 = rows[1] ?? []
  const maxCols = Math.max(row1.length, row2.length)

  for (let c = 0; c < maxCols; c += 1) {
    const h1 = String(row1[c] ?? '').trim()
    const h2 = String(row2[c] ?? '').trim()
    const merged = `${h1} ${h2}`.toLowerCase()
    if (/商品\s*标题/.test(merged) || /^标题$/.test(h1) || /^标题$/.test(h2)) {
      return c
    }
  }

  for (let c = 0; c < maxCols; c += 1) {
    const h1 = String(row1[c] ?? '').trim().toLowerCase()
    const h2 = String(row2[c] ?? '').trim().toLowerCase()
    const merged = `${h1} ${h2}`
    if (/product\s*title|title/.test(merged)) {
      return c
    }
  }

  return -1
}

function setCellValue(ws: XLSX.WorkSheet, rowIndex: number, colIndex: number, value: string): void {
  const addr = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
  const cell = ws[addr] as XLSX.CellObject | undefined
  if (!cell) {
    ws[addr] = { t: 's', v: value }
    return
  }

  cell.t = 's'
  cell.v = value
  if ('w' in cell) {
    delete cell.w
  }
}

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null)
  const [mode, setMode] = useState<ProcessMode>('selected')
  const [selectedSheets, setSelectedSheets] = useState<string[]>([])
  const [isProcessing, setIsProcessing] = useState(false)
  const [errorMessage, setErrorMessage] = useState('')
  const [successMessage, setSuccessMessage] = useState('')
  const [results, setResults] = useState<SheetResult[]>([])

  const sheetNames = useMemo(() => workbook?.SheetNames ?? [], [workbook])

  const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const currentFile = event.target.files?.[0] ?? null
    setFile(currentFile)
    setWorkbook(null)
    setSelectedSheets([])
    setResults([])
    setErrorMessage('')
    setSuccessMessage('')

    if (!currentFile) {
      return
    }

    try {
      const buffer = await currentFile.arrayBuffer()
      const wb = XLSX.read(buffer, { type: 'array' })
      setWorkbook(wb)
      setSelectedSheets(wb.SheetNames.slice(0, 5))
    } catch (error) {
      setErrorMessage(`读取 Excel 失败：${String((error as Error)?.message || error)}`)
    }
  }

  const toggleSheet = (sheetName: string) => {
    setSelectedSheets((prev) => {
      if (prev.includes(sheetName)) {
        return prev.filter((item) => item !== sheetName)
      }
      return [...prev, sheetName]
    })
  }

  const processExcel = () => {
    if (!workbook || !file) {
      setErrorMessage('请先上传 Excel 文件。')
      return
    }

    const targetSheets = mode === 'all' ? workbook.SheetNames : selectedSheets
    if (targetSheets.length === 0) {
      setErrorMessage('请至少选择一个 Sheet。')
      return
    }

    setIsProcessing(true)
    setErrorMessage('')
    setSuccessMessage('')
    setResults([])

    try {
      const titleMap = new Map<string, TitleState>()
      const usedPrefixes = new Set<string>()
      const resultList: SheetResult[] = []

      targetSheets.forEach((sheetName) => {
        const ws = workbook.Sheets[sheetName]
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as unknown[][]

        if (rows.length < 2) {
          resultList.push({
            sheetName,
            rowsFilled: 0,
            targetCols: [],
            skipped: true,
            reason: '表头不足 2 行'
          })
          return
        }

        const targetCols = findSkuTargetColumns(rows)
        if (targetCols.length === 0) {
          resultList.push({
            sheetName,
            rowsFilled: 0,
            targetCols: [],
            skipped: true,
            reason: '未找到表头：第一行 SKU信息 且第二行 SKU属性/SKU编码'
          })
          return
        }

        const titleCol = findTitleColumn(rows)
        if (titleCol < 0) {
          resultList.push({
            sheetName,
            rowsFilled: 0,
            targetCols,
            skipped: true,
            reason: '未识别到商品标题列'
          })
          return
        }

        let filledRows = 0
        for (let r = 2; r < rows.length; r += 1) {
          const row = rows[r] ?? []
          const title = normalizeTitle(row[titleCol])
          if (!title) {
            continue
          }

          if (!titleMap.has(title)) {
            let prefix = randomBase36(12)
            while (usedPrefixes.has(prefix)) {
              prefix = randomBase36(12)
            }
            usedPrefixes.add(prefix)
            titleMap.set(title, {
              prefix,
              usedColors: new Set<string>()
            })
          }

          const titleState = titleMap.get(title)
          if (!titleState) {
            continue
          }

          const colorAbbr = pickColorAbbr(titleState.usedColors)
          titleState.usedColors.add(colorAbbr)

          let skuValue = `${titleState.prefix}_${colorAbbr}`
          if (skuValue.length > 20) {
            skuValue = skuValue.slice(0, 20)
          }

          targetCols.forEach((col) => {
            setCellValue(ws, r, col, skuValue)
          })

          filledRows += 1
        }

        resultList.push({
          sheetName,
          rowsFilled: filledRows,
          targetCols,
          skipped: false
        })
      })

      setResults(resultList)

      const okCount = resultList.filter((item) => !item.skipped).length
      if (okCount === 0) {
        setErrorMessage('处理完成，但没有可写入的 Sheet，请检查表头格式。')
        return
      }

      const outputName = file.name.replace(/\.(xlsx|xls)$/i, '') + '_sku_filled.xlsx'
      const outputArray = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' })
      saveAs(
        new Blob([outputArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        outputName
      )
      setSuccessMessage(`处理完成：成功 ${okCount} 个 Sheet，文件已下载：${outputName}`)
    } catch (error) {
      setErrorMessage(`处理失败：${String((error as Error)?.message || error)}`)
    } finally {
      setIsProcessing(false)
    }
  }

  return (
    <main className="page-shell">
      <section className="tool-card">
        <div className="head-block">
          <h1>Excel SKU 批量处理工具</h1>
          <p>上传 Excel 后，按商品标题分组生成 SKU：12 位随机前缀 + 下划线 + 颜色缩写，总长度限制为 20。</p>
        </div>

        <div className="form-row">
          <label htmlFor="excel-file">Excel 文件</label>
          <input id="excel-file" type="file" accept=".xlsx,.xls" onChange={onFileChange} />
        </div>

        {sheetNames.length > 0 && (
          <>
            <div className="form-row">
              <label>处理模式</label>
              <div className="radio-group">
                <label>
                  <input
                    type="radio"
                    name="mode"
                    checked={mode === 'selected'}
                    onChange={() => setMode('selected')}
                  />
                  指定 Sheet
                </label>
                <label>
                  <input type="radio" name="mode" checked={mode === 'all'} onChange={() => setMode('all')} />
                  全部 Sheet
                </label>
              </div>
            </div>

            <div className="form-row">
              <label>Sheet 列表</label>
              <div className="sheet-grid">
                {sheetNames.map((name) => (
                  <label key={name} className="sheet-item">
                    <input
                      type="checkbox"
                      checked={selectedSheets.includes(name)}
                      onChange={() => toggleSheet(name)}
                      disabled={mode === 'all'}
                    />
                    <span>{name}</span>
                  </label>
                ))}
              </div>
            </div>
          </>
        )}

        <div className="action-row">
          <button type="button" onClick={processExcel} disabled={!workbook || isProcessing}>
            {isProcessing ? '处理中...' : '开始处理并下载'}
          </button>
        </div>

        {errorMessage && <div className="msg error">{errorMessage}</div>}
        {successMessage && <div className="msg success">{successMessage}</div>}

        {results.length > 0 && (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Sheet</th>
                  <th>状态</th>
                  <th>填充行数</th>
                  <th>目标列（1-based）</th>
                  <th>说明</th>
                </tr>
              </thead>
              <tbody>
                {results.map((item) => (
                  <tr key={item.sheetName}>
                    <td>{item.sheetName}</td>
                    <td>{item.skipped ? '跳过' : '成功'}</td>
                    <td>{item.rowsFilled}</td>
                    <td>{item.targetCols.map((col) => col + 1).join(', ') || '-'}</td>
                    <td>{item.reason || '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>
    </main>
  )
}

export default App

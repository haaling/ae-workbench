const fs = require('fs')
const path = require('path')
const XLSX = require('xlsx')

const DEFAULTS = {
  orders: '/Users/haaling/Downloads/aliexpress-exports/2026-02_haaling/haaling_202602_store_订单列表.xlsx',
  alipay:
    '/Users/haaling/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/LINDA955_4427/msg/file/2026-05/支付宝交易明细(20260429-20260529).csv',
  output:
    '/Users/haaling/dxm/performance-calculator-web/test-data/支付宝交易明细(20260429-20260529)_测试已加备注_UTF8.csv',
  shopName: 'DXM官方店',
  keepBlankEvery: 10,
  keepUnmatchedEvery: 9,
  matchedPoolRatio: 0.7
}

function parseArgs(argv) {
  const config = { ...DEFAULTS }

  for (let i = 0; i < argv.length; i += 1) {
    const current = argv[i]
    const next = argv[i + 1]

    if (current === '--orders' && next) {
      config.orders = next
      i += 1
      continue
    }
    if (current === '--alipay' && next) {
      config.alipay = next
      i += 1
      continue
    }
    if (current === '--output' && next) {
      config.output = next
      i += 1
      continue
    }
    if (current === '--shop' && next) {
      config.shopName = next
      i += 1
      continue
    }
    if (current === '--blankEvery' && next) {
      config.keepBlankEvery = Math.max(0, Number(next) || 0)
      i += 1
      continue
    }
    if (current === '--unmatchedEvery' && next) {
      config.keepUnmatchedEvery = Math.max(0, Number(next) || 0)
      i += 1
      continue
    }
    if (current === '--matchedRatio' && next) {
      const ratio = Number(next)
      config.matchedPoolRatio = Number.isFinite(ratio) ? Math.min(1, Math.max(0.05, ratio)) : DEFAULTS.matchedPoolRatio
      i += 1
      continue
    }
    if (current === '--help') {
      printHelpAndExit()
    }
  }

  return config
}

function printHelpAndExit() {
  const message = [
    'Usage: node scripts/generate-alipay-test-data.cjs [options]',
    '',
    'Options:',
    '  --orders <path>         AliExpress 订单 Excel 路径',
    '  --alipay <path>         支付宝交易明细 CSV 路径',
    '  --output <path>         输出 CSV 路径 (UTF-8 BOM)',
    '  --shop <name>           备注店铺名前缀，默认 DXM官方店',
    '  --blankEvery <num>      每 N 行保留空备注，默认 10',
    '  --unmatchedEvery <num>  每 N 行写入不匹配备注，默认 9',
    '  --matchedRatio <0-1>    匹配订单池占比，默认 0.7'
  ].join('\n')

  console.log(message)
  process.exit(0)
}

function normalize(value) {
  return String(value ?? '').trim()
}

function normalizeOrderNo(value) {
  const text = normalize(value)
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

  return sanitized
}

function parseCsvLine(line) {
  const result = []
  let current = ''
  let inQuotes = false

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index]
    if (char === '"') {
      if (inQuotes && line[index + 1] === '"') {
        current += '"'
        index += 1
      } else {
        inQuotes = !inQuotes
      }
      continue
    }

    if (char === ',' && !inQuotes) {
      result.push(current)
      current = ''
      continue
    }

    current += char
  }

  result.push(current)
  return result
}

function toCsvLine(columns) {
  return columns
    .map((column) => {
      const text = String(column ?? '')
      if (/[",\n\r]/.test(text)) {
        return `"${text.replace(/"/g, '""')}"`
      }
      return text
    })
    .join(',')
}

function decodeAlipayContent(buffer) {
  try {
    return new TextDecoder('gb18030').decode(buffer)
  } catch (error) {
    return new TextDecoder('gbk').decode(buffer)
  }
}

function extractOrderIds(ordersPath) {
  const workbook = XLSX.readFile(ordersPath, { cellDates: false })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false })

  if (!rows.length) {
    throw new Error('订单表没有可用数据行。')
  }

  const headers = Object.keys(rows[0])
  const orderColumn =
    headers.find((header) => {
      const lowered = normalize(header).toLowerCase()
      return lowered.includes('订单号') || lowered.includes('订单编号') || lowered.includes('order')
    }) || headers[0]

  const orderIds = Array.from(new Set(rows.map((row) => normalizeOrderNo(row[orderColumn])).filter(Boolean)))
  if (!orderIds.length) {
    throw new Error('订单表未提取到可用订单号。')
  }

  return { orderColumn, orderIds }
}

function injectRemarks(config, content, orderIds) {
  const lines = content.replace(/^\uFEFF/, '').split(/\r?\n/)
  let firstDataIndex = -1

  for (let index = 0; index < lines.length; index += 1) {
    if (/^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2},/.test(lines[index])) {
      firstDataIndex = index
      break
    }
  }

  if (firstDataIndex < 0) {
    throw new Error('支付宝文件中未检测到交易数据行。')
  }

  const outputLines = [...lines]
  const matchedPoolSize = Math.max(1, Math.floor(orderIds.length * config.matchedPoolRatio))

  let dataRows = 0
  let matched = 0
  let unmatched = 0
  let blanks = 0

  for (let index = firstDataIndex; index < lines.length; index += 1) {
    const line = lines[index]
    if (!/^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2},/.test(line)) {
      continue
    }

    const columns = parseCsvLine(line)
    if (columns.length < 3) {
      continue
    }

    const remarkIndex = columns.length - 2
    let remark = ''

    const hitBlankRule = config.keepBlankEvery > 0 && dataRows % config.keepBlankEvery === 0
    const hitUnmatchedRule = config.keepUnmatchedEvery > 0 && dataRows % config.keepUnmatchedEvery === 0

    if (hitBlankRule) {
      remark = ''
      blanks += 1
    } else if (hitUnmatchedRule) {
      remark = `${config.shopName}-AE_UNMATCH_${String(dataRows).padStart(4, '0')}`
      unmatched += 1
    } else {
      const orderNo = orderIds[dataRows % matchedPoolSize]
      remark = `${config.shopName}-${orderNo}`
      matched += 1
    }

    columns[remarkIndex] = remark
    outputLines[index] = toCsvLine(columns)
    dataRows += 1
  }

  return {
    outputText: `\uFEFF${outputLines.join('\n')}`,
    summary: {
      dataRows,
      matched,
      unmatched,
      blanks,
      matchedPoolSize
    }
  }
}

function main() {
  const config = parseArgs(process.argv.slice(2))

  if (!fs.existsSync(config.orders)) {
    throw new Error(`订单文件不存在: ${config.orders}`)
  }
  if (!fs.existsSync(config.alipay)) {
    throw new Error(`支付宝文件不存在: ${config.alipay}`)
  }

  const { orderColumn, orderIds } = extractOrderIds(config.orders)
  const buffer = fs.readFileSync(config.alipay)
  const decoded = decodeAlipayContent(buffer)

  const { outputText, summary } = injectRemarks(config, decoded, orderIds)
  fs.mkdirSync(path.dirname(config.output), { recursive: true })
  fs.writeFileSync(config.output, outputText, 'utf8')

  console.log('生成完成')
  console.log('output:', config.output)
  console.log('order column:', orderColumn)
  console.log('orders:', orderIds.join(', '))
  console.log('rows:', summary.dataRows)
  console.log('matched:', summary.matched)
  console.log('unmatched:', summary.unmatched)
  console.log('blank:', summary.blanks)
}

try {
  main()
} catch (error) {
  console.error('生成失败:', error instanceof Error ? error.message : error)
  process.exit(1)
}
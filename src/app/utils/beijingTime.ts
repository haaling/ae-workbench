import { normalizeCellValue } from '../calculatorDomain'

export function formatBeijingDateTime(value: unknown): string {
  if (!value) {
    return ''
  }

  const date = value instanceof Date ? value : new Date(String(value))
  if (Number.isNaN(date.getTime())) {
    return normalizeCellValue(value)
  }

  const parts = new Intl.DateTimeFormat('zh-CN', {
    timeZone: 'Asia/Shanghai',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  }).formatToParts(date)

  const read = (type: Intl.DateTimeFormatPartTypes) => parts.find((part) => part.type === type)?.value || ''
  return `${read('year')}-${read('month')}-${read('day')} ${read('hour')}:${read('minute')}:${read('second')}`
}

export function getCurrentBeijingPeriod(): string {
  const now = new Date()
  const parts = new Intl.DateTimeFormat('zh-CN', {
    timeZone: 'Asia/Shanghai',
    year: 'numeric',
    month: '2-digit'
  }).formatToParts(now)
  const year = parts.find((part) => part.type === 'year')?.value || String(now.getFullYear())
  const month = parts.find((part) => part.type === 'month')?.value || String(now.getMonth() + 1).padStart(2, '0')
  return `${year}-${month}`
}

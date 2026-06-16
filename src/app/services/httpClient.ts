export type JsonRequestOptions = Omit<RequestInit, 'headers'> & {
  headers?: Record<string, string>
}

export async function requestJson(
  url: string,
  options: JsonRequestOptions = {}
): Promise<Record<string, unknown>> {
  const method = (options.method || 'GET').toUpperCase()
  const hasBody = options.body !== undefined && options.body !== null
  const shouldSetJsonContentType = hasBody && method !== 'GET' && method !== 'HEAD'

  const mergedHeaders: Record<string, string> = {
    ...(options.headers || {})
  }

  if (shouldSetJsonContentType && !mergedHeaders['Content-Type']) {
    mergedHeaders['Content-Type'] = 'application/json'
  }

  const response = await fetch(url, {
    ...options,
    headers: mergedHeaders
  })

  const payload = await response.json().catch(() => ({ success: false, message: '响应解析失败' })) as Record<string, unknown>
  if (!response.ok || !payload.success) {
    if (response.status === 429) {
      throw new Error('请求过于频繁（429），请稍后重试。')
    }
    const message = typeof payload.message === 'string' ? payload.message : `请求失败（HTTP ${response.status}）`
    throw new Error(message)
  }

  return payload
}

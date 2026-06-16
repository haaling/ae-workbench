import { requestJson } from './httpClient'

export function createAuthService(authBaseUrl: string) {
  const normalizedBaseUrl = authBaseUrl.replace(/\/$/, '')

  return {
    async loginByAccount(input: { account: string; password: string }): Promise<Record<string, unknown>> {
      return requestJson(`${normalizedBaseUrl}/auth/login`, {
        method: 'POST',
        body: JSON.stringify({
          account: input.account,
          username: input.account,
          password: input.password
        })
      })
    }
  }
}

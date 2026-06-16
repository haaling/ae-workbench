import { requestJson } from './httpClient'

export function createTenantService(tenantApiBaseUrl: string, authToken: string) {
  const normalizedBaseUrl = `${tenantApiBaseUrl.replace(/\/$/, '')}/tenant`

  const authHeaders = () => ({ Authorization: `Bearer ${authToken}` })

  return {
    async getCompanies() {
      return requestJson(`${normalizedBaseUrl}/companies`, { headers: authHeaders() })
    },
    async createCompany(body: Record<string, unknown>) {
      return requestJson(`${normalizedBaseUrl}/companies`, {
        method: 'POST',
        headers: authHeaders(),
        body: JSON.stringify(body)
      })
    },
    async patchCompany(companyId: string, body: Record<string, unknown>) {
      return requestJson(`${normalizedBaseUrl}/companies/${companyId}`, {
        method: 'PATCH',
        headers: authHeaders(),
        body: JSON.stringify(body)
      })
    },
    async patchCompanySubsidiaries(companyId: string, subsidiaries: string[]) {
      return requestJson(`${normalizedBaseUrl}/companies/${companyId}/subsidiaries`, {
        method: 'PATCH',
        headers: authHeaders(),
        body: JSON.stringify({ subsidiaries })
      })
    },
    async getUsers() {
      return requestJson(`${normalizedBaseUrl}/users`, { headers: authHeaders() })
    },
    async createUser(body: Record<string, unknown>) {
      return requestJson(`${normalizedBaseUrl}/users`, {
        method: 'POST',
        headers: authHeaders(),
        body: JSON.stringify(body)
      })
    },
    async getStores() {
      return requestJson(`${normalizedBaseUrl}/stores`, { headers: authHeaders() })
    },
    async createStore(body: Record<string, unknown>) {
      return requestJson(`${normalizedBaseUrl}/stores`, {
        method: 'POST',
        headers: authHeaders(),
        body: JSON.stringify(body)
      })
    },
    async bindStoreEmployees(storeId: string, employeeIds: string[]) {
      return requestJson(`${normalizedBaseUrl}/stores/${storeId}/employees`, {
        method: 'PUT',
        headers: authHeaders(),
        body: JSON.stringify({ employeeIds })
      })
    },
    async getEmployees() {
      return requestJson(`${normalizedBaseUrl}/employees`, { headers: authHeaders() })
    },
    async createEmployee(body: Record<string, unknown>) {
      return requestJson(`${normalizedBaseUrl}/employees`, {
        method: 'POST',
        headers: authHeaders(),
        body: JSON.stringify(body)
      })
    },
    async getPerformanceFinalResults() {
      return requestJson(`${normalizedBaseUrl}/performance/final-results`, { headers: authHeaders() })
    },
    async getPerformanceWorkflows(mineOnly: boolean) {
      const suffix = mineOnly ? '?mine=1' : ''
      return requestJson(`${normalizedBaseUrl}/performance/workflows${suffix}`, { headers: authHeaders() })
    },
    async getWorkflowReviewers() {
      return requestJson(`${normalizedBaseUrl}/performance/workflows/reviewers`, { headers: authHeaders() })
    },
    async createWorkflow(body: Record<string, unknown>) {
      return requestJson(`${normalizedBaseUrl}/performance/workflows`, {
        method: 'POST',
        headers: authHeaders(),
        body: JSON.stringify(body)
      })
    },
    async pushWorkflow(workflowId: string, assignedToUserId: string) {
      return requestJson(`${normalizedBaseUrl}/performance/workflows/${workflowId}/push`, {
        method: 'PATCH',
        headers: authHeaders(),
        body: JSON.stringify({ assignedToUserId })
      })
    },
    async confirmWorkflow(workflowId: string) {
      return requestJson(`${normalizedBaseUrl}/performance/workflows/${workflowId}/confirm`, {
        method: 'PATCH',
        headers: authHeaders()
      })
    },
    async archiveWorkflow(workflowId: string) {
      return requestJson(`${normalizedBaseUrl}/performance/workflows/${workflowId}/archive`, {
        method: 'PATCH',
        headers: authHeaders()
      })
    }
  }
}

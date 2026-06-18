import { useCallback, useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import '../App.css'
import * as Domain from './calculatorDomain'
import {
  canProcessCalculation,
  getExportFileName,
  getProcessDisabledReason as getProcessDisabledReasonText,
  toDateText
} from './utils/appHelpers'
import { formatBeijingDateTime, getCurrentBeijingPeriod } from './utils/beijingTime'
import { createAuthService } from './services/authService'
import { createTenantService } from './services/tenantService'
import { calculatePerformanceResult } from './calculation/calculatePerformanceResult'
import {
  type AdminEmployee,
  type AdminBranchStore,
  type AdminCompanyProfile,
  type AdminManagedUser,
  type PerformanceSnapshot,
  AdminManagementSection,
  EmployeeManagementSection,
  ActionPanel,
  ResultPreviewPanel,
  StoreMetaSection
} from './components/AppSections'
import { MyPerformanceView } from './components/MyPerformanceView'
import type { PerformanceWorkflowItem, WorkflowReviewer } from './types/workflow'
import {
  DetailUploadCard as DetailUploadCardInternal,
  FreightUploadCard as FreightUploadCardInternal,
  OrdersUploadCard as OrdersUploadCardInternal,
  SimpleUploadCard as SimpleUploadCardInternal
} from './components/UploadCards'
import { PerformanceManagementView } from './components/PerformanceManagementView'

type RowData = Domain.RowData
type MultiUploadItem = Domain.MultiUploadItem
type IncomeUploadItem = Domain.IncomeUploadItem
type TableKind = Domain.TableKind
type ProcessResult = Domain.ProcessResult
type PersistedUploadBundle = Domain.PersistedUploadBundle
type AppView = 'calculator' | 'employee' | 'admin' | 'my-performance' | 'performance'

type AuthUser = {
  id: string
  username: string
  email: string
  name?: string
  fullName?: string
  nickname?: string
  role: string
  companyId: string | null
  companyName?: string
  company?: {
    id?: string
    name?: string
  }
}

const AUTH_TOKEN_STORAGE_KEY = 'workbench_access_token'
const AUTH_USER_STORAGE_KEY = 'workbench_auth_user'
const AUTH_API_BASE_URL_KEY = 'workbench_auth_api_base_url'
const TENANT_API_BASE_URL_KEY = 'workbench_tenant_api_base_url'
const AUTH_GUEST_MODE_KEY = 'workbench_guest_mode'
const STORE_ID_STORAGE_KEY = 'workbench_store_id'

const LEGACY_AUTH_TOKEN_STORAGE_KEY = 'saas_access_token'
const LEGACY_AUTH_USER_STORAGE_KEY = 'saas_auth_user'
const LEGACY_AUTH_API_BASE_URL_KEY = 'saas_api_base_url'
const LEGACY_TENANT_API_BASE_URL_KEY = 'saas_tenant_api_base_url'
const LEGACY_AUTH_GUEST_MODE_KEY = 'saas_guest_mode'
const LEGACY_STORE_ID_STORAGE_KEY = 'saas_store_id'
const DEFAULT_AUTH_API_BASE_URL = 'https://workbench-tenant-server-production.up.railway.app/api'
const DEFAULT_TENANT_API_BASE_URL = 'https://workbench-tenant-server-production.up.railway.app/api'
const LEGACY_AUTH_SERVER_HOST = 'dianxiaomi-auth-server-production.up.railway.app'
const WORKFLOW_REVIEWER_CACHE_KEY_PREFIX = 'workbench_workflow_reviewers_cache_v1'
const WORKFLOW_REVIEWER_CACHE_TTL_MS = 6 * 60 * 60 * 1000
const CHROME_EXPORT_PLUGIN_ZIP_PATH = '/aliexpress-order-export-extension.zip'

const {
  DEFAULT_ORDER_ID_HINTS,
  DEFAULT_ORDER_STATUS_HINTS,
  DEFAULT_ORDER_TIME_HINTS,
  DEFAULT_FREIGHT_ID_HINTS,
  DEFAULT_FREIGHT_FULFILLMENT_HINTS,
  DEFAULT_FREIGHT_WAYBILL_HINTS,
  DEFAULT_FREIGHT_CNY_HINTS,
  DEFAULT_FREIGHT_USD_HINTS,
  DEFAULT_REFUND_TYPE_HINTS,
  DEFAULT_INCOME_FLOW_TYPE_HINTS,
  DEFAULT_REFUND_PRODUCT_NAME_HINTS,
  DEFAULT_REFUND_SKU_ID_HINTS,
  normalizeCellValue,
  isExcludedIncomeDetailColumn,
  inferIncomeDetailAmountColumns,
  readUploadCache,
  writeUploadCache,
  clearUploadCache,
  pickDefaultColumn,
  pickOptionalColumn,
  resolveColumnSelection,
  toUploadedState,
  getSelectedSheet,
  validateDetailFiles,
} = Domain

function App() {
  const roleLabelMap: Record<string, string> = {
    super_admin: '超级管理员',
    company_admin: '公司管理者',
    finance: '财务',
    branch_manager: '分公司总经理',
    team_lead: '组长',
    employee: '普通员工',
    general_manager: '总经理',
    manager: '经理',
    gm: '总经理'
  }

  const readStorageWithLegacyFallback = (key: string, legacyKey: string): string => {
    return normalizeCellValue(localStorage.getItem(key)) || normalizeCellValue(localStorage.getItem(legacyKey))
  }

  const normalizeAuthBaseUrl = (value: string | null | undefined): string => {
    const normalized = normalizeCellValue(value)
    if (!normalized) {
      return ''
    }
    // Auto-migrate legacy auth endpoint to the standalone workbench service.
    if (normalized.includes(LEGACY_AUTH_SERVER_HOST)) {
      return DEFAULT_AUTH_API_BASE_URL
    }
    return normalized
  }

  const [authApiBaseUrl] = useState(
    normalizeAuthBaseUrl(readStorageWithLegacyFallback(AUTH_API_BASE_URL_KEY, LEGACY_AUTH_API_BASE_URL_KEY)) ||
    normalizeAuthBaseUrl(import.meta.env.VITE_WORKBENCH_API_BASE_URL) ||
    // Backward compatibility for old deployment config naming.
    normalizeCellValue(import.meta.env.VITE_SAAS_API_BASE_URL) ||
    DEFAULT_AUTH_API_BASE_URL
  )
  const [tenantApiBaseUrl] = useState(
    readStorageWithLegacyFallback(TENANT_API_BASE_URL_KEY, LEGACY_TENANT_API_BASE_URL_KEY) ||
    normalizeCellValue(import.meta.env.VITE_WORKBENCH_API_BASE_URL) ||
    DEFAULT_TENANT_API_BASE_URL
  )
  const [authToken, setAuthToken] = useState(
    readStorageWithLegacyFallback(AUTH_TOKEN_STORAGE_KEY, LEGACY_AUTH_TOKEN_STORAGE_KEY)
  )
  const [isGuestMode, setIsGuestMode] = useState(
    readStorageWithLegacyFallback(AUTH_GUEST_MODE_KEY, LEGACY_AUTH_GUEST_MODE_KEY) === '1'
  )
  const [authUser, setAuthUser] = useState<AuthUser | null>(() => {
    try {
      const raw = localStorage.getItem(AUTH_USER_STORAGE_KEY) || localStorage.getItem(LEGACY_AUTH_USER_STORAGE_KEY)
      return raw ? (JSON.parse(raw) as AuthUser) : null
    } catch {
      return null
    }
  })
  const [loginAccount, setLoginAccount] = useState('')
  const [loginPassword, setLoginPassword] = useState('')
  const [isLoginSubmitting, setIsLoginSubmitting] = useState(false)
  const [adminCompanyProfile, setAdminCompanyProfile] = useState<AdminCompanyProfile | null>(null)
  const [adminCompanies, setAdminCompanies] = useState<AdminCompanyProfile[]>([])
  const [managedUsers, setManagedUsers] = useState<AdminManagedUser[]>([])
  const [employees, setEmployees] = useState<AdminEmployee[]>([])
  const [performanceSnapshots, setPerformanceSnapshots] = useState<PerformanceSnapshot[]>([])
  const [branchStores, setBranchStores] = useState<AdminBranchStore[]>([])
  const [isTenantDataLoading, setIsTenantDataLoading] = useState(false)
  const [isStoreEmployeeOptionsLoading, setIsStoreEmployeeOptionsLoading] = useState(false)
  const [errorMessage, setErrorMessage] = useState('')
  const [successToast, setSuccessToast] = useState('')

  useEffect(() => {
    if (!successToast) {
      return
    }

    const timer = window.setTimeout(() => {
      setSuccessToast('')
    }, 2500)

    return () => {
      window.clearTimeout(timer)
    }
  }, [successToast])

  console.log('authUser:', authUser)

  const accountLabel = normalizeCellValue(authUser?.username || authUser?.email)
  const personNameLabel = useMemo(() => {
    const profileName = normalizeCellValue(authUser?.name || authUser?.fullName || authUser?.nickname)
    if (profileName) {
      return profileName
    }

    const legalRepresentative = normalizeCellValue(adminCompanyProfile?.legalRepresentative)
    if (legalRepresentative) {
      return legalRepresentative
    }

    if (!accountLabel) {
      return ''
    }

    const exact = employees.find((item) => {
      const name = normalizeCellValue(item.name)
      const employeeCode = normalizeCellValue(item.employeeCode)
      return name === accountLabel || employeeCode === accountLabel
    })
    if (exact) {
      return normalizeCellValue(exact.name)
    }

    const fuzzy = employees.find((item) => {
      const name = normalizeCellValue(item.name)
      const employeeCode = normalizeCellValue(item.employeeCode)
      return (name && (name.includes(accountLabel) || accountLabel.includes(name))) ||
        (employeeCode && (employeeCode.includes(accountLabel) || accountLabel.includes(employeeCode)))
    })
    return normalizeCellValue(fuzzy?.name)
  }, [authUser, accountLabel, employees, adminCompanyProfile])

  const totalCompanyLabel = useMemo(() => {
    const fromProfile = normalizeCellValue(adminCompanyProfile?.companyName)
    if (fromProfile) {
      return fromProfile
    }

    const currentCompanyId = normalizeCellValue(authUser?.companyId)
    const fromCompanyList = adminCompanies.find((item) => item.id === currentCompanyId)?.companyName
    if (normalizeCellValue(fromCompanyList)) {
      return normalizeCellValue(fromCompanyList)
    }

    const fromAuthPayload = normalizeCellValue(authUser?.companyName || authUser?.company?.name)
    if (fromAuthPayload) {
      return fromAuthPayload
    }

    return currentCompanyId ? '未设置公司名称' : '未绑定公司'
  }, [adminCompanyProfile, adminCompanies, authUser])

  const displayPersonName = isGuestMode ? 'guest' : personNameLabel || '未设置'
  const displayAccountName = isGuestMode ? 'guest' : accountLabel || '未设置'
  const displayTotalCompany = isGuestMode ? '本地模式' : totalCompanyLabel
  const actualRole = isGuestMode ? '' : normalizeCellValue(authUser?.role)
  const displayRole = isGuestMode
    ? '本地模式用户'
    : roleLabelMap[actualRole] || actualRole || '未设置身份'
  const isAdminUser = !isGuestMode && ['super_admin', 'company_admin'].includes(actualRole)
  const canAccessEmployeeManagement = !isGuestMode && actualRole === 'company_admin'
  const canAccessAdminPage = !isGuestMode && actualRole === 'super_admin'
  const canAccessMyPerformance = !isGuestMode && ['employee', 'branch_manager', 'team_lead'].includes(actualRole)
  const canAccessCalculator = isGuestMode || actualRole !== 'employee'
  const canOperateWorkflow = ['super_admin', 'company_admin', 'finance'].includes(actualRole)
  const canAccessPerformanceManagement = !isGuestMode && canOperateWorkflow
  const [activeView, setActiveView] = useState<AppView>('calculator')
  const [performanceListTab, setPerformanceListTab] = useState<'current' | 'history'>('current')
  const [employeeCompanyTab, setEmployeeCompanyTab] = useState<'subsidiary' | 'store' | 'employee'>('subsidiary')
  const [employeeStaffTab, setEmployeeStaffTab] = useState<'create' | 'list' | 'bind'>('create')

  const resolvedActiveView = useMemo<AppView>(() => {
    if (activeView === 'my-performance' && !canAccessMyPerformance) {
      return canAccessCalculator ? 'calculator' : 'employee'
    }
    if (activeView === 'calculator' && !canAccessCalculator) {
      return canAccessMyPerformance ? 'my-performance' : 'employee'
    }
    if (activeView === 'admin' && !canAccessAdminPage) {
      return canAccessEmployeeManagement ? 'employee' : canAccessMyPerformance ? 'my-performance' : 'calculator'
    }
    if (activeView === 'employee' && !canAccessEmployeeManagement) {
      return canAccessMyPerformance ? 'my-performance' : 'calculator'
    }
    if (activeView === 'performance' && !canAccessPerformanceManagement) {
      return canAccessCalculator ? 'calculator' : canAccessMyPerformance ? 'my-performance' : 'employee'
    }
    return activeView
  }, [
    activeView,
    canAccessAdminPage,
    canAccessEmployeeManagement,
    canAccessMyPerformance,
    canAccessCalculator,
    canAccessPerformanceManagement
  ])

  const resetAuthSession = () => {
    setAuthToken('')
    setAuthUser(null)
    setIsGuestMode(false)
    setActiveView('calculator')
    localStorage.removeItem(AUTH_TOKEN_STORAGE_KEY)
    localStorage.removeItem(AUTH_USER_STORAGE_KEY)
    localStorage.removeItem(AUTH_GUEST_MODE_KEY)
    localStorage.removeItem(STORE_ID_STORAGE_KEY)
    localStorage.removeItem(LEGACY_AUTH_TOKEN_STORAGE_KEY)
    localStorage.removeItem(LEGACY_AUTH_USER_STORAGE_KEY)
    localStorage.removeItem(LEGACY_AUTH_GUEST_MODE_KEY)
    localStorage.removeItem(LEGACY_STORE_ID_STORAGE_KEY)
  }

  const authService = useMemo(() => createAuthService(normalizeCellValue(authApiBaseUrl)), [authApiBaseUrl])
  const tenantService = useMemo(() => createTenantService(tenantApiBaseUrl, authToken), [tenantApiBaseUrl, authToken])

  const parseCompanyExtras = useCallback((notes: unknown): { companyCode: string; legalRepresentative: string; subsidiaries: string[] } => {
    const text = normalizeCellValue(notes)
    if (!text) {
      return { companyCode: '', legalRepresentative: '', subsidiaries: [] }
    }

    try {
      const parsed = JSON.parse(text) as { companyCode?: string; legalRepresentative?: string; subsidiaries?: string[] }
      return {
        companyCode: normalizeCellValue(parsed.companyCode),
        legalRepresentative: normalizeCellValue(parsed.legalRepresentative),
        subsidiaries: Array.isArray(parsed.subsidiaries)
          ? parsed.subsidiaries.map((item) => normalizeCellValue(item)).filter(Boolean)
          : []
      }
    } catch {
      return { companyCode: '', legalRepresentative: '', subsidiaries: [] }
    }
  }, [])

  const mapCompanyProfile = useCallback((company: Record<string, unknown>): AdminCompanyProfile => {
    const extras = parseCompanyExtras(company.notes)
    return {
      id: normalizeCellValue(company._id),
      companyName: normalizeCellValue(company.companyName),
      companyCode: extras.companyCode,
      legalRepresentative: extras.legalRepresentative,
      subsidiaries: extras.subsidiaries,
      createdAt: normalizeCellValue(company.createdAt),
      expireDate: normalizeCellValue(company.expireDate),
      maxUsers: Number(company.maxUsers || 0) || undefined,
      totalUsers: Number(company.totalUsers || 0) || 0,
      activeUsers: Number(company.activeUsers || 0) || 0,
      adminUsers: Number(company.adminUsers || 0) || 0,
      status: normalizeCellValue(company.status) || 'active'
    }
  }, [parseCompanyExtras])

  const mapManagedUser = (user: Record<string, unknown>): AdminManagedUser => ({
    id: normalizeCellValue(user._id || user.id),
    username: normalizeCellValue(user.username),
    email: normalizeCellValue(user.email),
    role: normalizeCellValue(user.role),
    createdAt: normalizeCellValue(user.createdAt),
    isActive: Boolean(user.isActive)
  })

  const mapBranchStore = (store: Record<string, unknown>): AdminBranchStore => {
    const metadata = (store.metadata || {}) as Record<string, unknown>
    const assignedEmployees = Array.isArray(store.employeeIds)
      ? (store.employeeIds as Array<Record<string, unknown>>)
      : []
    return {
      id: normalizeCellValue(store._id || store.id),
      subsidiary: normalizeCellValue(metadata.subsidiary) || '未分公司',
      shopName: normalizeCellValue(store.storeName),
      storeIdOnPlatform: normalizeCellValue(store.storeIdOnPlatform),
      businessLicenseName: normalizeCellValue(metadata.businessLicenseName || metadata.businessLicense),
      assignedEmployeeNames: assignedEmployees.map((item) => normalizeCellValue(item.name)).filter(Boolean),
      employeeIds: assignedEmployees.map((item) => normalizeCellValue(item._id || item.id)).filter(Boolean),
      status: normalizeCellValue(store.status) || 'active',
      createdAt: normalizeCellValue(store.createdAt)
    }
  }

  const composeCompanyNotes = (input: {
    companyCode?: string
    legalRepresentative?: string
    subsidiaries?: string[]
  }): string => {
    const subsidiaries = Array.isArray(input.subsidiaries)
      ? input.subsidiaries.map((item) => normalizeCellValue(item)).filter(Boolean)
      : []
    return JSON.stringify({
      companyCode: normalizeCellValue(input.companyCode),
      legalRepresentative: normalizeCellValue(input.legalRepresentative),
      subsidiaries: Array.from(new Set(subsidiaries))
    })
  }

  const mapEmployee = (employee: Record<string, unknown>): AdminEmployee => ({
    id: normalizeCellValue(employee._id || employee.id),
    userId: normalizeCellValue(employee.userId),
    name: normalizeCellValue(employee.name),
    subsidiary: normalizeCellValue(employee.subsidiary) || '总公司',
    employeeCode: normalizeCellValue(employee.employeeCode),
    status: normalizeCellValue(employee.status) || 'active',
    createdAt: normalizeCellValue(employee.createdAt)
  })

  const mapPerformanceSnapshot = (
    row: Record<string, unknown>,
    storeMap: Map<string, string>
  ): PerformanceSnapshot => {
    const storeId = normalizeCellValue(row.storeId)
    return {
      id: normalizeCellValue(row._id || row.id),
      period: normalizeCellValue(row.period),
      storeLabel: storeMap.get(storeId) || storeId || '未绑定店铺',
      rowCount: Number(row.rowCount || 0) || 0,
      createdAt: formatBeijingDateTime(row.createdAt)
    }
  }

  const mapWorkflowReviewer = useCallback((user: Record<string, unknown>): WorkflowReviewer => ({
    id: normalizeCellValue(user._id || user.id),
    username: normalizeCellValue(user.username),
    role: normalizeCellValue(user.role)
  }), [])

  const mapPerformanceWorkflow = useCallback((
    row: Record<string, unknown>,
    storeMap: Map<string, string>,
    subsidiaryMap: Map<string, string>
  ): PerformanceWorkflowItem => {
    const storeRef = (row.storeId || {}) as Record<string, unknown>
    const storeMetadata = (storeRef.metadata || {}) as Record<string, unknown>
    const assignedUser = (row.assignedToUser || {}) as Record<string, unknown>
    const submittedUser = (row.submittedBy || {}) as Record<string, unknown>
    const resolvedStoreId = normalizeCellValue(storeRef._id || row.storeId)
    const storeName = normalizeCellValue(storeRef.storeName) || storeMap.get(resolvedStoreId) || resolvedStoreId || '未绑定店铺'
    const subsidiaryLabel = normalizeCellValue(storeMetadata.subsidiary) || subsidiaryMap.get(resolvedStoreId) || '未分配分公司'

    return {
      id: normalizeCellValue(row._id || row.id),
      period: normalizeCellValue(row.period),
      status: normalizeCellValue(row.status) || 'draft',
      subsidiaryLabel,
      storeLabel: storeName,
      rowCountCalculated: Number(row.rowCountCalculated || 0) || 0,
      rowCountUploaded: Number(row.rowCountUploaded || 0) || 0,
      assignedToUserName: normalizeCellValue(assignedUser.username),
      submittedByName: normalizeCellValue(submittedUser.username),
      pushedAt: formatBeijingDateTime(row.pushedAt),
      confirmedAt: formatBeijingDateTime(row.confirmedAt),
      archivedAt: formatBeijingDateTime(row.archivedAt),
      createdAt: formatBeijingDateTime(row.createdAt),
      uploadedRows: Array.isArray(row.uploadedRows) ? (row.uploadedRows as RowData[]) : []
    }
  }, [])

  async function refreshEmployeeStoreData() {
    if (!authToken || isGuestMode) {
      return
    }

    const [storesPayload, employeesPayload] = await Promise.all([
      tenantService.getStores(),
      tenantService.getEmployees()
    ])

    const stores = (storesPayload.data as { stores?: Array<Record<string, unknown>> } | undefined)?.stores || []
    const employeeRows = (employeesPayload.data as { employees?: Array<Record<string, unknown>> } | undefined)?.employees || []

    setBranchStores(stores.map(mapBranchStore))
    setEmployees(employeeRows.map(mapEmployee))
  }

  useEffect(() => {
    const shouldLoadManagementData = resolvedActiveView === 'admin' || resolvedActiveView === 'employee'
    if (!shouldLoadManagementData) {
      return
    }

    console.log('Loading tenant data for admin/employee view...')

    let cancelled = false

    const loadTenantData = async () => {
      setIsTenantDataLoading(true)
      try {
        const companiesPayload = await tenantService.getCompanies()
        const usersPayload = await tenantService.getUsers()
        const storesPayload = await tenantService.getStores()
        const employeesPayload = await tenantService.getEmployees()
        const performancePayload = await tenantService.getPerformanceFinalResults()

        if (cancelled) {
          return
        }

        const companies = (companiesPayload.data as { companies?: Array<Record<string, unknown>> } | undefined)?.companies || []
        const users = (usersPayload.data as { users?: Array<Record<string, unknown>> } | undefined)?.users || []
        const stores = (storesPayload.data as { stores?: Array<Record<string, unknown>> } | undefined)?.stores || []
        const employeeRows = (employeesPayload.data as { employees?: Array<Record<string, unknown>> } | undefined)?.employees || []
        const performanceRows = (performancePayload.data as { rows?: Array<Record<string, unknown>> } | undefined)?.rows || []

        const currentCompanyId = normalizeCellValue(authUser?.companyId)
        const selectedCompany =
          companies.find((item) => normalizeCellValue(item._id) === currentCompanyId) ||
          companies[0] ||
          null

        const storeMap = new Map(
          stores.map((store) => [
            normalizeCellValue(store._id),
            normalizeCellValue(store.storeName) || normalizeCellValue(store._id)
          ])
        )

        setAdminCompanies(companies.map(mapCompanyProfile))
        setAdminCompanyProfile(selectedCompany ? mapCompanyProfile(selectedCompany) : null)
        setManagedUsers(users.map(mapManagedUser))
        setEmployees(employeeRows.map(mapEmployee))
        setPerformanceSnapshots(performanceRows.map((item) => mapPerformanceSnapshot(item, storeMap)))
        setBranchStores(stores.map(mapBranchStore))
      } catch (error) {
        if (!cancelled) {
          const message = error instanceof Error ? error.message : '未知错误'
          if (/(用户不存在或已被禁用|令牌已过期|无效的令牌|未提供认证令牌)/.test(message)) {
            resetAuthSession()
            setErrorMessage('登录状态已失效，请重新登录。')
            return
          }
          setErrorMessage(`加载管理员配置失败：${message}`)
        }
      } finally {
        if (!cancelled) {
          setIsTenantDataLoading(false)
        }
      }
    }

    void loadTenantData()

    return () => {
      cancelled = true
    }
  }, [isAdminUser, authToken, authUser?.companyId, tenantService, mapCompanyProfile, resolvedActiveView])

  useEffect(() => {
    const shouldLoadStoreAndEmployeeOptions =
      resolvedActiveView === 'calculator'
    if (!authToken || isGuestMode || !shouldLoadStoreAndEmployeeOptions) {
      return
    }
    if (branchStores.length > 0 && employees.length > 0) {
      return
    }

    let cancelled = false

    const loadStoreEmployeeOptions = async () => {
      setIsStoreEmployeeOptionsLoading(true)
      try {
        const storesPayload = await tenantService.getStores()
        const employeesPayload = await tenantService.getEmployees()
        if (cancelled) {
          return
        }

        const stores = (storesPayload.data as { stores?: Array<Record<string, unknown>> } | undefined)?.stores || []
        const employeeRows = (employeesPayload.data as { employees?: Array<Record<string, unknown>> } | undefined)?.employees || []

        setBranchStores(stores.map(mapBranchStore))
        setEmployees(employeeRows.map(mapEmployee))
      } catch (error) {
        if (!cancelled) {
          const message = error instanceof Error ? error.message : '未知错误'
          setErrorMessage(`加载店铺与员工列表失败：${message}`)
        }
      } finally {
        if (!cancelled) {
          setIsStoreEmployeeOptionsLoading(false)
        }
      }
    }

    void loadStoreEmployeeOptions()

    return () => {
      cancelled = true
    }
  }, [
    authToken,
    isGuestMode,
    resolvedActiveView,
    branchStores.length,
    employees.length,
    tenantService
  ])

  useEffect(() => {
    const inManagementView = resolvedActiveView === 'admin' || resolvedActiveView === 'employee'
    if (!authToken || isGuestMode || !isAdminUser || inManagementView) {
      return
    }
    if (adminCompanyProfile || adminCompanies.length > 0) {
      return
    }

    let cancelled = false

    const loadCompanyForSidebar = async () => {
      try {
        const companiesPayload = await tenantService.getCompanies()
        if (cancelled) {
          return
        }

        const companies = (companiesPayload.data as { companies?: Array<Record<string, unknown>> } | undefined)?.companies || []
        const currentCompanyId = normalizeCellValue(authUser?.companyId)
        const selectedCompany =
          companies.find((item) => normalizeCellValue(item._id) === currentCompanyId) ||
          companies[0] ||
          null

        setAdminCompanies(companies.map(mapCompanyProfile))
        setAdminCompanyProfile(selectedCompany ? mapCompanyProfile(selectedCompany) : null)
      } catch {
        // Keep sidebar fallback text if company list cannot be loaded.
      }
    }

    void loadCompanyForSidebar()

    return () => {
      cancelled = true
    }
  }, [
    authToken,
    isGuestMode,
    isAdminUser,
    resolvedActiveView,
    adminCompanyProfile,
    adminCompanies.length,
    tenantService,
    authUser?.companyId,
    mapCompanyProfile
  ])

  async function handleCreateCompanyProfile(input: {
    companyName: string
    legalRepresentative: string
    adminUsername: string
    adminPassword: string
    adminEmail: string
  }) {
    const companyName = normalizeCellValue(input.companyName)
    const legalRepresentative = normalizeCellValue(input.legalRepresentative)
    const adminUsername = normalizeCellValue(input.adminUsername)
    const adminPassword = normalizeCellValue(input.adminPassword)
    const adminEmail = normalizeCellValue(input.adminEmail)
    if (!companyName || !adminUsername || !adminPassword) {
      setErrorMessage('公司名称、管理员账号、管理员密码为必填项。')
      return
    }
    if (adminPassword.length < 6) {
      setErrorMessage('管理员密码至少 6 位。')
      return
    }

    try {
      const expireDate = new Date()
      expireDate.setFullYear(expireDate.getFullYear() + 1)

      const payload = await tenantService.createCompany({
        companyName,
        expireDate: expireDate.toISOString(),
        maxUsers: 20,
        adminUsername,
        adminPassword,
        adminEmail,
        notes: composeCompanyNotes({ legalRepresentative, subsidiaries: [] })
      })

      const company = (payload.data as { company?: Record<string, unknown> } | undefined)?.company
      if (!company) {
        throw new Error('接口返回缺少 company 数据')
      }

      const adminUser = (payload.data as { adminUser?: Record<string, unknown> } | undefined)?.adminUser

      const mappedCompany = mapCompanyProfile(company)
      setAdminCompanies((prev) => [mappedCompany, ...prev.filter((item) => item.id !== mappedCompany.id)])
      setAdminCompanyProfile(mapCompanyProfile(company))
      if (adminUser) {
        setManagedUsers((prev) => [mapManagedUser(adminUser), ...prev])
      }
      setErrorMessage('公司与管理员创建成功。')
    } catch (error) {
      setErrorMessage(`保存公司信息失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  async function handleAddManagedUser(input: { username: string; email: string; password: string; role: string }) {
    const username = normalizeCellValue(input.username)
    const email = normalizeCellValue(input.email)
    const password = normalizeCellValue(input.password)
    const role = normalizeCellValue(input.role)

    if (!username || !email || !role || !password) {
      setErrorMessage('添加用户失败：用户名、邮箱、初始密码、身份为必填项。')
      return
    }
    if (password.length < 6) {
      setErrorMessage('添加用户失败：初始密码至少 6 位。')
      return
    }

    try {
      const body: Record<string, unknown> = {
        username,
        email,
        password,
        role
      }
      if (actualRole === 'super_admin' && adminCompanyProfile?.id) {
        body.companyId = adminCompanyProfile.id
      }

      const payload = await tenantService.createUser(body)

      const user = (payload.data as { user?: Record<string, unknown> } | undefined)?.user
      if (!user) {
        throw new Error('接口返回缺少 user 数据')
      }

      setManagedUsers((prev) => [mapManagedUser(user), ...prev])

      try {
        const employeeBody: Record<string, unknown> = {
          name: username,
          employeeCode: '',
          notes: 'auto-created-from-add-user',
          userId: normalizeCellValue(user._id || user.id)
        }
        if (actualRole === 'super_admin' && adminCompanyProfile?.id) {
          employeeBody.companyId = adminCompanyProfile.id
        }

        const employeePayload = await tenantService.createEmployee(employeeBody)
        const employee = (employeePayload.data as { employee?: Record<string, unknown> } | undefined)?.employee
        if (!employee) {
          throw new Error('接口返回缺少 employee 数据')
        }

        setEmployees((prev) => [mapEmployee(employee), ...prev])
        await refreshEmployeeStoreData()
        setErrorMessage('用户与员工档案添加成功。')
      } catch (employeeError) {
        setErrorMessage(`用户添加成功，但员工档案创建失败：${employeeError instanceof Error ? employeeError.message : '未知错误'}`)
      }
    } catch (error) {
      setErrorMessage(`添加用户失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  async function handleAddBranchStore(input: { subsidiary: string; shopName: string; businessLicense: string }) {
    const subsidiary = normalizeCellValue(input.subsidiary)
    const shopName = normalizeCellValue(input.shopName)
    const businessLicense = normalizeCellValue(input.businessLicense)

    if (!subsidiary || !shopName || !businessLicense) {
      setErrorMessage('分公司、店铺名、营业执照编号为必填项。')
      return
    }
    if (actualRole === 'super_admin' && !adminCompanyProfile?.id) {
      setErrorMessage('请先创建或选择公司信息，再添加店铺。')
      return
    }

    try {
      const body: Record<string, unknown> = {
        storeName: shopName,
        platform: 'aliexpress',
        metadata: {
          subsidiary,
          businessLicenseName: businessLicense
        }
      }
      if (actualRole === 'super_admin' && adminCompanyProfile?.id) {
        body.companyId = adminCompanyProfile.id
      }

      const payload = await tenantService.createStore(body)

      const store = (payload.data as { store?: Record<string, unknown> } | undefined)?.store
      if (!store) {
        throw new Error('接口返回缺少 store 数据')
      }

      setBranchStores((prev) => [mapBranchStore(store), ...prev])
      await refreshEmployeeStoreData()
      setErrorMessage('分公司店铺信息添加成功。')
    } catch (error) {
      setErrorMessage(`添加分公司店铺失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  async function handleSaveSubsidiaries(subsidiaries: string[]) {
    if (!adminCompanyProfile?.id) {
      setErrorMessage('当前账号未绑定可修改的公司。')
      return
    }

    try {
      let payload: Record<string, unknown>
      try {
        payload = await tenantService.patchCompanySubsidiaries(adminCompanyProfile.id, subsidiaries)
      } catch {
        // Backward compatibility for environments where the dedicated endpoint is not yet deployed.
        payload = await tenantService.patchCompany(adminCompanyProfile.id, {
          notes: composeCompanyNotes({
            companyCode: adminCompanyProfile.companyCode,
            legalRepresentative: adminCompanyProfile.legalRepresentative,
            subsidiaries
          })
        })
      }

      const company = (payload.data as { company?: Record<string, unknown> } | undefined)?.company
      if (!company) {
        throw new Error('接口返回缺少 company 数据')
      }

      const mapped = mapCompanyProfile(company)
      setAdminCompanyProfile(mapped)
      setAdminCompanies((prev) => prev.map((item) => (item.id === mapped.id ? { ...item, ...mapped } : item)))
      setErrorMessage('分公司信息已保存。')
    } catch (error) {
      setErrorMessage(`保存分公司失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  async function handleCreateStoreWithAssignments(input: {
    subsidiary: string
    shopName: string
    businessLicenseName: string
    employeeIds: string[]
  }) {
    const subsidiary = normalizeCellValue(input.subsidiary) || '总公司'
    const shopName = normalizeCellValue(input.shopName)
    const businessLicenseName = normalizeCellValue(input.businessLicenseName)
    const employeeIds = Array.isArray(input.employeeIds)
      ? input.employeeIds.map((item) => normalizeCellValue(item)).filter(Boolean)
      : []

    if (!shopName || !businessLicenseName) {
      setErrorMessage('店铺名和营业执照名为必填项。')
      return
    }

    if (actualRole === 'super_admin' && !adminCompanyProfile?.id) {
      setErrorMessage('请先创建或选择公司信息，再添加店铺。')
      return
    }

    try {
      const body: Record<string, unknown> = {
        storeName: shopName,
        platform: 'aliexpress',
        metadata: {
          subsidiary,
          businessLicenseName
        }
      }
      if (actualRole === 'super_admin' && adminCompanyProfile?.id) {
        body.companyId = adminCompanyProfile.id
      }

      const createdPayload = await tenantService.createStore(body)

      const createdStore = (createdPayload.data as { store?: Record<string, unknown> } | undefined)?.store
      if (!createdStore) {
        throw new Error('接口返回缺少 store 数据')
      }

      let resolvedEmployeeIds = employeeIds
      if (employeeIds.length > 0) {
        const createdEmployees: AdminEmployee[] = []
        const normalizedIds: string[] = []

        for (const rawEmployeeId of employeeIds) {
          if (!rawEmployeeId.startsWith('user:')) {
            normalizedIds.push(rawEmployeeId)
            continue
          }

          const userId = normalizeCellValue(rawEmployeeId.replace(/^user:/, ''))
          const matchedUser = managedUsers.find((user) => user.id === userId)
          if (!matchedUser) {
            throw new Error('账号信息不存在，请刷新后重试。')
          }

          const employeePayload = await tenantService.createEmployee({
            name: normalizeCellValue(matchedUser.username),
            employeeCode: '',
            notes: 'auto-created-from-store-assignment'
          })
          const employee = (employeePayload.data as { employee?: Record<string, unknown> } | undefined)?.employee
          const nextEmployeeId = normalizeCellValue(employee?._id || employee?.id)
          if (!nextEmployeeId) {
            throw new Error('接口返回缺少 employee 数据')
          }

          normalizedIds.push(nextEmployeeId)
          createdEmployees.push(mapEmployee(employee as Record<string, unknown>))
        }

        resolvedEmployeeIds = Array.from(new Set(normalizedIds))
        if (createdEmployees.length > 0) {
          setEmployees((prev) => {
            const merged = [...prev]
            for (const item of createdEmployees) {
              const index = merged.findIndex((existing) => existing.id === item.id)
              if (index >= 0) {
                merged[index] = item
              } else {
                merged.unshift(item)
              }
            }
            return merged
          })
        }
      }

      let nextStore = createdStore
      if (resolvedEmployeeIds.length > 0) {
        const bindPayload = await tenantService.bindStoreEmployees(
          normalizeCellValue(createdStore._id || createdStore.id),
          resolvedEmployeeIds
        )
        nextStore = (bindPayload.data as { store?: Record<string, unknown> } | undefined)?.store || createdStore
      }

      const mapped = mapBranchStore(nextStore)
      setBranchStores((prev) => [mapped, ...prev.filter((item) => item.id !== mapped.id)])
      await refreshEmployeeStoreData()
      setErrorMessage('店铺创建成功，并已完成员工分配。')
    } catch (error) {
      setErrorMessage(`店铺创建失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  async function handleBindEmployeeStores(input: { employeeId: string; storeIds: string[] }) {
    const rawEmployeeId = normalizeCellValue(input.employeeId)
    const targetStoreIds = Array.isArray(input.storeIds)
      ? input.storeIds.map((item) => normalizeCellValue(item)).filter(Boolean)
      : []

    if (!rawEmployeeId) {
      setErrorMessage('请选择需要绑定店铺的员工。')
      return
    }

    let employeeId = rawEmployeeId

    if (rawEmployeeId.startsWith('user:')) {
      const userId = normalizeCellValue(rawEmployeeId.replace(/^user:/, ''))
      const matchedUser = managedUsers.find((user) => user.id === userId)
      if (!matchedUser) {
        setErrorMessage('账号信息不存在，请刷新后重试。')
        return
      }

      try {
        const employeePayload = await tenantService.createEmployee({
          name: normalizeCellValue(matchedUser.username),
          employeeCode: '',
            notes: 'auto-created-from-user-binding',
            userId
        })
        const employee = (employeePayload.data as { employee?: Record<string, unknown> } | undefined)?.employee
        const nextEmployeeId = normalizeCellValue(employee?._id || employee?.id)
        if (!nextEmployeeId) {
          throw new Error('接口返回缺少 employee 数据')
        }

        employeeId = nextEmployeeId
        setEmployees((prev) => {
          const mapped = mapEmployee(employee as Record<string, unknown>)
          return [mapped, ...prev.filter((item) => item.id !== mapped.id)]
        })
      } catch (error) {
        setErrorMessage(`自动创建员工档案失败：${error instanceof Error ? error.message : '未知错误'}`)
        return
      }
    }

    try {
      const updatedStores: AdminBranchStore[] = []
      for (const store of branchStores) {
        const hasEmployee = store.employeeIds.includes(employeeId)
        const shouldHaveEmployee = targetStoreIds.includes(store.id)
        if (hasEmployee === shouldHaveEmployee) {
          continue
        }

        const nextEmployeeIds = shouldHaveEmployee
          ? Array.from(new Set([...store.employeeIds, employeeId]))
          : store.employeeIds.filter((id) => id !== employeeId)

        const payload = await tenantService.bindStoreEmployees(store.id, nextEmployeeIds)
        const updatedStore = (payload.data as { store?: Record<string, unknown> } | undefined)?.store
        if (updatedStore) {
          updatedStores.push(mapBranchStore(updatedStore))
        }
      }

      if (updatedStores.length > 0) {
        const updatedMap = new Map(updatedStores.map((item) => [item.id, item]))
        setBranchStores((prev) => prev.map((item) => updatedMap.get(item.id) || item))
      }
      await refreshEmployeeStoreData()
      setErrorMessage('员工店铺绑定已更新。')
    } catch (error) {
      setErrorMessage(`保存员工店铺绑定失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  async function handleAddStaff(input: {
    name: string
    username: string
    email?: string
    password: string
    role: string
    subsidiary: string
    employeeCode: string
    notes: string
  }) {
    const name = normalizeCellValue(input.name)
    const username = normalizeCellValue(input.username)
    const email = normalizeCellValue(input.email)
    const password = normalizeCellValue(input.password)
    const role = normalizeCellValue(input.role)
    const subsidiary = normalizeCellValue(input.subsidiary) || '总公司'
    const employeeCode = normalizeCellValue(input.employeeCode)
    const notes = normalizeCellValue(input.notes)

    if (!name || !username || !password || !role) {
      setErrorMessage('创建人员失败：姓名、账号、初始密码、身份为必填项。')
      return
    }
    if (password.length < 6) {
      setErrorMessage('创建人员失败：初始密码至少 6 位。')
      return
    }

    try {
      const userBody: Record<string, unknown> = {
        username,
        password,
        role
      }
      if (email) {
        userBody.email = email
      }
      if (actualRole === 'super_admin' && adminCompanyProfile?.id) {
        userBody.companyId = adminCompanyProfile.id
      }

      const userPayload = await tenantService.createUser(userBody)
      const user = (userPayload.data as { user?: Record<string, unknown> } | undefined)?.user
      if (!user) {
        throw new Error('接口返回缺少 user 数据')
      }
      setManagedUsers((prev) => [mapManagedUser(user), ...prev])

      const employeeBody: Record<string, unknown> = {
        name,
        subsidiary,
        employeeCode,
        notes,
        userId: normalizeCellValue(user._id || user.id)
      }
      if (actualRole === 'super_admin' && adminCompanyProfile?.id) {
        employeeBody.companyId = adminCompanyProfile.id
      }

      const employeePayload = await tenantService.createEmployee(employeeBody)
      const employee = (employeePayload.data as { employee?: Record<string, unknown> } | undefined)?.employee
      if (!employee) {
        throw new Error('接口返回缺少 employee 数据')
      }

      setEmployees((prev) => [mapEmployee(employee), ...prev])
      await refreshEmployeeStoreData()
      setSuccessToast('新增员工成功（账号+档案）。')
    } catch (error) {
      setErrorMessage(`创建人员失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  async function handleResignEmployee(employeeId: string) {
    const normalizedEmployeeId = normalizeCellValue(employeeId)
    if (!normalizedEmployeeId) {
      return
    }

    if (!window.confirm('确认将该员工标记为已离职，并停用其账号吗？')) {
      return
    }

    try {
      const payload = await tenantService.resignEmployee(normalizedEmployeeId)
      const employee = (payload.data as { employee?: Record<string, unknown> } | undefined)?.employee
      const user = (payload.data as { user?: Record<string, unknown> | null } | undefined)?.user

      if (employee) {
        const mappedEmployee = mapEmployee(employee)
        setEmployees((prev) => prev.map((item) => item.id === mappedEmployee.id ? mappedEmployee : item))
      }

      if (user) {
        const mappedUser = mapManagedUser(user)
        setManagedUsers((prev) => prev.map((item) => item.id === mappedUser.id ? mappedUser : item))
      }

      await refreshEmployeeStoreData()
      setSuccessToast(user ? '员工已标记离职，账号已停用。' : '员工已标记离职。')
    } catch (error) {
      setErrorMessage(`员工离职操作失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  async function handleUpdateCompanyProfile(input: {
    companyName: string
    legalRepresentative: string
    maxUsers: number
    expireDate: string
  }) {
    if (!adminCompanyProfile?.id) {
      setErrorMessage('当前账号未绑定可修改的公司。')
      return
    }

    const companyName = normalizeCellValue(input.companyName)
    const legalRepresentative = normalizeCellValue(input.legalRepresentative)
    const expireDate = normalizeCellValue(input.expireDate)
    const maxUsers = Number(input.maxUsers)
    const canManageCompanyPolicy = actualRole === 'super_admin'

    if (!companyName) {
      setErrorMessage('公司名称不能为空。')
      return
    }
    if (canManageCompanyPolicy && (!Number.isFinite(maxUsers) || maxUsers < 1)) {
      setErrorMessage('最大账号数必须大于 0。')
      return
    }

    try {
      const updateBody: Record<string, unknown> = {
        companyName,
        notes: composeCompanyNotes({
          companyCode: adminCompanyProfile.companyCode,
          legalRepresentative,
          subsidiaries: adminCompanyProfile.subsidiaries || []
        })
      }
      if (canManageCompanyPolicy) {
        updateBody.maxUsers = maxUsers
        updateBody.expireDate = expireDate
      }

      const payload = await tenantService.patchCompany(adminCompanyProfile.id, updateBody)

      const company = (payload.data as { company?: Record<string, unknown> } | undefined)?.company
      if (!company) {
        throw new Error('接口返回缺少 company 数据')
      }

      const mapped = mapCompanyProfile(company)
      setAdminCompanyProfile(mapped)
      setAdminCompanies((prev) => prev.map((item) => (item.id === mapped.id ? { ...item, ...mapped } : item)))
      setErrorMessage('公司信息更新成功。')
    } catch (error) {
      setErrorMessage(`更新公司信息失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  const cachedUploadData = useMemo(() => readUploadCache(), [])

  const [ordersFiles, setOrdersFiles] = useState<MultiUploadItem[]>(
    Array.isArray(cachedUploadData?.ordersFiles) ? cachedUploadData.ordersFiles : []
  )
  const [incomeFiles, setIncomeFiles] = useState<IncomeUploadItem[]>(
    Array.isArray(cachedUploadData?.incomeFiles) ? cachedUploadData.incomeFiles : []
  )
  const [refundFiles, setRefundFiles] = useState<IncomeUploadItem[]>(
    Array.isArray(cachedUploadData?.refundFiles) ? cachedUploadData.refundFiles : []
  )
  const [freightFiles, setFreightFiles] = useState<MultiUploadItem[]>(
    Array.isArray(cachedUploadData?.freightFiles) ? cachedUploadData.freightFiles : []
  )
  const [alipayFiles, setAlipayFiles] = useState<MultiUploadItem[]>(
    Array.isArray(cachedUploadData?.alipayFiles) ? cachedUploadData.alipayFiles : []
  )
  const [offlineFiles, setOfflineFiles] = useState<MultiUploadItem[]>(
    Array.isArray(cachedUploadData?.offlineFiles) ? cachedUploadData.offlineFiles : []
  )

  const [ordersIdColumn, setOrdersIdColumn] = useState(normalizeCellValue(cachedUploadData?.ordersIdColumn))
  const [ordersStatusColumn, setOrdersStatusColumn] = useState(normalizeCellValue(cachedUploadData?.ordersStatusColumn))
  const [ordersTimeColumn, setOrdersTimeColumn] = useState(normalizeCellValue(cachedUploadData?.ordersTimeColumn))
  const [incomeDetailAmountColumns, setIncomeDetailAmountColumns] = useState<string[]>(
    Array.isArray(cachedUploadData?.incomeDetailAmountColumns) ? cachedUploadData.incomeDetailAmountColumns : []
  )
  const [refundDetailAmountColumns, setRefundDetailAmountColumns] = useState<string[]>(
    Array.isArray(cachedUploadData?.refundDetailAmountColumns) ? cachedUploadData.refundDetailAmountColumns : []
  )
  const [refundTypeColumn, setRefundTypeColumn] = useState(normalizeCellValue(cachedUploadData?.refundTypeColumn))
  const [refundProductNameColumn, setRefundProductNameColumn] = useState(normalizeCellValue(cachedUploadData?.refundProductNameColumn))
  const [refundSkuIdColumn, setRefundSkuIdColumn] = useState(normalizeCellValue(cachedUploadData?.refundSkuIdColumn))
  const [freightOrderColumn, setFreightOrderColumn] = useState(normalizeCellValue(cachedUploadData?.freightOrderColumn))
  const [freightFulfillmentColumn, setFreightFulfillmentColumn] = useState(normalizeCellValue(cachedUploadData?.freightFulfillmentColumn))
  const [freightWaybillColumn, setFreightWaybillColumn] = useState(normalizeCellValue(cachedUploadData?.freightWaybillColumn))
  const [freightAmountCnyColumn, setFreightAmountCnyColumn] = useState(normalizeCellValue(cachedUploadData?.freightAmountCnyColumn))
  const [freightAmountUsdColumn, setFreightAmountUsdColumn] = useState(normalizeCellValue(cachedUploadData?.freightAmountUsdColumn))

  const [isProcessing, setIsProcessing] = useState(false)
  const [isWorkflowSubmitting, setIsWorkflowSubmitting] = useState(false)
  const [isWorkflowListLoading, setIsWorkflowListLoading] = useState(false)
  const [result, setResult] = useState<ProcessResult | null>(null)
  const [lastCalculatedAt, setLastCalculatedAt] = useState('')
  const [calculationCount, setCalculationCount] = useState(0)
  const [shopId, setShopId] = useState(normalizeCellValue(cachedUploadData?.shopId))
  const [shopName, setShopName] = useState(normalizeCellValue(cachedUploadData?.shopName))
  const [subsidiary, setSubsidiary] = useState(normalizeCellValue(cachedUploadData?.subsidiary))
  const [selectedEmployeeId, setSelectedEmployeeId] = useState('')
  const [usdExchangeRate, setUsdExchangeRate] = useState(normalizeCellValue(cachedUploadData?.usdExchangeRate) || '7.20')
  const [uploadedPerformanceRows, setUploadedPerformanceRows] = useState<RowData[]>([])
  const [uploadedPerformanceFileName, setUploadedPerformanceFileName] = useState('')
  const [workflowPeriod, setWorkflowPeriod] = useState(
    normalizeCellValue(cachedUploadData?.workflowPeriod) || getCurrentBeijingPeriod()
  )
  const [workflowReviewers, setWorkflowReviewers] = useState<WorkflowReviewer[]>([])
  const [selectedReviewerUserId, setSelectedReviewerUserId] = useState('')
  const [performanceWorkflows, setPerformanceWorkflows] = useState<PerformanceWorkflowItem[]>([])
  const [uploadCacheReady] = useState(true)
  const reviewerCacheCompanyId = normalizeCellValue(authUser?.companyId || authUser?.company?.id)

  const isEmployeeReviewOnly = actualRole === 'employee'
  const workflowStatusLabelMap: Record<string, string> = {
    draft: '草稿',
    pushed: '待员工核对',
    confirmed: '已确认待归档',
    archived: '已归档'
  }

  const employeePendingWorkflows = performanceWorkflows.filter((item) => item.status === 'pushed')
  const employeeHistoryWorkflows = performanceWorkflows.filter((item) => item.status === 'confirmed' || item.status === 'archived')
  const financeArchiveReadyWorkflows = performanceWorkflows.filter((item) => item.status === 'confirmed')

  const pickDefaultReviewerByEmployee = (employeeId: string): string => {
    if (!employeeId || workflowReviewers.length === 0) {
      return ''
    }
    const employee = employees.find((item) => item.id === employeeId)
    const linkedUserId = normalizeCellValue(employee?.userId)
    if (linkedUserId) {
      const linkedReviewer = workflowReviewers.find((item) => item.id === linkedUserId)
      if (linkedReviewer) {
        return linkedReviewer.id
      }
    }
    const employeeName = normalizeCellValue(employee?.name)
    const employeeCode = normalizeCellValue(employee?.employeeCode)
    if (!employeeName && !employeeCode) {
      return ''
    }

    const exact = workflowReviewers.find((item) => {
      const username = normalizeCellValue(item.username)
      return (employeeName && username === employeeName) || (employeeCode && username === employeeCode)
    })
    if (exact) {
      return exact.id
    }

    const fuzzy = workflowReviewers.find((item) => {
      const username = normalizeCellValue(item.username)
      return (employeeName && username.includes(employeeName)) || (employeeCode && username.includes(employeeCode))
    })
    return fuzzy?.id || ''
  }

  const pickDefaultReviewerByEmployeeFrom = useCallback((employeeId: string, reviewers: WorkflowReviewer[]): string => {
    if (!employeeId || reviewers.length === 0) {
      return ''
    }
    const employee = employees.find((item) => item.id === employeeId)
    const linkedUserId = normalizeCellValue(employee?.userId)
    if (linkedUserId) {
      const linkedReviewer = reviewers.find((item) => item.id === linkedUserId)
      if (linkedReviewer) {
        return linkedReviewer.id
      }
    }
    const employeeName = normalizeCellValue(employee?.name)
    const employeeCode = normalizeCellValue(employee?.employeeCode)
    if (!employeeName && !employeeCode) {
      return ''
    }

    const exact = reviewers.find((item) => {
      const username = normalizeCellValue(item.username)
      return (employeeName && username === employeeName) || (employeeCode && username === employeeCode)
    })
    if (exact) {
      return exact.id
    }

    const fuzzy = reviewers.find((item) => {
      const username = normalizeCellValue(item.username)
      return (employeeName && username.includes(employeeName)) || (employeeCode && username.includes(employeeCode))
    })
    return fuzzy?.id || ''
  }, [employees])

  const refreshWorkflowData = async (showLoading = true, forceReviewerRefresh = false) => {
    if (!authToken || isGuestMode) {
      setPerformanceWorkflows([])
      setWorkflowReviewers([])
      return
    }

    if (showLoading) {
      setIsWorkflowListLoading(true)
    }
    try {
      const workflowPayload = await tenantService.getPerformanceWorkflows(!canOperateWorkflow)
      let reviewerRows: Array<Record<string, unknown>> = []
      const reviewerCacheKey = reviewerCacheCompanyId
        ? `${WORKFLOW_REVIEWER_CACHE_KEY_PREFIX}:${reviewerCacheCompanyId}`
        : ''

      if (canOperateWorkflow) {
        if (!forceReviewerRefresh && reviewerCacheKey) {
          try {
            const cachedRaw = localStorage.getItem(reviewerCacheKey)
            if (cachedRaw) {
              const cached = JSON.parse(cachedRaw) as {
                users?: Array<Record<string, unknown>>
                updatedAt?: number
              }
              const updatedAt = Number(cached.updatedAt || 0)
              const users = Array.isArray(cached.users) ? cached.users : []
              if (users.length > 0 && Date.now() - updatedAt < WORKFLOW_REVIEWER_CACHE_TTL_MS) {
                reviewerRows = users
              }
            }
          } catch {
            reviewerRows = []
          }
        }

        try {
          if (reviewerRows.length === 0) {
            const reviewerPayload = await tenantService.getWorkflowReviewers()
            reviewerRows = (reviewerPayload.data as { users?: Array<Record<string, unknown>> } | undefined)?.users || []
          }
        } catch {
          if (reviewerRows.length === 0) {
            reviewerRows = []
          }
        }

        if (reviewerRows.length === 0) {
          // Backward compatibility: some deployments do not expose the dedicated reviewers endpoint.
          const usersPayload = await tenantService.getUsers()
          const users = (usersPayload.data as { users?: Array<Record<string, unknown>> } | undefined)?.users || []
          reviewerRows = users.filter((user) => {
            const isActive = Boolean(user.isActive)
            return isActive
          })
        }

        if (reviewerRows.length > 0 && reviewerCacheKey) {
          try {
            localStorage.setItem(
              reviewerCacheKey,
              JSON.stringify({ users: reviewerRows, updatedAt: Date.now() })
            )
          } catch {
            // Ignore cache write failures (quota/private mode).
          }
        }
      }

      const workflowRows = (workflowPayload.data as { workflows?: Array<Record<string, unknown>> } | undefined)?.workflows || []
      const storeMap = new Map(branchStores.map((item) => [item.id, item.shopName]))
      const subsidiaryMap = new Map(branchStores.map((item) => [item.id, item.subsidiary]))

      const mappedWorkflows = workflowRows.map((row) => mapPerformanceWorkflow(row, storeMap, subsidiaryMap))
      const mappedReviewers = reviewerRows.map(mapWorkflowReviewer)
      setPerformanceWorkflows(mappedWorkflows)
      setWorkflowReviewers(mappedReviewers)

      if (selectedEmployeeId) {
        const defaultReviewerUserId = pickDefaultReviewerByEmployeeFrom(selectedEmployeeId, mappedReviewers)
        const hasCurrentReviewer = mappedReviewers.some((item) => item.id === selectedReviewerUserId)
        if (defaultReviewerUserId && (!selectedReviewerUserId || !hasCurrentReviewer)) {
          setSelectedReviewerUserId(defaultReviewerUserId)
        }
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知错误'
      if (/(无权限|403)/.test(message)) {
        setPerformanceWorkflows([])
        setWorkflowReviewers([])
      } else {
        setErrorMessage(`加载绩效核对流程失败：${message}`)
      }
    } finally {
      if (showLoading) {
        setIsWorkflowListLoading(false)
      }
    }
  }

  const subsidiaryOptions = useMemo(() => {
    const profileSubsidiaries = Array.isArray(adminCompanyProfile?.subsidiaries)
      ? adminCompanyProfile.subsidiaries
      : []
    const storeSubsidiaries = branchStores
      .map((item) => normalizeCellValue(item.subsidiary))
      .filter(Boolean)

    const merged = Array.from(new Set([
      ...profileSubsidiaries.map((item) => normalizeCellValue(item)).filter(Boolean),
      ...storeSubsidiaries
    ]))

    return merged
  }, [adminCompanyProfile, branchStores])

  const employeeOptions = useMemo(() => {
    if (!subsidiary) {
      return [] as Array<{ id: string; label: string }>
    }

    const idsInSubsidiary = new Set(
      branchStores
        .filter((item) => normalizeCellValue(item.subsidiary) === subsidiary)
        .flatMap((item) => item.employeeIds)
    )

    const scopedEmployees = idsInSubsidiary.size > 0
      ? employees.filter((item) => idsInSubsidiary.has(item.id))
      : employees

    const findAccountUsernameByEmployee = (employee: AdminEmployee): string => {
      const employeeName = normalizeCellValue(employee.name)
      const employeeCode = normalizeCellValue(employee.employeeCode)
      if (!employeeName && !employeeCode) {
        return ''
      }

      const exact = workflowReviewers.find((reviewer) => {
        const username = normalizeCellValue(reviewer.username)
        return (employeeName && username === employeeName) || (employeeCode && username === employeeCode)
      })
      if (exact) {
        return normalizeCellValue(exact.username)
      }

      const fuzzy = workflowReviewers.find((reviewer) => {
        const username = normalizeCellValue(reviewer.username)
        return (employeeName && username.includes(employeeName)) || (employeeCode && username.includes(employeeCode))
      })
      return normalizeCellValue(fuzzy?.username)
    }

    return scopedEmployees.map((item) => ({
      id: item.id,
      label: `${item.name}${findAccountUsernameByEmployee(item) ? `（${findAccountUsernameByEmployee(item)}）` : item.employeeCode ? `（${item.employeeCode}）` : ''}`
    }))
  }, [subsidiary, branchStores, employees, workflowReviewers])

  const storeOptions = useMemo(() => {
    if (!subsidiary || !selectedEmployeeId) {
      return [] as Array<{ id: string; name: string }>
    }

    return branchStores
      .filter((item) => normalizeCellValue(item.subsidiary) === subsidiary && item.employeeIds.includes(selectedEmployeeId))
      .map((item) => ({ id: item.id, name: item.shopName }))
  }, [subsidiary, selectedEmployeeId, branchStores])

  const selectedBusinessLicenseName = useMemo(() => {
    if (!shopId) {
      return ''
    }
    const selectedStore = branchStores.find((item) => item.id === shopId)
    return normalizeCellValue(selectedStore?.businessLicenseName)
  }, [shopId, branchStores])

  useEffect(() => {
    localStorage.setItem(AUTH_API_BASE_URL_KEY, authApiBaseUrl)
    localStorage.removeItem(LEGACY_AUTH_API_BASE_URL_KEY)
  }, [authApiBaseUrl])

  useEffect(() => {
    localStorage.setItem(TENANT_API_BASE_URL_KEY, tenantApiBaseUrl)
    localStorage.removeItem(LEGACY_TENANT_API_BASE_URL_KEY)
  }, [tenantApiBaseUrl])

  useEffect(() => {
    if (!uploadCacheReady) {
      return
    }

    const payload: PersistedUploadBundle = {
      ordersFiles,
      incomeFiles,
      refundFiles,
      freightFiles,
      alipayFiles,
      offlineFiles,
      ordersIdColumn,
      ordersStatusColumn,
      ordersTimeColumn,
      incomeDetailAmountColumns,
      refundDetailAmountColumns,
      refundTypeColumn,
      refundProductNameColumn,
      refundSkuIdColumn,
      freightOrderColumn,
      freightFulfillmentColumn,
      freightWaybillColumn,
      freightAmountCnyColumn,
      freightAmountUsdColumn,
      shopId,
      shopName,
      subsidiary,
      usdExchangeRate,
      workflowPeriod
    }

    const writeResult = writeUploadCache(payload)
    if (!writeResult.ok && writeResult.reason === 'quota') {
      console.warn('本地缓存空间不足，建议减少上传文件数量或先清空本地缓存。')
    }
  }, [
    uploadCacheReady,
    ordersFiles,
    incomeFiles,
    refundFiles,
    freightFiles,
    alipayFiles,
    offlineFiles,
    ordersIdColumn,
    ordersStatusColumn,
    ordersTimeColumn,
    incomeDetailAmountColumns,
    refundDetailAmountColumns,
    refundTypeColumn,
    refundProductNameColumn,
    refundSkuIdColumn,
    freightOrderColumn,
    freightFulfillmentColumn,
    freightWaybillColumn,
    freightAmountCnyColumn,
    freightAmountUsdColumn,
    shopId,
    shopName,
    subsidiary,
    usdExchangeRate,
    workflowPeriod
  ])

  function clearLocalUploadCache() {
    clearUploadCache()
    setOrdersFiles([])
    setIncomeFiles([])
    setRefundFiles([])
    setFreightFiles([])
    setAlipayFiles([])
    setOfflineFiles([])
    setOrdersIdColumn('')
    setOrdersStatusColumn('')
    setOrdersTimeColumn('')
    setIncomeDetailAmountColumns([])
    setRefundDetailAmountColumns([])
    setRefundTypeColumn('')
    setRefundProductNameColumn('')
    setRefundSkuIdColumn('')
    setFreightOrderColumn('')
    setFreightFulfillmentColumn('')
    setFreightWaybillColumn('')
    setFreightAmountCnyColumn('')
    setFreightAmountUsdColumn('')
    setShopId('')
    setShopName('')
    setSelectedEmployeeId('')
    setResult(null)
    setLastCalculatedAt('')
    setCalculationCount(0)
    setErrorMessage('已清空本地缓存与当前上传数据。')
  }

  const ordersSheetRows = useMemo(
    () => ordersFiles.flatMap((item) => getSelectedSheet(item.file)?.rows || []),
    [ordersFiles]
  )
  const ordersHeaders = useMemo(
    () => Array.from(new Set(ordersFiles.flatMap((item) => getSelectedSheet(item.file)?.headers || []))),
    [ordersFiles]
  )
  const freightSheetRows = useMemo(
    () => freightFiles.flatMap((item) => getSelectedSheet(item.file)?.rows || []),
    [freightFiles]
  )
  const freightHeaders = useMemo(
    () => Array.from(new Set(freightFiles.flatMap((item) => getSelectedSheet(item.file)?.headers || []))),
    [freightFiles]
  )
  const incomeHeaders = useMemo(() => {
    return Array.from(new Set(incomeFiles.flatMap((item) => getSelectedSheet(item.file)?.headers || [])))
  }, [incomeFiles])
  const refundHeaders = useMemo(() => {
    return Array.from(new Set(refundFiles.flatMap((item) => getSelectedSheet(item.file)?.headers || [])))
  }, [refundFiles])
  const effectiveIncomeDetailAmountColumns = useMemo(() => {
    const valid = incomeDetailAmountColumns.filter(
      (column) => incomeHeaders.includes(column) && !isExcludedIncomeDetailColumn(column)
    )
    if (valid.length > 0) {
      return valid
    }
    return inferIncomeDetailAmountColumns(incomeHeaders)
  }, [incomeDetailAmountColumns, incomeHeaders])
  const effectiveRefundDetailAmountColumns = useMemo(() => {
    const valid = refundDetailAmountColumns.filter(
      (column) => refundHeaders.includes(column) && !isExcludedIncomeDetailColumn(column)
    )
    if (valid.length > 0) {
      return valid
    }
    return inferIncomeDetailAmountColumns(refundHeaders)
  }, [refundDetailAmountColumns, refundHeaders])
  const effectiveRefundTypeColumn = resolveColumnSelection(
    refundTypeColumn,
    refundHeaders,
    DEFAULT_REFUND_TYPE_HINTS
  )
  const effectiveRefundProductNameColumn = resolveColumnSelection(
    refundProductNameColumn,
    refundHeaders,
    DEFAULT_REFUND_PRODUCT_NAME_HINTS
  )
  const effectiveRefundSkuIdColumn = resolveColumnSelection(
    refundSkuIdColumn,
    refundHeaders,
    DEFAULT_REFUND_SKU_ID_HINTS
  )

  const effectiveOrdersStatusColumn = resolveColumnSelection(
    ordersStatusColumn,
    ordersHeaders,
    DEFAULT_ORDER_STATUS_HINTS
  )
  const effectiveOrdersTimeColumn = resolveColumnSelection(
    ordersTimeColumn,
    ordersHeaders,
    DEFAULT_ORDER_TIME_HINTS
  )
  const effectiveFreightFulfillmentColumn = resolveColumnSelection(
    freightFulfillmentColumn,
    freightHeaders,
    DEFAULT_FREIGHT_FULFILLMENT_HINTS
  )
  const effectiveFreightWaybillColumn = resolveColumnSelection(
    freightWaybillColumn,
    freightHeaders,
    DEFAULT_FREIGHT_WAYBILL_HINTS
  )
  const effectiveFreightAmountCnyColumn = resolveSafeFreightAmountColumnSelection(
    freightAmountCnyColumn,
    freightHeaders,
    DEFAULT_FREIGHT_CNY_HINTS
  )
  const effectiveFreightAmountUsdColumn = resolveSafeFreightAmountColumnSelection(
    freightAmountUsdColumn,
    freightHeaders,
    DEFAULT_FREIGHT_USD_HINTS
  )

  function isLikelyOrderLikeHeader(column: string): boolean {
    const text = normalizeCellValue(column).toLowerCase()
    if (!text) {
      return false
    }

    return [
      '订单号',
      '订单编号',
      '订单id',
      '交易单号',
      '交易订单号',
      '商家订单号',
      '客户订单号',
      '平台订单号',
      'order',
      'orderid',
      'trade',
      'waybill'
    ].some((hint) => text.includes(hint))
  }

  function resolveSafeFreightAmountColumnSelection(selected: string, headers: string[], hints: string[]): string {
    if (selected && headers.includes(selected) && !isLikelyOrderLikeHeader(selected)) {
      return selected
    }

    const picked = pickOptionalColumn(
      headers.filter((header) => !isLikelyOrderLikeHeader(header)),
      hints
    )

    return picked || ''
  }

  function scoreDecodedCsvText(text: string): number {
    const sample = text.slice(0, 4000)
    const headerHints = [
      '交易时间',
      '交易分类',
      '交易对方',
      '商品说明',
      '金额',
      '备注',
      '订单号',
      '订单状态',
      '收/付款方式',
      '交易订单号',
      '商家订单号'
    ]

    let score = 0
    headerHints.forEach((hint) => {
      if (sample.includes(hint)) {
        score += 8
      }
    })

    if (sample.includes('��')) {
      score -= 30
    }

    const mojibakeMatches = sample.match(/[æäåçéïð]/g)
    if (mojibakeMatches) {
      score -= mojibakeMatches.length
    }

    return score
  }

  function readWorkbookFromUpload(fileName: string, buffer: ArrayBuffer): XLSX.WorkBook {
    const isCsv = /\.csv$/i.test(fileName)
    if (!isCsv) {
      return XLSX.read(buffer, { type: 'array' })
    }

    const bytes = new Uint8Array(buffer)
    const hasUtf8Bom = bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf
    const candidateEncodings = hasUtf8Bom ? ['utf-8', 'gb18030', 'gbk'] : ['utf-8', 'gb18030', 'gbk']

    let bestText = ''
    let bestScore = Number.NEGATIVE_INFINITY

    candidateEncodings.forEach((encoding) => {
      try {
        const decoded = new TextDecoder(encoding as string).decode(buffer)
        const score = scoreDecodedCsvText(decoded)
        if (score > bestScore) {
          bestScore = score
          bestText = decoded
        }
      } catch {
        // Ignore unsupported decoders and keep trying the next candidate.
      }
    })

    return XLSX.read(bestText, { type: 'string' })
  }

  async function handleFileUpload(event: React.ChangeEvent<HTMLInputElement>, table: TableKind) {
    const files = Array.from(event.target.files || [])
    if (files.length === 0) {
      return
    }

    setErrorMessage('')

    try {
      for (const file of files) {
        const buffer = await file.arrayBuffer()
        const workbook = readWorkbookFromUpload(file.name, buffer)
        const uploaded = toUploadedState(file.name, workbook)
        const defaultSheetHeaders = uploaded.sheets[0]?.headers || []

        if (table === 'orders') {
          setOrdersIdColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_ORDER_ID_HINTS))
          setOrdersStatusColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_ORDER_STATUS_HINTS))
          setOrdersTimeColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_ORDER_TIME_HINTS))
          setOrdersFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded }]
          })
          continue
        }

        if (table === 'income') {
          const orderColumn = pickDefaultColumn(defaultSheetHeaders, DEFAULT_ORDER_ID_HINTS)
          const flowTypeColumn = pickOptionalColumn(defaultSheetHeaders, DEFAULT_INCOME_FLOW_TYPE_HINTS)
          setIncomeDetailAmountColumns((prev) => {
            if (prev.length > 0) {
              return prev
            }
            return inferIncomeDetailAmountColumns(defaultSheetHeaders)
          })
          setIncomeFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded, orderColumn, flowTypeColumn }]
          })
          continue
        }

        if (table === 'refund') {
          const orderColumn = pickDefaultColumn(defaultSheetHeaders, DEFAULT_ORDER_ID_HINTS)
          setRefundTypeColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_REFUND_TYPE_HINTS))
          setRefundProductNameColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_REFUND_PRODUCT_NAME_HINTS))
          setRefundSkuIdColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_REFUND_SKU_ID_HINTS))
          setRefundDetailAmountColumns((prev) => {
            if (prev.length > 0) {
              return prev
            }
            return inferIncomeDetailAmountColumns(defaultSheetHeaders)
          })
          setRefundFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded, orderColumn }]
          })
          continue
        }

        if (table === 'freight') {
          setFreightOrderColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_FREIGHT_ID_HINTS))
          setFreightFulfillmentColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_FREIGHT_FULFILLMENT_HINTS))
          setFreightWaybillColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_FREIGHT_WAYBILL_HINTS))
          setFreightAmountCnyColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_FREIGHT_CNY_HINTS))
          setFreightAmountUsdColumn((prev) => prev || pickDefaultColumn(defaultSheetHeaders, DEFAULT_FREIGHT_USD_HINTS))
          setFreightFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded }]
          })
          continue
        }

        if (table === 'alipay') {
          setAlipayFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded }]
          })
          continue
        }

        if (table === 'offline') {
          setOfflineFiles((prev) => {
            const id = `${file.name}_${file.lastModified}_${file.size}_${prev.length}`
            return [...prev, { id, file: uploaded }]
          })
          continue
        }

      }

      event.target.value = ''
    } catch (error) {
      console.error(error)
      setErrorMessage(`读取文件失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  function removeOrdersFile(itemId: string) {
    setErrorMessage('')
    setOrdersFiles((prev) => prev.filter((item) => item.id !== itemId))
  }

  function removeFreightFile(itemId: string) {
    setErrorMessage('')
    setFreightFiles((prev) => prev.filter((item) => item.id !== itemId))
  }

  function removeSimpleUploadFile(itemId: string, table: 'alipay' | 'offline') {
    setErrorMessage('')
    if (table === 'alipay') {
      setAlipayFiles((prev) => prev.filter((item) => item.id !== itemId))
      return
    }
    setOfflineFiles((prev) => prev.filter((item) => item.id !== itemId))
  }

  function removeIncomeFile(itemId: string, table: 'income' | 'refund' = 'income') {
    setErrorMessage('')
    const setter = table === 'income' ? setIncomeFiles : setRefundFiles
    setter((prev) => prev.filter((item) => item.id !== itemId))
  }

  function canProcess(): boolean {
    return canProcessCalculation({
      ordersFilesLength: ordersFiles.length,
      incomeFiles,
      refundFiles,
      effectiveIncomeDetailAmountColumnsLength: effectiveIncomeDetailAmountColumns.length,
      effectiveRefundDetailAmountColumnsLength: effectiveRefundDetailAmountColumns.length,
      ordersIdColumn,
      effectiveOrdersStatusColumn
    })
  }

  function getProcessDisabledReason(): string {
    return getProcessDisabledReasonText({
      ordersFilesLength: ordersFiles.length,
      incomeFiles,
      refundFiles,
      effectiveIncomeDetailAmountColumnsLength: effectiveIncomeDetailAmountColumns.length,
      effectiveRefundDetailAmountColumnsLength: effectiveRefundDetailAmountColumns.length,
      ordersIdColumn,
      effectiveOrdersStatusColumn
    })
  }

  function runCalculation() {
    if (canOperateWorkflow && !normalizeCellValue(workflowPeriod)) {
      setErrorMessage('请先选择业绩月份（北京时间）后再计算。')
      return
    }

    if (ordersFiles.length === 0 || incomeFiles.length === 0 || refundFiles.length === 0) {
      setErrorMessage('请先上传订单表、订单明细/收支明细、放退款明细三个区块的表。')
      return
    }

    if (
      !ordersIdColumn ||
      !effectiveOrdersStatusColumn ||
      effectiveIncomeDetailAmountColumns.length === 0 ||
      effectiveRefundDetailAmountColumns.length === 0
    ) {
      setErrorMessage('请先确认订单表、订单明细/收支明细、放退款明细的关键字段。')
      return
    }

    const incomeContentError = validateDetailFiles(incomeFiles, '订单明细/收支明细')
    if (incomeContentError) {
      setErrorMessage(incomeContentError)
      return
    }
    const refundContentError = validateDetailFiles(refundFiles, '放退款订单明细')
    if (refundContentError) {
      setErrorMessage(refundContentError)
      return
    }

    setIsProcessing(true)
    setErrorMessage('')

    try {
      const nextResult = calculatePerformanceResult({
        ordersSheetRows,
        freightSheetRows,
        incomeFiles,
        refundFiles,
        alipayFiles,
        offlineFiles,
        ordersIdColumn,
        effectiveOrdersStatusColumn,
        effectiveOrdersTimeColumn,
        effectiveIncomeDetailAmountColumns,
        effectiveRefundDetailAmountColumns,
        effectiveRefundTypeColumn,
        effectiveRefundProductNameColumn,
        effectiveRefundSkuIdColumn,
        freightOrderColumn,
        effectiveFreightAmountCnyColumn,
        effectiveFreightAmountUsdColumn,
        effectiveFreightFulfillmentColumn,
        effectiveFreightWaybillColumn,
        shopName,
        usdExchangeRate
      })

      setResult(nextResult)
      setLastCalculatedAt(formatBeijingDateTime(new Date()))
      setCalculationCount((prev) => prev + 1)
    } catch (error) {
      console.error(error)
      setErrorMessage(`计算失败：${error instanceof Error ? error.message : '未知错误'}`)
    } finally {
      setIsProcessing(false)
    }
  }

  async function handleUploadCheckedPerformanceFile(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }

    setErrorMessage('')
    try {
      const buffer = await file.arrayBuffer()
      const workbook = readWorkbookFromUpload(file.name, buffer)
      const uploaded = toUploadedState(file.name, workbook)
      const firstSheet = uploaded.sheets[0]
      if (!firstSheet || firstSheet.rows.length === 0) {
        throw new Error('上传文件为空，或未识别到有效数据行')
      }

      setUploadedPerformanceRows(firstSheet.rows)
      setUploadedPerformanceFileName(file.name)
      event.target.value = ''
      window.alert(`绩效文件已加载：${file.name}（${firstSheet.rows.length} 行）`)
    } catch (error) {
      window.alert(`读取绩效文件失败：${error instanceof Error ? error.message : '未知错误'}`)
    }
  }

  async function createAndPushPerformanceWorkflow() {
    if (!canOperateWorkflow) {
      setErrorMessage('当前角色无权限推送绩效核对。')
      return
    }
    if (!shopId) {
      setErrorMessage('请先选择店铺后再推送绩效核对。')
      return
    }
    if (!workflowPeriod) {
      setErrorMessage('请填写绩效期间（例如 2026-06）。')
      return
    }
    if (uploadedPerformanceRows.length === 0) {
      setErrorMessage('请先上传财务核对后的绩效文件。')
      return
    }
    let resolvedReviewerUserId = selectedReviewerUserId
    if (!resolvedReviewerUserId) {
      const selectedId = normalizeCellValue(selectedEmployeeId)
      if (selectedId) {
        resolvedReviewerUserId = pickDefaultReviewerByEmployee(selectedId)
        if (!resolvedReviewerUserId) {
          const selectedEmployee = employees.find((item) => item.id === selectedId)
          const displayName = normalizeCellValue(selectedEmployee?.name) || normalizeCellValue(selectedEmployee?.employeeCode) || '当前选择员工'
          const linkedUserId = normalizeCellValue(selectedEmployee?.userId)

          let reason = '该员工未匹配到可用账号。'
          if (!linkedUserId) {
            reason = '该员工档案未关联账号（employee.userId 为空）。'
          } else {
            const linkedUser = managedUsers.find((item) => item.id === linkedUserId)
            if (!linkedUser) {
              reason = '该员工关联的账号不存在或不在当前公司。'
            } else if (linkedUser.isActive === false) {
              reason = '该员工关联账号已停用。'
                      }
          }

          setErrorMessage(`你已选择员工“${displayName}”，但无法推送核对：${reason}`)
          return
        }
      } else {
        const boundEmployeeIds = branchStores.find((item) => item.id === shopId)?.employeeIds || []
        const uniqueCandidateEmployeeIds = Array.from(new Set(boundEmployeeIds.filter(Boolean)))
        for (const employeeId of uniqueCandidateEmployeeIds) {
          const matched = pickDefaultReviewerByEmployee(employeeId)
          if (matched) {
            resolvedReviewerUserId = matched
            break
          }
        }
      }
    }

    if (!resolvedReviewerUserId) {
      setErrorMessage('未匹配到可核对员工账号，请检查员工是否已关联启用中的账号（employee.userId），或姓名/编号与账号名是否可匹配。')
      return
    }

    setIsWorkflowSubmitting(true)
    setErrorMessage('')
    try {
      const createPayload = await tenantService.createWorkflow({
        storeId: shopId,
        period: workflowPeriod,
        summary: result?.summary || {},
        calculatedRows: result?.aggregatedRows || [],
        uploadedRows: uploadedPerformanceRows
      })

      const workflow = (createPayload.data as { workflow?: Record<string, unknown> } | undefined)?.workflow
      const workflowId = normalizeCellValue(workflow?._id || workflow?.id)
      if (!workflowId) {
        throw new Error('创建流程成功但缺少 workflowId')
      }

      if (resolvedReviewerUserId !== selectedReviewerUserId) {
        setSelectedReviewerUserId(resolvedReviewerUserId)
      }

      await tenantService.pushWorkflow(workflowId, resolvedReviewerUserId)

      window.alert('绩效已推送给指定员工，等待其核对确认。')
      await refreshWorkflowData()
    } catch (error) {
      window.alert(`推送绩效核对失败：${error instanceof Error ? error.message : '未知错误'}`)
    } finally {
      setIsWorkflowSubmitting(false)
    }
  }

  async function confirmPerformanceWorkflow(workflowId: string) {
    if (!workflowId) {
      return
    }

    setIsWorkflowSubmitting(true)
    setErrorMessage('')
    try {
      await tenantService.confirmWorkflow(workflowId)
      window.alert('已完成核对确认，等待财务归档。')
      await refreshWorkflowData()
    } catch (error) {
      window.alert(`核对确认失败：${error instanceof Error ? error.message : '未知错误'}`)
    } finally {
      setIsWorkflowSubmitting(false)
    }
  }

  async function archivePerformanceWorkflow(workflowId: string) {
    if (!canOperateWorkflow) {
      setErrorMessage('当前角色无权限归档绩效。')
      return
    }
    if (!workflowId) {
      return
    }

    setIsWorkflowSubmitting(true)
    setErrorMessage('')
    try {
      await tenantService.archiveWorkflow(workflowId)
      window.alert('绩效已落库存档。')
      await refreshWorkflowData()
    } catch (error) {
      window.alert(`落库存档失败：${error instanceof Error ? error.message : '未知错误'}`)
    } finally {
      setIsWorkflowSubmitting(false)
    }
  }

  function exportAggregatedWorkbook() {
    if (!result) {
      setErrorMessage('请先执行计算，再导出结果。')
      return
    }

    const workbook = XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.aggregatedRows),
      '订单聚合表'
    )

    const dateText = toDateText(new Date())
    XLSX.writeFile(workbook, getExportFileName('订单聚合表', dateText, { shopName, shopId, subsidiary }))
  }

  function exportBusinessDetailWorkbook() {
    if (!result) {
      setErrorMessage('请先执行计算，再导出结果。')
      return
    }

    const workbook = XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.performanceRows),
      '业务明细表'
    )

    const dateText = toDateText(new Date())
    XLSX.writeFile(workbook, getExportFileName('业务明细表', dateText, { shopName, shopId, subsidiary }))
  }

  function exportOtherSheetsWorkbook() {
    if (!result) {
      setErrorMessage('请先执行计算，再导出结果。')
      return
    }

    const workbook = XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.ordersWithoutIncomeRows),
      '订单表有但收支缺失'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.incomeOnlyOrdersRows),
      '仅订单明细有'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.refundOnlyOrdersRows),
      '仅放退款明细有'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.ordersWithoutFreightRows),
      '无线上线下运费订单'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.ordersWithoutAlipayRows),
      '订单有但支付宝缺失'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.unmatchedAlipayRows),
      '支付宝无对应订单'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.alipayMultiplicityRows),
      '支付宝匹配异常'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(result.offlineFreightDiagnosticRows),
      '线下运费解析诊断'
    )

    const dateText = toDateText(new Date())
    XLSX.writeFile(workbook, getExportFileName('其他对账表', dateText, { shopName, shopId, subsidiary }))
  }

  async function handleLogin(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault()
    setErrorMessage('')

    const apiBaseUrl = normalizeCellValue(authApiBaseUrl)
    if (!apiBaseUrl) {
      setErrorMessage('系统未配置 API 地址，请联系管理员。')
      return
    }
    if (!loginAccount || !loginPassword) {
      setErrorMessage('请输入账号和密码。')
      return
    }

    setIsLoginSubmitting(true)
    try {
      const payload = await authService.loginByAccount({
        account: loginAccount.trim(),
        password: loginPassword
      })

      const loginData = payload.data as { token?: unknown; user?: AuthUser } | undefined
      const token = normalizeCellValue(loginData?.token)
      const user = loginData?.user
      if (!token || !user) {
        throw new Error('登录响应缺少 token 或 user 信息')
      }

      setAuthToken(token)
      setAuthUser(user)

      localStorage.setItem(AUTH_API_BASE_URL_KEY, apiBaseUrl)
      localStorage.setItem(AUTH_TOKEN_STORAGE_KEY, token)
      localStorage.setItem(AUTH_USER_STORAGE_KEY, JSON.stringify(user))
      localStorage.removeItem(LEGACY_AUTH_API_BASE_URL_KEY)
      localStorage.removeItem(LEGACY_AUTH_TOKEN_STORAGE_KEY)
      localStorage.removeItem(LEGACY_AUTH_USER_STORAGE_KEY)
      setErrorMessage('')
    } catch (error) {
      setErrorMessage(`登录失败：${error instanceof Error ? error.message : '未知错误'}`)
    } finally {
      setIsLoginSubmitting(false)
    }
  }

  function handleLogout() {
    resetAuthSession()
    setErrorMessage('已退出登录。')
  }

  function enterGuestMode() {
    setIsGuestMode(true)
    localStorage.setItem(AUTH_GUEST_MODE_KEY, '1')
    localStorage.removeItem(LEGACY_AUTH_GUEST_MODE_KEY)
    setErrorMessage('已进入本地模式：可计算和导出，本模式下不上传 SaaS。')
  }

  if ((!authToken || !authUser) && !isGuestMode) {
    return (
      <main className="page login-page">
        <section className="login-shell">
          <header className="page-header login-header">
            <h2>业绩计算工作台</h2>
          </header>

          <section className="upload-card login-card">
            <form className="auth-form" onSubmit={handleLogin}>
              <div className="control-row">
                <label>账号</label>
                <input
                  className="compact-input"
                  type="text"
                  value={loginAccount}
                  onChange={(event) => setLoginAccount(event.target.value)}
                  placeholder="请输入账号，例如：super_admin"
                />
              </div>
              <div className="control-row">
                <label>密码</label>
                <input
                  className="compact-input"
                  type="password"
                  value={loginPassword}
                  onChange={(event) => setLoginPassword(event.target.value)}
                  placeholder="请输入密码"
                />
              </div>
              <div className="auth-actions">
                <button type="submit" className="auth-primary-button" disabled={isLoginSubmitting}>
                  {isLoginSubmitting ? '登录中...' : '登录并进入工作台'}
                </button>
                <button type="button" className="auth-secondary-button" onClick={enterGuestMode}>
                  我没有账号，先进入本地模式
                </button>
              </div>
            </form>
          </section>

          {errorMessage && (
            <div className="error-modal-backdrop" role="presentation" onClick={() => setErrorMessage('')}>
              <div className="error-box error-modal" role="alertdialog" aria-modal="true" onClick={(event) => event.stopPropagation()}>
                <h3>操作提示</h3>
                <p>{errorMessage}</p>
                <div className="error-modal-actions">
                  <button type="button" className="ghost" onClick={() => setErrorMessage('')}>我知道了</button>
                </div>
              </div>
            </div>
          )}
        </section>
      </main>
    )
  }

  return (
    <main className="workspace-shell">
      <aside className="workspace-sidebar" aria-label="主菜单">
        <div>
            <div className="workspace-brand">
          <div className="workspace-user-box">
          <div className="header-user-row">
            <span>用户</span>
            <strong>{displayPersonName}</strong>
          </div>
          <div className="header-user-row">
            <span>账号</span>
            <strong>{displayAccountName}</strong>
          </div>
          <div className="header-user-row">
            <span>总公司</span>
            <strong>{displayTotalCompany}</strong>
          </div>
          <div className="header-user-row">
            <span>身份</span>
            <strong>{displayRole}</strong>
          </div>
        </div>
        </div>

        <nav className="workspace-menu" role="navigation" aria-label="功能菜单">
          {canAccessCalculator && (
            <button
              type="button"
              className={`workspace-menu-item ${resolvedActiveView === 'calculator' ? 'is-active' : ''}`}
              onClick={() => setActiveView('calculator')}
            >
              业绩计算
            </button>
          )}
          {canAccessPerformanceManagement && (
            <div className="workspace-menu-group">
              <button
                type="button"
                className={`workspace-menu-item ${resolvedActiveView === 'performance' ? 'is-active with-children' : 'with-children'}`}
                onClick={() => setActiveView('performance')}
              >
                绩效管理
                <span className="workspace-menu-caret">⌄</span>
              </button>
              {resolvedActiveView === 'performance' && (
                <div className="workspace-submenu-list">
                  <button
                    type="button"
                    className={`workspace-submenu-item ${performanceListTab === 'current' ? 'is-active' : ''}`}
                    onClick={() => setPerformanceListTab('current')}
                  >
                    当月绩效
                  </button>
                  <button
                    type="button"
                    className={`workspace-submenu-item ${performanceListTab === 'history' ? 'is-active' : ''}`}
                    onClick={() => setPerformanceListTab('history')}
                  >
                    历史绩效
                  </button>
                </div>
              )}
            </div>
          )}
          {canAccessMyPerformance && (
            <button
              type="button"
              className={`workspace-menu-item ${resolvedActiveView === 'my-performance' ? 'is-active' : ''}`}
              onClick={() => setActiveView('my-performance')}
            >
              我的业绩
            </button>
          )}
          {canAccessEmployeeManagement && (
            <div className="workspace-menu-group">
              <button
                type="button"
                className={`workspace-menu-item ${resolvedActiveView === 'employee' ? 'is-active with-children' : 'with-children'}`}
                onClick={() => setActiveView('employee')}
              >
                公司管理
                <span className="workspace-menu-caret">⌄</span>
              </button>
              {resolvedActiveView === 'employee' && (
                <div className="workspace-submenu-list">
                  <button
                    type="button"
                    className={`workspace-submenu-item ${employeeCompanyTab === 'subsidiary' ? 'is-active' : ''}`}
                    onClick={() => setEmployeeCompanyTab('subsidiary')}
                  >
                    分公司管理
                  </button>
                  <button
                    type="button"
                    className={`workspace-submenu-item ${employeeCompanyTab === 'store' ? 'is-active' : ''}`}
                    onClick={() => setEmployeeCompanyTab('store')}
                  >
                    店铺管理
                  </button>
                  <button
                    type="button"
                    className={`workspace-submenu-item ${employeeCompanyTab === 'employee' ? 'is-active' : ''}`}
                    onClick={() => setEmployeeCompanyTab('employee')}
                  >
                    员工管理
                  </button>

                  {employeeCompanyTab === 'employee' && (
                    <div className="workspace-submenu-list workspace-submenu-list-nested">
                      <button
                        type="button"
                        className={`workspace-submenu-item ${employeeStaffTab === 'create' ? 'is-active' : ''}`}
                        onClick={() => setEmployeeStaffTab('create')}
                      >
                        新增员工
                      </button>
                      <button
                        type="button"
                        className={`workspace-submenu-item ${employeeStaffTab === 'list' ? 'is-active' : ''}`}
                        onClick={() => setEmployeeStaffTab('list')}
                      >
                        员工列表
                      </button>
                      <button
                        type="button"
                        className={`workspace-submenu-item ${employeeStaffTab === 'bind' ? 'is-active' : ''}`}
                        onClick={() => setEmployeeStaffTab('bind')}
                      >
                        绑定店铺
                      </button>
                    </div>
                  )}
                </div>
              )}
            </div>
          )}
          {canAccessAdminPage && (
            <button
              type="button"
              className={`workspace-menu-item ${resolvedActiveView === 'admin' ? 'is-active' : ''}`}
              onClick={() => setActiveView('admin')}
            >
              管理员页面
            </button>
          )}
        </nav>

        </div>

         <button
            type="button"
            className="workspace-menu-item workspace-menu-item-logout"
            onClick={handleLogout}
          >
            退出登录
          </button>

      </aside>

      <section className="page workspace-page">

        {errorMessage && (
          <div className="error-modal-backdrop" role="presentation" onClick={() => setErrorMessage('')}>
            <div className="error-box error-modal" role="alertdialog" aria-modal="true" onClick={(event) => event.stopPropagation()}>
              <h3>操作提示</h3>
              <p>{errorMessage}</p>
              <div className="error-modal-actions">
                <button type="button" className="ghost" onClick={() => setErrorMessage('')}>我知道了</button>
              </div>
            </div>
          </div>
        )}

        {successToast && (
          <div className="success-toast" role="status" aria-live="polite">
            {successToast}
          </div>
        )}

        {resolvedActiveView === 'admin' && canAccessAdminPage && (
          <section className="upload-card onboarding-shell">
            <h2>超管页面</h2>
            <p>仅超级管理员可访问：创建总公司、创建公司管理员、查看公司入驻列表。</p>
            <AdminManagementSection
              companyProfile={adminCompanyProfile}
              companyList={adminCompanies}
              canCreateCompany={actualRole === 'super_admin'}
              managedUsers={managedUsers}
              branchStores={branchStores}
              isLoading={isTenantDataLoading}
              onCreateCompanyProfile={handleCreateCompanyProfile}
              onAddManagedUser={handleAddManagedUser}
              onAddBranchStore={handleAddBranchStore}
            />
          </section>
        )}

        {resolvedActiveView === 'employee' && canAccessEmployeeManagement && (
          <EmployeeManagementSection
            companyProfile={adminCompanyProfile}
            managedUsers={managedUsers}
            employees={employees}
            branchStores={branchStores}
            performanceSnapshots={performanceSnapshots}
            canEditCompany={canAccessEmployeeManagement}
            canEditCompanyPolicy={false}
            isLoading={isTenantDataLoading}
            companyTab={employeeCompanyTab}
            employeeTab={employeeStaffTab}
            onAddStaff={handleAddStaff}
            onSaveSubsidiaries={handleSaveSubsidiaries}
            onCreateStore={handleCreateStoreWithAssignments}
            onBindEmployeeStores={handleBindEmployeeStores}
            onResignEmployee={handleResignEmployee}
            onUpdateCompanyProfile={handleUpdateCompanyProfile}
          />
        )}

        {resolvedActiveView === 'my-performance' && canAccessMyPerformance && (
          <MyPerformanceView
            pendingWorkflows={employeePendingWorkflows}
            historyWorkflows={employeeHistoryWorkflows}
            workflowStatusLabelMap={workflowStatusLabelMap}
            isSubmitting={isWorkflowSubmitting}
            onConfirmWorkflow={(workflowId) => {
              void confirmPerformanceWorkflow(workflowId)
            }}
          />
        )}

        {resolvedActiveView === 'calculator' && canAccessCalculator && (
                    
          <>
          <header className="page-header with-plugin-download">
            <a className="header-plugin-download" href={CHROME_EXPORT_PLUGIN_ZIP_PATH} download>
              下载 Chrome 导出插件 ZIP
            </a>
            <h2>业绩计算工作台</h2>
          <p>
            上传订单表、订单明细/收支明细表、放退款订单明细、线上运费、支付宝记录、线下发货记录后，系统会自动筛选并生成业绩与对账结果。
          </p>
          <div className="actions-row">
            <button type="button" className="ghost" onClick={clearLocalUploadCache}>
              清空本地缓存
            </button>
          </div>
          </header>

            <div className="admin-form-grid admin-form-grid-3 common-info">
              <div className="control-row">
                <label>业绩月份</label>
                <input
                  type="month"
                  className="compact-input"
                  value={workflowPeriod}
                  onChange={(event) => setWorkflowPeriod(event.target.value)}
                  placeholder="例如：2026-06"
                  disabled={isEmployeeReviewOnly}
                />
              </div>
              <div className="control-row">
                <label>美元汇率</label>
                <input
                  className="compact-input"
                  value={usdExchangeRate}
                  onChange={(event) => setUsdExchangeRate(event.target.value)}
                  placeholder="例如：6.8"
                  disabled={isEmployeeReviewOnly}
                />
              </div>

            </div>

          <StoreMetaSection
            shopId={shopId}
            subsidiary={subsidiary}
            selectedEmployeeId={selectedEmployeeId}
            subsidiaryOptions={subsidiaryOptions}
            employeeOptions={employeeOptions}
            storeOptions={storeOptions}
            selectedBusinessLicenseName={selectedBusinessLicenseName}
            isLoading={isStoreEmployeeOptionsLoading}
            readOnly={isEmployeeReviewOnly}
            onSubsidiaryChange={(value) => {
              setSubsidiary(value)
              setSelectedEmployeeId('')
              setSelectedReviewerUserId('')
              setShopId('')
              setShopName('')
            }}
            onEmployeeChange={(value) => {
              setSelectedEmployeeId(value)
              setSelectedReviewerUserId(pickDefaultReviewerByEmployee(value))
              setShopId('')
              setShopName('')
            }}
            onStoreChange={(store) => {
              setShopId(store.id)
              setShopName(store.name)
            }}
          />


          {!isEmployeeReviewOnly && (
          <section className="upload-sections">
            <OrdersUploadCardInternal
              files={ordersFiles}
              onFileUpload={handleFileUpload}
              onRemove={removeOrdersFile}
            />
            <FreightUploadCardInternal
              files={freightFiles}
              onFileUpload={handleFileUpload}
              onRemove={removeFreightFile}
            />
            <DetailUploadCardInternal
              title="3) 订单明细/收支明细表"
              description={
                <>
                  请上传
                  <span className="emphasis-russia">【俄罗斯】结算月至今的订单明细表；【非俄罗斯】结算月至今的订单收支明细表。</span>
                </>
              }
              table="income"
              files={incomeFiles}
              emptyText="尚未上传订单明细/收支明细文件"
              onFileUpload={handleFileUpload}
              onRemove={removeIncomeFile}
            />
            <DetailUploadCardInternal
              title="4) 放退款订单明细"
              description={
                <>
                  请上传
                  <span className="emphasis-russia">【俄罗斯】和【非俄罗斯】结算月至今</span>
                  的放退款明细文件。
                </>
              }
              table="refund"
              files={refundFiles}
              emptyText="尚未上传放退款订单明细文件"
              onFileUpload={handleFileUpload}
              onRemove={removeIncomeFile}
            />
            <SimpleUploadCardInternal
              title="5) 支付宝订单记录（采购）"
              description="支持多文件上传。备注中自动提取连续16位数字作为订单号。"
              table="alipay"
              files={alipayFiles}
              emptyText="尚未上传支付宝订单记录"
              onFileUpload={handleFileUpload}
              onRemove={removeSimpleUploadFile}
            />
            <SimpleUploadCardInternal
              title="6) 线下发货订单记录"
              description="支持多文件上传。备注按“店铺名-订单号”匹配，系统会自动摘取短线后的订单号。"
              table="offline"
              files={offlineFiles}
              emptyText="尚未上传线下发货记录"
              onFileUpload={handleFileUpload}
              onRemove={removeSimpleUploadFile}
            />
          </section>
          )}

          {!isEmployeeReviewOnly && (
          <ActionPanel
            isProcessing={isProcessing}
            canProcess={canProcess()}
            lastCalculatedAt={lastCalculatedAt}
            calculationCount={calculationCount}
            processDisabledReason={getProcessDisabledReason()}
            onRunCalculation={runCalculation}
          />
          )}

          {!isEmployeeReviewOnly && result && (
            <ResultPreviewPanel
              result={result}
              onExportAggregated={exportAggregatedWorkbook}
              onExportBusinessDetail={exportBusinessDetailWorkbook}
              onExportOtherSheets={exportOtherSheetsWorkbook}
            />
          )}

          {canOperateWorkflow && (
            <PerformanceManagementView
              mode="push"
              listTab={performanceListTab}
              uploadedPerformanceFileName={uploadedPerformanceFileName}
              workflowReviewers={workflowReviewers}
              uploadedPerformanceRows={uploadedPerformanceRows}
              performanceWorkflows={performanceWorkflows}
              financeArchiveReadyWorkflows={financeArchiveReadyWorkflows}
              workflowStatusLabelMap={workflowStatusLabelMap}
              isWorkflowSubmitting={isWorkflowSubmitting}
              isWorkflowListLoading={isWorkflowListLoading}
              onUploadCheckedPerformanceFile={handleUploadCheckedPerformanceFile}
              onCreateAndPushPerformanceWorkflow={() => {
                void createAndPushPerformanceWorkflow()
              }}
              onRefreshWorkflowData={() => {
                void refreshWorkflowData(true, true)
              }}
              onArchivePerformanceWorkflow={(workflowId) => {
                void archivePerformanceWorkflow(workflowId)
              }}
            />
          )}

          </>
        )}

        {resolvedActiveView === 'performance' && canAccessPerformanceManagement && (
          <PerformanceManagementView
            mode="list"
            listTab={performanceListTab}
            uploadedPerformanceFileName={uploadedPerformanceFileName}
            workflowReviewers={workflowReviewers}
            uploadedPerformanceRows={uploadedPerformanceRows}
            performanceWorkflows={performanceWorkflows}
            financeArchiveReadyWorkflows={financeArchiveReadyWorkflows}
            workflowStatusLabelMap={workflowStatusLabelMap}
            isWorkflowSubmitting={isWorkflowSubmitting}
            isWorkflowListLoading={isWorkflowListLoading}
            onUploadCheckedPerformanceFile={handleUploadCheckedPerformanceFile}
            onCreateAndPushPerformanceWorkflow={() => {
              void createAndPushPerformanceWorkflow()
            }}
            onRefreshWorkflowData={() => {
              void refreshWorkflowData(true, true)
            }}
            onArchivePerformanceWorkflow={(workflowId) => {
              void archivePerformanceWorkflow(workflowId)
            }}
          />
        )}
      </section>
    </main>
  )
}

export default App

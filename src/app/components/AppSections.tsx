import { useCallback, useEffect, useMemo, useState } from 'react'
import type * as Domain from '../calculatorDomain'
import { normalizeCellValue, normalizeOrderNo } from '../calculatorDomain'

type ProcessResult = Domain.ProcessResult

export type AdminCompanyProfile = {
  id: string
  companyName: string
  companyCode: string
  legalRepresentative: string
  subsidiaries?: string[]
  createdAt: string
  expireDate?: string
  maxUsers?: number
  totalUsers?: number
  activeUsers?: number
  adminUsers?: number
  status?: string
}

export type AdminManagedUser = {
  id: string
  username: string
  email: string
  role: string
  createdAt: string
}

export type AdminBranchStore = {
  id: string
  subsidiary: string
  shopName: string
  storeIdOnPlatform: string
  businessLicenseName: string
  assignedEmployeeNames: string[]
  employeeIds: string[]
  status?: string
  createdAt: string
}

export type AdminEmployee = {
  id: string
  name: string
  employeeCode: string
  status: string
  createdAt: string
}

export type PerformanceSnapshot = {
  id: string
  period: string
  storeLabel: string
  rowCount: number
  createdAt: string
}

function formatRoleLabel(role: string): string {
  const normalized = normalizeCellValue(role)
  const roleMap: Record<string, string> = {
    company_admin: '公司管理者',
    finance: '财务',
    branch_manager: '分公司总经理',
    team_lead: '组长',
    employee: '普通员工'
  }
  return roleMap[normalized] || normalized || '-'
}

type StoreMetaSectionProps = {
  shopId: string
  subsidiary: string
  selectedEmployeeId: string
  subsidiaryOptions: string[]
  employeeOptions: Array<{ id: string; label: string }>
  storeOptions: Array<{ id: string; name: string }>
  selectedBusinessLicenseName: string
  isLoading?: boolean
  readOnly?: boolean
  onSubsidiaryChange: (value: string) => void
  onEmployeeChange: (value: string) => void
  onStoreChange: (store: { id: string; name: string }) => void
}

export function StoreMetaSection(props: StoreMetaSectionProps) {
  const {
    shopId,
    subsidiary,
    selectedEmployeeId,
    subsidiaryOptions,
    employeeOptions,
    storeOptions,
    selectedBusinessLicenseName,
    isLoading = false,
    readOnly = false,
    onSubsidiaryChange,
    onEmployeeChange,
    onStoreChange
  } = props

  return (
    <section className="upload-card store-meta-card">
      <h2>店铺信息</h2>
      <div className="store-meta-grid">
        <div className="control-row">
          <label>所属分公司</label>
          <select
            className="compact-select"
            value={subsidiary}
            onChange={(event) => onSubsidiaryChange(event.target.value)}
            disabled={readOnly || isLoading}
          >
            <option value="">{isLoading ? '加载中...' : '请选择分公司'}</option>
            {subsidiaryOptions.map((item) => (
              <option key={item} value={item}>{item}</option>
            ))}
          </select>
        </div>
        <div className="control-row">
          <label>人员</label>
          <select
            className="compact-select"
            value={selectedEmployeeId}
            onChange={(event) => onEmployeeChange(event.target.value)}
            disabled={readOnly || isLoading || !subsidiary}
          >
            <option value="">{isLoading ? '加载中...' : subsidiary ? '请选择人员' : '请先选择分公司'}</option>
            {employeeOptions.map((item) => (
              <option key={item.id} value={item.id}>{item.label}</option>
            ))}
          </select>
        </div>
        <div className="control-row">
          <label>店铺</label>
          <div className="store-select-inline">
            <select
              className="compact-select"
              value={shopId}
              onChange={(event) => {
                const nextId = normalizeCellValue(event.target.value)
                const next = storeOptions.find((item) => item.id === nextId)
                if (next) {
                  onStoreChange(next)
                  return
                }
                onStoreChange({ id: '', name: '' })
              }}
              disabled={readOnly || isLoading || !subsidiary || !selectedEmployeeId}
            >
              <option value="">{isLoading ? '加载中...' : subsidiary && selectedEmployeeId ? '请选择店铺' : '请先选择分公司和人员'}</option>
              {storeOptions.map((item) => (
                <option key={item.id} value={item.id}>{item.name}</option>
              ))}
            </select>
            {shopId && (
              <span className="store-license-inline">
                营业执照：{selectedBusinessLicenseName || '未设置'}
              </span>
            )}
          </div>
        </div>
      </div>
      {isLoading && <p className="hint-text">加载中...</p>}
    </section>
  )
}

type AdminManagementSectionProps = {
  companyProfile: AdminCompanyProfile | null
  companyList: AdminCompanyProfile[]
  canCreateCompany: boolean
  managedUsers: AdminManagedUser[]
  branchStores: AdminBranchStore[]
  isLoading?: boolean
  onCreateCompanyProfile: (input: {
    companyName: string
    legalRepresentative: string
    adminUsername: string
    adminPassword: string
    adminEmail: string
  }) => void
  onAddManagedUser: (input: { username: string; email: string; password: string; role: string }) => void
  onAddBranchStore: (input: { subsidiary: string; shopName: string; businessLicense: string }) => void
}

export function AdminManagementSection(props: AdminManagementSectionProps) {
  const {
    companyProfile,
    companyList,
    canCreateCompany,
    managedUsers,
    branchStores,
    isLoading = false,
    onCreateCompanyProfile,
    onAddManagedUser,
    onAddBranchStore
  } = props

  const [companyName, setCompanyName] = useState('')
  const [legalRepresentative, setLegalRepresentative] = useState('')
  const [adminUsername, setAdminUsername] = useState('')
  const [adminPassword, setAdminPassword] = useState('')
  const [adminEmail, setAdminEmail] = useState('')

  const [newUsername, setNewUsername] = useState('')
  const [newUserEmail, setNewUserEmail] = useState('')
  const [newUserPassword, setNewUserPassword] = useState('')
  const [newUserRole, setNewUserRole] = useState('finance')

  const [subsidiary, setSubsidiary] = useState('')
  const [shopNameInput, setShopNameInput] = useState('')
  const [businessLicense, setBusinessLicense] = useState('')

  const handleCreateCompanyProfile = () => {
    if (!canCreateCompany) {
      return
    }
    onCreateCompanyProfile({
      companyName,
      legalRepresentative,
      adminUsername,
      adminPassword,
      adminEmail
    })
    setCompanyName('')
    setLegalRepresentative('')
    setAdminUsername('')
    setAdminPassword('')
    setAdminEmail('')
  }

  const handleAddUser = () => {
    onAddManagedUser({
      username: newUsername,
      email: newUserEmail,
      password: newUserPassword,
      role: newUserRole
    })
    setNewUsername('')
    setNewUserEmail('')
    setNewUserPassword('')
    setNewUserRole('finance')
  }

  const handleAddBranchStore = () => {
    onAddBranchStore({
      subsidiary,
      shopName: shopNameInput,
      businessLicense
    })
    setSubsidiary('')
    setShopNameInput('')
    setBusinessLicense('')
  }

  return (
    <section className="upload-card admin-section">
      <h2>管理员配置</h2>
      <p>管理员可一次性录入公司信息（录入后锁定），并维护用户、分公司与店铺营业执照。</p>
      {isLoading && <p className="hint-text">加载中...</p>}

      <div className="admin-grid">
        <section className="admin-card">
          <h3>公司列表</h3>
          <div className="admin-list">
            {companyList.length === 0 && <p className="hint-text">{isLoading ? '加载中...' : '暂无公司记录'}</p>}
            {companyList.map((item) => (
              <div key={item.id} className="meta-row admin-list-item">
                <span className="meta-key">
                  {item.companyName}
                  {item.status ? `（${item.status}）` : ''}
                </span>
                <span className="meta-value">
                  用户 {item.totalUsers || 0} / 活跃 {item.activeUsers || 0} / 管理员 {item.adminUsers || 0}
                </span>
              </div>
            ))}
          </div>

          {companyProfile && (
            <div className="admin-readonly-list" style={{ marginTop: 12 }}>
              <div className="meta-row">
                <span className="meta-key">当前公司</span>
                <span className="meta-value">{companyProfile.companyName}</span>
              </div>
              <div className="meta-row">
                <span className="meta-key">法人/负责人</span>
                <span className="meta-value">{companyProfile.legalRepresentative || '-'}</span>
              </div>
            </div>
          )}
        </section>

        <section className="admin-card">
          <h3>新建公司</h3>
          {canCreateCompany ? (
            <div className="admin-form-grid">
              <div className="control-row">
                <label>公司名称</label>
                <input
                  className="compact-input"
                  value={companyName}
                  onChange={(event) => setCompanyName(event.target.value)}
                  placeholder="例如：深圳某某科技有限公司"
                />
              </div>
              <div className="control-row">
                <label>管理员账号</label>
                <input
                  className="compact-input"
                  value={adminUsername}
                  onChange={(event) => setAdminUsername(event.target.value)}
                  placeholder="例如：acme_admin"
                />
              </div>
              <div className="control-row">
                <label>管理员初始密码</label>
                <input
                  className="compact-input"
                  type="password"
                  value={adminPassword}
                  onChange={(event) => setAdminPassword(event.target.value)}
                  placeholder="至少 6 位"
                />
              </div>
              <div className="control-row">
                <label>管理员邮箱（可选）</label>
                <input
                  className="compact-input"
                  value={adminEmail}
                  onChange={(event) => setAdminEmail(event.target.value)}
                  placeholder="例如：admin@company.com"
                />
              </div>
              <div className="control-row">
                <label>法人/负责人（可选）</label>
                <input
                  className="compact-input"
                  value={legalRepresentative}
                  onChange={(event) => setLegalRepresentative(event.target.value)}
                  placeholder="例如：张三"
                />
              </div>
              <button type="button" className="ghost" onClick={handleCreateCompanyProfile}>
                新建公司并创建管理员
              </button>
            </div>
          ) : (
            <p className="hint-text">仅超级管理员可新建公司。</p>
          )}
        </section>

        <section className="admin-card">
          <h3>添加用户</h3>
          <p className="hint-text">添加用户后会自动创建同名员工档案，无需再单独建档。</p>
          <div className="admin-form-grid">
            <div className="control-row">
              <label>用户名</label>
              <input
                className="compact-input"
                value={newUsername}
                onChange={(event) => setNewUsername(event.target.value)}
                placeholder="例如：finance01"
              />
            </div>
            <div className="control-row">
              <label>邮箱</label>
              <input
                className="compact-input"
                value={newUserEmail}
                onChange={(event) => setNewUserEmail(event.target.value)}
                placeholder="例如：finance@example.com"
              />
            </div>
            <div className="control-row">
              <label>初始密码</label>
              <input
                className="compact-input"
                type="password"
                value={newUserPassword}
                onChange={(event) => setNewUserPassword(event.target.value)}
                placeholder="至少 6 位"
              />
            </div>
            <div className="control-row">
              <label>身份</label>
              <select
                className="compact-select"
                value={newUserRole}
                onChange={(event) => setNewUserRole(event.target.value)}
              >
                <option value="company_admin">公司管理者</option>
                <option value="finance">财务</option>
                <option value="branch_manager">分公司总经理</option>
                <option value="team_lead">组长</option>
                <option value="employee">普通员工</option>
              </select>
            </div>
            <button type="button" className="ghost" onClick={handleAddUser}>
              添加用户
            </button>
          </div>
          <div className="admin-list">
            {managedUsers.length === 0 && <p className="hint-text">{isLoading ? '加载中...' : '暂无用户'}</p>}
            {managedUsers.map((user) => (
              <div key={user.id} className="meta-row admin-list-item">
                <span className="meta-key">{user.username}（{formatRoleLabel(user.role)}）</span>
                <span className="meta-value">{user.email}</span>
              </div>
            ))}
          </div>
        </section>

        <section className="admin-card admin-card-full">
          <h3>分公司 / 店铺 / 营业执照</h3>
          <div className="admin-form-grid admin-form-grid-3">
            <div className="control-row">
              <label>分公司</label>
              <input
                className="compact-input"
                value={subsidiary}
                onChange={(event) => setSubsidiary(event.target.value)}
                placeholder="例如：华南分公司"
              />
            </div>
            <div className="control-row">
              <label>店铺名</label>
              <input
                className="compact-input"
                value={shopNameInput}
                onChange={(event) => setShopNameInput(event.target.value)}
                placeholder="例如：DXM官方店"
              />
            </div>
            <div className="control-row">
              <label>营业执照编号</label>
              <input
                className="compact-input"
                value={businessLicense}
                onChange={(event) => setBusinessLicense(event.target.value)}
                placeholder="例如：91440300MA5XXXXX"
              />
            </div>
            <button type="button" className="ghost" onClick={handleAddBranchStore}>
              添加分公司店铺信息
            </button>
          </div>
          <div className="admin-list">
            {branchStores.length === 0 && <p className="hint-text">{isLoading ? '加载中...' : '暂无分公司店铺记录'}</p>}
            {branchStores.map((item) => (
              <div key={item.id} className="meta-row admin-list-item">
                <span className="meta-key">{item.subsidiary} / {item.shopName}</span>
                <span className="meta-value">营业执照：{item.businessLicenseName || '-'}</span>
              </div>
            ))}
          </div>
        </section>
      </div>
    </section>
  )
}

type EmployeeManagementSectionProps = {
  companyProfile: AdminCompanyProfile | null
  managedUsers: AdminManagedUser[]
  employees: AdminEmployee[]
  branchStores: AdminBranchStore[]
  performanceSnapshots: PerformanceSnapshot[]
  canEditCompany: boolean
  canEditCompanyPolicy: boolean
  isLoading?: boolean
  companyTab: 'subsidiary' | 'store' | 'employee'
  employeeTab: 'create' | 'list' | 'bind'
  onAddStaff: (input: {
    name: string
    username: string
    email?: string
    password: string
    role: string
    employeeCode: string
    notes: string
  }) => void
  onSaveSubsidiaries: (subsidiaries: string[]) => void
  onCreateStore: (input: {
    subsidiary: string
    shopName: string
    businessLicenseName: string
    employeeIds: string[]
  }) => void
  onBindEmployeeStores: (input: { employeeId: string; storeIds: string[] }) => void
  onUpdateCompanyProfile: (input: {
    companyName: string
    legalRepresentative: string
    maxUsers: number
    expireDate: string
  }) => void
}

export function EmployeeManagementSection(props: EmployeeManagementSectionProps) {
  "use no memo";

  const {
    companyProfile,
    managedUsers,
    employees,
    branchStores,
    canEditCompany,
    canEditCompanyPolicy,
    isLoading = false,
    companyTab,
    employeeTab,
    onAddStaff,
    onSaveSubsidiaries,
    onCreateStore,
    onBindEmployeeStores,
    onUpdateCompanyProfile
  } = props

  const [staffName, setStaffName] = useState('')
  const [staffUsername, setStaffUsername] = useState('')
  const [staffEmail, setStaffEmail] = useState('')
  const [staffPassword, setStaffPassword] = useState('')
  const [staffRole, setStaffRole] = useState('employee')
  const [staffEmployeeCode, setStaffEmployeeCode] = useState('')
  const [staffNotes, setStaffNotes] = useState('')

  const [subsidiaryName, setSubsidiaryName] = useState('')
  const [subsidiaryManager, setSubsidiaryManager] = useState('')
  const [companySubsidiaries, setCompanySubsidiaries] = useState<string[]>([])

  const [storeSubsidiary, setStoreSubsidiary] = useState('总公司')
  const [storeName, setStoreName] = useState('')
  const [businessLicenseName, setBusinessLicenseName] = useState('')
  const [selectedEmployeeIds, setSelectedEmployeeIds] = useState<string[]>([])

  const [bindEmployeeId, setBindEmployeeId] = useState('')
  const [bindingStoreIds, setBindingStoreIds] = useState<string[]>([])

  const [companyName, setCompanyName] = useState(companyProfile?.companyName || '')
  const [legalRepresentative, setLegalRepresentative] = useState(companyProfile?.legalRepresentative || '')
  const [maxUsers, setMaxUsers] = useState(String(companyProfile?.maxUsers || 20))
  const [expireDate, setExpireDate] = useState(
    normalizeCellValue(companyProfile?.expireDate).slice(0, 10)
  )

  useEffect(() => {
    let cancelled = false
    queueMicrotask(() => {
      if (cancelled) {
        return
      }
      setCompanyName(companyProfile?.companyName || '')
      setLegalRepresentative(companyProfile?.legalRepresentative || '')
      setMaxUsers(String(companyProfile?.maxUsers || 20))
      setExpireDate(normalizeCellValue(companyProfile?.expireDate).slice(0, 10))
      setCompanySubsidiaries(Array.isArray(companyProfile?.subsidiaries) ? companyProfile.subsidiaries : [])
    })

    return () => {
      cancelled = true
    }
  }, [companyProfile])

  useEffect(() => {
    let cancelled = false
    queueMicrotask(() => {
      if (cancelled) {
        return
      }
      if (!bindEmployeeId) {
        setBindingStoreIds([])
        return
      }

      const assignedStoreIds = branchStores
        .filter((store) => store.employeeIds.includes(bindEmployeeId))
        .map((store) => store.id)
      setBindingStoreIds(assignedStoreIds)
    })

    return () => {
      cancelled = true
    }
  }, [bindEmployeeId, branchStores])

  const accountUsernameByEmployeeId = useMemo(() => {
    const map = new Map<string, string>()
    for (const employee of employees) {
      const employeeName = normalizeCellValue(employee.name)
      const employeeCode = normalizeCellValue(employee.employeeCode)
      if (!employeeName && !employeeCode) {
        continue
      }

      const exact = managedUsers.find((user) => {
        const username = normalizeCellValue(user.username)
        return (employeeName && username === employeeName) || (employeeCode && username === employeeCode)
      })
      if (exact) {
        map.set(employee.id, normalizeCellValue(exact.username))
        continue
      }

      const fuzzy = managedUsers.find((user) => {
        const username = normalizeCellValue(user.username)
        return (employeeName && username.includes(employeeName)) || (employeeCode && username.includes(employeeCode))
      })
      if (fuzzy) {
        map.set(employee.id, normalizeCellValue(fuzzy.username))
      }
    }
    return map
  }, [employees, managedUsers])

  const formatEmployeeDisplayLabel = useCallback((employee: AdminEmployee): string => {
    const account = normalizeCellValue(accountUsernameByEmployeeId.get(employee.id))
    if (account) {
      return `${employee.name}（${account}）`
    }
    if (employee.employeeCode) {
      return `${employee.name}（${employee.employeeCode}）`
    }
    return employee.name
  }, [accountUsernameByEmployeeId])

  const employeeBindingOptions = useMemo(() => {
    return employees.map((employee) => ({
      id: employee.id,
      label: formatEmployeeDisplayLabel(employee)
    }))
  }, [employees, formatEmployeeDisplayLabel])

  const handleAddStaff = () => {
    onAddStaff({
      name: staffName,
      username: staffUsername,
      email: staffEmail,
      password: staffPassword,
      role: staffRole,
      employeeCode: staffEmployeeCode,
      notes: staffNotes
    })
    setStaffName('')
    setStaffUsername('')
    setStaffEmail('')
    setStaffPassword('')
    setStaffRole('employee')
    setStaffEmployeeCode('')
    setStaffNotes('')
  }

  const handleSaveSubsidiary = () => {
    const normalized = normalizeCellValue(subsidiaryName)
    const manager = normalizeCellValue(subsidiaryManager)
    if (!normalized || normalized === '总公司') {
      return
    }
    const label = manager ? `${normalized}（负责人：${manager}）` : normalized
    const next = Array.from(new Set([...companySubsidiaries, label]))
    setCompanySubsidiaries(next)
    onSaveSubsidiaries(next)
    setSubsidiaryName('')
    setSubsidiaryManager('')
  }

  const handleCreateStore = () => {
    onCreateStore({
      subsidiary: normalizeCellValue(storeSubsidiary) || '总公司',
      shopName: storeName,
      businessLicenseName,
      employeeIds: selectedEmployeeIds
    })
    setStoreName('')
    setBusinessLicenseName('')
    setSelectedEmployeeIds([])
  }

  const toggleEmployeeSelection = (employeeId: string) => {
    setSelectedEmployeeIds((prev) => {
      if (prev.includes(employeeId)) {
        return prev.filter((item) => item !== employeeId)
      }
      return [...prev, employeeId]
    })
  }

  const toggleBindingStoreSelection = (storeId: string) => {
    setBindingStoreIds((prev) => {
      if (prev.includes(storeId)) {
        return prev.filter((item) => item !== storeId)
      }
      return [...prev, storeId]
    })
  }

  const handleBindEmployeeStores = () => {
    if (!bindEmployeeId) {
      return
    }
    onBindEmployeeStores({
      employeeId: bindEmployeeId,
      storeIds: bindingStoreIds
    })
  }

  const handleUpdateCompany = () => {
    onUpdateCompanyProfile({
      companyName,
      legalRepresentative,
      maxUsers: Number(maxUsers) || 0,
      expireDate
    })
  }

  return (
    <section className="upload-card admin-section">
      <h2>公司管理</h2>
      <p>按“分公司管理、店铺管理、员工管理”维护公司组织架构与员工店铺关系。</p>
      {isLoading && <p className="hint-text">加载中...</p>}

      <div className="admin-grid">
        {companyTab === 'subsidiary' && (
          <>

            <section className="admin-card">
              <h3>分公司管理</h3>
              <div className="admin-form-grid">
                <div className="control-row">
                  <label>分公司名称</label>
                  <input className="compact-input" value={subsidiaryName} onChange={(event) => setSubsidiaryName(event.target.value)} placeholder="例如：华南分公司" />
                </div>
                <div className="control-row">
                  <label>分公司管理者</label>
                  <input className="compact-input" value={subsidiaryManager} onChange={(event) => setSubsidiaryManager(event.target.value)} placeholder="例如：王主管" />
                </div>
                <button type="button" className="ghost" onClick={handleSaveSubsidiary} disabled={!canEditCompany}>保存分公司</button>
              </div>
              <div className="admin-list">
                <h4>分公司列表</h4>
                <div className="meta-row admin-list-item"><span className="meta-key">总公司</span><span className="meta-value">默认组织</span></div>
                {companySubsidiaries.length === 0 && <p className="hint-text">{isLoading ? '加载中...' : '暂无分公司记录'}</p>}
                {companySubsidiaries.map((item) => (
                  <div key={item} className="meta-row admin-list-item"><span className="meta-key">{item}</span><span className="meta-value">分公司</span></div>
                ))}
              </div>
            </section>


            <section className="admin-card">
              <h3>公司信息维护</h3>
              {companyProfile ? (
                <><div className="admin-form-grid">
                  <div className="control-row"><label>公司名称</label><input className="compact-input" value={companyName} onChange={(event) => setCompanyName(event.target.value)} disabled={!canEditCompany} /></div>
                  <div className="control-row"><label>法人/负责人</label><input className="compact-input" value={legalRepresentative} onChange={(event) => setLegalRepresentative(event.target.value)} disabled={!canEditCompany} /></div>
                  <div className="control-row"><label>最大账号数</label><span className="compact-span">{maxUsers}</span></div>
                  <div className="control-row"><label>到期日期</label><span className="compact-span">{expireDate}</span></div>
                </div>
                  <button type="button" className="ghost" onClick={handleUpdateCompany} disabled={!canEditCompany}>保存公司信息</button></>


              ) : (
                <p className="hint-text">当前账号未绑定公司信息。</p>
              )}
              {!canEditCompanyPolicy && <p className="hint-text">最大账号数和到期时间仅超级管理员可修改。</p>}
              {!canEditCompany && <p className="hint-text">当前角色不可修改公司信息。</p>}
            </section>

          </>
        )}

        {companyTab === 'store' && (
          <>
            <section className="admin-card admin-card-full">
              <h3>店铺管理（录入并分配员工）</h3>
              <div className="admin-form-grid admin-form-grid-3">
                <div className="control-row">
                  <label>所属组织</label>
                  <select className="compact-select" value={storeSubsidiary} onChange={(event) => setStoreSubsidiary(event.target.value)}>
                    <option value="总公司">总公司</option>
                    {companySubsidiaries.map((item) => (<option key={item} value={item}>{item}</option>))}
                  </select>
                </div>
                <div className="control-row"><label>店铺名</label><input className="compact-input" value={storeName} onChange={(event) => setStoreName(event.target.value)} placeholder="例如：DXM官方店" /></div>
                <div className="control-row"><label>营业执照名</label><input className="compact-input" value={businessLicenseName} onChange={(event) => setBusinessLicenseName(event.target.value)} placeholder="例如：深圳某某科技有限公司" /></div>
              </div>
              <div className="control-row">
                <label>分配员工（可多选）</label>
                <div className="checkbox-grid">
                  {employeeBindingOptions.length === 0 && <span className="hint-text">{isLoading ? '加载中...' : '暂无员工可分配'}</span>}
                  {employeeBindingOptions.map((employee) => (
                    <label key={employee.id} className="checkbox-item">
                      <input type="checkbox" checked={selectedEmployeeIds.includes(employee.id)} onChange={() => toggleEmployeeSelection(employee.id)} />
                      <span>{employee.label}</span>
                    </label>
                  ))}
                </div>
              </div>
              <button type="button" className="ghost" onClick={handleCreateStore}>保存店铺并分配员工</button>
            </section>

            <section className="admin-card admin-card-full">
              <h3>店铺列表</h3>
              <div className="admin-list">
                {branchStores.length === 0 && <p className="hint-text">{isLoading ? '加载中...' : '暂无店铺记录'}</p>}
                {branchStores.map((store) => (
                  <div key={store.id} className="meta-row admin-list-item">
                    <span className="meta-key">{store.subsidiary} / {store.shopName}</span>
                    <span className="meta-value">平台ID：{store.storeIdOnPlatform || '-'} · 营业执照：{store.businessLicenseName || '未填写营业执照名'} · {store.assignedEmployeeNames.length > 0 ? `员工：${store.assignedEmployeeNames.join('、')}` : '未分配员工'}</span>
                  </div>
                ))}
              </div>
            </section>
          </>
        )}

        {companyTab === 'employee' && (
          <section>
            <h3>员工管理</h3>

            {employeeTab === 'create' && (
              <section className="admin-card admin-card-full">
                <h3>添加员工</h3>
                <div className="admin-form-grid">
                  <div className="control-row"><label>姓名</label><input className="compact-input" value={staffName} onChange={(event) => setStaffName(event.target.value)} placeholder="例如：王小明" /></div>
                  <div className="control-row"><label>账号</label><input className="compact-input" value={staffUsername} onChange={(event) => setStaffUsername(event.target.value)} placeholder="例如：wangxm01" /></div>
                  <div className="control-row"><label>邮箱（可选）</label><input className="compact-input" value={staffEmail} onChange={(event) => setStaffEmail(event.target.value)} placeholder="不填将自动生成内部邮箱" /></div>
                  <div className="control-row"><label>初始密码</label><input className="compact-input" type="password" value={staffPassword} onChange={(event) => setStaffPassword(event.target.value)} placeholder="至少 6 位" /></div>
                  <div className="control-row"><label>身份</label><select className="compact-select" value={staffRole} onChange={(event) => setStaffRole(event.target.value)}><option value="company_admin">公司管理者</option><option value="finance">财务</option><option value="branch_manager">分公司总经理</option><option value="team_lead">组长</option><option value="employee">普通员工</option></select></div>
                  <div className="control-row"><label>员工编号（可选）</label><input className="compact-input" value={staffEmployeeCode} onChange={(event) => setStaffEmployeeCode(event.target.value)} placeholder="例如：EMP-1024" /></div>
                  <div className="control-row"><label>备注（可选）</label><input className="compact-input" value={staffNotes} onChange={(event) => setStaffNotes(event.target.value)} placeholder="例如：负责俄罗斯站点" /></div>
                  <button type="button" className="ghost" onClick={handleAddStaff}>创建人员</button>
                </div>
              </section>
            )}


            {employeeTab === 'bind' && (
              <section className="admin-card admin-card-full">
                <h3>员工绑定店铺</h3>
                <p className="hint-text">已展示全量员工，可直接绑定店铺。</p>
                <div className="admin-form-grid">
                  <div className="control-row">
                    <label>选择员工</label>
                    <select className="compact-select" value={bindEmployeeId} onChange={(event) => setBindEmployeeId(event.target.value)}>
                      <option value="">请选择员工</option>
                      {employeeBindingOptions.map((employee) => (
                        <option key={employee.id} value={employee.id}>{employee.label}</option>
                      ))}
                    </select>
                  </div>
                  <div className="control-row">
                    <label>绑定店铺（可多选）</label>
                    <div className="checkbox-grid">
                      {branchStores.length === 0 && <span className="hint-text">{isLoading ? '加载中...' : '暂无店铺可绑定'}</span>}
                      {branchStores.map((store) => (
                        <label key={store.id} className="checkbox-item">
                          <input type="checkbox" checked={bindingStoreIds.includes(store.id)} onChange={() => toggleBindingStoreSelection(store.id)} disabled={!bindEmployeeId} />
                          <span>{store.shopName}（{store.subsidiary}）</span>
                        </label>
                      ))}
                    </div>
                  </div>
                  <button type="button" className="ghost" onClick={handleBindEmployeeStores} disabled={!bindEmployeeId}>保存员工店铺绑定</button>
                </div>
              </section>
            )}

            {employeeTab === 'list' && (
              <>
                <section className="admin-card">
                  <h3>员工列表</h3>
                  <div className="admin-list">
                    {managedUsers.length === 0 && <p className="hint-text">{isLoading ? '加载中...' : '暂无账号'}</p>}
                    {managedUsers.map((user) => (<div key={user.id} className="meta-row admin-list-item"><span className="meta-key">{user.username}（{formatRoleLabel(user.role)}）</span><span className="meta-value">{user.email}</span></div>))}
                  </div>
                </section>
              </>
            )}


          </section>
        )}
      </div>
    </section>
  )
}

type ActionPanelProps = {
  isProcessing: boolean
  canProcess: boolean
  lastCalculatedAt: string
  calculationCount: number
  processDisabledReason: string
  onRunCalculation: () => void
}

type ResultPreviewPanelProps = {
  result: ProcessResult
  onExportAggregated: () => void
  onExportBusinessDetail: () => void
  onExportOtherSheets: () => void
}

export function ActionPanel(props: ActionPanelProps) {
  const {
    isProcessing,
    canProcess,
    lastCalculatedAt,
    calculationCount,
    processDisabledReason,
    onRunCalculation
  } = props

  return (
    <section className="action-panel">
      <div className="action-panel-primary">
        <button type="button" onClick={onRunCalculation} disabled={!canProcess || isProcessing}>
          {isProcessing ? '计算中...' : '计算业绩'}
        </button>
      </div>
      <div className="action-status-grid">
        {lastCalculatedAt && (
          <div className="meta-row action-status-item">
            <span className="meta-key">最近重新生成</span>
            <span className="meta-value">第 {calculationCount} 次（{lastCalculatedAt}）</span>
          </div>
        )}
        {!isProcessing && !canProcess && (
          <div className="meta-row action-status-item">
            <span className="meta-key">当前不可执行原因</span>
            <span className="meta-value">{processDisabledReason}</span>
          </div>
        )}
      </div>
    </section>
  )
}

type SummaryMetricCardProps = {
  title: string
  value: number
}

function SummaryMetricCard(props: SummaryMetricCardProps) {
  const { title, value } = props
  return (
    <article>
      <h3>{title}</h3>
      <p>{value}</p>
    </article>
  )
}

type SummarySectionProps = {
  title: string
  metrics: Array<{ title: string; value: number }>
  className?: string
}

function SummarySection(props: SummarySectionProps) {
  const { title, metrics, className } = props
  return (
    <section className={`summary-section ${className || ''}`}>
      <h3 className="overview-title">{title}</h3>
      <div className="stats-grid">
        {metrics.map((metric) => (
          <SummaryMetricCard key={metric.title} title={metric.title} value={metric.value} />
        ))}
      </div>
    </section>
  )
}

export function ResultPreviewPanel(props: ResultPreviewPanelProps) {
  const {
    result,
    onExportAggregated,
    onExportBusinessDetail,
    onExportOtherSheets
  } = props
  const [hoverRowIndex, setHoverRowIndex] = useState<number | null>(null)
  const [hoverColumnKey, setHoverColumnKey] = useState<string>('')
  const tableColumns = Object.keys(result.aggregatedRows[0] || {})

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

  const orderMetrics = [
    { title: '订单总数', value: result.summary.orderCount },
    { title: '放款订单数', value: result.summary.payoutOrderCount },
    { title: '取消订单退款数', value: result.summary.cancelRefundOrderCount },
    { title: '纠纷订单数', value: result.summary.disputeOrderCount }
  ]

  const incomeMetrics = [
    { title: '订单预计可得', value: result.summary.totalIncomeBeforeFreight },
    { title: '净利润（扣运费+采购）', value: result.summary.totalIncomeAfterOnlineFreight },
    { title: '开票金额合计', value: result.summary.totalInvoicedPurchaseAmount }
  ]

  return (
    <section className="result-panel">

      <div className="summary-grid">
        <SummarySection title="订单相关" className="simple" metrics={orderMetrics} />
        <SummarySection title="收支相关"  metrics={incomeMetrics} />
      </div>

      <h3 className="table-title">订单聚合表（页面预览）</h3>
      <div className="result-toolbar">
        <button type="button" className="ghost" onClick={onExportAggregated}>
          导出订单聚合表
        </button>
        <button type="button" className="ghost" onClick={onExportBusinessDetail}>
          导出业务明细表
        </button>
        <button type="button" className="ghost" onClick={onExportOtherSheets}>
          导出其他对账Sheet
        </button>
      </div>
      <div
        className="detail-table-wrap"
        tabIndex={0}
        onKeyDown={handleTableCopyShortcut}
        onMouseLeave={() => {
          setHoverRowIndex(null)
          setHoverColumnKey('')
        }}
      >
        <table className="detail-table">
          <thead>
            <tr>
              {tableColumns.map((col) => (
                <th
                  key={col}
                  className={hoverColumnKey === col ? 'is-col-hover' : ''}
                  onMouseEnter={() => setHoverColumnKey(col)}
                >
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {result.aggregatedRows.map((row, index) => (
              <tr key={`${normalizeOrderNo(row.订单号)}_${index}`}>
                {tableColumns.map((col) => (
                  <td
                    key={col}
                    className={[
                      hoverRowIndex === index ? 'is-row-hover' : '',
                      hoverColumnKey === col ? 'is-col-hover' : ''
                    ].filter(Boolean).join(' ')}
                    onMouseEnter={() => {
                      setHoverRowIndex(index)
                      setHoverColumnKey(col)
                    }}
                  >
                    {normalizeCellValue(row[col])}
                  </td>
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
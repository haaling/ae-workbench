
# 速卖通绩效计算 SaaS —— MVP 版本需求文档

| 版本 | 日期       | 作者 | 变更说明      |
| ---- | ---------- | ---- | ------------- |
| v1.0 | 2026-06-08 | 产品 | 初版 MVP 需求 |

---

## 1. 项目背景与目标

**背景**：
我们内部研发了一套基于速卖通账单的员工绩效计算工具，已部署在 Web 端。通过上传订单和账单文件，自动计算每笔订单的利润、员工提成等指标，显著提升了多店铺团队的绩效管理效率。现计划将此工具商品化，销售给其他多店铺的速卖通卖家。

**MVP 目标**：

- 最快速度实现可售卖的 SaaS 版本，满足 1~3 家种子客户的使用。
- 验证商业模式（线下收款、手动开通），暂不引入在线支付和开放注册。
- 建立数据隔离和基本鉴权体系，保证不同客户之间的数据安全。
- 控制开发成本，数据库初期采用 MongoDB，并设计合理的数据生命周期，避免存储膨胀。

---

## 2. 核心业务流程

1. 我方超级管理员在后台创建客户公司，并添加该公司的财务账号（用户名/密码）。
2. 客户线下付款后，超级管理员手动激活公司账号，设置到期时间和最大用户数。
3. 财务人员登录系统，绑定速卖通店铺，上传平台导出的账单文件。
4. 系统解析文件，生成订单级绩效明细，并在 7 天内允许财务核对定稿。
5. 同时系统自动聚合生成绩效汇总数据，用于历史趋势看板。
6. 原始上传文件保留 2 年，明细 7 天后自动清理，汇总数据保留 2 年。

---

## 3. 用户角色与权限

| 角色                 | 说明                  | 权限范围                                                   |
| -------------------- | --------------------- | ---------------------------------------------------------- |
| **超级管理员** | 我方内部运营人员      | 管理所有公司、用户、查看全平台日志（写死账号，不开放注册） |
| **公司管理员** | 客户公司高级财务/主管 | 看全公司数据，管理本公司的店铺、普通财务账号               |
| **普通财务**   | 客户公司操作人员      | 只能上传账单、查看本公司绩效看板，不可管理用户             |
| **只读财务**   | 仅查看的财务人员      | 只读本公司数据，不可上传或修改                             |

---

## 4. 功能需求（MVP 范围）

### 4.1 超级管理员后台（我方使用）

- 公司管理
  - 新增/编辑/禁用公司，设置公司名称、到期日期、最大可创建用户数。
- 用户管理
  - 为指定公司添加用户，设置角色（公司管理员/普通财务/只读）。
  - 重置用户密码。
- 操作日志查看
  - 按公司、用户、时间、动作类型筛选日志。
- 资源监控（可选）
  - 查看各公司上传文件记录、存储使用情况。

### 4.2 客户侧功能（财务使用）

- 登录/登出（JWT 认证）
- 店铺管理
  - 添加、编辑、删除本公司的店铺（平台、店铺名、店铺ID）。
- 账单上传与绩效计算
  - 选择店铺、上传速卖通账单文件（支持 .csv / .xlsx）。
  - 系统按预设公式计算每笔订单的绩效（订单金额、退款、佣金、其他成本、净利润）。
  - 生成批次号，记录上传人及时间。
- 绩效核对（7天明细）
  - 展示最近 7 天的订单级绩效明细，支持按店铺、员工、批次筛选。
  - 财务可对异议订单标记，并支持删除批次后重新上传（覆盖更新）。
- 绩效看板与报表
  - 基于汇总表展示月度/季度绩效趋势（按店铺、员工维度）。
  - 支持导出月度绩效报表（Excel/PDF，基于汇总表生成）。
- 个人中心
  - 修改密码。

---

## 5. 数据存储设计（MongoDB）

### 5.1 集合（表）清单与最少字段

#### 5.1.1 `companies` 公司表

```json
{
  "_id": ObjectId,
  "company_name": "XX科技有限公司",
  "status": 1,            // 1=正常, 0=禁用, 2=到期
  "expire_date": ISODate("2027-06-08"),
  "max_users": 5,
  "created_at": ISODate,
  "updated_at": ISODate
}
```

#### 5.1.2 `users` 用户表

```json
{
  "_id": ObjectId,
  "company_id": ObjectId, // 数据隔离依据
  "username": "finance01",
  "password_hash": "$2b$...",
  "role": "admin",        // admin / finance / readonly
  "status": 1,            // 1=启用, 0=禁用
  "last_login_at": ISODate,
  "created_at": ISODate,
  "updated_at": ISODate
}
```

- 唯一索引：`username`（全局唯一）

#### 5.1.3 `stores` 店铺表

```json
{
  "_id": ObjectId,
  "company_id": ObjectId,
  "platform": "aliexpress",
  "store_name": "速卖通店A",
  "store_id_on_platform": "123456",
  "created_at": ISODate
}
```

- 复合索引：`{company_id: 1, platform: 1}`

#### 5.1.4 `performance_details` 绩效明细表（7天TTL）

```json
{
  "_id": ObjectId,
  "company_id": ObjectId,
  "store_id": ObjectId,
  "order_number": "8001234567",
  "employee": "张三",
  "order_amount": 29.99,
  "refund_amount": 0,
  "commission": 2.40,
  "other_cost": 1.00,
  "net_profit": 26.59,
  "batch_id": "batch_20260608_001",
  "uploaded_by": ObjectId,
  "created_at": ISODate  // 设置 TTL 索引，7天后自动删除
}
```

- TTL 索引：`created_at` 字段，`expireAfterSeconds: 604800`
- 复合索引：`{company_id:1, created_at:-1}`、`{store_id:1, created_at:-1}`
- 数据保留策略：**仅保留最近 7 天**，供财务核对，到期自动清理

#### 5.1.5 `performance_summary` 绩效汇总表（保留 2 年）

```json
{
  "_id": ObjectId,
  "company_id": ObjectId,
  "store_id": ObjectId,
  "employee": "张三",
  "year": 2026,
  "month": 6,
  "total_orders": 1200,
  "total_sales": 35988.00,
  "total_refunds": 1200.00,
  "total_commission": 2879.04,
  "total_other_cost": 1200.00,
  "total_net_profit": 30708.96,
  "last_updated": ISODate
}
```

- 唯一复合索引：`{company_id:1, store_id:1, employee:1, year:1, month:1}`
- 更新策略：每次明细生成后，使用 `$merge` 或业务代码累加当月数据

#### 5.1.6 `upload_files` 原始文件记录表（保留 2 年）

```json
{
  "_id": ObjectId,
  "company_id": ObjectId,
  "user_id": ObjectId,
  "file_name": "aliexpress_202606.csv",
  "oss_key": "uploads/company_xxx/20260608/batch_xxx.csv",
  "file_size": 1024000,
  "batch_id": "batch_20260608_001",
  "uploaded_at": ISODate
}
```

- 文件实际存储在对象存储（如阿里云 OSS / AWS S3），本表仅存元数据
- 索引：`{company_id:1, uploaded_at:-1}`
- 清理策略：后端定时任务删除 `uploaded_at` 超过 2 年的文件及本表记录

#### 5.1.7 `operation_logs` 操作日志表（可永久保留或定期归档）

```json
{
  "_id": ObjectId,
  "company_id": ObjectId,       // 可为 null
  "user_id": ObjectId,
  "action": "upload_file",      // login_success, login_failed, upload_file, delete_record, create_user, disable_user
  "target_type": "file",
  "target_id": "batch_xxx",
  "description": "上传了 202606 账单，包含 1000 条订单",
  "ip_address": "192.168.1.1",
  "created_at": ISODate
}
```

- 索引：`{company_id:1, created_at:-1}`、`{user_id:1, created_at:-1}`

### 5.2 数据隔离核心原则

- 所有查询与写入操作必须在业务层强制注入 `company_id`，值来源于当前登录用户的 JWT。
- 前端请求绝不包含 `company_id`，防止篡改。
- 登录后所有的 API 调用均以此为依据，从代码层面实现租户隔离。

---

## 6. 数据生命周期管理总结

| 数据类型         | 存储位置                        | 保留时长  | 清理方式                                 |
| ---------------- | ------------------------------- | --------- | ---------------------------------------- |
| 订单绩效明细     | MongoDB `performance_details` | 7 天      | TTL 自动删除                             |
| 绩效月度汇总     | MongoDB `performance_summary` | 2 年      | 定时任务按年月清理过期文档               |
| 原始上传账单文件 | 对象存储（OSS/S3）              | 2 年      | 定时任务删除文件及 `upload_files` 记录 |
| 操作日志         | MongoDB `operation_logs`      | 建议 2 年 | 可选定时归档或删除                       |

---

## 7. 存储规模估算（单客户示例与全量参考）

假设场景：**10,000 家店铺，日均每店 50 单**，每个财务每月上传一次账单。

- **`performance_details` 稳态**：7 天留存 ≈ 350 万文档，磁盘占用约 300~400 MB。
- **`performance_summary` 2 年**：约 120 万文档，磁盘 < 100 MB。
- **对象存储原始文件**：10,000 × 1 次/月 × 24 月 = 24 万个文件，按 1MB/文件计算约 240 GB，成本极低。

结论：即使扩展到数千店铺，数据库压力极小，主要开销是廉价的对象存储。

---

## 8. 非功能性需求

- **安全**
  - 密码哈希存储（bcrypt）。
  - JWT 过期时间设置合理（如 7 天），支持刷新。
  - API 请求频率限制（上传接口每分钟 N 次）。
  - 文件上传校验格式及魔数，限制大小（最大 20MB）。
- **性能**
  - 所有查询必须带时间范围，默认查近 3 个月。
  - 看板数据查汇总表，禁止扫明细。
- **可扩展性**
  - 预留其他平台字段（`stores.platform`）。
  - 绩效公式可配置（初期可写死在代码中，后期可移至 `company_settings` 集合）。

---

## 9. 验收标准（MVP 交付点）

1. 超级管理员能创建公司、添加用户并设置到期日，到期后用户无法上传。
2. 财务能登录、绑定店铺、上传账单文件，并看到 7 天内的订单绩效明细。
3. 7 天后明细自动消失，但月度汇总仍然存在，看板能展示历史趋势。
4. 不同公司的财务登录后只能看到自己公司的店铺和数据。
5. 所有关键操作（登录、上传、删除）均在日志中可查。
6. 上传的原始文件在 OSS 中保存，2 年内可下载核查。

---

## 10. 附录：未来迭代方向（不在 MVP 范围）

- 在线支付与自动续费
- 用户自助注册与试用
- 速卖通 API 自动同步，告别手动上传
- 自定义绩效公式编辑器
- 员工子账号及自助查询
- 移动端适配/微信小程序
- 多语言与多币种支持

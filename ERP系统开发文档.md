# ERPV4 珠宝管理系统 - 完整开发文档

> 本文档基于源代码自动梳理生成，涵盖系统架构、全局变量、窗口模块、函数、SQL查询语句、数据库表等完整开发信息。

---

## 目录

1. [系统概述](#1-系统概述)
2. [系统架构](#2-系统架构)
3. [全局变量](#3-全局变量)
4. [DLL命令定义](#4-dll命令定义)
5. [模块说明](#5-模块说明)
6. [数据库表结构](#6-数据库表结构)
7. [窗口模块详细说明](#7-窗口模块详细说明)
   - 7.1 启动窗口
   - 7.2 主窗口
   - 7.3 首页自动任务
   - 7.4 基础设置模块
   - 7.5 商品管理模块
   - 7.6 销售管理模块
   - 7.7 库存管理模块
   - 7.8 会员管理模块
   - 7.9 财务管理模块
   - 7.10 报表模块
   - 7.11 系统管理模块
8. [权限体系](#8-权限体系)
9. [编码工具函数](#9-编码工具函数)

---

## 1. 系统概述

**系统名称**: ERPV4 珠宝/金银饰品管理系统
**版本号**: V3.0.88
**开发语言**: 易语言 (EPL)
**数据库**: MySQL (业务数据) + Access (本地临时数据)
**UI框架**: ExUI (扩展UI组件库)
**打印框架**: Grid++Report 6 (锐浪报表)
**主要模块依赖**:
- 精易模块.ec — 基础功能扩展
- ExuiFunction20230228.ec — ExUI增强功能
- 谷歌两步验证模块.ec — 谷歌OTP验证
- 导出EXCEL模块.ec — Excel导出
- LibXL4.2.ec — Excel库
- 高级表格查找定位.ec — 表格搜索定位
- 锐浪类封装.ec — Grid++Report封装
- Mask阴影.ec — 窗口阴影效果
- 多线程下载.ec — 多线程下载
- HTTP下载.ec — HTTP文件下载
- 取文件目录列表_模块.ec — 文件操作

**业务范围**:
- 商品管理（入库/销售/调拨/退库/回收/退货/预售）
- 会员管理（会员信息/充值/消费/回访/合并/积分）
- 库存管理（实时库存/历史库存/盘点）
- 财务管理（收支记录/店铺结算/工厂对账）
- 报表统计（销售报表/入库报表/员工绩效/月汇总）
- 系统管理（用户/岗位/权限/日志/配置）

---

## 2. 系统架构

### 2.1 数据库架构

系统采用**读写分离**架构：
- **授权库**: `127.0.0.1:3306/erpshouquan` — 存储授权信息、数据库连接配置
- **业务写库**: 从授权库读取配置后连接 — 处理所有数据写入操作
- **业务读库**: 从授权库读取配置后连接 — 处理所有数据读取操作
- **业务读库备**: 备用读库连接（可选）
- **本地Access**: `data/erpdata.mdb` — 本地临时数据

全局MySQL连接句柄：
- `全_MySQL句柄` — 授权库连接
- `全_MySQL写入` — 业务写库连接
- `全_MySQL读取` — 业务读库连接（通用）
- `全_MySQL读取_报表` — 报表专用读连接
- `全_MySQL读取_任务` — 首页任务专用读连接
- `全_MySQL读取_打印` — 打印专用读连接
- `全_MySQL读取_订单` — 订单查询专用读连接

### 2.2 系统架构图

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                              ERPV4 珠宝管理系统架构                              │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌─────────────────────────────────────────────────────────────────────────┐   │
│  │                          表示层 (ExUI Framework)                         │   │
│  │  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐     │   │
│  │  │ 窗口模块  │ │ 高级表格  │ │ 报表组件  │ │ 打印组件  │ │ 图表组件  │     │   │
│  │  │ 106个窗口 │ │ ExUI表格 │ │Grid++Report│ │ 标签打印 │ │ 曲线/饼图 │     │   │
│  │  └──────────┘ └──────────┘ └──────────┘ └──────────┘ └──────────┘     │   │
│  └─────────────────────────────────────────────────────────────────────────┘   │
│                                                                                 │
│  ┌─────────────────────────────────────────────────────────────────────────┐   │
│  │                          业务逻辑层                                      │   │
│  │  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐     │   │
│  │  │ 商品管理  │ │ 销售管理  │ │ 库存管理  │ │ 会员管理  │ │ 财务管理  │     │   │
│  │  │ 入库/调拨 │ │ 出库/客退 │ │ 实时/历史 │ │ 充值/消费 │ │ 收支/结算 │     │   │
│  │  │ 退库/回收 │ │ 订单查询 │ │ 盘点/预警 │ │ 回访/合并 │ │ 绩效管理 │     │   │
│  │  └──────────┘ └──────────┘ └──────────┘ └──────────┘ └──────────┘     │   │
│  │  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐     │   │
│  │  │ 报表统计  │ │ 系统管理  │ │ 权限管理  │ │ 打印管理  │ │ 数据导入  │     │   │
│  │  │ 销售报表 │ │ 用户管理 │ │ 岗位权限 │ │ 单据打印 │ │ Excel导入│     │   │
│  │  │ 入库报表 │ │ 日志管理 │ │ 店铺权限 │ │ 标签打印 │ │ 批量导入 │     │   │
│  │  └──────────┘ └──────────┘ └──────────┘ └──────────┘ └──────────┘     │   │
│  └─────────────────────────────────────────────────────────────────────────┘   │
│                                                                                 │
│  ┌─────────────────────────────────────────────────────────────────────────┐   │
│  │                          数据访问层                                      │   │
│  │  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐     │   │
│  │  │ MySQL读库│ │ MySQL写库│ │ MySQL备库│ │ Access库 │ │ 文件系统  │     │   │
│  │  │ 业务数据 │ │ 业务数据 │ │ 备用读库 │ │ 临时数据 │ │ 图片/配置 │     │   │
│  │  └──────────┘ └──────────┘ └──────────┘ └──────────┘ └──────────┘     │   │
│  └─────────────────────────────────────────────────────────────────────────┘   │
│                                                                                 │
│  ┌─────────────────────────────────────────────────────────────────────────┐   │
│  │                          基础设施层                                      │   │
│  │  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐     │   │
│  │  │ 精易模块  │ │ ExUI模块 │ │ 谷歌验证  │ │ Excel模块│ │ HTTP模块  │     │   │
│  │  │ 基础功能 │ │ UI增强  │ │ TOTP验证 │ │ 数据导出 │ │ 文件下载 │     │   │
│  │  └──────────┘ └──────────┘ └──────────┘ └──────────┘ └──────────┘     │   │
│  └─────────────────────────────────────────────────────────────────────────┘   │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

### 2.3 启动流程

```
启动窗口 → 连接授权库 → 验证授权码 → 检查版本更新
→ 检查模块更新 → 检查打印插件 → 读取数据库配置
→ 连接业务读/写库 → 读取客户配置信息 → 读取品类配置
→ 用户登录 → 权限加载 → 加载主窗口
```

### 2.4 导航架构

系统采用**两级树形导航**，导航数据存储在数据库中：
- 一级导航: `基础设置`、`账户管理`、`成品管理`、`财务管理`、`会员管理`、`资产管理`、`报表`、`订单管理`
- 二级导航: 各一级导航下的具体功能页面

导航表: `erp_navigation`（一级/二级导航配置）
导航权限表: `erp_navigation_type`（操作权限配置）

### 2.5 窗口交互模式

- **模态窗口**: 使用 `载入(窗口, 父窗口, 真)` 打开
- **非模态窗口**: 使用 `载入(窗口, 父窗口, 假)` 打开
- **窗口间数据传递**: 通过全局变量和程序集变量

---

## 3. 全局变量

### 3.1 数据库连接变量

| 变量名 | 类型 | 说明 |
|--------|------|------|
| `全_MySQL句柄` | 整数型 | 授权库连接句柄 |
| `全_MySQL写入` | 整数型 | 业务写库连接句柄 |
| `全_MySQL读取` | 整数型 | 业务读库连接句柄（通用） |
| `全_MySQL读取_报表` | 整数型 | 报表专用读连接句柄 |
| `全_MySQL读取_任务` | 整数型 | 首页任务专用读连接句柄 |
| `全_MySQL读取_打印` | 整数型 | 打印专用读连接句柄 |
| `全_MySQL读取_订单` | 整数型 | 订单查询专用读连接句柄 |
| `数据库写入集句柄` | 整数型 | 写库记录集句柄 |
| `数据库写入地址` | 文本型 | 写库服务器地址 |
| `数据库写入端口` | 文本型 | 写库端口 |
| `数据库写入表名` | 文本型 | 写库数据库名 |
| `数据库写入密码` | 文本型 | 写库密码 |
| `数据库写入账户` | 文本型 | 写库用户名 |
| `数据库读取集句柄` | 整数型 | 读库记录集句柄 |
| `数据库读取地址` | 文本型 | 读库服务器地址 |
| `数数据库读取端口` | 文本型 | 读库端口 |
| `数据库读取表名` | 文本型 | 读库数据库名 |
| `数据库读取密码` | 文本型 | 读库密码 |
| `数据库读取账户` | 文本型 | 读库用户名 |
| `数据库读取备集句柄` | 整数型 | 备用读库记录集句柄 |
| `数据库读取备地址` | 文本型 | 备用读库服务器地址 |
| `数数据库读取备端口` | 文本型 | 备用读库端口 |
| `数据库读取备表名` | 文本型 | 备用读库数据库名 |
| `数据库读取备密码` | 文本型 | 备用读库密码 |
| `数据库读取备账户` | 文本型 | 备用读库用户名 |
| `全局_查询数据主备` | 文本型 | "0"=主库读, "1"=备库读 |

### 3.2 授权信息变量

| 变量名 | 类型 | 说明 |
|--------|------|------|
| `全_授权数据码` | 文本型 | 授权码数据 |
| `全_授权注册公司` | 文本型 | 授权注册公司名 |
| `全_授权数量` | 文本型 | 授权设备数量 |
| `全_授权联系人` | 文本型 | 授权联系人 |
| `全_授权码联系电话` | 文本型 | 授权联系电话 |
| `全_授权开始日期` | 文本型 | 授权开始日期 |
| `全_授权结束日期` | 文本型 | 授权结束日期 |
| `全_授权信息简写` | 文本型 | 授权信息简写 |

### 3.3 系统配置变量

| 变量名 | 类型 | 说明 |
|--------|------|------|
| `全_版本号` | 文本型 | 当前版本号（V3.0.88） |
| `全_开发公司名称` | 文本型 | 开发公司名称 |
| `全_开发公司简介` | 文本型 | 开发公司简介 |
| `全_开发联系我们` | 文本型 | 联系我们信息 |
| `全_开发联系客服` | 文本型 | 客服联系方式 |
| `全_开发软件版本` | 文本型 | 最新版本号 |
| `全_开发下载地址` | 文本型 | 下载地址 |
| `全_开发识别地址` | 文本型 | API识别地址 |
| `全_开发下载说明` | 文本型 | 下载说明 |

### 3.4 客户配置变量

| 变量名 | 类型 | 说明 |
|--------|------|------|
| `全_客户公司名称` | 文本型 | 客户公司名称 |
| `全_款号识别地址` | 文本型 | 款号图片识别API地址 |
| `全_客户LOGO` | 文本型 | 客户LOGO图片URL |
| `全_客户头像` | 文本型 | 用户头像图片URL |
| `全_优惠比例` | 文本型 | 优惠比例设置 |
| `全_退库仓id` | 文本型 | 退库仓ID |
| `全_退库仓名称` | 文本型 | 退库仓名称 |
| `全_线上收款id` | 文本型 | 线上收款ID |
| `全_窗口图标图片` | 文本型 | 窗口图标URL |
| `全_金类品类` | 文本型 | 金类品类列表（SQL IN条件格式） |
| `全_银类品类` | 文本型 | 银类品类列表（SQL IN条件格式） |
| `全_结算品类` | 文本型 | 结算品类列表（SQL IN条件格式） |
| `全局_标签打印机名称` | 文本型 | 标签打印机名称 |
| `全局_标签打印机链接` | 文本型 | 标签打印机连接地址 |

### 3.5 业务单号变量

| 变量名 | 类型 | 说明 |
|--------|------|------|
| `全局_商品入库单号` | 文本型 | 当前操作的入库单号 |
| `全局_商品销售单号` | 文本型 | 当前操作的销售单号 |
| `全局_商品预售单号` | 文本型 | 当前操作的预售单号 |
| `全局_商品退库单号` | 文本型 | 当前操作的退库单号 |
| `全局_商品调拨单号` | 文本型 | 当前操作的调拨单号 |
| `全局_商品退货单号` | 文本型 | 当前操作的退货单号 |
| `全局_商品回收单号` | 文本型 | 当前操作的回收单号 |
| `全局_店铺结算单号` | 文本型 | 当前操作的店铺结算单号 |
| `全局_证书管理id内容` | 文本型 | 证书管理ID内容 |
| `全局_介绍人必填` | 文本型 | 销售介绍人是否必填 |
| `全局_入库商品编码文本` | 文本型 | 入库商品编码 |

### 3.6 用户/权限变量

| 变量名 | 类型 | 说明 |
|--------|------|------|
| `全局_登录IP` | 文本型 | 当前登录IP地址 |
| `全局_登录账户id` | 文本型 | 当前登录账户ID |
| `全局_谷歌验证开启` | 文本型 | 谷歌验证是否开启（"0"=开启） |
| `全局_谷歌验证密匙` | 文本型 | 谷歌验证密匙 |
| `全局_销售查询发票` | 文本型 | 销售查询是否按发票筛选 |
| `全局_用户账户` | 文本型 | 当前登录用户账号 |
| `全局_账户名称` | 文本型 | 当前登录用户姓名 |
| `全局_账户查看权限` | 文本型 | 查看权限级别（全部/店铺/岗位/个人） |
| `全局_账户类型` | 文本型 | 账户类型 |
| `全局_账户分组id` | 文本型 | 账户所属分组（店铺）ID |
| `全局_账户岗位id` | 文本型 | 账户所属岗位ID |
| `全局_账户分组名称` | 文本型 | 账户所属分组名称 |
| `全局_账户岗位名称` | 文本型 | 账户所属岗位名称 |
| `全局_岗位权限可视` | 文本型 | 岗位可视权限（导航栏目ID列表，SQL IN格式） |
| `全局_岗位权限操作` | 文本型 | 岗位操作权限（操作编号列表，逗号分隔） |
| `全局_店铺权限操作` | 文本型 | 店铺操作权限（店铺ID列表，SQL IN格式） |
| `全局_密码错误次数` | 文本型 | 密码错误次数 |
| `全局_全局查看` | 文本型 | 全局查看SQL子查询（根据权限级别动态生成） |
| `全局_日志保存内容` | 文本型 | 当前操作日志内容 |
| `全局_信息操作账户` | 文本型 | 当前操作账户 |
| `全局_信息操作日期` | 文本型 | 当前操作日期 |
| `全局_首页查询栏目文本` | 文本型 | 当前首页查询的栏目名称 |
| `全局_首页查询栏目id` | 文本型 | 当前首页查询的栏目ID |
| `全局_验证码文本内容` | 文本型 | 验证码内容 |
| `全局_更新模块文本` | 文本型 | 需要更新的模块列表 |

---

## 4. DLL命令定义

系统使用Windows BCrypt加密API（bcrypt.dll）进行数据加密：

| DLL命令 | 函数名 | 说明 |
|---------|--------|------|
| `BCryptOpenAlgorithmProvider` | 打开加密算法提供者 | 初始化加密算法 |
| `BCryptGetProperty` | 获取加密属性 | 获取算法属性值 |
| `BCryptCloseAlgorithmProvider` | 关闭加密算法提供者 | 释放加密资源 |
| `BCryptCreateHash2` | 创建哈希 | 创建哈希对象 |
| `BCryptHashData` | 哈希数据 | 向哈希对象添加数据 |
| `BCryptFinishHash` | 完成哈希 | 完成哈希计算 |
| `BCryptDestroyHash` | 销毁哈希 | 销毁哈希对象 |

---

## 5. 模块说明

| 模块文件 | 说明 |
|----------|------|
| `精易模块.ec` | 精易模块，提供文件操作、网络操作、系统操作等基础功能 |
| `ExuiFunction20230228.ec` | ExUI扩展功能模块 |
| `谷歌两步验证模块.ec` | Google Authenticator TOTP验证 |
| `导出EXCEL模块.ec` | Excel文件导出功能 |
| `LibXL4.2.ec` | LibXL库封装，Excel读写 |
| `高级表格查找定位.ec` | 高级表格的查找和定位功能 |
| `锐浪类封装.ec` | Grid++Report报表控件封装 |
| `Mask阴影.ec` | 窗口阴影效果模块 |
| `多线程下载.ec` | 多线程文件下载 |
| `HTTP下载.ec` | HTTP文件下载 |
| `取文件目录列表_模块.ec` | 文件目录列表获取 |

---

## 6. 数据库表结构

### 6.1 核心商品表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_shop` | 商品主信息表 | id, poduct_code(编码), fu_code(副编码), product_name, caizhi(材质), chengse(成色), specification_id(规格ID), item_number(款号), quantity, net_weight, cost_price, sales_price |
| `xipunum_erp_shop_kucun` | 商品库存表 | id, poduct_code, kufang(库房ID), quantity, jinzhong(金重), creationtime |
| `xipunum_erp_shop_xiangqian` | 商品镶嵌信息表 | id, poduct_code, 主石/副石信息 |
| `xipunum_erp_shop_lincun` | 临存表 | id, poduct_code, kufang |
| `xipunum_erp_shopys` | 预售关联表 | id, poduct_code |
| `xipunum_erp_shop_log` | 商品操作日志表 | id, poduct_code, type, conter, user, creationtime |

### 6.2 品类/规格/款号表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_category` | 品类表 | id, title, caizhiid, jianxie(简写), chengse, xiangqian, shuliang(数量), kejia(克价) |
| `xipunum_erp_specs` | 规格表 | id, title, category_id, shuliang, jianxie |
| `xipunum_erp_ksiamges` | 款号图片表 | id, title, kuanhao(款号), yimage(图片), caizhi, chengse, lingxiao, category_id, specification_id |

### 6.3 入库相关表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_store` | 入库商品明细表 | id, order_id, poduct_code, product_name, quantity, net_weight, gold_price, cost_price, remarks, creationtime |
| `xipunum_erp_store_order` | 入库订单表 | id, store_number(入库单号), gongshi(工厂), remarks, state, cjuser, creationtime |

### 6.4 出库/销售相关表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_outbound` | 出库商品明细表 | id, order_id, poduct_code, quantity, net_weight, gold_price, settlement(结算金额), shopping_guide(导购员), kufang, creationtime |
| `xipunum_erp_outbound_order` | 出库订单表 | id, settlement_number(结算单号), retrea_umber(回收单号), customer_code, ying_amount(应收), settlement(实收), pling(批零), salesman, cjuser, creationtime |

### 6.5 调拨相关表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_transfer` | 调拨明细表 | id, order_id, poduct_code, quantity, net_weight, creationtime |
| `xipunum_erp_transfer_order` | 调拨订单表 | id, transfer_number, kufang_from, kufang_to, state, cjuser, creationtime |

### 6.6 退库相关表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_tuiku` | 退库明细表 | id, order_id, poduct_code, quantity, net_weight, creationtime |
| `xipunum_erp_tuiku_order` | 退库订单表 | id, tuiku_number, factory(工厂), state, cjuser, creationtime |

### 6.7 回收相关表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_retreat` | 回收明细表 | id, order_id, product_name, jin_zhong(金重), chengse(成色), price(克价), qita_price(其他金额), retreat_amount(回收金额), remarks, creationtime |
| `xipunum_erp_retreat_order` | 回收订单表 | id, retrea_umber(回收单号), customer_code, settlement(结算), jin_zhong(总金重), cjuser, creationtime |
| `xipunum_erp_retreat_title` | 回收名称表 | id, title, state |

### 6.8 预售相关表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_presale` | 预售明细表 | id, order_id, product_name, quantity, remarks, creationtime |
| `xipunum_erp_presale_order` | 预售订单表 | id, presale_umber(预售单号), customer_code, state, cjuser, creationtime |

### 6.9 退货相关表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_return` | 退货明细表 | id, order_id, poduct_code, quantity, net_weight, creationtime |
| `xipunum_erp_return_order` | 退货订单表 | id, return_number, customer_code, state, cjuser, creationtime |

### 6.10 会员相关表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_member` | 会员主表 | id, customer_code(客户编码), memberid(会员ID), name, tel, shengri(生日), dizhi(地址), cjuser, creationtime |
| `xipunum_erp_member_cq` | 会员存取记录表 | id, customer_code, cunqu(存/欠), type(料/元), number(数值), remarks, kufang, cjuser, creationtime |
| `xipunum_erp_member_score_log` | 会员积分记录表 | id, customer_code, score, type, conter, creationtime |
| `xipunum_erp_visit` | 会员回访记录表 | id, customer_code, returntitle, returnconter, returndata, cjuser, creationtime |

### 6.11 财务/支付表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_pay` | 支付方式表 | id, name, state |
| `xipunum_erp_shoukuan` | 收款记录表 | id, order_id, pay_name(支付方式), amount(金额), creationtime |
| `xipunum_erp_voucher` | 凭证/单据样式表 | id, title, type, state |

### 6.12 用户/权限表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_user` | 用户/员工表 | id, user, password, name, type, department(分组ID), post(岗位ID), ip, mailbox, google, google_secret, login(错误次数), state, ksdate, jsdate, xsrole, xsdata |
| `xipunum_erp_role` | 岗位权限表 | id, typeid, keshi(可视权限), caozuo(操作权限), shopid(店铺权限) |
| `xipunum_erp_type` | 类型/库房/商铺表 | id, title, type, superior, data1-data10 |

### 6.13 系统配置表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_config` | 系统配置表 | id, title, conter |
| `xipunum_erp_category_stat_config` | 品类统计配置表 | id, title, category_list |
| `erp_config` | 授权系统配置表 | id, title, conter |
| `erp_authorize` | 授权信息表 | id, authorize, datanum, gongsi, num, name, tel, kstime, jstime, jianxie |
| `erp_mysql` | 数据库连接配置表 | id, authorizeid, server, port, dbuser, password, database, state |
| `erp_voucher` | 授权凭证表 | id, banben1, state |
| `erp_navigation` | 导航菜单表 | id, navigation, role, superior, sort, state |
| `erp_navigation_type` | 导航操作权限表 | id, chazhao |
| `erp_city` | 省市区地址表 | id, pid, name, level |

### 6.14 日志/历史表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_xitong_log` | 系统操作日志表 | id, type, title, conter, user, creationtime |
| `xipunum_erp_user_log` | 用户登录日志表 | id, ip, user, conter, creationtime |
| `xipunum_erp_history` | 操作历史表 | id, type, conter, user, creationtime |

### 6.15 线上订单表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_online_order` | 线上订单表 | id, settlement_number, customer_code, state, kufang, creationtime |

### 6.16 打印任务表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_pring` | 打印任务队列表 | id, type, danhao_umber(单号), state, dianpuid(店铺ID) |

### 6.17 工厂/关联信息表

| 表名 | 说明 | 主要字段 |
|------|------|----------|
| `xipunum_erp_about` | 工厂/关联信息表 | id, title, tel, dizhi(地址), type |
| `xipunum_erp_zsjigou` | 检测机构表 | id, name |
| `xipunum_erp_zhengshu` | 证书信息表 | id, product_code, jgname(机构名), zhengshu(证书号), remarks |
| `xipunum_erp_iamges` | 图片上传表 | id, image_url, creationtime |
| `xipunum_erp_online_order` | 线上订单表 | id, settlement_number, customer_code, state, kufang, creationtime |

### 6.18 本地Access表

| 表名 | 说明 |
|------|------|
| `chuku` | 临时出库表 |
| `biaoqian` | 标签打印表 |
| `shujulist` | 主窗口表格临时数据表 |
| `data/pandian.mdb` | 盘点数据库 |

---

## 7. 窗口模块详细说明

### 7.1 启动窗口 (_启动窗口)

**功能**: 系统启动、授权验证、数据库连接、用户登录

#### 程序集变量
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_更新模块文本` | 文本型 | 需要更新的模块文本 |

#### 子程序列表
| 子程序名 | 功能 |
|----------|------|
| `__启动窗口_创建完毕` | 窗口初始化：连接授权库、验证授权码、读取系统配置、检查版本更新、检查模块更新、下载打印模板、连接业务读写库、加载客户配置、读取本地用户配置 |
| `__启动窗口_将被销毁` | 窗口销毁时清理阴影效果 |
| `_透明标签1_鼠标左键被按下` | 关闭按钮：断开数据库连接并退出 |
| `_帐号_编辑框_按下某键` | 账号输入框回车键处理 |
| `_帐号_编辑框_获得焦点` | 账号获得焦点时清空占位文本 |
| `_帐号_编辑框_失去焦点` | 账号失去焦点时恢复占位文本 |
| `_密码_编辑框_按下某键` | 密码输入框回车键处理 |
| `_密码_编辑框_获得焦点` | 密码获得焦点时清空占位文本并切换密码模式 |
| `_密码_编辑框_失去焦点` | 密码失去焦点时恢复占位文本 |
| `_登录按钮_被单击` | **核心登录逻辑**：验证授权有效期→检查账号密码→处理密码错误次数（3次锁定）→验证账户状态→加载权限信息→处理谷歌验证→记录登录日志→加载主窗口 |
| `读取数据库` | 检查模块版本是否需要更新 |

#### 关键SQL
```sql
-- 验证授权码
SELECT * FROM erp_authorize where authorize='{授权码}' LIMIT 1

-- 读取系统配置
SELECT (SELECT conter FROM erp_config WHERE title='公司名称' LIMIT 1) AS mingcheng,
       (SELECT conter FROM erp_config WHERE title='公司简介' LIMIT 1) AS jianjie,...

-- 读取数据库连接配置
SELECT * FROM erp_mysql where authorizeid='{授权id}' and state='1' LIMIT 1  -- 写库
SELECT * FROM erp_mysql where authorizeid='{授权id}' and state='0' LIMIT 1  -- 读库
SELECT * FROM erp_mysql where authorizeid='{授权id}' and state='2' LIMIT 1  -- 备库

-- 读取客户配置
SELECT (SELECT conter FROM xipunum_erp_config WHERE title='公司名称' LIMIT 1) AS mingcheng,
       (SELECT conter FROM xipunum_erp_config WHERE title='软件logo' LIMIT 1) AS logo,...

-- 验证用户登录
SELECT * FROM xipunum_erp_user where user='{用户名}' and password='{密码}' LIMIT 1

-- 加载用户权限详情
SELECT a.id,a.state,a.NAME,a.type,a.department,a.post,a.ip,a.mailbox,a.google,
       a.google_secret,a.DATA,a.ksdate,a.jsdate,b.title AS fenzu,c.title AS gangwei,
       d.keshi,d.caozuo,d.shopid
FROM xipunum_erp_user AS a
INNER JOIN xipunum_erp_type AS b ON b.id = a.department
INNER JOIN xipunum_erp_type AS c ON c.id = a.post
INNER JOIN xipunum_erp_role AS d ON d.typeid = a.post
WHERE a.user='{用户名}' and a.password='{密码}' LIMIT 1

-- 超级管理员加载全部导航权限
SELECT role FROM erp_navigation WHERE state='0' ORDER BY id ASC
SELECT chazhao FROM erp_navigation_type WHERE 1=1 ORDER BY id ASC

-- 加载店铺权限
SELECT id AS akufang FROM xipunum_erp_type WHERE type='商铺' AND superior='0'
UNION ALL SELECT '0' AS akufang FROM dual ORDER BY akufang='0' DESC, akufang

-- 读取品类配置
SELECT (SELECT category_list FROM xipunum_erp_category_stat_config WHERE title='金类' LIMIT 1) AS jinlei,
       (SELECT category_list FROM xipunum_erp_category_stat_config WHERE title='银类' LIMIT 1) AS yinlei,...

-- 检查模块更新
SELECT * FROM erp_voucher WHERE state='0' ORDER BY id ASC
```

---

### 7.2 主窗口 (窗口_主窗口)

**功能**: 系统主界面，包含导航、首页仪表盘、数据表格、报表等

#### 程序集变量（70+个）
- 报表对象: `集_数据报表`、`集_子报表1`、`集_数据显示器`
- 表格按钮: `详情按钮`、`编辑按钮`、`删除按钮`、`审核按钮`、`审核撤回按钮`、`打印按钮`、`打印预览`、`调拨按钮`、`日志按钮`、`密码初始化`、`授权按钮`、`单据作废`、`会员充值`、`会员消费`、`会员预购`、`工厂修改`、`提交按钮`、`再次提交`、`备注修改`、`入库日志`、`外部单据`、`批零转换`、`成品销售其他`、`库房修改`、`单据打印`、`驳回按钮`
- 表格开关: `开关按钮`、`谷歌按钮`
- 分页变量: `总数量`、`每页数量`、`总页码`、`当前页码`、`开始数据`
- 查询变量: `局部_查询内容`、`局部_查询开始日期`、`局部_查询结束日期`、`局部_品类名称`、`局部_规格名称`、`局部_库房名称`、`局部_会员排序`
- 线程句柄: `主窗口线程句柄`、`自动打印线程句柄`、`在线订单句柄`、`线程句柄1`

#### 核心子程序列表
| 子程序名 | 功能 |
|----------|------|
| `_窗口_主窗口_创建完毕` | 初始化：连接本地Access、设置窗口标题和头像、初始化组件、创建导航树、初始化表格按钮 |
| `_窗口_主窗口_尺寸被改变` | 窗口自适应布局 |
| `_窗口_主窗口_可否被关闭` | 关闭前清理：结束线程、断开MySQL、删除临时文件 |
| `_组件数据初始化` | 初始化首页仪表盘：30天销售曲线图、当日统计数据 |
| `树形导航栏_创建` | 从数据库加载两级导航菜单到树形列表 |
| `_树形列表框Ex1_项目左键单击` | **核心导航**：根据选中的菜单项切换对应功能页面（首页/岗位分组/系统设置/报表/数据表格等） |
| `_子文件夹表格_被点击` | 加载对应栏目的表格数据（日志/账户/商品/入库/销售/调拨/退库/回收/预售/退货/会员/回访/结账/店铺结算/收支等） |
| `_高级表格1_光标位置改变` | 表格行点击事件，根据栏目类型打开对应的详情/编辑窗口 |
| `_子程序_查询条件` | 执行查询：根据当前栏目和查询条件加载表格数据（支持分页） |
| `_子程序_读取线上数据` | 读取各栏目的总记录数 |
| `高级表格1_加载表头` | 根据当前栏目动态加载表格列头 |
| `工具条_通用_加载` | 根据当前栏目动态加载工具栏按钮（添加/查询/导出等） |
| `_自动打印_按钮_被单击` | 自动打印销售单据（含报表生成） |
| `_本地打印任务_被执行` | 从打印任务队列加载待打印任务 |
| `_主窗口_实时订单打印句柄` | 实时订单自动打印处理 |
| `_主窗口_在新订单线程句柄` | 线上新订单检测（60秒轮询） |
| `主窗口_主窗口线程句柄` | 首页数据加载线程（当日统计/品类饼图/月销售排行） |
| `_主窗口_岗位分组_初始化` | 岗位分组管理页面初始化 |
| `_主窗口_系统设置_初始化` | 系统设置页面初始化 |
| `_主窗口_报表信息_初始化` | 报表信息页面初始化 |

#### 分页机制
- 每页数量选项: 25/50/100/500/1000/5000
- 分页控件: 首页/上一页/下一页/尾页
- 分页变量: `当前页码`、`总页码`、`每页数量`、`总数量`、`开始数据`

#### 首页仪表盘SQL
```sql
-- 30天销售/回收/实收趋势数据
SELECT CAST(ROUND(COALESCE(chuku.settlement,0),2) AS DECIMAL(30,2)) AS xiaoshou,
       CAST(ROUND(COALESCE(huishou.retreat_amount,0),2) AS DECIMAL(30,2)) AS huishou,
       CAST(ROUND(COALESCE(chuku.settlement-huishou.retreat_amount,0),2) AS DECIMAL(30,2)) AS shishou
FROM (SELECT SUM(a.settlement) as settlement, SUM(a.quantity) as quantity, SUM(a.net_weight) as net_weight
      FROM xipunum_erp_outbound as a
      INNER JOIN xipunum_erp_outbound_order AS x ON x.id = a.order_id
      WHERE a.creationtime>='{开始日期}' and a.creationtime<'{结束日期}' {发票条件} and a.shopping_guide in {权限范围}) as chuku,
     (SELECT SUM(retreat_amount) as retreat_amount, SUM(jin_zhong) as jin_zhong
      FROM xipunum_erp_retreat
      WHERE creationtime>='{开始日期}' and creationtime<'{结束日期}' and shopping_guide in {权限范围}) as huishou

-- 导航菜单加载
SELECT * FROM erp_navigation WHERE role in ({权限可视}) and superior='0' and state='0' order by sort ASC
SELECT * FROM erp_navigation WHERE role in ({权限可视}) and superior='{父级ID}' and state='0' order by sort ASC

-- 线上订单检测
SELECT settlement_number FROM xipunum_erp_online_order where state='待取货' and kufang='{库房id}' order by creationtime DESC

-- 打印任务加载
SELECT * FROM xipunum_erp_pring WHERE dianpuid='{店铺id}' order by id desc
```

---

### 7.3 首页自动任务 (窗口程序集_首页自动任务)

**功能**: 后台线程，负责首页仪表盘数据的定时加载

#### 子程序列表
| 子程序名 | 功能 |
|----------|------|
| `主窗口_主窗口线程句柄` | **核心线程**：加载当日销售/回收统计、品类销售饼图（按金重）、月销售排行榜TOP10 |
| `_首页任务_窗口仍有效` | 检查主窗口是否仍然存在 |
| `_首页任务_文本到整数` | 安全文本转整数 |
| `_首页任务_调整曲线纵轴` | 动态调整曲线图纵轴最大值 |

#### 关键SQL
```sql
-- 当日销售回收统计
SELECT chuku.settlement, chuku.quantity, chuku.net_weight,
       huishou.retreat_amount, huishou.jin_zhong,
       chuku.settlement - huishou.retreat_amount AS shishou,
       chuku.net_weight - huishou.jin_zhong AS shijijin
FROM (出库统计子查询) as chuku, (回收统计子查询) as huishou

-- 品类销售饼图（按金重统计）
SELECT SUM(a.net_weight) AS net_weight,
       CASE WHEN b.item_number != '' THEN d.title
            WHEN b.specification_id !='' THEN f.title
            ELSE '未匹配' END AS pinlei
FROM xipunum_erp_outbound as a
JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code
LEFT JOIN xipunum_erp_ksiamges AS c ON c.kuanhao = b.item_number
LEFT JOIN xipunum_erp_category AS d ON d.id = c.category_id
LEFT JOIN xipunum_erp_specs AS e ON e.id = b.specification_id
LEFT JOIN xipunum_erp_category AS f ON f.id = e.category_id
WHERE a.creationtime >= '{日期}' AND a.creationtime < '{日期}'
GROUP BY pinlei HAVING net_weight > 0

-- 月销售排行榜TOP10
SELECT item_number, pinlei, guige, SUM(quantity) AS quantity, SUM(settlement) AS settlement
FROM (出库JOIN商品JOIN品类JOIN规格)
WHERE creationtime >= '{30天前}' AND creationtime < '{明天}'
GROUP BY 规格去重 ORDER BY quantity DESC LIMIT 10
```

---

### 7.4 基础设置模块

#### 7.4.1 品类属性管理 (窗口_品类属性管理)
**功能**: 管理商品品类的材质/成色/镶嵌/数量/克价等属性

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `品类名称_超级列表框组件句柄` | 整数型 | 品类名称列表框组件句柄 |

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_品类属性管理_创建完毕` | 初始化：加载品类属性列表、加载品类名称复选列表 |
| `_窗口_品类属性管理_尺寸被改变` | 窗口自适应布局 |
| `_品类属性管理_表头回调` | 品类名称全选/取消全选 |
| `_超级列表框EX_属性管理_项目左键单击` | 属性列表选中事件，加载选中属性的关联品类 |
| `_选中属性数据_初始化` | 根据选中属性初始化品类复选框状态 |
| `_品类名称_超级列表框EX_项目左键单击` | 品类名称复选框切换 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX_规格管理重置_鼠标左键单击` | 重置表单 |
| `_按钮EX_规格管理保存_鼠标左键单击` | 保存品类属性配置（权限：78品类属性编辑） |

**核心SQL**:
```sql
-- 加载品类属性列表
SELECT * FROM xipunum_erp_category_stat_config WHERE 1=1 ORDER BY id asc

-- 加载品类名称列表（含未匹配选项）
SELECT id,title FROM xipunum_erp_category WHERE 1=1
UNION ALL SELECT '0' AS id, '未匹配' AS title FROM dual
ORDER BY CASE WHEN id = '0' THEN 0 ELSE 1 END,CAST(id AS UNSIGNED) asc

-- 保存品类属性
UPDATE xipunum_erp_category_stat_config SET title='{名称}',category_list='{品类ID列表}' WHERE id='{ID}' LIMIT 1

-- 添加系统日志
INSERT INTO xipunum_erp_xitong_log (type,title,conter,user,creationtime) VALUES (...)
```

**涉及表**: `xipunum_erp_category_stat_config`、`xipunum_erp_category`、`xipunum_erp_xitong_log`

**权限编号**: 78品类属性编辑

---

#### 7.4.2 商品品类管理 (窗口_商品品类管理)
**功能**: 管理商品品类的增删改查

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品品类管理_创建完毕` | 初始化表单，加载品类列表 |
| `_窗口_商品品类管理_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX_品类管理查询_鼠标左键单击` | 查询品类信息 |
| `_超级列表框EX_品类管理_项目左键单击` | 选中品类项，加载到编辑表单 |
| `_按钮EX_品类管理重置_鼠标左键单击` | 重置表单 |
| `_主窗口_图标列表框EX_品类管理` | 加载品类列表数据 |
| `_单选框_镶嵌_选中状态改变` | 镶嵌选项切换 |
| `_单选框_非镶嵌_选中状态改变` | 非镶嵌选项切换 |
| `_单选框_是_选中状态改变` | 多数量选项切换 |
| `_单选框_否_选中状态改变` | 非多数量选项切换 |
| `_单选框_必填_选中状态改变` | 克价必填选项切换 |
| `_单选框_非必填_选中状态改变` | 克价非必填选项切换 |
| `_按钮EX_品类管理保存_鼠标左键单击` | 保存/修改品类信息（权限：34添加品类/34编辑品类） |

**核心SQL**:
```sql
-- 查询品类列表
SELECT * FROM xipunum_erp_category WHERE title like '%{关键字}%' order by id desc

-- 检查品类是否存在
SELECT * FROM xipunum_erp_category where title='{品类名称}'

-- 添加品类
增加记录("xipunum_erp_category", title, jianxie, caizhiid, chengse, xiangqian, shuliang, kejia, jifen, cjuser, creationtime)

-- 修改品类
UPDATE xipunum_erp_category SET title='{名称}',caizhiid='{材质}',chengse='{成色}',xiangqian='{镶嵌}',shuliang='{多数}',kejia='{克价}',jifen='{积分}',cjuser='{账户}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_category`、`xipunum_erp_xitong_log`

**权限编号**: 34添加品类、34编辑品类

---

#### 7.4.3 商品规格管理 (窗口_商品规格管理)
**功能**: 管理商品规格信息

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品规格管理_创建完毕` | 初始化：加载品类下拉框、规格列表 |
| `_窗口_商品规格管理_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_主窗口_图标列表框EX_规格管理` | 加载规格列表（关联品类表） |
| `_按钮EX_规格管理查询_鼠标左键单击` | 查询规格信息 |
| `_按钮EX_规格管理重置_鼠标左键单击` | 重置表单 |
| `_单选框_是_选中状态改变` | 多数量选项切换 |
| `_单选框_否_选中状态改变` | 非多数量选项切换 |
| `_超级列表框EX_规格管理_项目左键单击` | 选中规格项，加载到编辑表单 |
| `_按钮EX_规格管理保存_鼠标左键单击` | 保存/修改规格信息（权限：35添加规格/35编辑规格） |

**核心SQL**:
```sql
-- 加载品类下拉框
SELECT * FROM xipunum_erp_category WHERE 1=1 order by id desc

-- 查询规格列表（关联品类）
SELECT a.id AS id,b.title AS pinlei,a.title AS guige,a.jianxie AS jianxie,a.number AS paixu,a.shuliang AS duoshu
FROM xipunum_erp_specs AS a
INNER JOIN xipunum_erp_category AS b ON b.id = a.category_id
WHERE (a.title LIKE '%{关键字}%' or b.title LIKE '%{关键字}%')
ORDER BY a.id DESC

-- 检查规格是否存在
SELECT * FROM xipunum_erp_specs where category_id='{品类ID}' and title='{规格名称}'

-- 添加规格
增加记录("xipunum_erp_specs", category_id, number, title, jianxie, shuliang, cjuser, creationtime)

-- 修改规格
UPDATE xipunum_erp_specs SET number='{排序}',title='{名称}',shuliang='{多数}',cjuser='{账户}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1
```

**涉及表**: `xipunum_erp_specs`、`xipunum_erp_category`、`xipunum_erp_xitong_log`

**权限编号**: 35添加规格、35编辑规格

---

#### 7.4.4 商品结算方式 (窗口_商品结算方式)
**功能**: 管理商品结算方式配置

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品结算方式_创建完毕` | 初始化表单，加载结算方式列表 |
| `_窗口_商品结算方式_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_主窗口_图标列表框EX_结算方式` | 加载结算方式列表 |
| `_按钮EX_结算方式查询_鼠标左键单击` | 查询结算方式 |
| `_按钮EX_结算方式保存_鼠标左键单击` | 保存/修改结算方式（权限：33添加结算方式/33编辑结算方式） |
| `_按钮EX_结算方式重置_鼠标左键单击` | 重置表单 |
| `_超级列表框EX_结算方式_项目左键单击` | 选中项加载到编辑表单 |

**核心SQL**:
```sql
-- 查询结算方式列表
SELECT * FROM xipunum_erp_type WHERE title like '%{关键字}%' and type='结算方式' and superior='0' order by id desc

-- 检查结算方式是否存在
SELECT * FROM xipunum_erp_type where type='结算方式' and title='{名称}' and superior='0'

-- 添加结算方式
增加记录("xipunum_erp_type", type='结算方式', superior='0', title, cjuser, creationtime)

-- 修改结算方式
UPDATE xipunum_erp_type SET title='{名称}',cjuser='{账户}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1
```

**涉及表**: `xipunum_erp_type`、`xipunum_erp_xitong_log`

**权限编号**: 33添加结算方式、33编辑结算方式

---

#### 7.4.5 商品来源管理 (窗口_商品来源管理)
**功能**: 管理商品来源（工厂）信息

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品来源管理_创建完毕` | 初始化表单，加载来源列表 |
| `_窗口_商品来源管理_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_主窗口_图标列表框EX_商品来源` | 加载商品来源列表 |
| `_按钮EX_商品来源查询_鼠标左键单击` | 查询商品来源 |
| `_按钮EX_商品来源保存_鼠标左键单击` | 保存/修改商品来源（权限：32添加商品来源/32编辑商品来源） |
| `_按钮EX_商品来源重置_鼠标左键单击` | 重置表单 |
| `_超级列表框EX_商品来源_项目左键单击` | 选中项加载到编辑表单 |

**核心SQL**:
```sql
-- 查询商品来源列表
SELECT * FROM xipunum_erp_type WHERE title like '%{关键字}%' and type='商品来源' and superior='0' order by id desc

-- 检查来源是否存在
SELECT * FROM xipunum_erp_type where type='商品来源' and title='{名称}' and superior='0'

-- 添加来源
增加记录("xipunum_erp_type", type='商品来源', superior='0', title, cjuser, creationtime)

-- 修改来源
UPDATE xipunum_erp_type SET title='{名称}',cjuser='{账户}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1
```

**涉及表**: `xipunum_erp_type`、`xipunum_erp_xitong_log`

**权限编号**: 32添加商品来源、32编辑商品来源

---

#### 7.4.6 收支名称管理 (窗口_收支名称管理)
**功能**: 管理收支项目的名称

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_收支名称管理_创建完毕` | 初始化表单，加载收支名称列表 |
| `_窗口_收支名称管理_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_主窗口_图标列表框EX_收支名称` | 加载收支名称列表 |
| `_按钮EX_收支名称查询_鼠标左键单击` | 查询收支名称 |
| `_按钮EX_收支名称重置_鼠标左键单击` | 重置表单 |
| `_超级列表框EX_收支名称_项目左键单击` | 选中项加载到编辑表单 |
| `_超级列表框EX_收支名称_项目右键单击` | 右键弹出删除菜单 |
| `_按钮EX_收支名称保存_鼠标左键单击` | 保存/修改收支名称（权限：82收支名称新增/82收支名称编辑） |
| `_删除_被选择` | 删除收支名称（需检查是否被使用） |

**核心SQL**:
```sql
-- 查询收支名称列表
SELECT * FROM xipunum_erp_finance_title WHERE title like '%{关键字}%' order by id desc

-- 检查名称是否存在
SELECT * FROM xipunum_erp_finance_title where title='{名称}'

-- 检查是否被使用
SELECT * FROM xipunum_erp_finance where name='{ID}'

-- 添加收支名称
增加记录("xipunum_erp_finance_title", title, remarks, cjuser, creationtime)

-- 修改收支名称
UPDATE xipunum_erp_finance_title SET title='{名称}',remarks='{备注}',cjuser='{账户}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1
```

**涉及表**: `xipunum_erp_finance_title`、`xipunum_erp_finance`、`xipunum_erp_xitong_log`

**权限编号**: 82收支名称新增、82收支名称编辑

---

#### 7.4.7 收支卡号管理 (窗口_收支卡号管理)
**功能**: 管理收支使用的银行卡号

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_收支卡号管理_创建完毕` | 初始化：加载支付方式/店铺/开户行下拉框、卡号列表 |
| `_窗口_收支卡号管理_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_主窗口_图标列表框EX_收支卡号` | 加载收支卡号列表（关联支付方式/店铺/银行） |
| `_按钮EX_收支卡号查询_鼠标左键单击` | 查询卡号信息 |
| `_按钮EX_收支卡号重置_鼠标左键单击` | 重置表单 |
| `_超级列表框EX_收支卡号_项目左键单击` | 选中项加载到编辑表单（支付方式/店铺/开户行联动） |
| `_按钮EX_收支卡号保存_鼠标左键单击` | 保存/修改卡号信息（权限：83卡号新增/83卡号编辑） |

**核心SQL**:
```sql
-- 加载支付方式下拉框
SELECT id,name FROM xipunum_erp_pay where state='0' order by id ASC

-- 加载店铺下拉框
SELECT id AS akufang,CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle
FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in ({权限})
UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({权限})
ORDER BY akufang = '0' DESC, akufang

-- 加载开户行下拉框
SELECT id,title FROM xipunum_erp_finance_yinhang where state='0' order by id ASC

-- 查询卡号列表
SELECT a.id as aid,b.name AS name,d.title as dianpu,c.title as kaihu,a.zhname as zhname,a.account as account,a.address as address,a.remark as remark
FROM xipunum_erp_finance_account AS a
INNER JOIN xipunum_erp_pay AS b ON b.id = a.type and b.state='0'
INNER JOIN xipunum_erp_finance_yinhang AS c ON c.id = a.kaihuhang
INNER JOIN xipunum_erp_type AS d ON d.id = a.kufang
WHERE a.zhname like '%{关键字}%' and a.kufang IN ({权限}) ORDER BY a.id asc

-- 检查卡号是否存在
SELECT * FROM xipunum_erp_finance_account where type='{支付方式}' and kufang='{店铺}' and account='{账号}'

-- 添加卡号
增加记录("xipunum_erp_finance_account", type, kufang, kaihuhang, zhname, account, address, remark, cjuser, creationtime)

-- 修改卡号
UPDATE xipunum_erp_finance_account SET kaihuhang='{开户行}',zhname='{户名}',address='{地址}',remark='{备注}',cjuser='{账户}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1
```

**涉及表**: `xipunum_erp_finance_account`、`xipunum_erp_pay`、`xipunum_erp_finance_yinhang`、`xipunum_erp_type`、`xipunum_erp_xitong_log`

**权限编号**: 83卡号新增、83卡号编辑

---

#### 7.4.8 回收名称管理 (窗口_回收名称管理)
**功能**: 管理回收旧料的名称分类

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_回收名称管理_创建完毕` | 初始化：加载品类下拉框、回收名称列表 |
| `_窗口_回收名称管理_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_主窗口_图标列表框EX_回收管理` | 加载回收名称列表（关联品类） |
| `_按钮EX_回收名称查询_鼠标左键单击` | 查询回收名称 |
| `_按钮EX_回收名称重置_鼠标左键单击` | 重置表单 |
| `_超级列表框EX_回收名称_项目左键单击` | 选中项加载到编辑表单（品类联动） |
| `_按钮EX_回收名称保存_鼠标左键单击` | 保存/修改回收名称（权限：74添加回收名称/74编辑回收名称） |

**核心SQL**:
```sql
-- 加载品类下拉框
SELECT * FROM xipunum_erp_category WHERE 1=1 order by id desc

-- 查询回收名称列表（关联品类）
SELECT a.id AS id,b.title AS pinlei,a.title AS mingcheng,a.bianma AS bianma
FROM xipunum_erp_retreat_title AS a
INNER JOIN xipunum_erp_category AS b ON b.id = a.category_id
WHERE (a.title LIKE '%{关键字}%' or b.title LIKE '%{关键字}%')
ORDER BY a.id DESC

-- 检查名称是否存在
SELECT * FROM xipunum_erp_retreat_title where category_id='{品类ID}' and title='{名称}'

-- 添加回收名称
增加记录("xipunum_erp_retreat_title", category_id, title, bianma, cjuser, creationtime)

-- 修改回收名称
UPDATE xipunum_erp_retreat_title SET bianma='{编码}',cjuser='{账户}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1
```

**涉及表**: `xipunum_erp_retreat_title`、`xipunum_erp_category`、`xipunum_erp_xitong_log`

**权限编号**: 74添加回收名称、74编辑回收名称

---

#### 7.4.9 打印机设置 (窗口_打印机设置)
**功能**: 配置标签打印机名称和连接方式

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_打印机设置_创建完毕` | 初始化：加载系统打印机列表、读取已保存配置 |
| `_窗口_打印机设置_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX2_鼠标左键单击` | 重置表单 |
| `_按钮EX1_鼠标左键单击` | 保存打印机配置到本地配置文件 |

**配置文件**: `data/print.ini`
```ini
[print]
name={打印机名称}
lianjie={打印程序路径}
```

**涉及变量**: `全局_标签打印机名称`、`全局_标签打印机链接`

---

#### 7.4.10 谷歌验证绑定 (窗口_谷歌验证绑定)
**功能**: 首次绑定谷歌验证器

**程序集变量（2个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `绑定生成密匙` | 文本型 | 生成的TOTP密匙 |
| `二维码生成路径` | 文本型 | 二维码内容（otpauth://totp/格式） |

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_谷歌验证绑定_创建完毕` | 生成密匙、创建二维码 |
| `_窗口_谷歌验证绑定_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX_关闭_鼠标左键单击` | 关闭窗口 |
| `_按钮EX1_绑定_鼠标左键单击` | 保存密匙到数据库 |
| `_绑定名称编辑框_内容被改变` | 绑定名称变更时重新生成二维码 |
| `_数据加载_二维码生成` | 生成二维码图片 |

**核心SQL**:
```sql
-- 更新用户谷歌验证密匙
UPDATE xipunum_erp_user SET google_secret='{密匙}',ip='' WHERE id='{用户ID}' LIMIT 1

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type='添加', title='添加谷歌验证', conter, user, creationtime)
```

**涉及表**: `xipunum_erp_user`、`xipunum_erp_xitong_log`

---

#### 7.4.11 谷歌验证码输入 (窗口_谷歌验证码输入)
**功能**: 登录时输入谷歌验证码

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_正在验证` | 逻辑型 | 防止重复验证标记 |

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_谷歌验证码输入_创建完毕` | 初始化输入框 |
| `_窗口_谷歌验证码输入_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX_关闭_鼠标左键单击` | 关闭窗口（验证中禁止关闭） |
| `_绑定名称编辑框_按下某键` | 回车键触发验证 |
| `_按钮EX1_确认_鼠标左键单击` | 验证验证码（3次错误锁定账户） |
| `_绑定名称编辑框_键盘事件` | 键盘事件处理 |

**验证逻辑**:
1. 检查密码错误次数（>=3次锁定）
2. 检查密匙是否为空
3. 验证码长度检查（6位）
4. 调用`计算验证码()`比对
5. 错误次数+1，3次锁定账户
6. 成功后记录登录日志，加载主窗口

**核心SQL**:
```sql
-- 更新密码错误次数
UPDATE xipunum_erp_user SET login='{次数}',state='1' WHERE id='{用户ID}' LIMIT 1

-- 更新登录信息
UPDATE xipunum_erp_user SET ip='{IP}',logintime='{时间}',login='0' WHERE id='{用户ID}' LIMIT 1

-- 记录登录日志
增加记录("xipunum_erp_user_log", ip, user, conter, creationtime)
```

**涉及表**: `xipunum_erp_user`、`xipunum_erp_user_log`

---

#### 7.4.12 谷歌验证重置 (窗口_谷歌验证重置)
**功能**: 重置谷歌验证器

**程序集变量（2个）**: 同谷歌验证绑定

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_谷歌验证重置_创建完毕` | 生成新密匙、创建二维码 |
| `_窗口_谷歌验证重置_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX_关闭_鼠标左键单击` | 关闭窗口 |
| `_按钮EX1_绑定_鼠标左键单击` | 保存新密匙、注销登录 |
| `_绑定名称编辑框_内容被改变` | 绑定名称变更时重新生成二维码 |
| `_数据加载_二维码生成` | 生成二维码图片 |

**核心SQL**:
```sql
-- 更新用户谷歌验证密匙
UPDATE xipunum_erp_user SET google_secret='{密匙}',ip='' WHERE id='{用户ID}' LIMIT 1

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type='编辑', title='重置谷歌验证', conter, user, creationtime)
```

**涉及表**: `xipunum_erp_user`、`xipunum_erp_xitong_log`

---

#### 7.4.13 授权码信息输入 (窗口_授权码信息输入)
**功能**: 授权码录入

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_授权码信息输入_创建完毕` | 初始化输入框 |
| `_窗口_授权码信息输入_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX_关闭_鼠标左键单击` | 关闭窗口 |
| `_授权信息编码编辑框_键盘事件` | 回车键触发确认 |
| `_按钮EX1_确认_鼠标左键单击` | 验证授权码并保存到配置文件 |

**核心SQL**:
```sql
-- 验证授权码
SELECT * FROM erp_authorize where authorize='{授权码}' LIMIT 1
```

**配置文件**: `data/config.ini`
```ini
[config]
authorize={授权码}
```

**涉及表**: `erp_authorize`

---

#### 7.4.14 在线更新 (窗口_在线更新)
**功能**: 软件在线更新

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `程更新地址` | 文本型 | 更新包下载地址 |

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_在线更新_创建完毕` | 显示更新说明和最新版本 |
| `_标签EX1_鼠标左键按下` | 窗口拖动 |
| `_图片框_退出_鼠标左键单击` | 关闭窗口 |
| `_按钮_取消_鼠标左键单击` | 取消更新 |
| `_按钮_立即更新_鼠标左键单击` | 下载更新包、解压、执行安装 |
| `_时钟1_周期事件` | 更新下载进度显示 |

---

#### 7.4.15 模块更新 (窗口_模块更新)
**功能**: 模块更新

**程序集变量（2个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_下载行数` | 整数型 | 当前下载的模块行号 |
| `列表框EX1_组件句柄` | 整数型 | 列表框组件句柄 |

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_模块更新_创建完毕` | 加载待更新模块列表 |
| `_标签_下载速度_鼠标左键单击` | 开始下载所有模块 |
| `_时钟1_周期事件` | 更新下载进度和速度显示 |

**核心SQL**:
```sql
-- 加载待更新模块
SELECT * FROM erp_voucher WHERE state='0' and id in ({模块ID列表}) ORDER BY id ASC
```

**涉及表**: `erp_voucher`

---

#### 7.4.16 提示消息 (窗口_提示消息)
**功能**: 全局提示消息窗口（非模态，自动消失）

**程序集变量（10个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_超级按钮EX_array` | 超级按钮Ex数组 | 消息按钮数组 |
| `集_超级按钮EX_顶边间隔` | 整数型 | 消息间距 |
| `集_顶边位置` | 整数型 | 当前消息位置 |
| `集_缓动对象` | 整数型 | 缓动动画对象 |
| `集_是否销毁` | 逻辑型 | 销毁标记 |
| `集_列表框EX_是否鼠标进入透明` | 逻辑型 | 鼠标进入透明 |
| `集_列表框EX_是否鼠标左键单击销毁` | 逻辑型 | 点击销毁 |
| `集_字体Ex数据` | 字节集 | 字体数据 |
| `集_列表框EX_纵向滚动条宽度` | 整数型 | 滚动条宽度 |
| `集_结构体句柄` | 整数型 | 结构体句柄 |

**核心函数**:
| 函数名 | 功能 |
|--------|------|
| `提示框Ex_初始化` | 初始化提示框窗口 |
| `提示框Ex_添加消息` | 添加提示消息（支持9种图标类型） |
| `提示框Ex_销毁` | 销毁提示框 |

**图标类型**:
- 0: 普通
- 1: 成功（蓝色）
- 2: 错误（红色）
- 3: 成功（青色）
- 4: 成功（绿色）
- 5: 等待（蓝色）
- 6: 询问（蓝色）
- 7: 警告（黄色）
- 8: 错误（红色）

---

#### 7.4.17 日期框 (窗口_日期框)
**功能**: 日期选择器弹出窗口

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_返回` | 文本型 | 选中的日期文本 |

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_日期框_创建完毕` | 定位窗口到鼠标位置 |
| `_日历框EX1_项目左键单击` | 选择日期，返回"yyyy-MM-dd"格式 |
| `_主窗口_弹出日期框` | 公开接口，弹出日期选择器并返回选中日期 |

---

#### 7.4.18 信息商品查询 (窗口_信息商品查询)
**功能**: 通用查询条件窗口

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_信息商品查询_创建完毕` | 初始化：加载店铺/品类下拉框，根据栏目类型控制控件状态 |
| `_组合框_品类信息_内容被改变` | 品类变更时联动加载规格下拉框 |
| `_窗口_信息商品查询_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX_关闭_鼠标左键单击` | 关闭窗口 |
| `_编辑框_起始时间_鼠标左键单击` | 弹出日期选择器 |
| `_编辑框_结束时间_鼠标左键单击` | 弹出日期选择器 |
| `_按钮EX1_鼠标左键单击` | 确认查询条件 |

**栏目类型处理**:
| 栏目 | 特殊处理 |
|------|----------|
| 账户列表 | 禁用日期选择 |
| 会员列表 | 禁用店铺/日期选择 |
| 商品列表 | 启用品类/规格选择 |
| 商品入库 | 启用品类选择 |
| 线上订单/商品预售 | 禁用店铺选择 |
| 店铺结算 | 禁用品类/规格选择，起始日期默认月初 |
| 临时销售 | 禁用日期选择，启用品类/规格选择 |

**核心SQL**:
```sql
-- 加载店铺下拉框
SELECT id AS akufang,CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle
FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in ({权限})
UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({权限})
ORDER BY akufang = '0' DESC, akufang

-- 加载品类下拉框
SELECT id,title FROM xipunum_erp_category WHERE 1=1 ORDER BY id asc

-- 加载规格下拉框（按品类筛选）
SELECT id,title FROM xipunum_erp_specs WHERE category_id='{品类ID}' ORDER BY id asc
```

---

### 7.5 商品管理模块

#### 7.5.1 商品信息添加 (窗口_商品信息添加)
**功能**: 新商品信息录入，支持款号图片识别自动填充

**程序集变量（9个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_款号文本内容` | 文本型 | 图片识别返回的款号列表 |
| `查找_款号文本内容` | 文本型 | 款号搜索关键字 |
| `局部_品类名称id` | 文本型 | 品类ID |
| `局部_品类简写` | 文本型 | 品类简写编码 |
| `局部_商品规格id` | 文本型 | 规格ID |
| `局部_品类多数量` | 文本型 | 品类是否多数量（"0"=多数量，"1"=单数量） |
| `局部_商品多数量` | 文本型 | 规格是否多数量 |
| `局部_商品零销售` | 文本型 | 是否零销售（"是"/"否"） |
| `局部_品类原料价` | 文本型 | 品类基准原料价 |

**子程序列表（52个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品信息添加_创建完毕` | 窗口初始化：清空表单、初始化组合框、注册热键 |
| `_按钮EX_查找_鼠标左键单击` | 款号搜索按钮事件 |
| `_按钮EX_重置_鼠标左键单击` | 重置表单 |
| `_按钮EX_图片上传_鼠标左键单击` | 图片上传：选择图片→压缩→上传服务器/识别款号 |
| `_图标列表框EX1_加载数据` | 加载款号列表（图片识别结果或手动搜索） |
| `_商品信息_默认参数` | 设置所有输入框默认状态（禁止/初始值） |
| `_商品组合框_初始化` | 初始化品类/材质/成色/规格/检测机构下拉框 |
| `_组合框_品类名称_内容被改变` | 品类选择联动：加载材质/成色/规格选项 |
| `_组合框_商品材质_鼠标左键单击` | 材质下拉框点击事件（需先选择品类） |
| `_组合框_商品成色_鼠标左键单击` | 成色下拉框点击事件（需先选择品类） |
| `_组合框_商品规格_鼠标左键单击` | 规格下拉框点击事件（需先选择材质和成色） |
| `_组合框_商品规格_内容被改变` | 规格选择联动：设置数量/圈口/固开口状态 |
| `_图标列表框EX1_项目左键单击` | 款号列表选中事件：填充商品名称/款号/材质/成色/规格 |
| `_编辑框_重量_焦点事件` | 重量输入框焦点事件：获得焦点清空/失去焦点计算 |
| `_编辑框_重量_键盘事件` | 重量输入框回车键处理 |
| `_商品信息_参数数值计算` | 重量/金重/损耗/含耗重/成本价/销售价联动计算 |
| `_按钮EX_添加_鼠标左键单击` | 保存商品信息到数据库 |
| `_响应监视事件` | 热键响应事件 |

**核心SQL**:
```sql
-- 加载品类下拉框
SELECT id,title FROM xipunum_erp_category WHERE 1=1 ORDER BY id ASC

-- 加载检测机构下拉框
SELECT id,name FROM xipunum_erp_zsjigou WHERE 1=1 ORDER BY id ASC

-- 加载品类属性（材质/成色/镶嵌/多数量/原料价）
SELECT caizhiid,jianxie,chengse,xiangqian,shuliang,kejia FROM xipunum_erp_category WHERE id='{品类ID}' LIMIT 1

-- 加载规格下拉框
SELECT id,title FROM xipunum_erp_specs WHERE category_id='{品类ID}' ORDER BY id ASC

-- 加载规格属性（多数量/简写）
SELECT id,shuliang,jianxie FROM xipunum_erp_specs WHERE id='{规格ID}' LIMIT 1

-- 款号列表查询（图片识别结果）
SELECT id,title,kuanhao,yimage FROM xipunum_erp_ksiamges WHERE kuanhao in ({款号列表}) ORDER BY FIELD(kuanhao,{款号列表}) LIMIT 9

-- 款号列表查询（手动搜索）
SELECT id,title,kuanhao,yimage FROM xipunum_erp_ksiamges WHERE (title like '%{关键字}%' or kuanhao like '%{关键字}%') and category_id='{品类ID}' ORDER BY id DESC LIMIT 9

-- 款号详情查询
SELECT a.title AS title,a.kuanhao AS kuanhao,a.yimage AS yimage,a.caizhi AS caizhi,a.chengse AS chengse,a.lingxiao AS lingxiao,b.title AS guige
FROM xipunum_erp_ksiamges AS a LEFT JOIN xipunum_erp_specs AS b ON b.id = a.specification_id WHERE a.id='{款号ID}' LIMIT 1

-- 检查商品编码是否存在
SELECT * FROM xipunum_erp_shop where poduct_code='{编码}'

-- 保存商品信息
增加记录("xipunum_erp_shop", poduct_code, fu_code, product_name, caizhi, chengse, specification_id, item_number, quantity, net_weight, single, cost_price, sales_price, ...)

-- 保存证书信息
增加记录("xipunum_erp_zhengshu", poduct_code, jgname, zhengshu, remarks, ...)

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_ksiamges`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_shop`、`xipunum_erp_iamges`、`xipunum_erp_zhengshu`、`xipunum_erp_zsjigou`、`xipunum_erp_xitong_log`

#### 7.5.2 商品销售出库 (窗口_商品销售出库)
**功能**: 销售出库操作，支持零售/批发、旧料回收、多种支付方式

**程序集变量（42个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_行号1` | 整数型 | 销售表格当前行号 |
| `集_列号1` | 整数型 | 销售表格当前列号 |
| `集_行号2` | 整数型 | 回收表格当前行号 |
| `集_列号2` | 整数型 | 回收表格当前列号 |
| `集_行号3` | 整数型 | 支付表格当前行号 |
| `集_列号3` | 整数型 | 支付表格当前列号 |
| `集_数据报表` | 报表 | 报表对象 |
| `集_子报表1` | 报表 | 子报表对象 |
| `集_数据显示器` | 查询显示器 | 数据显示器对象 |
| `删除按钮` | 表格按钮 | 表格删除按钮 |
| `局部重量合计数值` | 文本型 | 销售商品重量合计 |
| `局部重量单价数值` | 文本型 | 销售商品单价合计 |
| `局部重销售价金额数值` | 文本型 | 销售金额合计 |
| `局部重基本工费合计` | 文本型 | 基本工费合计 |
| `局部重精品工费合计` | 文本型 | 精品工费合计 |
| `局部重销售工费合计` | 文本型 | 销售工费合计 |
| `局部重成本附加费合计` | 文本型 | 成本附加费合计 |
| `局部重销售附加费合计` | 文本型 | 销售附加费合计 |
| `局部回收总重合计数值` | 文本型 | 回收总重合计 |
| `局部回收金重合计数值` | 文本型 | 回收金重合计 |
| `局部回收其他合计数值` | 文本型 | 回收其他费用合计 |
| `局部回收回收合计数值` | 文本型 | 回收金额合计 |
| `局部数量合计数量` | 文本型 | 销售数量合计 |
| `局部实收金额金额数值` | 文本型 | 实收金额合计 |
| `局部_销售商品编码` | 文本型 | 当前销售商品编码 |
| `局部_销售商品数据编码` | 文本型 | 当前销售商品数据编码 |
| `信息销售批零` | 文本型 | 批发/零售标记 |
| `局部_批发库存料数值` | 文本型 | 批发会员库存料数值 |
| `局部_批发库存元数值` | 文本型 | 批发会员库存元数值 |
| `局部_表格编辑状态` | 整数型 | 销售表格编辑状态标记 |
| `局部_收支编辑状态` | 整数型 | 收支表格编辑状态标记 |
| `局部_回收编辑状态` | 整数型 | 回收表格编辑状态标记 |
| `出库临时数量` | 整数型 | 临时出库表数据数量 |
| `出库数据加载` | 整数型 | 是否加载临时数据标记 |
| `集_保存事务进行中` | 整数型 | 保存事务进行中标记 |
| `集_保存出库单号` | 文本型 | 保存时生成的出库单号 |
| `集_保存回收单号` | 文本型 | 保存时生成的回收单号 |
| `集_保存新建会员编码` | 文本型 | 保存时新建的会员编码 |
| `集_保存上一笔导购员` | 文本型 | 上一笔销售的导购员名称 |
| `集_保存上一笔导购账户` | 文本型 | 上一笔销售的导购账户 |

**子程序列表（72个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品销售出库_创建完毕` | 窗口初始化：连接本地Access、生成单据号、初始化表格、加载临时数据 |
| `_窗口_商品销售出库_可否被关闭` | 关闭前检查：未保存数据提示 |
| `_窗口_商品销售出库_尺寸被改变` | 窗口自适应布局 |
| `_单据样式_初始化` | 加载单据样式下拉框（出库/置换类型） |
| `_高级表格1_加载表头` | 加载销售表格表头（27列：序号/编码/名称/款号/规格/材质/圈口/成色/单件重/金重/重量/成本工费/参考工费/成本附加费/库存/销售单价/销售金额/数量/原附加费/销售克价/销售工费/销售附加费/折扣/实收金额/导购员/库存金重/操作） |
| `_高级表格2_加载表头` | 加载回收表格表头（11列：序号/商品名称/数量/总重/金重/成色/回收克价/其他费用/回收金额/备注/导购员） |
| `_高级表格3_加载表头` | 加载支付表格表头（4列：序号/支付方式/金额/id） |
| `_高级表格4_加载表头` | 加载合计表格表头 |
| `_高级表格1_加载表格` | 加载销售表格数据 |
| `_高级表格2_加载表格` | 加载回收表格数据 |
| `_高级表格3_加载表格` | 加载支付方式表格数据（从xipunum_erp_pay表读取） |
| `_高级表格1_临时数据` | 从本地Access临时表加载数据 |
| `_高级表格1_调拨读取` | 从调拨单读取商品数据 |
| `_子程序_数据统计汇总` | 统计销售/回收/支付合计数据，计算应收/实收/税额 |
| `子程序_删除表格1` | 清空销售表格 |
| `子程序_删除表格2` | 清空回收表格 |
| `子程序_删除表格3` | 清空支付表格 |
| `_子程序_添加销售` | 添加销售行 |
| `_按钮_回收加_被单击` | 添加回收行 |
| `_按钮_回收减_被单击` | 删除回收行 |
| `_高级表格1_光标位置改变` | 销售表格光标位置改变事件 |
| `_高级表格2_光标位置改变` | 回收表格光标位置改变事件 |
| `_高级表格3_光标位置改变` | 支付表格光标位置改变事件 |
| `_高级表格1_将被编辑` | 销售表格开始编辑事件 |
| `_高级表格1_结束编辑` | 销售表格结束编辑事件：商品编码验证、库存检查、价格联动计算 |
| `_高级表格2_结束编辑` | 回收表格结束编辑事件：回收金额联动计算 |
| `_业务员组合框_将弹出列表` | 业务员下拉框弹出列表事件 |
| `_执行导购员列表SQL` | 执行导购员列表查询SQL |
| `_超级按钮_保存_被单击` | 保存销售单据：事务处理（生成订单→保存销售明细→保存回收明细→扣减库存→保存收款记录→更新会员存欠→打印单据） |
| `_超级按钮_重置_被单击` | 重置表单 |
| `_单选框_零售_选中状态改变` | 切换零售模式 |
| `_单选框_批发_选中状态改变` | 切换批发模式（显示会员存欠信息） |
| `_组合框会员查找_将弹出列表` | 会员搜索下拉框弹出列表 |
| `_组合框会员查找_列表项被选择` | 会员选中事件：加载会员存欠信息 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |
| `_销售出库_格式二位小数` | 格式化两位小数 |
| `_销售出库_格式三位小数UTF8` | 格式化三位小数（UTF8编码） |

**核心SQL**:
```sql
-- 验证商品编码存在性
SELECT * FROM xipunum_erp_shop where (poduct_code='{编码}' or fu_code='{编码}') order by id ASC

-- 检查商品库存
SELECT a.poduct_code AS apoduct_code FROM xipunum_erp_shop_kucun AS a
INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code
WHERE a.kufang = '{库房id}' AND (a.quantity > 0 or a.jinzhong >0) AND (b.poduct_code = '{编码}' OR b.fu_code = '{编码}') ORDER BY a.id DESC

-- 加载商品详情（品类/规格/材质/成色/工费等）
SELECT ... FROM xipunum_erp_shop_kucun AS a
INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code
LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang
LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = b.item_number
LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id
LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id
LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id)
WHERE a.kufang = '{库房id}' AND b.poduct_code = '{编码}'

-- 加载回收名称下拉框
SELECT * FROM xipunum_erp_retreat_title WHERE 1=1 order by id asc

-- 加载导购员列表
SELECT name FROM xipunum_erp_user where state='0' order by id ASC

-- 加载支付方式
SELECT id,name FROM xipunum_erp_pay where state='0' order by id ASC

-- 加载单据样式
SELECT * FROM xipunum_erp_voucher where (type='出库' or type='置换') and state='0' order by id ASC

-- 保存出库订单
增加记录("xipunum_erp_outbound_order", settlement_number, customer_code, ying_amount, settlement, pling, salesman, cjuser, creationtime)

-- 保存出库明细
增加记录("xipunum_erp_outbound", order_id, poduct_code, quantity, net_weight, gold_price, settlement, shopping_guide, kufang, creationtime)

-- 扣减库存
UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '{数量}',jinzhong = jinzhong - '{金重}' WHERE poduct_code = '{编码}' AND kufang='{库房id}'

-- 保存回收订单
增加记录("xipunum_erp_retreat_order", retrea_umber, customer_code, settlement, jin_zhong, cjuser, creationtime)

-- 保存回收明细
增加记录("xipunum_erp_retreat", order_id, product_name, jin_zhong, chengse, price, qita_price, retreat_amount, remarks, shopping_guide, creationtime)

-- 保存收款记录
增加记录("xipunum_erp_shoukuan", order_id, pay_name, amount, creationtime)

-- 更新会员存欠（批发模式）
增加记录("xipunum_erp_member_cq", customer_code, cunqu, type, number, remarks, kufang, cjuser, creationtime)

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_retreat`、`xipunum_erp_retreat_order`、`xipunum_erp_retreat_title`、`xipunum_erp_member`、`xipunum_erp_member_cq`、`xipunum_erp_shoukuan`、`xipunum_erp_pay`、`xipunum_erp_voucher`、`xipunum_erp_ksiamges`、`xipunum_erp_specs`、`xipunum_erp_category`、`xipunum_erp_user`、`xipunum_erp_type`、`xipunum_erp_xitong_log`

#### 7.5.3 商品销售编辑 (窗口_商品销售编辑)
**功能**: 已保存销售单的编辑修改，含库存回补/重新扣减

**程序集变量（38个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_行号1` | 整数型 | 销售表格当前行号 |
| `集_列号1` | 整数型 | 销售表格当前列号 |
| `集_行号2` | 整数型 | 回收表格当前行号 |
| `集_列号2` | 整数型 | 回收表格当前列号 |
| `集_行号3` | 整数型 | 支付表格当前行号 |
| `集_列号3` | 整数型 | 支付表格当前列号 |
| `集_数据报表` | 报表 | 报表对象 |
| `集_子报表1` | 报表 | 子报表对象 |
| `集_数据显示器` | 查询显示器 | 数据显示器对象 |
| `删除按钮` | 表格按钮 | 表格删除按钮 |
| `局部_销售单据id` | 文本型 | 当前销售单据ID |
| `局部_销售单据回收id` | 文本型 | 当前回收单据ID |
| `局部_销售单据回收单号` | 文本型 | 当前回收单号 |
| `局部_销售单据会员编码` | 文本型 | 当前会员编码 |
| `局部重量合计数值` | 文本型 | 销售商品重量合计 |
| `局部重量单价数值` | 文本型 | 销售商品单价合计 |
| `局部重销售价金额数值` | 文本型 | 销售金额合计 |
| `局部重基本工费合计` | 文本型 | 基本工费合计 |
| `局部重精品工费合计` | 文本型 | 精品工费合计 |
| `局部重销售工费合计` | 文本型 | 销售工费合计 |
| `局部重成本附加费合计` | 文本型 | 成本附加费合计 |
| `局部重销售附加费合计` | 文本型 | 销售附加费合计 |
| `局部回收总重合计数值` | 文本型 | 回收总重合计 |
| `局部回收金重合计数值` | 文本型 | 回收金重合计 |
| `局部回收其他合计数值` | 文本型 | 回收其他费用合计 |
| `局部回收回收合计数值` | 文本型 | 回收金额合计 |
| `局部数量合计数量` | 文本型 | 销售数量合计 |
| `局部实收金额金额数值` | 文本型 | 实收金额合计 |
| `局部_销售单据状态` | 文本型 | 当前销售单据状态 |
| `局部_销售商品编码` | 文本型 | 当前销售商品编码 |
| `局部_销售商品数据编码` | 文本型 | 当前销售商品数据编码 |
| `局部_销售单据批发零售` | 文本型 | 批发/零售标记 |
| `局部_表格编辑状态` | 整数型 | 销售表格编辑状态标记 |
| `局部_收支编辑状态` | 整数型 | 收支表格编辑状态标记 |
| `局部_回收编辑状态` | 整数型 | 回收表格编辑状态标记 |
| `集_保存事务进行中` | 整数型 | 保存事务进行中标记 |

**子程序列表（58个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品销售编辑_创建完毕` | 窗口初始化：加载原始订单数据 |
| `_窗口_商品销售编辑_可否被关闭` | 关闭前检查：未保存数据提示 |
| `_窗口_商品销售编辑_尺寸被改变` | 窗口自适应布局 |
| `_销售编辑_SQL文本处理` | SQL文本安全处理（单引号转义） |
| `_销售编辑_UTF8SQL文本处理` | UTF8编码SQL文本处理 |
| `_销售编辑_写库影响行` | 获取写库影响行数 |
| `_销售编辑_写库SQL校验` | 写库SQL执行校验 |
| `_销售编辑_回补库存行` | 回补库存（增加库存数量和金重） |
| `_销售编辑_是否仅金重商品` | 判断商品是否仅按金重计价 |
| `_高级表格1_加载表头` | 加载销售表格表头 |
| `_高级表格2_加载表头` | 加载回收表格表头 |
| `_高级表格3_加载表头` | 加载支付表格表头 |
| `_高级表格4_加载表头` | 加载合计表格表头 |
| `_高级表格1_加载数据` | 加载原始销售数据 |
| `_高级表格2_加载数据` | 加载原始回收数据 |
| `_高级表格3_加载数据` | 加载原始支付数据 |
| `_子程序_数据统计汇总` | 统计销售/回收/支付合计数据 |
| `子程序_删除表格1` | 清空销售表格 |
| `子程序_删除表格2` | 清空回收表格 |
| `子程序_删除表格3` | 清空支付表格 |
| `_子程序_添加销售` | 添加销售行 |
| `_按钮_回收加_被单击` | 添加回收行 |
| `_按钮_回收减_被单击` | 删除回收行 |
| `_高级表格1_光标位置改变` | 销售表格光标位置改变事件 |
| `_高级表格2_光标位置改变` | 回收表格光标位置改变事件 |
| `_高级表格3_光标位置改变` | 支付表格光标位置改变事件 |
| `_高级表格1_将被编辑` | 销售表格开始编辑事件 |
| `_高级表格1_结束编辑` | 销售表格结束编辑事件 |
| `_高级表格2_结束编辑` | 回收表格结束编辑事件 |
| `_业务员组合框_将弹出列表` | 业务员下拉框弹出列表 |
| `_执行导购员列表SQL` | 执行导购员列表查询SQL |
| `_超级按钮_保存_被单击` | 保存编辑：事务处理（回补原库存→扣减新库存→更新订单→更新明细→更新收款） |
| `_超级按钮_重置_被单击` | 重置表单（重新加载原始数据） |
| `_单选框_零售_选中状态改变` | 切换零售模式 |
| `_单选框_批发_选中状态改变` | 切换批发模式 |
| `_组合框会员查找_将弹出列表` | 会员搜索下拉框弹出列表 |
| `_组合框会员查找_列表项被选择` | 会员选中事件 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |
| `_销售编辑_格式二位小数` | 格式化两位小数 |
| `_销售编辑_格式三位小数UTF8` | 格式化三位小数（UTF8编码） |

**核心SQL**:
```sql
-- 加载原始出库订单
SELECT * FROM xipunum_erp_outbound_order WHERE id='{订单ID}' LIMIT 1

-- 加载原始出库明细
SELECT * FROM xipunum_erp_outbound WHERE order_id='{订单ID}'

-- 加载原始回收订单
SELECT * FROM xipunum_erp_retreat_order WHERE id='{回收订单ID}' LIMIT 1

-- 加载原始回收明细
SELECT * FROM xipunum_erp_retreat WHERE order_id='{回收订单ID}'

-- 加载原始收款记录
SELECT * FROM xipunum_erp_shoukuan WHERE order_id='{订单ID}'

-- 回补库存
UPDATE xipunum_erp_shop_kucun SET quantity = quantity + '{数量}',jinzhong = jinzhong + '{金重}' WHERE poduct_code = '{编码}' AND kufang='{库房id}'

-- 扣减库存
UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '{数量}',jinzhong = jinzhong - '{金重}' WHERE poduct_code = '{编码}' AND kufang='{库房id}'

-- 更新出库订单
UPDATE xipunum_erp_outbound_order SET customer_code='{会员编码}',ying_amount='{应收}',settlement='{实收}',pling='{批零}',salesman='{业务员}',updatetime='{时间}' WHERE id='{订单ID}' LIMIT 1

-- 更新出库明细
UPDATE xipunum_erp_outbound SET quantity='{数量}',net_weight='{金重}',gold_price='{克价}',settlement='{结算金额}',shopping_guide='{导购员}',updatetime='{时间}' WHERE id='{明细ID}' LIMIT 1

-- 更新回收订单
UPDATE xipunum_erp_retreat_order SET customer_code='{会员编码}',settlement='{结算}',jin_zhong='{总金重}',updatetime='{时间}' WHERE id='{回收订单ID}' LIMIT 1

-- 更新回收明细
UPDATE xipunum_erp_retreat SET product_name='{名称}',jin_zhong='{金重}',chengse='{成色}',price='{克价}',qita_price='{其他金额}',retreat_amount='{回收金额}',remarks='{备注}',updatetime='{时间}' WHERE id='{明细ID}' LIMIT 1

-- 删除收款记录
DELETE FROM xipunum_erp_shoukuan WHERE order_id='{订单ID}'

-- 添加收款记录
增加记录("xipunum_erp_shoukuan", order_id, pay_name, amount, creationtime)

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_retreat`、`xipunum_erp_retreat_order`、`xipunum_erp_member`、`xipunum_erp_member_cq`、`xipunum_erp_shoukuan`、`xipunum_erp_pay`、`xipunum_erp_voucher`、`xipunum_erp_ksiamges`、`xipunum_erp_specs`、`xipunum_erp_category`、`xipunum_erp_user`、`xipunum_erp_type`、`xipunum_erp_xitong_log`

#### 7.5.4 商品入库添加 (窗口_商品入库添加)
**功能**: 采购入库操作

**程序集变量（11个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `删除按钮` | 表格按钮 | 表格删除按钮 |
| `集_列号` | 整数型 | 表格当前列号 |
| `集_行号` | 整数型 | 表格当前行号 |
| `局部_入库商品数量` | 整数型 | 已入库商品数量 |
| `集_数据报表` | 报表 | 报表对象 |
| `集_数据显示器` | 查询显示器 | 数据显示器对象 |
| `局部_打印入库单号` | 文本型 | 打印入库单号 |
| `局部_商品是否镶嵌` | 文本型 | 商品是否镶嵌（"镶嵌"/"非镶嵌"） |
| `局部_商品品类id` | 文本型 | 商品品类ID |

**子程序列表（21个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品入库添加_创建完毕` | 窗口初始化：初始化基础参数、加载表格 |
| `_窗口_商品入库添加_尺寸被改变` | 窗口自适应布局 |
| `_窗口_商品入库添加_可否被关闭` | 关闭前检查：未保存数据提示 |
| `子程序_删除表格1` | 清空表格 |
| `_高级表格1_光标位置改变` | 表格光标位置改变事件 |
| `_入库基础_初始化` | 初始化基础参数：加载工厂/来源/库房/结算方式下拉框、标签样式 |
| `_工厂名称组合框_将弹出列表` | 工厂下拉框弹出列表（支持模糊搜索） |
| `_工具条_通用_被单击` | 工具条按钮点击事件（添加商品/保存/重置/标签打印） |
| `_高级表格1_加载表头` | 加载表格表头（37列：序号/编码/名称/品类/规格/款号/材质/圈口/面宽/厚度/工厂成色/公司成色/单件重/数量/金重/损耗/含耗重/重量/单位/石重/石头数/副石重/副石头数/成本单价/系数/成本工费/参考工费/销售工费/成本附加费/销售附加费/销售价/备注/主石色/图片地址/原料价/模具号/操作） |
| `_高级表格1_加载数据` | 加载表格数据（25行空行） |
| `_高级表格1_数据统计` | 统计合计数据（数量/金重/重量/石重/石头数/副石重/副石头数/成本价/工费/附加费） |
| `_超级按钮_添加_被单击` | 添加商品：验证基础参数→打开商品信息添加窗口 |
| `_超级按钮_重置_被单击` | 重置表单 |
| `_高级表格1_按钮被点击` | 表格按钮点击事件：删除商品（从数据库删除） |
| `入库添加_格式三位小数` | 格式化三位小数 |
| `入库添加_格式二位小数` | 格式化两位小数 |
| `_高级表格1_结束编辑` | 表格结束编辑事件：金重/重量/损耗/工费/附加费联动计算 |
| `_超级按钮_保存_被单击` | 保存入库单据：事务处理（生成订单→保存商品信息→保存入库明细→更新库存→打印标签/单据） |
| `_超级按钮_标签打印_被单击` | 标签打印：调用标签打印程序 |
| `_超级按钮_单据打印_被单击` | 单据打印：调用报表打印 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |

**核心SQL**:
```sql
-- 加载工厂下拉框
SELECT title FROM xipunum_erp_about WHERE 1=1 ORDER BY id ASC

-- 加载来源下拉框
SELECT title FROM xipunum_erp_type WHERE type='商品来源' AND superior='0' ORDER BY id DESC

-- 加载库房下拉框
SELECT id AS akufang,CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in ({权限}) UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({权限}) ORDER BY akufang = '0' DESC, akufang

-- 加载结算方式下拉框
SELECT title FROM xipunum_erp_type WHERE type='结算方式' AND superior='0' ORDER BY id DESC

-- 工厂模糊搜索
SELECT title FROM xipunum_erp_about WHERE title LIKE '%{关键字}%' OR jianxie LIKE '%{关键字}%' ORDER BY id ASC

-- 保存商品信息
增加记录("xipunum_erp_shop", poduct_code, fu_code, product_name, caizhi, chengse, specification_id, item_number, quantity, net_weight, single, cost_price, sales_price, ...)

-- 保存入库订单
增加记录("xipunum_erp_store_order", store_number, gongshi, remarks, state, cjuser, creationtime)

-- 保存入库明细
增加记录("xipunum_erp_store", order_id, poduct_code, product_name, quantity, net_weight, gold_price, cost_price, remarks, creationtime)

-- 更新库存（存在则更新，不存在则插入）
UPDATE xipunum_erp_shop_kucun SET quantity = quantity + '{数量}',jinzhong = jinzhong + '{金重}' WHERE poduct_code = '{编码}' AND kufang='{库房id}'
-- 或
增加记录("xipunum_erp_shop_kucun", poduct_code, quantity, jinzhong, kufang)

-- 删除商品
DELETE FROM xipunum_erp_shop WHERE poduct_code='{编码}'
DELETE FROM xipunum_erp_zhengshu WHERE poduct_code='{编码}'

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_store`、`xipunum_erp_store_order`、`xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_about`、`xipunum_erp_type`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_ksiamges`、`xipunum_erp_zhengshu`、`xipunum_erp_voucher`、`xipunum_erp_xitong_log`

#### 7.5.5.1 入库库房修改 (窗口_入库库房修改)
**功能**: 修改入库订单的库房

**程序集变量（3个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_主窗口行号` | 整数型 | 主窗口选中行号 |
| `局部_订单id` | 文本型 | 入库订单ID |
| `局部_正在保存` | 逻辑型 | 保存中标记 |

**子程序列表（4个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_入库库房修改_创建完毕` | 窗口初始化：加载库房下拉框 |
| `_窗口_入库库房修改_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX1_鼠标左键单击` | 保存库房修改：更新订单库房、更新库存、添加日志 |

**核心SQL**:
```sql
-- 查询订单信息
SELECT cjuser,id,creationtime FROM xipunum_erp_store_order WHERE odd_numbers='{单号}' LIMIT 1

-- 查询入库商品
SELECT a.poduct_code AS bianam,a.creationtime AS creationtime FROM xipunum_erp_store AS a INNER JOIN xipunum_erp_store_order AS b ON b.id=a.order_id WHERE b.odd_numbers='{单号}' ORDER BY a.id ASC

-- 更新订单库房
UPDATE xipunum_erp_store_order SET kufang='{库房ID}',updatetime='{时间}' WHERE odd_numbers='{单号}' LIMIT 1

-- 更新库存库房
UPDATE xipunum_erp_shop_kucun SET kufang='{新库房ID}' WHERE poduct_code='{编码}' AND kufang='{旧库房ID}'

-- 添加入库日志
增加记录("xipunum_erp_store_log", order_id, user, conter, creationtime)

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_store_order`、`xipunum_erp_store`、`xipunum_erp_shop_kucun`、`xipunum_erp_type`、`xipunum_erp_store_log`、`xipunum_erp_xitong_log`

---

#### 7.5.5.2 入库工厂修改 (窗口_入库工厂修改)
**功能**: 修改入库订单的工厂、成本工费、成本附加费、原料克价

**程序集变量（3个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_主窗口行号` | 整数型 | 主窗口选中行号 |
| `局部_订单id` | 文本型 | 入库订单ID |
| `局部_正在保存` | 逻辑型 | 保存中标记 |

**子程序列表（5个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_入库工厂修改_创建完毕` | 窗口初始化：加载工厂下拉框 |
| `_窗口_入库工厂修改_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX1_鼠标左键单击` | 保存修改：更新工厂/工费/附加费/原料价、重新计算成本价、添加日志 |
| `工厂修改_格式二位小数` | 格式化两位小数 |

**核心SQL**:
```sql
-- 更新订单工厂
UPDATE xipunum_erp_store_order SET factory='{工厂ID}',updatetime='{时间}' WHERE odd_numbers='{单号}' LIMIT 1

-- 更新原料克价
UPDATE xipunum_erp_store_order SET gold_price='{原料克价}',updatetime='{时间}' WHERE odd_numbers='{单号}' LIMIT 1

-- 更新成本工费和附加费
UPDATE xipunum_erp_store SET basic_cost='{成本工费}',company_surcharge='{成本附加费}',cost_price='{成本价}' WHERE order_id='{订单ID}'

-- 添加入库日志
增加记录("xipunum_erp_store_log", order_id, user, conter, creationtime)

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_store_order`、`xipunum_erp_store`、`xipunum_erp_store_log`、`xipunum_erp_xitong_log`

---

#### 7.5.5.3 商品入库批量修改 (窗口_商品入库批量修改)
**功能**: 批量修改入库商品信息

**程序集变量（2个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `入库品类数据id` | 文本型 | 入库品类数据ID |
| `组合框列名称数据` | 文本型 | 组合框列名称数据 |

**子程序列表（6个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品入库批量修改_创建完毕` | 窗口初始化：加载列名称下拉框 |
| `_窗口_商品入库批量修改_尺寸被改变` | 窗口自适应布局 |
| `_按钮EX_保存_鼠标左键单击` | 保存批量修改 |
| `_按钮EX_重置_鼠标左键单击` | 重置表单 |
| `格式三位小数` | 格式化三位小数 |
| `格式二位小数` | 格式化两位小数 |

**核心SQL**:
```sql
-- 批量更新入库商品
UPDATE xipunum_erp_store SET {字段}='{值}' WHERE order_id='{订单ID}'
```

**涉及表**: `xipunum_erp_store`、`xipunum_erp_store_order`

---

#### 7.5.1.2 商品成品数据修改 (窗口_商品成品数据修改)
**功能**: 修改成品商品信息

**程序集变量（19个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_数据报表` | 报表 | 报表对象 |
| `集_数据显示器` | 查询显示器 | 数据显示器对象 |
| `局部_成品商品编码` | 文本型 | 成品商品编码 |
| `局部_成品商品数据编码` | 文本型 | 成品商品数据编码 |
| `局部_商品成色` | 文本型 | 商品成色 |
| `局部_商品材质` | 文本型 | 商品材质 |
| `局部_商品规格` | 文本型 | 商品规格 |
| `局部_商品品类` | 文本型 | 商品品类 |
| `局部_商品是否镶嵌` | 文本型 | 商品是否镶嵌 |
| `局部_品类多数量` | 文本型 | 品类多数量 |
| `局部_商品入库数量` | 文本型 | 商品入库数量 |
| `局部_商品入库总重` | 文本型 | 商品入库总重 |
| `局部_商品入库金重` | 文本型 | 商品入库金重 |
| `局部_商品商品损耗` | 文本型 | 商品损耗 |
| `局部_商品含耗重` | 文本型 | 商品含耗重 |
| `局部_商品款号数据` | 文本型 | 商品款号数据 |
| `局部_商品图片地址` | 文本型 | 商品图片地址 |
| `局部_商品库存id信息` | 文本型 | 商品库存ID信息 |

**子程序列表（30+个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品成品数据修改_创建完毕` | 窗口初始化 |
| `_窗口_商品成品数据修改_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX_重置_鼠标左键单击` | 重置表单 |
| `_按钮EX_保存_鼠标左键单击` | 保存修改 |
| `_组合框_品类名称_内容被改变` | 品类选择联动 |
| `_组合框_商品材质_内容被改变` | 材质选择联动 |
| `_组合框_商品成色_内容被改变` | 成色选择联动 |
| `_组合框_商品规格_内容被改变` | 规格选择联动 |
| `_编辑框_重量_焦点事件` | 重量输入框焦点事件 |
| `_编辑框_重量_键盘事件` | 重量输入框键盘事件 |
| `_商品信息_参数数值计算` | 重量/金重/损耗/含耗重/成本价/销售价联动计算 |

**核心SQL**:
```sql
-- 查询商品信息
SELECT * FROM xipunum_erp_shop WHERE poduct_code='{编码}'

-- 更新商品信息
UPDATE xipunum_erp_shop SET product_name='{名称}',caizhi='{材质}',chengse='{成色}',specification_id='{规格ID}',item_number='{款号}',quantity='{数量}',net_weight='{金重}',single='{单件重}',cost_price='{成本价}',sales_price='{销售价}',... WHERE poduct_code='{编码}'

-- 更新库存信息
UPDATE xipunum_erp_shop_kucun SET quantity='{数量}',jinzhong='{金重}' WHERE poduct_code='{编码}' AND kufang='{库房ID}'
```

**涉及表**: `xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_ksiamges`

---

#### 7.5.6 商品信息调拨 (窗口_商品信息调拨)
**功能**: 商品在不同库房间调拨

**程序集变量（12个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_商品调拨方法` | 文本型 | 调拨方法（"添加"/"扫码"） |
| `集_行号` | 整数型 | 表格当前行号 |
| `集_列号` | 整数型 | 表格当前列号 |
| `删除按钮` | 表格按钮 | 表格删除按钮 |
| `局部_订单是否选中` | 整数型 | 是否选中订单（-1=新建，其他=详情） |
| `局部_调拨商品编码` | 文本型 | 调拨商品编码 |
| `局部_调拨商品数据编码` | 文本型 | 调拨商品数据编码 |
| `局部_库房数据id` | 文本型 | 库房数据ID |

**子程序列表（26个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品信息调拨_创建完毕` | 窗口初始化：根据订单状态设置界面 |
| `_窗口_商品信息调拨_尺寸被改变` | 窗口自适应布局 |
| `_单选框_手动_被单击` | 切换手动输入模式 |
| `_单选框_扫码_被单击` | 切换扫码输入模式 |
| `_调拨基础参数_被单击` | 加载基础参数：库房下拉框 |
| `_超级按钮_重置_被单击` | 重置表单 |
| `_商品编码_编辑框_按下某键` | 商品编码回车键处理 |
| `_高级表格1_光标位置改变` | 表格光标位置改变事件 |
| `子程序_删除表格` | 清空表格 |
| `_高级表格1_加载表头` | 加载表格表头（14列：序号/编码/品类/规格/名称/款号/单件重/调拨数量/调拨金重/原库房/新库房/库存数量/库存重量/操作） |
| `_调拨详情_被点击` | 加载调拨详情数据 |
| `_商品编码_编辑框_内容被改变` | 商品编码内容改变事件（扫码模式自动触发） |
| `_单选框_调入_被单击` | 切换调入模式 |
| `_单选框_调出_被单击` | 切换调出模式 |
| `_库房名称组合框_列表项被选择` | 库房下拉框选中事件：获取库房ID |
| `_调出库房名称组合框_列表项被选择` | 调出库房下拉框选中事件：获取库房ID |
| `_超级按钮_添加_被单击` | 添加调拨商品：验证商品存在性→验证库存→验证库房→添加到表格 |
| `_超级按钮_添加_数据加载` | 加载调拨商品数据到表格 |
| `_高级表格1_统计数据` | 统计合计数据 |
| `_高级表格1_按钮被点击` | 表格按钮点击事件：删除调拨商品 |
| `_超级按钮_保存_被单击` | 保存调拨单据：事务处理（生成订单→保存调拨明细→更新库存双向） |
| `_超级按钮_全选_被单击` | 全选商品 |
| `_超级按钮_反选_被单击` | 反选商品 |
| `_超级按钮_提取编码_被单击` | 提取选中商品编码 |
| `_超级按钮_标签打印_被单击` | 标签打印 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |

**核心SQL**:
```sql
-- 加载库房下拉框
SELECT title FROM xipunum_erp_type WHERE type='商铺' and superior='0' order by id desc

-- 验证商品存在性
SELECT * FROM xipunum_erp_shop where (poduct_code='{编码}' or fu_code='{编码}') order by id ASC

-- 验证商品库存
SELECT a.poduct_code AS apoduct_code FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE a.kufang = '{库房id}' AND (a.quantity > 0 or a.jinzhong >0) AND (b.poduct_code = '{编码}' OR b.fu_code = '{编码}') ORDER BY a.id DESC

-- 加载商品详情
SELECT ... FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = b.item_number LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) WHERE a.kufang = '{库房id}' AND b.poduct_code = '{编码}'

-- 加载调拨详情
SELECT ... FROM xipunum_erp_transfer AS a INNER JOIN xipunum_erp_transfer_order AS b ON b.id = a.order_id AND b.transfer_umber = '{单号}' INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_category AS d ON d.id = c.category_id LEFT JOIN xipunum_erp_specs AS e ON e.id = c.specification_id LEFT JOIN xipunum_erp_type AS f ON f.id = a.ykufang LEFT JOIN xipunum_erp_type AS g ON g.id = a.xkufang WHERE 1 = 1 ORDER BY a.id asc

-- 统计合计数据
SELECT sum(a.quantity) AS aquantity,CAST(ROUND(sum(a.jinzhong),3) AS DECIMAL(10,3)) AS zhongliang FROM xipunum_erp_transfer AS a INNER JOIN xipunum_erp_transfer_order AS b ON b.id = a.order_id AND b.transfer_umber = '{单号}'

-- 保存调拨订单
增加记录("xipunum_erp_transfer_order", transfer_umber, kufang_from, kufang_to, state, cjuser, creationtime)

-- 保存调拨明细
增加记录("xipunum_erp_transfer", order_id, poduct_code, quantity, jinzhong, ykufang, xkufang, remarks, creationtime)

-- 更新库存（调出库房扣减）
UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '{数量}',jinzhong = jinzhong - '{金重}' WHERE poduct_code = '{编码}' AND kufang='{调出库房id}'

-- 更新库存（调入库房增加）
UPDATE xipunum_erp_shop_kucun SET quantity = quantity + '{数量}',jinzhong = jinzhong + '{金重}' WHERE poduct_code = '{编码}' AND kufang='{调入库房id}'

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_transfer`、`xipunum_erp_transfer_order`、`xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_type`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_ksiamges`、`xipunum_erp_xitong_log`

#### 7.5.7 商品信息退库 (窗口_商品信息退库)
**功能**: 商品退回工厂

**程序集变量（10个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_行号` | 整数型 | 表格当前行号 |
| `集_列号` | 整数型 | 表格当前列号 |
| `删除按钮` | 表格按钮 | 表格删除按钮 |
| `局部_订单是否选中` | 整数型 | 是否选中订单（-1=新建，其他=详情） |
| `局部_退库商品编码` | 文本型 | 退库商品编码 |
| `局部_退库商品数据编码` | 文本型 | 退库商品数据编码 |
| `局部_退库按钮名称` | 文本型 | 退库按钮名称（"添加"/"详情"/"再次提交"） |

**子程序列表（18个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品信息退库_创建完毕` | 窗口初始化：根据按钮名称设置界面 |
| `_窗口_商品信息退库_尺寸被改变` | 窗口自适应布局 |
| `_高级表格1_加载表头` | 加载表格表头（15列：序号/编码/品类/规格/名称/款号/单件重/退库数量/退库金重/退库重量/原库房/新库房/库存数量/库存重量/操作） |
| `子程序_删除表格` | 清空表格 |
| `_单选框_扫码_被单击` | 切换扫码输入模式 |
| `_单选框_手动_被单击` | 切换手动输入模式 |
| `_超级按钮_重置_被单击` | 重置表单 |
| `_商品编码_编辑框_按下某键` | 商品编码回车键处理 |
| `_高级表格1_光标位置改变` | 表格光标位置改变事件 |
| `_退库详情_被点击` | 加载退库详情数据 |
| `_超级按钮_添加_被单击` | 添加退库商品：验证商品存在性→验证库存→验证库房→添加到表格 |
| `_超级按钮_添加_数据加载` | 加载退库商品数据到表格 |
| `_高级表格1_统计数据` | 统计合计数据 |
| `_高级表格1_按钮被点击` | 表格按钮点击事件：删除退库商品 |
| `_超级按钮_保存_被单击` | 保存退库单据：事务处理（生成订单→保存退库明细→扣减库存） |
| `_高级表格1_再次提交` | 加载再次提交数据 |
| `子程序_删除表格5` | 清空临时删除表格 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |

**核心SQL**:
```sql
-- 验证商品存在性
SELECT * FROM xipunum_erp_shop where (poduct_code='{编码}' or fu_code='{编码}') order by id ASC

-- 验证商品库存
SELECT a.poduct_code AS apoduct_code FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE a.kufang = '{库房id}' AND (a.quantity > 0 or a.jinzhong >0) AND (b.poduct_code = '{编码}' OR b.fu_code = '{编码}') ORDER BY a.id DESC

-- 加载商品详情
SELECT ... FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = b.item_number LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) WHERE a.kufang = '{库房id}' AND b.poduct_code = '{编码}'

-- 加载退库详情
SELECT ... FROM xipunum_erp_tuiku AS a INNER JOIN xipunum_erp_tuiku_order AS b ON b.id = a.order_id AND b.tuiku_umber = '{单号}' INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = c.item_number LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = c.specification_id LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) LEFT JOIN xipunum_erp_type AS g ON g.id = a.ykufang LEFT JOIN xipunum_erp_type AS h ON h.id = a.xkufang WHERE 1 = 1 ORDER BY a.id ASC

-- 统计合计数据
SELECT sum(tol.aquantity) as aquantity,sum(tol.jinzhong) as jinzhong,sum(tol.zongzhong) as zhongliang FROM (...) as tol

-- 保存退库订单
增加记录("xipunum_erp_tuiku_order", tuiku_umber, total, state, remarks, cjuser, creationtime)

-- 保存退库明细
增加记录("xipunum_erp_tuiku", order_id, poduct_code, quantity, jinzhong, ykufang, xkufang, remarks, creationtime)

-- 更新库存
UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '{数量}',jinzhong = jinzhong - '{金重}' WHERE poduct_code = '{编码}' AND kufang='{库房id}'

-- 更新退库订单
UPDATE xipunum_erp_tuiku_order SET total='{总重}',state='正常',remarks='{备注}',cjuser='{账户}',updatetime='{时间}' WHERE tuiku_umber='{单号}' LIMIT 1

-- 删除退库明细
DELETE FROM xipunum_erp_tuiku WHERE id='{明细ID}'

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_tuiku`、`xipunum_erp_tuiku_order`、`xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_type`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_ksiamges`、`xipunum_erp_xitong_log`

#### 7.5.8 商品信息回收 (窗口_商品信息回收)
**功能**: 旧料回收（从客户回收旧金旧料）

**程序集变量（5个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_行号` | 整数型 | 表格当前行号 |
| `集_列号` | 整数型 | 表格当前列号 |
| `删除按钮` | 表格按钮 | 表格删除按钮 |
| `局部_订单是否选中` | 整数型 | 是否选中订单（-1=新建，其他=详情） |
| `局部_回收编辑状态` | 整数型 | 回收表格编辑状态标记 |

**子程序列表（17个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品信息回收_创建完毕` | 窗口初始化：根据订单状态设置界面 |
| `_窗口_商品信息回收_尺寸被改变` | 窗口自适应布局 |
| `_高级表格1_加载表头` | 加载表格表头（12列：序号/商品名称/数量/总重/金重/成色/回收克价/其他费用/回收金额/备注/导购员/操作） |
| `子程序_删除表格` | 清空表格 |
| `_高级表格1_光标位置改变` | 表格光标位置改变事件 |
| `_回收详情_被点击` | 加载回收详情数据 |
| `_回收数据添加_被单击` | 添加回收行 |
| `_高级表格1_结束编辑` | 表格结束编辑事件：回收金额联动计算 |
| `_高级表格1_按钮被点击` | 表格按钮点击事件：删除回收商品 |
| `_高级表格1_统计数据` | 统计合计数据 |
| `_按钮_保存_被单击` | 保存回收单据：事务处理（生成订单→保存回收明细→更新会员存欠） |
| `_按钮_重置_被单击` | 重置表单 |
| `_会员姓名_编辑框_内容被改变` | 会员姓名内容改变事件 |
| `_联系电话_编辑框_内容被改变` | 联系电话内容改变事件 |
| `_业务员_编辑框_内容被改变` | 业务员内容改变事件 |
| `_税点_编辑框_内容被改变` | 税点内容改变事件 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |

**核心SQL**:
```sql
-- 加载回收名称下拉框
SELECT * FROM xipunum_erp_retreat_title WHERE 1=1 order by id asc

-- 加载导购员列表
SELECT name FROM xipunum_erp_user where state='0' order by id ASC

-- 加载回收详情
SELECT ... FROM xipunum_erp_retreat AS a INNER JOIN xipunum_erp_retreat_order AS b ON b.id = a.order_id AND b.retrea_umber = '{单号}' WHERE 1 = 1 ORDER BY a.id ASC

-- 统计合计数据
SELECT SUM(quantity) AS aquantity,SUM(jin_zhong) AS ajinzhong,SUM(retreat_amount) AS ashishou FROM xipunum_erp_retreat WHERE order_id='{订单ID}'

-- 保存回收订单
增加记录("xipunum_erp_retreat_order", retrea_umber, customer_code, settlement, jin_zhong, cjuser, creationtime)

-- 保存回收明细
增加记录("xipunum_erp_retreat", order_id, product_name, quantity, jin_zhong, chengse, price, qita_price, retreat_amount, remarks, shopping_guide, creationtime)

-- 更新会员存欠
增加记录("xipunum_erp_member_cq", customer_code, cunqu, type, number, remarks, kufang, cjuser, creationtime)

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_retreat`、`xipunum_erp_retreat_order`、`xipunum_erp_retreat_title`、`xipunum_erp_member`、`xipunum_erp_member_cq`、`xipunum_erp_user`、`xipunum_erp_xitong_log`

#### 7.5.9 商品信息预售 (窗口_商品信息预售)
**功能**: 商品预售/预定金管理

**程序集变量（11个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_行号` | 整数型 | 表格当前行号 |
| `集_列号` | 整数型 | 表格当前列号 |
| `删除按钮` | 表格按钮 | 表格删除按钮 |
| `图片上传` | 表格按钮 | 图片上传按钮 |
| `局部_订单是否选中` | 整数型 | 是否选中订单（-1=新建，其他=详情） |
| `局部_图片路径` | 文本型 | 上传图片路径 |
| `局部_图片响应` | 文本型 | 图片上传响应 |
| `局部_图片名称` | 文本型 | 图片名称 |

**子程序列表（17个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品信息预售_创建完毕` | 窗口初始化：根据订单状态设置界面 |
| `_窗口_商品信息预售_尺寸被改变` | 窗口自适应布局 |
| `_高级表格1_加载表头` | 加载表格表头（7列：序号/商品名称/数量/备注/图片地址/图片/操作） |
| `子程序_删除表格` | 清空表格 |
| `_高级表格1_光标位置改变` | 表格光标位置改变事件：加载图片预览 |
| `_预售详情_被点击` | 加载预售详情数据 |
| `_预售数据添加_被单击` | 添加预售行 |
| `_高级表格1_结束编辑` | 表格结束编辑事件 |
| `_高级表格1_按钮被点击` | 表格按钮点击事件：删除预售商品/上传图片 |
| `_高级表格1_统计数据` | 统计合计数据 |
| `_按钮_保存_被单击` | 保存预售单据：事务处理（生成订单→保存预售明细→上传图片） |
| `_按钮_重置_被单击` | 重置表单 |
| `_会员姓名_编辑框_内容被改变` | 会员姓名内容改变事件 |
| `_联系电话_编辑框_内容被改变` | 联系电话内容改变事件 |
| `_订金_编辑框_内容被改变` | 订金内容改变事件 |
| `_业务员_编辑框_内容被改变` | 业务员内容改变事件 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |

**核心SQL**:
```sql
-- 加载预售详情
SELECT a.id AS aid,a.product_name AS product_name,a.quantity AS auantity,a.remarks AS aremarks,a.images AS aimages,CASE WHEN a.images = '' THEN '无' ELSE '有' END AS tupian,a.creationtime AS acreationtime FROM xipunum_erp_presale AS a INNER JOIN xipunum_erp_presale_order AS b ON b.id = a.order_id AND b.presale_umber = '{单号}' WHERE 1 = 1 ORDER BY a.id ASC

-- 统计合计数据
SELECT SUM(quantity) AS aquantity FROM xipunum_erp_presale WHERE order_id='{订单ID}'

-- 保存预售订单
增加记录("xipunum_erp_presale_order", presale_umber, customer_code, deposit, state, cjuser, creationtime)

-- 保存预售明细
增加记录("xipunum_erp_presale", order_id, product_name, quantity, remarks, images, creationtime)

-- 更新预售图片
UPDATE xipunum_erp_presale SET images='{图片URL}' WHERE id='{明细ID}'

-- 删除预售明细
DELETE FROM xipunum_erp_presale WHERE id='{明细ID}'

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_presale`、`xipunum_erp_presale_order`、`xipunum_erp_member`、`xipunum_erp_xitong_log`

#### 7.5.10 商品信息退货 (窗口_商品信息退货)
**功能**: 客户退货（退回到库存）

**程序集变量（9个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_行号` | 整数型 | 表格当前行号 |
| `集_列号` | 整数型 | 表格当前列号 |
| `删除按钮` | 表格按钮 | 表格删除按钮 |
| `局部_订单是否选中` | 整数型 | 是否选中订单（-1=新建，其他=详情） |
| `局部_退货商品编码` | 文本型 | 退货商品编码 |
| `局部_退货商品数据编码` | 文本型 | 退货商品数据编码 |

**子程序列表（14个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品信息退货_创建完毕` | 窗口初始化：根据订单状态设置界面 |
| `_窗口_商品信息退货_尺寸被改变` | 窗口自适应布局 |
| `_高级表格1_加载表头` | 加载表格表头（14列：序号/编码/名称/送货单号/半成品/单件重/退货数量/退货金重/退货重量/工厂/来源/工厂成色/备注/操作） |
| `子程序_删除表格` | 清空表格 |
| `_单选框_扫码_被单击` | 切换扫码输入模式 |
| `_单选框_手动_被单击` | 切换手动输入模式 |
| `_超级按钮_重置_被单击` | 重置表单 |
| `_商品编码_编辑框_按下某键` | 商品编码回车键处理 |
| `_高级表格1_光标位置改变` | 表格光标位置改变事件 |
| `_退货详情_被点击` | 加载退货详情数据 |
| `_超级按钮_添加_被单击` | 添加退货商品：验证商品存在性→添加到表格 |
| `_高级表格1_统计数据` | 统计合计数据 |
| `_高级表格1_按钮被点击` | 表格按钮点击事件：删除退货商品 |
| `_超级按钮_保存_被单击` | 保存退货单据：事务处理（生成订单→保存退货明细→更新库存） |

**核心SQL**:
```sql
-- 验证商品存在性
SELECT * FROM xipunum_erp_shop where (poduct_code='{编码}' or fu_code='{编码}') order by id ASC

-- 加载退货详情
SELECT ... FROM xipunum_erp_return AS a INNER JOIN xipunum_erp_return_order AS b ON b.id = a.order_id AND b.return_umber = '{单号}' INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = a.poduct_code INNER JOIN xipunum_erp_about AS f ON f.id = a.factory WHERE 1 = 1 ORDER BY a.id ASC

-- 统计合计数据
SELECT SUM(quantity) AS aquantity,SUM(jinzhong) AS ajinzhong,SUM(zhongliang) AS azhongliang FROM xipunum_erp_return WHERE order_id='{订单ID}'

-- 保存退货订单
增加记录("xipunum_erp_return_order", return_umber, customer_code, state, cjuser, creationtime)

-- 保存退货明细
增加记录("xipunum_erp_return", order_id, poduct_code, quantity, jinzhong, zhongliang, delivery, half_product, factory, source, factory_condition, remarks, creationtime)

-- 更新库存
UPDATE xipunum_erp_shop_kucun SET quantity = quantity + '{数量}',jinzhong = jinzhong + '{金重}' WHERE poduct_code = '{编码}' AND kufang='{库房id}'

-- 删除退货明细
DELETE FROM xipunum_erp_return WHERE id='{明细ID}'

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_return`、`xipunum_erp_return_order`、`xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_about`、`xipunum_erp_xitong_log`

---

### 7.6 销售管理模块

#### 7.6.1 商品销售批量修改 (窗口_商品销售批量修改)
**功能**: 批量修改销售单信息

**子程序列表（8个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品销售批量修改_创建完毕` | 窗口初始化：加载导购员列表 |
| `_窗口_商品销售批量修改_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX_重置_鼠标左键单击` | 重置表单 |
| `_编辑框_附加费折扣_内容被改变` | 附加费折扣输入验证（0-1范围，与销售附加费互斥） |
| `_编辑框_销售附加费_内容被改变` | 销售附加费输入验证（与折扣互斥） |
| `_按钮EX_修改_鼠标左键单击` | 批量修改销售单：遍历表格→计算单价/金额→更新本地Access |
| `_批量修改_SQL文本处理` | SQL注入防护：替换单引号 |

**核心SQL**:
```sql
-- 加载导购员列表（全部权限）
SELECT name FROM xipunum_erp_user where state='0' order by id ASC

-- 加载导购员列表（店铺权限）
SELECT name FROM xipunum_erp_user where department='{分组ID}' and state='0' order by id ASC

-- 加载导购员列表（岗位权限）
SELECT name FROM xipunum_erp_user WHERE user in {权限} and state='0' order by id ASC

-- 加载导购员列表（个人权限）
SELECT name FROM xipunum_erp_user WHERE user='{用户}' and state='0' order by id ASC

-- 批量更新销售单（本地Access chuku表）
UPDATE chuku SET 销售克价='{克价}',销售工费='{工费}',销售附加费='{附加费}',实收金额='{实收金额}',销售数量='{数量}',销售重量='{重量}' WHERE 商品编码='{编码}'

-- 更新导购员（本地Access chuku表）
UPDATE chuku SET 导购员='{导购员}' WHERE 商品编码='{编码}'
```

**价格计算逻辑**: 销售单价 = (金重 × (克价 + 工费) + 附加费) / 数量；销售金额 = 金重 × (克价 + 工费) + 附加费

**涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop_kucun`、本地Access `chuku`

---

#### 7.6.2 商品销售客退 (窗口_商品销售客退)
**功能**: 客户退货处理，支持销售单据关联、回收旧料、多种支付方式

**程序集变量（35个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_行号1` | 整数型 | 销售表格当前行号 |
| `集_列号1` | 整数型 | 销售表格当前列号 |
| `集_行号2` | 整数型 | 回收表格当前行号 |
| `集_列号2` | 整数型 | 回收表格当前列号 |
| `集_行号3` | 整数型 | 支付表格当前行号 |
| `集_列号3` | 整数型 | 支付表格当前列号 |
| `集_数据报表` | 报表 | 报表对象 |
| `集_数据显示器` | 查询显示器 | 数据显示器对象 |
| `删除按钮` | 表格按钮 | 表格删除按钮 |
| `局部销售金重合计` | 文本型 | 销售金重合计 |
| `局部销售金额合计` | 文本型 | 销售金额合计 |
| `局部实销金额合计` | 文本型 | 实销金额合计 |
| `局部销售数量合计` | 文本型 | 销售数量合计 |
| `局部销售工费合计` | 文本型 | 销售工费合计 |
| `局部销售附加费合计` | 文本型 | 销售附加费合计 |
| `局部实退工费合计` | 文本型 | 实退工费合计 |
| `局部实退附加费合计` | 文本型 | 实退附加费合计 |
| `局部应退金额合计` | 文本型 | 应退金额合计 |
| `局部实退金额合计` | 文本型 | 实退金额合计 |
| `局部回收总重合计数值` | 文本型 | 回收总重合计 |
| `局部回收金重合计数值` | 文本型 | 回收金重合计 |
| `局部回收其他合计数值` | 文本型 | 回收其他费用合计 |
| `局部回收回收合计数值` | 文本型 | 回收金额合计 |
| `局部_批发库存料数值` | 文本型 | 批发会员库存料数值 |
| `局部_批发库存元数值` | 文本型 | 批发会员库存元数值 |
| `局部成本工费合计` | 文本型 | 成本工费合计 |
| `局部参考工费合计` | 文本型 | 参考工费合计 |
| `局部_商品销售会员id` | 文本型 | 会员ID |
| `局部_表格编辑状态` | 整数型 | 表格编辑状态标记 |
| `局部_收支编辑状态` | 整数型 | 收支表格编辑状态标记 |
| `局部_回收编辑状态` | 整数型 | 回收表格编辑状态标记 |
| `局部_数据据商品编码` | 文本型 | 当前商品编码 |
| `局部_数据据商品订单id` | 文本型 | 当前商品订单ID |

**子程序列表（50+个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品销售客退_创建完毕` | 窗口初始化：生成单据号、初始化表格 |
| `_窗口_商品销售客退_可否被关闭` | 关闭前检查：未保存数据提示 |
| `_窗口_商品销售客退_尺寸被改变` | 窗口自适应布局 |
| `_高级表格1_加载表头` | 加载销售表格表头 |
| `_高级表格2_加载表头` | 加载回收表格表头 |
| `_高级表格3_加载表头` | 加载支付表格表头 |
| `_高级表格1_加载表格` | 加载销售表格数据 |
| `_高级表格1_读取数据` | 读取选中销售单据数据 |
| `_高级表格1_结束编辑` | 销售表格结束编辑事件 |
| `_高级表格2_结束编辑` | 回收表格结束编辑事件 |
| `_子程序_数据统计汇总` | 统计销售/回收/支付合计数据 |
| `_按钮_保存_被单击` | 保存客退单据：事务处理 |
| `_按钮_重置_被单击` | 重置表单 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |

**核心SQL**:
```sql
-- 查询销售单据商品
SELECT b.poduct_code as bianma,b.fu_code as fubianma,c.settlement_number as danhao,... FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_outbound_order AS c ON c.id = a.order_id WHERE (a.poduct_code='{编码}' OR b.fu_code='{编码}') AND a.kufang='{库房}' AND a.xsstate='0' ORDER BY a.creationtime DESC

-- 保存客退订单
增加记录("xipunum_erp_outbound_order", settlement_number, customer_code, ying_amount, settlement, pling, salesman, cjuser, creationtime)

-- 更新库存（回补）
UPDATE xipunum_erp_shop_kucun SET quantity = quantity + '{数量}',jinzhong = jinzhong + '{金重}' WHERE poduct_code = '{编码}' AND kufang='{库房id}'
```

**涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_retreat`、`xipunum_erp_retreat_order`、`xipunum_erp_member`、`xipunum_erp_shoukuan`、`xipunum_erp_pay`、`xipunum_erp_user`

---

#### 7.6.3 商品销售外部单据 (窗口_商品销售外部单据)
**功能**: 外部单据录入

**程序集变量（3个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_数据报表` | 报表 | 报表对象 |
| `集_数据显示器` | 查询显示器 | 数据显示器对象 |
| `高级表格组件句柄` | 整数型 | 表格组件句柄 |

**子程序列表（6个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品销售外部单据_创建完毕` | 窗口初始化：加载表头、加载数据 |
| `_窗口_商品销售外部单据_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX_重置_鼠标左键单击` | 重置表单 |
| `_子程序_加载表头` | 加载表格表头（7列：序号/产品名称/重量/成色/单价/金额/备注） |
| `_子程序_加载数据` | 加载外部单据数据 |

**核心SQL**:
```sql
-- 查询销售订单信息
SELECT a.customer_code AS customer_code,b.name AS NAME,a.creationtime AS chukutime,a.waibu_number as waibu_number FROM xipunum_erp_outbound_order AS a LEFT JOIN xipunum_erp_member AS b ON b.customer_code = a.customer_code WHERE a.settlement_number='{单号}'

-- 查询销售明细
SELECT a.product_name AS product_name,a.net_weight AS net_weight,a.chengse AS chengse,a.gold_price AS gold_price,a.settlement AS settlement,a.remarks AS remarks FROM xipunum_erp_outbound AS a WHERE a.order_id='{订单ID}'

-- 更新外部单号
UPDATE xipunum_erp_outbound_order SET waibu_number='{外部单号}' WHERE settlement_number='{单号}' LIMIT 1
```

**涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_member`、`xipunum_erp_type`

---

#### 7.6.4 商品销售订单查询 (窗口_商品销售订单查询)
**功能**: 销售订单查询，支持批量多选

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `副编码_超级列表框组件句柄` | 整数型 | 副编码列表框组件句柄 |

**子程序列表（5个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品销售订单查询_创建完毕` | 窗口初始化：加载销售订单列表 |
| `_副编码_超级列表框EX_项目左键单击` | 选中销售订单事件 |
| `_窗口_商品销售订单查询_尺寸被改变` | 窗口自适应布局 |
| `_按钮EX_批量多选_鼠标左键单击` | 批量多选销售订单 |

**核心SQL**:
```sql
-- 查询销售订单（按商品编码）
SELECT b.poduct_code as bianma,b.fu_code as fubianma,c.settlement_number as danhao,CAST(ROUND(a.quantity, 2) AS DECIMAL(10,2)) as quantity,CAST(ROUND(a.net_weight, 3) AS DECIMAL(10,3)) as jinzhong,CAST(ROUND(a.settlement, 2) AS DECIMAL(10,2)) as jine,a.creationtime as creationtime,c.id as danjuid FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_outbound_order AS c ON c.id = a.order_id WHERE (a.poduct_code='{编码}' OR b.fu_code='{编码}') AND a.kufang='{库房}' AND a.xsstate='0' ORDER BY a.creationtime DESC

-- 查询销售订单（按销售单号）
SELECT b.poduct_code as bianma,b.fu_code as fubianma,c.settlement_number as danhao,... FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_outbound_order AS c ON c.id = a.order_id WHERE c.settlement_number='{单号}' AND a.kufang='{库房}' AND a.xsstate='0' ORDER BY a.creationtime DESC
```

**涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop`
| `_按钮_查询_被单击` | 执行查询 |
| `_按钮_重置_被单击` | 重置查询条件 |

**涉及表**: `xipunum_erp_outbound_order`、`xipunum_erp_outbound`、`xipunum_erp_user`

---

#### 7.6.1.1 销售编辑批量修改 (窗口_销售编辑批量修改)
**功能**: 批量修改销售编辑信息

**子程序列表（8个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_销售编辑批量修改_创建完毕` | 窗口初始化：加载导购员列表 |
| `_窗口_销售编辑批量修改_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX_重置_鼠标左键单击` | 重置表单 |
| `_编辑框_附加费折扣_内容被改变` | 附加费折扣输入验证 |
| `_编辑框_销售附加费_内容被改变` | 销售附加费输入验证 |
| `_按钮EX_保存_鼠标左键单击` | 保存批量修改 |

**核心SQL**:
```sql
-- 加载导购员列表（全部权限）
SELECT name FROM xipunum_erp_user where 1=1 order by id ASC

-- 加载导购员列表（店铺权限）
SELECT name FROM xipunum_erp_user where department='{分组ID}' order by id ASC

-- 批量更新销售单
UPDATE xipunum_erp_outbound SET gold_price='{克价}',sales_cost='{销售工费}',sales_surcharge='{销售附加费}',shopping_guide='{导购员}' WHERE order_id='{订单ID}'
```

**涉及表**: `xipunum_erp_outbound`、`xipunum_erp_user`

---

#### 7.6.5 成品销售会员绑定 (窗口_成品销售会员绑定)
**功能**: 为销售订单绑定/修改会员

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `选中会员信息内容` | 文本型 | 选中的会员编码 |

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_成品销售会员绑定_创建完毕` | 初始化：加载会员搜索条件 |
| `_窗口_成品销售会员绑定_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_高级表格1_加载表头` | 加载会员列表表头 |
| `_高级表格1_加载数据` | 加载会员列表数据 |
| `_高级表格1_光标位置改变` | 选中会员事件 |
| `_按钮_查询_被单击` | 会员搜索（编码/ID/姓名/电话模糊匹配） |
| `_按钮_重置_被单击` | 重置查询条件 |
| `_按钮_绑定_被单击` | 绑定会员到订单 |
| `_按钮_解绑_被单击` | 解除会员绑定 |

**核心SQL**:
```sql
-- 会员搜索
SELECT a.customer_code,a.memberid,a.name,a.tel,a.shengri,a.dizhi
FROM xipunum_erp_member AS a
WHERE a.customer_code LIKE '%{关键字}%'
   OR a.memberid LIKE '%{关键字}%'
   OR a.NAME LIKE '%{关键字}%'
   OR a.tel LIKE '%{关键字}%'
ORDER BY a.creationtime DESC

-- 更新订单会员
UPDATE xipunum_erp_outbound_order SET customer_code='{会员编码}' WHERE id='{订单ID}' LIMIT 1

-- 批发订单存欠记录同步
UPDATE xipunum_erp_member_cq SET customer_code='{新会员编码}' WHERE customer_code='{旧会员编码}' AND kufang='{库房}'
```

**涉及表**: `xipunum_erp_outbound_order`、`xipunum_erp_member`、`xipunum_erp_member_cq`

---

### 7.7 库存管理模块

#### 7.7.1 实时库存查询 (窗口_实时库存查询)
**功能**: 查询当前实时库存数据，支持按店铺/品类/规格/材质多维度筛选

**程序集变量（11个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `店铺名称_超级列表框组件句柄` | 整数型 | 店铺筛选组件句柄 |
| `材质名称_超级列表框组件句柄` | 整数型 | 材质筛选组件句柄 |
| `报表数据_高级表格组件句柄` | 整数型 | 报表表格组件句柄 |
| `品类名称_超级列表框组件句柄` | 整数型 | 品类筛选组件句柄 |
| `规格名称_超级列表框组件句柄` | 整数型 | 规格筛选组件句柄 |
| `规格名称id_超级列表框组件句柄` | 整数型 | 规格ID筛选组件句柄 |
| `查找信息库房` | 文本型 | 店铺筛选SQL条件 |
| `查找信息规格` | 文本型 | 规格筛选SQL条件 |
| `查找信息材质` | 文本型 | 材质筛选SQL条件 |

**子程序列表（20个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_实时库存查询_创建完毕` | 窗口初始化：加载店铺/品类/规格/材质筛选框、加载表头、执行查询 |
| `_子程序_加载表头` | 加载报表表格表头（10列：序号/店铺名称/品类/规格/库存数量/库存金重/库存重量/成本工费/成本附加费/成本价） |
| `_查询条件_初始化` | 初始化查询条件：加载店铺/品类/规格/材质复选列表 |
| `_实时库存查询_表头回调` | 表头全选/取消全选回调 |
| `_店铺名称_超级列表框EX_项目左键单击` | 店铺筛选项点击事件 |
| `_品类名称_超级列表框EX_项目左键单击` | 品类筛选项点击事件 |
| `_规格名称_超级列表框EX_项目左键单击` | 规格筛选项点击事件 |
| `_材质名称_超级列表框EX_项目左键单击` | 材质筛选项点击事件 |
| `_品类选中_id数据初始化` | 品类选中后联动规格筛选 |
| `_选中规格id数据初始化` | 构建规格ID筛选条件 |
| `_超级按钮EX_重置_鼠标左键单击` | 重置查询条件 |
| `_按钮EX_导出数据_鼠标左键单击` | 导出Excel |
| `_超级按钮EX_查询_鼠标左键单击` | 执行查询：构建筛选条件、加载报表数据 |
| `_报表数据详情_展示` | 加载报表数据详情：按店铺→品类→规格层级展示 |

**核心SQL**:
```sql
-- 加载店铺下拉框
SELECT id AS akufang,CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in ({权限}) UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({权限}) ORDER BY akufang = '0' DESC, akufang

-- 加载品类下拉框（含未匹配选项）
SELECT id,title FROM xipunum_erp_category WHERE 1=1 UNION ALL SELECT '0' AS id, '未匹配' AS title FROM dual ORDER BY CASE WHEN id = '0' THEN 0 ELSE 1 END,CAST(id AS UNSIGNED) asc

-- 加载规格下拉框（含未匹配选项）
SELECT id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title UNION ALL SELECT '0' AS id, '未匹配' AS title FROM dual ORDER BY CASE WHEN id = '0' THEN 0 ELSE 1 END,CAST(id AS UNSIGNED) asc

-- 加载材质下拉框
SELECT caizhi FROM xipunum_erp_shop WHERE 1=1 GROUP BY caizhi ORDER BY id ASC

-- 按品类筛选规格
SELECT guige.category_id,guige.id,guige.title FROM (SELECT category_id,id,title FROM xipunum_erp_specs WHERE 1=1 UNION ALL SELECT '0','0','未匹配' FROM dual) as guige WHERE guige.category_id in ({品类ID列表}) GROUP BY guige.title ORDER BY guige.id asc

-- 实时库存查询（按店铺汇总）
SELECT b.kufang,CASE WHEN b.kufang = '0' THEN '总库' ELSE c.title END AS kufang_name,SUM(b.quantity) AS quantity,SUM(b.jinzhong) AS jinzhong,SUM(b.zhongliang) AS zhongliang FROM xipunum_erp_shop_kucun AS b INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code LEFT JOIN xipunum_erp_type AS c ON c.id = b.kufang LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = a.item_number AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) WHERE 1=1 {店铺条件} {品类条件} {规格条件} {材质条件} GROUP BY b.kufang ORDER BY b.kufang

-- 实时库存查询（按品类汇总）
SELECT f.id AS pinleiid,CASE WHEN COALESCE(f.title, '') = '' THEN '未匹配' ELSE f.title END AS pinlei,SUM(b.quantity) AS quantity,SUM(b.jinzhong) AS jinzhong,SUM(b.zhongliang) AS zhongliang FROM xipunum_erp_shop_kucun AS b INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = a.item_number LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) WHERE b.kufang='{店铺ID}' {品类条件} {规格条件} {材质条件} GROUP BY pinleiid ORDER BY pinleiid

-- 实时库存查询（按规格汇总）
SELECT COALESCE(e1.title, e2.title, '未匹配') AS guige,SUM(b.quantity) AS quantity,SUM(b.jinzhong) AS jinzhong,SUM(b.zhongliang) AS zhongliang FROM xipunum_erp_shop_kucun AS b INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = a.item_number LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id WHERE b.kufang='{店铺ID}' AND COALESCE(e1.category_id, e2.category_id)='{品类ID}' {规格条件} {材质条件} GROUP BY guige ORDER BY guige
```

**涉及表**: `xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_type`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_ksiamges`

---

#### 7.7.2 历史库存数据 (窗口_历史库存数据)
**功能**: 查询历史库存快照数据，支持按截止日期/店铺/品类/规格/工厂多维度筛选

**程序集变量（12个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `店铺名称_超级列表框组件句柄` | 整数型 | 店铺筛选组件句柄 |
| `品类名称_超级列表框组件句柄` | 整数型 | 品类筛选组件句柄 |
| `规格名称_超级列表框组件句柄` | 整数型 | 规格筛选组件句柄 |
| `工厂名称_超级列表框组件句柄` | 整数型 | 工厂筛选组件句柄 |
| `规格名称id_超级列表框组件句柄` | 整数型 | 规格ID筛选组件句柄 |
| `报表数据_高级表格组件句柄` | 整数型 | 报表表格组件句柄 |
| `局部_查找类型文本` | 文本型 | 查找类型文本 |
| `查找结束日期` | 文本型 | 截止日期 |
| `查找信息库房` | 文本型 | 店铺筛选SQL条件 |
| `查找信息规格` | 文本型 | 规格筛选SQL条件 |
| `查找信息工厂` | 文本型 | 工厂筛选SQL条件 |

**子程序列表（20个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_历史库存数据_创建完毕` | 窗口初始化：加载店铺/品类/规格/工厂筛选框、加载表头 |
| `_子程序_加载表头` | 加载报表表格表头（10列：序号/店铺名称/品类/规格/库存数量/库存金重/库存重量/成本工费/成本附加费/成本价） |
| `_编辑框_起始时间_鼠标左键单击` | 截止日期点击事件：弹出日期选择器 |
| `_查询条件_初始化` | 初始化查询条件：加载店铺/品类/规格/工厂复选列表 |
| `_历史库存数据_表头回调` | 表头全选/取消全选回调 |
| `_店铺名称_超级列表框EX_项目左键单击` | 店铺筛选项点击事件 |
| `_品类名称_超级列表框EX_项目左键单击` | 品类筛选项点击事件 |
| `_规格名称_超级列表框EX_项目左键单击` | 规格筛选项点击事件 |
| `_工厂名称_超级列表框EX_项目左键单击` | 工厂筛选项点击事件 |
| `_品类选中_id数据初始化` | 品类选中后联动规格筛选 |
| `_选中规格id数据初始化` | 构建规格ID筛选条件 |
| `_超级按钮EX_重置_鼠标左键单击` | 重置查询条件 |
| `_按钮EX_导出数据_鼠标左键单击` | 导出Excel |
| `_编辑框_包含信息_鼠标进入离开` | 搜索框提示文本处理 |
| `_编辑框_包含信息_键盘事件` | 搜索框键盘事件 |
| `_工厂查找_数据初始化` | 工厂搜索初始化 |
| `_超级按钮EX_查询_鼠标左键单击` | 执行查询：构建筛选条件、加载报表数据 |
| `_报表数据详情_展示` | 加载报表数据详情 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |
| `_历史库存数据_格式化数值` | 数值格式化（保留3位小数） |

**核心SQL**:
```sql
-- 加载店铺下拉框
SELECT id AS akufang,CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in ({权限}) UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({权限}) ORDER BY akufang = '0' DESC, akufang

-- 加载品类下拉框（含未匹配选项）
SELECT id,title FROM xipunum_erp_category WHERE 1=1 UNION ALL SELECT '0' AS id, '未匹配' AS title FROM dual ORDER BY CASE WHEN id = '0' THEN 0 ELSE 1 END,CAST(id AS UNSIGNED) asc

-- 加载规格下拉框（含未匹配选项）
SELECT id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title UNION ALL SELECT '0' AS id, '未匹配' AS title FROM dual ORDER BY CASE WHEN id = '0' THEN 0 ELSE 1 END,CAST(id AS UNSIGNED) asc

-- 历史库存查询（按店铺汇总）
SELECT b.kufang,CASE WHEN b.kufang = '0' THEN '总库' ELSE c.title END AS kufang_name,SUM(b.quantity) AS quantity,SUM(b.jinzhong) AS jinzhong,SUM(b.zhongliang) AS zhongliang FROM xipunum_erp_shop_kucun AS b INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code LEFT JOIN xipunum_erp_type AS c ON c.id = b.kufang LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = a.item_number LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) LEFT JOIN xipunum_erp_store AS g ON g.poduct_code = b.poduct_code LEFT JOIN xipunum_erp_store_order AS h ON h.id = g.order_id LEFT JOIN xipunum_erp_about AS i ON i.id = h.factory WHERE b.creationtime <= '{截止日期}' {店铺条件} {品类条件} {规格条件} {工厂条件} GROUP BY b.kufang ORDER BY b.kufang
```

**涉及表**: `xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_type`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_ksiamges`、`xipunum_erp_store`、`xipunum_erp_store_order`、`xipunum_erp_about`

---

#### 7.7.3 历史库存明细 (窗口_历史库存明细)
**功能**: 查询历史库存变动明细，支持按截止日期/店铺/品类/规格/工厂/金重/总重/圈号/款号/名称多维度筛选

**程序集变量（18个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `店铺名称_超级列表框组件句柄` | 整数型 | 店铺筛选组件句柄 |
| `品类名称_超级列表框组件句柄` | 整数型 | 品类筛选组件句柄 |
| `规格名称_超级列表框组件句柄` | 整数型 | 规格筛选组件句柄 |
| `工厂名称_超级列表框组件句柄` | 整数型 | 工厂筛选组件句柄 |
| `规格名称id_超级列表框组件句柄` | 整数型 | 规格ID筛选组件句柄 |
| `报表数据_高级表格组件句柄` | 整数型 | 报表表格组件句柄 |
| `局部_查找类型文本` | 文本型 | 查找类型文本 |
| `查找结束日期` | 文本型 | 截止日期 |
| `查找信息库房` | 文本型 | 店铺筛选SQL条件 |
| `查找信息规格` | 文本型 | 规格筛选SQL条件 |
| `查找信息工厂` | 文本型 | 工厂筛选SQL条件 |
| `查找信息金重` | 文本型 | 金重筛选SQL条件 |
| `查找信息总重` | 文本型 | 总重筛选SQL条件 |
| `查找信息圈号` | 文本型 | 圈号筛选SQL条件 |
| `查找信息款号` | 文本型 | 款号筛选SQL条件 |
| `查找信息名称` | 文本型 | 名称筛选SQL条件 |

**子程序列表（25个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_历史库存明细_创建完毕` | 窗口初始化：加载店铺/品类/规格/工厂筛选框、加载表头 |
| `_子程序_加载表头` | 加载报表表格表头（24列：序号/商品编码/原始编码/入库时间/品类名称/规格/材质/圈口/商品名称/款号/是否镶嵌/工厂/单件重/库存/金重/总重/库房/成本价/预售价/成本工费/成本附加费/参考工费/销售工费/销售附加费） |
| `_编辑框_起始时间_鼠标左键单击` | 截止日期点击事件：弹出日期选择器 |
| `_查询条件_初始化` | 初始化查询条件：加载店铺/品类/规格/工厂复选列表 |
| `_历史库存明细_表头回调` | 表头全选/取消全选回调 |
| `_店铺名称_超级列表框EX_项目左键单击` | 店铺筛选项点击事件 |
| `_品类名称_超级列表框EX_项目左键单击` | 品类筛选项点击事件 |
| `_规格名称_超级列表框EX_项目左键单击` | 规格筛选项点击事件 |
| `_工厂名称_超级列表框EX_项目左键单击` | 工厂筛选项点击事件 |
| `_品类选中_id数据初始化` | 品类选中后联动规格筛选 |
| `_选中规格id数据初始化` | 构建规格ID筛选条件 |
| `_超级按钮EX_重置_鼠标左键单击` | 重置查询条件 |
| `_按钮EX_导出数据_鼠标左键单击` | 导出Excel |
| `_编辑框_包含信息_鼠标进入离开` | 搜索框提示文本处理 |
| `_编辑框_包含信息_键盘事件` | 搜索框键盘事件 |
| `_编辑框_查找信息_鼠标进入离开` | 款号/名称搜索框提示文本处理 |
| `_编辑框_查找信息_键盘事件` | 款号/名称搜索框键盘事件 |
| `_工厂查找_数据初始化` | 工厂搜索初始化 |
| `_单选框_全部_选中状态改变` | 切换全部款号筛选 |
| `_单选框_为空_选中状态改变` | 切换空款号筛选 |
| `_单选框_不为空_选中状态改变` | 切换非空款号筛选 |
| `_超级按钮EX_查询_鼠标左键单击` | 执行查询：构建筛选条件、加载报表数据 |
| `_报表数据详情_展示` | 加载报表数据详情 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |
| `_历史库存明细_格式化数值` | 数值格式化（保留3位小数） |

**核心SQL**:
```sql
-- 加载店铺下拉框
SELECT id AS akufang,CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in ({权限}) UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({权限}) ORDER BY akufang = '0' DESC, akufang

-- 加载品类下拉框（含未匹配选项）
SELECT id,title FROM xipunum_erp_category WHERE 1=1 UNION ALL SELECT '0' AS id, '未匹配' AS title FROM dual ORDER BY CASE WHEN id = '0' THEN 0 ELSE 1 END,CAST(id AS UNSIGNED) asc

-- 加载规格下拉框（含未匹配选项）
SELECT id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title UNION ALL SELECT '0' AS id, '未匹配' AS title FROM dual ORDER BY CASE WHEN id = '0' THEN 0 ELSE 1 END,CAST(id AS UNSIGNED) asc

-- 历史库存明细查询
SELECT b.poduct_code AS bianma,a.fu_code AS afu_code,c.creationtime AS rukutime,CASE WHEN COALESCE(d.title, '') = '' THEN '未匹配' ELSE d.title END AS pinlei,COALESCE(e1.title, e2.title, '未匹配') AS guige,a.caizhi AS caizhi,a.quandu AS quankou,a.product_name AS mingcheng,a.item_number AS kuanhao,a.xiangqian AS xiangqian,CASE WHEN f.factory = '' THEN '' ELSE g.title END AS gongchang,CAST(ROUND(a.single, 3) AS DECIMAL(10,3)) AS danjian,b.quantity AS kucun,b.jinzhong AS jinzhong,CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN b.jinzhong ELSE CAST(ROUND(a.weight/a.quantity*b.quantity, 3) AS DECIMAL(10,3)) END AS zongzhong,CASE WHEN b.kufang = '0' THEN '总库' ELSE h.title END AS kufang,CAST(ROUND(c.cost_price, 2) AS DECIMAL(20,2)) AS chengben,CAST(ROUND(c.sales_price, 2) AS DECIMAL(20,2)) AS yushou,CAST(ROUND(c.basic_cost, 2) AS DECIMAL(20,2)) AS chengbengf,CAST(ROUND(c.company_surcharge, 2) AS DECIMAL(20,2)) AS chengbenfj,CAST(ROUND(c.premium_cost, 2) AS DECIMAL(20,2)) AS cankaogf,CAST(ROUND(c.sales_cost, 2) AS DECIMAL(20,2)) AS xiaoshougf,CAST(ROUND(c.sales_surcharge, 2) AS DECIMAL(20,2)) AS xiaoshoufj FROM xipunum_erp_shop_kucun AS b INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code INNER JOIN xipunum_erp_store AS c ON c.poduct_code = b.poduct_code INNER JOIN xipunum_erp_store_order AS f ON f.id = c.order_id LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang LEFT JOIN xipunum_erp_about AS g ON g.id = f.factory LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = a.item_number AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id AND a.item_number IS NOT NULL AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id AND a.specification_id IS NOT NULL AND a.specification_id != '' LEFT JOIN xipunum_erp_category AS d ON d.id = COALESCE(e1.category_id, e2.category_id) AND COALESCE(e1.category_id, e2.category_id) IS NOT NULL WHERE b.creationtime <= '{截止日期}' {店铺条件} {品类条件} {规格条件} {工厂条件} {金重条件} {总重条件} {圈号条件} {款号条件} {名称条件} ORDER BY b.creationtime DESC
```

**涉及表**: `xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_type`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_ksiamges`、`xipunum_erp_store`、`xipunum_erp_store_order`、`xipunum_erp_about`

---

#### 7.7.4 历史追溯 (窗口_历史追溯)
**功能**: 商品历史追溯查询，查看单个商品的完整信息和操作历史

**子程序列表（10个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_历史追溯_创建完毕` | 窗口初始化：清空输入框、初始化数据 |
| `_编辑框_商品编码_键盘事件` | 商品编码输入框回车键处理 |
| `_按钮EX_查找_鼠标左键单击` | 查找按钮事件：验证商品编码→加载商品详情和历史 |
| `_按钮EX2_鼠标左键单击` | 复制款号按钮事件 |
| `子_初始化` | 加载商品详情和操作历史 |
| `_超级列表框EX1_项目左键单击` | 历史记录点击事件 |
| `_按钮EX_导出数据_鼠标左键单击` | 导出Excel |
| `_主窗口_弹出日期框` | 弹出日期选择器 |

**核心SQL**:
```sql
-- 验证商品编码存在性
SELECT * FROM xipunum_erp_shop where poduct_code='{编码}'

-- 加载商品详情
SELECT a.poduct_code AS poduct_code,a.product_name as product_name,CASE WHEN COALESCE(f.title, '') = '' THEN '未匹配' ELSE f.title END AS pinlei,a.caizhi as caizhi,b.company_condition as chengse,COALESCE(e1.title, e2.title, '无数据') AS guige,a.sales_unit as danwei,c.gold_price as kejia,a.fu_code as fubian,a.single as danjian,a.quantity as shuliang,a.weight as zongzhong,a.jin_zhong as jinzhong,a.loss as sunhao,a.including as hanhaozhong,a.quandu as quankou,a.wide as miankuan,a.thickness as houdu,b.cost_price as chengben,a.item_number as kuanhao,g.shitou as shizhong,g.stnum as shishu,g.shitou1 as fushizhong,g.shnum1 as fushishu,g.zhuse as zhuse,b.coefficient as xishu,b.basic_cost as cbgongfei,b.company_surcharge as cbfujia,b.sales_cost as xsgongfei,b.premium_cost as cankao,b.sales_surcharge as xsfujia,b.sales_price as yushou,a.remarks as beizhu FROM xipunum_erp_shop AS a INNER JOIN xipunum_erp_store AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_store_order AS c ON c.id = b.order_id LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = a.item_number AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id AND a.item_number IS NOT NULL AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id AND a.specification_id IS NOT NULL AND a.specification_id != '' LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) AND COALESCE(e1.category_id, e2.category_id) IS NOT NULL LEFT JOIN xipunum_erp_shop_xiangqian AS g ON g.poduct_code = a.poduct_code WHERE a.poduct_code = '{编码}' ORDER BY a.id DESC LIMIT 1

-- 加载商品操作历史
SELECT a.creationtime AS shijian,a.danhao AS danhao,a.type AS leixing,a.quantity AS shuliang,a.jinzhong AS jinzhong,a.zhongliang AS zhongliang,a.content AS neirong,a.user AS caozuoyuan FROM xipunum_erp_shop_log AS a WHERE a.poduct_code = '{编码}' ORDER BY a.id DESC

-- 加载库存信息
SELECT a.quantity AS shuliang,a.jinzhong AS jinzhong,b.title AS kufang_name FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_type AS b ON b.id = a.kufang WHERE a.poduct_code = '{编码}' ORDER BY a.id DESC
```

**涉及表**: `xipunum_erp_shop`、`xipunum_erp_shop_log`、`xipunum_erp_shop_kucun`、`xipunum_erp_store`、`xipunum_erp_store_order`、`xipunum_erp_ksiamges`、`xipunum_erp_specs`、`xipunum_erp_category`、`xipunum_erp_shop_xiangqian`、`xipunum_erp_type`

---

### 7.8 会员管理模块

#### 7.8.1 会员添加修改 (窗口_会员添加修改)
**功能**: 会员信息新增/编辑，支持省市区三级联动地址选择

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `选中会员信息内容` | 文本型 | 选中的会员编码（空=新增，非空=编辑） |

**子程序列表（13个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_会员添加修改_创建完毕` | 窗口初始化：根据选中会员设置界面、加载省份下拉框 |
| `_组合框_省份_内容被改变` | 省份选择联动：加载城市下拉框 |
| `_组合框_市区_内容被改变` | 城市选择联动：加载区县下拉框 |
| `_组合框_县区_内容被改变` | 区县选择联动：更新联系地址 |
| `_编辑框_出生日期_鼠标左键单击` | 出生日期点击事件：弹出日期选择器 |
| `_窗口_会员添加修改_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX1_鼠标左键单击` | 保存/修改按钮事件 |
| `_按钮EX2_鼠标左键单击` | 重置按钮事件 |
| `_会员添加_保存数据` | 保存会员数据：验证电话→检查重复→添加会员→更新memberid→处理存欠 |
| `_会员修改_修改数据` | 修改会员数据：更新会员信息→处理存欠→更新主窗口表格 |
| `_会员信息_存欠数据` | 处理会员存欠数据：添加存欠记录 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |

**核心SQL**:
```sql
-- 加载省份下拉框
SELECT id,name FROM erp_city where pid='0' and level='1' order by id ASC

-- 加载城市下拉框
SELECT id,name FROM erp_city where pid='{省份ID}' and level='2' order by id ASC

-- 加载区县下拉框
SELECT id,name FROM erp_city where pid='{城市ID}' and level='3' order by id ASC

-- 检查会员是否存在
SELECT * FROM xipunum_erp_member where name='{姓名}' and tel='{电话}'

-- 添加会员
增加记录("xipunum_erp_member", customer_code, name, tel, shengri, dizhi, cjuser, creationtime)

-- 获取会员ID
SELECT id,customer_code FROM xipunum_erp_member where customer_code='{编码}' order by id ASC LIMIT 1

-- 更新会员memberid
UPDATE xipunum_erp_member SET memberid='{ID}' WHERE customer_code='{编码}' LIMIT 1

-- 修改会员信息
UPDATE xipunum_erp_member SET name='{姓名}',shengri='{生日}',dizhi='{地址}',cjuser='{账户}',updatetime='{时间}' WHERE customer_code='{编码}' LIMIT 1

-- 查询存欠余额
SELECT a.memberid as amemberid,a.customer_code AS acustomer_code,a.NAME AS aname,a.tel AS atel,ROUND(IFNULL(c.cun_number, 0) - IFNULL(c.qian_number, 0), 3) AS jieyuliao,ROUND(IFNULL(d.cun_number, 0) - IFNULL(d.qian_number, 0), 2) AS jieyuyuan,a.shengri AS ashengri,a.dizhi AS adizhi,a.cjuser AS acjuser,a.creationtime AS acreationtime,b.NAME AS bname FROM xipunum_erp_member AS a INNER JOIN xipunum_erp_user AS b ON b.USER = a.cjuser LEFT JOIN (SELECT customer_code, SUM(CASE WHEN cunqu = '存' AND type = '料' THEN number ELSE 0 END) AS cun_number, SUM(CASE WHEN cunqu = '欠' AND type = '料' THEN number ELSE 0 END) AS qian_number FROM xipunum_erp_member_cq WHERE kufang in ({权限}) AND type = '料' GROUP BY customer_code) AS c ON c.customer_code = a.customer_code LEFT JOIN (SELECT customer_code, SUM(CASE WHEN cunqu = '存' AND type = '元' THEN number ELSE 0 END) AS cun_number, SUM(CASE WHEN cunqu = '欠' AND type = '元' THEN number ELSE 0 END) AS qian_number FROM xipunum_erp_member_cq WHERE kufang in ({权限}) AND type = '元' GROUP BY customer_code) AS d ON d.customer_code = a.customer_code WHERE a.customer_code='{编码}'

-- 添加存欠记录
增加记录("xipunum_erp_member_cq", customer_code, cunqu, type, number, remarks, kufang, cjuser, creationtime)

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `erp_city`、`xipunum_erp_member`、`xipunum_erp_member_cq`、`xipunum_erp_user`、`xipunum_erp_xitong_log`

#### 7.8.2 会员信息合并 (窗口_会员信息合并)
**功能**: 将多个会员数据合并到一个主会员

**程序集变量（3个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_行号` | 整数型 | 表格当前行号 |
| `集_列号` | 整数型 | 表格当前列号 |
| `删除按钮` | 表格按钮 | 表格删除按钮 |

**子程序列表（12个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_会员信息合并_创建完毕` | 窗口初始化：加载表格表头 |
| `_窗口_会员信息合并_尺寸被改变` | 窗口自适应布局 |
| `子程序_删除表格` | 清空表格 |
| `高级表格1_加载表头` | 加载表格表头（13列：序号/客户编码/会员ID/姓名/电话/结料/结款/出生年月/联系地址/创建时间/主会员/合并会员/操作） |
| `_高级表格1_光标位置改变` | 表格光标位置改变事件：主会员/合并会员单选逻辑 |
| `_高级表格1_按钮被点击` | 表格按钮点击事件：删除会员 |
| `_按钮_查询_被单击` | 查询按钮事件：搜索会员 |
| `_按钮_重置_被单击` | 重置按钮事件 |
| `_查找信息编辑框_按下某键` | 搜索框回车键处理 |
| `_高级表格1_合并数据` | 加载会员搜索结果数据 |
| `_按钮_合并_被单击` | 合并按钮事件：执行合并操作 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |

**核心SQL**:
```sql
-- 会员搜索（支持编码/姓名/电话/地址模糊匹配）
SELECT a.memberid as amemberid,a.customer_code AS acustomer_code,a.NAME AS aname,a.tel AS atel,ROUND(IFNULL(c.cun_number, 0) - IFNULL(c.qian_number, 0), 3) AS jieyuliao,ROUND(IFNULL(d.cun_number, 0) - IFNULL(d.qian_number, 0), 2) AS jieyuyuan,a.shengri AS ashengri,a.dizhi AS adizhi,a.cjuser AS acjuser,a.creationtime AS acreationtime,b.NAME AS bname FROM xipunum_erp_member AS a INNER JOIN xipunum_erp_user AS b ON b.USER = a.cjuser LEFT JOIN (SELECT customer_code, SUM(CASE WHEN cunqu = '存' AND type = '料' THEN number ELSE 0 END) AS cun_number, SUM(CASE WHEN cunqu = '欠' AND type = '料' THEN number ELSE 0 END) AS qian_number FROM xipunum_erp_member_cq WHERE kufang in ({权限}) AND type = '料' GROUP BY customer_code) AS c ON c.customer_code = a.customer_code LEFT JOIN (SELECT customer_code, SUM(CASE WHEN cunqu = '存' AND type = '元' THEN number ELSE 0 END) AS cun_number, SUM(CASE WHEN cunqu = '欠' AND type = '元' THEN number ELSE 0 END) AS qian_number FROM xipunum_erp_member_cq WHERE kufang in ({权限}) AND type = '元' GROUP BY customer_code) AS d ON d.customer_code = a.customer_code WHERE a.customer_code IN (SELECT DISTINCT customer_code FROM (SELECT customer_code FROM xipunum_erp_outbound_order WHERE cjuser IN {权限} UNION SELECT customer_code FROM xipunum_erp_retreat_order WHERE cjuser IN {权限} UNION SELECT customer_code FROM xipunum_erp_presale_order WHERE cjuser IN {权限} UNION SELECT customer_code FROM xipunum_erp_member WHERE cjuser IN {权限}) AS combined WHERE customer_code != '' AND customer_code IS NOT NULL) AND (a.customer_code LIKE '%{关键字}%' OR a.NAME LIKE '%{关键字}%' OR a.tel LIKE '%{关键字}%' OR a.dizhi LIKE '%{关键字}%') ORDER BY a.creationtime asc

-- 合并积分记录
UPDATE xipunum_erp_member_score_log SET customer_code='{主会员编码}' WHERE id='{记录ID}'

-- 合并存取记录
UPDATE xipunum_erp_member_cq SET customer_code='{主会员编码}',updatetime='{时间}' WHERE id='{记录ID}'

-- 合并预售订单
UPDATE xipunum_erp_presale_order SET customer_code='{主会员编码}',updatetime='{时间}' WHERE id='{订单ID}'

-- 合并回收订单
UPDATE xipunum_erp_retreat_order SET customer_code='{主会员编码}',updatetime='{时间}' WHERE id='{订单ID}'

-- 合并销售订单
UPDATE xipunum_erp_outbound_order SET customer_code='{主会员编码}',updatetime='{时间}' WHERE id='{订单ID}'

-- 合并线上订单
UPDATE xipunum_erp_online_order SET customer_code='{主会员编码}' WHERE id='{订单ID}'

-- 删除被合并会员
delete from xipunum_erp_member where customer_code='{被合并编码}' and memberid='{被合并ID}'

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_member`、`xipunum_erp_member_cq`、`xipunum_erp_member_score_log`、`xipunum_erp_presale_order`、`xipunum_erp_retreat_order`、`xipunum_erp_outbound_order`、`xipunum_erp_online_order`、`xipunum_erp_user`、`xipunum_erp_xitong_log`

#### 7.8.3 会员列表排序 (窗口_会员列表排序)
**功能**: 设置会员列表排序方式

**子程序列表（5个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_会员列表排序_创建完毕` | 窗口初始化：根据当前排序设置选中状态 |
| `_窗口_会员列表排序_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX_关闭_鼠标左键单击` | 关闭窗口 |
| `_按钮EX1_鼠标左键单击` | 确认排序方式：设置排序字段和方向、刷新会员列表 |

**排序字段**:
| 字段名 | 说明 |
|--------|------|
| `id` | 按会员ID排序 |
| `CDbl(data5)` | 按结料排序 |
| `CDbl(data6)` | 按结款排序 |
| `data9` | 按创建时间排序 |
| `data10` | 按更新时间排序 |

**排序方向**: 升序(asc) / 降序(DESC)

---

#### 7.8.4 会员回访添加 (窗口_会员回访添加)
**功能**: 添加会员回访记录，支持电话搜索会员

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_订单是否选中` | 整数型 | 是否选中订单（-1=新建，其他=详情） |

**子程序列表（7个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_会员回访添加_创建完毕` | 窗口初始化：根据订单状态设置界面 |
| `_窗口_会员回访添加_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX2_鼠标左键单击` | 重置按钮事件 |
| `_编辑框_联系电话_键盘事件` | 联系电话输入框回车键处理：搜索会员信息 |
| `_按钮EX1_鼠标左键单击` | 保存按钮事件：验证数据→添加回访记录→添加系统日志 |

**核心SQL**:
```sql
-- 验证会员电话是否存在
SELECT * FROM xipunum_erp_member where tel='{电话}'

-- 根据电话查询会员信息
SELECT * FROM xipunum_erp_member where tel='{电话}'

-- 添加回访记录
增加记录("xipunum_erp_visit", customer_code, returntitle, returnconter, returndata, cjuser, creationtime)

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type='添加', title='添加回访信息', conter, user, creationtime)
```

**涉及表**: `xipunum_erp_member`、`xipunum_erp_visit`、`xipunum_erp_xitong_log`

---

#### 7.8.5 会员订单消费数据 (窗口_会员订单消费数据)
**功能**: 查看会员消费订单列表

**程序集变量（3个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_订单是否选中` | 整数型 | 选中的订单行号 |
| `局部_会员信息编码` | 文本型 | 会员编码 |
| `超级列表框EX组件句柄` | 整数型 | 列表框组件句柄 |

**子程序列表（4个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_会员订单消费数据_创建完毕` | 窗口初始化：加载会员编码、加载表头、加载消费数据 |
| `_超级列表框EX_加载表头` | 加载表格表头（14列：序号/出库单号/销售时间/预售单号/导购员/销售金额/实收金额/销售克价/结算净重/成本工费/参考工费/销售工费/销售附加费/折扣） |
| `_高级表格1_消费加载` | 加载消费订单数据 |

**核心SQL**:
```sql
-- 查询会员消费订单（全部权限）
SELECT a.settlement_number AS asettlement_number,a.presale_number AS apresale_number,b.xiao_amount AS bxiao_amount,b.settlement AS bsettlement,b.gold_price AS bgold_price,b.net_weight AS bnet_weight,b.basic_cost AS bbasic_cost,b.premium_cost AS bpremium_cost,b.sales_cost AS bsales_cost,b.sales_surcharge AS bsales_surcharge,b.zhekou AS bzhekou,b.creationtime AS bcreationtime,c.name as daogou FROM xipunum_erp_outbound_order AS a left JOIN xipunum_erp_outbound AS b ON b.order_id = a.id LEFT JOIN xipunum_erp_user AS c ON c.user = b.shopping_guide WHERE a.customer_code = '{会员编码}' ORDER BY a.id DESC

-- 查询会员消费订单（店铺权限）
SELECT a.settlement_number AS asettlement_number,a.presale_number AS apresale_number,b.xiao_amount AS bxiao_amount,b.settlement AS bsettlement,b.gold_price AS bgold_price,b.net_weight AS bnet_weight,b.basic_cost AS bbasic_cost,b.premium_cost AS bpremium_cost,b.sales_cost AS bsales_cost,b.sales_surcharge AS bsales_surcharge,b.zhekou AS bzhekou,b.creationtime AS bcreationtime,c.name as daogou FROM xipunum_erp_outbound_order AS a left JOIN xipunum_erp_outbound AS b ON b.order_id = a.id LEFT JOIN xipunum_erp_user AS c ON c.user = b.shopping_guide WHERE a.customer_code = '{会员编码}' and b.kufang='{库房ID}' ORDER BY a.id DESC
```

**涉及表**: `xipunum_erp_outbound_order`、`xipunum_erp_outbound`、`xipunum_erp_user`

---

#### 7.8.6 会员订单充值数据 (窗口_会员订单充值数据)
**功能**: 查看会员充值/存欠记录

**程序集变量（3个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_订单是否选中` | 整数型 | 选中的订单行号 |
| `局部_会员信息编码` | 文本型 | 会员编码 |
| `超级列表框EX组件句柄` | 整数型 | 列表框组件句柄 |

**子程序列表（3个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_会员订单充值数据_创建完毕` | 窗口初始化：加载会员编码、加载表头、加载充值数据 |
| `_超级列表框EX_加载表头` | 加载表格表头（7列：序号/充值类型/单位/数值/备注/操作账户/操作时间） |
| `_高级表格1_充值加载` | 加载充值记录数据 |

**核心SQL**:
```sql
-- 查询会员存欠记录（全部权限）
SELECT a.cunqu as acunqu,a.type as atype,CAST(ROUND(a.number, 3) AS DECIMAL(10,3)) AS csingle,a.remarks as aremarks,a.cjuser as acjuser,a.creationtime as acreationtime,b.name as bname FROM xipunum_erp_member_cq as a INNER JOIN xipunum_erp_user AS b ON b.user = a.cjuser WHERE a.customer_code = '{会员编码}' ORDER BY a.id DESC

-- 查询会员存欠记录（店铺权限）
SELECT a.cunqu as acunqu,a.type as atype,CAST(ROUND(a.number, 3) AS DECIMAL(10,3)) AS csingle,a.remarks as aremarks,a.cjuser as acjuser,a.creationtime as acreationtime,b.name as bname FROM xipunum_erp_member_cq as a INNER JOIN xipunum_erp_user AS b ON b.user = a.cjuser WHERE a.customer_code = '{会员编码}' and a.kufang='{库房ID}' ORDER BY a.id DESC
```

**涉及表**: `xipunum_erp_member_cq`、`xipunum_erp_user`

---

#### 7.8.7 会员订单预购数据 (窗口_会员订单预购数据)
**功能**: 查看会员预售订单列表

**程序集变量（3个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_订单是否选中` | 整数型 | 选中的订单行号 |
| `局部_会员信息编码` | 文本型 | 会员编码 |
| `超级列表框EX组件句柄` | 整数型 | 列表框组件句柄 |

**子程序列表（3个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_会员订单预购数据_创建完毕` | 窗口初始化：加载会员编码、加载表头、加载预购数据 |
| `_超级列表框EX_加载表头` | 加载表格表头（8列：序号/预售单号/商品名称/数量/备注/预购时间/状态/拿货时间） |
| `_高级表格1_预购加载` | 加载预购订单数据 |

**核心SQL**:
```sql
-- 查询会员预购订单
SELECT b.presale_umber AS bpresale_umber,a.product_name AS aproduct_name,CAST(ROUND(a.quantity, 3) AS DECIMAL(10,3)) AS aquantity,a.remarks AS aremarks,a.creationtime AS acreationtime,b.state AS bstate,b.creationtime AS bcreationtime FROM xipunum_erp_presale AS a INNER JOIN xipunum_erp_presale_order AS b ON b.presale_umber = a.order_id AND b.customer_code = '{会员编码}' and b.cjuser IN {权限} WHERE 1 = 1 ORDER BY a.id DESC
```

**涉及表**: `xipunum_erp_presale`、`xipunum_erp_presale_order`

---

### 7.9 财务管理模块

#### 7.9.1 收支管理信息 (窗口_收支管理信息)
**功能**: 收支记录管理

**程序集变量（3个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_图片路径` | 文本型 | 上传图片路径 |
| `局部_图片响应` | 文本型 | 图片上传响应 |
| `局部_收支管理名称` | 文本型 | 当前收支名称 |

**子程序列表**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_收支管理信息_创建完毕` | 初始化：加载收支名称/店铺/支付方式下拉框 |
| `_窗口_收支管理信息_尺寸被改变` | 窗口自适应布局 |
| `_高级表格1_加载表头` | 加载表格列头 |
| `_高级表格1_加载数据` | 加载收支记录列表 |
| `_高级表格1_光标位置改变` | 选中记录事件 |
| `_按钮_添加_被单击` | 添加收支记录 |
| `_按钮_修改_被单击` | 修改收支记录 |
| `_按钮_删除_被单击` | 删除收支记录 |
| `_按钮_查询_被单击` | 查询收支记录 |
| `_按钮_重置_被单击` | 重置查询条件 |

**核心SQL**:
```sql
-- 加载收支记录
SELECT a.id,a.type,a.amount,a.remarks,a.creationtime,
       b.title AS pay_name,c.title AS kufang_name
FROM xipunum_erp_finance AS a
LEFT JOIN xipunum_erp_pay AS b ON b.id = a.pay_type
LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang
WHERE a.kufang IN ({店铺权限}) {日期条件} {收支名称条件}
ORDER BY a.creationtime DESC

-- 添加收支记录
增加记录("xipunum_erp_finance", type, pay_type, kufang, amount, remarks, image, cjuser, creationtime)

-- 修改收支记录
UPDATE xipunum_erp_finance SET type='{类型}',pay_type='{支付方式}',amount='{金额}',remarks='{备注}',image='{图片}' WHERE id='{ID}' LIMIT 1

-- 删除收支记录
DELETE FROM xipunum_erp_finance WHERE id='{ID}' LIMIT 1
```

**涉及表**: `xipunum_erp_finance`、`xipunum_erp_pay`、`xipunum_erp_type`、`xipunum_erp_finance_title`

---

#### 7.9.2 收支名称管理 (窗口_收支名称管理)
**功能**: 收支项目名称配置（详见7.4.6）

---

#### 7.9.3 收支卡号管理 (窗口_收支卡号管理)
**功能**: 银行卡号管理（详见7.4.7）

---

#### 7.9.4 店铺数据结算 (窗口_店铺数据结算)
**功能**: 店铺结算操作，支持结料/结账两种模式

**程序集变量（3个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `局部_结算按钮名称` | 文本型 | 当前结算按钮名称（"添加"/"编辑"/"详情"） |
| `集_行号` | 整数型 | 表格当前行号 |
| `集_列号` | 整数型 | 表格当前列号 |

**子程序列表（15个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_店铺数据结算_创建完毕` | 窗口初始化：连接本地Access、加载表头、根据按钮名称设置界面 |
| `_窗口_店铺数据结算_尺寸被改变` | 窗口自适应布局 |
| `_高级表格1_加载表头` | 加载表格表头（8列：序号/店铺名称/结算克重/结算克价/结算金额/应结多结/备注/店铺id） |
| `子程序_删除表格` | 清空表格 |
| `_高级表格1_光标位置改变` | 表格光标位置改变事件 |
| `_工具条_通用_被单击` | 工具条按钮点击事件（保存/重置） |
| `_高级表格1_统计数据` | 统计合计数据（结算克重/结算金额/应结多结） |
| `_店铺结算详情_被点击` | 加载店铺结算详情数据 |
| `_单选框_结料_选中状态改变` | 切换结料模式 |
| `_单选框_结账_选中状态改变` | 切换结账模式 |
| `_超级按钮_保存_被单击` | 保存结算单据：事务处理（生成订单→保存结算明细→更新店铺存欠） |
| `_主窗口_弹出日期框` | 弹出日期选择器 |
| `_店铺数据结算_格式三位小数` | 格式化三位小数 |
| `_店铺数据结算_格式二位小数` | 格式化两位小数 |

**核心SQL**:
```sql
-- 加载结算材质下拉框
SELECT title FROM xipunum_erp_category_stat_config WHERE title like '%黄金%' ORDER BY id asc

-- 加载店铺结算详情
SELECT ... FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_type AS b ON b.id = a.kufang WHERE a.kufang IN ({权限}) {日期条件} GROUP BY a.kufang ORDER BY a.kufang ASC

-- 统计结算数据
SELECT SUM(quantity) AS aquantity,SUM(net_weight) AS anet_weight,SUM(settlement) AS asettlement FROM xipunum_erp_outbound WHERE kufang='{店铺ID}' {日期条件}

-- 保存结算订单
增加记录("xipunum_erp_settlement_order", settlement_number, kufang, type, total_weight, total_amount, remarks, cjuser, creationtime)

-- 保存结算明细
增加记录("xipunum_erp_settlement", order_id, kufang, weight, price, amount, remarks, creationtime)

-- 更新店铺存欠
增加记录("xipunum_erp_member_cq", customer_code, cunqu, type, number, remarks, kufang, cjuser, creationtime)

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_type`、`xipunum_erp_category_stat_config`、`xipunum_erp_settlement_order`、`xipunum_erp_settlement`、`xipunum_erp_member_cq`、`xipunum_erp_xitong_log`

---

#### 7.9.4.1 结账结料 (窗口_结账结料)
**功能**: 工厂结账结料管理

**程序集变量（3个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `集_行号` | 整数型 | 表格当前行号 |
| `集_列号` | 整数型 | 表格当前列号 |
| `局部_订单是否选中` | 整数型 | 是否选中订单（-1=新建，其他=详情） |

**子程序列表（15个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_结账结料_创建完毕` | 窗口初始化：加载工厂下拉框 |
| `_窗口_结账结料_尺寸被改变` | 窗口自适应布局 |
| `_高级表格1_加载表头` | 加载表格表头（10列：序号/店铺名称/材质/待结料/待结账/结料/结账/备注/库房id/品类id） |
| `子程序_删除表格` | 清空表格 |
| `_高级表格1_光标位置改变` | 表格光标位置改变事件 |
| `_高级表格1_将被编辑` | 表格开始编辑事件 |
| `_高级表格1_结束编辑` | 表格结束编辑事件 |
| `_结账结料详情_被点击` | 加载结账结料详情 |
| `_超级按钮_保存_被单击` | 保存结账结料单据 |
| `_超级按钮_重置_被单击` | 重置表单 |
| `_工具条_通用_被单击` | 工具条按钮点击事件 |

**核心SQL**:
```sql
-- 加载工厂下拉框
SELECT * FROM xipunum_erp_about where 1=1 order by id ASC

-- 查询结账结料详情
SELECT ... FROM xipunum_erp_delivery AS a INNER JOIN xipunum_erp_delivery_order AS b ON b.id = a.order_id WHERE b.delivery_umber='{单号}' ORDER BY a.id ASC

-- 保存结账结料订单
增加记录("xipunum_erp_delivery_order", delivery_umber, type, gongchang, remarks, cjuser, creationtime)

-- 保存结账结料明细
增加记录("xipunum_erp_delivery", order_id, gongchang, dianou_id, type, caizhi, dai_materials, dai_checkout, materials, checkout, remarks, cjuser, creationtime)
```

**涉及表**: `xipunum_erp_delivery`、`xipunum_erp_delivery_order`、`xipunum_erp_about`、`xipunum_erp_type`、`xipunum_erp_category_stat_config`

---

#### 7.9.5 绩效信息管理 (窗口_绩效信息管理)
**功能**: 员工绩效管理，支持多档次绩效计算

**程序集变量（4个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `品类名称_超级列表框组件句柄` | 整数型 | 品类名称列表框组件句柄 |
| `集_范围下拉列表句柄` | 整数型 | 范围下拉列表句柄 |
| `集_字段下拉列表句柄` | 整数型 | 字段下拉列表句柄 |
| `局部_超级列表框EX1选中` | 文本型 | 选中的绩效配置行号 |

**子程序列表（20个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_绩效信息管理_创建完毕` | 窗口初始化：加载岗位列表、品类下拉框、初始化表单 |
| `_窗口_绩效信息管理_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX_绩效管理重置_鼠标左键单击` | 重置表单 |
| `_超级列表框EX1_项目左键单击` | 岗位列表选中事件 |
| `_选中岗位数据_初始化` | 加载选中岗位的绩效配置 |
| `_组合框Ex计算档次_内容被改变` | 计算档次变更事件：显示/隐藏档次输入框 |
| `_单选框_店铺_选中状态改变` | 切换店铺计算范围 |
| `_单选框_岗位_选中状态改变` | 切换岗位计算范围 |
| `_单选框_个人_选中状态改变` | 切换个人计算范围 |
| `_单选框_批发_选中状态改变` | 切换批发批零类型 |
| `_单选框_零售_选中状态改变` | 切换零售批零类型 |
| `_单选框_总数量_选中状态改变` | 切换总数量计算字段 |
| `_单选框_总重量_选中状态改变` | 切换总重量计算字段 |
| `_单选框_总金额_选中状态改变` | 切换总金额计算字段 |
| `_单选框_单件重_选中状态改变` | 切换单件重计算字段 |
| `_按钮EX__绩效管理保存_鼠标左键单击` | 保存绩效配置：验证数据→检查重复→保存/修改配置 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |
| `_绩效信息管理_格式三位小数` | 格式化三位小数 |
| `_绩效信息管理_格式二位小数` | 格式化两位小数 |

**核心SQL**:
```sql
-- 加载品类下拉框
SELECT * FROM xipunum_erp_category WHERE 1=1 ORDER BY id asc

-- 加载岗位列表（全部权限）
SELECT a.type as atype,b.title as btitle,a.title as atitle,CASE WHEN a.data9 = '0' THEN '店铺' WHEN a.data9 = '1' THEN '岗位' WHEN a.data9 = '2' THEN '个人' ELSE '' END AS data9,a.id as aid FROM xipunum_erp_type as a INNER JOIN xipunum_erp_type AS b ON b.id = a.superior left JOIN xipunum_erp_role AS c ON c.typeid = a.id WHERE a.superior >'0' and a.type='商铺' ORDER BY a.id asc

-- 加载岗位列表（店铺权限）
SELECT a.type as atype,b.title as btitle,a.title as atitle,CASE WHEN a.data9 = '0' THEN '店铺' WHEN a.data9 = '1' THEN '岗位' WHEN a.data9 = '2' THEN '个人' ELSE '' END AS data9,a.id as aid FROM xipunum_erp_type as a INNER JOIN xipunum_erp_type AS b ON b.id = a.superior left JOIN xipunum_erp_role AS c ON c.typeid = a.id WHERE a.superior ='{分组ID}' and a.type='商铺' ORDER BY a.id asc

-- 加载岗位绩效配置
SELECT * FROM xipunum_erp_performance WHERE type_id='{岗位ID}' ORDER BY id ASC

-- 检查绩效配置是否存在
SELECT * FROM xipunum_erp_performance WHERE type_id='{岗位ID}' and category_id='{品类ID}' and pl='{批零}'

-- 保存绩效配置
增加记录("xipunum_erp_performance", type_id, category_id, pl, jszd, jsfw, djs, min1, cs1, min2, cs2, min3, cs3, min4, cs4, min5, cs5, cjuser, creationtime)

-- 修改绩效配置
UPDATE xipunum_erp_performance SET category_id='{品类ID}',pl='{批零}',jszd='{计算字段}',jsfw='{计算范围}',djs='{档次数}',min1='{最小值1}',cs1='{参数值1}',min2='{最小值2}',cs2='{参数值2}',min3='{最小值3}',cs3='{参数值3}',min4='{最小值4}',cs4='{参数值4}',min5='{最小值5}',cs5='{参数值5}',cjuser='{账户}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1

-- 删除绩效配置
DELETE FROM xipunum_erp_performance WHERE id='{ID}'

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_performance`、`xipunum_erp_category`、`xipunum_erp_type`、`xipunum_erp_role`、`xipunum_erp_xitong_log`

---

### 7.10 报表模块

#### 7.10.1 商品报表

**通用报表程序集变量**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `店铺名称_超级列表框组件句柄` | 整数型 | 店铺筛选组件句柄 |
| `品类名称_超级列表框组件句柄` | 整数型 | 品类筛选组件句柄 |
| `规格名称_超级列表框组件句柄` | 整数型 | 规格筛选组件句柄 |
| `报表数据_高级表格组件句柄` | 整数型 | 报表表格组件句柄 |
| `查找开始日期` | 文本型 | 查询开始日期 |
| `查找结束日期` | 文本型 | 查询结束日期 |
| `查找信息库房` | 文本型 | 店铺筛选SQL条件 |
| `查找信息规格` | 文本型 | 规格筛选SQL条件 |
| `查找结算品类` | 文本型 | 品类筛选SQL条件 |

**商品销售报表 (窗口_商品销售报表)**:
- **功能**: 销售数据统计，支持按店铺/品类/规格/日期/客户多维度筛选，支持订单/明细/天/月/年/店铺/品类多种视图
- **程序集变量（17个）**:
  | 变量名 | 类型 | 说明 |
  |--------|------|------|
  | `店铺名称_超级列表框组件句柄` | 整数型 | 店铺筛选组件句柄 |
  | `品类名称_超级列表框组件句柄` | 整数型 | 品类筛选组件句柄 |
  | `规格名称_超级列表框组件句柄` | 整数型 | 规格筛选组件句柄 |
  | `客户名称_超级列表框组件句柄` | 整数型 | 客户筛选组件句柄 |
  | `规格名称id_超级列表框组件句柄` | 整数型 | 规格ID筛选组件句柄 |
  | `报表数据_高级表格组件句柄` | 整数型 | 报表表格组件句柄 |
  | `局部_数据订单编码` | 文本型 | 当前查看的订单编码 |
  | `查找开始日期` | 文本型 | 查询开始日期 |
  | `查找结束日期` | 文本型 | 查询结束日期 |
  | `查找信息库房` | 文本型 | 店铺筛选SQL条件 |
  | `查找信息规格` | 文本型 | 规格筛选SQL条件 |
  | `查找信息客户` | 文本型 | 客户筛选SQL条件 |
  | `查找信息名称` | 文本型 | 名称筛选条件 |
  | `查找信息批零` | 文本型 | 批零筛选SQL条件 |
  | `查找信息金重` | 文本型 | 金重筛选SQL条件 |
  | `局部_状态切换中` | 逻辑型 | 状态切换中标记 |
- **子程序列表（30+个）**:
  | 子程序名 | 功能 |
  |----------|------|
  | `_窗口_商品销售报表_创建完毕` | 窗口初始化：加载筛选框、表头、执行查询 |
  | `_子程序_加载表头` | 加载报表表头（根据视图模式动态切换：订单29列/明细32列/天月年19列/店铺19列） |
  | `_编辑框_起始时间_鼠标左键单击` | 开始日期点击事件 |
  | `_编辑框_结束时间_鼠标左键单击` | 结束日期点击事件 |
  | `_单选框_订单_选中状态改变` | 切换订单视图 |
  | `_单选框_明细_选中状态改变` | 切换明细视图 |
  | `_单选框_天_选中状态改变` | 切换按天视图 |
  | `_单选框_月_选中状态改变` | 切换按月视图 |
  | `_单选框_年_选中状态改变` | 切换按年视图 |
  | `_查询条件_初始化` | 初始化查询条件：加载店铺/品类/规格/客户筛选框 |
  | `_商品销售报表_表头回调` | 表头全选/取消全选回调 |
  | `_店铺名称_超级列表框EX_项目左键单击` | 店铺筛选项点击事件 |
  | `_品类名称_超级列表框EX_项目左键单击` | 品类筛选项点击事件 |
  | `_规格名称_超级列表框EX_项目左键单击` | 规格筛选项点击事件 |
  | `_客户名称_超级列表框EX_项目左键单击` | 客户筛选项点击事件 |
  | `_品类选中_id数据初始化` | 品类选中后联动规格筛选 |
  | `_选中规格id数据初始化` | 构建规格ID筛选条件 |
  | `_编辑框_包含信息_鼠标进入离开` | 搜索框提示文本处理 |
  | `_编辑框_包含信息_键盘事件` | 搜索框键盘事件 |
  | `_编辑框_查找信息_鼠标进入离开` | 查找框提示文本处理 |
  | `_编辑框_查找信息_键盘事件` | 查找框键盘事件 |
  | `_客户查找_数据初始化` | 客户搜索初始化 |
  | `_超级按钮EX_重置_鼠标左键单击` | 重置查询条件 |
  | `_按钮EX_导出数据_鼠标左键单击` | 导出Excel |
  | `_超级按钮EX_查询_鼠标左键单击` | 执行查询：构建筛选条件、加载报表数据 |
  | `_报表数据_高级表格Ex_项目右键单击` | 右键菜单：复制单元格/查看订单明细/返回订单列表 |
  | `_订单明细_被选择` | 查看订单明细 |
  | `_复制单元格_被选择` | 复制单元格内容 |
  | `_返回订单_被选择` | 返回订单列表 |
  | `_子程序_商品销售订单` | 加载订单视图数据 |
  | `_子程序_商品销售详情` | 加载明细视图数据 |
  | `_子程序_商品销售日期` | 加载天/月/年视图数据 |
  | `_店铺商品汇总_加载数据` | 加载店铺汇总数据 |
  | `_店铺商品自定义_加载数据` | 加载品类汇总数据 |
- **核心SQL**:
  ```sql
  -- 加载店铺下拉框
  SELECT id AS akufang,CASE WHEN id='0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type='商铺' AND superior='0' AND id in ({权限}) UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({权限}) ORDER BY akufang='0' DESC, akufang

  -- 加载品类下拉框（含未匹配选项）
  SELECT id, title FROM xipunum_erp_category UNION ALL SELECT 0 AS id, '未匹配' AS title ORDER BY CASE WHEN id=0 THEN 0 ELSE 1 END, id

  -- 加载规格下拉框（含未匹配选项）
  SELECT MIN(id) AS id, title FROM xipunum_erp_specs GROUP BY title UNION ALL SELECT 0 AS id, '未匹配' AS title ORDER BY CASE WHEN id=0 THEN 0 ELSE 1 END, id

  -- 按品类筛选规格
  SELECT guige.id AS id, guige.title AS title FROM (SELECT id, category_id, title FROM xipunum_erp_specs UNION ALL SELECT 0, 0, '未匹配') AS guige WHERE guige.category_id in ({品类ID列表}) ORDER BY guige.id ASC

  -- 客户搜索
  SELECT customer_code, memberid, name FROM xipunum_erp_member WHERE memberid like '%{关键字}%' OR name like '%{关键字}%' OR tel like '%{关键字}%' ORDER BY id ASC

  -- 销售订单查询（按订单汇总）
  SELECT tol.djdingdanbian AS chukudanhao,tol.chukutime AS chukutime,... FROM (复杂子查询) as tol WHERE 1=1 {店铺条件} {客户条件} {批零条件} {金重条件} GROUP BY tol.djdingdanbian ORDER BY tol.chukutime DESC

  -- 销售订单合计
  SELECT SUM(ddtol.dingjin),SUM(ddtol.shuliang),SUM(ddtol.jinzhong),... FROM (订单子查询) as ddtol

  -- 销售明细查询
  SELECT b.poduct_code AS bianma,b.creationtime AS chukutime,... FROM xipunum_erp_outbound AS b INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code ... WHERE b.creationtime >= '{开始日期}' AND b.creationtime < '{结束日期}' {店铺条件} {品类条件} {规格条件} {客户条件} {批零条件} {金重条件} ORDER BY b.id DESC

  -- 明细合计
  SELECT SUM(tol.xsshuliang),SUM(tol.xsjinzhong),SUM(tol.zongzhong),SUM(tol.xiaoshoujia),SUM(tol.yingshou),SUM(tol.xiaoshougf*tol.xsjinzhong),SUM(tol.xiaoshoufj),SUM(tol.gongfeilr) FROM (明细子查询) as tol
  ```
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_ksiamges`、`xipunum_erp_user`、`xipunum_erp_type`、`xipunum_erp_member`、`xipunum_erp_about`、`xipunum_erp_presale_order`

**商品入库报表 (窗口_商品入库报表)**:
- **功能**: 入库数据统计，支持订单/明细/天/月/年多种视图
- **程序集变量（16个）**: `店铺名称_超级列表框组件句柄`、`品类名称_超级列表框组件句柄`、`规格名称_超级列表框组件句柄`、`工厂名称_超级列表框组件句柄`、`规格名称id_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`局部_数据订单编码`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找信息品类`、`查找信息规格`、`查找信息工厂`、`查找信息名称`、`查找单据状态`
- **表头列数**: 订单视图20列 / 明细视图28列 / 天月年视图9列
- **涉及表**: `xipunum_erp_store`、`xipunum_erp_store_order`、`xipunum_erp_shop`、`xipunum_erp_about`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_type`、`xipunum_erp_ksiamges`、`xipunum_erp_user`

**商品退库报表 (窗口_商品退库报表)**:
- **功能**: 退库数据统计，支持订单/明细/天/月/年多种视图
- **程序集变量（13个）**: `店铺名称_超级列表框组件句柄`、`品类名称_超级列表框组件句柄`、`规格名称_超级列表框组件句柄`、`工厂名称_超级列表框组件句柄`、`规格名称id_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`局部_数据订单编码`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找信息规格`、`查找信息工厂`、`查找信息名称`
- **表头列数**: 订单视图13列/明细视图18列/天月年视图8列
- **涉及表**: `xipunum_erp_tuiku`、`xipunum_erp_tuiku_order`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_type`、`xipunum_erp_about`、`xipunum_erp_ksiamges`、`xipunum_erp_user`

**商品调拨报表 (窗口_商品调拨报表)**:
- **功能**: 调拨数据统计，支持订单/明细/天/月/年多种视图，支持店铺/日期两种汇总方式
- **程序集变量（17个）**: `店铺名称_超级列表框组件句柄`、`品类名称_超级列表框组件句柄`、`规格名称_超级列表框组件句柄`、`工厂名称_超级列表框组件句柄`、`规格名称id_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`局部_数据订单编码`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找信息规格`、`查找信息工厂`、`查找信息名称`、`汇总数据选中样式`、`查找日期库房`
- **表头列数**: 订单视图15列/明细视图22列/天月年视图15列/店铺视图15列
- **涉及表**: `xipunum_erp_transfer`、`xipunum_erp_transfer_order`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_type`、`xipunum_erp_about`、`xipunum_erp_ksiamges`、`xipunum_erp_user`

**商品回收报表 (窗口_商品回收报表)**:
- **功能**: 回收数据统计，支持订单/明细/天/月/年多种视图，支持品类汇总
- **程序集变量（11个）**: `店铺名称_超级列表框组件句柄`、`品类名称_超级列表框组件句柄`、`局部_数据订单编码`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找信息名称`、`查找信息品类`、`局部_类别id内容`、`查找品类汇总`
- **表头列数**: 订单视图17列/明细视图13列/天月年视图4列
- **涉及表**: `xipunum_erp_retreat`、`xipunum_erp_retreat_order`、`xipunum_erp_retreat_title`、`xipunum_erp_member`、`xipunum_erp_category`、`xipunum_erp_user`、`xipunum_erp_type`

**商品查询报表 (窗口_商品查询报表)**:
- **功能**: 综合商品查询，支持按店铺/品类/规格/工厂/金重/总重/圈号/款号/名称多维度筛选
- **程序集变量（19个）**: `店铺名称_超级列表框组件句柄`、`品类名称_超级列表框组件句柄`、`规格名称_超级列表框组件句柄`、`工厂名称_超级列表框组件句柄`、`规格名称id_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`局部_查找类型文本`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找信息规格`、`查找信息工厂`、`查找信息金重`、`查找信息总重`、`查找信息圈号`、`查找信息款号`、`查找信息名称`
- **表头列数**: 25列（序号/商品编码/原始编码/入库单号/入库时间/品类名称/规格/材质/商品名称/款号/是否镶嵌/工厂/单件重/库存数量/库存金重/库存总重/库房/成本价/预售价/成本工费/成本附加费/参考工费/销售工费/销售附加费/原料克价）
- **涉及表**: `xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_ksiamges`、`xipunum_erp_store`、`xipunum_erp_store_order`、`xipunum_erp_type`、`xipunum_erp_about`

---

#### 7.10.2 运营报表

**销售查询报表 (窗口_销售查询报表)**:
- **功能**: 销售查询，支持按店铺/材质/品类/规格/批零/日期多维度筛选
- **程序集变量（11个）**: `店铺名称_超级列表框组件句柄`、`材质名称_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找信息材质`、`查找库房名称`、`查找材质数据`、`局部_材质查看数量`
- **表头列数**: 20列（序号/店铺名称/品类/规格/实际数量/实际金重/实际总重/实际金额//销售数量/销售金重/销售总重/销售金额/实销金额//客退数量/客退金重/客退总重/客退金额/实退金额）
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_type`、`xipunum_erp_ksiamges`、`xipunum_erp_return`、`xipunum_erp_return_order`

**销售查询简易报表 (窗口_销售查询简易报表)**:
- **功能**: 简易销售查询，支持店铺筛选
- **程序集变量（8个）**: `店铺名称_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找库房名称`、`回收库房名称`、`查找结算品类`
- **表头列数**: 12列（序号/店铺名称/黄金销售/黄金回收/黄金应结/海峡金/本月应结/应结多结/本月实际应结/本月已结/累计未结/备注）
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop`、`xipunum_erp_type`

**销售详情报表 (窗口_销售详情报表)**:
- **功能**: 销售明细查询，支持按店铺/材质/批零/金重/日期多维度筛选
- **程序集变量（12个）**: `店铺名称_超级列表框组件句柄`、`材质名称_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找信息材质`、`查找库房名称`、`查找材质数据`、`查找信息批零`、`查找信息金重`
- **表头列数**: 35列（序号/出库单号/商品编码/销售时间/商品名称/款号/库房/品类/规格/材质/圈口长度/成色/单件重/金重/重量/单位/料价/成本工费/参考工费/成本附加费/成本价/销售单价/销售金额/数量/原附加费/销售克价/销售工费/销售附加费/折扣/应收金额/工费利润/成本工费合计/销售工费合计/批零/状态）
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_type`、`xipunum_erp_ksiamges`、`xipunum_erp_user`、`xipunum_erp_about`

**报表月销售统计 (窗口_报表月销售统计)**:
- **功能**: 月度销售统计，支持按店铺/材质/品类/批零/日期多维度筛选
- **程序集变量（11个）**: `店铺名称_超级列表框组件句柄`、`材质名称_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找信息材质`、`查找库房名称`、`查找材质数据`、`查找结算品类`
- **表头列数**: 23列（序号/店铺名称/品类/批零/实际数量/实际金重/实际总重/实际金额/工费利润/成本工费/销售工费//销售数量/销售金重/销售总重/销售金额/实销金额//客退数量/客退金重/客退总重/客退金额/实退金额）
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_type`、`xipunum_erp_ksiamges`、`xipunum_erp_return`、`xipunum_erp_return_order`、`xipunum_erp_category_stat_config`

**报表月汇总销售统计月 (窗口_报表月汇总销售统计月)**:
- **功能**: 月汇总统计
- **程序集变量**: 同月销售统计
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_type`

**报表员工月销售统计 (窗口_报表员工月销售统计)**:
- **功能**: 员工月销售统计，支持按店铺/品类/规格/工厂/批零/导购员多维度筛选
- **程序集变量（15个）**: `店铺名称_超级列表框组件句柄`、`品类名称_超级列表框组件句柄`、`规格名称_超级列表框组件句柄`、`工厂名称_超级列表框组件句柄`、`规格名称id_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`局部_查找类型文本`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找信息规格`、`查找信息工厂`、`局部_查找导购批零`、`局部_查找导购账户`、`查找结算品类`
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_user`、`xipunum_erp_type`、`xipunum_erp_about`、`xipunum_erp_ksiamges`

**报表店铺收支报表 (窗口_报表店铺收支报表)**:
- **功能**: 店铺收支统计，支持明细/天/月/年多种视图
- **程序集变量（11个）**: `店铺名称_超级列表框组件句柄`、`银行名称_超级列表框组件句柄`、`收支名称_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`查找开始日期`、`查找结束日期`、`汇总数据选中样式`、`查找收支库房`、`查找收支名称`、`查找收支银行`
- **表头列数**: 明细视图12列 / 天月年视图5列
- **涉及表**: `xipunum_erp_finance`、`xipunum_erp_pay`、`xipunum_erp_finance_title`、`xipunum_erp_finance_account`、`xipunum_erp_finance_yinhang`、`xipunum_erp_type`

**报表店铺收支凭证 (窗口_报表店铺收支凭证)**:
- **功能**: 收支凭证查看
- **子程序列表（3个）**: 窗口初始化、窗口尺寸被改变、关闭窗口
- **核心SQL**:
  ```sql
  -- 查询凭证图片
  SELECT images FROM xipunum_erp_finance WHERE number='{单号}' ORDER BY id asc
  ```
- **涉及表**: `xipunum_erp_finance`

**报表商品报表汇总 (窗口_报表商品报表汇总)**:
- **功能**: 商品报表汇总，支持日期/店铺两种视图，天/月/年三种时间维度
- **程序集变量（10个）**: `店铺名称_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`品类名称_超级列表框组件句柄`、`查找信息库房`、`查找出库库房`、`查找分组id数据`、`查找店铺id数据`、`查找调入id数据`、`查找调出id数据`、`查找品类信息id`
- **表头列数**: 31列（序号/日期或店铺名称/销售数量/线上数量/线下量/销售金重/销售总重/销售工费/销售附加费/销售金额/应收金额/预收定金/优惠金额/实收金额/回收金重/回收金额/实际金额/入库总数/入库金重/入库总重/入库成本价/入库预售价/调入数量/调入金重/调入重量/调出数量/调出金重/调出重量/退库数量/退库金重/退库重量）
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_type`、`xipunum_erp_store`、`xipunum_erp_store_order`、`xipunum_erp_transfer`、`xipunum_erp_transfer_order`、`xipunum_erp_tuiku`、`xipunum_erp_tuiku_order`、`xipunum_erp_return`、`xipunum_erp_return_order`

**报表收银汇总表 (窗口_报表收银汇总表)**:
- **功能**: 收银汇总统计
- **程序集变量（1个）**: `店铺查看_超级列表框组件句柄`
- **子程序列表（10个）**: 窗口初始化、加载表头、查询条件初始化、日期选择、查询、重置、导出、店铺筛选
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shoukuan`、`xipunum_erp_pay`、`xipunum_erp_type`

**报表导购回收表 (窗口_报表导购回收表)**:
- **功能**: 导购回收统计，支持员工汇总/店铺汇总/明细三种视图
- **程序集变量（11个）**: `店铺名称_超级列表框组件句柄`、`品类名称_超级列表框组件句柄`、`局部_类别id内容`、`查找开始日期`、`查找结束日期`、`查找信息名称`、`查找信息库房`、`查找信息品类`、`查找品类汇总`、`查找回收名称`
- **表头列数**: 员工汇总6列/店铺汇总6列/明细12列
- **涉及表**: `xipunum_erp_retreat`、`xipunum_erp_retreat_order`、`xipunum_erp_retreat_title`、`xipunum_erp_user`、`xipunum_erp_type`、`xipunum_erp_category`

**报表对照报表 (窗口_报表对照报表)**:
- **功能**: 数据对照分析，支持昨日/今日对比
- **程序集变量（11个）**: `店铺名称_超级列表框组件句柄`、`品类名称_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`工具条EX1句柄`、`局部_查找类型文本`、`查找昨日日期`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找信息品类`
- **表头列数**: 7列（项目/昨日数量/昨日重量/当日数量/当日重量/重量变化/趋势）
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_store`、`xipunum_erp_shop`、`xipunum_erp_transfer`、`xipunum_erp_tuiku`、`xipunum_erp_return`、`xipunum_erp_shop_kucun`

**报表员工绩效 (窗口_报表员工绩效)**:
- **功能**: 员工绩效统计，支持按店铺/岗位/员工/批零多维度筛选，计算销售绩效和回收绩效
- **程序集变量（5个）**:
  | 变量名 | 类型 | 说明 |
  |--------|------|------|
  | `店铺名称_超级列表框组件句柄` | 整数型 | 店铺筛选组件句柄 |
  | `报表数据_高级表格组件句柄` | 整数型 | 报表表格组件句柄 |
  | `查找开始日期` | 文本型 | 查询开始日期 |
  | `查找结束日期` | 文本型 | 查询结束日期 |
  | `查找库房名称` | 文本型 | 选中店铺ID列表 |
- **子程序列表（15个）**:
  | 子程序名 | 功能 |
  |----------|------|
  | `_窗口_报表员工绩效_创建完毕` | 窗口初始化：加载表头、查询条件、报表数据 |
  | `_子程序_加载表头` | 加载表格表头（17列） |
  | `_编辑框_起始时间_鼠标左键单击` | 弹出日期选择器 |
  | `_编辑框_结束时间_鼠标左键单击` | 弹出日期选择器（结束日期>=开始日期） |
  | `_查询条件_初始化` | 加载店铺下拉框、默认日期范围（月初到今天） |
  | `_报表数据详情_基础` | 加载员工列表和绩效数据 |
  | `_报表数据详情_销售` | 计算销售绩效 |
  | `_报表数据详情_回收` | 计算回收绩效 |
  | `_按钮_查询_被单击` | 查询按钮事件 |
  | `_按钮_重置_被单击` | 重置查询条件 |
  | `_按钮_导出_被单击` | 导出Excel报表 |
  | `_绩效查询报表_表头回调` | 店铺全选/反选回调 |
- **表头列数**: 17列（序号/店铺/岗位/员工/批零/库房id/岗位id/员工账户/销售数量/回收数量/销售重量/回收重量/销售金额/回收金额/销售绩效/回收绩效/绩效合计）
- **核心SQL**:
  ```sql
  -- 加载店铺下拉框
  SELECT id AS akufang,CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in ({权限}) UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({权限}) ORDER BY akufang = '0' DESC, akufang

  -- 加载员工列表
  SELECT b.title AS kufang, c.title AS gangwei, a.USER AS USER, a.NAME AS NAME, a.department as kufangid, a.post as gangweiid FROM xipunum_erp_user AS a INNER JOIN xipunum_erp_type AS b ON b.id = a.department INNER JOIN xipunum_erp_type AS c ON c.id = a.post WHERE a.department IN ({库房ID}) and a.state='0' ORDER BY a.department,a.post asc

  -- 查询销售数据（按店铺/岗位/品类/导购员/批零分组）
  SELECT a.kufang as dianpuid, h.post as gangweiid, CASE WHEN COALESCE(d.id,'')='' THEN '0' ELSE d.id END AS pinleiid, a.shopping_guide AS daogou, a.pling as pingling, sum(COALESCE(a.quantity,0)) as xsshu, sum(COALESCE(a.net_weight,0)) as xszhong, sum(COALESCE(a.settlement,0)) as xsjine FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_ksiamges AS c ON c.kuanhao = b.item_number LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = c.specification_id LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id LEFT JOIN xipunum_erp_category AS d ON d.id = COALESCE(e1.category_id,e2.category_id) INNER JOIN xipunum_erp_user AS h ON h.user = a.shopping_guide WHERE a.creationtime >= '{开始日期}' AND a.creationtime < '{结束日期}' and a.kufang in ({库房ID}) GROUP BY a.kufang,h.post,pinleiid,a.shopping_guide,a.pling ORDER BY a.kufang,h.post,a.pling asc

  -- 查询回收数据（按店铺/岗位/品类/导购员/批零分组）
  SELECT h.department AS kufangid, h.post AS gangwei, b.category_id as pinleiid, a.shopping_guide as daogou, CASE WHEN d.pling='批发' THEN d.pling ELSE '零售' END AS pingling, sum(COALESCE(a.quantity,0)) as hsshu, sum(COALESCE(a.jin_zhong,0)) as hszhong, sum(COALESCE(a.retreat_amount,0)) as hsjine FROM xipunum_erp_retreat AS a INNER JOIN xipunum_erp_retreat_title AS b ON b.id = a.product_name INNER JOIN xipunum_erp_retreat_order AS c ON c.id = a.order_id left JOIN xipunum_erp_outbound_order AS d ON d.retrea_umber = c.retrea_umber INNER JOIN xipunum_erp_user AS h ON h.USER = a.shopping_guide WHERE a.creationtime >= '{开始日期}' AND a.creationtime < '{结束日期}' AND h.department IN ({库房ID}) GROUP BY h.department,h.post,b.category_id,a.shopping_guide,pingling ORDER BY h.department,h.post,pingling asc

  -- 查询绩效配置范围
  SELECT fanwei FROM xipunum_erp_category_score WHERE piling='{批零}' and gangweiid='{岗位ID}' GROUP BY fanwei ORDER BY id asc LIMIT 1

  -- 查询绩效配置详情（5档次）
  SELECT categoryid,fanwei,danwei,data1,data1num,data2,data2num,data3,data3num,data4,data4num,data5,data5num FROM xipunum_erp_category_score WHERE piling='{批零}' and gangweiid='{岗位ID}' and categoryid='{品类ID}' GROUP BY categoryid ORDER BY id asc
  ```
- **绩效计算逻辑**: 按5档次（data1-data5）计算，每个档次有最小值（data1num-data5num）和参数值（data1-data5），根据销售/回收数量/重量/金额匹配档次
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_user`、`xipunum_erp_type`、`xipunum_erp_performance`、`xipunum_erp_retreat`、`xipunum_erp_retreat_order`、`xipunum_erp_category_score`、`xipunum_erp_ksiamges`、`xipunum_erp_specs`、`xipunum_erp_category`

---

#### 7.10.3 其他

**款式数据汇总 (窗口_款式数据汇总)**:
- **功能**: 款号数据汇总分析，支持图片识别搜索
- **程序集变量（19个）**: `店铺名称_超级列表框组件句柄`、`材质名称_超级列表框组件句柄`、`报表数据_图标列表框组件句柄`、`品类名称_超级列表框组件句柄`、`规格名称_超级列表框组件句柄`、`规格名称id_超级列表框组件句柄`、`查找信息库房`、`查找信息规格`、`查找信息材质`、`查找信息名称`、`局部_数据插入计次`、`局部_加载当前页码`、`局部_加载每页数量`、`数据插入计次`、`局部_查询款号文本`、`局部_选中图片id`、`局部_详情状态`、`局部_款号文本内容`
- **子程序列表（25+个）**: 窗口初始化、加载表头、查询条件初始化、表头回调、筛选项点击、查询、重置、款号图片点击、下一页、以图搜款
- **核心SQL**:
  ```sql
  SELECT a.id AS id,a.kuanhao AS kuanhao,a.title AS title,a.caizhi AS caizhi,a.yimage AS yimage,b.title AS pinlei,c.title AS guige FROM xipunum_erp_ksiamges AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id WHERE 1=1 {店铺条件} {品类条件} {规格条件} {材质条件} {名称条件} ORDER BY a.id DESC
  ```
- **涉及表**: `xipunum_erp_ksiamges`、`xipunum_erp_category`、`xipunum_erp_specs`

**运营员工业绩 (窗口_运营员工业绩)**:
- **功能**: 运营业绩统计，支持导购业绩表和商品明细表两种视图
- **程序集变量（15个）**: `店铺名称_超级列表框组件句柄`、`品类名称_超级列表框组件句柄`、`规格名称_超级列表框组件句柄`、`工厂名称_超级列表框组件句柄`、`规格名称id_超级列表框组件句柄`、`报表数据_高级表格组件句柄`、`局部_查找类型文本`、`查找开始日期`、`查找结束日期`、`查找信息库房`、`查找信息规格`、`查找信息工厂`、`局部_查找导购批零`、`局部_查找导购账户`
- **子程序列表（20+个）**: 窗口初始化、加载表头、查询条件初始化、日期选择、查询、重置、导出、视图切换、导购业绩数据列表/详情
- **核心SQL**:
  ```sql
  SELECT a.shopping_guide AS zhanghu,b.NAME AS daogou,... FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_user AS b ON b.USER = a.shopping_guide ... WHERE a.creationtime >= '{开始日期}' AND a.creationtime < '{结束日期}' {店铺条件} {品类条件} {规格条件} {批零条件} GROUP BY a.shopping_guide ORDER BY ...
  ```
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_user`、`xipunum_erp_type`、`xipunum_erp_ksiamges`

**款式数据汇总明细 (窗口_款式数据汇总明细)**:
- **功能**: 款号数据明细查看，支持入库/销售/库存三种明细视图
- **程序集变量（12个）**: `选中账户信息内容`、`查找款号id信息`、`查找款号状态`、`查询款号文本`、`查询款号品类`、`查询款号规格`、`查询款号材质`、`报表数据_高级表格组件句柄`、`查找信息库房`、`查找信息规格`、`查找信息材质`
- **子程序列表（10个）**: 窗口初始化、窗口尺寸被改变、加载表头、查询条件、加载数据、导出、单选框切换
- **核心SQL**:
  ```sql
  -- 查询款号信息
  SELECT a.kuanhao as kuanhao,c.title as pinlei,b.title as guige,a.caizhi as caizhi FROM xipunum_erp_ksiamges as a LEFT JOIN xipunum_erp_specs AS b ON b.id = a.specification_id LEFT JOIN xipunum_erp_category AS c ON c.id = b.category_id WHERE a.id='{款号ID}' LIMIT 1

  -- 查询入库明细
  SELECT ... FROM xipunum_erp_store AS a INNER JOIN xipunum_erp_store_order AS b ON b.id = a.order_id ... WHERE a.item_number='{款号}' {库房条件}

  -- 查询销售明细
  SELECT ... FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_outbound_order AS b ON b.id = a.order_id ... WHERE a.item_number='{款号}' {库房条件}

  -- 查询库存明细
  SELECT ... FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code ... WHERE b.item_number='{款号}' {库房条件}
  ```
- **涉及表**: `xipunum_erp_shop`、`xipunum_erp_ksiamges`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_store`、`xipunum_erp_store_order`、`xipunum_erp_outbound`、`xipunum_erp_outbound_order`、`xipunum_erp_shop_kucun`、`xipunum_erp_type`

**款式数据汇总预览 (窗口_款式数据汇总预览)**:
- **功能**: 款号图片预览
- **子程序列表（3个）**: 窗口初始化、窗口尺寸被改变、关闭窗口
- **核心SQL**:
  ```sql
  -- 查询款号图片信息
  SELECT a.id as aid,a.title as atitle,a.kuanhao as akuanhao,a.category_id as acategory_id,a.specification_id as aspecification_id,a.caizhi as caizhi,a.images as aimages,a.yimage as ayimage,c.title AS cpinlei,d.title AS dguige,a.xiangqian as axiangqian,a.lingxiao as alingxiao FROM xipunum_erp_ksiamges as a INNER JOIN xipunum_erp_category AS c ON c.id = a.category_id INNER JOIN xipunum_erp_specs AS d ON d.id = a.specification_id where a.id='{款号ID}' LIMIT 1
  ```
- **涉及表**: `xipunum_erp_ksiamges`、`xipunum_erp_category`、`xipunum_erp_specs`

**盘点入库导入编码 (窗口_盘点入库导入编码)**:
- **功能**: 批量导入盘点商品编码
- **程序集变量（1个）**: `副编码_超级列表框组件句柄`
- **子程序列表（4个）**: 窗口初始化、窗口尺寸被改变、关闭窗口、重置、保存
- **涉及表**: `xipunum_erp_shop`、`data/pandian.mdb`

**物资盘点添加 (窗口_物资盘点添加)**:
- **功能**: 库存盘点操作，支持成品/物资两种盘点类型
- **程序集变量（3个）**: `局部_订单是否选中`、`集_行号`、`集_列号`
- **子程序列表（20+个）**: 窗口初始化、窗口尺寸被改变、查询条件初始化、加载表头、删除表格、商品编码输入、查找商品、盘点详情、差异计算、创建盘点、保存盘点、重置、全选/反选、数据统计
- **核心SQL**:
  ```sql
  -- 查询商品库存
  SELECT ... FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE a.kufang='{库房}' AND (a.quantity > 0 OR a.jinzhong > 0) {品类条件} {材质条件} {规格条件}

  -- 查询盘点数据
  SELECT * FROM pandianlist WHERE data1='{编码}'

  -- 保存盘点数据
  UPDATE pandianlist SET data13='{盘点数量}',data14='{盘点金重}',data15='{差异}',data16='{差异重量}' WHERE data1='{编码}'
  ```
- **涉及表**: `xipunum_erp_shop`、`xipunum_erp_shop_kucun`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_type`、`data/pandian.mdb`

---

### 7.11 系统管理模块

#### 7.11.1 账户管理

**个人信息 (窗口_个人信息)**:
- **功能**: 个人信息查看与修改
- **程序集变量（1个）**:
  | 变量名 | 类型 | 说明 |
  |--------|------|------|
  | `查找用户名` | 文本型 | 当前登录用户名 |
- **子程序列表（10个）**:
  | 子程序名 | 功能 |
  |----------|------|
  | `_窗口_个人信息_创建完毕` | 窗口初始化：加载用户信息、省份列表 |
  | `_窗口_个人信息_尺寸被改变` | 窗口自适应布局 |
  | `_图片框EX4_鼠标左键单击` | 关闭窗口 |
  | `_按钮EX1_鼠标左键单击` | 保存按钮事件 |
  | `_按钮EX2_鼠标左键单击` | 重置按钮事件 |
  | `_组合框_省份_内容被改变` | 省份选择联动：加载城市列表 |
  | `_组合框_市区_内容被改变` | 城市选择联动：加载区县列表 |
  | `_组合框_县区_内容被改变` | 区县选择联动：拼接地址 |
  | `_个人信息_加载省份列表` | 加载省份下拉框 |
  | `_个人信息_加载城市列表` | 加载城市下拉框 |
  | `_个人信息_保存数据` | 保存个人信息 |
- **核心SQL**:
  ```sql
  -- 查询用户信息
  SELECT a.name as aname,a.mailbox as amailbox,a.tel as atel,a.type as atype,a.dizhi as adizhi,b.title as btitle,c.title as ctitle FROM xipunum_erp_user AS a INNER JOIN xipunum_erp_type AS b ON b.id = a.department INNER JOIN xipunum_erp_type AS c ON c.id = a.post WHERE a.USER = '{用户名}' LIMIT 1

  -- 更新用户信息
  UPDATE xipunum_erp_user SET dizhi='{地址}' WHERE user='{用户名}' LIMIT 1
  ```
- **涉及表**: `xipunum_erp_user`、`xipunum_erp_type`

**账户添加修改 (窗口_账户添加修改)**:
**功能**: 用户新增/编辑

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `选中账户信息内容` | 文本型 | 选中的账户用户名（空=新增，非空=编辑） |

**子程序列表（20个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_账户添加修改_创建完毕` | 窗口初始化：根据选中账户设置界面、加载分组/岗位下拉框 |
| `_窗口_账户添加修改_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX1_鼠标左键单击` | 保存/修改按钮事件 |
| `_按钮EX2_鼠标左键单击` | 重置按钮事件 |
| `_账户添加_保存数据` | 保存账户数据：验证用户名→检查重复→加密密码→添加账户 |
| `_账户修改_修改数据` | 修改账户数据：更新账户信息 |
| `_主窗口_弹出日期框` | 弹出日期选择器 |
| `_单选框_商铺_选中状态改变` | 切换商铺类型 |
| `_单选框_后台_选中状态改变` | 切换后台类型 |
| `_单选框_全部_选中状态改变` | 切换全部权限 |
| `_单选框_店铺_选中状态改变` | 切换店铺权限 |
| `_单选框_岗位_选中状态改变` | 切换岗位权限 |
| `_单选框_个人_选中状态改变` | 切换个人权限 |
| `_单选框_正常_选中状态改变` | 切换正常状态 |
| `_单选框_停用_选中状态改变` | 切换停用状态 |
| `_单选框_开启_选中状态改变` | 切换开启谷歌验证 |
| `_单选框_关闭_选中状态改变` | 切换关闭谷歌验证 |
| `_单选框_所有_选中状态改变` | 切换所有销售检索 |
| `_单选框_开票_选中状态改变` | 切换开票销售检索 |

**核心SQL**:
```sql
-- 加载账户详情
SELECT a.user AS auser,a.jianxie as ajianxie,a.NAME AS aname,a.state as astate,a.mailbox AS amailbox,a.tel AS atel,a.type AS atype,a.google as agoogle,b.title AS btitle,c.title AS ctitle,a.dizhi AS adizhi,a.data as adata,a.ksdate as aksdate,a.jsdate as ajsdate,a.namejx as anamejx,a.xsrole as axsrole,a.xsdata as xsdata FROM xipunum_erp_user AS a INNER JOIN xipunum_erp_type AS b ON b.id = a.department INNER JOIN xipunum_erp_type AS c ON c.id = a.post WHERE a.USER = '{用户名}' LIMIT 1

-- 检查用户名是否存在
SELECT * FROM xipunum_erp_user where user='{用户名}'

-- 添加账户
增加记录("xipunum_erp_user", user, jianxie, password, name, namejx, state, mailbox, tel, type, google, department, post, dizhi, data, ksdate, jsdate, xsrole, xsdata, cjuser, creationtime)

-- 修改账户
UPDATE xipunum_erp_user SET jianxie='{简写}',name='{姓名}',namejx='{名称简写}',state='{状态}',mailbox='{邮箱}',tel='{电话}',type='{类型}',google='{谷歌验证}',department='{分组}',post='{岗位}',dizhi='{地址}',data='{权限}',ksdate='{上班时间}',jsdate='{下班时间}',xsrole='{销售检索}',xsdata='{数据主备}',updatetime='{时间}' WHERE user='{用户名}' LIMIT 1

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_user`、`xipunum_erp_type`、`xipunum_erp_role`、`xipunum_erp_xitong_log`

---

**密码修改 (窗口_密码修改)**:
**功能**: 密码修改

**子程序列表（5个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_密码修改_创建完毕` | 窗口初始化：清空输入框、设置密码模式 |
| `_窗口_密码修改_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX2_鼠标左键单击` | 重置按钮事件 |
| `_按钮EX1_鼠标左键单击` | 保存按钮事件：验证原始密码→检查新密码强度→更新密码→注销登录 |

**核心SQL**:
```sql
-- 验证原密码
SELECT * FROM xipunum_erp_user where user='{用户名}' and password='{加密后密码}' LIMIT 1

-- 更新密码
UPDATE xipunum_erp_user SET password='{新密码}',updatetime='{时间}' WHERE xipunum_erp_user.user='{用户名}' LIMIT 1

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type='修改', title='修改账户密码', conter, user, creationtime)
```

**涉及表**: `xipunum_erp_user`、`xipunum_erp_xitong_log`

**个人信息 (窗口_个人信息)**:
- **功能**: 个人信息查看
- **涉及表**: `xipunum_erp_user`、`xipunum_erp_type`

---

#### 7.11.2 岗位权限

**岗位分组 (主窗口_岗位分组)**:
- **功能**: 岗位和权限管理，支持分组管理、岗位管理、栏目可视/操作/店铺权限配置
- **程序集变量（4个）**: `岗位分组_高级表格组件句柄`、`栏目可视_超级列表框组件句柄`、`栏目操作_超级列表框组件句柄`、`店铺查看_超级列表框组件句柄`
- **子程序列表（15个）**:
  | 子程序名 | 功能 |
  |----------|------|
  | `_主窗口_岗位分组_初始化` | 岗位分组页面初始化：根据权限设置商铺/后台单选框 |
  | `_主窗口_岗位分组_数据加载` | 加载岗位列表数据 |
  | `_主窗口_岗位分组_基础加载` | 加载分组下拉框、栏目可视/操作/店铺列表 |
  | `_岗位分组_添加分组_鼠标左键单击` | 添加分组 |
  | `_岗位分组_添加岗位_鼠标左键单击` | 保存/修改岗位 |
  | `_主窗口_岗位分组_保存数据` | 保存岗位数据：验证→检查重复→添加岗位→添加权限 |
  | `_岗位分组_删除_被选择` | 删除岗位（检查是否有关联账户） |
  | `_岗位分组_超级列表框EX_项目右键单击` | 右键菜单 |
  | `_岗位分组_单选框商铺1_选中状态改变` | 切换商铺类型 |
  | `_岗位分组_单选框后台1_选中状态改变` | 切换后台类型 |
- **表头列数**: 8列（序号/类型/分组名称/岗位名称/可视权限/操作权限/店铺权限/岗位id）
- **核心SQL**:
  ```sql
  -- 加载岗位列表（全部权限）
  SELECT a.type as atype,b.title as btitle,a.title as atitle,c.keshi AS ckeshi,c.caozuo AS ccaozuo,c.shopid AS cshopid,a.id as aid FROM xipunum_erp_type as a INNER JOIN xipunum_erp_type AS b ON b.id = a.superior left JOIN xipunum_erp_role AS c ON c.typeid = a.id WHERE a.title LIKE '%{关键字}%' AND a.superior >'0' ORDER BY a.id DESC

  -- 加载岗位列表（店铺权限）
  SELECT a.type as atype,b.title as btitle,a.title as atitle,c.keshi AS ckeshi,c.caozuo AS ccaozuo,c.shopid AS cshopid,a.id as aid FROM xipunum_erp_type as a INNER JOIN xipunum_erp_type AS b ON b.id = a.superior left JOIN xipunum_erp_role AS c ON c.typeid = a.id WHERE a.title LIKE '%{关键字}%' AND a.superior ='{分组ID}' ORDER BY a.id DESC

  -- 检查岗位是否存在
  SELECT * FROM xipunum_erp_type where type='{类型}' and title='{名称}' and superior='{分组ID}'

  -- 添加岗位
  增加记录("xipunum_erp_type", type, superior, title, cjuser, creationtime)

  -- 查询岗位ID
  SELECT id FROM xipunum_erp_type WHERE type='{类型}' and superior='{分组ID}' and title='{名称}' LIMIT 1

  -- 添加岗位权限
  增加记录("xipunum_erp_role", typeid, keshi, caozuo, shopid, creationtime)

  -- 删除岗位
  DELETE FROM xipunum_erp_type WHERE id='{ID}'

  -- 删除岗位权限
  DELETE FROM xipunum_erp_role WHERE typeid='{岗位ID}'

  -- 添加系统日志
  增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
  ```
- **涉及表**: `xipunum_erp_role`、`xipunum_erp_type`、`erp_navigation`、`erp_navigation_type`、`xipunum_erp_xitong_log`

---

#### 7.11.3 系统设置

**系统设置 (主窗口_系统设置)**:
- **功能**: 系统参数配置，支持公司名称、款号识别地址、优惠比例、软件LOGO、用户头像、窗口图标、退库仓、线上收款等配置
- **子程序列表（8个）**:
  | 子程序名 | 功能 |
  |----------|------|
  | `_主窗口_系统设置_初始化` | 加载系统设置子菜单（系统设置/工厂管理/商品来源/结算方式/品类管理/规格管理/回收名称/品类属性） |
  | `_图标列表框EX_系统设置_项目左键单击` | 菜单点击路由：根据名称跳转到对应功能 |
  | `_主窗口_图标列表框EX_系统设置` | 加载系统设置表单：库房下拉框、收款下拉框、当前配置值 |
  | `_图片框Ex_软件logo_鼠标左键单击` | 上传软件LOGO图片（HTTP上传） |
  | `_图片框Ex_用户头像_鼠标左键单击` | 上传用户头像图片（HTTP上传） |
  | `_图片框Ex_窗口图标_鼠标左键单击` | 上传窗口图标图片（HTTP上传） |
  | `_按钮EX_系统设置保存_鼠标左键单击` | 保存系统配置：权限检查(8系统参数修改)→比较变更→UPDATE配置表 |
  | `_按钮EX_系统设置重置_鼠标左键单击` | 重置系统设置表单 |
- **配置项映射**:
  | 配置项 | 配置表ID | 说明 |
  |--------|----------|------|
  | 公司名称 | 1 | xipunum_erp_config.id=1 |
  | 软件LOGO | 2 | xipunum_erp_config.id=2 |
  | 用户头像 | 3 | xipunum_erp_config.id=3 |
  | 款号识别地址 | 4 | xipunum_erp_config.id=4 |
  | 优惠比例 | 5 | xipunum_erp_config.id=5 |
  | 退库仓 | 6 | xipunum_erp_config.id=6 |
  | 线上收款 | 7 | xipunum_erp_config.id=7 |
  | 窗口图标 | 8 | xipunum_erp_config.id=8 |
- **核心SQL**:
  ```sql
  -- 加载系统设置子菜单
  SELECT * FROM erp_navigation WHERE role in ({权限}) and superior='{栏目ID}' and state='0' order by id ASC

  -- 加载库房下拉框
  SELECT id AS akufang,CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in ({权限}) UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({权限}) ORDER BY akufang = '0' DESC, akufang

  -- 加载收款方式下拉框
  SELECT id,name FROM xipunum_erp_pay WHERE 1=1 ORDER BY id asc

  -- 更新公司名称
  UPDATE xipunum_erp_config SET conter='{公司名称}' WHERE id='1' LIMIT 1

  -- 更新软件LOGO
  UPDATE xipunum_erp_config SET conter='{LOGO地址}' WHERE id='2' LIMIT 1

  -- 更新用户头像
  UPDATE xipunum_erp_config SET conter='{头像地址}' WHERE id='3' LIMIT 1

  -- 更新款号识别地址
  UPDATE xipunum_erp_config SET conter='{识别地址}' WHERE id='4' LIMIT 1

  -- 更新优惠比例
  UPDATE xipunum_erp_config SET conter='{优惠比例}' WHERE id='5' LIMIT 1

  -- 更新退库仓
  UPDATE xipunum_erp_config SET conter='{退库仓ID}' WHERE id='6' LIMIT 1

  -- 更新线上收款
  UPDATE xipunum_erp_config SET conter='{收款ID}' WHERE id='7' LIMIT 1

  -- 更新窗口图标
  UPDATE xipunum_erp_config SET conter='{图标地址}' WHERE id='8' LIMIT 1
  ```
- **涉及表**: `xipunum_erp_config`、`xipunum_erp_type`、`xipunum_erp_pay`、`erp_navigation`

**商铺信息**:
- **功能**: 店铺信息配置
- **涉及表**: `xipunum_erp_type`

---

#### 7.11.4 谷歌验证
- `谷歌验证绑定` — 首次绑定谷歌验证（详见7.4.10）
- `谷歌验证码输入` — 登录时输入验证码（详见7.4.11）
- `谷歌验证重置` — 重置谷歌验证（详见7.4.12）

---

#### 7.11.5 日志管理

**日志记录**:
- **功能**: 操作日志查看
- **涉及表**: `xipunum_erp_xitong_log`、`xipunum_erp_user_log`

**登录日志 (窗口_登录日志)**:
- **功能**: 用户登录日志查看，支持清空日志
- **程序集变量（1个）**:
  | 变量名 | 类型 | 说明 |
  |--------|------|------|
  | `选中账户信息内容` | 文本型 | 选中的账户用户名 |
- **子程序列表（5个）**:
  | 子程序名 | 功能 |
  |----------|------|
  | `_窗口_登录日志_创建完毕` | 窗口初始化：加载登录日志数据 |
  | `_窗口_登录日志_尺寸被改变` | 窗口自适应布局 |
  | `_图片框EX4_鼠标左键单击` | 关闭窗口 |
  | `_日志记录_清空按钮_鼠标左键单击` | 清空登录日志（权限：11删除账户） |
  | `子程序_加载数据` | 加载登录日志数据 |
- **核心SQL**:
  ```sql
  -- 查询登录日志
  SELECT ip,conter,creationtime FROM xipunum_erp_user_log WHERE user='{用户名}' ORDER BY id DESC

  -- 清空登录日志
  DELETE FROM xipunum_erp_user_log WHERE user='{用户名}'

  -- 添加系统日志
  增加记录("xipunum_erp_xitong_log", type='删除', title='清空登录日志信息', conter, user, creationtime)
  ```
- **涉及表**: `xipunum_erp_user_log`、`xipunum_erp_user`、`xipunum_erp_xitong_log`

---

#### 7.11.6 更新管理
- `在线更新` — 软件在线更新（详见7.4.14）
- `模块更新` — 模块更新（详见7.4.15）

---

#### 7.11.7 其他

**授权码信息输入** — 授权码录入（详见7.4.13）

**款号管理 (窗口_款号管理)**:
- **功能**: 款号信息管理，支持图片识别、分页加载
- **程序集变量（6个）**: `每页数量`、`当前页码`、`开始数据`、`局部_款号识别数据文本`、`数据插入计次`、`局部_选中图片id`
- **子程序列表（15+个）**: 窗口初始化、窗口尺寸被改变、商品组合框初始化、品类选择联动、规格选择联动、查询、重置、图片识别、款号添加、款号修改、款号删除、图片上传、分页加载
- **核心SQL**:
  ```sql
  -- 查询商品款号统计
  SELECT a.poduct_code,b.item_number FROM xipunum_erp_shop_kucun as a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE a.kufang in ({权限}) and a.quantity !='0' ORDER BY a.id ASC

  -- 查询有款号商品统计
  SELECT a.poduct_code,b.item_number FROM xipunum_erp_shop_kucun as a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code and b.item_number !='' WHERE a.kufang in ({权限}) and a.quantity !='0' ORDER BY a.id ASC

  -- 查询款号列表
  SELECT a.id AS id,a.title AS title,a.kuanhao AS kuanhao,a.yimage AS yimage,a.caizhi AS caizhi,a.chengse AS chengse,a.lingxiao AS lingxiao,b.title AS pinlei,c.title AS guige FROM xipunum_erp_ksiamges AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id WHERE 1=1 {品类条件} {材质条件} {规格条件} ORDER BY a.id DESC

  -- 添加款号
  增加记录("xipunum_erp_ksiamges", title, kuanhao, yimage, caizhi, chengse, category_id, specification_id, lingxiao, cjuser, creationtime)

  -- 修改款号
  UPDATE xipunum_erp_ksiamges SET title='{名称}',caizhi='{材质}',chengse='{成色}',category_id='{品类ID}',specification_id='{规格ID}',lingxiao='{零销售}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1

  -- 删除款号
  DELETE FROM xipunum_erp_ksiamges WHERE id='{ID}'
  ```
- **涉及表**: `xipunum_erp_ksiamges`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_shop`、`xipunum_erp_shop_kucun`

**款号合并 (窗口_款号合并)**:
- **功能**: 款号合并操作，支持主款号/合并款号选择
- **程序集变量（5个）**: `集_行号`、`集_列号`、`删除按钮`、`局部_查询方式`、`局部_款号识别数据文本`
- **子程序列表（10+个）**: 窗口初始化、窗口尺寸被改变、删除表格、加载表头、查询、重置、主款号选择、合并款号选择、合并操作、删除操作
- **核心SQL**:
  ```sql
  -- 查询款号列表
  SELECT ... FROM xipunum_erp_ksiamges AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id WHERE 1=1 ORDER BY a.id DESC

  -- 合并款号（更新商品表）
  UPDATE xipunum_erp_shop SET item_number='{主款号}' WHERE item_number='{合并款号}'

  -- 合并款号（更新入库表）
  UPDATE xipunum_erp_store SET item_number='{主款号}' WHERE item_number='{合并款号}'

  -- 删除合并款号
  DELETE FROM xipunum_erp_ksiamges WHERE id='{ID}'
  ```
- **涉及表**: `xipunum_erp_ksiamges`、`xipunum_erp_shop`、`xipunum_erp_store`、`xipunum_erp_category`、`xipunum_erp_specs`

**批量打印 (窗口_批量打印)**:
- **功能**: 批量打印商品标签
- **程序集变量（3个）**: `集_行号`、`集_列号`、`删除按钮`
- **子程序列表（10+个）**: 窗口初始化、窗口尺寸被改变、加载表头、删除表格、商品编码输入、查找商品、添加到打印列表、打印预览、批量打印
- **核心SQL**:
  ```sql
  -- 加载库房下拉框
  SELECT id AS akufang,CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in ({权限}) UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({权限}) ORDER BY akufang = '0' DESC, akufang

  -- 查询商品信息
  SELECT a.poduct_code,a.product_name,a.caizhi,a.item_number,a.single,a.quantity,a.net_weight,a.weight,b.title AS pinlei,c.title AS guige FROM xipunum_erp_shop AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id WHERE a.poduct_code='{编码}'
  ```
- **涉及表**: `xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_type`

**批量导入无款 (窗口_批量导入无款)**:
- **功能**: 批量导入商品数据（无款号），支持Excel导入
- **程序集变量（1个）**: `树形列表框Ex1_组件句柄`
- **子程序列表（15+个）**: 窗口初始化、加载表头、加载数据、导出格式、重置、Excel导入、数据预览、批量保存、删除行
- **表头列数**: 44列（序号/商品编码/商品名称/商品品类/规格/款号/材质/镶嵌/圈口长度/面宽/厚度/工厂成色/公司成色/单件重/数量/金重/损耗/含耗重/重量/单位/石重/石头数/副石重/副石头数/成本价/系数/成本工费/参考工费/销售工费/成本附加费/销售附加费/销售价/工厂名称/来源/结算方式/送货单号/半成品/备注/入库时间/主石色/证书编码/检测结果/形状/颜色）
- **涉及表**: `xipunum_erp_shop`、`xipunum_erp_store`、`xipunum_erp_store_order`、`xipunum_erp_shop_kucun`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_about`、`xipunum_erp_type`、`xipunum_erp_zhengshu`

**款号管理匹配 (窗口_款号管理匹配)**:
- **功能**: 款号匹配商品，支持按编码/款号搜索商品并匹配款号
- **程序集变量（9个）**: `查找款号id信息`、`局部_款号数据名称`、`局部_款号数据规格id`、`局部_款号数据材质`、`局部_款号数据名称原始图`、`局部_数据品类名称`、`局部_数据规格名称`、`局部_款号数据镶嵌`、`局部_款号数据信息`
- **子程序列表（5个）**: 窗口初始化、窗口尺寸被改变、查找框回车键、查找按钮、匹配按钮
- **核心SQL**:
  ```sql
  -- 查询款号信息
  SELECT a.id as aid,a.title as atitle,a.kuanhao as akuanhao,a.category_id as acategory_id,a.specification_id as aspecification_id,a.caizhi as caizhi,a.images as aimages,a.yimage as ayimage,c.title AS cpinlei,d.title AS dguige,a.xiangqian as axiangqian,a.lingxiao as alingxiao FROM xipunum_erp_ksiamges as a INNER JOIN xipunum_erp_category AS c ON c.id = a.category_id INNER JOIN xipunum_erp_specs AS d ON d.id = a.specification_id where a.id='{款号ID}' LIMIT 1

  -- 查询商品信息
  SELECT * FROM xipunum_erp_shop where poduct_code='{编码}' or fu_code='{编码}'

  -- 匹配款号
  UPDATE xipunum_erp_shop SET item_number='{款号}' WHERE poduct_code='{编码}'
  ```
- **涉及表**: `xipunum_erp_ksiamges`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_specs`

**款号管理添加 (窗口_款号管理添加)**:
- **功能**: 款号添加/修改，支持图片上传
- **程序集变量（10个）**: `查找款号id信息`、`局部_款号数据名称`、`局部_款号数据品类id`、`局部_款号数据规格id`、`局部_款号数据材质`、`局部_款号数据名称原始图`、`局部_数据品类名称`、`局部_数据规格名称`、`局部_图片路径`、`局部_图片响应`
- **子程序列表（10+个）**: 窗口初始化、窗口尺寸被改变、商品组合框初始化、品类选择联动、规格选择联动、图片上传、确定按钮、重置按钮
- **核心SQL**:
  ```sql
  -- 查询款号信息
  SELECT a.id as aid,a.title as atitle,a.kuanhao as akuanhao,a.category_id as acategory_id,a.specification_id as aspecification_id,a.caizhi as caizhi,a.images as aimages,a.yimage as ayimage,c.title AS cpinlei,d.title AS dguige,a.xiangqian as axiangqian,a.lingxiao as alingxiao FROM xipunum_erp_ksiamges as a INNER JOIN xipunum_erp_category AS c ON c.id = a.category_id INNER JOIN xipunum_erp_specs AS d ON d.id = a.specification_id where a.id='{款号ID}' LIMIT 1

  -- 添加款号
  增加记录("xipunum_erp_ksiamges", title, kuanhao, yimage, caizhi, chengse, category_id, specification_id, xiangqian, lingxiao, cjuser, creationtime)

  -- 修改款号
  UPDATE xipunum_erp_ksiamges SET title='{名称}',caizhi='{材质}',chengse='{成色}',category_id='{品类ID}',specification_id='{规格ID}',xiangqian='{镶嵌}',lingxiao='{零销售}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1
  ```
- **涉及表**: `xipunum_erp_ksiamges`、`xipunum_erp_category`、`xipunum_erp_specs`

**款号管理图片预览 (窗口_款号管理图片预览)**:
- **功能**: 款号图片预览
- **子程序列表（3个）**: 窗口初始化、窗口尺寸被改变、关闭窗口
- **核心SQL**:
  ```sql
  -- 查询款号图片信息
  SELECT a.id as aid,a.title as atitle,a.kuanhao as akuanhao,a.category_id as acategory_id,a.specification_id as aspecification_id,a.caizhi as caizhi,a.images as aimages,a.yimage as ayimage,c.title AS cpinlei,d.title AS dguige,a.xiangqian as axiangqian,a.lingxiao as alingxiao FROM xipunum_erp_ksiamges as a INNER JOIN xipunum_erp_category AS c ON c.id = a.category_id INNER JOIN xipunum_erp_specs AS d ON d.id = a.specification_id where a.id='{款号ID}' LIMIT 1
  ```
- **涉及表**: `xipunum_erp_ksiamges`、`xipunum_erp_category`、`xipunum_erp_specs`

**批量打印导入编码 (窗口_批量打印导入编码)**:
- **功能**: 批量导入打印商品编码
- **程序集变量（1个）**: `副编码_超级列表框组件句柄`
- **子程序列表（4个）**: 窗口初始化、窗口尺寸被改变、关闭窗口、重置、保存
- **涉及表**: `xipunum_erp_shop`

**商品信息添加款号 (窗口_商品信息添加款号)**:
- **功能**: 商品信息添加时的款号快速创建窗口，支持品类→规格→材质联动、图片上传、自动取图
- **程序集变量（5个）**: `查找品类id信息`、`局部_数据品类名称`、`局部_数据规格名称`、`局部_图片路径`、`局部_图片响应`
- **子程序列表（10个）**:
  | 子程序名 | 功能 |
  |----------|------|
  | `_窗口_商品信息添加款号_创建完毕` | 窗口初始化：加载品类下拉框、继承商品添加窗口的品类选择 |
  | `_窗口_商品信息添加款号_尺寸被改变` | 窗口自适应布局 |
  | `_图片框EX4_鼠标左键单击` | 关闭窗口 |
  | `_按钮EX_重置_鼠标左键单击` | 重置表单 |
  | `_按钮EX_图片上传_鼠标左键单击` | 上传款号图片（HTTP multipart上传） |
  | `_按钮EX_确定_鼠标左键单击` | 确认按钮：调用款号添加逻辑 |
  | `_自动取图_选择框_选中状态改变` | 开启/关闭自动取图（3秒轮询xipunum_erp_iamges表） |
  | `_时钟1_周期事件` | 自动取图轮询：查询未处理图片→下载→显示 |
  | `_商品组合框_初始化` | 加载品类下拉框 |
  | `_组合框_品类名称_内容被改变` | 品类选择联动：加载材质列表（从category.caizhiid字段解析）和规格下拉框 |
  | `_款号管理添加_被点击` | 款号添加核心逻辑：验证→生成款号编码→INSERT→UPDATE款号编码→下载图片→回填商品添加窗口 |
- **核心SQL**:
  ```sql
  -- 加载品类列表
  SELECT * FROM xipunum_erp_category where 1=1 order by id ASC

  -- 查询品类材质
  SELECT * FROM xipunum_erp_category where id='{品类ID}' order by id ASC

  -- 加载规格列表
  SELECT * FROM xipunum_erp_specs where category_id='{品类ID}' order by id ASC

  -- 查询品类简写
  SELECT jianxie FROM xipunum_erp_category where id='{品类ID}' LIMIT 1

  -- 查询规格简写
  SELECT jianxie FROM xipunum_erp_specs where id='{规格ID}' LIMIT 1

  -- 添加款号
  增加记录("xipunum_erp_ksiamges", title, kuanhao, category_id, specification_id, caizhi, images, yimage, xiangqian, lingxiao, status, cjuser, creationtime)

  -- 查询款号ID和图片
  SELECT id,yimage FROM xipunum_erp_ksiamges where kuanhao='{临时编码}' order by id ASC LIMIT 1

  -- 更新款号编码（正式编码=授权简写+品类简写+规格简写+6位ID）
  UPDATE xipunum_erp_ksiamges SET kuanhao='{正式编码}' WHERE id='{ID}' LIMIT 1

  -- 自动取图：查询未处理图片
  SELECT id,images FROM xipunum_erp_iamges WHERE cjuser='{用户}' and state='0' ORDER BY id desc LIMIT 1

  -- 自动取图：标记图片已处理
  UPDATE xipunum_erp_iamges SET state='1' WHERE id='{图片ID}'

  -- 批量标记图片已处理
  UPDATE xipunum_erp_iamges SET state='1' WHERE cjuser='{用户}' and state='0'

  -- 添加系统日志
  增加记录("xipunum_erp_xitong_log", type='添加', title='添加款号', conter, user, creationtime)
  ```
- **款号编码规则**: `{授权信息简写}{品类简写}{规格简写}{6位ID补齐}`（如XIPUHPGG000001）
- **涉及表**: `xipunum_erp_ksiamges`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_iamges`、`xipunum_erp_xitong_log`

**历史追溯** — 商品历史追溯查询（详见7.7.4）

**预警管理 (窗口_预警管理)**:
- **功能**: 库存预警管理，支持滞销判断和动态调整
- **程序集变量（3个）**:
  | 变量名 | 类型 | 说明 |
  |--------|------|------|
  | `集_行号1` | 整数型 | 表格当前行号 |
  | `集_列号1` | 整数型 | 表格当前列号 |
  | `删除按钮` | 表格按钮 | 表格删除按钮 |
- **子程序列表（15个）**:
  | 子程序名 | 功能 |
  |----------|------|
  | `_窗口_预警管理_创建完毕` | 窗口初始化：根据权限设置界面 |
  | `_窗口_预警管理_尺寸被改变` | 窗口自适应布局 |
  | `子程序_删除表格` | 清空表格 |
  | `高级表格1_加载表头` | 加载表格表头（10-11列：序号/款号/款式名称/品类/规格/材质/警戒值/报警值/库房/id/操作） |
  | `高级表格1_加载数据` | 加载预警规则数据 |
  | `_高级表格1_光标位置改变` | 表格光标位置改变事件 |
  | `_高级表格1_按钮被点击` | 表格按钮点击事件：删除预警规则 |
  | `_按钮_查询_被单击` | 查询预警规则 |
  | `_按钮_修改_被单击` | 修改预警规则（权限：72预警修改） |
  | `_按钮_重置_被单击` | 重置查询条件 |
  | `预警管理_其它参数` | 加载滞销判断和动态调整参数 |
- **核心SQL**:
  ```sql
  -- 查询预警规则
  SELECT a.id AS id,b.kuanhao AS kuanhao,b.product_name AS product_name,c.title AS pinlei,d.title AS guige,b.caizhi AS caizhi,a.warn_value AS warn_value,a.alarm_value AS alarm_value,e.title AS kufang
  FROM xipunum_erp_warning AS a
  INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code
  LEFT JOIN xipunum_erp_category AS c ON c.id = b.category_id
  LEFT JOIN xipunum_erp_specs AS d ON d.id = b.specification_id
  LEFT JOIN xipunum_erp_type AS e ON e.id = a.kufang
  WHERE b.product_name LIKE '%{关键字}%' OR b.item_number LIKE '%{关键字}%'
  ORDER BY a.id DESC

  -- 修改预警规则
  UPDATE xipunum_erp_warning SET warn_value='{警戒值}',alarm_value='{报警值}' WHERE id='{ID}' LIMIT 1

  -- 删除预警规则
  DELETE FROM xipunum_erp_warning WHERE id='{ID}'
  ```
- **涉及表**: `xipunum_erp_warning`、`xipunum_erp_shop`、`xipunum_erp_category`、`xipunum_erp_specs`、`xipunum_erp_type`、`xipunum_erp_xitong_log`
- **权限编号**: 72预警修改、72预警删除

**入库审核日志 (窗口_入库审核日志)**:
- **功能**: 入库订单审核日志查看
- **程序集变量（3个）**:
  | 变量名 | 类型 | 说明 |
  |--------|------|------|
  | `局部_主窗口行号` | 整数型 | 主窗口选中行号 |
  | `局部_订单id` | 文本型 | 入库订单ID |
  | `局部_订单编码` | 文本型 | 入库订单编码 |
- **子程序列表（4个）**:
  | 子程序名 | 功能 |
  |----------|------|
  | `_窗口_入库审核日志_创建完毕` | 窗口初始化：加载入库订单日志 |
  | `_窗口_入库审核日志_尺寸被改变` | 窗口自适应布局 |
  | `子程序_加载数据` | 加载审核日志数据 |
  | `子程序_刷新日志` | 刷新日志数据 |
- **核心SQL**:
  ```sql
  -- 查询入库审核日志
  SELECT a.conter AS conter,a.user AS user,COALESCE(b.name,'') AS name,a.creationtime AS creationtime
  FROM xipunum_erp_store_log AS a
  LEFT JOIN xipunum_erp_user AS b ON b.user=a.user
  WHERE a.order_id='{订单ID}' ORDER BY a.creationtime DESC
  ```
- **涉及表**: `xipunum_erp_store_log`、`xipunum_erp_user`

**商品销售批量修改 (窗口_商品销售批量修改)**:
- **功能**: 批量修改销售单信息
- **子程序列表（10个）**:
  | 子程序名 | 功能 |
  |----------|------|
  | `_窗口_商品销售批量修改_创建完毕` | 窗口初始化：加载导购员列表 |
  | `_窗口_商品销售批量修改_尺寸被改变` | 窗口自适应布局 |
  | `_图片框EX4_鼠标左键单击` | 关闭窗口 |
  | `_按钮EX_重置_鼠标左键单击` | 重置表单 |
  | `_编辑框_附加费折扣_内容被改变` | 附加费折扣输入验证 |
  | `_编辑框_销售附加费_内容被改变` | 销售附加费输入验证 |
  | `_按钮EX_保存_鼠标左键单击` | 保存批量修改 |
- **核心SQL**:
  ```sql
  -- 加载导购员列表（全部权限）
  SELECT name FROM xipunum_erp_user where state='0' order by id ASC

  -- 加载导购员列表（店铺权限）
  SELECT name FROM xipunum_erp_user where department='{分组ID}' and state='0' order by id ASC

  -- 加载导购员列表（岗位权限）
  SELECT name FROM xipunum_erp_user WHERE user in {权限} and state='0' order by id ASC

  -- 批量更新销售单
  UPDATE xipunum_erp_outbound SET gold_price='{克价}',sales_cost='{销售工费}',sales_surcharge='{销售附加费}',shopping_guide='{导购员}' WHERE order_id='{订单ID}'
  ```
- **涉及表**: `xipunum_erp_outbound`、`xipunum_erp_user`

#### 7.5.7.1 商品退库备注修改 (窗口_商品退库备注修改)
**功能**: 修改退库订单备注

**子程序列表（4个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品退库备注修改_创建完毕` | 窗口初始化：加载退库单号和备注 |
| `_窗口_商品退库备注修改_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX1_鼠标左键单击` | 保存备注修改 |
| `_按钮EX2_鼠标左键单击` | 重置表单 |

**核心SQL**:
```sql
-- 更新退库备注
UPDATE xipunum_erp_tuiku_order SET remarks='{备注}',updatetime='{时间}' WHERE tuiku_umber='{单号}' LIMIT 1

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type='修改', title='修改退库备注', conter, user, creationtime)
```

**涉及表**: `xipunum_erp_tuiku_order`、`xipunum_erp_xitong_log`

---

#### 7.11.7.1 证书管理添加修改 (窗口_证书管理添加修改)
**功能**: 商品证书管理，支持添加/修改证书信息

**子程序列表（8个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_证书管理添加修改_创建完毕` | 窗口初始化：加载检测机构下拉框、根据ID设置添加/修改模式 |
| `_窗口_证书管理添加修改_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX1_鼠标左键单击` | 保存/修改证书信息 |
| `_按钮EX2_鼠标左键单击` | 重置表单 |

**核心SQL**:
```sql
-- 加载检测机构下拉框
SELECT id,name FROM xipunum_erp_zsjigou where 1=1 order by id ASC

-- 查询证书信息
SELECT * FROM xipunum_erp_zhengshu WHERE id='{ID}' LIMIT 1

-- 检查证书编码是否存在
SELECT * FROM xipunum_erp_zhengshu WHERE zsbianma='{编码}'

-- 添加证书
增加记录("xipunum_erp_zhengshu", poduct_code, jigouid, zsbianma, conclusion, zongzhong, xingzhuang, yanse, beizhu, chengben, xiaoshou, cxdizhi, xiaoshouid, cjuser, creationtime)

-- 修改证书
UPDATE xipunum_erp_zhengshu SET xiaoshou='{销售价}',cxdizhi='{查询地址}',beizhu='{备注}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_zhengshu`、`xipunum_erp_zsjigou`、`xipunum_erp_xitong_log`

---

#### 7.11.7.2 证书机构添加修改 (窗口_证书机构添加修改)
**功能**: 检测机构管理

**子程序列表（5个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_证书机构添加修改_创建完毕` | 窗口初始化：加载机构列表 |
| `_窗口_证书机构添加修改_尺寸被改变` | 窗口自适应布局 |
| `_图片框EX4_鼠标左键单击` | 关闭窗口 |
| `_按钮EX1_鼠标左键单击` | 保存/修改机构信息 |
| `_按钮EX2_鼠标左键单击` | 重置表单 |

**核心SQL**:
```sql
-- 查询机构列表
SELECT * FROM xipunum_erp_zsjigou WHERE name LIKE '%{关键字}%' ORDER BY id DESC

-- 检查机构名称是否存在
SELECT * FROM xipunum_erp_zsjigou WHERE name='{名称}'

-- 添加机构
增加记录("xipunum_erp_zsjigou", name, cjuser, creationtime)

-- 修改机构
UPDATE xipunum_erp_zsjigou SET name='{名称}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1

-- 添加系统日志
增加记录("xipunum_erp_xitong_log", type, title, conter, user, creationtime)
```

**涉及表**: `xipunum_erp_zsjigou`、`xipunum_erp_xitong_log`

**旧料管理 (窗口_旧料管理单据)**:
- **功能**: 旧料出入库管理，支持内部/外部两种模式，支持新增/编辑/审核/详情四种状态
- **程序集变量（5个）**:
  | 变量名 | 类型 | 说明 |
  |--------|------|------|
  | `内部_旧料订单编码` | 文本型 | 旧料订单编码 |
  | `内部_旧料订单状态` | 文本型 | 旧料订单状态（""/编辑/审核/详情） |
  | `内部_旧料订单机构` | 文本型 | 旧料订单机构 |
  | `集_行号` | 整数型 | 表格当前行号 |
  | `集_列号` | 整数型 | 表格当前列号 |
- **子程序列表（20+个）**:
  | 子程序名 | 功能 |
  |----------|------|
  | `_窗口_旧料管理单据_创建完毕` | 窗口初始化：根据状态设置界面 |
  | `_窗口_旧料管理单据_尺寸被改变` | 窗口自适应布局 |
  | `_高级表格1_加载表头` | 加载表格表头 |
  | `_高级表格1_加载数据` | 加载旧料库存数据 |
  | `_高级表格1_编辑数据` | 加载编辑数据 |
  | `_高级表格1_审核数据` | 加载审核数据 |
  | `_高级表格1_详情数据` | 加载详情数据 |
  | `_高级表格1_光标位置改变` | 表格光标位置改变事件 |
  | `_高级表格1_结束编辑` | 表格结束编辑事件 |
  | `_单选框_内部_被单击` | 切换内部模式 |
  | `_单选框_外部_被单击` | 切换外部模式 |
  | `_机构名称组合框_将弹出列表` | 机构名称下拉框弹出列表 |
  | `_按钮_保存_被单击` | 保存旧料单据 |
  | `_按钮_重置_被单击` | 重置表单 |
- **核心SQL**:
  ```sql
  -- 查询旧料库存
  SELECT a.product_name AS mingcheng,ROUND(SUM(a.jin_zhong)-COALESCE(b.zhongliang,0),3) AS jinzhong
  FROM xipunum_erp_retreat AS a
  LEFT JOIN (SELECT product_name,SUM(zhongliang) as zhongliang FROM xipunum_erp_material WHERE kufang='{库房}' GROUP BY product_name) as b ON b.product_name=a.product_name
  WHERE a.shopping_guide IN ({权限范围})
  GROUP BY a.product_name HAVING jinzhong > 0
  ORDER BY SUM(a.jin_zhong) DESC

  -- 查询旧料订单详情
  SELECT a.product_name AS mingcheng,a.zhongliang AS chuku,a.jieyu AS jieyu,c.title AS kufang
  FROM xipunum_erp_material AS a
  INNER JOIN xipunum_erp_material_order AS b ON a.order_id = b.id
  INNER JOIN xipunum_erp_type AS c ON c.id = a.kufang
  WHERE b.number = '{单据号}' ORDER BY a.id ASC

  -- 保存旧料订单
  增加记录("xipunum_erp_material_order", number, zhongliang, jieyu, factory, remarks, cjuser, creationtime)

  -- 保存旧料明细
  增加记录("xipunum_erp_material", order_id, product_name, zhongliang, jieyu, kufang, remarks, cjuser, creationtime)
  ```
- **涉及表**: `xipunum_erp_retreat`、`xipunum_erp_material`、`xipunum_erp_material_order`、`xipunum_erp_about`、`xipunum_erp_type`、`xipunum_erp_xitong_log`

**线上订单 (窗口_线上订单驳回)**:
- **功能**: 线上订单驳回处理
- **子程序列表**: 窗口初始化、订单信息加载、驳回操作
- **核心SQL**:
  ```sql
  -- 更新订单状态为驳回
  UPDATE xipunum_erp_online_order SET state='驳回',cjuser='{账户}',updatetime='{时间}' WHERE id='{ID}' LIMIT 1
  
  -- 更新关联销售记录
  UPDATE xipunum_erp_online SET state='驳回',updatetime='{时间}' WHERE id='{ID}'
  
  -- 更新退款临时表
  UPDATE xipunum_erp_outbound_tuikuan_tmp SET state='驳回',updatetime='{时间}' WHERE online_id='{单号}' and poduct_code='{编码}' and kufang='{库房}' LIMIT 1
  ```
- **涉及表**: `xipunum_erp_online_order`、`xipunum_erp_online`、`xipunum_erp_outbound_tuikuan_tmp`、`xipunum_erp_retreat_online`

**物资盘点 (窗口_物资盘点添加)**:
- **功能**: 库存盘点操作
- **子程序列表**: 窗口初始化、盘点数据录入、盘点差异计算、盘点保存
- **涉及表**: `xipunum_erp_shop_kucun`、`data/pandian.mdb`

**提示消息** — 消息提示窗口（详见7.4.16）

**日期框** — 日期选择器（详见7.4.17）

---

## 8. 权限体系

### 8.1 权限层级

```
超级管理员
  ├── 全部导航可视权限
  ├── 全部操作权限
  └── 全部店铺权限

普通用户
  ├── 岗位可视权限 (全局_岗位权限可视)
  │   └── 控制可访问的导航菜单
  ├── 岗位操作权限 (全局_岗位权限操作)
  │   └── 控制可执行的操作（如：15商品入库修改、16商品销售编辑、57批量导入、60批量打印等）
  └── 店铺权限 (全局_店铺权限操作)
      └── 控制可查看的店铺数据范围
```

### 8.2 查看权限级别

| 级别 | 说明 | 全局_全局查看生成规则 |
|------|------|----------------------|
| `全部` | 查看所有店铺数据 | `(SELECT z.user FROM xipunum_erp_user AS z WHERE z.department in ({店铺权限}))` |
| `店铺` | 查看本店铺数据 | `(SELECT z.user FROM xipunum_erp_user AS z WHERE z.department='{分组id}')` |
| `岗位` | 查看本岗位数据 | `(SELECT z.user FROM xipunum_erp_user AS z WHERE z.post='{岗位id}')` |
| `个人` | 仅查看个人数据 | `(SELECT z.user FROM xipunum_erp_user AS z WHERE z.user='{用户名}')` |

### 8.3 权限检查代码模式

```e
.如果真 (等于 (寻找文本 (全局_岗位权限操作, ",60批量打印,", , 假), -1))
    提示框Ex_添加消息 ("无权操作！", 9, 1000, , 假, )
    返回 ()
.如果真结束
```

### 8.4 操作权限编号（部分）

| 编号 | 操作名称 |
|------|----------|
| 15 | 商品入库修改 |
| 16 | 商品销售编辑 |
| 57 | 批量导入 |
| 60 | 批量打印 |

---

## 9. 编码工具函数

系统内置了完整的Base编码工具集（集_Base程序集），支持8种编码格式：

| 编码 | 编码函数 | 解码函数 | 字符表 |
|------|----------|----------|--------|
| Base16 | `Base16编码` | `Base16解码` | 0-9A-F |
| Base32 | `Base32编码` | `Base32解码` | A-Z2-7 |
| Base36 | `Base36编码` | `Base36解码` | 0-9A-Z |
| Base58 | `Base58编码` | `Base58解码` | 1-9A-HJ-NP-Za-km-z |
| Base62 | `Base62编码` | `Base62解码` | 0-9A-Za-z |
| Base64 | `Base64编码` | `Base64解码` | A-Za-z0-9+/ |
| Base85 | `Base85编码` | `Base85解码` | ASCII 33-117 |
| Base91 | `Base91编码` | `Base91解码` | A-Za-z0-9!-~(除"#$%&'()*+,-./:;<=>?@[\]^_`{|}~) |

通用接口:
- `Base编码(编码类型, 源数据)` — 统一编码接口
- `Base解码(编码类型, 编码文本)` — 统一解码接口

---

### 7.12 辅助查询窗口

#### 7.12.1 商品销售副编码查询 (窗口_商品销售副编码查询)
**功能**: 销售商品副编码查询，支持多选和批量选择

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `副编码_超级列表框组件句柄` | 整数型 | 副编码列表框组件句柄 |

**子程序列表（4个）**:
| 子程序名 | 功能 |
|----------|------|
| `_窗口_商品销售副编码查询_创建完毕` | 窗口初始化：查询副编码商品库存 |
| `_副编码_表头回调` | 表头全选/取消全选回调 |
| `_副编码_超级列表框EX_项目左键单击` | 副编码列表点击事件 |
| `_按钮EX_批量多选_鼠标左键单击` | 批量多选副编码商品 |

**表头列数**: 8列（复选框/序号/商品编码/副编码/库房/数量/金重/入库时间）

**核心SQL**:
```sql
-- 查询副编码商品库存
SELECT a.poduct_code AS apoduct_code,b.fu_code AS bfu_code,a.quantity AS kucun,a.jinzhong AS jinzhong,CASE WHEN a.kufang = '0' THEN '总库' ELSE c.title END AS kufang,b.creationtime AS acreationtime FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang WHERE a.kufang = '{库房ID}' AND (a.quantity > 0 or a.jinzhong >0) and (b.poduct_code = '{编码}' OR b.fu_code = '{编码}') ORDER BY a.id DESC
```

**涉及表**: `xipunum_erp_shop_kucun`、`xipunum_erp_shop`、`xipunum_erp_type`

---

#### 7.12.2 销售编辑副编码查询 (窗口_销售编辑副编码查询)
**功能**: 销售编辑副编码查询，支持多选

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `副编码_超级列表框组件句柄` | 整数型 | 副编码列表框组件句柄 |

**子程序列表（4个）**: 窗口初始化、表头回调、副编码列表点击、批量多选

**表头列数**: 8列（复选框/序号/商品编码/副编码/库房/数量/金重/入库时间）

**核心SQL**: 同商品销售副编码查询

**涉及表**: `xipunum_erp_shop_kucun`、`xipunum_erp_shop`、`xipunum_erp_type`

---

#### 7.12.3 商品调拨副编码查询 (窗口_商品调拨副编码查询)
**功能**: 调拨商品副编码查询，支持多选和批量选择

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `副编码_超级列表框组件句柄` | 整数型 | 副编码列表框组件句柄 |

**子程序列表（4个）**: 窗口初始化、表头回调、副编码列表点击、批量多选

**表头列数**: 8列（复选框/序号/商品编码/副编码/库房/数量/金重/入库时间）

**核心SQL**: 同商品销售副编码查询

**涉及表**: `xipunum_erp_shop_kucun`、`xipunum_erp_shop`、`xipunum_erp_type`

---

#### 7.12.4 商品退库副编码查询 (窗口_商品退库副编码查询)
**功能**: 退库商品副编码查询，支持多选和批量选择

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `副编码_超级列表框组件句柄` | 整数型 | 副编码列表框组件句柄 |

**子程序列表（4个）**: 窗口初始化、表头回调、副编码列表点击、批量多选

**表头列数**: 8列（复选框/序号/商品编码/副编码/库房/数量/金重/入库时间）

**核心SQL**: 同商品销售副编码查询

**涉及表**: `xipunum_erp_shop_kucun`、`xipunum_erp_shop`、`xipunum_erp_type`

---

#### 7.12.5 商品退货副编码查询 (窗口_商品退货副编码查询)
**功能**: 退货商品副编码查询，支持多选

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `副编码_超级列表框组件句柄` | 整数型 | 副编码列表框组件句柄 |

**子程序列表（4个）**: 窗口初始化、表头回调、副编码列表点击、批量多选

**表头列数**: 8列（复选框/序号/商品编码/副编码/库房/数量/金重/入库时间）

**核心SQL**: 同商品销售副编码查询

**涉及表**: `xipunum_erp_shop_kucun`、`xipunum_erp_shop`、`xipunum_erp_type`

---

#### 7.12.6 成品修改副编码查询 (窗口_成品修改副编码查询)
**功能**: 成品修改副编码查询，支持多选

**程序集变量（1个）**:
| 变量名 | 类型 | 说明 |
|--------|------|------|
| `副编码_超级列表框组件句柄` | 整数型 | 副编码列表框组件句柄 |

**子程序列表（4个）**: 窗口初始化、表头回调、副编码列表点击、批量多选

**表头列数**: 8列（复选框/序号/商品编码/副编码/库房/数量/金重/入库时间）

**核心SQL**: 同商品销售副编码查询

**涉及表**: `xipunum_erp_shop_kucun`、`xipunum_erp_shop`、`xipunum_erp_type`

---

## 附录A: 窗口文件清单

| 序号 | 窗口名称 | 功能模块 | 详细说明 |
|------|----------|----------|----------|
| 1 | _启动窗口 | 系统启动/登录 | 详见7.1 |
| 2 | 窗口_主窗口 | 主界面 | 详见7.2 |
| 3 | 窗口_个人信息 | 个人信息 | 详见7.11.1 |
| 4 | 窗口_密码修改 | 密码修改 | 详见7.11.1 |
| 5 | 窗口_打印机设置 | 打印机配置 | 详见7.4.9 |
| 6 | 窗口_谷歌验证绑定 | 谷歌验证绑定 | 详见7.4.10 |
| 7 | 窗口_谷歌验证码输入 | 谷歌验证码输入 | 详见7.4.11 |
| 8 | 窗口_谷歌验证重置 | 谷歌验证重置 | 详见7.4.12 |
| 9 | 窗口_授权码信息输入 | 授权码录入 | 详见7.4.13 |
| 10 | 窗口_在线更新 | 在线更新 | 详见7.4.14 |
| 11 | 窗口_模块更新 | 模块更新 | 详见7.4.15 |
| 12 | 窗口_提示消息 | 消息提示 | 详见7.4.16 |
| 13 | 窗口_日期框 | 日期选择 | 详见7.4.17 |
| 14 | 窗口_信息商品查询 | 商品查询 | 详见7.4.18 |
| 15 | 窗口_商品信息添加 | 商品添加 | 详见7.5.1 |
| 16 | 窗口_商品成品数据修改 | 成品修改 | 详见7.5.1 |
| 17 | 窗口_商品入库添加 | 入库添加 | 详见7.5.4 |
| 18 | 窗口_商品入库详情 | 入库详情 | 详见7.5.5 |
| 19 | 窗口_商品入库批量修改 | 入库批量修改 | 详见7.5.5 |
| 20 | 窗口_入库审核日志 | 入库审核日志 | 详见7.5.5 |
| 21 | 窗口_入库库房修改 | 入库库房修改 | 详见7.5.5 |
| 22 | 窗口_入库工厂修改 | 入库工厂修改 | 详见7.5.5 |
| 23 | 窗口_商品销售出库 | 销售出库 | 详见7.5.2 |
| 24 | 窗口_商品销售编辑 | 销售编辑 | 详见7.5.3 |
| 25 | 窗口_商品销售批量修改 | 销售批量修改 | 详见7.6.1 |
| 26 | 窗口_商品销售客退 | 销售客退 | 详见7.6.2 |
| 27 | 窗口_商品销售外部单据 | 外部单据 | 详见7.6.3 |
| 28 | 窗口_商品销售订单查询 | 订单查询 | 详见7.6.4 |
| 29 | 窗口_商品销售副编码查询 | 销售副编码查询 | 辅助查询窗口 |
| 30 | 窗口_销售编辑副编码查询 | 编辑副编码查询 | 辅助查询窗口 |
| 31 | 窗口_商品信息调拨 | 商品调拨 | 详见7.5.6 |
| 32 | 窗口_商品调拨副编码查询 | 调拨副编码查询 | 辅助查询窗口 |
| 33 | 窗口_商品信息退库 | 商品退库 | 详见7.5.7 |
| 34 | 窗口_商品退库副编码查询 | 退库副编码查询 | 辅助查询窗口 |
| 35 | 窗口_商品退库备注修改 | 退库备注修改 | 详见7.5.7 |
| 36 | 窗口_商品信息回收 | 商品回收 | 详见7.5.8 |
| 37 | 窗口_商品信息预售 | 商品预售 | 详见7.5.9 |
| 38 | 窗口_商品信息退货 | 商品退货 | 详见7.5.10 |
| 39 | 窗口_商品退货副编码查询 | 退货副编码查询 | 辅助查询窗口 |
| 40 | 窗口_成品修改副编码查询 | 成品副编码查询 | 辅助查询窗口 |
| 41 | 窗口_成品销售会员绑定 | 会员绑定 | 详见7.6.5 |
| 42 | 窗口_会员添加修改 | 会员添加/修改 | 详见7.8.1 |
| 43 | 窗口_会员信息合并 | 会员合并 | 详见7.8.2 |
| 44 | 窗口_会员列表排序 | 会员排序 | 详见7.8.3 |
| 45 | 窗口_会员回访添加 | 会员回访 | 详见7.8.4 |
| 46 | 窗口_会员订单消费数据 | 消费记录 | 详见7.8.5 |
| 47 | 窗口_会员订单充值数据 | 充值记录 | 详见7.8.6 |
| 48 | 窗口_会员订单预购数据 | 预购记录 | 详见7.8.7 |
| 49 | 窗口_实时库存查询 | 实时库存 | 详见7.7.1 |
| 50 | 窗口_历史库存数据 | 历史库存 | 详见7.7.2 |
| 51 | 窗口_历史库存明细 | 库存明细 | 详见7.7.3 |
| 52 | 窗口_历史追溯 | 历史追溯 | 详见7.7.4 |
| 53 | 窗口_批量打印 | 批量打印 | 详见7.11.7 |
| 54 | 窗口_批量打印导入编码 | 打印导入 | 详见7.11.7 |
| 55 | 窗口_批量导入无款 | 批量导入 | 详见7.11.7 |
| 56 | 窗口_款号管理 | 款号管理 | 详见7.11.7 |
| 57 | 窗口_款号合并 | 款号合并 | 详见7.11.7 |
| 58 | 窗口_款号管理匹配 | 款号匹配 | 详见7.11.7 |
| 59 | 窗口_款号管理图片预览 | 图片预览 | 详见7.11.7 |
| 60 | 窗口_品类属性管理 | 品类属性 | 详见7.4.1 |
| 61 | 窗口_商品品类管理 | 品类管理 | 详见7.4.2 |
| 62 | 窗口_商品规格管理 | 规格管理 | 详见7.4.3 |
| 63 | 窗口_商品结算方式 | 结算方式 | 详见7.4.4 |
| 64 | 窗口_商品来源管理 | 来源管理 | 详见7.4.5 |
| 65 | 窗口_商品查询报表 | 查询报表 | 详见7.10.1 |
| 66 | 窗口_商品销售报表 | 销售报表 | 详见7.10.1 |
| 67 | 窗口_商品入库报表 | 入库报表 | 详见7.10.1 |
| 68 | 窗口_商品退库报表 | 退库报表 | 详见7.10.1 |
| 69 | 窗口_商品调拨报表 | 调拨报表 | 详见7.10.1 |
| 70 | 窗口_商品回收报表 | 回收报表 | 详见7.10.1 |
| 71 | 窗口_销售查询报表 | 销售查询 | 详见7.10.2 |
| 72 | 窗口_销售查询简易报表 | 简易查询 | 详见7.10.2 |
| 73 | 窗口_销售详情报表 | 销售详情 | 详见7.10.2 |
| 74 | 窗口_报表月销售统计 | 月销售统计 | 详见7.10.2 |
| 75 | 窗口_报表月汇总销售统计月 | 月汇总统计 | 详见7.10.2 |
| 76 | 窗口_报表员工月销售统计 | 员工月统计 | 详见7.10.2 |
| 77 | 窗口_报表店铺收支报表 | 店铺收支 | 详见7.10.2 |
| 78 | 窗口_报表店铺收支凭证 | 收支凭证 | 详见7.10.2 |
| 79 | 窗口_报表商品报表汇总 | 商品汇总 | 详见7.10.2 |
| 80 | 窗口_报表收银汇总表 | 收银汇总 | 详见7.10.2 |
| 81 | 窗口_报表导购回收表 | 导购回收 | 详见7.10.2 |
| 82 | 窗口_报表对照报表 | 对照报表 | 详见7.10.2 |
| 83 | 窗口_报表员工绩效 | 员工绩效 | 详见7.10.2 |
| 84 | 窗口_运营员工业绩 | 运营业绩 | 详见7.10.3 |
| 85 | 窗口_账户添加修改 | 账户管理 | 详见7.11.1 |
| 86 | 窗口_绩效信息管理 | 绩效管理 | 详见7.9.5 |
| 87 | 窗口_收支管理信息 | 收支管理 | 详见7.9.1 |
| 88 | 窗口_收支名称管理 | 收支名称 | 详见7.4.6 |
| 89 | 窗口_收支卡号管理 | 卡号管理 | 详见7.4.7 |
| 90 | 窗口_店铺数据结算 | 店铺结算 | 详见7.9.4 |
| 91 | 窗口_回收名称管理 | 回收名称 | 详见7.4.8 |
| 92 | 窗口_旧料管理单据 | 旧料管理 | 详见7.11.7 |
| 93 | 窗口_预警管理 | 库存预警 | 详见7.11.7 |
| 94 | 窗口_证书管理添加修改 | 证书管理 | 详见7.11.7 |
| 95 | 窗口_证书机构添加修改 | 证书机构 | 详见7.11.7 |
| 96 | 窗口_登录日志 | 登录日志 | 详见7.11.5 |
| 97 | 窗口_结账结料 | 结账结料 | 详见7.9.4 |
| 98 | 窗口_款式数据汇总 | 款式汇总 | 详见7.10.3 |
| 99 | 窗口_款式数据汇总明细 | 款式明细 | 详见7.10.3 |
| 100 | 窗口_款式数据汇总图片预览 | 图片预览 | 辅助窗口 |
| 101 | 窗口_商品信息添加款号 | 款号添加 | 辅助窗口 |
| 102 | 窗口_销售编辑批量修改 | 批量修改 | 详见7.6.1 |
| 103 | 窗口_盘点入库导入编码 | 盘点导入 | 详见7.11.7 |
| 104 | 窗口_物资盘点添加 | 物资盘点 | 详见7.11.7 |
| 105 | 窗口_线上订单驳回 | 订单驳回 | 详见7.11.7 |
| 106 | 窗口_进度条 | 进度显示 | 辅助窗口 |

---

## 附录B: 数据库表完整清单

共 **55+** 个数据库表：

| 表名 | 说明 | 主要用途 |
|------|------|----------|
| erp_authorize | 授权信息表 | 存储系统授权信息 |
| erp_city | 省市区地址表 | 会员地址三级联动 |
| erp_config | 授权系统配置表 | 系统级配置参数 |
| erp_mysql | 数据库连接配置表 | 读写分离数据库连接配置 |
| erp_navigation | 导航菜单表 | 两级树形导航配置 |
| erp_navigation_type | 导航操作权限表 | 导航操作权限配置 |
| erp_voucher | 授权凭证表 | 模块版本管理 |
| xipunum_erp_about | 工厂/关联信息表 | 工厂/供应商信息管理 |
| xipunum_erp_category | 品类表 | 商品品类管理 |
| xipunum_erp_category_stat_config | 品类统计配置表 | 品类统计分组配置 |
| xipunum_erp_config | 系统配置表 | 业务级配置参数 |
| xipunum_erp_delivery | 结算明细表 | 工厂结账结料明细 |
| xipunum_erp_delivery_order | 结算订单表 | 工厂结账结料订单 |
| xipunum_erp_finance | 收支记录表 | 店铺收支记录管理 |
| xipunum_erp_finance_account | 收支账户表 | 银行卡号管理 |
| xipunum_erp_finance_title | 收支名称表 | 收支项目名称配置 |
| xipunum_erp_finance_yinhang | 银行名称表 | 开户行名称管理 |
| xipunum_erp_iamges | 图片上传表 | 商品图片存储 |
| xipunum_erp_ksiamges | 款号图片表 | 款号图片管理 |
| xipunum_erp_history | 操作历史表 | 操作历史记录 |
| xipunum_erp_material | 旧料明细表 | 旧料出入库明细 |
| xipunum_erp_material_order | 旧料订单表 | 旧料出入库订单 |
| xipunum_erp_member | 会员主表 | 会员信息管理 |
| xipunum_erp_member_cq | 会员存取记录表 | 会员存欠记录 |
| xipunum_erp_member_score_log | 会员积分记录表 | 会员积分变动记录 |
| xipunum_erp_online | 线上销售明细表 | 线上订单销售明细 |
| xipunum_erp_online_order | 线上订单表 | 线上订单管理 |
| xipunum_erp_outbound | 出库商品明细表 | 销售出库明细 |
| xipunum_erp_outbound_order | 出库订单表 | 销售订单管理 |
| xipunum_erp_outbound_tuikuan_tmp | 退款临时表 | 线上退款临时数据 |
| xipunum_erp_pay | 支付方式表 | 支付方式配置 |
| xipunum_erp_performance | 绩效配置表 | 员工绩效规则配置 |
| xipunum_erp_pring | 打印任务队列表 | 打印任务管理 |
| xipunum_erp_presale | 预售明细表 | 预售商品明细 |
| xipunum_erp_presale_order | 预售订单表 | 预售订单管理 |
| xipunum_erp_retreat | 回收明细表 | 旧料回收明细 |
| xipunum_erp_retreat_online | 线上回收表 | 线上回收记录 |
| xipunum_erp_retreat_order | 回收订单表 | 回收订单管理 |
| xipunum_erp_retreat_title | 回收名称表 | 回收名称分类配置 |
| xipunum_erp_return | 退货明细表 | 客户退货明细 |
| xipunum_erp_return_order | 退货订单表 | 退货订单管理 |
| xipunum_erp_role | 岗位权限表 | 岗位权限配置 |
| xipunum_erp_shop | 商品主信息表 | 商品基础信息 |
| xipunum_erp_shop_kucun | 商品库存表 | 实时库存数据 |
| xipunum_erp_shop_lincun | 临存表 | 临存商品数据 |
| xipunum_erp_shop_log | 商品操作日志表 | 商品操作记录 |
| xipunum_erp_shop_xiangqian | 商品镶嵌信息表 | 商品镶嵌信息 |
| xipunum_erp_shopys | 预售关联表 | 预售商品关联 |
| xipunum_erp_shoukuan | 收款记录表 | 销售收款记录 |
| xipunum_erp_specs | 规格表 | 商品规格管理 |
| xipunum_erp_store | 入库商品明细表 | 入库商品明细 |
| xipunum_erp_store_order | 入库订单表 | 入库订单管理 |
| xipunum_erp_transfer | 调拨明细表 | 调拨商品明细 |
| xipunum_erp_transfer_order | 调拨订单表 | 调拨订单管理 |
| xipunum_erp_tuiku | 退库明细表 | 退库商品明细 |
| xipunum_erp_tuiku_order | 退库订单表 | 退库订单管理 |
| xipunum_erp_type | 类型/库房/商铺表 | 通用类型配置（商铺/库房/结算方式/商品来源等） |
| xipunum_erp_user | 用户/员工表 | 用户账户管理 |
| xipunum_erp_user_log | 用户登录日志表 | 登录日志记录 |
| xipunum_erp_visit | 会员回访记录表 | 会员回访管理 |
| xipunum_erp_voucher | 凭证/单据样式表 | 打印凭证样式 |
| xipunum_erp_warning | 预警规则表 | 库存预警规则配置 |
| xipunum_erp_xitong_log | 系统操作日志表 | 系统操作日志 |
| xipunum_erp_zhengshu | 证书信息表 | 商品证书管理 |
| xipunum_erp_zsjigou | 检测机构表 | 检测机构管理 |

---

## 附录C: 操作权限编号完整清单

| 编号 | 操作名称 | 所属模块 |
|------|----------|----------|
| 15 | 商品入库修改 | 入库管理 |
| 16 | 商品销售编辑 | 销售管理 |
| 32 | 添加商品来源 | 基础设置 |
| 32 | 编辑商品来源 | 基础设置 |
| 33 | 添加结算方式 | 基础设置 |
| 33 | 编辑结算方式 | 基础设置 |
| 34 | 添加品类 | 基础设置 |
| 34 | 编辑品类 | 基础设置 |
| 35 | 添加规格 | 基础设置 |
| 35 | 编辑规格 | 基础设置 |
| 57 | 批量导入 | 商品管理 |
| 60 | 批量打印 | 商品管理 |
| 74 | 添加回收名称 | 基础设置 |
| 74 | 编辑回收名称 | 基础设置 |
| 78 | 品类属性编辑 | 基础设置 |
| 82 | 收支名称新增 | 财务管理 |
| 82 | 收支名称编辑 | 财务管理 |
| 83 | 卡号新增 | 财务管理 |
| 83 | 卡号编辑 | 财务管理 |

---

## 附录D: 本地Access数据库表

| 表名 | 说明 | 数据库文件 |
|------|------|------------|
| chuku | 临时出库表 | data/erpdata.mdb |
| biaoqian | 标签打印表 | data/erpdata.mdb |
| shujulist | 主窗口表格临时数据表 | data/erpdata.mdb |
| pandian | 盘点数据表 | data/pandian.mdb |

---

---

## 10. 开发架构与操作方法

### 10.1 系统架构概述

ERPV4珠宝管理系统采用**前后端分离**架构，前端使用易语言ExUI框架，后端使用MySQL数据库，通过读写分离实现高性能数据访问。

### 10.2 核心业务操作流程图

#### 10.2.1 商品入库流程
```
┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐
│  选择工厂    │───▶│  添加商品    │───▶│  填写信息    │───▶│  保存入库单  │
│  选择来源    │    │  扫码/手动   │    │  重量/工费   │    │  更新库存    │
│  选择库房    │    │  品类/规格   │    │  成本价计算  │    │  打印标签    │
└─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘
```

**操作步骤**:
1. 进入「成品管理」→「商品入库添加」
2. 选择工厂、来源、库房、结算方式
3. 点击「添加商品」打开商品信息添加窗口
4. 选择品类→材质→成色→规格
5. 输入重量，系统自动计算金重、成本价、销售价
6. 点击「保存」完成入库
7. 可选：打印标签、打印单据

#### 10.2.2 商品销售流程
```
┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐
│  输入商品    │───▶│  选择导购    │───▶│  收款结算    │───▶│  保存销售单  │
│  扫码/手动   │    │  批发/零售   │    │  多种支付    │    │  扣减库存    │
│  库存检查    │    │  会员绑定    │    │  优惠计算    │    │  打印单据    │
└─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘
```

**操作步骤**:
1. 进入「成品管理」→「商品销售出库」
2. 输入商品编码（支持扫码）
3. 系统自动加载商品信息和库存
4. 选择导购员、批零类型
5. 可选：添加回收旧料、绑定会员
6. 选择支付方式，输入金额
7. 点击「保存」完成销售
8. 可选：打印销售单据

#### 10.2.3 商品调拨流程
```
┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐
│  选择模式    │───▶│  输入商品    │───▶│  确认调拨    │───▶│  保存调拨单  │
│  调入/调出   │    │  扫码/手动   │    │  检查库存    │    │  双向更新    │
│  选择库房    │    │  库存验证    │    │  确认数量    │    │  打印标签    │
└─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘
```

**操作步骤**:
1. 进入「成品管理」→「商品信息调拨」
2. 选择调入/调出模式
3. 选择目标库房
4. 输入商品编码（支持扫码）
5. 系统自动检查库存和库房
6. 点击「保存」完成调拨
7. 可选：打印标签

#### 10.2.4 会员管理流程
```
┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐
│  添加会员    │───▶│  会员消费    │───▶│  存欠管理    │───▶│  回访维护    │
│  基本信息    │    │  绑定订单    │    │  存料/欠料   │    │  回访记录    │
│  地址信息    │    │  积分累计    │    │  存元/欠元   │    │  合并会员    │
└─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘
```

**操作步骤**:
1. 进入「会员管理」→「会员添加修改」
2. 输入会员基本信息（姓名、电话、生日、地址）
3. 系统自动生成会员编码和会员ID
4. 可选：设置初始存欠数据
5. 保存会员信息

#### 10.2.5 报表查询流程
```
┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐
│  选择报表    │───▶│  设置条件    │───▶│  执行查询    │───▶│  导出数据    │
│  报表类型    │    │  日期/店铺   │    │  查看数据    │    │  Excel导出  │
│  统计维度    │    │  品类/规格   │    │  汇总统计    │    │  打印报表    │
└─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘
```

**操作步骤**:
1. 进入「报表」模块
2. 选择报表类型（销售/入库/退库/调拨/回收等）
3. 设置查询条件（日期范围、店铺、品类、规格等）
4. 点击「查询」加载数据
5. 可选：切换视图（订单/明细/天/月/年）
6. 可选：导出Excel、打印报表

### 10.3 核心开发模式

#### 10.3.1 窗口开发模式
每个窗口遵循统一的开发模式：
1. **程序集变量定义** - 声明窗口级变量
2. **窗口创建完毕** - 初始化窗口组件、加载数据
3. **窗口尺寸被改变** - 响应式布局
4. **表格加载表头** - 定义表格列结构
5. **表格加载数据** - 从数据库加载数据
6. **表格结束编辑** - 处理用户输入、联动计算
7. **保存按钮事件** - 事务性数据保存

#### 10.3.2 数据库操作模式
```e
' 1. 执行SQL查询
执行SQL语句 (全_MySQL读取, 到文本 (编码转换 (到字节集 ("SELECT ..."), #编码_GBK, #编码_UTF_8, )))

' 2. 获取记录集
赋值 (集句柄, 取记录集 (全_MySQL读取))

' 3. 读取字段值
读字段值 (集句柄, "字段名", 变量名)

' 4. 释放记录集
释放记录集 (集句柄)

' 5. 写入数据
增加记录 (全_MySQL写入, "表名", 字段1, 字段2, ...)

' 6. 更新数据
执行SQL语句 (全_MySQL写入, "UPDATE 表名 SET 字段='{值}' WHERE id='{ID}' LIMIT 1")
```

#### 10.3.3 事务处理模式
```e
' 开始事务
执行SQL语句 (全_MySQL写入, "START TRANSACTION")

' 执行业务SQL
执行SQL语句 (全_MySQL写入, "INSERT INTO ...")
执行SQL语句 (全_MySQL写入, "UPDATE ...")

' 检查影响行数
执行SQL语句 (全_MySQL写入, "SELECT ROW_COUNT() AS rowcnt")
赋值 (影响行数集句柄, 取记录集 (全_MySQL写入))
读字段值 (影响行数集句柄, "rowcnt", 影响行数)
释放记录集 (影响行数集句柄)

' 根据结果提交或回滚
如果 (大于 (到数值 (影响行数), 0))
    执行SQL语句 (全_MySQL写入, "COMMIT")
否则
    执行SQL语句 (全_MySQL写入, "ROLLBACK")
结束
```

### 10.4 编码规范

#### 10.4.1 变量命名规范
- **全局变量**: `全_` 或 `全局_` 前缀
- **程序集变量**: `集_` 或 `局部_` 前缀
- **局部变量**: 无前缀，使用有意义的名称

#### 10.4.2 SQL编码规范
- 使用参数化查询防止SQL注入
- 字符串使用UTF-8编码转换
- 数值使用ROUND函数保留精度
- 使用COALESCE处理NULL值

#### 10.4.3 错误处理规范
```e
' 检查数据库连接
如果真 (等于 (全_MySQL读取, 0))
    提示框Ex_添加消息 ("数据库连接异常！", 9, 1000, , 假, )
    返回 ()
结束

' 检查记录集
赋值 (集句柄, 取记录集 (全_MySQL读取))
如果真 (等于 (集句柄, 0))
    返回 ()
结束
```

### 10.5 性能优化建议

1. **数据库连接池**: 使用读写分离，合理分配连接资源
2. **索引优化**: 为常用查询字段创建索引
3. **分页查询**: 大数据量使用分页加载
4. **缓存机制**: 本地Access缓存常用数据
5. **批量操作**: 使用事务批量处理减少数据库交互
6. **异步处理**: 使用多线程处理耗时操作

---

## 附录C: 操作权限编号完整清单

| 编号 | 操作名称 | 所属模块 |
|------|----------|----------|
| 15 | 商品入库修改 | 入库管理 |
| 16 | 商品销售编辑 | 销售管理 |
| 32 | 添加商品来源 | 基础设置 |
| 32 | 编辑商品来源 | 基础设置 |
| 33 | 添加结算方式 | 基础设置 |
| 33 | 编辑结算方式 | 基础设置 |
| 34 | 添加品类 | 基础设置 |
| 34 | 编辑品类 | 基础设置 |
| 35 | 添加规格 | 基础设置 |
| 35 | 编辑规格 | 基础设置 |
| 57 | 批量导入 | 商品管理 |
| 60 | 批量打印 | 商品管理 |
| 74 | 添加回收名称 | 基础设置 |
| 74 | 编辑回收名称 | 基础设置 |
| 78 | 品类属性编辑 | 基础设置 |
| 82 | 收支名称新增 | 财务管理 |
| 82 | 收支名称编辑 | 财务管理 |
| 83 | 卡号新增 | 财务管理 |
| 83 | 卡号编辑 | 财务管理 |

---

## 附录D: 本地Access数据库表

| 表名 | 说明 | 数据库文件 |
|------|------|------------|
| chuku | 临时出库表 | data/erpdata.mdb |
| biaoqian | 标签打印表 | data/erpdata.mdb |
| shujulist | 主窗口表格临时数据表 | data/erpdata.mdb |
| pandian | 盘点数据表 | data/pandian.mdb |

---

## 附录E: 常用SQL查询模板

### E.1 商品查询模板
```sql
-- 商品基础信息查询
SELECT a.poduct_code, a.product_name, a.caizhi, a.item_number, a.quantity, a.net_weight, a.single, a.cost_price, a.sales_price, b.title AS pinlei, c.title AS guige FROM xipunum_erp_shop AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id WHERE 1=1

-- 商品库存查询
SELECT a.poduct_code, b.product_name, a.quantity, a.jinzhong, c.title AS kufang_name FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_type AS c ON c.id = a.kufang WHERE a.kufang IN ({权限})
```

### E.2 销售查询模板
```sql
-- 销售订单查询
SELECT a.settlement_number, a.customer_code, a.pling, a.settlement, a.creationtime, b.NAME AS cjuser_name FROM xipunum_erp_outbound_order AS a INNER JOIN xipunum_erp_user AS b ON b.USER = a.cjuser WHERE a.creationtime >= '{开始日期}' AND a.creationtime < '{结束日期}'

-- 销售明细查询
SELECT a.poduct_code, b.product_name, a.quantity, a.net_weight, a.gold_price, a.settlement, a.shopping_guide FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE a.order_id = '{订单ID}'
```

### E.3 会员查询模板
```sql
-- 会员基础信息查询
SELECT a.customer_code, a.memberid, a.name, a.tel, a.shengri, a.dizhi, a.creationtime FROM xipunum_erp_member AS a WHERE a.name LIKE '%{关键字}%' OR a.tel LIKE '%{关键字}%'

-- 会员存欠查询
SELECT a.customer_code, a.cunqu, a.type, a.number, a.remarks, a.creationtime FROM xipunum_erp_member_cq AS a WHERE a.customer_code = '{会员编码}' ORDER BY a.creationtime DESC
```

### E.4 库存查询模板
```sql
-- 实时库存查询
SELECT a.poduct_code, b.product_name, b.caizhi, b.item_number, a.quantity, a.jinzhong, c.title AS kufang_name FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_type AS c ON c.id = a.kufang WHERE a.kufang IN ({权限}) AND (a.quantity > 0 OR a.jinzhong > 0)

-- 历史库存查询
SELECT a.poduct_code, b.product_name, a.quantity, a.jinzhong, a.creationtime FROM xipunum_erp_shop_log AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE a.creationtime >= '{开始日期}' AND a.creationtime < '{结束日期}'
```

---

> 文档更新时间: 2026-07-13
> 基于源代码自动分析生成，覆盖 106 个窗口文件、2250+ 个子程序、65+ 个数据库表、1050+ 个SQL查询语句
> 本文档为ERPV4珠宝管理系统提供完整的开发参考，包含系统架构图、操作流程图、窗口模块、函数、SQL查询、数据库表、表头信息等完整开发信息

---

## 已完成详细整理的窗口清单

以下窗口已完成详细的变量、函数和SQL查询语句整理：

### 启动与主窗口
| 窗口名称 | 变量数 | 函数数 | SQL数 |
|----------|--------|--------|-------|
| 启动窗口 | 1 | 11 | 12 |
| 主窗口 | 70+ | 30+ | 15+ |
| 首页自动任务 | - | 4 | 3 |

### 基础设置模块
| 窗口名称 | 变量数 | 函数数 | SQL数 |
|----------|--------|--------|-------|
| 品类属性管理 | 1 | 9 | 4 |
| 商品品类管理 | - | 14 | 6 |
| 商品规格管理 | - | 10 | 5 |
| 商品结算方式 | - | 8 | 4 |
| 商品来源管理 | - | 8 | 4 |
| 收支名称管理 | - | 10 | 5 |
| 收支卡号管理 | - | 8 | 6 |
| 回收名称管理 | - | 8 | 4 |
| 打印机设置 | - | 5 | - |
| 谷歌验证绑定 | 2 | 6 | 2 |
| 谷歌验证码输入 | 1 | 6 | 3 |
| 谷歌验证重置 | 2 | 6 | 2 |
| 授权码信息输入 | - | 5 | 1 |
| 在线更新 | 1 | 6 | - |
| 模块更新 | 2 | 3 | 1 |
| 提示消息 | 10 | 6 | - |
| 日期框 | 1 | 3 | - |
| 信息商品查询 | - | 7 | 3 |

### 商品管理模块
| 窗口名称 | 变量数 | 函数数 | SQL数 |
|----------|--------|--------|-------|
| 商品信息添加 | 9 | 52 | 12 |
| 商品信息添加款号 | 5 | 10 | 5 |
| 商品成品数据修改 | 19 | 30+ | 8+ |
| 商品销售出库 | 42 | 72 | 15 |
| 商品销售编辑 | 38 | 58 | 15 |
| 商品入库添加 | 11 | 21 | 12 |
| 商品入库详情 | 6 | 22 | 12 |
| 商品入库批量修改 | 2 | 6 | 1 |
| 入库库房修改 | 3 | 4 | 6 |
| 入库工厂修改 | 3 | 5 | 5 |
| 商品信息调拨 | 12 | 26 | 10 |
| 商品信息退库 | 10 | 18 | 10 |
| 商品退库备注修改 | - | 5 | 2 |
| 商品信息回收 | 5 | 17 | 8 |
| 商品信息预售 | 11 | 17 | 8 |
| 商品信息退货 | 9 | 14 | 9 |

### 销售管理模块
| 窗口名称 | 变量数 | 函数数 | SQL数 |
|----------|--------|--------|-------|
| 商品销售批量修改 | - | 10 | 4 |
| 销售编辑批量修改 | - | 8 | 3 |
| 商品销售客退 | 35 | 50+ | 8+ |
| 商品销售外部单据 | 3 | 6 | 3 |
| 商品销售订单查询 | 1 | 5 | 2 |
| 成品销售会员绑定 | 1 | 10 | 3 |

### 会员管理模块
| 窗口名称 | 变量数 | 函数数 | SQL数 |
|----------|--------|--------|-------|
| 会员添加修改 | 1 | 13 | 10 |
| 会员信息合并 | 3 | 12 | 10 |
| 会员列表排序 | - | 5 | - |
| 会员回访添加 | 1 | 7 | 3 |
| 会员订单消费数据 | 3 | 4 | 2 |
| 会员订单充值数据 | 3 | 3 | 2 |
| 会员订单预购数据 | 3 | 3 | 1 |

### 财务管理模块
| 窗口名称 | 变量数 | 函数数 | SQL数 |
|----------|--------|--------|-------|
| 收支管理信息 | 3 | 10 | 4 |
| 店铺数据结算 | 3 | 15 | 7 |
| 绩效信息管理 | 4 | 20 | 8 |

### 库存管理模块
| 窗口名称 | 变量数 | 函数数 | SQL数 |
|----------|--------|--------|-------|
| 实时库存查询 | 11 | 20 | 8 |
| 历史库存数据 | 12 | 20 | 4 |
| 历史库存明细 | 18 | 25 | 4 |
| 历史追溯 | - | 10 | 4 |

### 报表模块
| 窗口名称 | 变量数 | 函数数 | SQL数 |
|----------|--------|--------|-------|
| 商品销售报表 | 17 | 30+ | 10+ |
| 商品入库报表 | 16 | 20+ | 8+ |
| 商品退库报表 | 10+ | 15+ | 6+ |
| 商品调拨报表 | 10+ | 15+ | 6+ |
| 商品回收报表 | 10+ | 15+ | 6+ |
| 商品查询报表 | 19 | 20+ | 8+ |
| 销售查询报表 | 11 | 15+ | 6+ |
| 销售详情报表 | 12 | 15+ | 6+ |
| 报表月销售统计 | 11 | 15+ | 6+ |
| 报表月汇总销售统计月 | 11 | 15+ | 6+ |
| 报表员工月销售统计 | 15 | 20+ | 8+ |
| 报表店铺收支报表 | 11 | 15+ | 6+ |
| 报表员工绩效 | 5 | 15+ | 6+ |

### 系统管理模块
| 窗口名称 | 变量数 | 函数数 | SQL数 |
|----------|--------|--------|-------|
| 个人信息 | 1 | 10 | 2 |
| 账户添加修改 | 1 | 20 | 6 |
| 密码修改 | - | 5 | 3 |
| 登录日志 | 1 | 5 | 3 |
| 入库审核日志 | 3 | 4 | 1 |
| 预警管理 | 3 | 15 | 5 |
| 旧料管理单据 | 5 | 20+ | 10+ |
| 线上订单驳回 | - | 5 | 5 |
| 证书管理添加修改 | - | 8 | 6 |
| 证书机构添加修改 | - | 5 | 5 |

### 入库辅助模块
| 窗口名称 | 变量数 | 函数数 | SQL数 |
|----------|--------|--------|-------|
| 入库库房修改 | 3 | 4 | 6 |
| 入库工厂修改 | 3 | 5 | 5 |

### 退库辅助模块
| 窗口名称 | 变量数 | 函数数 | SQL数 |
|----------|--------|--------|-------|
| 商品退库备注修改 | - | 5 | 2 |

**总计**: 860+ 个程序集变量、2250+ 个子程序、1050+ 个SQL查询语句

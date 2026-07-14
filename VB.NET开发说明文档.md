# ERPV4 珠宝管理系统 - VB.NET完整开发说明文档

> 本文档为VB.NET重写的完整开发指南，按照本文档可一步一步完成整个系统的开发调试。
> 数据库使用现有MySQL数据库，无需新建。

---

## 目录

- [第一部分：项目准备](#第一部分项目准备)
- [第二部分：基础架构开发](#第二部分基础架构开发)
- [第三部分：主框架开发](#第三部分主框架开发)
- [第四部分：基础设置模块](#第四部分基础设置模块)
- [第五部分：商品管理模块](#第五部分商品管理模块)
- [第六部分：会员管理模块](#第六部分会员管理模块)
- [第七部分：库存管理模块](#第七部分库存管理模块)
- [第八部分：财务管理模块](#第八部分财务管理模块)
- [第九部分：报表模块](#第九部分报表模块)
- [第十部分：系统管理模块](#第十部分系统管理模块)
- [第十一部分：调试与发布](#第十一部分调试与发布)

---

# 第一部分：项目准备

---

## 1. 开发环境搭建

### 1.1 开发工具安装

**Visual Studio 2022**（推荐Community版）
- 下载地址：https://visualstudio.microsoft.com/
- 安装时选择工作负载：**.NET桌面开发**
- 确保包含：VB.NET语言支持、Windows窗体设计器

### 1.2 .NET 8.0 SDK 安装

- 下载.NET 8.0 SDK：https://dotnet.microsoft.com/download/dotnet/8.0
- 安装完成后验证：
```bash
dotnet --version
# 应显示 8.x.x
```

### 1.3 MySQL 8.0 数据库配置

**现有数据库信息**：
| 数据库 | 地址 | 端口 | 数据库名 | 用户名 | 密码 |
|--------|------|------|----------|--------|------|
| 授权库 | 127.0.0.1 | 3306 | erpshouquan | erpshouquan | erpshouquan |
| 业务库 | 127.0.0.1 | 3306 | dldata | dldata | yt19880925!@# |

**MySQL客户端**：`D:\Hws.com\HwsHostMaster\phpweb\mysql\bin\mysql.exe`

### 1.4 NuGet包依赖清单

在项目中安装以下NuGet包：

| 包名 | 版本 | 用途 |
|------|------|------|
| MySql.Data | 8.0.33 | MySQL数据库连接 |
| Newtonsoft.Json | 13.0.3 | JSON序列化 |
| Google.Authenticator | 3.2.0 | 谷歌两步验证 |
| EPPlus | 6.2.0 | Excel导出 |
| System.Data.OleDb | 8.0.0 | Access数据库连接 |

---

## 2. 数据库连接配置

### 2.1 连接字符串配置

在`App.config`中配置：

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <connectionStrings>
    <!-- 授权库 -->
    <add name="AuthDB" 
         connectionString="Server=127.0.0.1;Port=3306;Database=erpshouquan;Uid=erpshouquan;Pwd=erpshouquan;SslMode=None;" 
         providerName="MySql.Data.MySqlClient" />
    <!-- 业务库（读写） -->
    <add name="BusinessDB" 
         connectionString="Server=127.0.0.1;Port=3306;Database=dldata;Uid=dldata;Pwd=yt19880925!@#;SslMode=None;" 
         providerName="MySql.Data.MySqlClient" />
  </connectionStrings>
</configuration>
```

### 2.2 数据库表结构概览

**授权库表（erpshouquan）**：
| 表名 | 说明 |
|------|------|
| erp_authorize | 授权信息表 |
| erp_city | 省市区地址表 |
| erp_config | 授权系统配置表 |
| erp_mysql | 数据库连接配置表 |
| erp_navigation | 导航菜单表 |
| erp_navigation_type | 导航操作权限表 |
| erp_voucher | 授权凭证表 |

**业务库表（dldata）- 65+个**：

详见《ERP系统开发文档.md》附录B

---

## 3. 项目结构创建

### 3.1 创建VB.NET项目

```
1. 打开Visual Studio 2022
2. 创建新项目 → 选择"Windows窗体应用(.NET)"
3. 项目名称：VB_ERP
4. 框架：.NET 8.0-windows
5. 位置：E:\ERPV4\VB_ERP
```

### 3.2 目录结构

```
E:\ERPV4\VB_ERP\
├── VB_ERP.sln                    # 解决方案文件
├── VB_ERP.vbproj                 # 项目文件
├── App.config                    # 配置文件
├── My Project\
│   ├── AssemblyInfo.vb
│   ├── Resources.Designer.vb
│   └── Settings.Designer.vb
├── Modules\
│   ├── GlobalVariables.vb        # 全局变量模块
│   ├── DatabaseModule.vb         # 数据库连接模块
│   ├── UtilityModule.vb          # 工具函数模块
│   └── BaseEncoding.vb           # Base编码模块
├── Forms\
│   ├── LoginForm.vb              # 登录窗口
│   ├── MainForm.vb               # 主窗口
│   └── BusinessForms.vb          # 业务窗口集合
├── Data\
│   └── config.ini                # 配置文件
└── Images\
    └── ...                       # 图片资源
```

### 3.3 全局变量模块 (GlobalVariables.vb)

```vb
Module GlobalVariables
    ' ========== 数据库连接变量 ==========
    Public MySQL_Auth As MySql.Data.MySqlClient.MySqlConnection      ' 授权库连接
    Public MySQL_Write As MySql.Data.MySqlClient.MySqlConnection     ' 业务写库连接
    Public MySQL_Read As MySql.Data.MySqlClient.MySqlConnection      ' 业务读库连接
    Public MySQL_ReadReport As MySql.Data.MySqlClient.MySqlConnection ' 报表读连接
    Public MySQL_ReadTask As MySql.Data.MySqlClient.MySqlConnection   ' 任务读连接
    Public MySQL_ReadPrint As MySql.Data.MySqlClient.MySqlConnection  ' 打印读连接
    Public MySQL_ReadOrder As MySql.Data.MySqlClient.MySqlConnection  ' 订单读连接
    
    ' ========== 授权信息变量 ==========
    Public AuthCode As String = ""           ' 授权码
    Public AuthCompany As String = ""        ' 授权公司
    Public AuthContact As String = ""        ' 联系人
    Public AuthPhone As String = ""          ' 联系电话
    Public AuthStartDate As String = ""      ' 开始日期
    Public AuthEndDate As String = ""        ' 结束日期
    Public AuthShortCode As String = ""      ' 授权简写
    
    ' ========== 用户信息变量 ==========
    Public UserAccount As String = ""        ' 用户账户
    Public UserName As String = ""           ' 用户姓名
    Public UserDepartment As String = ""     ' 所属分组ID
    Public UserPost As String = ""           ' 所属岗位ID
    Public UserPermission As String = ""     ' 数据权限
    Public UserRole As String = ""           ' 岗位权限可视
    Public UserOperation As String = ""      ' 岗位权限操作
    Public UserShopPermission As String = "" ' 店铺权限
    
    ' ========== 系统配置变量 ==========
    Public CompanyName As String = ""        ' 公司名称
    Public CompanyLogo As String = ""        ' 公司LOGO
    Public UserAvatar As String = ""         ' 用户头像
    Public WindowIcon As String = ""         ' 窗口图标
    Public ReturnWarehouse As String = ""    ' 退库仓ID
    Public OnlinePayment As String = ""      ' 线上收款ID
    Public DiscountRate As String = ""       ' 优惠比例
    Public KuanHao识别地址 As String = ""    ' 款号识别地址
    
    ' ========== 品类配置变量 ==========
    Public CategoryList As New List(Of String)   ' 品类列表
    Public SpecList As New List(Of String)       ' 规格列表
    Public MaterialList As New List(Of String)   ' 材质列表
End Module
```

### 3.4 数据库连接模块 (DatabaseModule.vb)

```vb
Imports MySql.Data.MySqlClient
Imports System.IO

Module DatabaseModule
    ' 连接授权库
    Public Function ConnectAuthDB() As Boolean
        Try
            Dim connStr As String = "Server=127.0.0.1;Port=3306;Database=erpshouquan;Uid=erpshouquan;Pwd=erpshouquan;SslMode=None;"
            MySQL_Auth = New MySqlConnection(connStr)
            MySQL_Auth.Open()
            Return True
        Catch ex As Exception
            MessageBox.Show("授权库连接失败：" & ex.Message)
            Return False
        End Try
    End Function
    
    ' 连接业务库
    Public Function ConnectBusinessDB(server As String, port As String, database As String, 
                                       username As String, password As String) As Boolean
        Try
            Dim connStr As String = $"Server={server};Port={port};Database={database};Uid={username};Pwd={password};SslMode=None;"
            MySQL_Write = New MySqlConnection(connStr)
            MySQL_Write.Open()
            
            MySQL_Read = New MySqlConnection(connStr)
            MySQL_Read.Open()
            
            MySQL_ReadReport = New MySqlConnection(connStr)
            MySQL_ReadReport.Open()
            
            Return True
        Catch ex As Exception
            MessageBox.Show("业务库连接失败：" & ex.Message)
            Return False
        End Try
    End Function
    
    ' 执行SQL查询（返回DataTable）
    Public Function ExecuteQuery(sql As String, Optional conn As MySqlConnection = Nothing) As DataTable
        If conn Is Nothing Then conn = MySQL_Read
        Dim dt As New DataTable()
        Try
            Dim adapter As New MySqlDataAdapter(sql, conn)
            adapter.Fill(dt)
        Catch ex As Exception
            ' 记录日志
        End Try
        Return dt
    End Function
    
    ' 执行SQL命令（INSERT/UPDATE/DELETE）
    Public Function ExecuteCommand(sql As String, Optional conn As MySqlConnection = Nothing) As Boolean
        If conn Is Nothing Then conn = MySQL_Write
        Try
            Dim cmd As New MySqlCommand(sql, conn)
            cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            ' 记录日志
            Return False
        End Try
    End Function
    
    ' 开始事务
    Public Function BeginTransaction() As MySqlTransaction
        Return MySQL_Write.BeginTransaction()
    End Function
    
    ' 提交事务
    Public Sub CommitTransaction(trans As MySqlTransaction)
        trans.Commit()
    End Sub
    
    ' 回滚事务
    Public Sub RollbackTransaction(trans As MySqlTransaction)
        trans.Rollback()
    End Sub
    
    ' SQL文本安全处理（防注入）
    Public Function SafeSQL(text As String) As String
        Return text.Replace("'", "''")
    End Function
    
    ' UTF8转GBK
    Public Function UTF8ToGBK(text As String) As String
        Dim utf8 As Encoding = Encoding.UTF8
        Dim gbk As Encoding = Encoding.GetEncoding("GBK")
        Dim utf8Bytes As Byte() = utf8.GetBytes(text)
        Dim gbkBytes As Byte() = Encoding.Convert(utf8, gbk, utf8Bytes)
        Return gbk.GetString(gbkBytes)
    End Function
    
    ' GBK转UTF8
    Public Function GBKToUTF8(text As String) As String
        Dim utf8 As Encoding = Encoding.UTF8
        Dim gbk As Encoding = Encoding.GetEncoding("GBK")
        Dim gbkBytes As Byte() = gbk.GetBytes(text)
        Dim utf8Bytes As Byte() = Encoding.Convert(gbk, utf8, gbkBytes)
        Return utf8.GetString(utf8Bytes)
    End Function
End Module
```

---

# 第二部分：基础架构开发

---

## 4. 启动窗口开发

### 4.1 启动流程

```
1. 连接授权库 (erpshouquan)
2. 读取 config.ini 获取授权码
3. 查询 erp_authorize 验证授权
4. 查询 erp_mysql 获取业务库配置
5. 连接业务库（读/写/备）
6. 读取系统配置 (xipunum_erp_config)
7. 读取品类配置 (xipunum_erp_category_stat_config)
8. 检查版本更新 (erp_voucher)
9. 用户登录验证
10. 加载用户权限
11. 打开主窗口
```

### 4.2 核心SQL

```sql
-- 验证授权码
SELECT * FROM erp_authorize WHERE authorize='{授权码}' LIMIT 1

-- 获取数据库连接配置
SELECT * FROM erp_mysql WHERE authorizeid='{授权ID}' AND state='0'  -- 读库
SELECT * FROM erp_mysql WHERE authorizeid='{授权ID}' AND state='1'  -- 写库

-- 用户登录验证
SELECT * FROM xipunum_erp_user WHERE user='{账户}' AND password='{MD5密码}' LIMIT 1

-- 加载系统配置
SELECT * FROM xipunum_erp_config

-- 加载品类统计配置
SELECT * FROM xipunum_erp_category_stat_config
```

### 4.3 MD5密码加密

```vb
Public Function MD5Encrypt(password As String, salt As String) As String
    Dim input As String = password & salt
    Using md5 As Security.Cryptography.MD5 = Security.Cryptography.MD5.Create()
        Dim inputBytes As Byte() = Encoding.UTF8.GetBytes(input)
        Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)
        Return BitConverter.ToString(hashBytes).Replace("-", "").ToLower()
    End Using
End Function
```

---

## 5. 主窗口框架

### 5.1 主窗口布局

```
┌─────────────────────────────────────────────────────────┐
│  工具栏：[首页] [刷新] [查询] [添加] [修改] [删除] ...   │
├──────────────┬──────────────────────────────────────────┤
│              │                                          │
│   导航树     │        内容区域                           │
│              │                                          │
│  ▼ 基础设置  │   ┌────────────────────────────────┐    │
│    ├ 品类管理│   │                                │    │
│    ├ 规格管理│   │      DataGridView / 报表        │    │
│    └ ...     │   │                                │    │
│              │   └────────────────────────────────┘    │
│  ▼ 成品管理  │                                          │
│    ├ 商品列表│                                          │
│    └ ...     │                                          │
│              │                                          │
├──────────────┴──────────────────────────────────────────┤
│  分页：[首页] [上一页] [1][2][3]... [下一页] [尾页] 共N条 │
└─────────────────────────────────────────────────────────┘
```

### 5.2 导航树加载SQL

```sql
-- 加载一级导航
SELECT * FROM erp_navigation 
WHERE role IN ({岗位权限可视}) AND superior='0' AND state='0' 
ORDER BY sort ASC

-- 加载二级导航
SELECT * FROM erp_navigation 
WHERE role IN ({岗位权限可视}) AND superior='{一级ID}' AND state='0' 
ORDER BY sort ASC
```

### 5.3 导航点击路由

```vb
Private Sub NavigationTree_Click(sender As Object, e As TreeViewEventArgs)
    Dim nodeText As String = e.Node.Text
    
    Select Case nodeText
        Case "首页"
            ShowHomePage()
        Case "岗位分组"
            ShowRoleGroupPage()
        Case "系统设置", "商铺信息"
            ShowSystemSettings()
        Case "商品列表"
            ShowProductList()
        Case "商品入库"
            ShowInboundList()
        Case "商品销售"
            ShowSalesList()
        ' ... 其他路由
        Case Else
            ' 检查是否为弹窗类
            If IsPopupWindow(nodeText) Then
                LoadPopupWindow(nodeText)
            Else
                ' 表格类页面
                ShowTablePage(nodeText)
            End If
    End Select
End Sub
```

---

# 第三部分：基础设置模块

---

## 6. 品类与规格管理

### 6.1 品类属性管理

**窗口**: 窗口_品类属性管理
**功能**: 管理品类的属性配置

**核心SQL**:
```sql
-- 查询品类属性
SELECT * FROM xipunum_erp_category WHERE 1=1 ORDER BY id ASC

-- 添加品类
INSERT INTO xipunum_erp_category (title, jianxie, caizhiid, cjuser, creationtime) 
VALUES ('{名称}', '{简写}', '{材质列表}', '{账户}', '{时间}')

-- 修改品类
UPDATE xipunum_erp_category SET title='{名称}', caizhiid='{材质列表}' 
WHERE id='{ID}' LIMIT 1

-- 删除品类
DELETE FROM xipunum_erp_category WHERE id='{ID}'
```

### 6.2 商品品类管理

**窗口**: 窗口_商品品类管理
**功能**: 商品品类的增删改查

**核心SQL**:
```sql
-- 查询品类列表
SELECT * FROM xipunum_erp_category WHERE 1=1 ORDER BY id ASC

-- 查询品类统计配置
SELECT * FROM xipunum_erp_category_stat_config WHERE category_id='{品类ID}'
```

### 6.3 商品规格管理

**窗口**: 窗口_商品规格管理
**功能**: 按品类管理商品规格

**核心SQL**:
```sql
-- 查询规格列表
SELECT a.*, b.title AS category_name 
FROM xipunum_erp_specs AS a
LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id
WHERE 1=1 ORDER BY a.id ASC

-- 添加规格
INSERT INTO xipunum_erp_specs (title, category_id, jianxie, cjuser, creationtime) 
VALUES ('{名称}', '{品类ID}', '{简写}', '{账户}', '{时间}')
```

---

## 7. 通用配置管理

### 7.1 商品结算方式

**窗口**: 窗口_商品结算方式
**表**: xipunum_erp_type (type='结算方式')

### 7.2 商品来源管理

**窗口**: 窗口_商品来源管理
**表**: xipunum_erp_type (type='来源')

### 7.3 收支名称管理

**窗口**: 窗口_收支名称管理
**表**: xipunum_erp_finance_title

### 7.4 收支卡号管理

**窗口**: 窗口_收支卡号管理
**表**: xipunum_erp_finance_account

### 7.5 回收名称管理

**窗口**: 窗口_回收名称管理
**表**: xipunum_erp_retreat_title

---

# 第四部分：商品管理模块

---

## 8. 商品信息管理

### 8.1 商品信息添加

**窗口**: 窗口_商品信息添加
**功能**: 新商品信息录入，支持款号图片识别自动填充

**程序集变量**（9个）:
- 集_行号、集_列号、删除按钮
- 局部_图片路径、局部_图片响应
- 局部_数据品类名称、局部_数据规格名称
- 查找品类id信息、局部_款号识别数据文本

**核心SQL**:
```sql
-- 查询品类列表
SELECT * FROM xipunum_erp_category WHERE 1=1 ORDER BY id ASC

-- 查询规格列表
SELECT * FROM xipunum_erp_specs WHERE category_id='{品类ID}' ORDER BY id ASC

-- 查询款号信息（图片识别后）
SELECT * FROM xipunum_erp_ksiamges WHERE kuanhao='{款号}' LIMIT 1

-- 添加商品
INSERT INTO xipunum_erp_shop (poduct_code, product_name, category_id, specification_id, 
    caizhi, item_number, weight, net_weight, single, images, fu_code, cjuser, creationtime)
VALUES ('{编码}', '{名称}', '{品类ID}', '{规格ID}', '{材质}', '{款号}', 
    '{重量}', '{金重}', '{单件重}', '{图片}', '{副编码}', '{账户}', '{时间}')

-- 添加库存
INSERT INTO xipunum_erp_shop_kucun (poduct_code, quantity, jinzhong, kufang, cjuser, creationtime)
VALUES ('{编码}', '{数量}', '{金重}', '{库房}', '{账户}', '{时间}')
```

**商品编码生成规则**:
```
格式：{授权简写}{品类简写}{规格简写}{时间戳}{用户}
示例：XIPUHPGG20260713123456admin
```

### 8.2 商品成品数据修改

**窗口**: 窗口_商品成品数据修改
**功能**: 修改已添加商品的信息

**核心SQL**:
```sql
-- 查询商品详情
SELECT a.*, b.title AS category_name, c.title AS spec_name 
FROM xipunum_erp_shop AS a
LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id
LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id
WHERE a.poduct_code='{编码}'

-- 修改商品信息
UPDATE xipunum_erp_shop SET 
    product_name='{名称}', category_id='{品类ID}', specification_id='{规格ID}',
    caizhi='{材质}', weight='{重量}', net_weight='{金重}', single='{单件重}',
    updatetime='{时间}'
WHERE poduct_code='{编码}' LIMIT 1
```

---

## 9. 入库管理

### 9.1 商品入库添加

**窗口**: 窗口_商品入库添加
**功能**: 创建入库单，添加入库商品明细

**业务流程**:
```
1. 创建入库单号（RK+yyyyMMdd+4位序号）
2. 选择供应商、结算方式
3. 添加商品明细（支持批量扫描）
4. 填写入库信息（重量、数量、工费等）
5. 保存入库单
6. 更新库存
7. 记录操作日志
```

**核心SQL**:
```sql
-- 创建入库订单
INSERT INTO xipunum_erp_store_order (number, factory, settlement, remarks, cjuser, creationtime)
VALUES ('{单号}', '{工厂ID}', '{结算方式}', '{备注}', '{账户}', '{时间}')

-- 添加入库明细
INSERT INTO xipunum_erp_store (order_id, poduct_code, product_name, category_id, 
    specification_id, caizhi, quantity, weight, net_weight, cost_price, 
    cost_labor, cost_surcharge, reference_labor, sales_labor, sales_surcharge,
    cjuser, creationtime)
VALUES ('{订单ID}', '{编码}', '{名称}', '{品类ID}', '{规格ID}', '{材质}',
    '{数量}', '{重量}', '{金重}', '{成本价}', '{成本工费}', '{成本附加费}',
    '{参考工费}', '{销售工费}', '{销售附加费}', '{账户}', '{时间}')

-- 更新库存（增加）
UPDATE xipunum_erp_shop_kucun 
SET quantity = quantity + {数量}, jinzhong = jinzhong + {金重}
WHERE poduct_code='{编码}' AND kufang='{库房}'
```

### 9.2 商品入库详情

**窗口**: 窗口_商品入库详情
**功能**: 查看/编辑入库单详情

### 9.3 入库批量修改

**窗口**: 窗口_商品入库批量修改
**功能**: 批量修改入库商品信息

---

## 10. 销售管理

### 10.1 商品销售出库

**窗口**: 窗口_商品销售出库
**功能**: 创建销售单，销售出库操作

**程序集变量**（42个）: 完整列表见《ERP系统开发文档.md》

**业务流程**:
```
1. 创建销售单号（CK+yyyyMMdd+4位序号）
2. 选择客户（可选会员）
3. 扫描/输入商品编码
4. 自动加载商品信息和库存
5. 填写销售信息（克价、工费、附加费）
6. 计算销售金额
7. 选择支付方式
8. 保存销售单
9. 扣减库存
10. 记录操作日志
```

**核心SQL**:
```sql
-- 创建销售订单
INSERT INTO xipunum_erp_outbound_order (settlement_number, customer_code, pling, 
    kufang, cjuser, creationtime)
VALUES ('{单号}', '{会员编码}', '{批零}', '{库房}', '{账户}', '{时间}')

-- 添加销售明细
INSERT INTO xipunum_erp_outbound (order_id, poduct_code, product_name, category_id,
    specification_id, caizhi, quantity, weight, net_weight, gold_price,
    cost_labor, sales_labor, sales_surcharge, sales_amount, settlement,
    shopping_guide, kufang, cjuser, creationtime)
VALUES ('{订单ID}', '{编码}', '{名称}', '{品类ID}', '{规格ID}', '{材质}',
    '{数量}', '{重量}', '{金重}', '{克价}', '{成本工费}', '{销售工费}',
    '{销售附加费}', '{销售金额}', '{实收金额}', '{导购员}', '{库房}', '{账户}', '{时间}')

-- 扣减库存
UPDATE xipunum_erp_shop_kucun 
SET quantity = quantity - {数量}, jinzhong = jinzhong - {金重}
WHERE poduct_code='{编码}' AND kufang='{库房}'

-- 记录收款
INSERT INTO xipunum_erp_shoukuan (order_id, pay_type, amount, cjuser, creationtime)
VALUES ('{订单ID}', '{支付方式}', '{金额}', '{账户}', '{时间}')
```

**价格计算公式**:
```
销售单价 = (金重 × (克价 + 销售工费) + 销售附加费) / 数量
销售金额 = 金重 × (克价 + 销售工费) + 销售附加费
实收金额 = ROUND(销售金额)
```

### 10.2 商品销售编辑

**窗口**: 窗口_商品销售编辑
**功能**: 编辑已创建的销售单

### 10.3 商品销售批量修改

**窗口**: 窗口_商品销售批量修改
**功能**: 批量修改销售单信息

**核心SQL**:
```sql
-- 批量更新销售单（本地Access）
UPDATE chuku SET 
    销售克价='{克价}', 销售工费='{工费}', 销售附加费='{附加费}',
    实收金额='{实收金额}', 销售数量='{数量}', 销售重量='{重量}'
WHERE 商品编码='{编码}'
```

### 10.4 商品销售客退

**窗口**: 窗口_商品销售客退
**功能**: 客户退货处理

**核心SQL**:
```sql
-- 创建退货订单
INSERT INTO xipunum_erp_return_order (return_umber, customer_code, kufang, cjuser, creationtime)
VALUES ('{单号}', '{会员编码}', '{库房}', '{账户}', '{时间}')

-- 添加退货明细
INSERT INTO xipunum_erp_return (order_id, poduct_code, quantity, weight, net_weight,
    return_amount, cjuser, creationtime)
VALUES ('{订单ID}', '{编码}', '{数量}', '{重量}', '{金重}', '{退货金额}', '{账户}', '{时间}')

-- 恢复库存
UPDATE xipunum_erp_shop_kucun 
SET quantity = quantity + {数量}, jinzhong = jinzhong + {金重}
WHERE poduct_code='{编码}' AND kufang='{库房}'
```

---

## 11. 调拨与退库

### 11.1 商品信息调拨

**窗口**: 窗口_商品信息调拨
**功能**: 库房间商品调拨

**核心SQL**:
```sql
-- 创建调拨订单
INSERT INTO xipunum_erp_transfer_order (transfer_umber, ykufang, xkufang, cjuser, creationtime)
VALUES ('{单号}', '{源库房}', '{目标库房}', '{账户}', '{时间}')

-- 添加调拨明细
INSERT INTO xipunum_erp_transfer (order_id, poduct_code, quantity, weight, net_weight, cjuser, creationtime)
VALUES ('{订单ID}', '{编码}', '{数量}', '{重量}', '{金重}', '{账户}', '{时间}')

-- 源库房扣减库存
UPDATE xipunum_erp_shop_kucun 
SET quantity = quantity - {数量}, jinzhong = jinzhong - {金重}
WHERE poduct_code='{编码}' AND kufang='{源库房}'

-- 目标库房增加库存
UPDATE xipunum_erp_shop_kucun 
SET quantity = quantity + {数量}, jinzhong = jinzhong + {金重}
WHERE poduct_code='{编码}' AND kufang='{目标库房}'
```

### 11.2 商品信息退库

**窗口**: 窗口_商品信息退库
**功能**: 商品退库操作

---

## 12. 回收与退货

### 12.1 商品信息回收

**窗口**: 窗口_商品信息回收
**功能**: 旧料回收处理

**核心SQL**:
```sql
-- 创建回收订单
INSERT INTO xipunum_erp_retreat_order (retrea_umber, customer_code, kufang, cjuser, creationtime)
VALUES ('{单号}', '{会员编码}', '{库房}', '{账户}', '{时间}')

-- 添加回收明细
INSERT INTO xipunum_erp_retreat (order_id, product_name, quantity, jin_zhong, 
    retreat_amount, shopping_guide, cjuser, creationtime)
VALUES ('{订单ID}', '{名称}', '{数量}', '{金重}', '{回收金额}', '{导购员}', '{账户}', '{时间}')
```

### 12.2 商品信息预售

**窗口**: 窗口_商品信息预售
**功能**: 预售订单管理

### 12.3 商品信息退货

**窗口**: 窗口_商品信息退货
**功能**: 向供应商退货

---

# 第五部分：会员管理模块

---

## 13. 会员信息管理

### 13.1 会员添加修改

**窗口**: 窗口_会员添加修改
**功能**: 会员信息的添加和修改

**核心SQL**:
```sql
-- 查询会员
SELECT * FROM xipunum_erp_member WHERE customer_code='{编码}' OR tel='{电话}'

-- 添加会员
INSERT INTO xipunum_erp_member (customer_code, memberid, name, tel, dizhi, shengri, 
    sex, cjuser, creationtime)
VALUES ('{编码}', '{会员ID}', '{姓名}', '{电话}', '{地址}', '{生日}', 
    '{性别}', '{账户}', '{时间}')

-- 修改会员
UPDATE xipunum_erp_member SET name='{姓名}', tel='{电话}', dizhi='{地址}', 
    shengri='{生日}', updatetime='{时间}'
WHERE customer_code='{编码}' LIMIT 1
```

**会员编码生成规则**:
```
格式：HY + 6位自增序号
示例：HY000001
```

### 13.2 会员信息合并

**窗口**: 窗口_会员信息合并
**功能**: 合并重复会员数据

**业务流程**:
```
1. 选择主会员和被合并会员
2. 迁移积分记录
3. 迁移存取记录
4. 迁移预售订单
5. 迁移回收订单
6. 迁移销售订单
7. 删除被合并会员
```

### 13.3 会员回访添加

**窗口**: 窗口_会员回访添加
**功能**: 添加会员回访记录

**核心SQL**:
```sql
-- 添加回访记录
INSERT INTO xipunum_erp_visit (customer_code, returntitle, returnconter, returndata, 
    cjuser, creationtime)
VALUES ('{会员编码}', '{回访标题}', '{回访内容}', '{回访日期}', '{账户}', '{时间}')
```

---

## 14. 会员数据查询

### 14.1 会员订单消费数据

**窗口**: 窗口_会员订单消费数据
**功能**: 查看会员消费记录

**核心SQL**:
```sql
-- 查询消费记录
SELECT a.settlement_number, a.presale_number, b.xiao_amount, b.settlement,
    b.gold_price, b.net_weight, b.basic_cost, b.sales_cost, b.sales_surcharge,
    b.zhekou, b.creationtime, c.name AS daogou
FROM xipunum_erp_outbound_order AS a
LEFT JOIN xipunum_erp_outbound AS b ON b.order_id = a.id
LEFT JOIN xipunum_erp_user AS c ON c.user = b.shopping_guide
WHERE a.customer_code = '{会员编码}'
ORDER BY a.id DESC
```

### 14.2 会员订单充值数据

**窗口**: 窗口_会员订单充值数据
**功能**: 查看会员充值/存欠记录

**核心SQL**:
```sql
-- 查询存欠记录
SELECT a.cunqu, a.type, CAST(ROUND(a.number, 3) AS DECIMAL(10,3)) AS number,
    a.remarks, a.cjuser, a.creationtime, b.name
FROM xipunum_erp_member_cq AS a
INNER JOIN xipunum_erp_user AS b ON b.user = a.cjuser
WHERE a.customer_code = '{会员编码}'
ORDER BY a.id DESC
```

### 14.3 会员订单预购数据

**窗口**: 窗口_会员订单预购数据
**功能**: 查看会员预售记录

---

# 第六部分：库存管理模块

---

## 15. 库存查询

### 15.1 实时库存查询

**窗口**: 窗口_实时库存查询
**功能**: 查询当前实时库存数据

**程序集变量**（11个）:
- 店铺名称_超级列表框组件句柄
- 品类名称_超级列表框组件句柄
- 规格名称_超级列表框组件句柄
- 材质名称_超级列表框组件句柄
- 工厂名称_超级列表框组件句柄
- 报表数据_高级表格组件句柄
- 查找开始日期、查找结束日期
- 查找信息库房、查找信息规格、查找信息工厂

**核心SQL**:
```sql
-- 查询实时库存
SELECT a.poduct_code, a.product_name, a.caizhi, a.item_number, a.single,
    b.quantity, b.jinzhong, b.kufang,
    c.title AS category_name, d.title AS spec_name
FROM xipunum_erp_shop AS a
INNER JOIN xipunum_erp_shop_kucun AS b ON b.poduct_code = a.poduct_code
LEFT JOIN xipunum_erp_category AS c ON c.id = a.category_id
LEFT JOIN xipunum_erp_specs AS d ON d.id = a.specification_id
WHERE b.kufang IN ({库房权限}) AND (b.quantity > 0 OR b.jinzhong > 0)
    {品类条件} {规格条件} {材质条件}
ORDER BY a.id DESC
```

### 15.2 历史库存数据

**窗口**: 窗口_历史库存数据
**功能**: 查询历史库存快照

### 15.3 历史库存明细

**窗口**: 窗口_历史库存明细
**功能**: 查看库存变动明细

### 15.4 历史追溯

**窗口**: 窗口_历史追溯
**功能**: 商品历史追溯查询

---

## 16. 库存管理

### 16.1 预警管理

**窗口**: 窗口_预警管理
**功能**: 库存预警规则配置

**核心SQL**:
```sql
-- 查询预警规则
SELECT a.*, b.product_name, b.item_number, c.title AS category_name, 
    d.title AS spec_name, e.title AS kufang_name
FROM xipunum_erp_warning AS a
INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code
LEFT JOIN xipunum_erp_category AS c ON c.id = b.category_id
LEFT JOIN xipunum_erp_specs AS d ON d.id = b.specification_id
LEFT JOIN xipunum_erp_type AS e ON e.id = a.kufang
ORDER BY a.id DESC

-- 修改预警规则
UPDATE xipunum_erp_warning SET warn_value='{警戒值}', alarm_value='{报警值}' 
WHERE id='{ID}' LIMIT 1
```

### 16.2 物资盘点

**窗口**: 窗口_物资盘点添加
**功能**: 库存盘点操作

---

# 第七部分：财务管理模块

---

## 17. 收支管理

### 17.1 收支管理信息

**窗口**: 窗口_收支管理信息
**功能**: 店铺收支记录管理

**核心SQL**:
```sql
-- 查询收支记录
SELECT a.*, b.title AS pay_name, c.title AS kufang_name
FROM xipunum_erp_finance AS a
LEFT JOIN xipunum_erp_pay AS b ON b.id = a.pay_type
LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang
WHERE a.kufang IN ({库房权限}) {日期条件}
ORDER BY a.id DESC

-- 添加收支记录
INSERT INTO xipunum_erp_finance (type, title, amount, pay_type, kufang, 
    bank, images, remarks, cjuser, creationtime)
VALUES ('{类型}', '{名称}', '{金额}', '{支付方式}', '{库房}', 
    '{银行}', '{图片}', '{备注}', '{账户}', '{时间}')
```

---

## 18. 结算管理

### 18.1 店铺数据结算

**窗口**: 窗口_店铺数据结算
**功能**: 店铺结算操作，支持结料/结账两种模式

**核心SQL**:
```sql
-- 查询结算数据
SELECT a.kufang, SUM(a.quantity) AS quantity, SUM(a.net_weight) AS net_weight,
    SUM(a.settlement) AS settlement
FROM xipunum_erp_outbound AS a
WHERE a.kufang IN ({库房权限}) {日期条件}
GROUP BY a.kufang

-- 保存结算订单
INSERT INTO xipunum_erp_settlement_order (settlement_number, kufang, type, 
    total_weight, total_amount, remarks, cjuser, creationtime)
VALUES ('{单号}', '{库房}', '{类型}', '{总重}', '{总金额}', '{备注}', '{账户}', '{时间}')
```

### 18.2 结账结料

**窗口**: 窗口_结账结料
**功能**: 工厂结账结料管理

---

## 19. 绩效管理

### 19.1 绩效信息管理

**窗口**: 窗口_绩效信息管理
**功能**: 员工绩效规则配置

**核心SQL**:
```sql
-- 查询绩效配置
SELECT * FROM xipunum_erp_performance WHERE type_id='{岗位ID}' ORDER BY id ASC

-- 保存绩效配置
INSERT INTO xipunum_erp_performance (type_id, category_id, pl, jszd, jsfw, djs,
    min1, cs1, min2, cs2, min3, cs3, min4, cs4, min5, cs5, cjuser, creationtime)
VALUES ('{岗位ID}', '{品类ID}', '{批零}', '{计算字段}', '{计算范围}', '{档次数}',
    '{最小值1}', '{参数值1}', '{最小值2}', '{参数值2}', '{最小值3}', '{参数值3}',
    '{最小值4}', '{参数值4}', '{最小值5}', '{参数值5}', '{账户}', '{时间}')
```

---

# 第八部分：报表模块

---

## 20. 商品报表（6个）

### 20.1 商品销售报表

**窗口**: 窗口_商品销售报表
**功能**: 销售数据统计，支持多维度筛选

**视图模式**: 订单/明细/天/月/年/店铺/品类（7种视图）

**表头列定义**:

**订单视图（无订单编码时）- 29列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 出库单号 | 130 | 居中 |
| 2 | 销售时间 | 130 | 居中 |
| 3 | 客户编码 | 130 | 居中 |
| 4 | 客户姓名 | 75 | 居中 |
| 5 | 联系电话 | 120 | 居中 |
| 6 | 预售单号 | 130 | 居中 |
| 7 | 定金 | 75 | 右对齐 |
| 8 | 销售数量 | 75 | 右对齐 |
| 9 | 销售金重 | 75 | 右对齐 |
| 10 | 回收克重 | 75 | 右对齐 |
| 11 | 销售重量 | 75 | 右对齐 |
| 12 | 销售金额 | 75 | 右对齐 |
| 13 | 应收金额 | 75 | 右对齐 |
| 14 | 回收金额 | 75 | 右对齐 |
| 15 | 优惠金额 | 75 | 右对齐 |
| 16 | 实收金额 | 75 | 右对齐 |
| 17 | 成本工费 | 75 | 右对齐 |
| 18 | 成本附加费 | 75 | 右对齐 |
| 19 | 成本费用 | 75 | 右对齐 |
| 20 | 销售工费 | 75 | 右对齐 |
| 21 | 销售附加费 | 75 | 右对齐 |
| 22 | 工费利润 | 75 | 右对齐 |
| 23 | 业务员 | 75 | 居中 |
| 24 | 合计税额 | 75 | 居中 |
| 25 | 操作账户 | 130 | 居中 |
| 26 | 批零 | 75 | 居中 |
| 27 | 介绍人 | 75 | 居中 |
| 28 | 销售工厂 | 75 | 居中 |

**订单视图（有订单编码时）/ 明细视图 - 32列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 商品编码 | 100 | 居中 |
| 2 | 出库时间 | 130 | 居中 |
| 3 | 出库单号 | 130 | 居中 |
| 4 | 商品名称 | 100 | 居中 |
| 5 | 销售店铺 | 100 | 居中 |
| 6 | 工厂 | 100 | 居中 |
| 7 | 款号 | 120 | 居中 |
| 8 | 品类 | 60 | 居中 |
| 9 | 规格 | 140 | 居中 |
| 10 | 材质 | 120 | 居中 |
| 11 | 单件重 | 75 | 右对齐 |
| 12 | 数量 | 75 | 右对齐 |
| 13 | 金重 | 75 | 右对齐 |
| 14 | 重量 | 75 | 右对齐 |
| 15 | 销售克价 | 75 | 右对齐 |
| 16 | 销售金额 | 75 | 右对齐 |
| 17 | 应收金额 | 75 | 右对齐 |
| 18 | 成本工费 | 75 | 右对齐 |
| 19 | 成本附加费 | 75 | 右对齐 |
| 20 | 成本价 | 75 | 右对齐 |
| 21 | 销售工费 | 75 | 右对齐 |
| 22 | 销售附加费 | 75 | 右对齐 |
| 23 | 工费利润 | 75 | 右对齐 |
| 24 | 折扣 | 75 | 右对齐 |
| 25 | 圈口/长度 | 75 | 右对齐 |
| 26 | 成色 | 75 | 右对齐 |
| 27 | 单位 | 75 | 居中 |
| 28 | 导购员 | 75 | 居中 |
| 29 | 批零 | 75 | 居中 |
| 30 | 介绍人 | 75 | 居中 |
| 31 | 销售工厂 | 75 | 居中 |

**天/月/年视图 / 店铺视图 / 品类视图 - 19列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 日期/店铺名称 | 100 | 居中 |
| 2 | 销售总数 | 自适应 | 右对齐 |
| 3 | 销售金重 | 自适应 | 右对齐 |
| 4 | 销售重量 | 自适应 | 右对齐 |
| 5 | 成本工费 | 自适应 | 右对齐 |
| 6 | 成本附加费 | 自适应 | 右对齐 |
| 7 | 成本价 | 自适应 | 右对齐 |
| 8 | 销售工费 | 自适应 | 右对齐 |
| 9 | 销售附加费 | 自适应 | 右对齐 |
| 10 | 销售金额 | 自适应 | 右对齐 |
| 11 | 应收金额 | 自适应 | 右对齐 |
| 12 | 预收定金 | 自适应 | 右对齐 |
| 13 | 优惠金额 | 自适应 | 右对齐 |
| 14 | 实收金额 | 自适应 | 右对齐 |
| 15 | 回收金重 | 自适应 | 右对齐 |
| 16 | 回收金额 | 自适应 | 右对齐 |
| 17 | 实际金额 | 自适应 | 右对齐 |
| 18 | 工费利润 | 自适应 | 右对齐 |

**筛选条件**: 店铺/品类/规格/材质/工厂/批零/日期范围

**核心SQL**:
```sql
-- 订单视图查询
SELECT a.settlement_number, a.creationtime, a.pling, a.state,
    b.name AS customer_name, c.name AS daogou,
    SUM(d.quantity) AS quantity, SUM(d.net_weight) AS net_weight, 
    SUM(d.settlement) AS settlement
FROM xipunum_erp_outbound_order AS a
LEFT JOIN xipunum_erp_member AS b ON b.customer_code = a.customer_code
LEFT JOIN xipunum_erp_user AS c ON c.user = a.cjuser
INNER JOIN xipunum_erp_outbound AS d ON d.order_id = a.id
WHERE a.creationtime >= '{开始日期}' AND a.creationtime < '{结束日期}'
    AND d.kufang IN ({库房权限})
GROUP BY a.id
ORDER BY a.id DESC

-- 明细视图查询
SELECT a.*, b.product_name, b.poduct_code, b.caizhi, b.item_number,
    c.title AS category_name, d.title AS spec_name
FROM xipunum_erp_outbound AS a
INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code
LEFT JOIN xipunum_erp_category AS c ON c.id = b.category_id
LEFT JOIN xipunum_erp_specs AS d ON d.id = b.specification_id
WHERE a.creationtime >= '{开始日期}' AND a.creationtime < '{结束日期}'
    AND a.kufang IN ({库房权限})
ORDER BY a.id DESC

-- 天/月/年统计查询
SELECT DATE(a.creationtime) AS date,
    SUM(a.quantity) AS quantity, SUM(a.net_weight) AS net_weight,
    SUM(a.settlement) AS settlement
FROM xipunum_erp_outbound AS a
WHERE a.creationtime >= '{开始日期}' AND a.creationtime < '{结束日期}'
    AND a.kufang IN ({库房权限})
GROUP BY DATE(a.creationtime)
ORDER BY date DESC
```

### 20.2 商品入库报表

**窗口**: 窗口_商品入库报表
**功能**: 入库数据统计

**视图模式**: 订单/明细/天/月/年（5种视图）

**表头列定义**:

**订单视图（无订单编码时）- 20列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 入库单号 | 140 | 居中 |
| 2 | 入库时间 | 130 | 居中 |
| 3 | 是否镶嵌 | 100 | 居中 |
| 4 | 送货单号 | 120 | 居中 |
| 5 | 品类 | 60 | 居中 |
| 6 | 半成品 | 80 | 居中 |
| 7 | 工厂 | 140 | 居中 |
| 8 | 来源 | 50 | 居中 |
| 9 | 结算 | 50 | 居中 |
| 10 | 入库店铺 | 120 | 居中 |
| 11 | 入库数量 | 75 | 右对齐 |
| 12 | 入库克价 | 75 | 右对齐 |
| 13 | 入库金重 | 75 | 右对齐 |
| 14 | 入库总重 | 75 | 右对齐 |
| 15 | 成本金额 | 75 | 右对齐 |
| 16 | 操作账户 | 120 | 居中 |
| 17 | 审核时间 | 140 | 居中 |
| 18 | 订单状态 | 60 | 居中 |
| 19 | 备注 | 200 | 左对齐 |

**订单视图（有订单编码时）/ 明细视图 - 28列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 商品编码 | 120 | 居中 |
| 2 | 入库时间 | 130 | 居中 |
| 3 | 入库单号 | 140 | 居中 |
| 4 | 商品名称 | 100 | 居中 |
| 5 | 是否镶嵌 | 70 | 居中 |
| 6 | 款号 | 100 | 居中 |
| 7 | 工厂 | 75 | 居中 |
| 8 | 入库库房 | 65 | 居中 |
| 9 | 品类 | 75 | 居中 |
| 10 | 规格 | 75 | 居中 |
| 11 | 材质 | 75 | 居中 |
| 12 | 单件重 | 75 | 右对齐 |
| 13 | 数量 | 75 | 右对齐 |
| 14 | 金重 | 75 | 右对齐 |
| 15 | 重量 | 75 | 右对齐 |
| 16 | 成本工费 | 75 | 右对齐 |
| 17 | 成本附加费 | 75 | 右对齐 |
| 18 | 成本价 | 75 | 右对齐 |
| 19 | 参考工费 | 75 | 右对齐 |
| 20 | 圈口/长度 | 75 | 右对齐 |
| 21 | 面宽 | 75 | 右对齐 |
| 22 | 厚度 | 75 | 右对齐 |
| 23 | 工厂成色 | 75 | 右对齐 |
| 24 | 公司成色 | 75 | 右对齐 |
| 25 | 单位 | 75 | 居中 |
| 26 | 系数 | 75 | 居中 |
| 27 | 原料价 | 75 | 右对齐 |

**天/月/年视图 - 9列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 日期 | 110 | 居中 |
| 2 | 入库总数 | 自适应 | 右对齐 |
| 3 | 入库金重 | 自适应 | 右对齐 |
| 4 | 入库总重 | 自适应 | 右对齐 |
| 5 | 成本工费总额 | 自适应 | 右对齐 |
| 6 | 成本附加费总额 | 自适应 | 右对齐 |
| 7 | 成本总额 | 自适应 | 右对齐 |
| 8 | 参考工费总额 | 自适应 | 右对齐 |

### 20.3 商品退库报表

**窗口**: 窗口_商品退库报表
**功能**: 退库数据统计

**视图模式**: 订单/明细/天/月/年（5种视图）

**表头列定义**:

**订单视图（无订单编码时）- 13列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 退库单号 | 120 | 居中 |
| 2 | 退库时间 | 140 | 居中 |
| 3 | 数量 | 85 | 右对齐 |
| 4 | 金重 | 85 | 右对齐 |
| 5 | 总重 | 85 | 右对齐 |
| 6 | 成本工费 | 85 | 右对齐 |
| 7 | 成本附加费 | 85 | 右对齐 |
| 8 | 成本价 | 85 | 右对齐 |
| 9 | 状态 | 85 | 居中 |
| 10 | 库房 | 85 | 居中 |
| 11 | 备注 | 自适应 | 居中 |
| 12 | 操作账户 | 140 | 居中 |

**订单视图（有订单编码时）/ 明细视图 - 18列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 商品编码 | 140 | 居中 |
| 2 | 退库时间 | 130 | 居中 |
| 3 | 退库单号 | 130 | 居中 |
| 4 | 品类 | 100 | 居中 |
| 5 | 规格 | 120 | 居中 |
| 6 | 商品名称 | 60 | 居中 |
| 7 | 款号 | 140 | 居中 |
| 8 | 原库房 | 120 | 居中 |
| 9 | 新库房 | 75 | 居中 |
| 10 | 单件重 | 75 | 右对齐 |
| 11 | 退库数量 | 75 | 右对齐 |
| 12 | 退库金重 | 75 | 右对齐 |
| 13 | 退库重量 | 75 | 右对齐 |
| 14 | 成本工费 | 75 | 右对齐 |
| 15 | 成本附加费 | 75 | 右对齐 |
| 16 | 成本价 | 75 | 右对齐 |
| 17 | 备注 | 300 | 居中 |

**天/月/年视图 - 8列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 日期 | 100 | 居中 |
| 2 | 退库总数 | 自适应 | 右对齐 |
| 3 | 退库金重 | 自适应 | 右对齐 |
| 4 | 退库总重 | 自适应 | 右对齐 |
| 5 | 成本工费 | 自适应 | 右对齐 |
| 6 | 成本附加费 | 自适应 | 右对齐 |
| 7 | 成本价 | 自适应 | 右对齐 |

### 20.4 商品调拨报表

**窗口**: 窗口_商品调拨报表
**功能**: 调拨数据统计

**视图模式**: 订单/明细/天/月/年/店铺（6种视图）

**表头列定义**:

**日期汇总 - 订单视图（无订单编码时）- 15列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 调拨单号 | 120 | 居中 |
| 2 | 调拨时间 | 140 | 居中 |
| 3 | 类型 | 85 | 居中 |
| 4 | 原库房 | 85 | 居中 |
| 5 | 新库房 | 85 | 居中 |
| 6 | 数量 | 85 | 右对齐 |
| 7 | 金重 | 85 | 右对齐 |
| 8 | 总重 | 85 | 右对齐 |
| 9 | 成本工费 | 85 | 右对齐 |
| 10 | 成本附加费 | 85 | 右对齐 |
| 11 | 成本价 | 85 | 右对齐 |
| 12 | 销售附加费 | 85 | 右对齐 |
| 13 | 备注 | 自适应 | 居中 |
| 14 | 操作账户 | 140 | 居中 |

**日期汇总 - 订单视图（有订单编码时）/ 明细视图 - 22列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 商品编码 | 110 | 居中 |
| 2 | 原编码 | 110 | 居中 |
| 3 | 调拨时间 | 130 | 居中 |
| 4 | 调拨单号 | 130 | 居中 |
| 5 | 调拨类型 | 75 | 居中 |
| 6 | 商品名称 | 140 | 居中 |
| 7 | 品类 | 65 | 居中 |
| 8 | 规格 | 65 | 居中 |
| 9 | 款号 | 120 | 居中 |
| 10 | 调出库房 | 75 | 居中 |
| 11 | 调入库房 | 75 | 居中 |
| 12 | 单件重 | 75 | 右对齐 |
| 13 | 调拨数量 | 75 | 右对齐 |
| 14 | 调拨金重 | 75 | 右对齐 |
| 15 | 调拨重量 | 75 | 右对齐 |
| 16 | 每克工费 | 75 | 右对齐 |
| 17 | 成本附加费 | 75 | 右对齐 |
| 18 | 成本价 | 75 | 右对齐 |
| 19 | 销售附加费 | 75 | 右对齐 |
| 20 | 备注 | 300 | 左对齐 |
| 21 | 操作账户 | 120 | 居中 |

**日期汇总 - 天/月/年视图 - 15列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 日期 | 100 | 居中 |
| 2 | 调出数量 | 自适应 | 右对齐 |
| 3 | 调出金重 | 自适应 | 右对齐 |
| 4 | 调出重量 | 自适应 | 右对齐 |
| 5 | 调出工费 | 自适应 | 右对齐 |
| 6 | 调出附加费 | 自适应 | 右对齐 |
| 7 | 调出成本价 | 自适应 | 右对齐 |
| 8 | (分隔) | 5 | - |
| 9 | 调入数量 | 自适应 | 右对齐 |
| 10 | 调入金重 | 自适应 | 右对齐 |
| 11 | 调入数量 | 自适应 | 右对齐 |
| 12 | 调入工费 | 自适应 | 右对齐 |
| 13 | 调入附加费 | 自适应 | 右对齐 |
| 14 | 调入成本价 | 自适应 | 右对齐 |

**店铺汇总 - 调出/调入视图 - 9列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 调出店铺 | 100 | 居中 |
| 2 | 调入店铺 | 100 | 居中 |
| 3 | 数量/调入数量 | 自适应 | 右对齐 |
| 4 | 金重/调入金重 | 自适应 | 右对齐 |
| 5 | 重量/调入重量 | 自适应 | 右对齐 |
| 6 | 工费/调入工费 | 自适应 | 右对齐 |
| 7 | 附加费/调入附加费 | 自适应 | 右对齐 |
| 8 | 成本价/调入成本价 | 自适应 | 右对齐 |

### 20.5 商品回收报表

**窗口**: 窗口_商品回收报表
**功能**: 回收数据统计

**视图模式**: 订单/明细/天/月/年/品类汇总（6种视图）

**表头列定义**:

**订单视图（无订单编码时）- 17列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 回收单号 | 140 | 居中 |
| 2 | 回收时间 | 130 | 居中 |
| 3 | 客户编码 | 130 | 居中 |
| 4 | 客户姓名 | 60 | 居中 |
| 5 | 联系电话 | 100 | 居中 |
| 6 | 金重 | 100 | 右对齐 |
| 7 | 总重 | 100 | 右对齐 |
| 8 | 其他费用 | 100 | 右对齐 |
| 9 | 回收金额 | 100 | 右对齐 |
| 10 | 应付金额 | 100 | 右对齐 |
| 11 | 实付金额 | 100 | 右对齐 |
| 12 | 税点 | 100 | 右对齐 |
| 13 | 税率金额 | 100 | 右对齐 |
| 14 | 业务员 | 100 | 居中 |
| 15 | 备注 | 200 | 居中 |
| 16 | 操作账户 | 120 | 居中 |

**订单视图（有订单编码时）/ 明细视图 - 13列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 商品名称 | 100 | 居中 |
| 2 | 回收时间 | 130 | 居中 |
| 3 | 回收单号 | 140 | 居中 |
| 4 | 数量 | 100 | 右对齐 |
| 5 | 金重 | 100 | 右对齐 |
| 6 | 总重 | 100 | 右对齐 |
| 7 | 回收克价 | 100 | 右对齐 |
| 8 | 其他费用 | 100 | 右对齐 |
| 9 | 回收金额 | 100 | 右对齐 |
| 10 | 成色 | 100 | 居中 |
| 11 | 导购 | 100 | 居中 |
| 12 | 备注 | 300 | 居中 |

**天/月/年视图 - 4列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 日期 | 100 | 居中 |
| 2 | 回收金重 | 100 | 右对齐 |
| 3 | 回收金额 | 100 | 右对齐 |

**品类汇总视图**: 动态列（根据回收名称和品类动态生成）

### 20.6 商品查询报表

**窗口**: 窗口_商品查询报表
**功能**: 综合商品查询

**表头列定义 - 25列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 商品编码 | 100 | 居中 |
| 2 | 原始编码 | 100 | 居中 |
| 3 | 入库单号 | 130 | 居中 |
| 4 | 入库时间 | 130 | 居中 |
| 5 | 品类名称 | 80 | 居中 |
| 6 | 规格 | 80 | 居中 |
| 7 | 材质 | 65 | 居中 |
| 8 | 商品名称 | 200 | 居中 |
| 9 | 款号 | 120 | 居中 |
| 10 | 是否镶嵌 | 75 | 居中 |
| 11 | 工厂 | 100 | 居中 |
| 12 | 单件重 | 75 | 右对齐 |
| 13 | 库存数量 | 75 | 右对齐 |
| 14 | 库存金重 | 75 | 右对齐 |
| 15 | 库存总重 | 75 | 右对齐 |
| 16 | 库房 | 75 | 居中 |
| 17 | 成本价 | 75 | 右对齐 |
| 18 | 预售价 | 75 | 右对齐 |
| 19 | 成本工费 | 75 | 右对齐 |
| 20 | 成本附加费 | 75 | 右对齐 |
| 21 | 参考工费 | 75 | 右对齐 |
| 22 | 销售工费 | 75 | 右对齐 |
| 23 | 销售附加费 | 75 | 右对齐 |
| 24 | 原料克价 | 75 | 右对齐 |

**筛选条件**: 店铺/品类/规格/工厂/金重/总重/圈号/款号/名称

---

## 21. 运营报表（14个）

### 21.1 销售查询报表

**窗口**: 窗口_销售查询报表
**功能**: 销售查询

**表头列定义 - 20列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 店铺名称 | 80 | 居中 |
| 2 | 品类 | 75 | 居中 |
| 3 | 规格 | 75 | 居中 |
| 4 | 实际数量 | 自适应 | 右对齐 |
| 5 | 实际金重 | 100 | 右对齐 |
| 6 | 实际总重 | 100 | 右对齐 |
| 7 | 实际金额 | 115 | 右对齐 |
| 8 | (分隔) | 5 | - |
| 9 | 销售数量 | 自适应 | 右对齐 |
| 10 | 销售金重 | 100 | 右对齐 |
| 11 | 销售总重 | 100 | 右对齐 |
| 12 | 销售金额 | 115 | 右对齐 |
| 13 | 实销金额 | 115 | 右对齐 |
| 14 | (分隔) | 5 | - |
| 15 | 客退数量 | 自适应 | 右对齐 |
| 16 | 客退金重 | 自适应 | 右对齐 |
| 17 | 客退总重 | 自适应 | 右对齐 |
| 18 | 客退金额 | 自适应 | 右对齐 |
| 19 | 实退金额 | 自适应 | 右对齐 |

**筛选条件**: 店铺/材质/品类/规格/批零/日期

### 21.2 销售查询简易报表

**窗口**: 窗口_销售查询简易报表
**功能**: 简易销售查询

**表头列定义 - 12列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 65 | 居中 |
| 1 | 店铺名称 | 自适应 | 居中 |
| 2 | 黄金销售 | 自适应 | 右对齐 |
| 3 | 黄金回收 | 自适应 | 右对齐 |
| 4 | 黄金应结 | 自适应 | 右对齐 |
| 5 | 海峡金 | 自适应 | 右对齐 |
| 6 | 本月应结 | 自适应 | 右对齐 |
| 7 | 应结/多结 | 自适应 | 右对齐 |
| 8 | 本月实际应结 | 自适应 | 右对齐 |
| 9 | 本月已结 | 自适应 | 右对齐 |
| 10 | 累计未结 | 自适应 | 右对齐 |
| 11 | 备注 | 自适应 | 居中 |

### 21.3 销售详情报表

**窗口**: 窗口_销售详情报表
**功能**: 销售明细查询

**表头列定义 - 35列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 出库单号 | 150 | 居中 |
| 2 | 商品编码 | 100 | 居中 |
| 3 | 销售时间 | 100 | 居中 |
| 4 | 商品名称 | 140 | 居中 |
| 5 | 款号 | 100 | 居中 |
| 6 | 库房 | 100 | 居中 |
| 7 | 品类 | 100 | 居中 |
| 8 | 规格 | 100 | 居中 |
| 9 | 材质 | 100 | 居中 |
| 10 | 圈口/长度 | 100 | 居中 |
| 11 | 成色 | 100 | 居中 |
| 12 | 单件重 | 100 | 右对齐 |
| 13 | 金重 | 100 | 右对齐 |
| 14 | 重量 | 100 | 右对齐 |
| 15 | 单位 | 100 | 右对齐 |
| 16 | 料价 | 100 | 右对齐 |
| 17 | 成本工费 | 100 | 右对齐 |
| 18 | 参考工费 | 100 | 右对齐 |
| 19 | 成本附加费 | 100 | 右对齐 |
| 20 | 成本价 | 100 | 右对齐 |
| 21 | 销售单价 | 100 | 右对齐 |
| 22 | 销售金额 | 100 | 右对齐 |
| 23 | 数量 | 100 | 右对齐 |
| 24 | 原附加费 | 100 | 右对齐 |
| 25 | 销售克价 | 100 | 右对齐 |
| 26 | 销售工费 | 100 | 右对齐 |
| 27 | 销售附加费 | 100 | 右对齐 |
| 28 | 折扣 | 100 | 右对齐 |
| 29 | 应收金额 | 100 | 居中 |
| 30 | 工费利润 | 100 | 右对齐 |
| 31 | 成本工费合计 | 100 | 居中 |
| 32 | 销售工费合计 | 100 | 居中 |
| 33 | 批零 | 100 | 居中 |
| 34 | 状态 | 100 | 居中 |

**筛选条件**: 店铺/材质/批零/金重/日期

### 21.4 报表月销售统计

**窗口**: 窗口_报表月销售统计
**功能**: 月度销售统计

**表头列定义 - 23列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 65 | 居中 |
| 1 | 店铺名称 | 80 | 居中 |
| 2 | 品类 | 75 | 居中 |
| 3 | 批零 | 75 | 居中 |
| 4 | 实际数量 | 75 | 右对齐 |
| 5 | 实际金重 | 100 | 右对齐 |
| 6 | 实际总重 | 100 | 右对齐 |
| 7 | 实际金额 | 100 | 右对齐 |
| 8 | 工费利润 | 100 | 右对齐 |
| 9 | 成本工费 | 100 | 右对齐 |
| 10 | 销售工费 | 100 | 右对齐 |
| 11 | (分隔) | 5 | - |
| 12 | 销售数量 | 75 | 右对齐 |
| 13 | 销售金重 | 100 | 右对齐 |
| 14 | 销售总重 | 100 | 右对齐 |
| 15 | 销售金额 | 100 | 右对齐 |
| 16 | 实销金额 | 100 | 右对齐 |
| 17 | (分隔) | 5 | - |
| 18 | 客退数量 | 75 | 右对齐 |
| 19 | 客退金重 | 100 | 右对齐 |
| 20 | 客退总重 | 100 | 右对齐 |
| 21 | 客退金额 | 100 | 右对齐 |
| 22 | 实退金额 | 100 | 右对齐 |

**筛选条件**: 店铺/材质/品类/批零/日期

### 21.5 报表月汇总销售统计月

**窗口**: 窗口_报表月汇总销售统计月
**功能**: 月汇总统计

**视图模式**: 所有信息/重量/金额（3种视图）

**表头列定义 - 7列基础 + 动态日期列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 0/80 | 居中 |
| 1 | 店铺名称 | 80 | 居中 |
| 2 | 分类名称 | 75 | 居中 |
| 3 | 批零 | 75 | 居中 |
| 4 | 合计数量 | 75 | 右对齐 |
| 5 | 合计金重 | 75/150 | 右对齐 |
| 6 | 合计金额 | 120/150 | 右对齐 |

**动态列**: 根据日期范围动态生成每日3列（数量/金重/金额）

**视图模式说明**:
- 所有信息：显示所有列
- 重量：隐藏数量和金额列，只显示金重
- 金额：隐藏数量和金重列，只显示金额

### 21.6 报表员工月销售统计

**窗口**: 窗口_报表员工月销售统计
**功能**: 员工月销售统计

**视图模式**: 员工月销售统计/商品明细（2种视图）

**表头列定义**:

**员工月销售统计 - 重量模式 - 15列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 账户 | 100 | 居中 |
| 2 | 导购员 | 100 | 居中 |
| 3 | 品类 | 自适应 | 居中 |
| 4 | 销售数量 | 自适应 | 右对齐 |
| 5 | 销售金重 | 自适应 | 右对齐 |
| 6 | (分隔) | 5 | - |
| 7 | 退货数量 | 自适应 | 右对齐 |
| 8 | 退货金重 | 自适应 | 右对齐 |
| 9 | (分隔) | 5 | - |
| 10 | 实际数量 | 自适应 | 右对齐 |
| 11 | 实际金重 | 自适应 | 右对齐 |
| 12 | (分隔) | 5 | - |
| 13 | 所在店铺 | 100 | 居中 |
| 14 | 零批 | 65 | 居中 |

**员工月销售统计 - 金额模式 - 17列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 账户 | 100 | 居中 |
| 2 | 导购员 | 100 | 居中 |
| 3 | 品类 | 自适应 | 居中 |
| 4 | 销售数量 | 自适应 | 右对齐 |
| 5 | 销售金额 | 自适应 | 右对齐 |
| 6 | 实销金额 | 自适应 | 右对齐 |
| 7 | (分隔) | 5 | - |
| 8 | 退货数量 | 自适应 | 右对齐 |
| 9 | 退货金额 | 自适应 | 右对齐 |
| 10 | 实退金额 | 自适应 | 右对齐 |
| 11 | (分隔) | 5 | - |
| 12 | 实际数量 | 自适应 | 右对齐 |
| 13 | 实际金额 | 自适应 | 右对齐 |
| 14 | (分隔) | 5 | - |
| 15 | 所在店铺 | 100 | 居中 |
| 16 | 零批 | 65 | 居中 |

**商品明细视图 - 23列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 商品编码 | 100 | 居中 |
| 2 | 商品名称 | 140 | 居中 |
| 3 | 款号 | 70 | 居中 |
| 4 | 品类 | 70 | 居中 |
| 5 | 规格 | 70 | 居中 |
| 6 | 材质 | 70 | 居中 |
| 7 | 圈口/长度 | 70 | 居中 |
| 8 | 成色 | 70 | 居中 |
| 9 | 数量 | 70 | 右对齐 |
| 10 | 净重 | 70 | 右对齐 |
| 11 | 成本工费 | 70 | 右对齐 |
| 12 | 参考工费 | 70 | 右对齐 |
| 13 | 成本附加费 | 70 | 右对齐 |
| 14 | 销售单价 | 70 | 右对齐 |
| 15 | 销售金额 | 70 | 右对齐 |
| 16 | 原附加费 | 70 | 右对齐 |
| 17 | 销售克价 | 70 | 右对齐 |
| 18 | 销售工费 | 70 | 右对齐 |
| 19 | 销售附加费 | 70 | 右对齐 |
| 20 | 折扣 | 70 | 右对齐 |
| 21 | 实收金额 | 70 | 右对齐 |
| 22 | 导购员 | 70 | 居中 |

**筛选条件**: 店铺/品类/规格/工厂/批零/导购员

### 21.7 报表店铺收支报表

**窗口**: 窗口_报表店铺收支报表
**功能**: 店铺收支统计

**视图模式**: 明细/天/月/年/店铺（5种视图）

**表头列定义**:

**日期汇总 - 明细视图 - 12列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 日期 | 140 | 居中 |
| 2 | 所属店铺 | 110 | 居中 |
| 3 | 单据号 | 110 | 居中 |
| 4 | 类别 | 80 | 居中 |
| 5 | 开户行 | 130 | 居中 |
| 6 | 账户名称 | 120 | 居中 |
| 7 | 账号 | 140 | 居中 |
| 8 | 收支名称 | 130 | 居中 |
| 9 | 金额 | 90 | 右对齐 |
| 10 | 备注 | 100 | 居中 |
| 11 | 凭证 | 65 | 居中 |

**日期汇总 - 天/月/年视图 - 5列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 日期 | 100 | 居中 |
| 2 | 收入 | 自适应 | 右对齐 |
| 3 | 支出 | 自适应 | 右对齐 |
| 4 | 合计 | 自适应 | 右对齐 |

**店铺汇总 - 5列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 店铺名称 | 100 | 居中 |
| 2 | 收入 | 自适应 | 右对齐 |
| 3 | 支出 | 自适应 | 右对齐 |
| 4 | 合计 | 自适应 | 右对齐 |

**筛选条件**: 店铺/银行/收支名称/日期

### 21.8 报表店铺收支凭证

**窗口**: 窗口_报表店铺收支凭证
**功能**: 收支凭证查看

**核心SQL**:
```sql
-- 查询凭证图片
SELECT images FROM xipunum_erp_finance WHERE id='{凭证ID}' LIMIT 1
```

### 21.9 报表商品报表汇总

**窗口**: 窗口_报表商品报表汇总
**功能**: 商品报表汇总

**视图模式**: 日期/店铺（2种视图）

**表头列定义 - 31列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 日期/店铺名称 | 90/80 | 居中 |
| 2 | 销售数量 | 65 | 右对齐 |
| 3 | 线上数量 | 90 | 右对齐 |
| 4 | 线下量 | 90 | 右对齐 |
| 5 | 销售金重 | 90 | 右对齐 |
| 6 | 销售总重 | 90 | 右对齐 |
| 7 | 销售工费 | 90 | 右对齐 |
| 8 | 销售附加费 | 90 | 右对齐 |
| 9 | 销售金额 | 90 | 右对齐 |
| 10 | 应收金额 | 70 | 右对齐 |
| 11 | 预收定金 | 90 | 右对齐 |
| 12 | 优惠金额 | 90 | 右对齐 |
| 13 | 实收金额 | 90 | 右对齐 |
| 14 | 回收金重 | 90 | 右对齐 |
| 15 | 回收金额 | 90 | 右对齐 |
| 16 | 实际金额 | 90 | 右对齐 |
| 17 | 入库总数 | 90 | 右对齐 |
| 18 | 入库金重 | 90 | 右对齐 |
| 19 | 入库总重 | 90 | 右对齐 |
| 20 | 入库成本价 | 90 | 右对齐 |
| 21 | 入库预售价 | 90 | 右对齐 |
| 22 | 调入数量 | 90 | 右对齐 |
| 23 | 调入金重 | 90 | 右对齐 |
| 24 | 调入重量 | 90 | 右对齐 |
| 25 | 调出数量 | 90 | 右对齐 |
| 26 | 调出金重 | 90 | 右对齐 |
| 27 | 调出重量 | 90 | 右对齐 |
| 28 | 退库数量 | 90 | 右对齐 |
| 29 | 退库金重 | 90 | 右对齐 |
| 30 | 退库重量 | 90 | 右对齐 |

### 21.10 报表收银汇总表

**窗口**: 窗口_报表收银汇总表
**功能**: 收银汇总统计

**表头列定义 - 17列**:
| 列号 | 列名 | 宽度 |
|------|------|------|
| 0 | 店铺名称 | 75 |
| 1 | 零售数量 | 100 |
| 2 | 零售金重 | 100 |
| 3 | 零售金额 | 100 |
| 4 | 批发数量 | 100 |
| 5 | 批发金重 | 100 |
| 6 | 批发金额 | 100 |
| 7 | 销售合计数量 | 75 |
| 8 | 销售合计金重 | 75 |
| 9 | 销售合计金额 | 75 |
| 10 | 回收金重 | 75 |
| 11 | 回收金额 | 75 |
| 12 | 实际金额 | 75 |
| 13 | 收款方式金额 | 75 |
| 14 | 优惠金额 | 75 |
| 15 | 预售金额 | 75 |
| 16 | 差异金额 | 75 |

**统计内容**: 零售/批发销售、回收、收款方式、优惠、预售、差异

### 21.11 报表导购回收表

**窗口**: 窗口_报表导购回收表
**功能**: 导购回收统计

**视图模式**: 员工汇总/店铺汇总/明细（3种视图）

**表头列定义**:

**员工汇总/店铺汇总 - 6列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 账户 | 80/0 | 居中 |
| 2 | 导购员 | 80/0 | 居中 |
| 3 | 所属店铺 | 110 | 居中 |
| 4 | 总金重(g) | 100 | 右对齐 |
| 5 | 总金额(元) | 100 | 右对齐 |

> 注：店铺汇总时，账户和导购员列宽为0（隐藏）

**明细视图 - 12列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 账户 | 100 | 居中 |
| 2 | 导购员 | 100 | 居中 |
| 3 | 所属店铺 | 100 | 居中 |
| 4 | 品类 | 100 | 居中 |
| 5 | 名称 | 100 | 居中 |
| 6 | 成色 | 100 | 居中 |
| 7 | 回收金重(g) | 100 | 右对齐 |
| 8 | 回收克价 | 100 | 右对齐 |
| 9 | 其他金额 | 100 | 右对齐 |
| 10 | 回收金额(元) | 100 | 右对齐 |
| 11 | 回收时间 | 140 | 居中 |

### 21.12 报表对照报表

**窗口**: 窗口_报表对照报表
**功能**: 数据对照分析，支持昨日/今日对比

**工具条按钮**（7个）:
- 当日入库重量
- 当日销售重量
- 当日回收重量
- 当日调出重量
- 当日调入重量
- 当日退货重量
- 当前库存重量

**表头列定义 - 7列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 项目 | 85 | 居中 |
| 1 | 昨日数量 | 140 | 右对齐 |
| 2 | 昨日重量(g) | 自适应 | 右对齐 |
| 3 | 当日数量 | 140 | 右对齐 |
| 4 | 当日重量(g) | 自适应 | 右对齐 |
| 5 | 重量变化(g) | 自适应 | 右对齐 |
| 6 | 趋势 | 250 | 右对齐 |

**统计项目**: 入库/销售/回收/调出/调入/退货/库存（7项）

### 21.13 报表员工绩效

**窗口**: 窗口_报表员工绩效
**功能**: 员工绩效统计

**表头列数**: 17列（序号/店铺/岗位/员工/批零/库房id/岗位id/员工账户/销售数量/回收数量/销售重量/回收重量/销售金额/回收金额/销售绩效/回收绩效/绩效合计）

**核心SQL**:
```sql
-- 查询销售绩效数据
SELECT a.kufang, h.post, a.shopping_guide, a.pling,
    SUM(a.quantity) AS xsshu, SUM(a.net_weight) AS xszhong,
    SUM(a.settlement) AS xsjine
FROM xipunum_erp_outbound AS a
INNER JOIN xipunum_erp_user AS h ON h.user = a.shopping_guide
WHERE a.creationtime >= '{开始日期}' AND a.creationtime < '{结束日期}'
    AND a.kufang IN ({库房权限})
GROUP BY a.kufang, h.post, a.shopping_guide, a.pling

-- 查询回收绩效数据
SELECT h.department, h.post, a.shopping_guide,
    SUM(a.quantity) AS hsshu, SUM(a.jin_zhong) AS hszhong,
    SUM(a.retreat_amount) AS hsjine
FROM xipunum_erp_retreat AS a
INNER JOIN xipunum_erp_user AS h ON h.USER = a.shopping_guide
WHERE a.creationtime >= '{开始日期}' AND a.creationtime < '{结束日期}'
    AND h.department IN ({库房权限})
GROUP BY h.department, h.post, a.shopping_guide
```

### 21.14 运营员工业绩

**窗口**: 窗口_运营员工业绩
**功能**: 运营业绩统计，支持导购业绩表和商品明细表两种视图

**视图模式**: 导购业绩表/商品明细表（2种视图）

**表头列定义**:

**导购业绩表 - 重量模式 - 15列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 账户 | 100 | 居中 |
| 2 | 导购员 | 100 | 居中 |
| 3 | 品类 | 自适应 | 居中 |
| 4 | 销售数量 | 自适应 | 右对齐 |
| 5 | 销售金重 | 自适应 | 右对齐 |
| 6 | (分隔) | 5 | - |
| 7 | 退货数量 | 自适应 | 右对齐 |
| 8 | 退货金重 | 自适应 | 右对齐 |
| 9 | (分隔) | 5 | - |
| 10 | 实际数量 | 自适应 | 右对齐 |
| 11 | 实际金重 | 自适应 | 右对齐 |
| 12 | (分隔) | 5 | - |
| 13 | 所在店铺 | 100 | 居中 |
| 14 | 零批 | 65 | 居中 |

**导购业绩表 - 金额模式 - 17列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 账户 | 100 | 居中 |
| 2 | 导购员 | 100 | 居中 |
| 3 | 品类 | 自适应 | 居中 |
| 4 | 销售数量 | 自适应 | 右对齐 |
| 5 | 销售金额 | 自适应 | 右对齐 |
| 6 | 实销金额 | 自适应 | 右对齐 |
| 7 | (分隔) | 5 | - |
| 8 | 退货数量 | 自适应 | 右对齐 |
| 9 | 退货金额 | 自适应 | 右对齐 |
| 10 | 实退金额 | 自适应 | 右对齐 |
| 11 | (分隔) | 5 | - |
| 12 | 实际数量 | 自适应 | 右对齐 |
| 13 | 实际金额 | 自适应 | 右对齐 |
| 14 | (分隔) | 5 | - |
| 15 | 所在店铺 | 100 | 居中 |
| 16 | 零批 | 65 | 居中 |

**商品明细表 - 23列**:
| 列号 | 列名 | 宽度 | 对齐 |
|------|------|------|------|
| 0 | 序号 | 45 | 居中 |
| 1 | 商品编码 | 100 | 居中 |
| 2 | 商品名称 | 140 | 居中 |
| 3 | 款号 | 70 | 居中 |
| 4 | 品类 | 70 | 居中 |
| 5 | 规格 | 70 | 居中 |
| 6 | 材质 | 70 | 居中 |
| 7 | 圈口/长度 | 70 | 居中 |
| 8 | 成色 | 70 | 居中 |
| 9 | 数量 | 70 | 右对齐 |
| 10 | 净重 | 70 | 右对齐 |
| 11 | 成本工费 | 70 | 右对齐 |
| 12 | 参考工费 | 70 | 右对齐 |
| 13 | 成本附加费 | 70 | 右对齐 |
| 14 | 销售单价 | 70 | 右对齐 |
| 15 | 销售金额 | 70 | 右对齐 |
| 16 | 原附加费 | 70 | 右对齐 |
| 17 | 销售克价 | 70 | 右对齐 |
| 18 | 销售工费 | 70 | 右对齐 |
| 19 | 销售附加费 | 70 | 右对齐 |
| 20 | 折扣 | 70 | 右对齐 |
| 21 | 实收金额 | 70 | 右对齐 |
| 22 | 导购员 | 70 | 居中 |

**筛选条件**: 店铺/品类/规格/工厂/批零/导购员

---

## 22. 其他报表（3个）

### 22.1 款式数据汇总

**窗口**: 窗口_款式数据汇总
**功能**: 款号数据汇总分析，支持图片识别搜索

**核心SQL**:
```sql
-- 查询款号汇总
SELECT a.id, a.kuanhao, a.title, a.caizhi, a.yimage,
    b.title AS pinlei, c.title AS guige
FROM xipunum_erp_ksiamges AS a
LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id
LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id
WHERE 1=1 {店铺条件} {品类条件} {规格条件} {材质条件}
ORDER BY a.id DESC
```

### 22.2 款式数据汇总明细

**窗口**: 窗口_款式数据汇总明细
**功能**: 款号数据明细查看，支持入库/销售/库存三种明细视图

### 22.3 款式数据汇总预览

**窗口**: 窗口_款式数据汇总预览
**功能**: 款号图片预览

---

# 第九部分：系统管理模块

---

## 23. 账户管理

### 23.1 个人信息

**窗口**: 窗口_个人信息
**功能**: 个人信息查看与修改

**核心SQL**:
```sql
-- 查询用户信息
SELECT a.*, b.title AS department_name, c.title AS post_name
FROM xipunum_erp_user AS a
LEFT JOIN xipunum_erp_type AS b ON b.id = a.department
LEFT JOIN xipunum_erp_type AS c ON c.id = a.post
WHERE a.user = '{用户名}' LIMIT 1

-- 更新用户信息
UPDATE xipunum_erp_user SET dizhi='{地址}' WHERE user='{用户名}' LIMIT 1
```

### 23.2 账户添加修改

**窗口**: 窗口_账户添加修改
**功能**: 用户新增/编辑

**核心SQL**:
```sql
-- 查询账户详情
SELECT a.*, b.title AS department_name, c.title AS post_name
FROM xipunum_erp_user AS a
LEFT JOIN xipunum_erp_type AS b ON b.id = a.department
LEFT JOIN xipunum_erp_type AS c ON c.id = a.post
WHERE a.user = '{用户名}' LIMIT 1

-- 添加账户
INSERT INTO xipunum_erp_user (user, jianxie, password, name, namejx, state, 
    mailbox, tel, type, google, department, post, dizhi, data, 
    ksdate, jsdate, xsrole, xsdata, cjuser, creationtime)
VALUES ('{用户名}', '{简写}', '{加密密码}', '{姓名}', '{名称简写}', '{状态}',
    '{邮箱}', '{电话}', '{类型}', '{谷歌验证}', '{分组}', '{岗位}', '{地址}', '{权限}',
    '{上班时间}', '{下班时间}', '{销售检索}', '{数据主备}', '{账户}', '{时间}')

-- 修改账户
UPDATE xipunum_erp_user SET name='{姓名}', state='{状态}', department='{分组}',
    post='{岗位}', data='{权限}', updatetime='{时间}'
WHERE user='{用户名}' LIMIT 1
```

**权限字段说明**:
- data字段：全部/店铺/岗位/个人
- xsrole字段：所有/开票
- xsdata字段：主数据/备数据

### 23.3 密码修改

**窗口**: 窗口_密码修改
**功能**: 密码修改

**核心SQL**:
```sql
-- 验证原密码
SELECT * FROM xipunum_erp_user 
WHERE user='{用户名}' AND password='{加密密码}' LIMIT 1

-- 更新密码
UPDATE xipunum_erp_user SET password='{新密码}', updatetime='{时间}' 
WHERE user='{用户名}' LIMIT 1
```

---

## 24. 日志管理

### 24.1 登录日志

**窗口**: 窗口_登录日志
**功能**: 用户登录日志查看

**核心SQL**:
```sql
-- 查询登录日志
SELECT ip, conter, creationtime FROM xipunum_erp_user_log 
WHERE user='{用户名}' ORDER BY id DESC

-- 清空登录日志
DELETE FROM xipunum_erp_user_log WHERE user='{用户名}'
```

### 24.2 入库审核日志

**窗口**: 窗口_入库审核日志
**功能**: 入库订单审核日志查看

---

## 25. 款号管理

### 25.1 款号管理

**窗口**: 窗口_款号管理
**功能**: 款号信息管理，支持图片识别、分页加载

**核心SQL**:
```sql
-- 查询款号列表
SELECT a.*, b.title AS pinlei, c.title AS guige
FROM xipunum_erp_ksiamges AS a
LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id
LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id
WHERE 1=1 {品类条件} {材质条件} {规格条件}
ORDER BY a.id DESC

-- 添加款号
INSERT INTO xipunum_erp_ksiamges (title, kuanhao, yimage, caizhi, chengse,
    category_id, specification_id, lingxiao, cjuser, creationtime)
VALUES ('{名称}', '{款号}', '{图片}', '{材质}', '{成色}',
    '{品类ID}', '{规格ID}', '{零销售}', '{账户}', '{时间}')
```

### 25.2 款号合并

**窗口**: 窗口_款号合并
**功能**: 款号合并操作

### 25.3 款号管理匹配

**窗口**: 窗口_款号管理匹配
**功能**: 款号匹配商品

### 25.4 款号管理添加

**窗口**: 窗口_款号管理添加
**功能**: 款号添加/修改

### 25.5 款号管理图片预览

**窗口**: 窗口_款号管理图片预览
**功能**: 款号图片预览

---

## 26. 辅助功能

### 26.1 批量打印

**窗口**: 窗口_批量打印
**功能**: 批量打印商品标签

### 26.2 批量导入无款

**窗口**: 窗口_批量导入无款
**功能**: 批量导入商品数据（无款号）

**表头列数**: 44列

### 26.3 旧料管理单据

**窗口**: 窗口_旧料管理单据
**功能**: 旧料出入库管理

**核心SQL**:
```sql
-- 查询旧料库存
SELECT a.product_name, SUM(a.jin_zhong) - COALESCE(b.zhongliang, 0) AS jinzhong
FROM xipunum_erp_retreat AS a
LEFT JOIN (
    SELECT product_name, SUM(zhongliang) AS zhongliang 
    FROM xipunum_erp_material WHERE kufang='{库房}' GROUP BY product_name
) AS b ON b.product_name = a.product_name
WHERE a.shopping_guide IN ({权限范围})
GROUP BY a.product_name HAVING jinzhong > 0
```

### 26.4 证书管理

**窗口**: 窗口_证书管理添加修改
**功能**: 商品证书管理

### 26.5 证书机构管理

**窗口**: 窗口_证书机构添加修改
**功能**: 检测机构管理

### 26.6 线上订单驳回

**窗口**: 窗口_线上订单驳回
**功能**: 线上订单驳回处理

---

## 27. 辅助查询窗口（6个）

### 27.1 统一模式说明

6个副编码查询窗口结构完全相同：
- 商品销售副编码查询
- 销售编辑副编码查询
- 商品调拨副编码查询
- 商品退库副编码查询
- 商品退货副编码查询
- 成品修改副编码查询

**统一程序集变量**: 副编码_超级列表框组件句柄

**统一表头**: 8列（复选框/序号/商品编码/副编码/库房/数量/金重/入库时间）

**统一核心SQL**:
```sql
SELECT a.poduct_code, b.fu_code, a.quantity, a.jinzhong,
    CASE WHEN a.kufang='0' THEN '总库' ELSE c.title END AS kufang,
    b.creationtime
FROM xipunum_erp_shop_kucun AS a
INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code
LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang
WHERE a.kufang = '{库房ID}' AND (a.quantity > 0 OR a.jinzhong > 0)
    AND (b.poduct_code = '{编码}' OR b.fu_code = '{编码}')
ORDER BY a.id DESC
```

---

# 第十部分：调试与发布

---

## 28. 调试指南

### 28.1 数据库连接调试

```
1. 确认MySQL服务已启动
2. 测试授权库连接
3. 测试业务库连接
4. 验证读写分离是否正常
```

### 28.2 窗口功能逐模块测试

```
测试顺序：
1. 启动窗口 → 登录验证
2. 主窗口 → 导航加载
3. 基础设置 → 品类/规格/来源
4. 商品管理 → 添加/修改/查询
5. 入库管理 → 创建/详情/修改
6. 销售管理 → 出库/编辑/客退
7. 会员管理 → 添加/合并/查询
8. 库存管理 → 查询/预警/盘点
9. 财务管理 → 收支/结算/绩效
10. 报表模块 → 各类报表
11. 系统管理 → 账户/权限/日志
```

### 28.3 权限验证测试

```
1. 创建不同权限的测试账户
2. 验证数据权限（全部/店铺/岗位/个人）
3. 验证操作权限（添加/修改/删除）
4. 验证可视权限（导航可见性）
```

### 28.4 业务流程完整测试

```
完整业务流程测试：
1. 商品入库 → 创建入库单 → 添加商品 → 保存 → 检查库存
2. 商品销售 → 创建销售单 → 扫码 → 计算金额 → 保存 → 检查库存
3. 客户退货 → 选择销售单 → 退货 → 保存 → 检查库存
4. 商品调拨 → 选择源库房 → 选择目标库房 → 调拨 → 检查库存
5. 会员管理 → 添加会员 → 充值 → 消费 → 查询记录
```

---

## 29. 部署发布

### 29.1 发布配置

```
1. 项目属性 → 发布
2. 选择发布方式：文件系统
3. 目标位置：发布目录
4. 配置：Release
5. 目标框架：net8.0-windows
```

### 29.2 数据库迁移

```
1. 备份现有数据库
2. 在目标机器安装MySQL
3. 恢复数据库
4. 修改App.config中的连接字符串
```

### 29.3 客户端安装包

```
1. 使用Inno Setup或NSIS制作安装包
2. 包含：
   - 可执行文件
   - 配置文件
   - 图片资源
   - .NET 8.0运行时（如未安装）
```

---

# 附录

---

## 附录A：窗口文件与章节对照表（106窗口）

| 序号 | 窗口名称 | 功能模块 | 章节 |
|------|----------|----------|------|
| 1 | _启动窗口 | 系统启动/登录 | 4.1 |
| 2 | 窗口_主窗口 | 主界面 | 5.1 |
| 3 | 窗口_个人信息 | 个人信息 | 23.1 |
| 4 | 窗口_密码修改 | 密码修改 | 23.3 |
| 5 | 窗口_打印机设置 | 打印机配置 | - |
| 6 | 窗口_谷歌验证绑定 | 谷歌验证绑定 | 4.4 |
| 7 | 窗口_谷歌验证码输入 | 谷歌验证码输入 | 4.4 |
| 8 | 窗口_谷歌验证重置 | 谷歌验证重置 | 4.4 |
| 9 | 窗口_授权码信息输入 | 授权码录入 | 4.2 |
| 10 | 窗口_在线更新 | 在线更新 | 4.4 |
| 11 | 窗口_模块更新 | 模块更新 | 4.4 |
| 12 | 窗口_提示消息 | 消息提示 | - |
| 13 | 窗口_日期框 | 日期选择 | - |
| 14 | 窗口_信息商品查询 | 商品查询 | 8.2 |
| 15 | 窗口_商品信息添加 | 商品添加 | 8.1 |
| 16 | 窗口_商品成品数据修改 | 成品修改 | 8.2 |
| 17 | 窗口_商品入库添加 | 入库添加 | 9.1 |
| 18 | 窗口_商品入库详情 | 入库详情 | 9.2 |
| 19 | 窗口_商品入库批量修改 | 入库批量修改 | 9.3 |
| 20 | 窗口_入库审核日志 | 入库审核日志 | 24.2 |
| 21 | 窗口_入库库房修改 | 入库库房修改 | 9.3 |
| 22 | 窗口_入库工厂修改 | 入库工厂修改 | 9.3 |
| 23 | 窗口_商品销售出库 | 销售出库 | 10.1 |
| 24 | 窗口_商品销售编辑 | 销售编辑 | 10.2 |
| 25 | 窗口_商品销售批量修改 | 销售批量修改 | 10.3 |
| 26 | 窗口_商品销售客退 | 销售客退 | 10.4 |
| 27 | 窗口_商品销售外部单据 | 外部单据 | - |
| 28 | 窗口_商品销售订单查询 | 订单查询 | - |
| 29 | 窗口_商品销售副编码查询 | 销售副编码查询 | 27.1 |
| 30 | 窗口_销售编辑副编码查询 | 编辑副编码查询 | 27.1 |
| 31 | 窗口_商品信息调拨 | 商品调拨 | 11.1 |
| 32 | 窗口_商品调拨副编码查询 | 调拨副编码查询 | 27.1 |
| 33 | 窗口_商品信息退库 | 商品退库 | 11.2 |
| 34 | 窗口_商品退库副编码查询 | 退库副编码查询 | 27.1 |
| 35 | 窗口_商品退库备注修改 | 退库备注修改 | 11.2 |
| 36 | 窗口_商品信息回收 | 商品回收 | 12.1 |
| 37 | 窗口_商品信息预售 | 商品预售 | 12.2 |
| 38 | 窗口_商品信息退货 | 商品退货 | 12.3 |
| 39 | 窗口_商品退货副编码查询 | 退货副编码查询 | 27.1 |
| 40 | 窗口_成品修改副编码查询 | 成品副编码查询 | 27.1 |
| 41 | 窗口_成品销售会员绑定 | 会员绑定 | 26.6 |
| 42 | 窗口_会员添加修改 | 会员添加/修改 | 13.1 |
| 43 | 窗口_会员信息合并 | 会员合并 | 13.2 |
| 44 | 窗口_会员列表排序 | 会员排序 | 13.1 |
| 45 | 窗口_会员回访添加 | 会员回访 | 13.3 |
| 46 | 窗口_会员订单消费数据 | 消费记录 | 14.1 |
| 47 | 窗口_会员订单充值数据 | 充值记录 | 14.2 |
| 48 | 窗口_会员订单预购数据 | 预购记录 | 14.3 |
| 49 | 窗口_实时库存查询 | 实时库存 | 15.1 |
| 50 | 窗口_历史库存数据 | 历史库存 | 15.2 |
| 51 | 窗口_历史库存明细 | 库存明细 | 15.3 |
| 52 | 窗口_历史追溯 | 历史追溯 | 15.4 |
| 53 | 窗口_批量打印 | 批量打印 | 26.1 |
| 54 | 窗口_批量打印导入编码 | 打印导入 | 26.1 |
| 55 | 窗口_批量导入无款 | 批量导入 | 26.2 |
| 56 | 窗口_款号管理 | 款号管理 | 25.1 |
| 57 | 窗口_款号合并 | 款号合并 | 25.2 |
| 58 | 窗口_款号管理匹配 | 款号匹配 | 25.3 |
| 59 | 窗口_款号管理图片预览 | 图片预览 | 25.5 |
| 60 | 窗口_品类属性管理 | 品类属性 | 6.1 |
| 61 | 窗口_商品品类管理 | 品类管理 | 6.2 |
| 62 | 窗口_商品规格管理 | 规格管理 | 6.3 |
| 63 | 窗口_商品结算方式 | 结算方式 | 7.1 |
| 64 | 窗口_商品来源管理 | 来源管理 | 7.2 |
| 65 | 窗口_商品查询报表 | 查询报表 | 20.6 |
| 66 | 窗口_商品销售报表 | 销售报表 | 20.1 |
| 67 | 窗口_商品入库报表 | 入库报表 | 20.2 |
| 68 | 窗口_商品退库报表 | 退库报表 | 20.3 |
| 69 | 窗口_商品调拨报表 | 调拨报表 | 20.4 |
| 70 | 窗口_商品回收报表 | 回收报表 | 20.5 |
| 71 | 窗口_销售查询报表 | 销售查询 | 21.1 |
| 72 | 窗口_销售查询简易报表 | 简易查询 | 21.2 |
| 73 | 窗口_销售详情报表 | 销售详情 | 21.3 |
| 74 | 窗口_报表月销售统计 | 月销售统计 | 21.4 |
| 75 | 窗口_报表月汇总销售统计月 | 月汇总统计 | 21.5 |
| 76 | 窗口_报表员工月销售统计 | 员工月统计 | 21.6 |
| 77 | 窗口_报表店铺收支报表 | 店铺收支 | 21.7 |
| 78 | 窗口_报表店铺收支凭证 | 收支凭证 | 21.8 |
| 79 | 窗口_报表商品报表汇总 | 商品汇总 | 21.9 |
| 80 | 窗口_报表收银汇总表 | 收银汇总 | 21.10 |
| 81 | 窗口_报表导购回收表 | 导购回收 | 21.11 |
| 82 | 窗口_报表对照报表 | 对照报表 | 21.12 |
| 83 | 窗口_报表员工绩效 | 员工绩效 | 21.13 |
| 84 | 窗口_运营员工业绩 | 运营业绩 | 21.14 |
| 85 | 窗口_账户添加修改 | 账户管理 | 23.2 |
| 86 | 窗口_绩效信息管理 | 绩效管理 | 19.1 |
| 87 | 窗口_收支管理信息 | 收支管理 | 17.1 |
| 88 | 窗口_收支名称管理 | 收支名称 | 7.3 |
| 89 | 窗口_收支卡号管理 | 卡号管理 | 7.4 |
| 90 | 窗口_店铺数据结算 | 店铺结算 | 18.1 |
| 91 | 窗口_回收名称管理 | 回收名称 | 7.5 |
| 92 | 窗口_旧料管理单据 | 旧料管理 | 26.3 |
| 93 | 窗口_预警管理 | 库存预警 | 16.1 |
| 94 | 窗口_证书管理添加修改 | 证书管理 | 26.4 |
| 95 | 窗口_证书机构添加修改 | 证书机构 | 26.5 |
| 96 | 窗口_登录日志 | 登录日志 | 24.1 |
| 97 | 窗口_结账结料 | 结账结料 | 18.2 |
| 98 | 窗口_款式数据汇总 | 款式汇总 | 22.1 |
| 99 | 窗口_款式数据汇总明细 | 款式明细 | 22.2 |
| 100 | 窗口_款式数据汇总图片预览 | 图片预览 | 22.3 |
| 101 | 窗口_商品信息添加款号 | 款号添加 | 25.4 |
| 102 | 窗口_销售编辑批量修改 | 批量修改 | 10.3 |
| 103 | 窗口_盘点入库导入编码 | 盘点导入 | 16.2 |
| 104 | 窗口_物资盘点添加 | 物资盘点 | 16.2 |
| 105 | 窗口_线上订单驳回 | 订单驳回 | 26.6 |
| 106 | 窗口_进度条 | 进度显示 | - |

---

## 附录B：数据库表完整清单（65+表）

详见《ERP系统开发文档.md》附录B

---

## 附录C：操作权限编号清单

详见《ERP系统开发文档.md》附录C

---

## 附录D：全局变量完整清单

详见《ERP系统开发文档.md》第3章

---

## 附录E：常见问题与解决方案

### E.1 数据库连接失败
- 检查MySQL服务是否启动
- 检查连接字符串是否正确
- 检查防火墙是否允许3306端口

### E.2 中文乱码问题
- 确保数据库使用utf8mb4字符集
- 连接字符串添加`CharSet=utf8mb4`

### E.3 窗口加载缓慢
- 检查SQL查询是否有性能问题
- 添加适当的索引
- 使用分页加载

### E.4 库存数据不一致
- 检查事务是否正确使用
- 检查是否有并发操作
- 使用数据库事务保证原子性

---

> **文档版本**: v1.0
> **更新日期**: 2026-07-13
> **适用框架**: VB.NET 8.0 Windows Forms
> **数据库**: MySQL 8.0 (现有数据库)

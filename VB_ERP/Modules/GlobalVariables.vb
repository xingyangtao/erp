' ============================================================================
' 全局变量模块
' 存储系统运行时的所有全局变量
' ============================================================================

Imports System.Linq

Module GlobalVariables

    ' ========== 系统版本常量 ==========
    Public Const LocalVersion As String = "V3.0.88"   ' 本地版本号

    ' ========== 授权信息变量 ==========
    Public AuthCode As String = ""                    ' 授权码（authorize字段）
    Public AuthDataNum As String = ""                 ' 授权数据码（datanum字段，用于MD5加密盐值）
    Public AuthCompany As String = ""                 ' 授权公司
    Public AuthContact As String = ""                 ' 联系人
    Public AuthPhone As String = ""                   ' 联系电话
    Public AuthStartDate As String = ""               ' 开始日期
    Public AuthEndDate As String = ""                 ' 结束日期
    Public AuthShortCode As String = ""               ' 授权简写
    Public AuthDeviceCount As String = ""             ' 授权设备数量

    ' ========== 用户信息变量 ==========
    Public UserAccount As String = ""                 ' 用户账户
    Public UserName As String = ""                    ' 用户姓名
    Public UserDepartment As String = ""              ' 所属分组ID
    Public UserDepartmentName As String = ""          ' 所属分组名称
    Public UserPost As String = ""                    ' 所属岗位ID
    Public UserPostName As String = ""                ' 所属岗位名称
    Public UserPermission As String = ""              ' 数据权限
    Public UserRole As String = ""                    ' 岗位权限可视
    Public UserOperation As String = ""               ' 岗位权限操作
    Public UserShopPermission As String = ""          ' 店铺权限操作
    Public SalesQueryInvoice As String = ""           ' 销售查询是否按发票筛选（axsrole）
    Public GlobalViewSQL As String = ""               ' 全局查看SQL子查询（shopping_guide IN ...）
    Public UserLoginIP As String = ""                 ' 登录IP
    Public UserLoginTime As String = ""               ' 登录时间
    Public UserAccountID As String = ""               ' 用户ID（数据库自增ID）

    ' ========== 系统配置变量 ==========
    Public ERPCompanyName As String = ""              ' 公司名称
    Public CompanyLogo As String = ""                 ' 公司LOGO
    Public UserAvatar As String = ""                  ' 用户头像
    Public WindowIcon As String = ""                  ' 窗口图标
    Public ReturnWarehouse As String = ""             ' 退库仓ID
    Public ReturnWarehouseName As String = ""         ' 退库仓名称
    Public OnlinePayment As String = ""               ' 线上收款ID
    Public DiscountRate As String = ""                ' 优惠比例
    Public KuanHaoRecognizeURL As String = ""         ' 款号识别地址

    ' ========== 开发商配置变量 ==========
    Public DevCompanyName As String = ""              ' 开发公司名称
    Public DevCompanyDesc As String = ""              ' 开发公司简介
    Public DevContactUs As String = ""                ' 联系我们
    Public DevCustomerService As String = ""          ' 联系客服
    Public DevSoftwareVersion As String = ""          ' 软件版本（服务器）
    Public DevDownloadURL As String = ""              ' 下载地址
    Public DevDownloadDesc As String = ""             ' 下载说明
    Public DevRecognizeURL As String = ""             ' 识别地址

    ' ========== 品类配置变量 ==========
    Public GoldCategoryList As String = ""            ' 金类品类列表（SQL IN格式）
    Public SilverCategoryList As String = ""          ' 银类品类列表（SQL IN格式）
    Public SettlementCategoryList As String = ""      ' 结算品类列表（SQL IN格式）

    ' ========== 打印机配置变量 ==========
    Public LabelPrinterName As String = ""            ' 标签打印机名称
    Public LabelPrinterConnection As String = ""      ' 标签打印机连接

    ' ========== 登录状态变量 ==========
    Public PasswordErrorCount As String = ""          ' 密码错误次数
    Public QueryDataPrimaryBackup As String = "1"     ' 查询数据主备（0=主库读, 1=备库读）
    Public IntroducerRequired As String = ""          ' 介绍人必填
    Public GoogleAuthSecretKey As String = ""         ' 谷歌验证密匙

    ' ========== 数据库连接配置变量 ==========
    Public DBWriteServer As String = ""               ' 写库服务器
    Public DBWritePort As String = ""                 ' 写库端口
    Public DBWriteDatabase As String = ""             ' 写库数据库名
    Public DBWriteUser As String = ""                 ' 写库用户名
    Public DBWritePassword As String = ""             ' 写库密码

    Public DBReadServer As String = ""                ' 读库服务器
    Public DBReadPort As String = ""                  ' 读库端口
    Public DBReadDatabase As String = ""              ' 读库数据库名
    Public DBReadUser As String = ""                  ' 读库用户名
    Public DBReadPassword As String = ""              ' 读库密码

    Public DBReadBackupServer As String = ""          ' 备用读库服务器
    Public DBReadBackupPort As String = ""            ' 备用读库端口
    Public DBReadBackupDatabase As String = ""        ' 备用读库数据库名
    Public DBReadBackupUser As String = ""            ' 备用读库用户名
    Public DBReadBackupPassword As String = ""        ' 备用读库密码

    ' ========== 品类配置变量 ==========
    Public CategoryList As New List(Of CategoryItem)  ' 品类列表
    Public SpecList As New List(Of SpecItem)          ' 规格列表

    ' ========== 品类数据结构 ==========
    Public Class CategoryItem
        Public Property ID As String
        Public Property Title As String
        Public Property JianXie As String
        Public Property CaiZhiID As String
    End Class

    ' ========== 规格数据结构 ==========
    Public Class SpecItem
        Public Property ID As String
        Public Property Title As String
        Public Property JianXie As String
        Public Property CategoryID As String
    End Class

    ' ========== 操作日志变量 ==========
    Public LogOperationDate As String = ""            ' 操作日期
    Public LogOperationAccount As String = ""         ' 操作账户
    Public LogContent As String = ""                  ' 日志内容

    ' ========== 业务传递变量 ==========
    Public InboundOrderNumber As String = ""           ' 全局_商品入库单号（详情窗口用）
    Public HomePageQueryText As String = ""            ' 全局_首页查询栏目文本
    Public InboundProductCodeText As String = ""       ' 全局_入库商品编码文本（提取编码用）
    Public TransferOrderNumber As String = ""          ' 全局_商品调拨单号
    Public ReturnOrderNumber As String = ""             ' 全局_商品退库单号
    Public RefundOrderNumber As String = ""             ' 全局_商品退货单号
    Public RecoveryOrderNumber As String = ""           ' 全局_商品回收单号
    Public PresaleOrderNumber As String = ""            ' 全局_商品预售单号
    Public AccountType As String = ""                  ' 全局_账户类型（后台/店铺）
    Public LogSaveContent As String = ""                ' 全局_日志保存内容
    Public InformationOperationDate As String = ""     ' 全局_信息操作日期
    Public InformationOperationAccount As String = ""   ' 全局_信息操作账户

    ' ========== 获取当前操作日期 ==========
    Public Function GetOperationDate() As String
        Return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
    End Function

    ' ========== 获取当前操作账户 ==========
    Public Function GetOperationAccount() As String
        Return UserAccount
    End Function

    ' ========== 生成订单号 ==========
    Public Function GenerateOrderNumber(prefix As String) As String
        Return prefix & DateTime.Now.ToString("yyyyMMdd") & New Random().Next(1000, 9999).ToString()
    End Function

    ' ========== 生成商品编码 ==========
    Public Function GenerateProductCode() As String
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        Return AuthShortCode & timestamp & UserAccount
    End Function

    ' ========== MD5加密（完全按照易语言：密码+数据码，GBK编码，32位小写） ==========
    Public Function MD5Encrypt(password As String, salt As String) As String
        Dim input As String = password & salt
        Using md5 As Security.Cryptography.MD5 = Security.Cryptography.MD5.Create()
            ' 使用GBK编码（与易语言一致）
            Dim gbk As Text.Encoding = Text.Encoding.GetEncoding("GBK")
            Dim inputBytes As Byte() = gbk.GetBytes(input)
            Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)
            ' 返回32位小写字符串
            Return BitConverter.ToString(hashBytes).Replace("-", "").ToLower()
        End Using
    End Function

    ' ========== 检查操作权限 ==========
    Public Function HasOperationPermission(permissionId As String) As Boolean
        Return UserOperation.Contains("," & permissionId & ",")
    End Function

    ' ========== 检查可视权限 ==========
    Public Function HasVisiblePermission(permissionId As String) As Boolean
        Return UserRole.Contains("," & permissionId & ",")
    End Function

    ' ========== 构建全局查看SQL（对应易语言 全局_全局查看） ==========
    Public Sub BuildGlobalViewSQL()
        Dim shopQuoted = GetQuotedShopPermission()
        Select Case UserPermission
            Case "全部"
                GlobalViewSQL = $" (SELECT z.user FROM xipunum_erp_user AS z WHERE z.department IN ({shopQuoted}))"
            Case "店铺"
                GlobalViewSQL = $" (SELECT z.user FROM xipunum_erp_user AS z WHERE z.department = '{UserDepartment}')"
            Case "岗位"
                GlobalViewSQL = $" (SELECT z.user FROM xipunum_erp_user AS z WHERE z.post = '{UserPost}')"
            Case Else
                GlobalViewSQL = $" (SELECT z.user FROM xipunum_erp_user AS z WHERE z.user = '{UserAccount}')"
        End Select
    End Sub

    Public Function GetQuotedShopPermission() As String
        If String.IsNullOrWhiteSpace(UserShopPermission) Then Return "''"
        Dim raw = UserShopPermission.Trim().TrimEnd(","c)
        If raw.StartsWith("'") Then Return raw
        Dim parts = raw.Split(","c, StringSplitOptions.RemoveEmptyEntries)
        Return String.Join(",", parts.Select(Function(p) $"'{p.Trim()}'"))
    End Function

    Public Function GetSalesInvoiceFilter() As String
        If SalesQueryInvoice = "1" Then Return " AND x.fapiao='1'"
        Return ""
    End Function

    ' ========== 获取店铺权限SQL条件 ==========
    Public Function GetShopPermissionSQL(field As String) As String
        If UserPermission = "全部" Then
            Return "1=1"
        ElseIf UserPermission = "店铺" Then
            Return $"{field} IN ({UserShopPermission})"
        ElseIf UserPermission = "岗位" Then
            Return $"{field} IN ({UserShopPermission})"
        ElseIf UserPermission = "个人" Then
            Return $"{field} = '{UserAccount}'"
        Else
            Return "1=0"
        End If
    End Function

    ' ========== 获取本机IP地址 ==========
    Public Function GetLocalIP() As String
        Try
            Dim hostName As String = System.Net.Dns.GetHostName()
            Dim ipEntry As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(hostName)
            For Each ip As System.Net.IPAddress In ipEntry.AddressList
                If ip.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                    Return ip.ToString()
                End If
            Next
            Return ""
        Catch
            Return ""
        End Try
    End Function

End Module

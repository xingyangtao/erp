' ============================================================================
' 登录窗口
' 完全按照易语言启动窗口实现：窗口程序集_启动窗口.form.e.txt
' 初始化流程：授权码→版本检查→模块更新→插件检查→数据库连接→登录
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.IO
Imports System.Text

Public Class LoginForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（对应易语言：局部_更新模块文本） ==========
    Private UpdateModuleText As String = ""           ' 更新模块文本
    Private AuthInfoID As String = ""                 ' 授权信息ID
    Private _isPasswordPlaceholder As Boolean = True  ' 密码框是否显示占位文本
    Private _isUpdatingPassword As Boolean = False    ' 防止递归调用

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeComponent()
        AddHandler Me.Load, AddressOf LoginForm_Load
        AddHandler btnLogin.Click, AddressOf btnLogin_Click
        AddHandler btnCancel.Click, AddressOf btnCancel_Click
        AddHandler btnTogglePassword.Click, AddressOf btnTogglePassword_Click
        AddHandler txtUsername.KeyDown, AddressOf txtUsername_KeyDown
        AddHandler txtPassword.KeyDown, AddressOf txtPassword_KeyDown
        AddHandler txtUsername.GotFocus, AddressOf txtUsername_GotFocus
        AddHandler txtUsername.LostFocus, AddressOf txtUsername_LostFocus
        AddHandler txtPassword.GotFocus, AddressOf txtPassword_GotFocus
        AddHandler txtPassword.LostFocus, AddressOf txtPassword_LostFocus
    End Sub

    ' ========== 密码显示/隐藏切换 ==========
    Private Sub btnTogglePassword_Click(sender As Object, e As EventArgs)
        If _isPasswordPlaceholder Then Return
        _isUpdatingPassword = True
        txtPassword.UseSystemPasswordChar = Not txtPassword.UseSystemPasswordChar
        If txtPassword.UseSystemPasswordChar Then
            btnTogglePassword.Text = "👁"
        Else
            btnTogglePassword.Text = "🙈"
        End If
        _isUpdatingPassword = False
    End Sub

    ' ========== 账号编辑框按键事件（对应易语言：_帐号_编辑框_按下某键） ==========
    Private Sub txtUsername_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtUsername.Text) Then
                ShowWarning("用户账户不能为空！")
                txtUsername.Focus()
                Return
            End If
            btnLogin_Click(sender, e)
        End If
    End Sub

    ' ========== 账号编辑框获得焦点（对应易语言：_帐号_编辑框_获得焦点） ==========
    Private Sub txtUsername_GotFocus(sender As Object, e As EventArgs)
        If txtUsername.Text = "请输入账号" Then
            txtUsername.Text = ""
        End If
    End Sub

    ' ========== 账号编辑框失去焦点（对应易语言：_帐号_编辑框_失去焦点） ==========
    Private Sub txtUsername_LostFocus(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(txtUsername.Text) Then
            txtUsername.Text = "请输入账号"
        End If
    End Sub

    ' ========== 密码编辑框按键事件（对应易语言：_密码_编辑框_按下某键） ==========
    Private Sub txtPassword_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtPassword.Text) OrElse txtPassword.Text = "请输入密码" Then
                ShowWarning("登陆密码不能为空！")
                txtPassword.Focus()
                Return
            End If
            btnLogin_Click(sender, e)
        End If
    End Sub

    ' ========== 密码编辑框获得焦点（对应易语言：_密码_编辑框_获得焦点） ==========
    Private Sub txtPassword_GotFocus(sender As Object, e As EventArgs)
        If _isUpdatingPassword Then Return
        If _isPasswordPlaceholder Then
            _isUpdatingPassword = True
            _isPasswordPlaceholder = False
            txtPassword.Text = ""
            txtPassword.UseSystemPasswordChar = True
            txtPassword.Font = New Font(txtPassword.Font.FontFamily, txtPassword.Font.Size)
            _isUpdatingPassword = False
        End If
    End Sub

    ' ========== 密码编辑框失去焦点（对应易语言：_密码_编辑框_失去焦点） ==========
    Private Sub txtPassword_LostFocus(sender As Object, e As EventArgs)
        If _isUpdatingPassword Then Return
        If String.IsNullOrEmpty(txtPassword.Text) Then
            _isUpdatingPassword = True
            _isPasswordPlaceholder = True
            txtPassword.UseSystemPasswordChar = False
            txtPassword.Text = "请输入密码"
            _isUpdatingPassword = False
        End If
    End Sub

    ' ========== 窗口加载（完全按照易语言__启动窗口_创建完毕流程） ==========
    Private Sub LoginForm_Load(sender As Object, e As EventArgs)
        Try
            ' 初始化界面
            lblStatus.Text = "正在初始化..."
            lblStatus.Refresh()

            ' ========== 第1步：获取外网IP（对应易语言：网页_取外网IP） ==========
            UserLoginIP = GetLocalIP()

            ' ========== 第2步：连接授权库（对应易语言：连接MySql） ==========
            lblStatus.Text = "正在连接授权数据库..."
            lblStatus.Refresh()
            If Not ConnectAuthDB() Then
                ShowError("数据库账户密码错误！")
                Application.Exit()
                Return
            End If

            ' ========== 第3步：读取config.ini获取授权码 ==========
            lblStatus.Text = "正在读取配置文件..."
            lblStatus.Refresh()
            Dim configPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "config.ini")
            Dim infoDataCode As String = ""

            If File.Exists(configPath) Then
                infoDataCode = ReadConfigValue("config", "authorize", "")
            End If

            If String.IsNullOrEmpty(infoDataCode) Then
                ' config.ini不存在或授权码为空，打开授权码输入窗口
                Dim authForm As New AuthorizationForm()
                Me.Hide()
                authForm.ShowDialog()
                Application.Exit()
                Return
            End If

            ' ========== 第4步：查询erp_authorize获取授权信息 ==========
            lblStatus.Text = "正在验证授权信息..."
            lblStatus.Refresh()
            If Not LoadAuthorizationInfo(infoDataCode) Then
                Return
            End If

            ' ========== 第5步：查询erp_config获取系统参数 ==========
            lblStatus.Text = "正在加载系统参数..."
            lblStatus.Refresh()
            LoadSystemParameters()

            ' ========== 第6步：检查版本更新 ==========
            lblStatus.Text = "正在检查版本更新..."
            lblStatus.Refresh()
            If CheckVersionUpdate() Then
                Return
            End If

            ' ========== 第7步：检查模块更新（对应易语言：读取数据库） ==========
            lblStatus.Text = "正在检查模块更新..."
            lblStatus.Refresh()
            CheckModuleUpdate()
            If Not String.IsNullOrEmpty(UpdateModuleText) Then
                ShowWarning("当前有模块需要更新！")
                Dim updateForm As New ModuleUpdateForm()
                updateForm.ShowDialog()
                Application.Exit()
                Return
            End If

            ' ========== 第8步：检查标签打印软件 ==========
            lblStatus.Text = "正在检查插件..."
            lblStatus.Refresh()
            If CheckLabelPrintSoftware() Then
                Return
            End If

            ' ========== 第9步：检查打印插件 ==========
            If CheckReportPlugin() Then
                Return
            End If

            ' ========== 第10步：读取数据库连接配置 ==========
            lblStatus.Text = "正在读取数据库配置..."
            lblStatus.Refresh()
            If Not LoadDatabaseConfig() Then
                Return
            End If

            ' ========== 第11步：连接业务数据库 ==========
            lblStatus.Text = "正在连接业务数据库..."
            lblStatus.Refresh()
            If Not ConnectToBusinessDB() Then
                Return
            End If

            ' ========== 第12步：查询xipunum_erp_config获取客户信息 ==========
            lblStatus.Text = "正在加载客户信息..."
            lblStatus.Refresh()
            LoadCustomerInfo()

            ' ========== 第13步：查询退库仓名称 ==========
            If Not String.IsNullOrEmpty(ReturnWarehouse) Then
                LoadReturnWarehouseName()
            End If

            ' ========== 第14步：查询xipunum_erp_category_stat_config获取品类配置 ==========
            lblStatus.Text = "正在加载品类配置..."
            lblStatus.Refresh()
            LoadCategoryStatConfig()

            ' ========== 第15步：读取打印机配置 ==========
            LoadPrinterConfig()

            ' ========== 第16步：下载并缓存图片 ==========
            lblStatus.Text = "正在下载图片资源..."
            lblStatus.Refresh()
            DownloadAndCacheImages()

            ' ========== 第17步：加载保存的用户名密码 ==========
            lblStatus.Text = "正在加载用户信息..."
            lblStatus.Refresh()
            LoadSavedCredentials()

            ' 设置版本信息（对应易语言：版本标签.标题）
            lblVersion.Text = $" 2025-{DateTime.Now:yyyy} {DevCompanyName} 版权所有 {LocalVersion}   当前IP:{UserLoginIP}"
            lblWelcome.Text = $"欢迎来到{ERPCompanyName}！"
            lblStatus.Text = "初始化完成，请登录"
            lblStatus.Refresh()

        Catch ex As Exception
            ShowError("初始化失败：" & ex.Message)
            Application.Exit()
        End Try
    End Sub

    ' ========== 加载授权信息（对应易语言：SELECT * FROM erp_authorize） ==========
    Private Function LoadAuthorizationInfo(infoDataCode As String) As Boolean
        Try
            ' 调试日志
            System.IO.File.AppendAllText(
                IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log"),
                $"[{DateTime.Now}] LoadAuthorizationInfo called with infoDataCode={infoDataCode}" & vbCrLf
            )

            Dim sql As String = $"SELECT * FROM erp_authorize WHERE authorize='{SafeSQL(infoDataCode)}' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Auth)

            If dt.Rows.Count = 0 Then
                ShowError("授权信息不存在！")
                MySQL_Auth.Close()
                Application.Exit()
                Return False
            End If

            Dim row As DataRow = dt.Rows(0)
            AuthInfoID = SafeString(row("id"))
            Dim authDataNum As String = SafeString(row("datanum"))

            ' 调试日志
            System.IO.File.AppendAllText(
                IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log"),
                $"[{DateTime.Now}] LoadAuth: authDataNum={authDataNum}, AuthInfoID={AuthInfoID}" & vbCrLf
            )
            Dim authCompany As String = SafeString(row("gongsi"))
            Dim authNum As String = SafeString(row("num"))
            Dim authName As String = SafeString(row("name"))
            Dim authTel As String = SafeString(row("tel"))
            Dim authStartTime As String = SafeString(row("kstime"))
            Dim authEndTime As String = SafeString(row("jstime"))
            Dim authJianXie As String = SafeString(row("jianxie"))
            DevSoftwareVersion = SafeString(row("banben"))
            DevDownloadURL = SafeString(row("xiazai"))
            DevDownloadDesc = SafeString(row("conter"))

            If String.IsNullOrEmpty(authDataNum) OrElse String.IsNullOrEmpty(AuthInfoID) Then
                ShowError("授权信息不存在！")
                MySQL_Auth.Close()
                Application.Exit()
                Return False
            End If

            ' 保存授权信息到全局变量（对应易语言：编码转换 UTF8->GBK）
            AuthCode = infoDataCode
            AuthDataNum = authDataNum

            ' 调试日志 - 确认赋值成功
            System.IO.File.AppendAllText(
                IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log"),
                $"[{DateTime.Now}] After assignment: AuthDataNum={AuthDataNum}" & vbCrLf
            )

            AuthCompany = GBKToUTF8(authCompany)
            AuthDeviceCount = GBKToUTF8(authNum)
            AuthContact = GBKToUTF8(authName)
            AuthPhone = GBKToUTF8(authTel)
            AuthStartDate = GBKToUTF8(authStartTime)
            AuthEndDate = GBKToUTF8(authEndTime)
            AuthShortCode = GBKToUTF8(authJianXie)
            DevDownloadURL = GBKToUTF8(DevDownloadURL)
            DevDownloadDesc = GBKToUTF8(DevDownloadDesc)

            Return True
        Catch ex As Exception
            ShowError("加载授权信息失败：" & ex.Message)
            Application.Exit()
            Return False
        End Try
    End Function

    ' ========== 加载系统参数（对应易语言：SELECT FROM erp_config） ==========
    Private Sub LoadSystemParameters()
        Try
            Dim sql As String = "SELECT (SELECT conter FROM erp_config WHERE title='公司名称' LIMIT 1) AS mingcheng," &
                                "(SELECT conter FROM erp_config WHERE title='公司简介' LIMIT 1) AS jianjie," &
                                "(SELECT conter FROM erp_config WHERE title='联系我们' LIMIT 1) AS lianxi," &
                                "(SELECT conter FROM erp_config WHERE title='联系客服' LIMIT 1) AS kefu," &
                                "(SELECT conter FROM erp_config WHERE title='软件版本' LIMIT 1) AS banben," &
                                "(SELECT conter FROM erp_config WHERE title='下载地址' LIMIT 1) AS xiazai," &
                                "(SELECT conter FROM erp_config WHERE title='识别地址' LIMIT 1) AS shibie"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Auth)

            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                DevCompanyName = GBKToUTF8(SafeString(row("mingcheng")))
                DevCompanyDesc = GBKToUTF8(SafeString(row("jianjie")))
                DevContactUs = GBKToUTF8(SafeString(row("lianxi")))
                DevCustomerService = GBKToUTF8(SafeString(row("kefu")))
                DevRecognizeURL = GBKToUTF8(SafeString(row("shibie")))
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 检查版本更新 ==========
    Private Function CheckVersionUpdate() As Boolean
        Try
            If Not String.IsNullOrEmpty(DevSoftwareVersion) AndAlso DevSoftwareVersion <> LocalVersion Then
                If ConfirmAction("您当前的版本较低，请前往更新最新版本") Then
                    Dim updateForm As New OnlineUpdateForm()
                    updateForm.Show()
                    Me.Close()
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    ' ========== 检查模块更新（对应易语言：读取数据库子程序） ==========
    Private Sub CheckModuleUpdate()
        Try
            Dim sql As String = "SELECT * FROM erp_voucher WHERE state='0' ORDER BY id ASC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Auth)

            If dt.Rows.Count = 0 Then
                UpdateModuleText = ""
                Return
            End If

            Dim needUpdateIds As String = ""
            For Each row As DataRow In dt.Rows
                Dim moduleId As String = SafeString(row("id"))
                Dim serverVersion As String = GBKToUTF8(SafeString(row("banben1")))

                ' 从config.ini读取本地版本
                Dim localModuleVersion As String = ReadConfigValue("mokuai", "mokuai" & moduleId, "")

                ' 如果本地版本为空或与服务器版本不同，则需要更新
                If String.IsNullOrEmpty(localModuleVersion) OrElse localModuleVersion <> serverVersion Then
                    needUpdateIds &= "'" & moduleId & "',"
                End If
            Next

            UpdateModuleText = needUpdateIds
        Catch ex As Exception
            UpdateModuleText = ""
        End Try
    End Sub

    ' ========== 检查标签打印软件（对应易语言：LABEL MAATRIX检查） ==========
    Private Function CheckLabelPrintSoftware() As Boolean
        Try
            Dim lmwPath1 As String = "C:\Program Files (x86)\LABEL MAATRIX\LMWPRINT.EXE"
            Dim lmwPath2 As String = "C:\Program Files (x86)\LABEL MAATRIX\lmw.exe"
            If Not File.Exists(lmwPath1) AndAlso Not File.Exists(lmwPath2) Then
                If ConfirmAction("标签打印软件未安装，是否安装打印软件") Then
                    Dim matrixPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sofe", "Matrix8.exe")
                    If File.Exists(matrixPath) Then
                        Process.Start(matrixPath)
                    End If
                    Application.Exit()
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    ' ========== 检查打印插件（对应易语言：Grid++Report检查） ==========
    Private Function CheckReportPlugin() As Boolean
        Try
            Dim gridReportPath As String = "C:\Grid++Report 6\grdesigner6.exe"
            If Not File.Exists(gridReportPath) Then
                If ConfirmAction("打印插件未安装，是否安装打印插件") Then
                    Dim gridReportSetup As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sofe", "gridreport6.8.exe")
                    If File.Exists(gridReportSetup) Then
                        Process.Start(gridReportSetup)
                    End If
                    Application.Exit()
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    ' ========== 加载数据库连接配置 ==========
    Private Function LoadDatabaseConfig() As Boolean
        Try
            ' 查询写库配置（state=1）
            Dim sql As String = $"SELECT * FROM erp_mysql WHERE authorizeid='{AuthInfoID}' AND state='1' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Auth)

            If dt.Rows.Count = 0 Then
                ShowError("业务写库连接配置不存在！")
                MySQL_Auth.Close()
                Application.Exit()
                Return False
            End If

            Dim row As DataRow = dt.Rows(0)
            DBWriteServer = SafeString(row("server"))
            DBWritePort = SafeString(row("port"))
            DBWriteDatabase = SafeString(row("database"))
            DBWritePassword = SafeString(row("password"))
            DBWriteUser = SafeString(row("dbuser"))

            ' 查询读库配置（state=0）
            sql = $"SELECT * FROM erp_mysql WHERE authorizeid='{AuthInfoID}' AND state='0' LIMIT 1"
            dt = DatabaseModule.ExecuteQuery(sql, MySQL_Auth)

            If dt.Rows.Count = 0 Then
                ShowError("业务读库连接配置不存在！")
                MySQL_Auth.Close()
                Application.Exit()
                Return False
            End If

            row = dt.Rows(0)
            DBReadServer = SafeString(row("server"))
            DBReadPort = SafeString(row("port"))
            DBReadDatabase = SafeString(row("database"))
            DBReadPassword = SafeString(row("password"))
            DBReadUser = SafeString(row("dbuser"))

            ' 查询备用读库配置（state=2）
            sql = $"SELECT * FROM erp_mysql WHERE authorizeid='{AuthInfoID}' AND state='2' LIMIT 1"
            dt = DatabaseModule.ExecuteQuery(sql, MySQL_Auth)

            If dt.Rows.Count > 0 Then
                row = dt.Rows(0)
                DBReadBackupServer = SafeString(row("server"))
                DBReadBackupPort = SafeString(row("port"))
                DBReadBackupDatabase = SafeString(row("database"))
                DBReadBackupPassword = SafeString(row("password"))
                DBReadBackupUser = SafeString(row("dbuser"))
            End If

            Return True
        Catch ex As Exception
            ShowError("加载数据库配置失败：" & ex.Message)
            Application.Exit()
            Return False
        End Try
    End Function

    ' ========== 连接业务数据库 ==========
    Private Function ConnectToBusinessDB() As Boolean
        ' 连接写库
        If Not ConnectBusinessDB(DBWriteServer, DBWritePort, DBWriteDatabase, DBWriteUser, DBWritePassword) Then
            ShowError("业务写库连接失败，请检查 erp_mysql 配置！")
            MySQL_Auth.Close()
            Application.Exit()
            Return False
        End If

        ' 连接读库（打印专用、订单查询专用）
        MySQL_ReadPrint = CreateConnection(DBReadServer, DBReadPort, DBReadDatabase, DBReadUser, DBReadPassword)
        MySQL_ReadOrder = CreateConnection(DBReadServer, DBReadPort, DBReadDatabase, DBReadUser, DBReadPassword)

        Return True
    End Function

    ' ========== 创建数据库连接 ==========
    Private Function CreateConnection(server As String, port As String, database As String, username As String, password As String) As MySqlConnection
        Try
            Dim connStr As String = $"Server={server};Port={port};Database={database};Uid={username};Pwd={password};SslMode=None;CharSet=utf8mb4;"
            Dim conn As New MySqlConnection(connStr)
            conn.Open()
            Return conn
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ' ========== 加载客户信息（对应易语言：SELECT FROM xipunum_erp_config） ==========
    Private Sub LoadCustomerInfo()
        Try
            Dim sql As String = "SELECT (SELECT conter FROM xipunum_erp_config WHERE title='公司名称' LIMIT 1) AS mingcheng," &
                                "(SELECT conter FROM xipunum_erp_config WHERE title='软件logo' LIMIT 1) AS logo," &
                                "(SELECT conter FROM xipunum_erp_config WHERE title='款号识别地址' LIMIT 1) AS shibie," &
                                "(SELECT conter FROM xipunum_erp_config WHERE title='用户头像' LIMIT 1) AS ico," &
                                "(SELECT conter FROM xipunum_erp_config WHERE title='优惠比例' LIMIT 1) AS bilinum," &
                                "(SELECT conter FROM xipunum_erp_config WHERE title='退库仓' LIMIT 1) AS tuikuid," &
                                "(SELECT conter FROM xipunum_erp_config WHERE title='线上收款' LIMIT 1) AS shoukuan," &
                                "(SELECT conter FROM xipunum_erp_config WHERE title='窗口图标' LIMIT 1) AS chuangkouimg," &
                                "(SELECT conter FROM xipunum_erp_config WHERE title='API网址' LIMIT 1) AS apiwangzhi"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)

            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                ERPCompanyName = GBKToUTF8(SafeString(row("mingcheng")))
                CompanyLogo = GBKToUTF8(SafeString(row("logo")))
                KuanHaoRecognizeURL = GBKToUTF8(SafeString(row("shibie")))
                UserAvatar = GBKToUTF8(SafeString(row("ico")))
                DiscountRate = GBKToUTF8(SafeString(row("bilinum")))
                ReturnWarehouse = GBKToUTF8(SafeString(row("tuikuid")))
                OnlinePayment = GBKToUTF8(SafeString(row("shoukuan")))
                WindowIcon = GBKToUTF8(SafeString(row("chuangkouimg")))
                DevRecognizeURL = GBKToUTF8(SafeString(row("apiwangzhi")))
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 加载退库仓名称（对应易语言：SELECT title FROM xipunum_erp_type） ==========
    Private Sub LoadReturnWarehouseName()
        Try
            Dim sql As String = $"SELECT title FROM xipunum_erp_type WHERE id='{SafeSQL(ReturnWarehouse)}' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)

            If dt.Rows.Count > 0 Then
                ReturnWarehouseName = GBKToUTF8(SafeString(dt.Rows(0)("title")))
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 加载品类统计配置（对应易语言：SELECT FROM xipunum_erp_category_stat_config） ==========
    Private Sub LoadCategoryStatConfig()
        Try
            Dim sql As String = "SELECT (SELECT category_list FROM xipunum_erp_category_stat_config WHERE title='金类' LIMIT 1) AS jinlei," &
                                "(SELECT category_list FROM xipunum_erp_category_stat_config WHERE title='银类' LIMIT 1) AS yinlei," &
                                "(SELECT category_list FROM xipunum_erp_category_stat_config WHERE title='黄金' LIMIT 1) AS jiesuan"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)

            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                Dim jinLei As String = GBKToUTF8(SafeString(row("jinlei")))
                Dim yinLei As String = GBKToUTF8(SafeString(row("yinlei")))
                Dim jieSuan As String = GBKToUTF8(SafeString(row("jiesuan")))

                ' 处理金类品类（对应易语言：替换[为'，]为'，,为','）
                If String.IsNullOrEmpty(jinLei) Then
                    GoldCategoryList = "''"
                Else
                    GoldCategoryList = "'" & jinLei.Replace("[", "'").Replace("]", "'").Replace(",", "','") & "'"
                End If

                ' 处理银类品类
                If String.IsNullOrEmpty(yinLei) Then
                    SilverCategoryList = "''"
                Else
                    SilverCategoryList = "'" & yinLei.Replace("[", "'").Replace("]", "'").Replace(",", "','") & "'"
                End If

                ' 处理结算品类
                If String.IsNullOrEmpty(jieSuan) Then
                    SettlementCategoryList = "''"
                Else
                    SettlementCategoryList = "'" & jieSuan.Replace("[", "'").Replace("]", "'").Replace(",", "','") & "'"
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 加载打印机配置（对应易语言：读取data\print.ini） ==========
    Private Sub LoadPrinterConfig()
        Try
            Dim printIniPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "print.ini")
            If File.Exists(printIniPath) Then
                LabelPrinterName = ReadConfigValue("print", "name", "")
                LabelPrinterConnection = ReadConfigValue("print", "lianjie", "")
            Else
                WriteConfigValue("print", "name", "")
                WriteConfigValue("print", "lianjie", "")
                LabelPrinterName = ""
                LabelPrinterConnection = ""
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 下载并缓存图片（对应易语言：HTTP读文件） ==========
    Private Sub DownloadAndCacheImages()
        Try
            ' 下载LOGO
            If Not String.IsNullOrEmpty(CompanyLogo) Then
                Dim logoPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", "logo.jpg")
                If ReadConfigValue("images", "logo", "") = "" Then
                    Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img"))
                    If File.Exists(logoPath) Then File.Delete(logoPath)
                    DownloadImageSync(CompanyLogo, logoPath)
                    WriteConfigValue("images", "logo", logoPath)
                End If
            End If

            ' 下载头像
            If Not String.IsNullOrEmpty(UserAvatar) Then
                Dim avatarPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", "touxiang.jpg")
                If ReadConfigValue("images", "ico", "") = "" Then
                    Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img"))
                    If File.Exists(avatarPath) Then File.Delete(avatarPath)
                    DownloadImageSync(UserAvatar, avatarPath)
                    WriteConfigValue("images", "ico", avatarPath)
                End If
            End If

            ' 下载窗口图标
            If Not String.IsNullOrEmpty(WindowIcon) Then
                Dim iconPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", "tubiao.jpg")
                If ReadConfigValue("images", "tubiao", "") = "" Then
                    Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img"))
                    If File.Exists(iconPath) Then File.Delete(iconPath)
                    DownloadImageSync(WindowIcon, iconPath)
                    WriteConfigValue("images", "tubiao", iconPath)
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 同步下载图片 ==========
    Private Sub DownloadImageSync(url As String, savePath As String)
        Try
            Using client As New System.Net.WebClient()
                client.DownloadFile(url, savePath)
            End Using
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 加载保存的凭证（对应易语言：读取data\user.ini） ==========
    Private Sub LoadSavedCredentials()
        Try
            Dim userIniPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "user.ini")
            If File.Exists(userIniPath) Then
                Dim savedUser As String = ReadIniFile(userIniPath, "user", "user", "")
                Dim savedPass As String = ReadIniFile(userIniPath, "user", "pass", "")

                txtUsername.Text = savedUser

                If String.IsNullOrEmpty(savedUser) Then
                    txtUsername.Text = "请输入账号"
                End If

                If String.IsNullOrEmpty(savedPass) Then
                    _isPasswordPlaceholder = True
                    txtPassword.UseSystemPasswordChar = False
                    txtPassword.Text = "请输入密码"
                Else
                    _isPasswordPlaceholder = False
                    txtPassword.Text = savedPass
                    txtPassword.UseSystemPasswordChar = True
                End If

                If Not String.IsNullOrEmpty(savedUser) AndAlso savedUser <> "请输入账号" Then
                    txtUsername.Focus()
                    chkRemember.Checked = True
                Else
                    chkRemember.Checked = False
                End If
            Else
                txtUsername.Text = "请输入账号"
                _isPasswordPlaceholder = True
                txtPassword.UseSystemPasswordChar = False
                txtPassword.Text = "请输入密码"
                chkRemember.Checked = False
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 登录按钮点击（完全按照易语言：_登录按钮_被单击） ==========
    Private Sub btnLogin_Click(sender As Object, e As EventArgs)
        ' 记录登录加载时间
        Dim loginStartTime As Integer = Environment.TickCount

        ' 获取当前日期
        Dim currentDate As String = DateTime.Now.ToString("yyyy-MM-dd")

        ' 检查授权开始日期
        If Not String.IsNullOrEmpty(AuthStartDate) AndAlso DateTime.Parse(currentDate) < DateTime.Parse(AuthStartDate) Then
            ShowWarning("授权尚未生效，请联系客服！")
            Return
        End If

        ' 检查授权结束日期
        If Not String.IsNullOrEmpty(AuthEndDate) AndAlso DateTime.Parse(currentDate) > DateTime.Parse(AuthEndDate) Then
            ShowWarning("授权已到期，请联系客服续期！")
            Return
        End If

        ' 验证用户名
        Dim username As String = txtUsername.Text.Trim()
        If String.IsNullOrEmpty(username) OrElse username = "请输入账号" Then
            ShowWarning("用户账户不能为空！")
            txtUsername.Focus()
            Return
        End If

        ' 验证密码
        Dim password As String = txtPassword.Text.Trim()
        If String.IsNullOrEmpty(password) OrElse password = "请输入密码" Then
            ShowWarning("登陆密码不能为空！")
            txtPassword.Focus()
            Return
        End If

        ' 禁用登录按钮
        btnLogin.Enabled = False

        ' 获取当前时间
        Dim currentTime As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        ' 编码转换（对应易语言：GBK转UTF8）
        Dim queryUsername As String = GBKToUTF8(username)

        ' 确保AuthDataNum已加载（对应易语言：直接使用全_授权数据码）
        If String.IsNullOrEmpty(AuthDataNum) Then
            Dim authSql As String = $"SELECT datanum FROM erp_authorize WHERE authorize='{SafeSQL(AuthCode)}' LIMIT 1"
            Dim authDt As DataTable = DatabaseModule.ExecuteQuery(authSql, MySQL_Auth)
            If authDt.Rows.Count > 0 Then
                AuthDataNum = SafeString(authDt.Rows(0)("datanum"))
            End If
        End If

        Dim queryPassword As String = MD5Encrypt(password, AuthDataNum)

        ' 调试日志
        System.IO.File.AppendAllText(
            IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log"),
            $"[{DateTime.Now}] Login: username={username}, queryUsername={queryUsername}, password={password}, AuthDataNum={AuthDataNum}, queryPassword={queryPassword}" & vbCrLf
        )

        ' 保存用户账户
        UserAccount = username

        ' 查询用户是否存在（对应易语言：SELECT id,login FROM xipunum_erp_user）
        Dim sql As String = $"SELECT id,login FROM xipunum_erp_user WHERE user='{SafeSQL(queryUsername)}' LIMIT 1"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)

        If dt.Rows.Count = 0 Then
            ShowError("此登录账户不存在！")
            txtUsername.Focus()
            btnLogin.Enabled = True
            Return
        End If

        Dim row As DataRow = dt.Rows(0)
        Dim userId As String = SafeString(row("id"))
        Dim loginCount As String = SafeString(row("login"))

        ' 检查错误次数（对应易语言：admin账户不检查）
        If username <> "admin" Then
            If Not String.IsNullOrEmpty(loginCount) AndAlso Integer.Parse(loginCount) >= 3 Then
                ShowError("账户登录错误次数超出3次，账户已被禁用，请联系管理员开启账户！")
                txtUsername.Focus()
                btnLogin.Enabled = True
                Return
            End If
        End If

        ' 保存登录账户ID和密码错误次数
        UserAccountID = GBKToUTF8(userId)
        PasswordErrorCount = GBKToUTF8(loginCount)

        ' 验证密码（对应易语言：SELECT * FROM xipunum_erp_user where user and password）
        Dim checkCount As Integer = 0
        sql = $"SELECT * FROM xipunum_erp_user WHERE user='{SafeSQL(queryUsername)}' AND password='{SafeSQL(queryPassword)}' LIMIT 1"
        dt = DatabaseModule.ExecuteQuery(sql)
        checkCount = dt.Rows.Count

        If checkCount = 0 Then
            ' 密码错误
            ShowError("登录密码输入错误，请输入正确的登录密码！")

            ' 增加错误次数
            Dim newErrorCount As Integer = If(String.IsNullOrEmpty(PasswordErrorCount), 0, Integer.Parse(PasswordErrorCount)) + 1
            Dim stateUpdate As String = ""
            If newErrorCount >= 3 Then
                stateUpdate = ",state='1'"
            End If

            ' 更新错误次数（对应易语言：UPDATE xipunum_erp_user SET login）
            If username <> "admin" Then
                DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_user SET login='{newErrorCount}'{stateUpdate} WHERE id='{SafeSQL(userId)}' LIMIT 1")

                ' 记录登录失败日志（对应易语言：增加记录 xipunum_erp_user_log）
                Dim failLogContent As String = $"账户:{UserAccount}在{currentTime}登录失败"
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_user_log (ip, user, conter, creationtime) VALUES ('{SafeSQL(UserLoginIP)}', '{SafeSQL(queryUsername)}', '{SafeSQL(failLogContent)}', '{currentDate}')")
            End If

            txtUsername.Focus()
            btnLogin.Enabled = True
            Return
        End If

        ' 保存凭证（对应易语言：写配置项 data\user.ini）
        Dim userIniPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "user.ini")
        If chkRemember.Checked Then
            WriteIniFile(userIniPath, "user", "user", txtUsername.Text)
            WriteIniFile(userIniPath, "user", "pass", txtPassword.Text)
        Else
            WriteIniFile(userIniPath, "user", "user", "")
            WriteIniFile(userIniPath, "user", "pass", "")
        End If

        ' 查询用户完整信息（对应易语言：SELECT a.*, b.title, c.title, d.*）
        sql = $"SELECT b.data8 as bitian,a.id AS aid,a.state AS astate,a.NAME AS aname,a.type AS atype," &
              $"a.department AS adepartment,a.post AS apost,a.ip AS aip,a.mailbox AS mailbox," &
              $"a.google AS agoogle,a.google_secret AS google_secret,a.DATA AS adata," &
              $"a.ksdate AS aksdate,a.jsdate AS ajsdate,b.title AS fenzu,c.title AS gangwei," &
              $"d.keshi AS keshi,d.caozuo AS caozuo,d.shopid AS shopid," &
              $"a.xsrole as axsrole,a.xsdata as xsdata " &
              $"FROM xipunum_erp_user AS a " &
              $"INNER JOIN xipunum_erp_type AS b ON b.id = a.department " &
              $"INNER JOIN xipunum_erp_type AS c ON c.id = a.post " &
              $"INNER JOIN xipunum_erp_role AS d ON d.typeid = a.post " &
              $"WHERE a.user='{SafeSQL(queryUsername)}' AND a.password='{SafeSQL(queryPassword)}' LIMIT 1"
        dt = DatabaseModule.ExecuteQuery(sql)

        If dt.Rows.Count = 0 Then
            btnLogin.Enabled = True
            Return
        End If

        row = dt.Rows(0)

        ' 读取用户信息
        Dim loginAccountId As String = SafeString(row("aid"))
        Dim loginAccountState As String = SafeString(row("astate"))
        Dim loginAccountName As String = SafeString(row("aname"))
        Dim accountType As String = SafeString(row("atype"))
        Dim accountDepartmentId As String = SafeString(row("adepartment"))
        Dim accountPostId As String = SafeString(row("apost"))
        Dim loginAccountIP As String = SafeString(row("aip"))
        Dim googleAuthEmail As String = SafeString(row("mailbox"))
        Dim googleAuthEnabled As String = SafeString(row("agoogle"))
        Dim googleAuthSecret As String = SafeString(row("google_secret"))
        Dim accountViewPermission As String = SafeString(row("adata"))
        Dim loginStartTime2 As String = SafeString(row("aksdate"))
        Dim loginEndTime As String = SafeString(row("ajsdate"))
        Dim accountDepartmentName As String = SafeString(row("fenzu"))
        Dim accountPostName As String = SafeString(row("gangwei"))
        Dim postPermissionView As String = SafeString(row("keshi"))
        Dim postPermissionOperation As String = SafeString(row("caozuo"))
        Dim shopPermissionOperation As String = SafeString(row("shopid"))
        Dim introducerRequired As String = SafeString(row("bitian"))
        Dim salesQueryInvoice As String = SafeString(row("axsrole"))
        Dim queryDataPrimaryBackup As String = SafeString(row("xsdata"))

        ' 检查账户状态（对应易语言：astate != "0"）
        If loginAccountState <> "0" Then
            ShowError("当前登陆账户已被停止使用！")
            txtUsername.Focus()
            btnLogin.Enabled = True
            Return
        End If

        ' 保存用户信息到全局变量（对应易语言：编码转换 UTF8->GBK）
        UserAccountID = GBKToUTF8(loginAccountId)
        UserName = GBKToUTF8(loginAccountName)
        UserPermission = GBKToUTF8(accountViewPermission)
        UserDepartment = GBKToUTF8(accountDepartmentId)
        UserPost = GBKToUTF8(accountPostId)
        UserDepartmentName = GBKToUTF8(accountDepartmentName)
        UserPostName = GBKToUTF8(accountPostName)
        IntroducerRequired = GBKToUTF8(introducerRequired)
        GoogleAuthSecretKey = GBKToUTF8(googleAuthSecret)
        SalesQueryInvoice = GBKToUTF8(salesQueryInvoice)
        UserRole = GBKToUTF8(postPermissionView)
        UserOperation = GBKToUTF8(postPermissionOperation)
        UserShopPermission = GBKToUTF8(shopPermissionOperation)
        QueryDataPrimaryBackup = GBKToUTF8(queryDataPrimaryBackup)

        ' 连接报表/任务读库（对应易语言：根据查询数据主备连接）
        If QueryDataPrimaryBackup = "0" Then
            MySQL_ReadReport = CreateConnection(DBReadServer, DBReadPort, DBReadDatabase, DBReadUser, DBReadPassword)
            MySQL_ReadTask = CreateConnection(DBReadServer, DBReadPort, DBReadDatabase, DBReadUser, DBReadPassword)
        Else
            MySQL_ReadReport = CreateConnection(DBReadBackupServer, DBReadBackupPort, DBReadBackupDatabase, DBReadBackupUser, DBReadBackupPassword)
            MySQL_ReadTask = CreateConnection(DBReadBackupServer, DBReadBackupPort, DBReadBackupDatabase, DBReadBackupUser, DBReadBackupPassword)
        End If

        ' 检查报表/任务读库连接
        If MySQL_ReadReport Is Nothing OrElse MySQL_ReadTask Is Nothing Then
            ShowError("报表/任务读库连接失败！")
            btnLogin.Enabled = True
            Return
        End If

        ' 超级管理员加载所有权限（对应易语言：全局_账户岗位名称 == "超级管理员"）
        If UserPostName = "超级管理员" Then
            ' 加载所有可视权限
            UserRole = ""
            sql = "SELECT role FROM erp_navigation WHERE state='0' ORDER BY id ASC"
            dt = DatabaseModule.ExecuteQuery(sql, MySQL_Auth)
            For Each roleRow As DataRow In dt.Rows
                UserRole &= GBKToUTF8(SafeString(roleRow("role"))) & ","
            Next

            ' 加载所有操作权限
            UserOperation = ""
            sql = "SELECT chazhao FROM erp_navigation_type WHERE 1=1 ORDER BY id ASC"
            dt = DatabaseModule.ExecuteQuery(sql, MySQL_Auth)
            For Each opRow As DataRow In dt.Rows
                UserOperation &= GBKToUTF8(SafeString(opRow("chazhao"))) & ","
            Next

            ' 加载所有店铺权限
            UserShopPermission = ""
            sql = "SELECT id AS akufang FROM xipunum_erp_type WHERE type='商铺' AND superior='0' UNION ALL SELECT '0' AS akufang FROM dual ORDER BY akufang='0' DESC, akufang"
            dt = DatabaseModule.ExecuteQuery(sql)
            For Each shopRow As DataRow In dt.Rows
                UserShopPermission &= GBKToUTF8(SafeString(shopRow("akufang"))) & ","
            Next

            ' 更新岗位权限（对应易语言：UPDATE xipunum_erp_role SET keshi, caozuo, shopid）
            DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_role SET keshi='{SafeSQL(UserRole)}',caozuo='{SafeSQL(UserOperation)}',shopid='{SafeSQL(UserShopPermission)}' WHERE typeid='{SafeSQL(accountPostId)}' LIMIT 1")
        End If

        ' 权限处理（去除末尾逗号和多余引号）
        UserRole = UserRole.Trim().TrimEnd(","c).Trim("'"c)
        UserOperation = UserOperation.Trim().TrimEnd(","c)
        UserShopPermission = UserShopPermission.Trim().TrimEnd(","c)

        ' 格式化可视权限（对应易语言：替换,为','）
        If String.IsNullOrEmpty(UserRole) Then
            UserRole = "''"
        Else
            UserRole = "'" & UserRole & "'"
            UserRole = UserRole.Replace(",", "','")
        End If

        ' 操作权限处理
        UserOperation = "," & UserOperation & ","

        ' 店铺权限处理（去除末尾逗号和多余引号）
        UserShopPermission = UserShopPermission.Trim().TrimEnd(","c).Trim("'"c)

        ' 格式化店铺权限
        If String.IsNullOrEmpty(UserShopPermission) Then
            UserShopPermission = "''"
        Else
            UserShopPermission = "'" & UserShopPermission & "'"
            UserShopPermission = UserShopPermission.Replace(",", "','")
        End If

        ' 构建全局查看SQL（对应易语言：根据账户查看权限设置全局_全局查看）
        Select Case GBKToUTF8(accountViewPermission)
            Case "全部"
                GlobalViewSQL = $" (SELECT z.user FROM xipunum_erp_user AS z WHERE z.department IN ({UserShopPermission}))"
            Case "店铺"
                GlobalViewSQL = $" (SELECT z.user FROM xipunum_erp_user AS z WHERE z.department='{accountDepartmentId}')"
            Case "岗位"
                GlobalViewSQL = $" (SELECT z.user FROM xipunum_erp_user AS z WHERE z.post='{accountPostId}')"
            Case "个人"
                GlobalViewSQL = $" (SELECT z.user FROM xipunum_erp_user AS z WHERE z.user='{queryUsername}')"
        End Select

        ' 检查登录时间（对应易语言：判断当前时间是否在上班期间）
        Dim currentJudgeTime As String = DateTime.Now.ToString("HH:mm:ss")
        If Not String.IsNullOrEmpty(loginStartTime2) AndAlso Not String.IsNullOrEmpty(loginEndTime) Then
            If currentJudgeTime < loginStartTime2 OrElse currentJudgeTime > loginEndTime Then
                ShowWarning("当前账户登录时间不在上班期间！")
                txtUsername.Focus()
                btnLogin.Enabled = True
                Return
            End If
        End If

        ' 谷歌验证（对应易语言：谷歌验证开启 == "0"）
        If googleAuthEnabled = "0" Then
            If String.IsNullOrEmpty(googleAuthSecret) Then
                ' 密匙为空，前往绑定
                ShowWarning("您的谷歌验证密匙为空,前往绑定谷歌验证密匙！")
                Me.Close()
                Dim bindForm As New GoogleAuthBindForm()
                bindForm.Show()
                Return
            End If

            ' 检查是否需要验证（IP为空或IP变化）
            If String.IsNullOrEmpty(loginAccountIP) OrElse loginAccountIP <> UserLoginIP Then
                Me.Close()
                Dim verifyForm As New GoogleAuthVerifyForm()
                verifyForm.Show()
                Return
            End If
        End If

        ' 登录成功，更新用户表（对应易语言：UPDATE xipunum_erp_user SET ip, logintime, login）
        DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_user SET ip='{SafeSQL(UserLoginIP)}',logintime='{SafeSQL(currentTime)}',login='0' WHERE id='{SafeSQL(loginAccountId)}' LIMIT 1")

        ' 记录登录成功日志（对应易语言：增加记录 xipunum_erp_user_log）
        Dim logContent As String = $"账户:{UserAccount}在{currentTime}登录成功"
        DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_user_log (ip, user, conter, creationtime) VALUES ('{SafeSQL(UserLoginIP)}', '{SafeSQL(queryUsername)}', '{SafeSQL(logContent)}', '{currentDate}')")

        ' 计算登录加载时间
        Dim loginLoadTime As Integer = Environment.TickCount - loginStartTime

        ' 显示登录成功提示
        ShowInfo("登陆成功！")

        ' 打开主窗口（对应易语言：载入(窗口_主窗口, , 假)）
        Dim mainForm As New MainForm()
        mainForm.WindowState = FormWindowState.Maximized
        mainForm.Show()
        mainForm.Activate()
        mainForm.BringToFront()

        ' 启用登录按钮并隐藏登录窗口
        btnLogin.Enabled = True
        Me.Hide()
    End Sub

    ' ========== 取消按钮点击 ==========
    Private Sub btnCancel_Click(sender As Object, e As EventArgs)
        CloseAllConnections()
        Application.Exit()
    End Sub

    ' ========== 窗口关闭 ==========
    Private Sub LoginForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        CloseAllConnections()
    End Sub

End Class

' ============================================================================
' 模块更新窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_模块更新.form.e.txt
' 功能：查询需要更新的模块，下载ZIP并解压到对应目录（支持中文文件名）
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.IO
Imports System.IO.Compression
Imports System.Net.Http
Imports System.Text

Public Class ModuleUpdateForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private downloadRowIndex As Integer = 0           ' 局部_下载行数

    ' ========== 控件声明 ==========
    Private dgvModules As New DataGridView()           ' 模块列表
    Private WithEvents btnUpdate As New Button()       ' 更新按钮
    Private lblStatus As New Label()                   ' 状态标签
    Private progressBar As New ProgressBar()           ' 总进度条

    ' ========== 构造函数 ==========
    Public Sub New()
        ' 注册编码提供程序以支持GBK
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        InitializeUI()
        AddHandler Me.Load, AddressOf ModuleUpdateForm_Load
        AddHandler btnUpdate.Click, AddressOf btnUpdate_Click
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "模块更新"
        Me.Size = New Drawing.Size(900, 550)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 标题
        Dim lblTitle As New Label()
        lblTitle.Text = "以下模块需要更新："
        lblTitle.Font = New Drawing.Font("微软雅黑", 12, Drawing.FontStyle.Bold)
        lblTitle.Location = New Drawing.Point(20, 15)
        lblTitle.AutoSize = True
        Me.Controls.Add(lblTitle)

        ' 模块列表
        dgvModules.Location = New Drawing.Point(20, 50)
        dgvModules.Size = New Drawing.Size(845, 350)
        dgvModules.ReadOnly = True
        dgvModules.AllowUserToAddRows = False
        dgvModules.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvModules)

        ' 初始化表格
        InitGrid()

        ' 状态标签
        lblStatus.Text = "就绪"
        lblStatus.Location = New Drawing.Point(20, 415)
        lblStatus.AutoSize = True
        Me.Controls.Add(lblStatus)

        ' 进度条
        progressBar.Location = New Drawing.Point(20, 440)
        progressBar.Size = New Drawing.Size(845, 20)
        progressBar.Minimum = 0
        progressBar.Maximum = 100
        Me.Controls.Add(progressBar)

        ' 更新按钮
        btnUpdate.Text = "开始更新"
        btnUpdate.Location = New Drawing.Point(380, 475)
        btnUpdate.Size = New Drawing.Size(120, 35)
        Me.Controls.Add(btnUpdate)
    End Sub

    ' ========== 初始化表格 ==========
    Private Sub InitGrid()
        dgvModules.Columns.Clear()
        dgvModules.Columns.Add("xuhao", "序号")
        dgvModules.Columns.Add("type", "类型")
        dgvModules.Columns.Add("role", "类别")
        dgvModules.Columns.Add("name", "名称")
        dgvModules.Columns.Add("old_version", "原版本")
        dgvModules.Columns.Add("new_version", "新版本")
        dgvModules.Columns.Add("size", "文件大小")
        dgvModules.Columns.Add("status", "状态")
        ' 隐藏列：下载地址、文件位置、模块ID
        dgvModules.Columns.Add("download_url", "下载地址")
        dgvModules.Columns.Add("save_path", "文件位置")
        dgvModules.Columns.Add("module_id", "模块ID")

        dgvModules.Columns("xuhao").Width = 50
        dgvModules.Columns("type").Width = 80
        dgvModules.Columns("role").Width = 80
        dgvModules.Columns("name").Width = 200
        dgvModules.Columns("old_version").Width = 80
        dgvModules.Columns("new_version").Width = 80
        dgvModules.Columns("size").Width = 100
        dgvModules.Columns("status").Width = 100
        dgvModules.Columns("download_url").Width = 0
        dgvModules.Columns("download_url").Visible = False
        dgvModules.Columns("save_path").Width = 0
        dgvModules.Columns("save_path").Visible = False
        dgvModules.Columns("module_id").Width = 0
        dgvModules.Columns("module_id").Visible = False
    End Sub

    ' ========== 窗口加载 ==========
    Private Sub ModuleUpdateForm_Load(sender As Object, e As EventArgs)
        LoadModuleList()
    End Sub

    ' ========== 加载模块列表 ==========
    Private Sub LoadModuleList()
        Try
            ' 查询需要更新的模块
            Dim sql As String = "SELECT * FROM erp_voucher WHERE state='0' ORDER BY id ASC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Auth)

            dgvModules.Rows.Clear()
            Dim index As Integer = 0
            For Each row As DataRow In dt.Rows
                Dim id As String = SafeString(row("id"))
                Dim serverVersion As String = GBKToUTF8(SafeString(row("banben1")))

                ' 从config.ini读取本地版本
                Dim localVersion As String = ReadConfigValue("mokuai", "mokuai" & id, "")

                ' 只显示需要更新的模块（版本不同）
                If Not String.IsNullOrEmpty(localVersion) AndAlso localVersion = serverVersion Then
                    Continue For
                End If

                index += 1
                Dim type As String = GBKToUTF8(SafeString(row("type")))
                Dim role As String = GBKToUTF8(SafeString(row("role")))
                Dim name As String = GBKToUTF8(SafeString(row("name")))
                Dim oldVersion As String = If(String.IsNullOrEmpty(localVersion), "无", localVersion)
                Dim size As Decimal = SafeDecimal(row("daxiao"))
                Dim unit As String = SafeString(row("danwei"))
                Dim downloadUrl As String = GBKToUTF8(SafeString(row("xzurl")))
                Dim savePath As String = GBKToUTF8(SafeString(row("weizhi")))

                ' 转换MB为KB
                Dim sizeStr As String
                If unit = "MB" Then
                    sizeStr = (size / 1024).ToString("F2") & "KB"
                Else
                    sizeStr = size.ToString() & unit
                End If

                dgvModules.Rows.Add(
                    index.ToString("000"),
                    type,
                    role,
                    name,
                    oldVersion,
                    serverVersion,
                    sizeStr,
                    "待更新",
                    downloadUrl,
                    savePath,
                    id
                )
            Next

            If dgvModules.Rows.Count = 0 Then
                lblStatus.Text = "所有模块已是最新版本"
                btnUpdate.Enabled = False
            Else
                lblStatus.Text = $"共 {dgvModules.Rows.Count} 个模块需要更新"
            End If
        Catch ex As Exception
            ShowError("加载模块列表失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 更新按钮点击 ==========
    Private Async Sub btnUpdate_Click(sender As Object, e As EventArgs)
        If dgvModules.Rows.Count = 0 Then
            ShowWarning("没有需要更新的模块！")
            Return
        End If

        btnUpdate.Enabled = False
        progressBar.Value = 0

        ' 创建目录
        Dim sofePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sofe")
        Dim voucherPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "voucher")
        Directory.CreateDirectory(sofePath)
        Directory.CreateDirectory(Path.Combine(voucherPath, "danju"))
        Directory.CreateDirectory(Path.Combine(voucherPath, "biaoqian"))

        ' 逐个下载并解压模块
        For i As Integer = 0 To dgvModules.Rows.Count - 1
            downloadRowIndex = i
            dgvModules.Rows(i).Cells("status").Value = "下载中..."
            lblStatus.Text = $"正在下载 {dgvModules.Rows(i).Cells("name").Value}..."
            lblStatus.Refresh()

            Try
                Dim downloadUrl As String = SafeString(dgvModules.Rows(i).Cells("download_url").Value)
                Dim savePath As String = SafeString(dgvModules.Rows(i).Cells("save_path").Value)
                Dim moduleId As String = SafeString(dgvModules.Rows(i).Cells("module_id").Value)
                Dim newVersion As String = SafeString(dgvModules.Rows(i).Cells("new_version").Value)

                If Not String.IsNullOrEmpty(downloadUrl) Then
                    ' 创建目标目录
                    Dim fullDirPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, savePath)
                    Directory.CreateDirectory(fullDirPath)

                    ' 下载ZIP文件
                    Dim zipPath As String = Path.Combine(fullDirPath, "data.zip")

                    Using client As New HttpClient()
                        Dim response = Await client.GetAsync(downloadUrl)
                        response.EnsureSuccessStatusCode()
                        Dim bytes = Await response.Content.ReadAsByteArrayAsync()
                        File.WriteAllBytes(zipPath, bytes)
                    End Using

                    ' 解压ZIP文件（使用GBK编码支持中文文件名）
                    dgvModules.Rows(i).Cells("status").Value = "解压中..."
                    dgvModules.Rows(i).Cells("status").Style.ForeColor = Drawing.Color.Blue
                    lblStatus.Refresh()

                    Dim gbk As Encoding = Encoding.GetEncoding("GBK")
                    ZipFile.ExtractToDirectory(zipPath, fullDirPath, gbk, True)

                    ' 删除ZIP文件
                    File.Delete(zipPath)

                    ' 更新进度
                    progressBar.Value = CInt((i + 1) * 100 / dgvModules.Rows.Count)
                    dgvModules.Rows(i).Cells("status").Value = "已完成"
                    dgvModules.Rows(i).Cells("status").Style.ForeColor = Drawing.Color.Green

                    ' 更新配置文件中的模块版本
                    WriteConfigValue("mokuai", "mokuai" & moduleId, newVersion)
                End If
            Catch ex As Exception
                dgvModules.Rows(i).Cells("status").Value = "失败: " & ex.Message
                dgvModules.Rows(i).Cells("status").Style.ForeColor = Drawing.Color.Red
            End Try
        Next

        lblStatus.Text = "更新完成！"
        ShowSuccess("模块更新完成！请重新启动程序。")

        ' 打开登录窗口
        Dim loginForm As New LoginForm()
        loginForm.Show()
        Me.Close()
    End Sub

End Class

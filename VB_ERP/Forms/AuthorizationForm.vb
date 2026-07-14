' ============================================================================
' 授权码信息输入窗口
' 完全按照易语言启动窗口实现：窗口程序集_窗口_授权码信息输入.form.e.txt
' 功能：创建data文件夹和config.ini文件，填写授权码
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.IO

Public Class AuthorizationForm
    Inherits System.Windows.Forms.Form

    ' ========== 控件声明 ==========
    Private WithEvents btnConfirm As New Button()
    Private WithEvents btnCancel As New Button()
    Private txtAuthCode As New TextBox()
    Private lblTitle As New Label()
    Private lblAuthCode As New Label()
    Private lblStatus As New Label()

    ' ========== 构造函数 ==========
    Public Sub New()
        ' 初始化界面
        InitializeUI()
        AddHandler Me.Load, AddressOf AuthorizationForm_Load
        AddHandler btnConfirm.Click, AddressOf btnConfirm_Click
        AddHandler btnCancel.Click, AddressOf btnCancel_Click
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "授权码信息输入"
        Me.Size = New Drawing.Size(450, 250)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False

        ' 标题
        lblTitle.Text = "请输入授权码"
        lblTitle.Font = New Drawing.Font("微软雅黑", 14, Drawing.FontStyle.Bold)
        lblTitle.Location = New Drawing.Point(150, 20)
        lblTitle.AutoSize = True
        Me.Controls.Add(lblTitle)

        ' 授权码标签
        lblAuthCode.Text = "授权码："
        lblAuthCode.Location = New Drawing.Point(30, 70)
        lblAuthCode.AutoSize = True
        Me.Controls.Add(lblAuthCode)

        ' 授权码输入框
        txtAuthCode.Location = New Drawing.Point(100, 67)
        txtAuthCode.Size = New Drawing.Size(300, 25)
        Me.Controls.Add(txtAuthCode)

        ' 状态标签
        lblStatus.Text = ""
        lblStatus.Location = New Drawing.Point(100, 100)
        lblStatus.AutoSize = True
        lblStatus.ForeColor = Drawing.Color.Blue
        Me.Controls.Add(lblStatus)

        ' 确认按钮
        btnConfirm.Text = "确认"
        btnConfirm.Location = New Drawing.Point(100, 140)
        btnConfirm.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnConfirm)

        ' 取消按钮
        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(220, 140)
        btnCancel.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnCancel)
    End Sub

    ' ========== 窗口加载 ==========
    Private Sub AuthorizationForm_Load(sender As Object, e As EventArgs)
        ' 创建data文件夹
        Dim dataPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data")
        If Not Directory.Exists(dataPath) Then
            Directory.CreateDirectory(dataPath)
        End If

        ' 检查config.ini是否存在
        Dim configPath As String = Path.Combine(dataPath, "config.ini")
        If File.Exists(configPath) Then
            ' 读取已有的授权码
            Dim existingCode As String = ReadConfigValue("config", "authorize", "")
            If Not String.IsNullOrEmpty(existingCode) Then
                txtAuthCode.Text = existingCode
            End If
        End If
    End Sub

    ' ========== 确认按钮点击 ==========
    Private Sub btnConfirm_Click(sender As Object, e As EventArgs)
        Dim authCode As String = txtAuthCode.Text.Trim()

        If String.IsNullOrEmpty(authCode) Then
            ShowWarning("请输入授权码！")
            txtAuthCode.Focus()
            Return
        End If

        ' 显示状态
        lblStatus.Text = "正在验证授权码..."
        lblStatus.Refresh()

        ' 验证授权码
        If Not ValidateAuthCode(authCode) Then
            lblStatus.Text = ""
            Return
        End If

        ' 保存授权码到config.ini
        lblStatus.Text = "正在保存配置..."
        lblStatus.Refresh()

        Dim configPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "config.ini")
        Dim dirPath As String = Path.GetDirectoryName(configPath)
        If Not Directory.Exists(dirPath) Then
            Directory.CreateDirectory(dirPath)
        End If

        ' 写入config.ini文件
        WriteConfigValue("config", "authorize", authCode)

        lblStatus.Text = "授权码保存成功！"
        lblStatus.Refresh()

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    ' ========== 验证授权码 ==========
    Private Function ValidateAuthCode(authCode As String) As Boolean
        Try
            ' 连接授权库
            If Not ConnectAuthDB() Then
                ShowError("无法连接授权数据库，请检查MySQL服务是否启动。")
                Return False
            End If

            ' 查询授权码
            Dim sql As String = $"SELECT * FROM erp_authorize WHERE authorize='{SafeSQL(authCode)}' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Auth)

            If dt.Rows.Count = 0 Then
                ShowError("授权码无效，请检查输入是否正确！")
                MySQL_Auth.Close()
                Return False
            End If

            Dim row As DataRow = dt.Rows(0)
            Dim authId As String = SafeString(row("id"))
            Dim authDataNum As String = SafeString(row("datanum"))

            If String.IsNullOrEmpty(authDataNum) OrElse String.IsNullOrEmpty(authId) Then
                ShowError("授权信息不完整，请联系客服！")
                MySQL_Auth.Close()
                Return False
            End If

            MySQL_Auth.Close()
            Return True
        Catch ex As Exception
            ShowError("验证授权码失败：" & ex.Message)
            Return False
        End Try
    End Function

    ' ========== 取消按钮点击 ==========
    Private Sub btnCancel_Click(sender As Object, e As EventArgs)
        CloseAllConnections()
        Application.Exit()
    End Sub

End Class

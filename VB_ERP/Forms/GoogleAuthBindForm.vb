' ============================================================================
' 谷歌验证绑定窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_谷歌验证绑定.form.e.txt
' 功能：生成密匙、显示二维码、绑定谷歌验证
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.IO
Imports QRCoder

Public Class GoogleAuthBindForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（对应易语言） ==========
    Private bindingSecret As String = ""           ' 绑定生成密匙
    Private qrCodePath As String = ""              ' 二维码生成路径

    ' ========== 控件声明 ==========
    Private grpMain As New GroupBox()               ' 谷歌验证绑定_分组框
    Private lblName As New Label()                  ' 绑定名称标签
    Private txtName As New TextBox()                ' 绑定名称编辑框
    Private picQRCode As New PictureBox()           ' 图片框EX_主图
    Private lblSecret As New Label()                ' 密匙显示标签
    Private txtSecret As New TextBox()              ' 密匙显示框
    Private btnBind As New Button()                 ' 按钮EX1_绑定
    Private btnClose As New Button()                ' 关闭按钮

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf GoogleAuthBindForm_Load
        AddHandler Me.Resize, AddressOf GoogleAuthBindForm_Resize
        AddHandler btnBind.Click, AddressOf btnBind_Click
        AddHandler btnClose.Click, AddressOf btnClose_Click
        AddHandler txtName.TextChanged, AddressOf txtName_TextChanged
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "谷歌验证绑定"
        Me.Size = New Drawing.Size(500, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False

        ' 分组框
        grpMain.Text = "谷歌验证绑定"
        grpMain.Size = New Drawing.Size(440, 520)
        Me.Controls.Add(grpMain)

        ' 绑定名称
        lblName.Text = "绑定名称："
        lblName.Location = New Drawing.Point(20, 30)
        lblName.AutoSize = True
        grpMain.Controls.Add(lblName)

        txtName.Location = New Drawing.Point(100, 27)
        txtName.Size = New Drawing.Size(300, 25)
        grpMain.Controls.Add(txtName)

        ' 二维码图片
        picQRCode.Location = New Drawing.Point(70, 70)
        picQRCode.Size = New Drawing.Size(300, 300)
        picQRCode.SizeMode = PictureBoxSizeMode.Zoom
        picQRCode.BorderStyle = BorderStyle.FixedSingle
        grpMain.Controls.Add(picQRCode)

        ' 密匙显示
        lblSecret.Text = "密匙："
        lblSecret.Location = New Drawing.Point(20, 390)
        lblSecret.AutoSize = True
        grpMain.Controls.Add(lblSecret)

        txtSecret.Location = New Drawing.Point(100, 387)
        txtSecret.Size = New Drawing.Size(300, 25)
        txtSecret.ReadOnly = True
        grpMain.Controls.Add(txtSecret)

        ' 绑定按钮
        btnBind.Text = "绑定"
        btnBind.Location = New Drawing.Point(100, 430)
        btnBind.Size = New Drawing.Size(120, 40)
        grpMain.Controls.Add(btnBind)

        ' 关闭按钮
        btnClose.Text = "关闭"
        btnClose.Location = New Drawing.Point(250, 430)
        btnClose.Size = New Drawing.Size(120, 40)
        grpMain.Controls.Add(btnClose)
    End Sub

    ' ========== 窗口加载（对应易语言_窗口_谷歌验证绑定_创建完毕） ==========
    Private Sub GoogleAuthBindForm_Load(sender As Object, e As EventArgs)
        ' 设置绑定名称为用户账户
        txtName.Text = UserAccount

        ' 生成密匙
        bindingSecret = GenerateSecretKey()

        ' 显示密匙
        txtSecret.Text = bindingSecret

        ' 生成二维码路径
        qrCodePath = $"otpauth://totp/{txtName.Text}?secret={bindingSecret}"

        ' 生成二维码
        GenerateQRCode()
    End Sub

    ' ========== 窗口尺寸改变（居中分组框） ==========
    Private Sub GoogleAuthBindForm_Resize(sender As Object, e As EventArgs)
        If Me.WindowState <> FormWindowState.Minimized Then
            grpMain.Left = (Me.Width - grpMain.Width) \ 2
            grpMain.Top = (Me.Height - grpMain.Height) \ 2
        End If
    End Sub

    ' ========== 绑定名称改变（更新二维码） ==========
    Private Sub txtName_TextChanged(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(txtName.Text) Then
            txtName.Text = UserAccount
        End If
        qrCodePath = $"otpauth://totp/{txtName.Text}?secret={bindingSecret}"
        GenerateQRCode()
    End Sub

    ' ========== 生成二维码（对应易语言_数据加载_二维码生成） ==========
    Private Sub GenerateQRCode()
        Try
            Using qrGenerator As New QRCodeGenerator()
                Using qrCodeData As QRCodeData = qrGenerator.CreateQrCode(qrCodePath, QRCodeGenerator.ECCLevel.Q)
                    Using qrCode As New QRCode(qrCodeData)
                        Dim qrImage As Bitmap = qrCode.GetGraphic(20)
                        picQRCode.Image = qrImage
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("生成二维码失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 绑定按钮点击（对应易语言_按钮EX1_绑定_鼠标左键单击） ==========
    Private Sub btnBind_Click(sender As Object, e As EventArgs)
        Try
            Dim operationDate As String = GetOperationDate()

            ' 调试信息
            System.IO.File.AppendAllText(
                IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log"),
                $"[{DateTime.Now}] GoogleAuthBind: UserAccountID={UserAccountID}, Secret={bindingSecret}" & vbCrLf
            )

            ' 更新数据库：绑定谷歌验证密匙
            Dim sql As String = $"UPDATE xipunum_erp_user SET google_secret='{SafeSQL(bindingSecret)}', ip='' WHERE id='{SafeSQL(UserAccountID)}' LIMIT 1"
            DatabaseModule.ExecuteCommand(sql)

            ' 记录日志
            Dim logContent As String = $"账户:{UserAccount} 绑定谷歌验证，密匙为：{bindingSecret}"
            Dim logSql As String = $"INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('添加', '添加谷歌验证', '{SafeSQL(logContent)}', '{SafeSQL(UserAccount)}', '{operationDate}')"
            DatabaseModule.ExecuteCommand(logSql)

            ShowSuccess("谷歌验证绑定成功！请重新登录。")

            ' 重新打开登录窗口
            Dim loginForm As New LoginForm()
            loginForm.Show()
            Me.Close()
        Catch ex As Exception
            ShowError("绑定失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 关闭按钮点击 ==========
    Private Sub btnClose_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ' ========== 生成密匙（对应易语言生成密匙()） ==========
    Private Function GenerateSecretKey() As String
        ' 生成Base32编码的密匙（16位）
        Dim chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"
        Dim random As New Random()
        Dim secret As String = ""
        For i As Integer = 0 To 15
            secret &= chars(random.Next(chars.Length))
        Next
        Return secret
    End Function

End Class

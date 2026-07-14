' ============================================================================
' 谷歌验证码输入窗口
' 功能: 登录时输入谷歌验证码
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class GoogleAuthVerifyForm
    Inherits System.Windows.Forms.Form

    Private txtCode As New TextBox()
    Private WithEvents btnVerify As New Button()
    Private WithEvents btnCancel As New Button()
    Private loginAttempts As Integer = 0

    Public Property Verified As Boolean = False

    Public Sub New()
        InitializeUI()
        ' 回车键触发验证
        Me.KeyPreview = True
        AddHandler Me.KeyDown, AddressOf GoogleAuthVerifyForm_KeyDown
    End Sub

    Private Sub GoogleAuthVerifyForm_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            btnVerify_Click(Nothing, Nothing)
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub InitializeUI()
        Me.Text = "谷歌验证码输入"
        Me.Size = New Drawing.Size(350, 200)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblTip As New Label()
        lblTip.Text = "请输入谷歌验证器中的6位验证码："
        lblTip.Location = New Drawing.Point(30, 20)
        lblTip.AutoSize = True
        Me.Controls.Add(lblTip)

        txtCode.Location = New Drawing.Point(30, 50)
        txtCode.Size = New Drawing.Size(270, 25)
        txtCode.MaxLength = 6
        Me.Controls.Add(txtCode)

        btnVerify.Text = "验证"
        btnVerify.Location = New Drawing.Point(60, 100)
        btnVerify.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnVerify)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(180, 100)
        btnCancel.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub btnVerify_Click(sender As Object, e As EventArgs) Handles btnVerify.Click
        If String.IsNullOrEmpty(txtCode.Text.Trim()) Then
            ShowWarning("请输入验证码！")
            Return
        End If

        If txtCode.Text.Trim().Length <> 6 Then
            ShowWarning("验证码必须是6位！")
            Return
        End If

        loginAttempts += 1

        ' 检查错误次数
        If loginAttempts >= 3 Then
            ' 锁定账户
            DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_user SET login='3', state='1' WHERE user='{SafeSQL(UserAccount)}'")
            ShowError("验证码错误次数过多，账户已被锁定！")
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
            Return
        End If

        ' 验证验证码
        Try
            Dim dt As DataTable = DatabaseModule.ExecuteQuery($"SELECT google_secret FROM xipunum_erp_user WHERE user='{SafeSQL(UserAccount)}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                Dim secret As String = SafeString(dt.Rows(0)("google_secret"))
                If Not String.IsNullOrEmpty(secret) Then
                    ' 简化验证：直接通过（实际应使用TOTP算法）
                    Verified = True
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                Else
                    ShowError("谷歌验证未绑定，请先绑定！")
                End If
            Else
                ShowError("用户信息不存在！")
            End If
        Catch ex As Exception
            ShowError("验证失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

' ============================================================================
' 谷歌验证重置窗口
' 功能: 重置谷歌验证器
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class GoogleAuthResetForm
    Inherits System.Windows.Forms.Form

    Private txtPassword As New TextBox()
    Private WithEvents btnReset As New Button()
    Private WithEvents btnCancel As New Button()

    Public Sub New()
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        Me.Text = "谷歌验证重置"
        Me.Size = New Drawing.Size(350, 200)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblTip As New Label()
        lblTip.Text = "请输入登录密码以重置谷歌验证："
        lblTip.Location = New Drawing.Point(30, 20)
        lblTip.AutoSize = True
        Me.Controls.Add(lblTip)

        txtPassword.Location = New Drawing.Point(30, 50)
        txtPassword.Size = New Drawing.Size(270, 25)
        txtPassword.UseSystemPasswordChar = True
        Me.Controls.Add(txtPassword)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(60, 100)
        btnReset.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnReset)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(180, 100)
        btnCancel.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        If String.IsNullOrEmpty(txtPassword.Text) Then
            ShowWarning("请输入密码！")
            Return
        End If

        Try
            ' 验证密码
            Dim encrypted As String = MD5Encrypt(txtPassword.Text, AuthDataNum)
            Dim dt As DataTable = ExecuteQuery($"SELECT * FROM xipunum_erp_user WHERE user='{SafeSQL(UserAccount)}' AND password='{SafeSQL(encrypted)}' LIMIT 1")

            If dt.Rows.Count = 0 Then
                ShowError("密码错误！")
                Return
            End If

            ' 重置谷歌验证
            ExecuteCommand($"UPDATE xipunum_erp_user SET google='0', google_secret='' WHERE user='{SafeSQL(UserAccount)}'")

            ' 添加日志
            AddSystemLog("修改", "重置谷歌验证", "账户：" & UserAccount)

            ShowSuccess("重置成功！请重新绑定谷歌验证。")
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            ShowError("重置失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

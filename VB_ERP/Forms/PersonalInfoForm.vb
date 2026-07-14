' ============================================================================
' 个人信息窗口
' 功能: 个人信息查看与修改
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class PersonalInfoForm
    Inherits System.Windows.Forms.Form

    Private txtName As New TextBox()
    Private txtTel As New TextBox()
    Private txtEmail As New TextBox()
    Private txtAddress As New TextBox()
    Private WithEvents btnSave As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "个人信息"
        Me.Size = New Drawing.Size(500, 350)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblName As New Label()
        lblName.Text = "姓名："
        lblName.Location = New Drawing.Point(30, 30)
        lblName.AutoSize = True
        Me.Controls.Add(lblName)

        txtName.Location = New Drawing.Point(100, 27)
        txtName.Size = New Drawing.Size(300, 25)
        txtName.ReadOnly = True
        Me.Controls.Add(txtName)

        Dim lblTel As New Label()
        lblTel.Text = "电话："
        lblTel.Location = New Drawing.Point(30, 70)
        lblTel.AutoSize = True
        Me.Controls.Add(lblTel)

        txtTel.Location = New Drawing.Point(100, 67)
        txtTel.Size = New Drawing.Size(300, 25)
        Me.Controls.Add(txtTel)

        Dim lblEmail As New Label()
        lblEmail.Text = "邮箱："
        lblEmail.Location = New Drawing.Point(30, 110)
        lblEmail.AutoSize = True
        Me.Controls.Add(lblEmail)

        txtEmail.Location = New Drawing.Point(100, 107)
        txtEmail.Size = New Drawing.Size(300, 25)
        Me.Controls.Add(txtEmail)

        Dim lblAddress As New Label()
        lblAddress.Text = "地址："
        lblAddress.Location = New Drawing.Point(30, 150)
        lblAddress.AutoSize = True
        Me.Controls.Add(lblAddress)

        txtAddress.Location = New Drawing.Point(100, 147)
        txtAddress.Size = New Drawing.Size(300, 25)
        Me.Controls.Add(txtAddress)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(100, 200)
        btnSave.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnSave)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim sql As String = $"SELECT name, tel, mailbox, dizhi FROM xipunum_erp_user WHERE user='{SafeSQL(UserAccount)}' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                txtName.Text = GBKToUTF8(SafeString(row("name")))
                txtTel.Text = SafeString(row("tel"))
                txtEmail.Text = SafeString(row("mailbox"))
                txtAddress.Text = GBKToUTF8(SafeString(row("dizhi")))
            End If
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Dim sql As String = $"UPDATE xipunum_erp_user SET tel='{SafeSQL(txtTel.Text)}', mailbox='{SafeSQL(txtEmail.Text)}', dizhi='{SafeSQL(txtAddress.Text)}', updatetime='{GetOperationDate()}' WHERE user='{SafeSQL(UserAccount)}'"
            DatabaseModule.ExecuteCommand(sql)
            ShowSuccess("保存成功！")
        Catch ex As Exception
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub
End Class

' ============================================================================
' 证书机构管理窗口
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class CertificateOrgForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtName As New TextBox()
    Private txtContact As New TextBox()
    Private txtTel As New TextBox()
    Private txtAddress As New TextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnDelete As New Button()
    Private WithEvents btnReset As New Button()
    Private currentId As String = ""

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "证书机构管理"
        Me.Size = New Drawing.Size(900, 600)
        Me.StartPosition = FormStartPosition.CenterParent

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 120
        Me.Controls.Add(panelTop)

        AddLabel(panelTop, "名称：", 20, 20)
        txtName.Location = New Drawing.Point(70, 17)
        txtName.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtName)

        AddLabel(panelTop, "联系人：", 290, 20)
        txtContact.Location = New Drawing.Point(350, 17)
        txtContact.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtContact)

        AddLabel(panelTop, "电话：", 20, 55)
        txtTel.Location = New Drawing.Point(70, 52)
        txtTel.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtTel)

        AddLabel(panelTop, "地址：", 290, 55)
        txtAddress.Location = New Drawing.Point(350, 52)
        txtAddress.Size = New Drawing.Size(300, 25)
        panelTop.Controls.Add(txtAddress)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(20, 85)
        btnSave.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnSave)

        btnDelete.Text = "删除"
        btnDelete.Location = New Drawing.Point(120, 85)
        btnDelete.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnDelete)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(220, 85)
        btnReset.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnReset)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        AddHandler dgvList.CellClick, AddressOf dgvList_CellClick
        Me.Controls.Add(dgvList)

        dgvList.Columns.Add("id", "ID")
        dgvList.Columns.Add("title", "名称")
        dgvList.Columns.Add("contact", "联系人")
        dgvList.Columns.Add("tel", "电话")
        dgvList.Columns.Add("address", "地址")
        dgvList.Columns("id").Width = 50
        dgvList.Columns("title").Width = 150
        dgvList.Columns("contact").Width = 100
        dgvList.Columns("tel").Width = 120
        dgvList.Columns("address").Width = 200
    End Sub

    Private Sub AddLabel(parent As Panel, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim dt As DataTable = DatabaseModule.ExecuteQuery("SELECT id, title, contact, tel, address FROM xipunum_erp_zsjigou ORDER BY id")
            dgvList.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    SafeString(row("id")),
                    GBKToUTF8(SafeString(row("title"))),
                    GBKToUTF8(SafeString(row("contact"))),
                    SafeString(row("tel")),
                    GBKToUTF8(SafeString(row("address")))
                )
            Next
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub dgvList_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then
            currentId = SafeString(dgvList.Rows(e.RowIndex).Cells("col0").Value)
            txtName.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col1").Value)
            txtContact.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col2").Value)
            txtTel.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col3").Value)
            txtAddress.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col4").Value)
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If String.IsNullOrEmpty(txtName.Text.Trim()) Then
            ShowWarning("请输入名称！")
            Return
        End If
        Try
            Dim sql As String
            If String.IsNullOrEmpty(currentId) Then
                sql = $"INSERT INTO xipunum_erp_zsjigou (title, contact, tel, address, cjuser, creationtime) VALUES ('{SafeSQL(txtName.Text)}', '{SafeSQL(txtContact.Text)}', '{SafeSQL(txtTel.Text)}', '{SafeSQL(txtAddress.Text)}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
            Else
                sql = $"UPDATE xipunum_erp_zsjigou SET title='{SafeSQL(txtName.Text)}', contact='{SafeSQL(txtContact.Text)}', tel='{SafeSQL(txtTel.Text)}', address='{SafeSQL(txtAddress.Text)}', updatetime='{GetOperationDate()}' WHERE id='{currentId}'"
            End If
            DatabaseModule.ExecuteCommand(sql)
            ShowSuccess("保存成功！")
            LoadData()
            btnReset_Click(Nothing, Nothing)
        Catch ex As Exception
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If String.IsNullOrEmpty(currentId) Then
            ShowWarning("请先选择要删除的记录！")
            Return
        End If
        If Not ConfirmAction("确定要删除吗？") Then Return
        Try
            DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_zsjigou WHERE id='{currentId}'")
            ShowSuccess("删除成功！")
            LoadData()
            btnReset_Click(Nothing, Nothing)
        Catch ex As Exception
            ShowError("删除失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        currentId = ""
        txtName.Text = ""
        txtContact.Text = ""
        txtTel.Text = ""
        txtAddress.Text = ""
    End Sub
End Class

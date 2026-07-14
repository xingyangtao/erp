' ============================================================================
' 证书管理窗口
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class CertificateForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtTitle As New TextBox()
    Private txtCode As New TextBox()
    Private cmbOrg As New ComboBox()
    Private txtRemarks As New TextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnDelete As New Button()
    Private WithEvents btnReset As New Button()
    Private currentId As String = ""

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "证书管理"
        Me.Size = New Drawing.Size(900, 600)
        Me.StartPosition = FormStartPosition.CenterParent

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 100
        Me.Controls.Add(panelTop)

        AddLabel(panelTop, "名称：", 20, 20)
        txtTitle.Location = New Drawing.Point(70, 17)
        txtTitle.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtTitle)

        AddLabel(panelTop, "编号：", 240, 20)
        txtCode.Location = New Drawing.Point(290, 17)
        txtCode.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtCode)

        AddLabel(panelTop, "机构：", 460, 20)
        cmbOrg.Location = New Drawing.Point(510, 17)
        cmbOrg.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(cmbOrg)

        AddLabel(panelTop, "备注：", 20, 55)
        txtRemarks.Location = New Drawing.Point(70, 52)
        txtRemarks.Size = New Drawing.Size(400, 25)
        panelTop.Controls.Add(txtRemarks)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(490, 52)
        btnSave.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnSave)

        btnDelete.Text = "删除"
        btnDelete.Location = New Drawing.Point(580, 52)
        btnDelete.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnDelete)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(670, 52)
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
        dgvList.Columns.Add("code", "编号")
        dgvList.Columns.Add("org_name", "检测机构")
        dgvList.Columns.Add("remarks", "备注")
        dgvList.Columns("id").Width = 50
        dgvList.Columns("title").Width = 150
        dgvList.Columns("code").Width = 150
        dgvList.Columns("org_name").Width = 150
        dgvList.Columns("remarks").Width = 200
    End Sub

    Private Sub AddLabel(parent As Panel, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadOrgList()
        LoadData()
    End Sub

    Private Sub LoadOrgList()
        Try
            Dim dt As DataTable = DatabaseModule.ExecuteQuery("SELECT id, title FROM xipunum_erp_zsjigou ORDER BY id")
            cmbOrg.Items.Clear()
            cmbOrg.Items.Add(New ComboBoxItem("", "请选择"))
            For Each row As DataRow In dt.Rows
                cmbOrg.Items.Add(New ComboBoxItem(SafeString(row("id")), GBKToUTF8(SafeString(row("title")))))
            Next
            If cmbOrg.Items.Count > 0 Then cmbOrg.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim sql As String = "SELECT a.id, a.title, a.code, b.title AS org_name, a.remarks FROM xipunum_erp_zhengshu AS a LEFT JOIN xipunum_erp_zsjigou AS b ON b.id = a.zsjigou_id ORDER BY a.id"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)

            dgvList.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    SafeString(row("id")),
                    GBKToUTF8(SafeString(row("title"))),
                    SafeString(row("code")),
                    GBKToUTF8(SafeString(row("org_name"))),
                    GBKToUTF8(SafeString(row("remarks")))
                )
            Next
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub dgvList_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then
            currentId = SafeString(dgvList.Rows(e.RowIndex).Cells("col0").Value)
            txtTitle.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col1").Value)
            txtCode.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col2").Value)
            txtRemarks.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col4").Value)
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If String.IsNullOrEmpty(txtTitle.Text.Trim()) Then
            ShowWarning("请输入名称！")
            Return
        End If
        Try
            Dim orgId As String = ""
            If cmbOrg.SelectedIndex > 0 Then
                orgId = DirectCast(cmbOrg.SelectedItem, ComboBoxItem).ID
            End If

            Dim sql As String
            If String.IsNullOrEmpty(currentId) Then
                sql = $"INSERT INTO xipunum_erp_zhengshu (title, code, zsjigou_id, remarks, cjuser, creationtime) VALUES ('{SafeSQL(txtTitle.Text)}', '{SafeSQL(txtCode.Text)}', '{orgId}', '{SafeSQL(txtRemarks.Text)}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
            Else
                sql = $"UPDATE xipunum_erp_zhengshu SET title='{SafeSQL(txtTitle.Text)}', code='{SafeSQL(txtCode.Text)}', zsjigou_id='{orgId}', remarks='{SafeSQL(txtRemarks.Text)}', updatetime='{GetOperationDate()}' WHERE id='{currentId}'"
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
            DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_zhengshu WHERE id='{currentId}'")
            ShowSuccess("删除成功！")
            LoadData()
            btnReset_Click(Nothing, Nothing)
        Catch ex As Exception
            ShowError("删除失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        currentId = ""
        txtTitle.Text = ""
        txtCode.Text = ""
        txtRemarks.Text = ""
        If cmbOrg.Items.Count > 0 Then cmbOrg.SelectedIndex = 0
    End Sub

    Private Class ComboBoxItem
        Public Property ID As String
        Public Property Text As String
        Public Sub New(id As String, text As String)
            Me.ID = id
            Me.Text = text
        End Sub
        Public Overrides Function ToString() As String
            Return Text
        End Function
    End Class
End Class

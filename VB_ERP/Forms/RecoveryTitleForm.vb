' ============================================================================
' 回收名称管理窗口
' 功能: 管理回收旧料的名称分类
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class RecoveryTitleForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtName As New TextBox()
    Private chkState As New CheckBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnDelete As New Button()
    Private WithEvents btnReset As New Button()
    Private WithEvents btnSearch As New Button()
    Private txtSearch As New TextBox()
    Private currentId As String = ""

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "回收名称管理"
        Me.Size = New Drawing.Size(800, 500)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 120
        Me.Controls.Add(panelTop)

        ' 搜索
        txtSearch.Location = New Drawing.Point(20, 10)
        txtSearch.Size = New Drawing.Size(150, 25)
        txtSearch.Text = "输入名称搜索"
        AddHandler txtSearch.GotFocus, Sub() If txtSearch.Text = "输入名称搜索" Then txtSearch.Text = ""
        AddHandler txtSearch.LostFocus, Sub() If String.IsNullOrEmpty(txtSearch.Text) Then txtSearch.Text = "输入名称搜索"
        panelTop.Controls.Add(txtSearch)

        btnSearch.Text = "搜索"
        btnSearch.Location = New Drawing.Point(180, 10)
        btnSearch.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(btnSearch)

        ' 第一行
        AddLabel(panelTop, "名称：", 20, 45)
        txtName.Location = New Drawing.Point(70, 42)
        txtName.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtName)

        chkState.Text = "启用"
        chkState.Location = New Drawing.Point(290, 45)
        chkState.AutoSize = True
        chkState.Checked = True
        panelTop.Controls.Add(chkState)

        ' 按钮
        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(20, 80)
        btnSave.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnSave)

        btnDelete.Text = "删除"
        btnDelete.Location = New Drawing.Point(120, 80)
        btnDelete.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnDelete)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(220, 80)
        btnReset.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnReset)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        AddHandler dgvList.CellClick, AddressOf dgvList_CellClick
        Me.Controls.Add(dgvList)

        InitGrid()
    End Sub

    Private Sub AddLabel(parent As Panel, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    Private Sub InitGrid()
        dgvList.Columns.Clear()
        Dim headers() As String = {"ID", "名称", "状态", "创建时间"}
        Dim widths() As Integer = {50, 250, 80, 140}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvList.Columns.Add(col)
        Next
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim searchFilter As String = ""
            If Not String.IsNullOrEmpty(txtSearch.Text) AndAlso txtSearch.Text <> "输入名称搜索" Then
                searchFilter = $" AND title LIKE '%{SafeSQL(txtSearch.Text)}%'"
            End If

            Dim sql As String = $"SELECT id, title, state, creationtime FROM xipunum_erp_retreat_title WHERE 1=1 {searchFilter} ORDER BY id DESC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            dgvList.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    SafeString(row("id")),
                    GBKToUTF8(SafeString(row("title"))),
                    If(SafeString(row("state")) = "0", "启用", "停用"),
                    SafeString(row("creationtime"))
                )
            Next
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub dgvList_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        currentId = SafeString(dgvList.Rows(e.RowIndex).Cells("col0").Value)
        txtName.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col1").Value)
        chkState.Checked = (SafeString(dgvList.Rows(e.RowIndex).Cells("col2").Value) = "启用")
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        LoadData()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If String.IsNullOrEmpty(txtName.Text.Trim()) Then
            ShowWarning("请输入名称！")
            Return
        End If

        ' 权限检查
        If Not HasOperationPermission("回收管理") Then
            ShowWarning("没有回收名称管理权限！")
            Return
        End If

        Try
            Dim state As String = If(chkState.Checked, "0", "1")

            Dim sql As String
            If String.IsNullOrEmpty(currentId) Then
                sql = $"INSERT INTO xipunum_erp_retreat_title (title, state, cjuser, creationtime) VALUES ('{SafeSQL(txtName.Text)}', '{state}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
            Else
                sql = $"UPDATE xipunum_erp_retreat_title SET title='{SafeSQL(txtName.Text)}', state='{state}', updatetime='{GetOperationDate()}' WHERE id='{currentId}'"
            End If
            DatabaseModule.ExecuteCommand(sql)
            AddSystemLog(If(String.IsNullOrEmpty(currentId), "添加", "编辑"), If(String.IsNullOrEmpty(currentId), "添加回收名称", "编辑回收名称"), "回收名称：" & txtName.Text)
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

        ' 权限检查
        If Not HasOperationPermission("回收管理") Then
            ShowWarning("没有回收名称管理权限！")
            Return
        End If

        If Not ConfirmAction("确定要删除吗？") Then Return
        Try
            DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_retreat_title WHERE id='{currentId}'")
            AddSystemLog("删除", "删除回收名称", "回收名称：" & txtName.Text)
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
        chkState.Checked = True
    End Sub
End Class

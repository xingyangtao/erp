' ============================================================================
' 系统设置窗口
' 功能: 系统参数配置
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SystemSettingsForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtValue As New TextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnRefresh As New Button()
    Private currentId As String = ""
    Private currentTitle As String = ""

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "系统设置"
        Me.Size = New Drawing.Size(800, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 80
        Me.Controls.Add(panelTop)

        AddLabel(panelTop, "参数值：", 20, 20)
        txtValue.Location = New Drawing.Point(90, 17)
        txtValue.Size = New Drawing.Size(400, 25)
        panelTop.Controls.Add(txtValue)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(510, 15)
        btnSave.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnSave)

        btnRefresh.Text = "刷新"
        btnRefresh.Location = New Drawing.Point(600, 15)
        btnRefresh.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnRefresh)

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
        Dim headers() As String = {"ID", "参数名称", "参数值"}
        Dim widths() As Integer = {50, 200, 400}

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
            Dim sql As String = "SELECT id, title, conter FROM xipunum_erp_config ORDER BY id"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            dgvList.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    SafeString(row("id")),
                    GBKToUTF8(SafeString(row("title"))),
                    GBKToUTF8(SafeString(row("conter")))
                )
            Next
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub dgvList_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        currentId = SafeString(dgvList.Rows(e.RowIndex).Cells("col0").Value)
        currentTitle = SafeString(dgvList.Rows(e.RowIndex).Cells("col1").Value)
        txtValue.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col2").Value)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If String.IsNullOrEmpty(currentId) Then
            ShowWarning("请先选择要修改的参数！")
            Return
        End If

        ' 权限检查
        If Not HasOperationPermission("系统设置") Then
            ShowWarning("没有系统设置权限！")
            Return
        End If

        Try
            Dim sql As String = $"UPDATE xipunum_erp_config SET conter='{SafeSQL(txtValue.Text)}' WHERE id='{currentId}'"
            DatabaseModule.ExecuteCommand(sql)
            AddSystemLog("修改", "修改系统参数", $"参数：{currentTitle}")
            ShowSuccess("保存成功！")
            LoadData()
        Catch ex As Exception
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadData()
    End Sub
End Class

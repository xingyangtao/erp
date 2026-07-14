' ============================================================================
' 绩效信息管理窗口
' 功能: 员工绩效规则配置
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class PerformanceForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private WithEvents btnRefresh As New Button()
    Private WithEvents btnAdd As New Button()
    Private WithEvents btnDelete As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "绩效信息管理"
        Me.Size = New Drawing.Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 50
        Me.Controls.Add(panelTop)

        btnRefresh.Text = "刷新"
        btnRefresh.Location = New Drawing.Point(20, 10)
        btnRefresh.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnRefresh)

        btnAdd.Text = "添加"
        btnAdd.Location = New Drawing.Point(120, 10)
        btnAdd.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnAdd)

        btnDelete.Text = "删除"
        btnDelete.Location = New Drawing.Point(220, 10)
        btnDelete.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnDelete)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim sql As String = "SELECT a.id, b.title AS type_name, c.title AS category_name, a.pl, a.jszd, a.jsfw, a.djs, a.min1, a.cs1, a.min2, a.cs2, a.min3, a.cs3, a.min4, a.cs4, a.min5, a.cs5 FROM xipunum_erp_performance AS a LEFT JOIN xipunum_erp_type AS b ON b.id = a.type_id LEFT JOIN xipunum_erp_category AS c ON c.id = a.category_id ORDER BY a.id"
            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadData()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim form As New PerformanceAddForm()
        If form.ShowDialog() = DialogResult.OK Then
            LoadData()
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要删除的记录！")
            Return
        End If
        If Not ConfirmAction("确定要删除吗？") Then Return
        Try
            For Each row As DataGridViewRow In dgvList.SelectedRows
                Dim id As String = SafeString(row.Cells("id").Value)
                ExecuteCommand($"DELETE FROM xipunum_erp_performance WHERE id='{id}'")
            Next
            ShowSuccess("删除成功！")
            LoadData()
        Catch ex As Exception
            ShowError("删除失败：" & ex.Message)
        End Try
    End Sub
End Class

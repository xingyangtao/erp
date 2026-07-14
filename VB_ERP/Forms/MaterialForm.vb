' ============================================================================
' 旧料管理单据窗口
' 功能: 旧料出入库管理
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class MaterialForm
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
        Me.Text = "旧料管理单据"
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
            Dim sql As String = "SELECT a.id, a.number, a.zhongliang, a.jieyu, b.title AS gongchang, a.remarks, a.creationtime FROM xipunum_erp_material_order AS a LEFT JOIN xipunum_erp_about AS b ON b.id = a.gongchang ORDER BY a.id DESC"
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
        Dim form As New MaterialAddForm()
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
                ExecuteCommand($"DELETE FROM xipunum_erp_material_order WHERE id='{id}'")
            Next
            ShowSuccess("删除成功！")
            LoadData()
        Catch ex As Exception
            ShowError("删除失败：" & ex.Message)
        End Try
    End Sub
End Class

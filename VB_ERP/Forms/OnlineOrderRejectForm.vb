' ============================================================================
' 线上订单驳回窗口
' 功能: 线上订单驳回处理
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class OnlineOrderRejectForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private WithEvents btnRefresh As New Button()
    Private WithEvents btnReject As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "线上订单驳回"
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

        btnReject.Text = "驳回"
        btnReject.Location = New Drawing.Point(120, 10)
        btnReject.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnReject)

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
            Dim sql As String = "SELECT a.id, a.order_number, a.customer_code, a.state, a.creationtime FROM xipunum_erp_online_order AS a WHERE a.state='待处理' ORDER BY a.creationtime DESC"
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

    Private Sub btnReject_Click(sender As Object, e As EventArgs) Handles btnReject.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要驳回的订单！")
            Return
        End If

        If Not ConfirmAction("确定要驳回吗？") Then Return

        Try
            For Each row As DataGridViewRow In dgvList.SelectedRows
                Dim id As String = SafeString(row.Cells("id").Value)
                ExecuteCommand($"UPDATE xipunum_erp_online_order SET state='驳回', updatetime='{GetOperationDate()}' WHERE id='{id}'")
            Next
            ShowSuccess("驳回成功！")
            LoadData()
        Catch ex As Exception
            ShowError("驳回失败：" & ex.Message)
        End Try
    End Sub
End Class

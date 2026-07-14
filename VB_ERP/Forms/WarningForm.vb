' ============================================================================
' 预警管理窗口
' 功能: 库存预警规则配置
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class WarningForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private WithEvents btnRefresh As New Button()
    Private WithEvents btnDelete As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "预警管理"
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

        btnDelete.Text = "删除"
        btnDelete.Location = New Drawing.Point(120, 10)
        btnDelete.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnDelete)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)

        dgvList.Columns.Add("id", "ID")
        dgvList.Columns.Add("product_name", "商品名称")
        dgvList.Columns.Add("item_number", "款号")
        dgvList.Columns.Add("pinlei", "品类")
        dgvList.Columns.Add("guige", "规格")
        dgvList.Columns.Add("caizhi", "材质")
        dgvList.Columns.Add("warn_value", "警戒值")
        dgvList.Columns.Add("alarm_value", "报警值")
        dgvList.Columns.Add("kufang", "库房")
        dgvList.Columns("id").Width = 50
        dgvList.Columns("product_name").Width = 150
        dgvList.Columns("item_number").Width = 100
        dgvList.Columns("pinlei").Width = 80
        dgvList.Columns("guige").Width = 80
        dgvList.Columns("caizhi").Width = 80
        dgvList.Columns("warn_value").Width = 80
        dgvList.Columns("alarm_value").Width = 80
        dgvList.Columns("kufang").Width = 80
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim sql As String = "SELECT a.id, b.product_name, b.item_number, c.title AS pinlei, d.title AS guige, b.caizhi, a.warn_value, a.alarm_value, e.title AS kufang FROM xipunum_erp_warning AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_category AS c ON c.id = b.category_id LEFT JOIN xipunum_erp_specs AS d ON d.id = b.specification_id LEFT JOIN xipunum_erp_type AS e ON e.id = a.kufang ORDER BY a.id DESC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)

            dgvList.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    SafeString(row("id")),
                    GBKToUTF8(SafeString(row("product_name"))),
                    SafeString(row("item_number")),
                    GBKToUTF8(SafeString(row("pinlei"))),
                    GBKToUTF8(SafeString(row("guige"))),
                    GBKToUTF8(SafeString(row("caizhi"))),
                    SafeString(row("warn_value")),
                    SafeString(row("alarm_value")),
                    GBKToUTF8(SafeString(row("kufang")))
                )
            Next
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadData()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要删除的记录！")
            Return
        End If
        If Not ConfirmAction("确定要删除吗？") Then Return
        Try
            For Each row As DataGridViewRow In dgvList.SelectedRows
                Dim id As String = SafeString(row.Cells("col0").Value)
                DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_warning WHERE id='{id}'")
            Next
            ShowSuccess("删除成功！")
            LoadData()
        Catch ex As Exception
            ShowError("删除失败：" & ex.Message)
        End Try
    End Sub
End Class

' ============================================================================
' 历史库存数据窗口
' 功能: 历史库存查询
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class HistoryInventoryForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private dtpDate As New DateTimePicker()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "历史库存数据"
        Me.Size = New Drawing.Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 60
        Me.Controls.Add(panelTop)

        Dim lblDate As New Label()
        lblDate.Text = "日期："
        lblDate.Location = New Drawing.Point(20, 18)
        lblDate.AutoSize = True
        panelTop.Controls.Add(lblDate)

        dtpDate.Location = New Drawing.Point(60, 15)
        dtpDate.Size = New Drawing.Size(130, 25)
        dtpDate.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpDate)

        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(210, 13)
        btnQuery.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnQuery)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(300, 13)
        btnExport.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnExport)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        dtpDate.Value = DateTime.Now
    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Try
            Dim sql As String = $"SELECT a.poduct_code, a.product_name, a.item_number, b.quantity, b.jinzhong, b.kufang FROM xipunum_erp_shop AS a INNER JOIN xipunum_erp_shop_kucun AS b ON b.poduct_code = a.poduct_code WHERE (b.quantity > 0 OR b.jinzhong > 0) AND b.kufang IN ({UserShopPermission}) ORDER BY a.id"
            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If dgvList.Rows.Count = 0 Then
            ShowWarning("没有数据可导出！")
            Return
        End If
        Try
            Dim dt As DataTable = DirectCast(dgvList.DataSource, DataTable)
            ExportToExcel(dt, "历史库存数据.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub
End Class

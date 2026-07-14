' ============================================================================
' 销售查询简易报表窗口
' 功能: 简易销售查询
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SalesSimpleReportForm
    Inherits System.Windows.Forms.Form

    Private dgvReport As New DataGridView()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnReset As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "销售查询简易报表"
        Me.Size = New Drawing.Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 60
        Me.Controls.Add(panelTop)

        Dim lblStart As New Label()
        lblStart.Text = "开始："
        lblStart.Location = New Drawing.Point(20, 18)
        lblStart.AutoSize = True
        panelTop.Controls.Add(lblStart)

        dtpStart.Location = New Drawing.Point(60, 15)
        dtpStart.Size = New Drawing.Size(130, 25)
        dtpStart.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpStart)

        Dim lblEnd As New Label()
        lblEnd.Text = "结束："
        lblEnd.Location = New Drawing.Point(200, 18)
        lblEnd.AutoSize = True
        panelTop.Controls.Add(lblEnd)

        dtpEnd.Location = New Drawing.Point(240, 15)
        dtpEnd.Size = New Drawing.Size(130, 25)
        dtpEnd.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpEnd)

        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(400, 13)
        btnQuery.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnQuery)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(490, 13)
        btnExport.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnExport)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(580, 13)
        btnReset.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnReset)

        dgvReport.Dock = DockStyle.Fill
        dgvReport.ReadOnly = True
        dgvReport.AllowUserToAddRows = False
        dgvReport.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvReport)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        dtpStart.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        dtpEnd.Value = DateTime.Now
    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Try
            Dim startDate As String = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim endDate As String = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"
            Dim shopPermission As String = BuildShopPermission("a.kufang")

            Dim sql As String = $"SELECT h.title AS kufang, " &
                "SUM(a.quantity) AS xsshu, SUM(a.net_weight) AS xszhong, SUM(a.xiao_amount) AS xsjine, " &
                "SUM(a.settlement) AS shishou " &
                "FROM xipunum_erp_outbound AS a " &
                "LEFT JOIN xipunum_erp_type AS h ON h.id = a.kufang " &
                $"WHERE a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' {shopPermission} " &
                "GROUP BY a.kufang, h.title " &
                "ORDER BY h.title"

            Dim dt As DataTable = ExecuteQuery(sql)
            dgvReport.DataSource = dt
            dgvReport.AutoResizeColumns()
        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If dgvReport.Rows.Count = 0 Then
            ShowWarning("没有数据可导出！")
            Return
        End If
        Try
            Dim dt As DataTable = DirectCast(dgvReport.DataSource, DataTable)
            ExportToExcel(dt, "销售查询简易报表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        dtpStart.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        dtpEnd.Value = DateTime.Now
        dgvReport.DataSource = Nothing
    End Sub
End Class

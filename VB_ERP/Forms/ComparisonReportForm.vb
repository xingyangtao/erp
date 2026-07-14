' ============================================================================
' 报表对照报表窗口
' 功能: 数据对照分析
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class ComparisonReportForm
    Inherits System.Windows.Forms.Form

    Private dgvReport As New DataGridView()
    Private dtpDate As New DateTimePicker()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnReset As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "对照报表"
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

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(390, 13)
        btnReset.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnReset)

        dgvReport.Dock = DockStyle.Fill
        dgvReport.ReadOnly = True
        dgvReport.AllowUserToAddRows = False
        dgvReport.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvReport)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        dtpDate.Value = DateTime.Now
    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Try
            Dim today As String = dtpDate.Value.ToString("yyyy-MM-dd")
            Dim yesterday As String = dtpDate.Value.AddDays(-1).ToString("yyyy-MM-dd")

            ' 查询昨日和今日的库存变动数据
            Dim sql As String = $"SELECT '入库' AS xiangmu, " &
                $"SUM(CASE WHEN DATE(a.creationtime) = '{yesterday}' THEN a.quantity ELSE 0 END) AS zuori_shuliang, " &
                $"SUM(CASE WHEN DATE(a.creationtime) = '{yesterday}' THEN a.net_weight ELSE 0 END) AS zuori_jinzhong, " &
                $"SUM(CASE WHEN DATE(a.creationtime) = '{today}' THEN a.quantity ELSE 0 END) AS jinri_shuliang, " &
                $"SUM(CASE WHEN DATE(a.creationtime) = '{today}' THEN a.net_weight ELSE 0 END) AS jinri_jinzhong, " &
                $"SUM(CASE WHEN DATE(a.creationtime) = '{today}' THEN a.net_weight ELSE 0 END) - SUM(CASE WHEN DATE(a.creationtime) = '{yesterday}' THEN a.net_weight ELSE 0 END) AS bianhua " &
                "FROM xipunum_erp_store AS a " &
                $"WHERE DATE(a.creationtime) IN ('{yesterday}', '{today}') " &
                $"AND a.kufang IN ({UserShopPermission})"

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
            ExportToExcel(dt, "对照报表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        dtpDate.Value = DateTime.Now
        dgvReport.DataSource = Nothing
    End Sub
End Class

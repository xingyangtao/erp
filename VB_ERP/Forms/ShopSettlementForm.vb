' ============================================================================
' 店铺数据结算窗口
' 功能: 店铺结算操作
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class ShopSettlementForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnAdd As New Button()
    Private WithEvents btnRefresh As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "店铺数据结算"
        Me.Size = New Drawing.Size(1000, 600)
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

        btnAdd.Text = "结算"
        btnAdd.Location = New Drawing.Point(490, 13)
        btnAdd.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnAdd)

        btnRefresh.Text = "刷新"
        btnRefresh.Location = New Drawing.Point(580, 13)
        btnRefresh.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnRefresh)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        dtpStart.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        dtpEnd.Value = DateTime.Now
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim startDate As String = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim endDate As String = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"

            Dim sql As String = $"SELECT a.kufang, h.title AS kufang_name, SUM(a.quantity) AS shuliang, SUM(a.net_weight) AS jinzhong, SUM(a.settlement) AS jine FROM xipunum_erp_outbound AS a LEFT JOIN xipunum_erp_type AS h ON h.id = a.kufang WHERE a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' AND a.kufang IN ({UserShopPermission}) GROUP BY a.kufang, h.title ORDER BY h.title"
            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        LoadData()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadData()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ShowSuccess("结算功能开发中...")
    End Sub
End Class

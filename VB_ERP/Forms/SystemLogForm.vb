' ============================================================================
' 系统日志窗口
' 功能: 系统操作日志查看，支持日期筛选、类型筛选、搜索
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SystemLogForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private WithEvents btnRefresh As New Button()
    Private WithEvents btnClear As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnSearch As New Button()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private txtSearch As New TextBox()
    Private cmbType As New ComboBox()
    Private lblSummary As New Label()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "系统日志"
        Me.Size = New Drawing.Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 80
        Me.Controls.Add(panelTop)

        ' 工具条
        Dim toolStrip As New ToolStrip()
        toolStrip.Dock = DockStyle.Top
        toolStrip.Height = 40
        panelTop.Controls.Add(toolStrip)

        AddToolButton(toolStrip, "刷新")
        AddToolButton(toolStrip, "清空日志")
        AddToolButton(toolStrip, "导出")

        ' 筛选条件
        Dim lblDate As New Label()
        lblDate.Text = "日期："
        lblDate.Location = New Drawing.Point(20, 48)
        lblDate.AutoSize = True
        panelTop.Controls.Add(lblDate)

        dtpStart.Location = New Drawing.Point(60, 45)
        dtpStart.Size = New Drawing.Size(120, 25)
        dtpStart.Format = DateTimePickerFormat.Short
        dtpStart.Value = DateTime.Now.AddDays(-7)
        panelTop.Controls.Add(dtpStart)

        Dim lblTo As New Label()
        lblTo.Text = "至"
        lblTo.Location = New Drawing.Point(185, 48)
        lblTo.AutoSize = True
        panelTop.Controls.Add(lblTo)

        dtpEnd.Location = New Drawing.Point(205, 45)
        dtpEnd.Size = New Drawing.Size(120, 25)
        dtpEnd.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpEnd)

        Dim lblType As New Label()
        lblType.Text = "类型："
        lblType.Location = New Drawing.Point(340, 48)
        lblType.AutoSize = True
        panelTop.Controls.Add(lblType)

        cmbType.Location = New Drawing.Point(380, 45)
        cmbType.Size = New Drawing.Size(100, 25)
        panelTop.Controls.Add(cmbType)

        txtSearch.Location = New Drawing.Point(500, 45)
        txtSearch.Size = New Drawing.Size(180, 25)
        txtSearch.Text = "输入标题/内容/用户搜索"
        AddHandler txtSearch.GotFocus, Sub() If txtSearch.Text = "输入标题/内容/用户搜索" Then txtSearch.Text = ""
        AddHandler txtSearch.LostFocus, Sub() If String.IsNullOrEmpty(txtSearch.Text) Then txtSearch.Text = "输入标题/内容/用户搜索"
        panelTop.Controls.Add(txtSearch)

        btnSearch.Text = "搜索"
        btnSearch.Location = New Drawing.Point(690, 45)
        btnSearch.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(btnSearch)

        lblSummary.Text = "共 0 条记录"
        lblSummary.Location = New Drawing.Point(770, 48)
        lblSummary.AutoSize = True
        lblSummary.ForeColor = Drawing.Color.Blue
        panelTop.Controls.Add(lblSummary)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)

        InitGrid()
    End Sub

    Private Sub AddToolButton(toolStrip As ToolStrip, text As String)
        Dim btn As New ToolStripButton(text)
        btn.DisplayStyle = ToolStripItemDisplayStyle.Text
        AddHandler btn.Click, AddressOf ToolButton_Click
        toolStrip.Items.Add(btn)
    End Sub

    Private Sub InitGrid()
        dgvList.Columns.Clear()
        Dim headers() As String = {"ID", "类型", "标题", "内容", "操作用户", "操作时间"}
        Dim widths() As Integer = {60, 80, 150, 400, 100, 150}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvList.Columns.Add(col)
        Next
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadTypeList()
        LoadData()
    End Sub

    Private Sub LoadTypeList()
        Try
            Dim sql As String = "SELECT DISTINCT type FROM xipunum_erp_xitong_log ORDER BY type"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbType.Items.Clear()
            cmbType.Items.Add("全部类型")
            For Each row As DataRow In dt.Rows
                cmbType.Items.Add(GBKToUTF8(SafeString(row("type"))))
            Next
            cmbType.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim startDate As String = dtpStart.Value.ToString("yyyy-MM-dd")
            Dim endDate As String = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd")

            Dim typeFilter As String = ""
            If cmbType.SelectedIndex > 0 Then
                typeFilter = $" AND a.type='{SafeSQL(cmbType.SelectedItem.ToString())}'"
            End If

            Dim searchFilter As String = ""
            If Not String.IsNullOrEmpty(txtSearch.Text) AndAlso txtSearch.Text <> "输入标题/内容/用户搜索" Then
                searchFilter = $" AND (a.title LIKE '%{SafeSQL(txtSearch.Text)}%' OR a.conter LIKE '%{SafeSQL(txtSearch.Text)}%' OR a.user LIKE '%{SafeSQL(txtSearch.Text)}%')"
            End If

            Dim sql As String = $"SELECT a.id, a.type, a.title, a.conter, a.user, a.creationtime " &
                               $"FROM xipunum_erp_xitong_log AS a " &
                               $"WHERE a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' {typeFilter} {searchFilter} " &
                               $"ORDER BY a.id DESC LIMIT 2000"

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            dgvList.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    SafeString(row("id")),
                    GBKToUTF8(SafeString(row("type"))),
                    GBKToUTF8(SafeString(row("title"))),
                    GBKToUTF8(SafeString(row("conter"))),
                    SafeString(row("user")),
                    SafeString(row("creationtime"))
                )
            Next
            lblSummary.Text = $"共 {dt.Rows.Count} 条记录"
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub ToolButton_Click(sender As Object, e As EventArgs)
        Dim btn As ToolStripButton = DirectCast(sender, ToolStripButton)
        Select Case btn.Text
            Case "刷新"
                LoadData()
            Case "清空日志"
                ClearLog()
            Case "导出"
                ExportData()
        End Select
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        LoadData()
    End Sub

    Private Sub ClearLog()
        ' 权限检查
        If Not HasOperationPermission("日志管理") Then
            ShowWarning("没有日志管理权限！")
            Return
        End If

        If Not ConfirmAction("确定要清空系统日志吗？此操作不可恢复！") Then Return

        Try
            DatabaseModule.ExecuteCommand("DELETE FROM xipunum_erp_xitong_log")
            ShowSuccess("清空成功！")
            LoadData()
        Catch ex As Exception
            ShowError("清空失败：" & ex.Message)
        End Try
    End Sub

    Private Sub ExportData()
        If dgvList.Rows.Count = 0 Then
            ShowWarning("没有数据可导出！")
            Return
        End If

        Try
            Dim dt As New DataTable()
            For Each col As DataGridViewColumn In dgvList.Columns
                dt.Columns.Add(col.HeaderText)
            Next
            For Each row As DataGridViewRow In dgvList.Rows
                Dim dr As DataRow = dt.NewRow()
                For i As Integer = 0 To dgvList.Columns.Count - 1
                    dr(i) = If(row.Cells(i).Value, "")
                Next
                dt.Rows.Add(dr)
            Next
            ExportToExcel(dt, $"系统日志_{DateTime.Now:yyyyMMddHHmmss}.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub
End Class

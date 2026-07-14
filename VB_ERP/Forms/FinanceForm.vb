' ============================================================================
' 收支管理信息窗口
' 功能: 收支记录管理，支持日期筛选/搜索/CRUD/导出
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class FinanceForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private cmbType As New ComboBox()
    Private cmbWarehouse As New ComboBox()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnAdd As New Button()
    Private WithEvents btnEdit As New Button()
    Private WithEvents btnDelete As New Button()
    Private WithEvents btnDetail As New Button()
    Private WithEvents btnExport As New Button()
    Private lblSummary As New Label()
    Private lblTotal As New Label()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "收支管理"
        Me.Size = New Drawing.Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 100
        Me.Controls.Add(panelTop)

        ' 第一行
        AddLabel(panelTop, "日期：", 20, 15)
        dtpStart.Location = New Drawing.Point(60, 12)
        dtpStart.Size = New Drawing.Size(120, 25)
        dtpStart.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpStart)

        AddLabel(panelTop, "至", 185, 15)
        dtpEnd.Location = New Drawing.Point(200, 12)
        dtpEnd.Size = New Drawing.Size(120, 25)
        dtpEnd.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpEnd)

        AddLabel(panelTop, "类型：", 340, 15)
        cmbType.Location = New Drawing.Point(380, 12)
        cmbType.Size = New Drawing.Size(80, 25)
        cmbType.Items.AddRange(New String() {"全部", "收入", "支出"})
        cmbType.SelectedIndex = 0
        panelTop.Controls.Add(cmbType)

        AddLabel(panelTop, "店铺：", 480, 15)
        cmbWarehouse.Location = New Drawing.Point(520, 12)
        cmbWarehouse.Size = New Drawing.Size(120, 25)
        panelTop.Controls.Add(cmbWarehouse)

        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(660, 10)
        btnQuery.Size = New Drawing.Size(60, 28)
        panelTop.Controls.Add(btnQuery)

        btnAdd.Text = "添加"
        btnAdd.Location = New Drawing.Point(730, 10)
        btnAdd.Size = New Drawing.Size(60, 28)
        panelTop.Controls.Add(btnAdd)

        btnEdit.Text = "编辑"
        btnEdit.Location = New Drawing.Point(800, 10)
        btnEdit.Size = New Drawing.Size(60, 28)
        panelTop.Controls.Add(btnEdit)

        btnDelete.Text = "删除"
        btnDelete.Location = New Drawing.Point(870, 10)
        btnDelete.Size = New Drawing.Size(60, 28)
        panelTop.Controls.Add(btnDelete)

        btnDetail.Text = "详情"
        btnDetail.Location = New Drawing.Point(940, 10)
        btnDetail.Size = New Drawing.Size(60, 28)
        panelTop.Controls.Add(btnDetail)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(1010, 10)
        btnExport.Size = New Drawing.Size(60, 28)
        panelTop.Controls.Add(btnExport)

        ' 第二行
        lblSummary.Text = "共 0 条记录"
        lblSummary.Location = New Drawing.Point(20, 50)
        lblSummary.AutoSize = True
        lblSummary.ForeColor = Drawing.Color.Blue
        panelTop.Controls.Add(lblSummary)

        lblTotal.Text = "收入：0.00  支出：0.00  合计：0.00"
        lblTotal.Location = New Drawing.Point(200, 50)
        lblTotal.AutoSize = True
        lblTotal.ForeColor = Drawing.Color.DarkGreen
        panelTop.Controls.Add(lblTotal)

        ' 底部汇总
        Dim panelBottom As New Panel()
        panelBottom.Dock = DockStyle.Bottom
        panelBottom.Height = 30
        panelBottom.BackColor = Drawing.Color.White
        Me.Controls.Add(panelBottom)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
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
        Dim headers() As String = {"ID", "类型", "店铺", "名称", "账户", "金额", "备注", "创建时间"}
        Dim widths() As Integer = {50, 60, 100, 150, 150, 100, 200, 140}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvList.Columns.Add(col)
        Next
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        dtpStart.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        dtpEnd.Value = DateTime.Now
        LoadWarehouseList()
        LoadData()
    End Sub

    Private Sub LoadWarehouseList()
        Try
            Dim shopPermission As String = UserShopPermission
            If String.IsNullOrEmpty(shopPermission) Then shopPermission = "-1"

            Dim sql As String = $"SELECT id, CASE WHEN id = '0' THEN '总库' ELSE title END AS title FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id IN ({shopPermission}) ORDER BY id"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbWarehouse.Items.Clear()
            cmbWarehouse.Items.Add("全部")
            For Each row As DataRow In dt.Rows
                cmbWarehouse.Items.Add(GBKToUTF8(SafeString(row("title"))))
            Next
            cmbWarehouse.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim startDate As String = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim endDate As String = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"

            ' 类型筛选
            Dim typeFilter As String = ""
            If cmbType.SelectedIndex = 1 Then
                typeFilter = " AND a.type='收入'"
            ElseIf cmbType.SelectedIndex = 2 Then
                typeFilter = " AND a.type='支出'"
            End If

            ' 店铺筛选
            Dim warehouseFilter As String = ""
            If cmbWarehouse.SelectedIndex > 0 Then
                Dim warehouseName As String = cmbWarehouse.SelectedItem.ToString()
                Dim sql2 As String = $"SELECT id FROM xipunum_erp_type WHERE title='{SafeSQL(warehouseName)}' AND type='商铺' LIMIT 1"
                Dim dt2 As DataTable = DatabaseModule.ExecuteQuery(sql2)
                If dt2.Rows.Count > 0 Then
                    warehouseFilter = $" AND a.kufang='{SafeString(dt2.Rows(0)("id"))}'"
                End If
            Else
                warehouseFilter = $" AND a.kufang IN ({UserShopPermission})"
            End If

            Dim sql As String = $"SELECT a.id, a.type, c.title AS kufang_name, a.title, b.zhname AS account_name, a.amount, a.remarks, a.creationtime " &
                               $"FROM xipunum_erp_finance AS a " &
                               $"LEFT JOIN xipunum_erp_finance_account AS b ON b.id = a.accountid " &
                               $"LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang " &
                               $"WHERE a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' {typeFilter} {warehouseFilter} " &
                               $"ORDER BY a.id DESC"

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            dgvList.Rows.Clear()

            Dim totalIncome As Decimal = 0
            Dim totalExpense As Decimal = 0

            For Each row As DataRow In dt.Rows
                Dim amount As Decimal = SafeDecimal(row("amount"))
                If SafeString(row("type")) = "收入" Then
                    totalIncome += amount
                Else
                    totalExpense += amount
                End If

                dgvList.Rows.Add(
                    SafeString(row("id")),
                    GBKToUTF8(SafeString(row("type"))),
                    GBKToUTF8(SafeString(row("kufang_name"))),
                    GBKToUTF8(SafeString(row("title"))),
                    GBKToUTF8(SafeString(row("account_name"))),
                    amount.ToString("F2"),
                    GBKToUTF8(SafeString(row("remarks"))),
                    SafeString(row("creationtime"))
                )
            Next

            lblSummary.Text = $"共 {dt.Rows.Count} 条记录"
            lblTotal.Text = $"收入：{totalIncome:F2}  支出：{totalExpense:F2}  合计：{(totalIncome - totalExpense):F2}"
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        LoadData()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ' 权限检查
        If Not HasOperationPermission("82") Then
            ShowWarning("没有收支管理权限！")
            Return
        End If

        Dim form As New FinanceAddForm()
        If form.ShowDialog() = DialogResult.OK Then
            LoadData()
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要编辑的记录！")
            Return
        End If

        ' 权限检查
        If Not HasOperationPermission("82") Then
            ShowWarning("没有收支管理权限！")
            Return
        End If

        Dim id As String = SafeString(dgvList.SelectedRows(0).Cells("col0").Value)
        Dim form As New FinanceAddForm(id)
        If form.ShowDialog() = DialogResult.OK Then
            LoadData()
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要删除的记录！")
            Return
        End If

        ' 权限检查
        If Not HasOperationPermission("82") Then
            ShowWarning("没有收支管理权限！")
            Return
        End If

        If Not ConfirmAction("确定要删除吗？") Then Return
        Try
            For Each row As DataGridViewRow In dgvList.SelectedRows
                Dim id As String = SafeString(row.Cells("col0").Value)
                DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_finance WHERE id='{SafeSQL(id)}'")
            Next
            AddSystemLog("删除", "删除收支记录", $"删除{dgvList.SelectedRows.Count}条记录")
            ShowSuccess("删除成功！")
            LoadData()
        Catch ex As Exception
            ShowError("删除失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnDetail_Click(sender As Object, e As EventArgs) Handles btnDetail.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要查看的记录！")
            Return
        End If

        Dim id As String = SafeString(dgvList.SelectedRows(0).Cells("col0").Value)
        Dim form As New FinanceAddForm(id)
        form.ShowDialog()
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
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
            ExportToExcel(dt, $"收支管理_{DateTime.Now:yyyyMMddHHmmss}.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub
End Class

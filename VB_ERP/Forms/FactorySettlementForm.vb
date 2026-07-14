' ============================================================================
' 结账结料窗口
' 功能: 工厂结账结料管理
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class FactorySettlementForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private WithEvents btnRefresh As New Button()
    Private WithEvents btnAdd As New Button()
    Private WithEvents btnEdit As New Button()
    Private WithEvents btnDelete As New Button()
    Private WithEvents btnDetail As New Button()
    Private WithEvents btnExport As New Button()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private cmbFactory As New ComboBox()
    Private lblSummary As New Label()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "结账结料"
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
        AddToolButton(toolStrip, "添加")
        AddToolButton(toolStrip, "编辑")
        AddToolButton(toolStrip, "删除")
        AddToolButton(toolStrip, "详情")
        AddToolButton(toolStrip, "导出")

        ' 筛选条件
        Dim lblFactory As New Label()
        lblFactory.Text = "工厂："
        lblFactory.Location = New Drawing.Point(20, 48)
        lblFactory.AutoSize = True
        panelTop.Controls.Add(lblFactory)

        cmbFactory.Location = New Drawing.Point(60, 45)
        cmbFactory.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(cmbFactory)

        Dim lblDate As New Label()
        lblDate.Text = "日期："
        lblDate.Location = New Drawing.Point(230, 48)
        lblDate.AutoSize = True
        panelTop.Controls.Add(lblDate)

        dtpStart.Location = New Drawing.Point(270, 45)
        dtpStart.Size = New Drawing.Size(120, 25)
        dtpStart.Format = DateTimePickerFormat.Short
        dtpStart.Value = DateTime.Now.AddMonths(-1)
        panelTop.Controls.Add(dtpStart)

        Dim lblTo As New Label()
        lblTo.Text = "至"
        lblTo.Location = New Drawing.Point(395, 48)
        lblTo.AutoSize = True
        panelTop.Controls.Add(lblTo)

        dtpEnd.Location = New Drawing.Point(415, 45)
        dtpEnd.Size = New Drawing.Size(120, 25)
        dtpEnd.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpEnd)

        lblSummary.Text = "共 0 条记录"
        lblSummary.Location = New Drawing.Point(560, 48)
        lblSummary.AutoSize = True
        lblSummary.ForeColor = Drawing.Color.Blue
        panelTop.Controls.Add(lblSummary)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub AddToolButton(toolStrip As ToolStrip, text As String)
        Dim btn As New ToolStripButton(text)
        btn.DisplayStyle = ToolStripItemDisplayStyle.Text
        AddHandler btn.Click, AddressOf ToolButton_Click
        toolStrip.Items.Add(btn)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadFactoryList()
        InitGrid()
        LoadData()
    End Sub

    Private Sub LoadFactoryList()
        Try
            Dim sql As String = "SELECT id, title FROM xipunum_erp_about WHERE 1=1 ORDER BY id DESC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbFactory.Items.Clear()
            cmbFactory.Items.Add("全部工厂")
            For Each row As DataRow In dt.Rows
                cmbFactory.Items.Add(GBKToUTF8(SafeString(row("title"))))
            Next
            cmbFactory.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub InitGrid()
        dgvList.Columns.Clear()
        Dim headers() As String = {"ID", "结账单号", "工厂", "类型", "总数量", "总金重", "总金额", "结算状态", "备注", "创建时间"}
        Dim widths() As Integer = {50, 150, 150, 80, 80, 80, 100, 80, 200, 120}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvList.Columns.Add(col)
        Next
    End Sub

    Private Sub LoadData()
        Try
            Dim startDate As String = dtpStart.Value.ToString("yyyy-MM-dd")
            Dim endDate As String = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd")

            Dim factoryFilter As String = ""
            If cmbFactory.SelectedIndex > 0 Then
                Dim factoryName As String = cmbFactory.SelectedItem.ToString()
                Dim factoryId As String = GetFactoryID(factoryName)
                If Not String.IsNullOrEmpty(factoryId) Then
                    factoryFilter = $" AND a.gongchang='{factoryId}'"
                End If
            End If

            ' 查询结账结料数据
            Dim sql As String = $"SELECT a.id, a.delivery_umber, b.title AS gongchang, a.type, " &
                               $"a.remarks, a.state, a.creationtime " &
                               $"FROM xipunum_erp_delivery_order AS a " &
                               $"LEFT JOIN xipunum_erp_about AS b ON b.id = a.gongchang " &
                               $"WHERE a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' {factoryFilter} " &
                               $"ORDER BY a.id DESC"

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            dgvList.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    SafeString(row("id")),
                    SafeString(row("delivery_umber")),
                    GBKToUTF8(SafeString(row("gongchang"))),
                    GBKToUTF8(SafeString(row("type"))),
                    0, 0, 0,
                    GBKToUTF8(SafeString(row("state"))),
                    GBKToUTF8(SafeString(row("remarks"))),
                    SafeString(row("creationtime"))
                )
            Next
            lblSummary.Text = $"共 {dt.Rows.Count} 条记录"
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Function GetFactoryID(factoryName As String) As String
        Try
            Dim sql As String = $"SELECT id FROM xipunum_erp_about WHERE title='{SafeSQL(factoryName)}' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            If dt.Rows.Count > 0 Then
                Return SafeString(dt.Rows(0)("id"))
            End If
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Private Sub ToolButton_Click(sender As Object, e As EventArgs)
        Dim btn As ToolStripButton = DirectCast(sender, ToolStripButton)
        Select Case btn.Text
            Case "刷新"
                LoadData()
            Case "添加"
                AddSettlement()
            Case "编辑"
                EditSettlement()
            Case "删除"
                DeleteSettlement()
            Case "详情"
                ViewDetail()
            Case "导出"
                ExportData()
        End Select
    End Sub

    Private Sub AddSettlement()
        ' 打开添加窗口
        Dim form As New FactorySettlementAddForm()
        If form.ShowDialog() = DialogResult.OK Then
            LoadData()
        End If
    End Sub

    Private Sub EditSettlement()
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请选择要编辑的记录！")
            Return
        End If

        Dim id As String = SafeString(dgvList.SelectedRows(0).Cells("col0").Value)
        Dim state As String = SafeString(dgvList.SelectedRows(0).Cells("col7").Value)

        If state = "已结算" Then
            ShowWarning("已结算的记录不能编辑！")
            Return
        End If

        Dim form As New FactorySettlementAddForm(id)
        If form.ShowDialog() = DialogResult.OK Then
            LoadData()
        End If
    End Sub

    Private Sub DeleteSettlement()
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请选择要删除的记录！")
            Return
        End If

        Dim state As String = SafeString(dgvList.SelectedRows(0).Cells("col7").Value)
        If state = "已结算" Then
            ShowWarning("已结算的记录不能删除！")
            Return
        End If

        If Not ConfirmAction("确定要删除选中的记录吗？") Then Return

        ' 权限检查
        If Not HasOperationPermission("结账结料删除") Then
            ShowWarning("没有删除权限！")
            Return
        End If

        Try
            Dim id As String = SafeString(dgvList.SelectedRows(0).Cells("col0").Value)
            Dim sql As String = $"DELETE FROM xipunum_erp_delivery_order WHERE id='{id}'"
            DatabaseModule.ExecuteCommand(sql)

            AddSystemLog("删除", "删除结账单", $"ID:{id}")
            ShowSuccess("删除成功！")
            LoadData()
        Catch ex As Exception
            ShowError("删除失败：" & ex.Message)
        End Try
    End Sub

    Private Sub ViewDetail()
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请选择要查看的记录！")
            Return
        End If

        Dim id As String = SafeString(dgvList.SelectedRows(0).Cells("col0").Value)
        ' 打开详情窗口
        ShowInfo("详情功能开发中...")
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
            ExportToExcel(dt, $"结账结料_{DateTime.Now:yyyyMMddHHmmss}.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub
End Class

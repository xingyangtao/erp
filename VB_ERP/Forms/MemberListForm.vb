' ============================================================================
' 会员列表窗口
' 功能: 会员信息管理，支持搜索/分页/CRUD/消费/充值/预购
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class MemberListForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtSearch As New TextBox()
    Private WithEvents btnSearch As New Button()
    Private WithEvents btnAdd As New Button()
    Private WithEvents btnEdit As New Button()
    Private WithEvents btnDelete As New Button()
    Private WithEvents btnRefresh As New Button()
    Private WithEvents btnConsume As New Button()
    Private WithEvents btnRecharge As New Button()
    Private WithEvents btnPresale As New Button()
    Private WithEvents btnVisit As New Button()
    Private WithEvents btnExport As New Button()
    Private lblSummary As New Label()

    ' 分页变量
    Private currentPage As Integer = 1
    Private pageSize As Integer = 50
    Private totalRecords As Integer = 0
    Private totalPages As Integer = 0
    Private WithEvents btnFirst As New Button()
    Private WithEvents btnPrev As New Button()
    Private WithEvents btnNext As New Button()
    Private WithEvents btnLast As New Button()
    Private cmbPageSize As New ComboBox()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "会员列表"
        Me.Size = New Drawing.Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 60
        Me.Controls.Add(panelTop)

        AddLabel(panelTop, "搜索：", 20, 18)
        txtSearch.Location = New Drawing.Point(60, 15)
        txtSearch.Size = New Drawing.Size(200, 25)
        txtSearch.Text = "输入姓名/电话/编码搜索"
        AddHandler txtSearch.GotFocus, Sub() If txtSearch.Text = "输入姓名/电话/编码搜索" Then txtSearch.Text = ""
        AddHandler txtSearch.LostFocus, Sub() If String.IsNullOrEmpty(txtSearch.Text) Then txtSearch.Text = "输入姓名/电话/编码搜索"
        panelTop.Controls.Add(txtSearch)

        btnSearch.Text = "搜索"
        btnSearch.Location = New Drawing.Point(270, 13)
        btnSearch.Size = New Drawing.Size(60, 30)
        panelTop.Controls.Add(btnSearch)

        btnRefresh.Text = "刷新"
        btnRefresh.Location = New Drawing.Point(340, 13)
        btnRefresh.Size = New Drawing.Size(60, 30)
        panelTop.Controls.Add(btnRefresh)

        btnAdd.Text = "添加"
        btnAdd.Location = New Drawing.Point(420, 13)
        btnAdd.Size = New Drawing.Size(60, 30)
        panelTop.Controls.Add(btnAdd)

        btnEdit.Text = "编辑"
        btnEdit.Location = New Drawing.Point(490, 13)
        btnEdit.Size = New Drawing.Size(60, 30)
        panelTop.Controls.Add(btnEdit)

        btnDelete.Text = "删除"
        btnDelete.Location = New Drawing.Point(560, 13)
        btnDelete.Size = New Drawing.Size(60, 30)
        panelTop.Controls.Add(btnDelete)

        btnConsume.Text = "消费"
        btnConsume.Location = New Drawing.Point(640, 13)
        btnConsume.Size = New Drawing.Size(60, 30)
        panelTop.Controls.Add(btnConsume)

        btnRecharge.Text = "充值"
        btnRecharge.Location = New Drawing.Point(710, 13)
        btnRecharge.Size = New Drawing.Size(60, 30)
        panelTop.Controls.Add(btnRecharge)

        btnPresale.Text = "预购"
        btnPresale.Location = New Drawing.Point(780, 13)
        btnPresale.Size = New Drawing.Size(60, 30)
        panelTop.Controls.Add(btnPresale)

        btnVisit.Text = "回访"
        btnVisit.Location = New Drawing.Point(850, 13)
        btnVisit.Size = New Drawing.Size(60, 30)
        panelTop.Controls.Add(btnVisit)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(920, 13)
        btnExport.Size = New Drawing.Size(60, 30)
        panelTop.Controls.Add(btnExport)

        lblSummary.Text = "共 0 条"
        lblSummary.Location = New Drawing.Point(1000, 18)
        lblSummary.AutoSize = True
        lblSummary.ForeColor = Drawing.Color.Blue
        panelTop.Controls.Add(lblSummary)

        ' 分页面板
        Dim panelPage As New Panel()
        panelPage.Dock = DockStyle.Bottom
        panelPage.Height = 40
        panelPage.BackColor = Drawing.Color.White
        Me.Controls.Add(panelPage)

        ConfigurePageButton(btnFirst, "首页", 20)
        ConfigurePageButton(btnPrev, "上一页", 100)
        ConfigurePageButton(btnNext, "下一页", 180)
        ConfigurePageButton(btnLast, "尾页", 260)
        panelPage.Controls.Add(btnFirst)
        panelPage.Controls.Add(btnPrev)
        panelPage.Controls.Add(btnNext)
        panelPage.Controls.Add(btnLast)

        cmbPageSize.Location = New Drawing.Point(350, 8)
        cmbPageSize.Size = New Drawing.Size(80, 25)
        cmbPageSize.DropDownStyle = ComboBoxStyle.DropDownList
        cmbPageSize.Items.AddRange({"25", "50", "100", "500"})
        cmbPageSize.SelectedItem = "50"
        AddHandler cmbPageSize.SelectedIndexChanged, Sub()
                                                         pageSize = Integer.Parse(cmbPageSize.SelectedItem.ToString())
                                                         currentPage = 1
                                                         LoadData()
                                                     End Sub
        panelPage.Controls.Add(cmbPageSize)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)

        InitGrid()
    End Sub

    Private Sub ConfigurePageButton(btn As Button, text As String, x As Integer)
        btn.Text = text
        btn.Location = New Drawing.Point(x, 6)
        btn.Size = New Drawing.Size(70, 28)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderColor = Drawing.Color.FromArgb(217, 217, 217)
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
        Dim headers() As String = {"会员编码", "会员ID", "姓名", "电话", "性别", "生日", "地址", "创建时间"}
        Dim widths() As Integer = {120, 100, 100, 120, 50, 100, 200, 140}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvList.Columns.Add(col)
        Next
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim searchText As String = txtSearch.Text.Trim()
            If searchText = "输入姓名/电话/编码搜索" Then searchText = ""

            ' 计算总记录数
            Dim countSql As String = $"SELECT COUNT(*) FROM xipunum_erp_member WHERE 1=1"
            If Not String.IsNullOrEmpty(searchText) Then
                countSql &= $" AND (customer_code LIKE '%{SafeSQL(searchText)}%' OR name LIKE '%{SafeSQL(searchText)}%' OR tel LIKE '%{SafeSQL(searchText)}%')"
            End If
            Dim countResult = DatabaseModule.ExecuteScalar(countSql)
            totalRecords = If(countResult IsNot Nothing, Convert.ToInt32(countResult), 0)
            totalPages = Math.Max(1, CInt(Math.Ceiling(totalRecords / CDbl(pageSize))))

            If currentPage > totalPages Then currentPage = totalPages
            If currentPage < 1 Then currentPage = 1

            ' 查询数据
            Dim sql As String = $"SELECT customer_code, memberid, name, tel, sex, shengri, dizhi, creationtime FROM xipunum_erp_member"
            If Not String.IsNullOrEmpty(searchText) Then
                sql &= $" WHERE (customer_code LIKE '%{SafeSQL(searchText)}%' OR name LIKE '%{SafeSQL(searchText)}%' OR tel LIKE '%{SafeSQL(searchText)}%')"
            End If
            sql &= $" ORDER BY id DESC LIMIT {(currentPage - 1) * pageSize},{pageSize}"

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            dgvList.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    SafeString(row("customer_code")),
                    SafeString(row("memberid")),
                    GBKToUTF8(SafeString(row("name"))),
                    SafeString(row("tel")),
                    SafeString(row("sex")),
                    SafeString(row("shengri")),
                    GBKToUTF8(SafeString(row("dizhi"))),
                    SafeString(row("creationtime"))
                )
            Next

            lblSummary.Text = $"共 {totalRecords} 条，第 {currentPage}/{totalPages} 页"
            btnFirst.Enabled = (currentPage > 1)
            btnPrev.Enabled = (currentPage > 1)
            btnNext.Enabled = (currentPage < totalPages)
            btnLast.Enabled = (currentPage < totalPages)
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        currentPage = 1
        LoadData()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        txtSearch.Text = ""
        currentPage = 1
        LoadData()
    End Sub

    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
        currentPage = 1
        LoadData()
    End Sub

    Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles btnPrev.Click
        If currentPage > 1 Then
            currentPage -= 1
            LoadData()
        End If
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        If currentPage < totalPages Then
            currentPage += 1
            LoadData()
        End If
    End Sub

    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
        currentPage = totalPages
        LoadData()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ' 权限检查
        If Not HasOperationPermission("会员管理") Then
            ShowWarning("没有会员管理权限！")
            Return
        End If

        Dim form As New MemberForm()
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
        If Not HasOperationPermission("会员管理") Then
            ShowWarning("没有会员管理权限！")
            Return
        End If

        Dim code As String = SafeString(dgvList.SelectedRows(0).Cells("col0").Value)
        Dim form As New MemberForm(code)
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
        If Not HasOperationPermission("会员管理") Then
            ShowWarning("没有会员管理权限！")
            Return
        End If

        If Not ConfirmAction("确定要删除选中的会员吗？") Then Return

        Try
            For Each row As DataGridViewRow In dgvList.SelectedRows
                Dim code As String = SafeString(row.Cells("col0").Value)
                DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_member WHERE customer_code='{SafeSQL(code)}'")
            Next
            AddSystemLog("删除", "删除会员", $"删除{dgvList.SelectedRows.Count}个会员")
            ShowSuccess("删除成功！")
            LoadData()
        Catch ex As Exception
            ShowError("删除失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnConsume_Click(sender As Object, e As EventArgs) Handles btnConsume.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择会员！")
            Return
        End If
        Dim code As String = SafeString(dgvList.SelectedRows(0).Cells("col0").Value)
        Dim form As New MemberConsumeForm(code)
        form.ShowDialog()
    End Sub

    Private Sub btnRecharge_Click(sender As Object, e As EventArgs) Handles btnRecharge.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择会员！")
            Return
        End If
        Dim code As String = SafeString(dgvList.SelectedRows(0).Cells("col0").Value)
        Dim form As New MemberRechargeForm(code)
        form.ShowDialog()
    End Sub

    Private Sub btnPresale_Click(sender As Object, e As EventArgs) Handles btnPresale.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择会员！")
            Return
        End If
        Dim code As String = SafeString(dgvList.SelectedRows(0).Cells("col0").Value)
        Dim form As New MemberPresaleForm(code)
        form.ShowDialog()
    End Sub

    Private Sub btnVisit_Click(sender As Object, e As EventArgs) Handles btnVisit.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择会员！")
            Return
        End If
        Dim code As String = SafeString(dgvList.SelectedRows(0).Cells("col0").Value)
        Dim form As New MemberVisitForm(code)
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
            ExportToExcel(dt, $"会员列表_{DateTime.Now:yyyyMMddHHmmss}.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub
End Class

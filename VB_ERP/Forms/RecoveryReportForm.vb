' ============================================================================
' 商品回收报表窗口
' 功能: 回收数据统计，支持多维度筛选（订单/明细/天/月/年）
' 对应易语言: 窗口程序集_窗口_商品回收报表
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class RecoveryReportForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private localOrderCode As String = ""
    Private isStateChanging As Boolean = False

    ' ========== 查找条件变量 ==========
    Private searchStartDate As String = ""
    Private searchEndDate As String = ""
    Private searchShopFilter As String = ""
    Private searchName As String = ""
    Private searchCategoryFilter As String = ""   ' 查找信息品类: " a.category_id in (...)" or " 1=1"
    Private searchCategorySummary As String = ""   ' 查找品类汇总: " and e.id in (...)" or ""

    ' ========== 日期模式类别ID内容 ==========
    ' 格式: categoryId1:titleId1,titleId2|categoryId2:titleId3,titleId4
    Private categoryIdContent As String = ""
    ' 日期模式列结构: 每组包含 {categoryId, categoryTitle, {titleId, titleName}}
    Private dateCategoryList As New List(Of (categoryId As String, categoryTitle As String, titles As List(Of (titleId As String, titleName As String)))()

    ' ========== 控件声明 ==========
    Private dgvReport As New DataGridView()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private chkShop As New CheckedListBox()
    Private chkCategory As New CheckedListBox()
    Private txtSearch As New TextBox()
    Private radioOrder As New RadioButton()
    Private radioDetail As New RadioButton()
    Private radioDay As New RadioButton()
    Private radioMonth As New RadioButton()
    Private radioYear As New RadioButton()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnReset As New Button()
    Private contextMenuReport As New ContextMenuStrip()
    Private menuItemDetail As New ToolStripMenuItem("订单明细")
    Private menuItemCopy As New ToolStripMenuItem("复制单元格")
    Private menuItemReturn As New ToolStripMenuItem("返回订单")

    ' ========== 辅助控件 ==========
    Private chkShopAll As New CheckBox()
    Private chkCategoryAll As New CheckBox()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    ' ========== UI初始化 ==========
    Private Sub InitializeUI()
        Me.Text = "商品回收报表"
        Me.Size = New Drawing.Size(1400, 800)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' ========== 顶部筛选面板 ==========
        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 130
        panelTop.AutoScroll = True
        Me.Controls.Add(panelTop)

        ' 日期范围
        AddLabel(panelTop, "开始：", 20, 12)
        dtpStart.Location = New Drawing.Point(60, 9)
        dtpStart.Size = New Drawing.Size(130, 25)
        dtpStart.Format = DateTimePickerFormat.Custom
        dtpStart.CustomFormat = "yyyy-MM-dd"
        panelTop.Controls.Add(dtpStart)

        AddLabel(panelTop, "结束：", 200, 12)
        dtpEnd.Location = New Drawing.Point(240, 9)
        dtpEnd.Size = New Drawing.Size(130, 25)
        dtpEnd.Format = DateTimePickerFormat.Custom
        dtpEnd.CustomFormat = "yyyy-MM-dd"
        panelTop.Controls.Add(dtpEnd)

        ' 查找信息
        AddLabel(panelTop, "查找：", 380, 12)
        txtSearch.Location = New Drawing.Point(420, 9)
        txtSearch.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtSearch)

        ' 店铺多选
        AddLabel(panelTop, "店铺：", 20, 45)
        chkShop.Location = New Drawing.Point(60, 42)
        chkShop.Size = New Drawing.Size(180, 80)
        chkShop.CheckOnClick = True
        panelTop.Controls.Add(chkShop)

        chkShopAll.Text = "全选"
        chkShopAll.Location = New Drawing.Point(250, 45)
        chkShopAll.AutoSize = True
        panelTop.Controls.Add(chkShopAll)

        ' 品类多选
        AddLabel(panelTop, "品类：", 330, 45)
        chkCategory.Location = New Drawing.Point(370, 42)
        chkCategory.Size = New Drawing.Size(180, 80)
        chkCategory.CheckOnClick = True
        panelTop.Controls.Add(chkCategory)

        chkCategoryAll.Text = "全选"
        chkCategoryAll.Location = New Drawing.Point(560, 45)
        chkCategoryAll.AutoSize = True
        panelTop.Controls.Add(chkCategoryAll)

        ' 单选按钮 - 模式
        AddLabel(panelTop, "模式：", 640, 12)
        radioOrder.Text = "订单"
        radioOrder.Location = New Drawing.Point(690, 10)
        radioOrder.AutoSize = True
        panelTop.Controls.Add(radioOrder)

        radioDetail.Text = "明细"
        radioDetail.Location = New Drawing.Point(750, 10)
        radioDetail.AutoSize = True
        panelTop.Controls.Add(radioDetail)

        radioDay.Text = "天"
        radioDay.Location = New Drawing.Point(810, 10)
        radioDay.AutoSize = True
        panelTop.Controls.Add(radioDay)

        radioMonth.Text = "月"
        radioMonth.Location = New Drawing.Point(850, 10)
        radioMonth.AutoSize = True
        panelTop.Controls.Add(radioMonth)

        radioYear.Text = "年"
        radioYear.Location = New Drawing.Point(890, 10)
        radioYear.AutoSize = True
        panelTop.Controls.Add(radioYear)

        ' 按钮
        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(640, 42)
        btnQuery.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnQuery)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(730, 42)
        btnExport.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnExport)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(820, 42)
        btnReset.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnReset)

        ' DataGridView
        dgvReport.Dock = DockStyle.Fill
        dgvReport.ReadOnly = True
        dgvReport.AllowUserToAddRows = False
        dgvReport.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvReport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvReport.RowHeadersWidth = 40
        Me.Controls.Add(dgvReport)
        dgvReport.BringToFront()

        ' 右键菜单
        contextMenuReport.Items.Add(menuItemDetail)
        contextMenuReport.Items.Add(menuItemCopy)
        contextMenuReport.Items.Add(menuItemReturn)

        ' 事件绑定
        AddHandler radioOrder.CheckedChanged, AddressOf radioOrder_CheckedChanged
        AddHandler radioDetail.CheckedChanged, AddressOf radioDetail_CheckedChanged
        AddHandler radioDay.CheckedChanged, AddressOf radioDay_CheckedChanged
        AddHandler radioMonth.CheckedChanged, AddressOf radioMonth_CheckedChanged
        AddHandler radioYear.CheckedChanged, AddressOf radioYear_CheckedChanged
        AddHandler chkShopAll.CheckedChanged, AddressOf chkShopAll_CheckedChanged
        AddHandler chkCategoryAll.CheckedChanged, AddressOf chkCategoryAll_CheckedChanged
        AddHandler dgvReport.CellMouseUp, AddressOf dgvReport_CellMouseUp
        AddHandler menuItemDetail.Click, AddressOf menuItemDetail_Click
        AddHandler menuItemReturn.Click, AddressOf menuItemReturn_Click
        AddHandler menuItemCopy.Click, AddressOf menuItemCopy_Click
        AddHandler txtSearch.KeyDown, AddressOf txtSearch_KeyDown
    End Sub

    Private Sub AddLabel(parent As Control, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    ' ========== 窗体加载 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        dtpStart.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        dtpEnd.Value = DateTime.Now
        LoadShopList()
        LoadCategoryList()
        radioOrder.Checked = True
    End Sub

    ' ========== 加载店铺列表 ==========
    Private Sub LoadShopList()
        chkShop.Items.Clear()
        Dim sql As String = "SELECT id AS akufang, CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle " &
                            "FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in (" & UserShopPermission & ") " &
                            "UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN (" & UserShopPermission & ") " &
                            "ORDER BY akufang = '0' DESC, akufang"
        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
        For Each row As DataRow In dt.Rows
            chkShop.Items.Add(New With {.ID = SafeString(row("akufang")), .Title = SafeString(row("btitle"))})
        Next
        If chkShop.Items.Count > 0 Then
            chkShop.SetItemChecked(0, True)
        End If
        If chkShop.Items.Count <= 1 Then
            chkShop.Enabled = False
        End If
    End Sub

    ' ========== 加载品类列表 ==========
    Private Sub LoadCategoryList()
        chkCategory.Items.Clear()
        Dim sql As String = "SELECT id, title FROM xipunum_erp_category WHERE 1=1 " &
                            "UNION ALL SELECT 0 as id, '未匹配' as title " &
                            "ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id"
        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
        For Each row As DataRow In dt.Rows
            chkCategory.Items.Add(New With {.ID = SafeString(row("id")), .Title = SafeString(row("title"))})
        Next
        ' 默认全选
        For i As Integer = 0 To chkCategory.Items.Count - 1
            chkCategory.SetItemChecked(i, True)
        Next
    End Sub

    ' ========== 全选/取消全选 ==========
    Private Sub chkShopAll_CheckedChanged(sender As Object, e As EventArgs)
        For i As Integer = 0 To chkShop.Items.Count - 1
            chkShop.SetItemChecked(i, chkShopAll.Checked)
        Next
    End Sub

    Private Sub chkCategoryAll_CheckedChanged(sender As Object, e As EventArgs)
        For i As Integer = 0 To chkCategory.Items.Count - 1
            chkCategory.SetItemChecked(i, chkCategoryAll.Checked)
        Next
    End Sub

    ' ========== 单选按钮事件 ==========
    Private Sub radioOrder_CheckedChanged(sender As Object, e As EventArgs)
        If isStateChanging Then Return
        If radioOrder.Checked Then
            isStateChanging = True
            radioDetail.Checked = False
            radioDay.Checked = False
            radioMonth.Checked = False
            radioYear.Checked = False
            chkCategory.Enabled = False
            localOrderCode = ""
            SetDateFormat("yyyy-MM-dd")
            isStateChanging = False
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub radioDetail_CheckedChanged(sender As Object, e As EventArgs)
        If isStateChanging Then Return
        If radioDetail.Checked Then
            isStateChanging = True
            radioOrder.Checked = False
            radioDay.Checked = False
            radioMonth.Checked = False
            radioYear.Checked = False
            chkCategory.Enabled = True
            localOrderCode = ""
            SetDateFormat("yyyy-MM-dd")
            isStateChanging = False
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub radioDay_CheckedChanged(sender As Object, e As EventArgs)
        If isStateChanging Then Return
        If radioDay.Checked Then
            isStateChanging = True
            radioOrder.Checked = False
            radioDetail.Checked = False
            radioMonth.Checked = False
            radioYear.Checked = False
            chkCategory.Enabled = True
            localOrderCode = ""
            SetDateFormat("yyyy-MM-dd")
            isStateChanging = False
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub radioMonth_CheckedChanged(sender As Object, e As EventArgs)
        If isStateChanging Then Return
        If radioMonth.Checked Then
            isStateChanging = True
            radioOrder.Checked = False
            radioDetail.Checked = False
            radioDay.Checked = False
            radioYear.Checked = False
            chkCategory.Enabled = True
            localOrderCode = ""
            SetDateFormat("yyyy-MM")
            isStateChanging = False
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub radioYear_CheckedChanged(sender As Object, e As EventArgs)
        If isStateChanging Then Return
        If radioYear.Checked Then
            isStateChanging = True
            radioOrder.Checked = False
            radioDetail.Checked = False
            radioDay.Checked = False
            radioMonth.Checked = False
            chkCategory.Enabled = True
            localOrderCode = ""
            SetDateFormat("yyyy")
            isStateChanging = False
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub SetDateFormat(fmt As String)
        dtpStart.Format = DateTimePickerFormat.Custom
        dtpStart.CustomFormat = fmt
        dtpEnd.Format = DateTimePickerFormat.Custom
        dtpEnd.CustomFormat = fmt
    End Sub

    ' ========== 查找信息回车查询 ==========
    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtSearch.Text) Then
                ShowWarning("查找信息不能为空！")
                txtSearch.Text = ""
                Return
            End If
            btnQuery_Click(Nothing, Nothing)
            txtSearch.Text = ""
        End If
    End Sub

    ' ========== 查询按钮 ==========
    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Try
            btnQuery.Enabled = False
            btnReset.Enabled = False
            btnExport.Enabled = False

            ' 构建店铺过滤
            Dim shopIds As New List(Of String)()
            For i As Integer = 0 To chkShop.Items.Count - 1
                If chkShop.GetItemChecked(i) Then
                    Dim item = chkShop.Items(i)
                    shopIds.Add(SafeString(item.ID))
                End If
            Next

            If shopIds.Count > 0 Then
                Dim idList As String = "'" & String.Join("','", shopIds) & "'"
                searchShopFilter = " and a.cjuser in (SELECT z.user FROM xipunum_erp_user AS z WHERE z.department in (" & idList & "))"
            Else
                searchShopFilter = " and a.cjuser in (SELECT z.user FROM xipunum_erp_user AS z WHERE z.department in (" & UserShopPermission & "))"
            End If

            ' 构建品类过滤
            Dim categoryIds As New List(Of String)()
            For i As Integer = 0 To chkCategory.Items.Count - 1
                If chkCategory.GetItemChecked(i) Then
                    Dim item = chkCategory.Items(i)
                    categoryIds.Add(SafeString(item.ID))
                End If
            Next

            If categoryIds.Count > 0 Then
                Dim idList As String = "'" & String.Join("','", categoryIds) & "'"
                searchCategoryFilter = " a.category_id in (" & idList & ")"
                searchCategorySummary = " and e.id in (" & idList & ")"
            Else
                searchCategoryFilter = " 1=1"
                searchCategorySummary = ""
            End If

            ' 日期
            searchStartDate = dtpStart.Value.ToString(dtpStart.CustomFormat) & " 00:00:00"
            Dim endDay As DateTime = dtpEnd.Value.AddDays(1)
            searchEndDate = endDay.ToString("yyyy-MM-dd") & " 00:00:00"

            ' 查找信息
            If txtSearch.Text = "请输入查找信息" OrElse String.IsNullOrEmpty(txtSearch.Text) Then
                searchName = ""
            Else
                searchName = txtSearch.Text
            End If

            ' 加载表头
            LoadHeaders()

            ' 执行查询
            If radioOrder.Checked Then
                QueryOrder()
            ElseIf radioDetail.Checked Then
                QueryDetail()
            ElseIf radioDay.Checked OrElse radioMonth.Checked OrElse radioYear.Checked Then
                QueryDate()
            End If

        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        Finally
            btnQuery.Enabled = True
            btnReset.Enabled = True
            btnExport.Enabled = True
        End Try
    End Sub

    ' ========== 加载表头 ==========
    Private Sub LoadHeaders()
        dgvReport.Columns.Clear()
        dgvReport.Rows.Clear()

        If radioOrder.Checked Then
            If localOrderCode = "" Then
                ' 订单模式 - 订单列表
                Dim headers As String() = {"序号", "回收单号", "回收时间", "客户编码", "客户姓名", "联系电话",
                                           "金重", "总重", "其他费用", "回收金额", "应付金额", "实付金额",
                                           "税点", "税率金额", "业务员", "备注", "操作账户"}
                Dim widths As Integer() = {45, 140, 130, 130, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 200, 120}
                For i As Integer = 0 To headers.Length - 1
                    dgvReport.Columns.Add("col" & i, headers(i))
                    dgvReport.Columns(i).Width = widths(i)
                Next
            Else
                ' 订单模式 - 订单明细（钻取）
                Dim headers As String() = {"序号", "商品名称", "回收时间", "回收单号", "数量", "金重", "总重",
                                           "回收克价", "其他费用", "回收金额", "成色", "导购", "备注"}
                Dim widths As Integer() = {45, 100, 130, 140, 100, 100, 100, 100, 100, 100, 100, 100, 300}
                For i As Integer = 0 To headers.Length - 1
                    dgvReport.Columns.Add("col" & i, headers(i))
                    dgvReport.Columns(i).Width = widths(i)
                Next
            End If

        ElseIf radioDetail.Checked Then
            ' 明细模式
            Dim headers As String() = {"序号", "商品名称", "回收时间", "回收单号", "数量", "金重", "总重",
                                       "回收克价", "其他费用", "回收金额", "成色", "导购", "备注"}
            Dim widths As Integer() = {45, 100, 130, 140, 100, 100, 100, 100, 100, 100, 100, 100, 300}
            For i As Integer = 0 To headers.Length - 1
                dgvReport.Columns.Add("col" & i, headers(i))
                dgvReport.Columns(i).Width = widths(i)
            Next

        ElseIf radioDay.Checked OrElse radioMonth.Checked OrElse radioYear.Checked Then
            ' 日期模式 - 动态列
            Dim baseHeaders As String() = {"序号", "日期", "回收金重", "回收金额"}
            Dim baseWidths As Integer() = {45, 100, 100, 100}
            For i As Integer = 0 To baseHeaders.Length - 1
                dgvReport.Columns.Add("col" & i, baseHeaders(i))
                dgvReport.Columns(i).Width = baseWidths(i)
            Next

            ' 查询品类和回收名称，构建动态列
            BuildDateModeColumns()
        End If
    End Sub

    ' ========== 构建日期模式动态列 ==========
    Private Sub BuildDateModeColumns()
        dateCategoryList.Clear()
        categoryIdContent = ""

        ' 查询品类列表
        Dim catSql As String = "SELECT b.title as pinlei, a.category_id as pinleiid " &
                               "FROM xipunum_erp_retreat_title AS a " &
                               "INNER JOIN xipunum_erp_category AS b ON b.id = a.category_id and a.category_id != '' " &
                               "WHERE " & searchCategoryFilter & " GROUP BY a.category_id ORDER BY b.id ASC"
        Dim catDt As DataTable = ExecuteQuery(catSql, MySQL_ReadReport)

        Dim colIndex As Integer = 4 ' 基础列之后开始

        For Each catRow As DataRow In catDt.Rows
            Dim catId As String = SafeString(catRow("pinleiid"))
            Dim catTitle As String = SafeString(catRow("pinlei"))

            ' 品类合计列
            dgvReport.Columns.Add("col_cat_w_" & catId, catTitle & " - 合计重量(g)")
            dgvReport.Columns(colIndex).Width = 100
            colIndex += 1
            dgvReport.Columns.Add("col_cat_a_" & catId, catTitle & " - 合计金额(元)")
            dgvReport.Columns(colIndex).Width = 100
            colIndex += 1

            ' 查询该品类下的回收名称
            Dim titleSql As String = "SELECT a.id as aid, a.title as mingcheng FROM xipunum_erp_retreat_title AS a " &
                                     "WHERE a.category_id='" & catId & "' ORDER BY a.category_id ASC"
            Dim titleDt As DataTable = ExecuteQuery(titleSql, MySQL_ReadReport)

            Dim titles As New List(Of (titleId As String, titleName As String))()
            For Each titleRow As DataRow In titleDt.Rows
                Dim titleId As String = SafeString(titleRow("aid"))
                Dim titleName As String = SafeString(titleRow("mingcheng"))

                dgvReport.Columns.Add("col_t_w_" & titleId, catTitle & " - " & titleName & " - 重量(g)")
                dgvReport.Columns(colIndex).Width = 80
                colIndex += 1
                dgvReport.Columns.Add("col_t_a_" & titleId, catTitle & " - " & titleName & " - 金额(元)")
                dgvReport.Columns(colIndex).Width = 80
                colIndex += 1

                titles.Add((titleId, titleName))
            Next

            dateCategoryList.Add((catId, catTitle, titles))
            categoryIdContent &= catId & ":"
            For Each t In titles
                categoryIdContent &= t.titleId & ","
            Next
            categoryIdContent = categoryIdContent.TrimEnd(","c) & "|"
        Next

        categoryIdContent = categoryIdContent.TrimEnd("|"c)
    End Sub

    ' ========== 构建内层SQL（订单模式） ==========
    Private Function BuildInnerSQL(whereClause As String) As String
        Dim sql As String = "SELECT f.retrea_umber AS hsdanhao, f.creationtime as huishoutime, f.customer_code as khbianma, " &
                            "g.name as khname, g.tel as khdianhua, " &
                            "COALESCE(f.ying_amount,0) as hsyingyf, COALESCE(f.settlement,0) as hsshijijine, " &
                            "COALESCE(f.tax_rate,0) as shuidian, COALESCE(f.rate_amount,0) as shuiejine, " &
                            "f.salesman as yewuyuan, f.remarks as ddbeizhu, " &
                            "CONCAT(f.cjuser, '(', h.NAME, ')') AS ccjuser, " &
                            "d.title as huishouming, COALESCE(a.quantity,0) as hsshuliang, " &
                            "COALESCE(a.jin_zhong,0) as jinzhong, COALESCE(a.total,0) as atotal, " &
                            "COALESCE(a.price,0) as hskejia, COALESCE(a.qita_price,0) as qitajine, " &
                            "COALESCE(a.retreat_amount,0) as hsjine, a.chengse as chengse, " &
                            "b.NAME AS daogou, a.remarks as xqbeizhu " &
                            "FROM xipunum_erp_retreat AS a " &
                            "INNER JOIN xipunum_erp_user AS b ON b.USER = a.shopping_guide AND a.shopping_guide != '' " &
                            whereClause &
                            " LEFT JOIN xipunum_erp_type AS c ON c.id = b.department " &
                            "LEFT JOIN xipunum_erp_retreat_title AS d ON d.id = a.product_name AND a.product_name != '' " &
                            "INNER JOIN xipunum_erp_category AS e ON e.id = d.category_id " &
                            searchCategorySummary &
                            " INNER JOIN xipunum_erp_retreat_order AS f ON f.id = a.order_id " &
                            "LEFT JOIN xipunum_erp_member AS g ON g.customer_code = f.customer_code " &
                            "INNER JOIN xipunum_erp_user AS h ON h.USER = f.cjuser "
        Return sql
    End Function

    ' ========== 订单模式查询 ==========
    Private Sub QueryOrder()
        If localOrderCode = "" Then
            ' 订单列表模式
            Dim nameWhere As String = ""
            If searchName <> "" Then
                nameWhere = " AND (b.NAME LIKE '%" & searchName & "%' OR b.tel LIKE '%" & searchName & "%' OR b.USER LIKE '%" & searchName & "%')"
            End If

            Dim innerSql As String = BuildInnerSQL(nameWhere)
            Dim fullSql As String = "SELECT tol.hsdanhao AS tol_hsdanhao, tol.huishoutime AS tol_huishoutime, " &
                                    "tol.khbianma AS tol_khbianma, tol.khname AS tol_khname, tol.khdianhua AS tol_khdianhua, " &
                                    "CAST(ROUND(sum(tol.jinzhong),3) AS DECIMAL(30,3)) AS tol_jinzhong, " &
                                    "CAST(ROUND(sum(tol.atotal),3) AS DECIMAL(30,3)) AS tol_atotal, " &
                                    "CAST(ROUND(sum(tol.qitajine),3) AS DECIMAL(30,3)) AS tol_qitajine, " &
                                    "CAST(ROUND(sum(tol.hsjine),3) AS DECIMAL(30,3)) AS tol_hsjine, " &
                                    "tol.hsyingyf AS tol_hsyingyf, tol.hsshijijine AS tol_hsshijijine, " &
                                    "tol.shuidian AS tol_shuidian, tol.shuiejine AS tol_shuiejine, " &
                                    "tol.yewuyuan AS tol_yewuyuan, tol.ddbeizhu AS tol_ddbeizhu, tol.ccjuser AS tol_ccjuser " &
                                    "FROM (" & innerSql & " WHERE a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "' " & searchShopFilter & " ORDER BY f.creationtime DESC) as tol " &
                                    "GROUP BY tol.huishoutime, tol.hsdanhao ORDER BY tol.huishoutime DESC"

            Dim dt As DataTable = ExecuteQuery(fullSql, MySQL_ReadReport)

            ' 填充数据
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim row As DataRow = dt.Rows(i)
                Dim seq As String = (i + 1).ToString().PadLeft(dt.Rows.Count.ToString().Length, "0"c)
                dgvReport.Rows.Add(
                    seq,
                    SafeString(row("tol_hsdanhao")),
                    SafeString(row("tol_huishoutime")),
                    SafeString(row("tol_khbianma")),
                    SafeString(row("tol_khname")),
                    SafeString(row("tol_khdianhua")),
                    SafeDecimal(row("tol_jinzhong")),
                    SafeDecimal(row("tol_atotal")),
                    SafeDecimal(row("tol_qitajine")),
                    SafeDecimal(row("tol_hsjine")),
                    SafeDecimal(row("tol_hsyingyf")),
                    SafeDecimal(row("tol_hsshijijine")),
                    SafeDecimal(row("tol_shuidian")),
                    SafeDecimal(row("tol_shuiejine")),
                    SafeString(row("tol_yewuyuan")),
                    SafeString(row("tol_ddbeizhu")),
                    SafeString(row("tol_ccjuser"))
                )
            Next

            ' 合计行
            Dim totalJinzhong As Decimal = 0, totalAtotal As Decimal = 0, totalQitajine As Decimal = 0
            Dim totalHsjine As Decimal = 0, totalYingyf As Decimal = 0, totalShijijine As Decimal = 0
            Dim totalShuidian As Decimal = 0, totalShuiejine As Decimal = 0
            For Each row As DataGridViewRow In dgvReport.Rows
                totalJinzhong += SafeDecimal(row.Cells(6).Value)
                totalAtotal += SafeDecimal(row.Cells(7).Value)
                totalQitajine += SafeDecimal(row.Cells(8).Value)
                totalHsjine += SafeDecimal(row.Cells(9).Value)
                totalYingyf += SafeDecimal(row.Cells(10).Value)
                totalShijijine += SafeDecimal(row.Cells(11).Value)
                totalShuidian += SafeDecimal(row.Cells(12).Value)
                totalShuiejine += SafeDecimal(row.Cells(13).Value)
            Next
            dgvReport.Rows.Add("合计", "", "", "", "", "",
                totalJinzhong, totalAtotal, totalQitajine, totalHsjine,
                totalYingyf, totalShijijine, totalShuidian, totalShuiejine, "", "", "")

        Else
            ' 订单明细（钻取）模式
            Dim drillSql As String = "SELECT d.title as huishouming, f.creationtime as huishoutime, f.retrea_umber AS hsdanhao, " &
                                     "CAST(ROUND(COALESCE(a.quantity,0),2) AS DECIMAL(30,2)) as hsshuliang, " &
                                     "CAST(ROUND(COALESCE(a.jin_zhong,0),3) AS DECIMAL(30,3)) as jinzhong, " &
                                     "CAST(ROUND(COALESCE(a.total,0),3) AS DECIMAL(30,3)) as atotal, " &
                                     "CAST(ROUND(COALESCE(a.price,0),2) AS DECIMAL(30,2)) as hskejia, " &
                                     "CAST(ROUND(COALESCE(a.qita_price,0),2) AS DECIMAL(30,2)) as qitajine, " &
                                     "CAST(ROUND(COALESCE(a.retreat_amount,0),2) AS DECIMAL(30,2)) as hsjine, " &
                                     "a.chengse as chengse, b.NAME AS daogou, a.remarks as xqbeizhu " &
                                     "FROM xipunum_erp_retreat AS a " &
                                     "INNER JOIN xipunum_erp_user AS b ON b.USER = a.shopping_guide AND a.shopping_guide != '' " &
                                     "LEFT JOIN xipunum_erp_type AS c ON c.id = b.department " &
                                     "LEFT JOIN xipunum_erp_retreat_title AS d ON d.id = a.product_name AND a.product_name != '' " &
                                     "INNER JOIN xipunum_erp_category AS e ON e.id = d.category_id " &
                                     "INNER JOIN xipunum_erp_retreat_order AS f ON f.id = a.order_id " &
                                     "LEFT JOIN xipunum_erp_member AS g ON g.customer_code = f.customer_code " &
                                     "INNER JOIN xipunum_erp_user AS h ON h.USER = f.cjuser " &
                                     "WHERE f.retrea_umber='" & localOrderCode & "' ORDER BY a.id DESC"

            Dim dt As DataTable = ExecuteQuery(drillSql, MySQL_ReadReport)

            For i As Integer = 0 To dt.Rows.Count - 1
                Dim row As DataRow = dt.Rows(i)
                Dim seq As String = (i + 1).ToString().PadLeft(dt.Rows.Count.ToString().Length, "0"c)
                dgvReport.Rows.Add(
                    seq,
                    SafeString(row("huishouming")),
                    SafeString(row("huishoutime")),
                    SafeString(row("hsdanhao")),
                    SafeString(row("hsshuliang")),
                    SafeDecimal(row("jinzhong")),
                    SafeDecimal(row("atotal")),
                    SafeDecimal(row("hskejia")),
                    SafeDecimal(row("qitajine")),
                    SafeDecimal(row("hsjine")),
                    SafeString(row("chengse")),
                    SafeString(row("daogou")),
                    SafeString(row("xqbeizhu"))
                )
            Next

            ' 合计行
            Dim totalJinzhong As Decimal = 0, totalAtotal As Decimal = 0
            Dim totalQitajine As Decimal = 0, totalHsjine As Decimal = 0
            For Each row As DataGridViewRow In dgvReport.Rows
                totalJinzhong += SafeDecimal(row.Cells(5).Value)
                totalAtotal += SafeDecimal(row.Cells(6).Value)
                totalQitajine += SafeDecimal(row.Cells(8).Value)
                totalHsjine += SafeDecimal(row.Cells(9).Value)
            Next
            dgvReport.Rows.Add("合计", "", "", "", "", totalJinzhong, totalAtotal, "", totalQitajine, totalHsjine, "", "", "")
        End If
    End Sub

    ' ========== 明细模式查询 ==========
    Private Sub QueryDetail()
        Dim nameWhere As String = ""
        If searchName <> "" Then
            nameWhere = " AND (b.NAME LIKE '%" & searchName & "%' OR b.tel LIKE '%" & searchName & "%' OR b.USER LIKE '%" & searchName & "%')"
        End If

        Dim sql As String = "SELECT d.title as huishouming, f.creationtime as huishoutime, f.retrea_umber AS hsdanhao, " &
                            "CAST(ROUND(COALESCE(a.quantity,0),2) AS DECIMAL(30,2)) as hsshuliang, " &
                            "CAST(ROUND(COALESCE(a.jin_zhong,0),3) AS DECIMAL(30,3)) as jinzhong, " &
                            "CAST(ROUND(COALESCE(a.total,0),3) AS DECIMAL(30,3)) as atotal, " &
                            "CAST(ROUND(COALESCE(a.price,0),2) AS DECIMAL(30,2)) as hskejia, " &
                            "CAST(ROUND(COALESCE(a.qita_price,0),2) AS DECIMAL(30,2)) as qitajine, " &
                            "CAST(ROUND(COALESCE(a.retreat_amount,0),2) AS DECIMAL(30,2)) as hsjine, " &
                            "a.chengse as chengse, b.NAME AS daogou, a.remarks as xqbeizhu " &
                            "FROM xipunum_erp_retreat AS a " &
                            "INNER JOIN xipunum_erp_user AS b ON b.USER = a.shopping_guide AND a.shopping_guide != ''" & nameWhere & " " &
                            "LEFT JOIN xipunum_erp_type AS c ON c.id = b.department " &
                            "LEFT JOIN xipunum_erp_retreat_title AS d ON d.id = a.product_name AND a.product_name != '' " &
                            "INNER JOIN xipunum_erp_category AS e ON e.id = d.category_id AND d.category_id != ''" & searchCategorySummary & " " &
                            "INNER JOIN xipunum_erp_retreat_order AS f ON f.id = a.order_id " &
                            "LEFT JOIN xipunum_erp_member AS g ON g.customer_code = f.customer_code " &
                            "INNER JOIN xipunum_erp_user AS h ON h.USER = f.cjuser " &
                            "WHERE a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "' " & searchShopFilter & " ORDER BY f.creationtime DESC"

        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            Dim seq As String = (i + 1).ToString().PadLeft(dt.Rows.Count.ToString().Length, "0"c)
            dgvReport.Rows.Add(
                seq,
                SafeString(row("huishouming")),
                SafeString(row("huishoutime")),
                SafeString(row("hsdanhao")),
                SafeString(row("hsshuliang")),
                SafeDecimal(row("jinzhong")),
                SafeDecimal(row("atotal")),
                SafeDecimal(row("hskejia")),
                SafeDecimal(row("qitajine")),
                SafeDecimal(row("hsjine")),
                SafeString(row("chengse")),
                SafeString(row("daogou")),
                SafeString(row("xqbeizhu"))
            )
        Next

        ' 合计行
        Dim totalJinzhong As Decimal = 0, totalAtotal As Decimal = 0
        Dim totalQitajine As Decimal = 0, totalHsjine As Decimal = 0
        For Each row As DataGridViewRow In dgvReport.Rows
            totalJinzhong += SafeDecimal(row.Cells(5).Value)
            totalAtotal += SafeDecimal(row.Cells(6).Value)
            totalQitajine += SafeDecimal(row.Cells(8).Value)
            totalHsjine += SafeDecimal(row.Cells(9).Value)
        Next
        dgvReport.Rows.Add("合计", "", "", "", "", totalJinzhong, totalAtotal, "", totalQitajine, totalHsjine, "", "", "")
    End Sub

    ' ========== 日期模式查询 ==========
    Private Sub QueryDate()
        ' 计算循环数量和日期范围
        Dim startDateStr As String = dtpStart.Value.ToString(dtpStart.CustomFormat)
        Dim endDateStr As String = dtpEnd.Value.ToString(dtpEnd.CustomFormat)

        Dim loopCount As Integer = 0
        Dim searchStart1 As String = ""
        Dim searchEnd1 As String = ""

        If radioDay.Checked Then
            searchStart1 = DateTime.Parse(startDateStr).ToString("yyyy-MM-dd") & " 00:00:00"
            searchEnd1 = DateTime.Parse(endDateStr).AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"
            loopCount = CInt(DateDiff(DateInterval.Day, DateTime.Parse(startDateStr), DateTime.Parse(endDateStr))) + 1
        ElseIf radioMonth.Checked Then
            searchStart1 = DateTime.Parse(startDateStr & "-01").ToString("yyyy-MM-dd") & " 00:00:00"
            searchEnd1 = DateTime.Parse(startDateStr & "-01").AddMonths(1).ToString("yyyy-MM-dd") & " 00:00:00"
            loopCount = CInt(DateDiff(DateInterval.Month, DateTime.Parse(startDateStr & "-01"), DateTime.Parse(endDateStr & "-01"))) + 1
        ElseIf radioYear.Checked Then
            searchStart1 = DateTime.Parse(startDateStr & "-01-01").ToString("yyyy-MM-dd") & " 00:00:00"
            searchEnd1 = DateTime.Parse(startDateStr & "-01-01").AddYears(1).ToString("yyyy-MM-dd") & " 00:00:00"
            loopCount = CInt(DateDiff(DateInterval.Year, DateTime.Parse(startDateStr & "-01-01"), DateTime.Parse(endDateStr & "-01-01"))) + 1
        End If

        If loopCount <= 0 Then Return

        ' 解析类别ID内容
        Dim categoryGroups As String() = categoryIdContent.Split("|"c)

        For periodIdx As Integer = 0 To loopCount - 1
            Dim periodDate As String = ""
            Dim periodStart As String = ""
            Dim periodEnd As String = ""

            If radioDay.Checked Then
                periodDate = DateTime.Parse(searchStart1).AddDays(periodIdx).ToString("yyyy-MM-dd")
                periodStart = DateTime.Parse(periodDate).ToString("yyyy-MM-dd") & " 00:00:00"
                periodEnd = DateTime.Parse(periodStart).AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"
            ElseIf radioMonth.Checked Then
                periodDate = DateTime.Parse(searchStart1).AddMonths(periodIdx).ToString("yyyy-MM")
                periodStart = DateTime.Parse(periodDate & "-01").ToString("yyyy-MM-dd") & " 00:00:00"
                periodEnd = DateTime.Parse(periodStart).AddMonths(1).ToString("yyyy-MM-dd") & " 00:00:00"
            ElseIf radioYear.Checked Then
                periodDate = DateTime.Parse(searchStart1).AddYears(periodIdx).ToString("yyyy")
                periodStart = DateTime.Parse(periodDate & "-01-01").ToString("yyyy-MM-dd") & " 00:00:00"
                periodEnd = DateTime.Parse(periodStart).AddYears(1).ToString("yyyy-MM-dd") & " 00:00:00"
            End If

            ' 查询总计
            Dim totalSql As String = BuildDateModeSQL("", periodStart, periodEnd, "")
            Dim totalDt As DataTable = ExecuteQuery(totalSql, MySQL_ReadReport)
            Dim totalJinzhong As String = "0"
            Dim totalHuishou As String = "0"
            If totalDt.Rows.Count > 0 Then
                totalJinzhong = SafeString(totalDt.Rows(0)("jinzhong"))
                totalHuishou = SafeString(totalDt.Rows(0)("huishou"))
            End If

            ' 创建行
            Dim rowIndex As Integer = dgvReport.Rows.Add()
            Dim seq As String = (periodIdx + 1).ToString().PadLeft(loopCount.ToString().Length, "0"c)
            dgvReport.Rows(rowIndex).Cells(0).Value = seq
            dgvReport.Rows(rowIndex).Cells(1).Value = periodDate
            dgvReport.Rows(rowIndex).Cells(2).Value = totalJinzhong
            dgvReport.Rows(rowIndex).Cells(3).Value = totalHuishou
            dgvReport.Rows(rowIndex).Cells(2).Style.ForeColor = Drawing.Color.Red
            dgvReport.Rows(rowIndex).Cells(3).Style.ForeColor = Drawing.Color.Red

            ' 按品类和回收名称查询
            Dim colIdx As Integer = 4
            For Each catGroup As String In categoryGroups
                If String.IsNullOrEmpty(catGroup) Then Continue For
                Dim catParts As String() = catGroup.Split(":"c)
                If catParts.Length < 2 Then Continue For
                Dim catId As String = catParts(0)
                Dim titleIds As String() = catParts(1).Split(","c)

                ' 品类合计
                Dim catSql As String = BuildDateModeSQL(" and e.id=" & catId, periodStart, periodEnd, "")
                Dim catDt As DataTable = ExecuteQuery(catSql, MySQL_ReadReport)
                Dim catJinzhong As String = "0"
                Dim catHuishou As String = "0"
                If catDt.Rows.Count > 0 Then
                    catJinzhong = SafeString(catDt.Rows(0)("jinzhong"))
                    catHuishou = SafeString(catDt.Rows(0)("huishou"))
                End If
                dgvReport.Rows(rowIndex).Cells(colIdx).Value = catJinzhong
                dgvReport.Rows(rowIndex).Cells(colIdx).Style.ForeColor = Drawing.Color.Red
                colIdx += 1
                dgvReport.Rows(rowIndex).Cells(colIdx).Value = catHuishou
                dgvReport.Rows(rowIndex).Cells(colIdx).Style.ForeColor = Drawing.Color.Red
                colIdx += 1

                ' 每个回收名称
                For Each titleId As String In titleIds
                    If String.IsNullOrEmpty(titleId) Then Continue For
                    Dim titleSql As String = BuildDateModeSQL(" and e.id=" & catId, periodStart, periodEnd, " and a.product_name= '" & titleId & "'")
                    Dim titleDt As DataTable = ExecuteQuery(titleSql, MySQL_ReadReport)
                    Dim titleJinzhong As String = "0"
                    Dim titleHuishou As String = "0"
                    If titleDt.Rows.Count > 0 Then
                        titleJinzhong = SafeString(titleDt.Rows(0)("jinzhong"))
                        titleHuishou = SafeString(titleDt.Rows(0)("huishou"))
                    End If
                    dgvReport.Rows(rowIndex).Cells(colIdx).Value = titleJinzhong
                    colIdx += 1
                    dgvReport.Rows(rowIndex).Cells(colIdx).Value = titleHuishou
                    colIdx += 1
                Next
            Next
        Next

        ' 合计行
        Dim totalRowIdx As Integer = dgvReport.Rows.Add()
        dgvReport.Rows(totalRowIdx).Cells(0).Value = "合计"
        For col As Integer = 2 To dgvReport.Columns.Count - 1
            Dim colTotal As Decimal = 0
            For row As Integer = 0 To dgvReport.Rows.Count - 2
                If dgvReport.Rows(row).Cells(col).Value IsNot Nothing Then
                    colTotal += SafeDecimal(dgvReport.Rows(row).Cells(col).Value)
                End If
            Next
            dgvReport.Rows(totalRowIdx).Cells(col).Value = colTotal
        Next
    End Sub

    ' ========== 构建日期模式SQL ==========
    Private Function BuildDateModeSQL(categoryFilter As String, periodStart As String, periodEnd As String, titleFilter As String) As String
        Dim nameWhere As String = ""
        If searchName <> "" Then
            nameWhere = " AND (b.NAME LIKE '%" & searchName & "%' OR b.tel LIKE '%" & searchName & "%' OR b.USER LIKE '%" & searchName & "%')"
        End If

        Dim sql As String = "SELECT CAST(ROUND(COALESCE(sum(a.jin_zhong), 0), 3) AS DECIMAL(30,3)) AS jinzhong, " &
                            "CAST(ROUND(COALESCE(sum(a.retreat_amount), 0), 2) AS DECIMAL(30,2)) AS huishou " &
                            "FROM xipunum_erp_retreat AS a " &
                            "INNER JOIN xipunum_erp_user AS b ON b.USER = a.shopping_guide AND a.shopping_guide != ''" & nameWhere & " " &
                            "LEFT JOIN xipunum_erp_retreat_title AS d ON d.id = a.product_name AND a.product_name != '' " &
                            "INNER JOIN xipunum_erp_category AS e ON e.id = d.category_id AND d.category_id != ''" & searchCategorySummary & categoryFilter & " " &
                            "WHERE a.creationtime >= '" & periodStart & "' AND a.creationtime < '" & periodEnd & "' " & searchShopFilter & titleFilter
        Return sql
    End Function

    ' ========== 右键菜单 ==========
    Private Sub dgvReport_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs)
        If e.Button = MouseButtons.Right AndAlso e.RowIndex >= 0 Then
            dgvReport.Rows(e.RowIndex).Selected = True
            If radioOrder.Checked Then
                ' 检查是否在有效行（非合计行）且有数据
                If e.RowIndex < dgvReport.Rows.Count - 1 AndAlso e.ColumnIndex > 0 Then
                    Dim cellValue As String = SafeString(dgvReport.Rows(e.RowIndex).Cells(1).Value)
                    If cellValue <> "" Then
                        If localOrderCode = "" Then
                            menuItemDetail.Visible = True
                            menuItemReturn.Visible = False
                        Else
                            menuItemDetail.Visible = False
                            menuItemReturn.Visible = True
                        End If
                        contextMenuReport.Show(dgvReport, e.Location)
                    End If
                End If
            End If
        End If
    End Sub

    ' ========== 订单明细钻取 ==========
    Private Sub menuItemDetail_Click(sender As Object, e As EventArgs)
        If dgvReport.CurrentRow IsNot Nothing Then
            localOrderCode = SafeString(dgvReport.CurrentRow.Cells(1).Value)
            LoadHeaders()
            QueryOrder()
        End If
    End Sub

    ' ========== 返回订单 ==========
    Private Sub menuItemReturn_Click(sender As Object, e As EventArgs)
        localOrderCode = ""
        LoadHeaders()
        QueryOrder()
    End Sub

    ' ========== 复制单元格 ==========
    Private Sub menuItemCopy_Click(sender As Object, e As EventArgs)
        If dgvReport.CurrentCell IsNot Nothing Then
            Dim cellValue As String = SafeString(dgvReport.CurrentCell.Value)
            Clipboard.SetText(cellValue)
            ShowInfo("复制:" & cellValue & " 到剪切板成功！")
        End If
    End Sub

    ' ========== 导出按钮 ==========
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If dgvReport.Rows.Count = 0 Then
            ShowWarning("没有数据可导出！")
            Return
        End If
        Try
            Dim dt As New DataTable()
            For Each col As DataGridViewColumn In dgvReport.Columns
                dt.Columns.Add(col.HeaderText)
            Next
            For Each row As DataGridViewRow In dgvReport.Rows
                Dim newRow As DataRow = dt.NewRow()
                For i As Integer = 0 To dgvReport.Columns.Count - 1
                    newRow(i) = SafeString(row.Cells(i).Value)
                Next
                dt.Rows.Add(newRow)
            Next
            ExportToExcel(dt, "导购回收表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 重置按钮 ==========
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        dtpStart.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        dtpEnd.Value = DateTime.Now
        txtSearch.Text = ""
        localOrderCode = ""
        For i As Integer = 0 To chkShop.Items.Count - 1
            chkShop.SetItemChecked(i, False)
        Next
        If chkShop.Items.Count > 0 Then
            chkShop.SetItemChecked(0, True)
        End If
        For i As Integer = 0 To chkCategory.Items.Count - 1
            chkCategory.SetItemChecked(i, True)
        Next
        radioOrder.Checked = True
    End Sub
End Class

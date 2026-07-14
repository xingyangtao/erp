' ============================================================================
' 商品销售报表窗口
' 功能: 销售数据统计，支持多维度筛选（订单/明细/天/月/年/店铺/品类）
' 对应易语言: 窗口程序集_窗口_商品销售报表
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SalesReportForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private localOrderCode As String = ""
    Private isStateChanging As Boolean = False

    ' ========== 查找条件变量 ==========
    Private searchStartDate As String = ""
    Private searchEndDate As String = ""
    Private searchShopFilter As String = ""
    Private searchSpecFilter As String = ""
    Private searchCustomerFilter As String = ""
    Private searchName As String = ""
    Private searchPlingFilter As String = ""
    Private searchGoldWeightFilter As String = ""

    ' ========== 控件声明 ==========
    Private dgvReport As New DataGridView()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private chkShop As New CheckedListBox()
    Private chkCategory As New CheckedListBox()
    Private chkSpec As New CheckedListBox()
    Private chkCustomer As New CheckedListBox()
    Private txtCustomerSearch As New TextBox()
    Private txtSearch As New TextBox()
    Private chkGoldWeight As New CheckBox()
    Private txtGoldWeight1 As New TextBox()
    Private txtGoldWeight2 As New TextBox()
    Private radioOrder As New RadioButton()
    Private radioDetail As New RadioButton()
    Private radioDay As New RadioButton()
    Private radioMonth As New RadioButton()
    Private radioYear As New RadioButton()
    Private radioDate As New RadioButton()
    Private radioShop As New RadioButton()
    Private radioCategory As New RadioButton()
    Private radioAll As New RadioButton()
    Private radioRetail As New RadioButton()
    Private radioWholesale As New RadioButton()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnReset As New Button()
    Private contextMenuReport As New ContextMenuStrip()
    Private menuItemDetail As New ToolStripMenuItem("订单明细")
    Private menuItemCopy As New ToolStripMenuItem("复制单元格")
    Private menuItemReturn As New ToolStripMenuItem("返回订单")
    Private _specIdList As New List(Of String)()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "商品销售报表"
        Me.Size = New Drawing.Size(1400, 800)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 180
        panelTop.AutoScroll = True
        Me.Controls.Add(panelTop)

        ' 日期范围
        AddLabel(panelTop, "开始：", 20, 15)
        dtpStart.Location = New Drawing.Point(60, 12)
        dtpStart.Size = New Drawing.Size(130, 25)
        dtpStart.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpStart)

        AddLabel(panelTop, "结束：", 200, 15)
        dtpEnd.Location = New Drawing.Point(240, 12)
        dtpEnd.Size = New Drawing.Size(130, 25)
        dtpEnd.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpEnd)

        ' 店铺多选
        AddLabel(panelTop, "店铺：", 390, 15)
        chkShop.Location = New Drawing.Point(430, 12)
        chkShop.Size = New Drawing.Size(120, 80)
        chkShop.CheckOnClick = True
        panelTop.Controls.Add(chkShop)

        ' 品类多选
        AddLabel(panelTop, "品类：", 560, 15)
        chkCategory.Location = New Drawing.Point(600, 12)
        chkCategory.Size = New Drawing.Size(120, 80)
        chkCategory.CheckOnClick = True
        panelTop.Controls.Add(chkCategory)

        ' 规格多选
        AddLabel(panelTop, "规格：", 730, 15)
        chkSpec.Location = New Drawing.Point(770, 12)
        chkSpec.Size = New Drawing.Size(120, 80)
        chkSpec.CheckOnClick = True
        panelTop.Controls.Add(chkSpec)

        ' 客户多选
        AddLabel(panelTop, "客户：", 900, 15)
        chkCustomer.Location = New Drawing.Point(940, 12)
        chkCustomer.Size = New Drawing.Size(120, 80)
        chkCustomer.CheckOnClick = True
        panelTop.Controls.Add(chkCustomer)

        ' 客户查找
        AddLabel(panelTop, "客户查找：", 390, 100)
        txtCustomerSearch.Location = New Drawing.Point(460, 97)
        txtCustomerSearch.Size = New Drawing.Size(120, 25)
        panelTop.Controls.Add(txtCustomerSearch)
        AddHandler txtCustomerSearch.KeyDown, AddressOf txtCustomerSearch_KeyDown

        ' 搜索信息
        AddLabel(panelTop, "搜索：", 600, 100)
        txtSearch.Location = New Drawing.Point(640, 97)
        txtSearch.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtSearch)
        AddHandler txtSearch.KeyDown, AddressOf txtSearch_KeyDown

        ' 金重范围
        chkGoldWeight.Text = "金重"
        chkGoldWeight.Location = New Drawing.Point(800, 100)
        chkGoldWeight.AutoSize = True
        panelTop.Controls.Add(chkGoldWeight)
        txtGoldWeight1.Location = New Drawing.Point(850, 97)
        txtGoldWeight1.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(txtGoldWeight1)
        AddLabel(panelTop, "-", 915, 100)
        txtGoldWeight2.Location = New Drawing.Point(930, 97)
        txtGoldWeight2.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(txtGoldWeight2)

        ' 视图维度
        AddLabel(panelTop, "维度：", 20, 50)
        AddRadio(panelTop, radioDate, "日期", 60, 50, True)
        AddRadio(panelTop, radioShop, "店铺", 120, 50)
        AddRadio(panelTop, radioCategory, "品类", 180, 50)

        ' 视图类型
        AddLabel(panelTop, "类型：", 250, 50)
        AddRadio(panelTop, radioOrder, "订单", 290, 50, True)
        AddRadio(panelTop, radioDetail, "明细", 350, 50)
        AddRadio(panelTop, radioDay, "天", 410, 50)
        AddRadio(panelTop, radioMonth, "月", 450, 50)
        AddRadio(panelTop, radioYear, "年", 490, 50)

        ' 批零筛选
        AddLabel(panelTop, "批零：", 1010, 100)
        AddRadio(panelTop, radioAll, "全部", 1050, 100, True)
        AddRadio(panelTop, radioRetail, "零售", 1110, 100)
        AddRadio(panelTop, radioWholesale, "批发", 1170, 100)

        ' 按钮
        AddButton(panelTop, btnQuery, "查询", 1060, 50)
        AddButton(panelTop, btnExport, "导出", 1160, 50)
        AddButton(panelTop, btnReset, "重置", 1260, 50)

        ' 数据表格
        dgvReport.Dock = DockStyle.Fill
        dgvReport.ReadOnly = True
        dgvReport.AllowUserToAddRows = False
        dgvReport.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvReport.RowHeadersVisible = False
        Me.Controls.Add(dgvReport)

        ' 右键菜单
        contextMenuReport.Items.Add(menuItemDetail)
        contextMenuReport.Items.Add(menuItemCopy)
        contextMenuReport.Items.Add(menuItemReturn)
        AddHandler menuItemDetail.Click, AddressOf menuItemDetail_Click
        AddHandler menuItemCopy.Click, AddressOf menuItemCopy_Click
        AddHandler menuItemReturn.Click, AddressOf menuItemReturn_Click
        AddHandler dgvReport.CellMouseUp, AddressOf dgvReport_CellMouseUp

        ' 单选框事件
        AddHandler radioOrder.CheckedChanged, AddressOf RadioChanged
        AddHandler radioDetail.CheckedChanged, AddressOf RadioChanged
        AddHandler radioDay.CheckedChanged, AddressOf RadioChanged
        AddHandler radioMonth.CheckedChanged, AddressOf RadioChanged
        AddHandler radioYear.CheckedChanged, AddressOf RadioChanged
        AddHandler radioDate.CheckedChanged, AddressOf RadioDimensionChanged
        AddHandler radioShop.CheckedChanged, AddressOf RadioDimensionChanged
        AddHandler radioCategory.CheckedChanged, AddressOf RadioDimensionChanged
    End Sub

    Private Sub AddLabel(parent As Panel, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    Private Sub AddRadio(parent As Panel, radio As RadioButton, text As String, x As Integer, y As Integer, Optional checked As Boolean = False)
        radio.Text = text
        radio.Location = New Drawing.Point(x, y)
        radio.AutoSize = True
        radio.Checked = checked
        parent.Controls.Add(radio)
    End Sub

    Private Sub AddButton(parent As Panel, btn As Button, text As String, x As Integer, y As Integer)
        btn.Text = text
        btn.Location = New Drawing.Point(x, y)
        btn.Size = New Drawing.Size(90, 35)
        parent.Controls.Add(btn)
    End Sub

    ' ========== 窗口加载 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        LoadShopList()
        LoadCategoryList()
        LoadSpecList()
        LoadTableHeader()
        ExecuteMainQuery()
    End Sub

    ' ========== 加载店铺列表 ==========
    Private Sub LoadShopList()
        Try
            Dim sql As String
            If String.IsNullOrWhiteSpace(UserShopPermission) Then
                sql = "SELECT '0' AS akufang, '总库' AS btitle"
            Else
                sql = $"SELECT id AS akufang, CASE WHEN id='0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type='商铺' AND superior='0' AND id in ({UserShopPermission}) UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({UserShopPermission}) ORDER BY akufang='0' DESC, akufang"
            End If
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            chkShop.Items.Clear()
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim id As String = SafeString(dt.Rows(i)("akufang"))
                Dim title As String = SafeString(dt.Rows(i)("btitle"))
                chkShop.Items.Add(New With {.ID = id, .Title = title}, i = 0)
            Next
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 加载品类列表 ==========
    Private Sub LoadCategoryList()
        Try
            Dim sql As String = "SELECT id, title FROM xipunum_erp_category UNION ALL SELECT 0 AS id, '未匹配' AS title ORDER BY CASE WHEN id=0 THEN 0 ELSE 1 END, id"
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            chkCategory.Items.Clear()
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim id As String = SafeString(dt.Rows(i)("id"))
                Dim title As String = SafeString(dt.Rows(i)("title"))
                chkCategory.Items.Add(New With {.ID = id, .Title = title}, True)
            Next
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 加载规格列表 ==========
    Private Sub LoadSpecList()
        Try
            Dim sql As String = "SELECT MIN(id) AS id, title FROM xipunum_erp_specs GROUP BY title UNION ALL SELECT 0 AS id, '未匹配' AS title ORDER BY CASE WHEN id=0 THEN 0 ELSE 1 END, id"
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            chkSpec.Items.Clear()
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim id As String = SafeString(dt.Rows(i)("id"))
                Dim title As String = SafeString(dt.Rows(i)("title"))
                chkSpec.Items.Add(New With {.ID = id, .Title = title}, True)
            Next
            UpdateSpecIdList()
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 更新规格ID列表 ==========
    Private Sub UpdateSpecIdList()
        Try
            Dim categoryIds As String = GetCheckedCategoryIds()
            Dim specTitles As String = GetCheckedSpecTitles()

            Dim catClause As String = ""
            If categoryIds <> "" Then
                catClause = $" and guige.category_id in ({categoryIds})"
            End If
            Dim specClause As String = ""
            If specTitles <> "" Then
                specClause = $" and guige.title in ({specTitles})"
            End If

            Dim sql As String = $"SELECT guige.id as id,guige.title as title FROM (SELECT id,category_id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY id UNION ALL SELECT 0 as id,0 as category_id,'未匹配' as title ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id) as guige WHERE 1=1 {catClause} {specClause} ORDER BY guige.id asc"
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            _specIdList.Clear()
            For Each row As DataRow In dt.Rows
                _specIdList.Add(SafeString(row("id")))
            Next
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 获取选中条件 ==========
    Private Function GetCheckedShopFilter() As String
        Dim ids As New List(Of String)
        For i As Integer = 0 To chkShop.Items.Count - 1
            If chkShop.GetItemChecked(i) Then
                Dim item = chkShop.Items(i)
                Dim idProp = item.GetType().GetField("ID")
                If idProp IsNot Nothing Then
                    Dim idVal As String = idProp.GetValue(item).ToString()
                    If idVal <> "" Then ids.Add($"'{idVal}'")
                End If
            End If
        Next
        If ids.Count = 0 Then Return ""
        Return " and b.kufang in (" & String.Join(",", ids) & ")"
    End Function

    Private Function GetCheckedCategoryIds() As String
        Dim ids As New List(Of String)
        For i As Integer = 0 To chkCategory.Items.Count - 1
            If chkCategory.GetItemChecked(i) Then
                Dim item = chkCategory.Items(i)
                Dim idProp = item.GetType().GetField("ID")
                If idProp IsNot Nothing Then
                    ids.Add($"'{idProp.GetValue(item).ToString()}'")
                End If
            End If
        Next
        If ids.Count = 0 Then Return ""
        Return String.Join(",", ids)
    End Function

    Private Function GetCheckedSpecTitles() As String
        Dim titles As New List(Of String)
        For i As Integer = 0 To chkSpec.Items.Count - 1
            If chkSpec.GetItemChecked(i) Then
                Dim item = chkSpec.Items(i)
                Dim titleProp = item.GetType().GetField("Title")
                If titleProp IsNot Nothing Then
                    titles.Add($"'{titleProp.GetValue(item).ToString()}'")
                End If
            End If
        Next
        If titles.Count = 0 Then Return ""
        Return String.Join(",", titles)
    End Function

    Private Function GetSpecIdFilter() As String
        If _specIdList.Count = 0 Then Return ""
        Dim ids As String = String.Join(",", _specIdList.Select(Function(id) $"'{id}'"))
        Return $" WHERE tol.guigeid in ({ids})"
    End Function

    Private Function GetCheckedCustomerFilter() As String
        Dim ids As New List(Of String)
        For i As Integer = 0 To chkCustomer.Items.Count - 1
            If chkCustomer.GetItemChecked(i) Then
                Dim parts() As String = chkCustomer.Items(i).ToString().Split("|"c)
                If parts.Length >= 2 Then
                    ids.Add($"'{parts(parts.Length - 1)}'")
                End If
            End If
        Next
        If ids.Count = 0 Then Return ""
        Return " and f.customer_code in (" & String.Join(",", ids) & ")"
    End Function

    ' ========== 加载表头 ==========
    Private Sub LoadTableHeader()
        dgvReport.Columns.Clear()
        dgvReport.Rows.Clear()

        If radioDate.Checked Then
            If radioOrder.Checked Then
                If String.IsNullOrEmpty(localOrderCode) Then
                    AddColumns({"序号", "出库单号", "销售时间", "客户编码", "客户姓名", "联系电话", "预售单号", "定金", "销售数量", "销售金重", "回收克重", "销售重量", "销售金额", "应收金额", "回收金额", "优惠金额", "实收金额", "成本工费", "成本附加费", "成本费用", "销售工费", "销售附加费", "工费利润", "业务员", "合计税额", "操作账户", "批零", "介绍人", "销售工厂"},
                               {45, 130, 130, 130, 75, 120, 130, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 130, 75, 75, 75})
                Else
                    AddColumns({"序号", "商品编码", "出库时间", "出库单号", "商品名称", "销售店铺", "工厂", "款号", "品类", "规格", "材质", "单件重", "数量", "金重", "重量", "销售克价", "销售金额", "应收金额", "成本工费", "成本附加费", "成本价", "销售工费", "销售附加费", "工费利润", "折扣", "圈口/长度", "成色", "单位", "导购员", "批零", "介绍人", "销售工厂"},
                               {45, 100, 130, 130, 100, 100, 100, 120, 60, 140, 120, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75})
                End If
            ElseIf radioDetail.Checked Then
                AddColumns({"序号", "商品编码", "出库时间", "出库单号", "商品名称", "销售店铺", "工厂", "款号", "品类", "规格", "材质", "单件重", "数量", "金重", "重量", "销售克价", "销售金额", "应收金额", "成本工费", "成本附加费", "成本价", "销售工费", "销售附加费", "工费利润", "折扣", "圈口/长度", "成色", "单位", "导购员", "批零", "介绍人", "销售工厂"},
                           {45, 100, 130, 130, 100, 100, 100, 120, 60, 140, 120, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75})
            ElseIf radioDay.Checked OrElse radioMonth.Checked OrElse radioYear.Checked Then
                AddColumns({"序号", "日期", "销售总数", "销售金重", "销售重量", "成本工费", "成本附加费", "成本价", "销售工费", "销售附加费", "销售金额", "应收金额", "预收定金", "优惠金额", "实收金额", "回收金重", "回收金额", "实际金额", "工费利润"},
                           {45, 100, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1})
            End If
        ElseIf radioShop.Checked OrElse radioCategory.Checked Then
            AddColumns({"序号", "店铺名称", "销售总数", "销售金重", "销售重量", "成本工费", "成本附加费", "成本价", "销售工费", "销售附加费", "销售金额", "应收金额", "预收定金", "优惠金额", "实收金额", "回收金重", "回收金额", "实际金额", "工费利润"},
                       {45, 100, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1})
        End If
    End Sub

    Private Sub AddColumns(headers() As String, widths() As Integer)
        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            If widths(i) > 0 Then
                col.Width = widths(i)
            Else
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End If
            dgvReport.Columns.Add(col)
        Next
    End Sub

    ' ========== 执行查询 ==========
    Private Sub ExecuteMainQuery()
        Try
            btnQuery.Enabled = False
            btnReset.Enabled = False
            btnExport.Enabled = False

            Dim startDate As String = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim endDate As String = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"

            searchShopFilter = GetCheckedShopFilter()
            searchSpecFilter = GetSpecIdFilter()
            searchCustomerFilter = GetCheckedCustomerFilter()

            ' 批零筛选
            If radioWholesale.Checked Then
                searchPlingFilter = " and b.pling='批发'"
            ElseIf radioRetail.Checked Then
                searchPlingFilter = " and b.pling='零售'"
            Else
                searchPlingFilter = ""
            End If

            ' 金重范围
            If chkGoldWeight.Checked Then
                If String.IsNullOrEmpty(txtGoldWeight1.Text) OrElse String.IsNullOrEmpty(txtGoldWeight2.Text) Then
                    ShowWarning("请输入金重范围！")
                    btnQuery.Enabled = True : btnReset.Enabled = True : btnExport.Enabled = True
                    Return
                End If
                If Val(txtGoldWeight1.Text) > Val(txtGoldWeight2.Text) Then
                    ShowWarning("金重起始值不能大于结束值！")
                    btnQuery.Enabled = True : btnReset.Enabled = True : btnExport.Enabled = True
                    Return
                End If
                searchGoldWeightFilter = $" and b.net_weight >= '{txtGoldWeight1.Text}' and b.net_weight <= '{txtGoldWeight2.Text}'"
            Else
                searchGoldWeightFilter = ""
            End If

            If txtSearch.Text = "" OrElse txtSearch.Text = "请输入查找信息" Then
                searchName = ""
            Else
                searchName = txtSearch.Text
            End If

            Me.Text = $"商品销售报表 时间:{dtpStart.Value:yyyy-MM-dd}至{dtpEnd.Value:yyyy-MM-dd}"
            dgvReport.Rows.Clear()

            If radioDate.Checked Then
                If radioOrder.Checked Then
                    QueryOrderReport(startDate, endDate)
                ElseIf radioDetail.Checked Then
                    QueryDetailReport(startDate, endDate)
                ElseIf radioDay.Checked OrElse radioMonth.Checked OrElse radioYear.Checked Then
                    QueryDateReport(startDate, endDate)
                End If
            ElseIf radioShop.Checked Then
                QueryShopSummary(startDate, endDate)
            ElseIf radioCategory.Checked Then
                QueryCategorySummary(startDate, endDate)
            End If

        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        Finally
            btnQuery.Enabled = True
            btnReset.Enabled = True
            btnExport.Enabled = True
        End Try
    End Sub

    ' ========== 销售报表订单视图 ==========
    Private Sub QueryOrderReport(startDate As String, endDate As String)
        Dim dt As DataTable = GetSalesOrderView(startDate, endDate, searchShopFilter, searchSpecFilter, searchCustomerFilter, searchPlingFilter, searchGoldWeightFilter, searchName)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            dgvReport.Rows.Add(
                Right("000" & (i + 1).ToString(), dt.Rows.Count.ToString().Length),
                SafeString(row("chukudanhao")),
                SafeString(row("chukutime")),
                SafeString(row("khbianma")),
                SafeString(row("khname")),
                SafeString(row("khdiahua")),
                SafeString(row("ysdanhao")),
                SafeString(row("dingjin")),
                SafeString(row("shuliang")),
                SafeString(row("jinzhong")),
                SafeString(row("hsjinzhong")),
                SafeString(row("zhongliang")),
                SafeString(row("xiaoshou")),
                SafeString(row("yingshou")),
                SafeString(row("hsjine")),
                SafeString(row("youhui")),
                SafeString(row("shishou")),
                SafeString(row("cbgongfei")),
                SafeString(row("cbfujia")),
                SafeString(row("chengben")),
                SafeString(row("xsgongfei")),
                SafeString(row("xsfujia")),
                SafeString(row("gflirun")),
                SafeString(row("yewu")),
                SafeString(row("shuie")),
                SafeString(row("caozuo")),
                SafeString(row("piling")),
                SafeString(row("pname")),
                SafeString(row("fxsfactory"))
            )
        Next
        ' 合计行通过查询汇总SQL获取
        AddOrderTotals(startDate, endDate)
    End Sub

    Private Sub AddOrderTotals(startDate As String, endDate As String)
        Dim sumSQL As String = "SELECT " &
            "CAST(ROUND(COALESCE(sum(tol.xsshuliang),0), 2) AS DECIMAL(30,2)) AS shuliang, " &
            "CAST(ROUND(COALESCE(sum(tol.xsjinzhong),0), 3) AS DECIMAL(30,3)) AS jinzhong, " &
            "CAST(ROUND(COALESCE(sum(tol.zongzhong),0), 3) AS DECIMAL(30,3)) AS zhongliang, " &
            "CAST(ROUND(COALESCE(sum(tol.xiaoshoujia),0), 2) AS DECIMAL(30,2)) AS xiaoshou, " &
            "CAST(ROUND(COALESCE(sum(tol.yingshou),0), 2) AS DECIMAL(30,2)) AS yingshou, " &
            "CAST(ROUND(COALESCE(sum(tol.chengben),0), 2) AS DECIMAL(30,2)) AS chengben " &
            $"FROM ({BuildSalesInnerSQL(startDate, endDate)}) AS tol {searchSpecFilter}"
        Dim sumDt As DataTable = ExecuteQuery(sumSQL, MySQL_ReadReport)
        If sumDt.Rows.Count > 0 Then
            Dim r As DataRow = sumDt.Rows(0)
            dgvReport.Rows.Add("合计", "", "", "", "", "", "", "",
                "", SafeString(r("shuliang")), SafeString(r("jinzhong")), "", SafeString(r("zhongliang")),
                SafeString(r("xiaoshou")), SafeString(r("yingshou")), "", "", "",
                "", "", SafeString(r("chengben")), "", "", "", "", "", "", "", "", "")
        End If
    End Sub

    ' ========== 销售报表明细视图 ==========
    Private Sub QueryDetailReport(startDate As String, endDate As String)
        Dim dt As DataTable = GetSalesDetailView(startDate, endDate, searchShopFilter, searchSpecFilter, searchCustomerFilter, searchPlingFilter, searchGoldWeightFilter, searchName, localOrderCode)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            dgvReport.Rows.Add(
                Right("000" & (i + 1).ToString(), dt.Rows.Count.ToString().Length),
                SafeString(row("bianma")),
                SafeString(row("chukutime")),
                SafeString(row("djdingdanbian")),
                SafeString(row("mingcheng")),
                SafeString(row("kufang")),
                SafeString(row("gongchang")),
                SafeString(row("kuanhao")),
                SafeString(row("pinlei")),
                SafeString(row("guige")),
                SafeString(row("caizhi")),
                SafeString(row("danzhong")),
                SafeString(row("xsshuliang")),
                SafeString(row("xsjinzhong")),
                SafeString(row("zongzhong")),
                SafeString(row("kejia")),
                SafeString(row("xiaoshoujia")),
                SafeString(row("yingshou")),
                SafeString(row("chengbengf")),
                SafeString(row("chengbenfj")),
                SafeString(row("chengben")),
                SafeString(row("xiaoshougf")),
                SafeString(row("xiaoshoufj")),
                SafeString(row("gflirun")),
                SafeString(row("zhekou")),
                SafeString(row("quankou")),
                SafeString(row("chengse")),
                SafeString(row("danwei")),
                SafeString(row("daogou")),
                SafeString(row("piling")),
                SafeString(row("pname")),
                SafeString(row("fxsfactory"))
            )
        Next
        AddDetailTotals(startDate, endDate)
    End Sub

    Private Sub AddDetailTotals(startDate As String, endDate As String)
        Dim sumSQL As String = "SELECT " &
            "CAST(ROUND(COALESCE(sum(tol.xsshuliang),0), 2) AS DECIMAL(30,2)) AS shuliang, " &
            "CAST(ROUND(COALESCE(sum(tol.xsjinzhong),0), 3) AS DECIMAL(30,3)) AS jinzhong, " &
            "CAST(ROUND(COALESCE(sum(tol.zongzhong),0), 3) AS DECIMAL(30,3)) AS zhongliang, " &
            "CAST(ROUND(COALESCE(sum(tol.xiaoshoujia),0), 2) AS DECIMAL(30,2)) AS xiaoshou, " &
            "CAST(ROUND(COALESCE(sum(tol.yingshou),0), 2) AS DECIMAL(30,2)) AS yingshou, " &
            "CAST(ROUND(COALESCE(sum(tol.chengben),0), 2) AS DECIMAL(30,2)) AS chengben " &
            $"FROM ({BuildSalesInnerSQL(startDate, endDate)}) AS tol {searchSpecFilter}"
        Dim sumDt As DataTable = ExecuteQuery(sumSQL, MySQL_ReadReport)
        If sumDt.Rows.Count > 0 Then
            Dim r As DataRow = sumDt.Rows(0)
            dgvReport.Rows.Add("合计", "", "", "", "", "", "", "", "", "", "", "",
                SafeString(r("shuliang")), SafeString(r("jinzhong")), SafeString(r("zhongliang")),
                "", SafeString(r("xiaoshou")), SafeString(r("yingshou")), "", "", SafeString(r("chengben")),
                "", "", "", "", "", "", "", "", "", "", "")
        End If
    End Sub

    ' ========== 构建销售内层SQL ==========
    Private Function BuildSalesInnerSQL(startDate As String, endDate As String) As String
        Dim orderFilter As String = ""
        If Not String.IsNullOrEmpty(localOrderCode) Then
            orderFilter = $" AND f.settlement_number='{SafeSQL(localOrderCode)}'"
        End If
        Dim searchFilter As String = ""
        If Not String.IsNullOrEmpty(searchName) Then
            searchFilter = $" AND (a.product_name LIKE '%{SafeSQL(searchName)}%' OR a.poduct_code LIKE '%{SafeSQL(searchName)}%' OR a.item_number LIKE '%{SafeSQL(searchName)}%')"
        End If
        Return "SELECT COALESCE(e1.id, e2.id, '0') AS guigeid, b.kufang AS kufangid, b.creationtime AS chukutime, " &
               "f.settlement_number AS djdingdanbian, f.customer_code as djkhbianma, k.`name` as djkhname, k.tel as djkhtel, " &
               "f.presale_number as djysdanhao, o.deposit as djdingjin, l.jin_zhong as hsjinzhong, l.settlement as hsjine, " &
               "f.salesman as djyewu, f.settlement as sdjhishou, f.taxamount as djshuijin, " &
               "CONCAT(a.cjuser, '(', j.name, ')') AS djccjuser, f.pling as djpiling, " &
               "b.poduct_code as bianma, a.product_name as mingcheng, " &
               "CASE WHEN b.kufang = '0' THEN '总库' ELSE h.title END AS kufang, g.title as gongchang, " &
               "a.item_number as kuanhao, CASE WHEN COALESCE(d.title, '') = '' THEN '未匹配' ELSE d.title END AS pinlei, " &
               "COALESCE(e1.title, e2.title, '未匹配') AS guige, a.caizhi as caizhi, " &
               "CAST(ROUND(COALESCE(a.single,0), 3) AS DECIMAL(30,3)) as danzhong, " &
               "CAST(ROUND(COALESCE(b.quantity,0), 2) AS DECIMAL(30,2)) as xsshuliang, " &
               "CAST(ROUND(COALESCE(b.net_weight,0), 3) AS DECIMAL(30,3)) as xsjinzhong, " &
               "CAST(ROUND(COALESCE(CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN b.net_weight ELSE a.weight / NULLIF(a.quantity,0) * b.quantity END,0), 3) AS DECIMAL(30,3)) AS zongzhong, " &
               "CAST(ROUND(COALESCE(b.gold_price,0), 2) AS DECIMAL(30,2)) as kejia, " &
               "CAST(ROUND(COALESCE(b.xiao_amount,0), 2) AS DECIMAL(30,2)) as xiaoshoujia, " &
               "CAST(ROUND(COALESCE(b.settlement,0), 2) AS DECIMAL(30,2)) as yingshou, " &
               "CAST(ROUND(COALESCE(c.basic_cost,0), 2) AS DECIMAL(30,2)) as chengbengf, " &
               "CAST(ROUND(COALESCE(c.company_surcharge,0), 2) AS DECIMAL(30,2)) as chengbenfj, " &
               "CAST(ROUND(COALESCE(c.cost_price,0), 2) AS DECIMAL(30,2)) as chengben, " &
               "CAST(ROUND(COALESCE(b.sales_cost,0), 2) AS DECIMAL(30,2)) as xiaoshougf, " &
               "CAST(ROUND(COALESCE(b.sales_surcharge,0), 2) AS DECIMAL(30,2)) as xiaoshoufj, " &
               "b.zhekou as zhekou, a.quandu as quankou, c.company_condition as chengse, " &
               "a.sales_unit as danwei, m.`name` as daogou, b.pling as xqpiling, " &
               "p.name AS pname, f.xsfactory AS fxsfactory, b.youhui AS youhui " &
               "FROM xipunum_erp_outbound AS b " &
               "INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code " &
               "INNER JOIN xipunum_erp_outbound_order AS f ON f.id = b.order_id " &
               "LEFT JOIN xipunum_erp_member AS k ON k.customer_code = f.customer_code " &
               "LEFT JOIN xipunum_erp_store AS c ON c.poduct_code = b.poduct_code " &
               "LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id " &
               "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = (SELECT specification_id FROM xipunum_erp_ksiamges WHERE kuanhao = a.item_number AND a.item_number != '' LIMIT 1) " &
               "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id " &
               "LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang " &
               "LEFT JOIN xipunum_erp_type AS g ON g.id = c.factory " &
               "LEFT JOIN xipunum_erp_user AS j ON j.user = a.cjuser " &
               "LEFT JOIN xipunum_erp_user AS m ON m.user = b.shopping_guide " &
               "LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = a.item_number " &
               "LEFT JOIN xipunum_erp_pay AS p ON p.id = b.pay_type " &
               $"WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}' " &
               $"{searchShopFilter} {searchCustomerFilter} {searchPlingFilter} {searchGoldWeightFilter} {orderFilter} {searchFilter}"
    End Function

    ' ========== 销售报表日期视图 ==========
    Private Sub QueryDateReport(startDate As String, endDate As String)
        Dim dateGroup As String = ""
        If radioDay.Checked Then
            dateGroup = "DATE(b.creationtime)"
        ElseIf radioMonth.Checked Then
            dateGroup = "DATE_FORMAT(b.creationtime, '%Y-%m')"
        ElseIf radioYear.Checked Then
            dateGroup = "DATE_FORMAT(b.creationtime, '%Y')"
        End If
        Dim dt As DataTable = GetSalesDateSummary(startDate, endDate, searchShopFilter, dateGroup)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            dgvReport.Rows.Add(
                Right("000" & (i + 1).ToString(), dt.Rows.Count.ToString().Length),
                SafeString(row("date")),
                SafeString(row("shuliang")),
                SafeString(row("jinzhong")),
                SafeString(row("zhongliang")),
                SafeString(row("cbgongfei")),
                SafeString(row("cbfujia")),
                SafeString(row("chengben")),
                SafeString(row("xsgongfei")),
                SafeString(row("xsfujia")),
                SafeString(row("xiaoshou")),
                SafeString(row("yingshou")),
                SafeString(row("dingjin")),
                SafeString(row("youhui")),
                SafeString(row("shishou")),
                SafeString(row("hsjinzhong")),
                SafeString(row("hsjine")),
                SafeString(row("shiji")),
                SafeString(row("gflirun"))
            )
        Next
    End Sub

    ' ========== 店铺汇总 ==========
    Private Sub QueryShopSummary(startDate As String, endDate As String)
        Dim dt As DataTable = GetSalesShopSummary(startDate, endDate, searchShopFilter)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            dgvReport.Rows.Add(
                Right("000" & (i + 1).ToString(), dt.Rows.Count.ToString().Length),
                SafeString(row("kufang")),
                SafeString(row("shuliang")),
                SafeString(row("jinzhong")),
                SafeString(row("zhongliang")),
                SafeString(row("cbgongfei")),
                SafeString(row("cbfujia")),
                SafeString(row("chengben")),
                SafeString(row("xsgongfei")),
                SafeString(row("xsfujia")),
                SafeString(row("xiaoshou")),
                SafeString(row("yingshou")),
                SafeString(row("dingjin")),
                SafeString(row("youhui")),
                SafeString(row("shishou")),
                SafeString(row("hsjinzhong")),
                SafeString(row("hsjine")),
                SafeString(row("shiji")),
                SafeString(row("gflirun"))
            )
        Next
    End Sub

    ' ========== 品类汇总 ==========
    Private Sub QueryCategorySummary(startDate As String, endDate As String)
        Dim dt As DataTable = GetSalesCategorySummary(startDate, endDate, searchShopFilter)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            dgvReport.Rows.Add(
                Right("000" & (i + 1).ToString(), dt.Rows.Count.ToString().Length),
                SafeString(row("pinlei")),
                SafeString(row("shuliang")),
                SafeString(row("jinzhong")),
                SafeString(row("zhongliang")),
                SafeString(row("cbgongfei")),
                SafeString(row("cbfujia")),
                SafeString(row("chengben")),
                SafeString(row("xsgongfei")),
                SafeString(row("xsfujia")),
                SafeString(row("xiaoshou")),
                SafeString(row("yingshou")),
                SafeString(row("dingjin")),
                SafeString(row("youhui")),
                SafeString(row("shishou")),
                SafeString(row("hsjinzhong")),
                SafeString(row("hsjine")),
                SafeString(row("shiji")),
                SafeString(row("gflirun"))
            )
        Next
    End Sub

    ' ========== 按钮事件 ==========
    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        LoadTableHeader()
        ExecuteMainQuery()
    End Sub

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
                If row.IsNewRow Then Continue For
                Dim dr As DataRow = dt.NewRow()
                For i As Integer = 0 To dgvReport.Columns.Count - 1
                    dr(i) = If(row.Cells(i).Value, "")
                Next
                dt.Rows.Add(dr)
            Next
            ExportToExcel(dt, "商品销售报表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        localOrderCode = ""
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        radioOrder.Checked = True
        radioDate.Checked = True
        radioAll.Checked = True
        txtSearch.Text = ""
        txtCustomerSearch.Text = ""
        txtGoldWeight1.Text = ""
        txtGoldWeight2.Text = ""
        chkGoldWeight.Checked = False
        chkCustomer.Items.Clear()
        For i As Integer = 0 To chkShop.Items.Count - 1
            chkShop.SetItemChecked(i, i = 0)
        Next
        For i As Integer = 0 To chkCategory.Items.Count - 1
            chkCategory.SetItemChecked(i, True)
        Next
        For i As Integer = 0 To chkSpec.Items.Count - 1
            chkSpec.SetItemChecked(i, True)
        Next
        LoadShopList()
        LoadCategoryList()
        LoadSpecList()
        LoadTableHeader()
        ExecuteMainQuery()
    End Sub

    ' ========== 单选框切换 ==========
    Private Sub RadioChanged(sender As Object, e As EventArgs)
        If Not isStateChanging Then
            localOrderCode = ""
            If radioMonth.Checked Then
                dtpStart.Format = DateTimePickerFormat.Custom
                dtpStart.CustomFormat = "yyyy-MM"
                dtpEnd.Format = DateTimePickerFormat.Custom
                dtpEnd.CustomFormat = "yyyy-MM"
            ElseIf radioYear.Checked Then
                dtpStart.Format = DateTimePickerFormat.Custom
                dtpStart.CustomFormat = "yyyy"
                dtpEnd.Format = DateTimePickerFormat.Custom
                dtpEnd.CustomFormat = "yyyy"
            Else
                dtpStart.Format = DateTimePickerFormat.Short
                dtpEnd.Format = DateTimePickerFormat.Short
            End If
            LoadTableHeader()
            ExecuteMainQuery()
        End If
    End Sub

    Private Sub RadioDimensionChanged(sender As Object, e As EventArgs)
        If Not isStateChanging Then
            localOrderCode = ""
            LoadTableHeader()
            ExecuteMainQuery()
        End If
    End Sub

    ' ========== 客户查找 ==========
    Private Sub txtCustomerSearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtCustomerSearch.Text) Then Return
            SearchCustomer(txtCustomerSearch.Text)
            txtCustomerSearch.Text = ""
        End If
    End Sub

    Private Sub SearchCustomer(searchText As String)
        Try
            Dim sql As String = $"SELECT customer_code, memberid, name FROM xipunum_erp_member WHERE memberid like '%{SafeSQL(searchText)}%' OR name like '%{SafeSQL(searchText)}%' OR tel like '%{SafeSQL(searchText)}%' ORDER BY id ASC"
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            chkCustomer.Items.Clear()
            If dt.Rows.Count = 0 Then
                ShowWarning("查询无此客户数据！")
                Return
            End If
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim name As String = SafeString(dt.Rows(i)("name"))
                Dim memberid As String = SafeString(dt.Rows(i)("memberid"))
                Dim code As String = SafeString(dt.Rows(i)("customer_code"))
                chkCustomer.Items.Add($"{name}({memberid})|{code}", False)
            Next
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 搜索回车 ==========
    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtSearch.Text) Then Return
            LoadTableHeader()
            ExecuteMainQuery()
            txtSearch.Text = ""
        End If
    End Sub

    ' ========== 右键菜单 ==========
    Private Sub dgvReport_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs)
        If e.Button = MouseButtons.Right AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If radioDate.Checked AndAlso radioOrder.Checked AndAlso Not String.IsNullOrEmpty(SafeString(dgvReport.Rows(e.RowIndex).Cells(1).Value)) Then
                dgvReport.Rows(e.RowIndex).Selected = True
                If String.IsNullOrEmpty(localOrderCode) Then
                    menuItemDetail.Visible = True
                    menuItemReturn.Visible = False
                Else
                    menuItemDetail.Visible = False
                    menuItemReturn.Visible = True
                End If
                contextMenuReport.Show(dgvReport, e.Location)
            End If
        End If
    End Sub

    Private Sub menuItemDetail_Click(sender As Object, e As EventArgs)
        If dgvReport.CurrentRow IsNot Nothing Then
            localOrderCode = SafeString(dgvReport.CurrentRow.Cells(1).Value)
            LoadTableHeader()
            ExecuteMainQuery()
        End If
    End Sub

    Private Sub menuItemReturn_Click(sender As Object, e As EventArgs)
        localOrderCode = ""
        LoadTableHeader()
        ExecuteMainQuery()
    End Sub

    Private Sub menuItemCopy_Click(sender As Object, e As EventArgs)
        If dgvReport.CurrentCell IsNot Nothing Then
            Dim cellValue As String = SafeString(dgvReport.CurrentCell.Value)
            Clipboard.SetText(cellValue)
            ShowInfo($"复制:{cellValue} 到剪切板成功！")
        End If
    End Sub
End Class

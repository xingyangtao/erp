' ============================================================================
' 商品入库报表窗口
' 功能: 入库数据统计，支持多维度筛选（订单/明细/天/月/年）
' 对应易语言: 窗口程序集_窗口_商品入库报表
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class InboundReportForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private localOrderCode As String = ""
    Private isStateChanging As Boolean = False

    ' ========== 查找条件变量 ==========
    Private searchStartDate As String = ""
    Private searchEndDate As String = ""
    Private searchShopFilter As String = ""
    Private searchCategoryFilter As String = ""
    Private searchSpecFilter As String = ""
    Private searchFactoryFilter As String = ""
    Private searchName As String = ""
    Private searchStateFilter As String = ""

    ' ========== 控件声明 ==========
    Private dgvReport As New DataGridView()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private chkShop As New CheckedListBox()
    Private chkCategory As New CheckedListBox()
    Private chkSpec As New CheckedListBox()
    Private chkFactory As New CheckedListBox()
    Private txtFactorySearch As New TextBox()
    Private txtSearch As New TextBox()
    Private radioOrder As New RadioButton()
    Private radioDetail As New RadioButton()
    Private radioDay As New RadioButton()
    Private radioMonth As New RadioButton()
    Private radioYear As New RadioButton()
    Private radioAll As New RadioButton()
    Private radioPending As New RadioButton()
    Private radioApproved As New RadioButton()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnReset As New Button()
    Private contextMenuReport As New ContextMenuStrip()
    Private menuItemDetail As New ToolStripMenuItem("订单明细")
    Private menuItemCopy As New ToolStripMenuItem("复制单元格")
    Private menuItemReturn As New ToolStripMenuItem("返回订单")

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "商品入库报表"
        Me.Size = New Drawing.Size(1400, 800)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' ========== 顶部筛选面板 ==========
        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 160
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

        ' 工厂多选
        AddLabel(panelTop, "工厂：", 900, 15)
        chkFactory.Location = New Drawing.Point(940, 12)
        chkFactory.Size = New Drawing.Size(120, 80)
        chkFactory.CheckOnClick = True
        panelTop.Controls.Add(chkFactory)

        ' 工厂查找
        AddLabel(panelTop, "工厂查找：", 390, 100)
        txtFactorySearch.Location = New Drawing.Point(460, 97)
        txtFactorySearch.Size = New Drawing.Size(120, 25)
        panelTop.Controls.Add(txtFactorySearch)
        AddHandler txtFactorySearch.KeyDown, AddressOf txtFactorySearch_KeyDown

        ' 搜索信息
        AddLabel(panelTop, "搜索：", 600, 100)
        txtSearch.Location = New Drawing.Point(640, 97)
        txtSearch.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtSearch)
        AddHandler txtSearch.KeyDown, AddressOf txtSearch_KeyDown

        ' 视图模式单选框
        AddLabel(panelTop, "视图：", 20, 50)
        AddRadio(panelTop, radioOrder, "订单", 60, 50, True)
        AddRadio(panelTop, radioDetail, "明细", 120, 50)
        AddRadio(panelTop, radioDay, "天", 180, 50)
        AddRadio(panelTop, radioMonth, "月", 220, 50)
        AddRadio(panelTop, radioYear, "年", 260, 50)

        ' 单据状态
        AddLabel(panelTop, "状态：", 800, 100)
        AddRadio(panelTop, radioAll, "所有", 840, 100, True)
        AddRadio(panelTop, radioPending, "待审", 900, 100)
        AddRadio(panelTop, radioApproved, "已审", 960, 100)

        ' 按钮
        AddButton(panelTop, btnQuery, "查询", 820, 50)
        AddButton(panelTop, btnExport, "导出", 920, 50)
        AddButton(panelTop, btnReset, "重置", 1020, 50)

        ' ========== 数据表格 ==========
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
            Dim sql As String = $"SELECT id AS akufang, CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id IN ({UserShopPermission}) UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({UserShopPermission}) ORDER BY akufang = '0' DESC, akufang"
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
            Dim sql As String = "SELECT id, title FROM xipunum_erp_category WHERE 1=1 UNION ALL SELECT 0 as id, '未匹配' as title ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id"
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
            Dim sql As String = "SELECT guige.id as id,guige.title as title FROM (SELECT id,category_id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title UNION ALL SELECT 0 as id,0 as category_id,'未匹配' as title ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id) as guige"
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            chkSpec.Items.Clear()
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim title As String = SafeString(dt.Rows(i)("title"))
                chkSpec.Items.Add(title, True)
            Next
            UpdateSpecIdList()
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 更新规格ID列表（根据选中品类筛选） ==========
    Private Sub UpdateSpecIdList()
        Try
            Dim categoryFilter As String = GetCheckedCategoryFilter()
            Dim specTitleFilter As String = GetCheckedSpecTitleFilter()

            Dim catClause As String = ""
            If categoryFilter <> "" Then
                catClause = $" and guige.category_id in ({categoryFilter})"
            End If

            Dim specClause As String = ""
            If specTitleFilter <> "" Then
                specClause = $" and guige.title in ({specTitleFilter})"
            End If

            Dim sql As String = $"SELECT guige.id as id,guige.title as title FROM (SELECT id,category_id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY id UNION ALL SELECT 0 as id,0 as category_id,'未匹配' as title ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id) as guige WHERE 1=1 {catClause} {specClause} ORDER BY guige.id asc"
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)

            ' 存储规格ID列表用于查询
            _specIdList.Clear()
            For Each row As DataRow In dt.Rows
                _specIdList.Add(SafeString(row("id")))
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private _specIdList As New List(Of String)()

    ' ========== 获取选中的店铺ID过滤条件 ==========
    Private Function GetCheckedShopFilter() As String
        Dim ids As New List(Of String)
        For i As Integer = 0 To chkShop.Items.Count - 1
            If chkShop.GetItemChecked(i) Then
                Dim item = chkShop.Items(i)
                Dim idProp = item.GetType().GetField("ID")
                If idProp IsNot Nothing Then
                    ids.Add("'" & idProp.GetValue(item).ToString() & "'")
                End If
            End If
        Next
        If ids.Count = 0 Then Return ""
        Return " and a.kufang in (" & String.Join(",", ids) & ")"
    End Function

    ' ========== 获取选中的品类ID过滤条件 ==========
    Private Function GetCheckedCategoryFilter() As String
        Dim ids As New List(Of String)
        For i As Integer = 0 To chkCategory.Items.Count - 1
            If chkCategory.GetItemChecked(i) Then
                Dim item = chkCategory.Items(i)
                Dim idProp = item.GetType().GetField("ID")
                If idProp IsNot Nothing Then
                    ids.Add("'" & idProp.GetValue(item).ToString() & "'")
                End If
            End If
        Next
        If ids.Count = 0 Then Return ""
        Return String.Join(",", ids)
    End Function

    ' ========== 获取选中的规格名称过滤条件 ==========
    Private Function GetCheckedSpecTitleFilter() As String
        Dim titles As New List(Of String)
        For i As Integer = 0 To chkSpec.Items.Count - 1
            If chkSpec.GetItemChecked(i) Then
                titles.Add("'" & chkSpec.Items(i).ToString() & "'")
            End If
        Next
        If titles.Count = 0 Then Return ""
        Return String.Join(",", titles)
    End Function

    ' ========== 获取选中的工厂ID过滤条件 ==========
    Private Function GetCheckedFactoryFilter() As String
        Dim ids As New List(Of String)
        For i As Integer = 0 To chkFactory.Items.Count - 1
            If chkFactory.GetItemChecked(i) Then
                Dim parts() As String = chkFactory.Items(i).ToString().Split("|"c)
                If parts.Length >= 2 Then
                    ids.Add("'" & parts(parts.Length - 1) & "'")
                End If
            End If
        Next
        If ids.Count = 0 Then Return ""
        Return " and f.factory in (" & String.Join(",", ids) & ")"
    End Function

    ' ========== 获取规格ID过滤条件（WHERE子句） ==========
    Private Function GetSpecIdFilter() As String
        If _specIdList.Count = 0 Then Return ""
        Dim ids As String = String.Join(",", _specIdList.Select(Function(id) $"'{id}'"))
        Return $" WHERE tol.guigeid in ({ids})"
    End Function

    ' ========== 加载表头 ==========
    Private Sub LoadTableHeader()
        dgvReport.Columns.Clear()
        dgvReport.DataSource = Nothing

        If radioOrder.Checked Then
            If String.IsNullOrEmpty(localOrderCode) Then
                ' 订单视图（无订单编码）
                AddColumns({"序号", "入库单号", "入库时间", "是否镶嵌", "送货单号", "品类", "半成品", "工厂", "来源", "结算", "入库店铺", "入库数量", "入库克价", "入库金重", "入库总重", "成本金额", "操作账户", "审核时间", "订单状态", "备注"},
                           {45, 140, 130, 100, 120, 60, 80, 140, 50, 50, 120, 75, 75, 75, 75, 75, 120, 140, 60, 200})
            Else
                ' 订单视图（有订单编码 - 显示详情列）
                AddColumns({"序号", "商品编码", "入库时间", "入库单号", "商品名称", "是否镶嵌", "款号", "工厂", "入库库房", "品类", "规格", "材质", "单件重", "数量", "金重", "重量", "成本工费", "成本附加费", "成本价", "参考工费", "圈口/长度", "面宽", "厚度", "工厂成色", "公司成色", "单位", "系数", "原料价"},
                           {45, 120, 130, 140, 100, 70, 100, 75, 65, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75})
            End If
        ElseIf radioDetail.Checked Then
            AddColumns({"序号", "商品编码", "入库时间", "入库单号", "商品名称", "是否镶嵌", "款号", "工厂", "入库库房", "品类", "规格", "材质", "单件重", "数量", "金重", "重量", "成本工费", "成本附加费", "成本价", "参考工费", "圈口/长度", "面宽", "厚度", "工厂成色", "公司成色", "单位", "系数", "原料价"},
                       {45, 120, 130, 140, 100, 70, 100, 75, 65, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75})
        ElseIf radioDay.Checked OrElse radioMonth.Checked OrElse radioYear.Checked Then
            AddColumns({"序号", "日期", "入库总数", "入库金重", "入库总重", "成本工费总额", "成本附加费总额", "成本总额", "参考工费总额"},
                       {45, 110, -1, -1, -1, -1, -1, -1, -1})
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

    ' ========== 构建内层SQL子查询 ==========
    Private Function BuildInnerSQL(startDate As String, endDate As String, shopFilter As String, categoryFilter As String, factoryFilter As String, stateFilter As String) As String
        Return "SELECT f.gold_price as kejia, b.creationtime as rukutime, f.shenhetime AS shenhetime, " &
               "f.odd_numbers AS dingdanbian, f.category_id AS djpinleiid, f.xiangqian as djxiangqian, " &
               "f.delivery as djdanhao, " &
               "CASE WHEN f.category_id = '0' THEN '未匹配' ELSE c.title END AS djpinlei, " &
               "f.half_product as djbanchengpin, " &
               "CASE WHEN f.factory = '' THEN '' ELSE g.title END AS djgongchang, " &
               "f.source as djlaiyuan, f.settlement as djjiesuan, " &
               "CONCAT(f.cjuser, '(', j.name, ')') AS ccjuser, " &
               "f.state as djstate, f.remarks as djbeizhu, " &
               "a.kufang as kufangid, " &
               "CASE WHEN a.kufang = '0' THEN '总库' ELSE h.title END AS kufang, " &
               "b.poduct_code AS bianma, " &
               "CASE WHEN a.product_name = '' or a.product_name IS NULL THEN i.title ELSE a.product_name END AS mingcheng, " &
               "A.xiangqian AS xqxiangqian, a.item_number as kuanhao, " &
               "CASE WHEN COALESCE(d.id, '') = '' THEN '0' ELSE d.id END AS pinleiid, " &
               "COALESCE(e1.id, e2.id, '0') AS guigeid, " &
               "CASE WHEN COALESCE(d.title, '') = '' THEN '未匹配' ELSE d.title END AS xqpinlei, " &
               "COALESCE(e1.title, e2.title, '未匹配') AS xqguige, " &
               "a.caizhi as caizhi, a.single as danjian, a.quantity as quantity, " &
               "a.jin_zhong as jinzhong, a.weight as zongzhong, " &
               "b.basic_cost as chengbengf, b.company_surcharge as chengbenfj, " &
               "CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN b.cost_price ELSE b.cost_price * a.quantity END AS chengben, " &
               "b.premium_cost as cankao, a.quandu as quankou, a.wide as miankuan, a.thickness as houdu, " &
               "b.factory_condition as gongchangcs, b.company_condition as gongsics, " &
               "a.sales_unit as danwei, b.coefficient as xishu " &
               "FROM xipunum_erp_store AS b " &
               "INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code " &
               "INNER JOIN xipunum_erp_store_order AS f ON f.id = b.order_id " &
               "LEFT JOIN xipunum_erp_category AS c ON c.id = f.category_id " &
               "LEFT JOIN xipunum_erp_type AS h ON h.id = a.kufang " &
               "INNER JOIN xipunum_erp_about AS g ON g.id = f.factory " &
               "LEFT JOIN xipunum_erp_user AS j ON j.user = f.cjuser " &
               "LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = a.item_number AND a.item_number != '' " &
               "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id " &
               "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id " &
               "LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id " &
               $"WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}' " &
               $"{shopFilter} {categoryFilter} {factoryFilter} {stateFilter}"
    End Function

    ' ========== 执行查询 ==========
    Private Sub ExecuteMainQuery()
        Try
            btnQuery.Enabled = False
            btnReset.Enabled = False
            btnExport.Enabled = False

            Dim startDate As String = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim endDate As String = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"

            ' 构建筛选条件
            searchShopFilter = GetCheckedShopFilter()
            Dim catFilter As String = GetCheckedCategoryFilter()
            searchCategoryFilter = If(catFilter <> "", $" and f.category_id in ({catFilter})", "")
            searchFactoryFilter = GetCheckedFactoryFilter()
            searchSpecFilter = GetSpecIdFilter()

            If radioAll.Checked Then
                searchStateFilter = ""
            ElseIf radioPending.Checked Then
                searchStateFilter = " and f.state='待审'"
            ElseIf radioApproved.Checked Then
                searchStateFilter = " and f.state='已审'"
            End If

            If txtSearch.Text = "" OrElse txtSearch.Text = "请输入查找信息" Then
                searchName = ""
            Else
                searchName = txtSearch.Text
            End If

            Me.Text = $"商品入库报表 截止时间:{dtpStart.Value:yyyy-MM-dd}"

            dgvReport.Rows.Clear()

            If radioOrder.Checked Then
                QueryOrderReport(startDate, endDate)
            ElseIf radioDetail.Checked Then
                QueryDetailReport(startDate, endDate)
            ElseIf radioDay.Checked OrElse radioMonth.Checked OrElse radioYear.Checked Then
                QueryDateReport(startDate, endDate)
            End If

        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        Finally
            btnQuery.Enabled = True
            btnReset.Enabled = True
            btnExport.Enabled = True
        End Try
    End Sub

    ' ========== 入库报表订单视图 ==========
    Private Sub QueryOrderReport(startDate As String, endDate As String)
        Dim innerSQL As String = BuildInnerSQL(startDate, endDate, searchShopFilter, searchCategoryFilter, searchFactoryFilter, searchStateFilter)

        If String.IsNullOrEmpty(localOrderCode) Then
            ' 订单汇总视图
            Dim sql As String = "SELECT tol.shenhetime AS shenhetime, tol.dingdanbian AS dingdanbian, tol.rukutime AS rukutime, " &
                "tol.djxiangqian AS djxiangqian, tol.djdanhao AS djdanhao, tol.djpinlei AS djpinlei, " &
                "tol.djbanchengpin AS djbanchengpin, tol.djgongchang AS djgongchang, tol.djlaiyuan AS djlaiyuan, " &
                "tol.djjiesuan AS djjiesuan, tol.kufang as kufang, " &
                "CAST(ROUND(COALESCE(sum(tol.quantity),0), 2) AS DECIMAL(30,2)) AS quantity, " &
                "CAST(ROUND(COALESCE(sum(tol.jinzhong),0), 3) AS DECIMAL(30,3)) AS jinzhong, " &
                "CAST(ROUND(COALESCE(sum(tol.zongzhong),0), 3) AS DECIMAL(30,3)) AS zongzhong, " &
                "CAST(ROUND(COALESCE(sum(tol.chengben),0), 2) AS DECIMAL(30,2)) AS chengben, " &
                "tol.ccjuser AS ccjuser, tol.djstate AS djstate, tol.djbeizhu AS djbeizhu, tol.kejia as kejia " &
                $"FROM ({innerSQL}) AS tol {searchSpecFilter} GROUP BY tol.dingdanbian ORDER BY tol.rukutime DESC"

            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim row As DataRow = dt.Rows(i)
                dgvReport.Rows.Add(
                    Right("000" & (i + 1).ToString(), dt.Rows.Count.ToString().Length),
                    SafeString(row("dingdanbian")),
                    SafeString(row("rukutime")),
                    SafeString(row("djxiangqian")),
                    SafeString(row("djdanhao")),
                    SafeString(row("djpinlei")),
                    SafeString(row("djbanchengpin")),
                    SafeString(row("djgongchang")),
                    SafeString(row("djlaiyuan")),
                    SafeString(row("djjiesuan")),
                    SafeString(row("kufang")),
                    SafeString(row("quantity")),
                    SafeString(row("kejia")),
                    SafeString(row("jinzhong")),
                    SafeString(row("zongzhong")),
                    SafeString(row("chengben")),
                    SafeString(row("ccjuser")),
                    SafeString(row("shenhetime")),
                    SafeString(row("djstate")),
                    SafeString(row("djbeizhu"))
                )
            Next

            ' 合计行
            Dim sumSQL As String = "SELECT " &
                "CAST(ROUND(COALESCE(sum(tol.quantity),0), 2) AS DECIMAL(30,2)) AS quantity, " &
                "CAST(ROUND(COALESCE(sum(tol.jinzhong),0), 3) AS DECIMAL(30,3)) AS jinzhong, " &
                "CAST(ROUND(COALESCE(sum(tol.zongzhong),0), 3) AS DECIMAL(30,3)) AS zongzhong, " &
                "CAST(ROUND(COALESCE(sum(tol.chengben),0), 2) AS DECIMAL(30,2)) AS chengben " &
                $"FROM ({innerSQL}) AS tol {searchSpecFilter}"
            Dim sumDt As DataTable = ExecuteQuery(sumSQL, MySQL_ReadReport)
            If sumDt.Rows.Count > 0 Then
                dgvReport.Rows.Add("合计", "", "", "", "", "", "", "", "", "", "",
                    SafeString(sumDt.Rows(0)("quantity")), "", SafeString(sumDt.Rows(0)("jinzhong")),
                    SafeString(sumDt.Rows(0)("zongzhong")), SafeString(sumDt.Rows(0)("chengben")), "", "", "", "")
            End If

        Else
            ' 订单详情视图（指定订单编码）
            Dim detailSQL As String = "SELECT f.gold_price as kejia, b.creationtime as rukutime, f.odd_numbers AS dingdanbian, " &
                "f.category_id AS djpinleiid, f.xiangqian as djxiangqian, f.delivery as djdanhao, " &
                "CASE WHEN f.category_id = '0' THEN '未匹配' ELSE c.title END AS djpinlei, " &
                "f.half_product as djbanchengpin, " &
                "CASE WHEN f.factory = '' THEN '' ELSE g.title END AS djgongchang, " &
                "f.source as djlaiyuan, f.settlement as djjiesuan, " &
                "CONCAT(f.cjuser, '(', j.name, ')') AS ccjuser, " &
                "f.state as djstate, f.remarks as djbeizhu, " &
                "a.kufang as kufangid, " &
                "CASE WHEN a.kufang = '0' THEN '总库' ELSE h.title END AS kufang, " &
                "b.poduct_code AS bianma, " &
                "CASE WHEN a.product_name = '' or a.product_name IS NULL THEN i.title ELSE a.product_name END AS mingcheng, " &
                "A.xiangqian AS xqxiangqian, a.item_number as kuanhao, " &
                "CASE WHEN COALESCE(d.id, '') = '' THEN '0' ELSE d.id END AS pinleiid, " &
                "COALESCE(e1.id, e2.id, '0') AS guigeid, " &
                "CASE WHEN COALESCE(d.title, '') = '' THEN '未匹配' ELSE d.title END AS xqpinlei, " &
                "COALESCE(e1.title, e2.title, '未匹配') AS xqguige, " &
                "a.caizhi as caizhi, a.single as danjian, a.quantity as quantity, " &
                "a.jin_zhong as jinzhong, a.weight as zongzhong, " &
                "b.basic_cost as chengbengf, b.company_surcharge as chengbenfj, " &
                "CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN b.cost_price ELSE b.cost_price * a.quantity END AS chengben, " &
                "b.premium_cost as cankao, a.quandu as quankou, a.wide as miankuan, a.thickness as houdu, " &
                "b.factory_condition as gongchangcs, b.company_condition as gongsics, " &
                "a.sales_unit as danwei, b.coefficient as xishu " &
                "FROM xipunum_erp_store AS b " &
                "INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code " &
                $"INNER JOIN xipunum_erp_store_order AS f ON f.id = b.order_id and f.odd_numbers='{SafeSQL(localOrderCode)}' " &
                "LEFT JOIN xipunum_erp_category AS c ON c.id = f.category_id " &
                "LEFT JOIN xipunum_erp_type AS h ON h.id = a.kufang " &
                "INNER JOIN xipunum_erp_about AS g ON g.id = f.factory " &
                "LEFT JOIN xipunum_erp_user AS j ON j.user = f.cjuser " &
                "LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = a.item_number AND a.item_number != '' " &
                "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id " &
                "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id " &
                "LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id"

            Dim dt As DataTable = ExecuteQuery(detailSQL, MySQL_ReadReport)
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim row As DataRow = dt.Rows(i)
                dgvReport.Rows.Add(
                    Right("000" & (i + 1).ToString(), dt.Rows.Count.ToString().Length),
                    SafeString(row("bianma")),
                    SafeString(row("rukutime")),
                    SafeString(row("dingdanbian")),
                    SafeString(row("mingcheng")),
                    SafeString(row("xqxiangqian")),
                    SafeString(row("kuanhao")),
                    SafeString(row("djgongchang")),
                    SafeString(row("kufang")),
                    SafeString(row("xqpinlei")),
                    SafeString(row("xqguige")),
                    SafeString(row("caizhi")),
                    SafeString(row("danjian")),
                    SafeString(row("quantity")),
                    SafeString(row("jinzhong")),
                    SafeString(row("zongzhong")),
                    SafeString(row("chengbengf")),
                    SafeString(row("chengbenfj")),
                    SafeString(row("chengben")),
                    SafeString(row("cankao")),
                    SafeString(row("quankou")),
                    SafeString(row("miankuan")),
                    SafeString(row("houdu")),
                    SafeString(row("gongchangcs")),
                    SafeString(row("gongsics")),
                    SafeString(row("danwei")),
                    SafeString(row("xishu")),
                    SafeString(row("kejia"))
                )
            Next

            ' 合计行
            Dim sumSQL As String = "SELECT " &
                "CAST(ROUND(COALESCE(sum(tol.quantity),0), 2) AS DECIMAL(30,2)) AS quantity, " &
                "CAST(ROUND(COALESCE(sum(tol.jinzhong),0), 3) AS DECIMAL(30,3)) AS jinzhong, " &
                "CAST(ROUND(COALESCE(sum(tol.zongzhong),0), 3) AS DECIMAL(30,3)) AS zongzhong, " &
                "CAST(ROUND(COALESCE(sum(tol.chengben),0), 2) AS DECIMAL(30,2)) AS chengben " &
                $"FROM ({detailSQL}) AS tol {searchSpecFilter}"
            Dim sumDt As DataTable = ExecuteQuery(sumSQL, MySQL_ReadReport)
            If sumDt.Rows.Count > 0 Then
                dgvReport.Rows.Add("合计", "", "", "", "", "", "", "", "", "", "", "",
                    SafeString(sumDt.Rows(0)("quantity")), SafeString(sumDt.Rows(0)("jinzhong")),
                    SafeString(sumDt.Rows(0)("zongzhong")), "", "", SafeString(sumDt.Rows(0)("chengben")), "", "", "", "", "", "", "", "", "")
            End If
        End If
    End Sub

    ' ========== 入库报表明细视图 ==========
    Private Sub QueryDetailReport(startDate As String, endDate As String)
        Dim innerSQL As String = BuildInnerSQL(startDate, endDate, searchShopFilter, searchCategoryFilter, searchFactoryFilter, searchStateFilter)

        Dim sql As String = "SELECT tol.bianma AS bianma, tol.rukutime AS rukutime, tol.dingdanbian AS dingdanbian, " &
            "tol.mingcheng AS mingcheng, tol.xqxiangqian AS xqxiangqian, tol.kuanhao AS kuanhao, " &
            "tol.djgongchang AS djgongchang, tol.xqpinlei AS xqpinlei, tol.xqguige AS xqguige, " &
            "tol.caizhi AS caizhi, tol.quankou AS quankou, tol.miankuan AS miankuan, tol.houdu AS houdu, " &
            "tol.gongchangcs AS gongchangcs, tol.gongsics AS gongsics, " &
            "CAST(ROUND(COALESCE(tol.danjian,0), 3) AS DECIMAL(30,3)) AS danjian, " &
            "CAST(ROUND(COALESCE(tol.quantity,0), 2) AS DECIMAL(30,2)) AS quantity, " &
            "CAST(ROUND(COALESCE(tol.jinzhong,0), 3) AS DECIMAL(30,3)) AS jinzhong, " &
            "CAST(ROUND(COALESCE(tol.zongzhong,0), 3) AS DECIMAL(30,3)) AS zongzhong, " &
            "tol.danwei AS danwei, " &
            "CAST(ROUND(COALESCE(tol.chengben,0), 2) AS DECIMAL(30,2)) AS chengben, " &
            "tol.xishu AS xishu, " &
            "CAST(ROUND(COALESCE(tol.chengbengf,0), 2) AS DECIMAL(30,2)) AS chengbengf, " &
            "CAST(ROUND(COALESCE(tol.cankao,0), 2) AS DECIMAL(30,2)) AS cankao, " &
            "CAST(ROUND(COALESCE(tol.chengbenfj,0), 2) AS DECIMAL(30,2)) AS chengbenfj, " &
            "tol.kufang AS kufang, tol.kejia as kejia " &
            $"FROM ({innerSQL}) AS tol {searchSpecFilter} ORDER BY tol.rukutime DESC"

        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            dgvReport.Rows.Add(
                Right("000" & (i + 1).ToString(), dt.Rows.Count.ToString().Length),
                SafeString(row("bianma")),
                SafeString(row("rukutime")),
                SafeString(row("dingdanbian")),
                SafeString(row("mingcheng")),
                SafeString(row("xqxiangqian")),
                SafeString(row("kuanhao")),
                SafeString(row("djgongchang")),
                SafeString(row("kufang")),
                SafeString(row("xqpinlei")),
                SafeString(row("xqguige")),
                SafeString(row("caizhi")),
                SafeString(row("danjian")),
                SafeString(row("quantity")),
                SafeString(row("jinzhong")),
                SafeString(row("zongzhong")),
                SafeString(row("chengbengf")),
                SafeString(row("chengbenfj")),
                SafeString(row("chengben")),
                SafeString(row("cankao")),
                SafeString(row("quankou")),
                SafeString(row("miankuan")),
                SafeString(row("houdu")),
                SafeString(row("gongchangcs")),
                SafeString(row("gongsics")),
                SafeString(row("danwei")),
                SafeString(row("xishu")),
                SafeString(row("kejia"))
            )
        Next

        ' 合计行
        Dim sumSQL As String = "SELECT " &
            "CAST(ROUND(COALESCE(sum(tol.quantity),0), 2) AS DECIMAL(30,2)) AS quantity, " &
            "CAST(ROUND(COALESCE(sum(tol.jinzhong),0), 3) AS DECIMAL(30,3)) AS jinzhong, " &
            "CAST(ROUND(COALESCE(sum(tol.zongzhong),0), 3) AS DECIMAL(30,3)) AS zongzhong, " &
            "CAST(ROUND(COALESCE(sum(tol.chengben),0), 2) AS DECIMAL(30,2)) AS chengben " &
            $"FROM ({innerSQL}) AS tol {searchSpecFilter}"
        Dim sumDt As DataTable = ExecuteQuery(sumSQL, MySQL_ReadReport)
        If sumDt.Rows.Count > 0 Then
            dgvReport.Rows.Add("合计", "", "", "", "", "", "", "", "", "", "", "",
                SafeString(sumDt.Rows(0)("quantity")), SafeString(sumDt.Rows(0)("jinzhong")),
                SafeString(sumDt.Rows(0)("zongzhong")), "", "", SafeString(sumDt.Rows(0)("chengben")), "", "", "", "", "", "", "", "", "")
        End If
    End Sub

    ' ========== 入库报表日期视图 ==========
    Private Sub QueryDateReport(startDate As String, endDate As String)
        Dim reportStart As String = dtpStart.Value.ToString("yyyy-MM-dd")
        Dim reportEnd As String = dtpEnd.Value.ToString("yyyy-MM-dd")

        Dim searchStart As String = ""
        Dim searchEnd As String = ""
        Dim loopCount As Integer = 0

        If radioDay.Checked Then
            searchStart = reportStart & " 00:00:00"
            searchEnd = DateTime.Parse(reportEnd).AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"
            loopCount = CInt((DateTime.Parse(reportEnd) - DateTime.Parse(reportStart)).TotalDays) + 1
        ElseIf radioMonth.Checked Then
            searchStart = reportStart & "-01 00:00:00"
            searchEnd = DateTime.Parse(reportEnd & "-01").AddMonths(1).ToString("yyyy-MM-dd") & " 00:00:00"
            loopCount = (DateTime.Parse(reportEnd & "-01").Year - DateTime.Parse(reportStart & "-01").Year) * 12 + DateTime.Parse(reportEnd & "-01").Month - DateTime.Parse(reportStart & "-01").Month + 1
        ElseIf radioYear.Checked Then
            searchStart = reportStart & "-01-01 00:00:00"
            searchEnd = DateTime.Parse(reportStart & "-01-01").AddYears(1).ToString("yyyy-MM-dd") & " 00:00:00"
            loopCount = DateTime.Parse(reportEnd & "-01-01").Year - DateTime.Parse(reportStart & "-01-01").Year + 1
        End If

        If loopCount <= 0 Then Return

        For i As Integer = 0 To loopCount - 1
            Dim loopDate As String = ""
            Dim loopStart As String = ""
            Dim loopEnd As String = ""

            If radioDay.Checked Then
                loopDate = DateTime.Parse(searchStart).AddDays(i).ToString("yyyy-MM-dd")
                loopStart = loopDate & " 00:00:00"
                loopEnd = DateTime.Parse(loopDate).AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"
            ElseIf radioMonth.Checked Then
                loopDate = DateTime.Parse(searchStart).AddMonths(i).ToString("yyyy-MM")
                loopStart = loopDate & "-01 00:00:00"
                loopEnd = DateTime.Parse(loopDate & "-01").AddMonths(1).ToString("yyyy-MM-dd") & " 00:00:00"
            ElseIf radioYear.Checked Then
                loopDate = DateTime.Parse(searchStart).AddYears(i).ToString("yyyy")
                loopStart = loopDate & "-01-01 00:00:00"
                loopEnd = DateTime.Parse(loopDate & "-01-01").AddYears(1).ToString("yyyy-MM-dd") & " 00:00:00"
            End If

            Dim innerSQL As String = BuildInnerSQL(loopStart, loopEnd, searchShopFilter, searchCategoryFilter, searchFactoryFilter, searchStateFilter)

            Dim sql As String = "SELECT " &
                "CAST(ROUND(COALESCE(sum(tol.quantity),0), 2) AS DECIMAL(30,2)) AS tol_rukunum, " &
                "CAST(ROUND(COALESCE(sum(tol.jinzhong),0), 3) AS DECIMAL(30,3)) AS tol_jinzhong, " &
                "CAST(ROUND(COALESCE(sum(tol.zongzhong),0), 3) AS DECIMAL(30,3)) AS tol_zhongliang, " &
                "CAST(ROUND(COALESCE(sum(tol.chengbengf*tol.jinzhong),0), 2) AS DECIMAL(30,2)) AS tol_chengbengf, " &
                "CAST(ROUND(COALESCE(sum(tol.chengbenfj),0), 2) AS DECIMAL(30,2)) AS tol_cbgongfei, " &
                "CAST(ROUND(COALESCE(sum(tol.chengben),0), 2) AS DECIMAL(30,2)) AS tol_chengben, " &
                "CAST(ROUND(COALESCE(sum(tol.cankao*tol.jinzhong),0), 2) AS DECIMAL(30,2)) AS tol_cankaogf " &
                $"FROM ({innerSQL}) AS tol {searchSpecFilter}"

            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                dgvReport.Rows.Add(
                    Right("000" & (i + 1).ToString(), loopCount.ToString().Length),
                    loopDate,
                    SafeString(row("tol_rukunum")),
                    SafeString(row("tol_jinzhong")),
                    SafeString(row("tol_zhongliang")),
                    SafeString(row("tol_chengbengf")),
                    SafeString(row("tol_cbgongfei")),
                    SafeString(row("tol_chengben")),
                    SafeString(row("tol_cankaogf"))
                )
            End If
        Next

        ' 合计行
        Dim totalInnerSQL As String = BuildInnerSQL(searchStart, searchEnd, searchShopFilter, searchCategoryFilter, searchFactoryFilter, searchStateFilter)
        Dim totalSQL As String = "SELECT " &
            "CAST(ROUND(COALESCE(sum(tol.quantity),0), 2) AS DECIMAL(30,2)) AS tol_rukunum, " &
            "CAST(ROUND(COALESCE(sum(tol.jinzhong),0), 3) AS DECIMAL(30,3)) AS tol_jinzhong, " &
            "CAST(ROUND(COALESCE(sum(tol.zongzhong),0), 3) AS DECIMAL(30,3)) AS tol_zhongliang, " &
            "CAST(ROUND(COALESCE(sum(tol.chengbengf*tol.jinzhong),0), 2) AS DECIMAL(30,2)) AS tol_chengbengf, " &
            "CAST(ROUND(COALESCE(sum(tol.chengbenfj),0), 2) AS DECIMAL(30,2)) AS tol_cbgongfei, " &
            "CAST(ROUND(COALESCE(sum(tol.chengben),0), 2) AS DECIMAL(30,2)) AS tol_chengben, " &
            "CAST(ROUND(COALESCE(sum(tol.cankao*tol.jinzhong),0), 2) AS DECIMAL(30,2)) AS tol_cankaogf " &
            $"FROM ({totalInnerSQL}) AS tol {searchSpecFilter}"
        Dim totalDt As DataTable = ExecuteQuery(totalSQL, MySQL_ReadReport)
        If totalDt.Rows.Count > 0 Then
            Dim row As DataRow = totalDt.Rows(0)
            dgvReport.Rows.Add("合计", "",
                SafeString(row("tol_rukunum")), SafeString(row("tol_jinzhong")), SafeString(row("tol_zhongliang")),
                SafeString(row("tol_chengbengf")), SafeString(row("tol_cbgongfei")), SafeString(row("tol_chengben")),
                SafeString(row("tol_cankaogf")))
        End If
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
            ' 构建DataTable用于导出
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
            ExportToExcel(dt, "商品入库报表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        localOrderCode = ""
        radioOrder.Checked = True
        radioAll.Checked = True
        txtSearch.Text = ""
        txtFactorySearch.Text = ""
        For i As Integer = 0 To chkShop.Items.Count - 1
            chkShop.SetItemChecked(i, i = 0)
        Next
        For i As Integer = 0 To chkCategory.Items.Count - 1
            chkCategory.SetItemChecked(i, True)
        Next
        For i As Integer = 0 To chkSpec.Items.Count - 1
            chkSpec.SetItemChecked(i, True)
        Next
        chkFactory.Items.Clear()
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

    ' ========== 工厂查找 ==========
    Private Sub txtFactorySearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtFactorySearch.Text) Then Return
            SearchFactory(txtFactorySearch.Text)
            txtFactorySearch.Text = ""
        End If
    End Sub

    Private Sub SearchFactory(searchText As String)
        Try
            Dim sql As String = $"SELECT id,title FROM xipunum_erp_about WHERE title like '%{SafeSQL(searchText)}%' OR jianxie like '%{SafeSQL(searchText)}%' OR name like '%{SafeSQL(searchText)}%' ORDER BY id asc"
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            chkFactory.Items.Clear()
            If dt.Rows.Count = 0 Then
                ShowWarning("查询无此信息数据！")
                Return
            End If
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim id As String = SafeString(dt.Rows(i)("id"))
                Dim title As String = SafeString(dt.Rows(i)("title"))
                chkFactory.Items.Add(title & "|" & id, False)
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
            If radioOrder.Checked AndAlso Not String.IsNullOrEmpty(SafeString(dgvReport.Rows(e.RowIndex).Cells(1).Value)) Then
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

' ============================================================================
' 商品退库报表窗口
' 功能: 退库数据统计，支持多维度筛选（订单/明细/天/月/年）
' 对应易语言: 窗口程序集_窗口_商品退库报表
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class ReturnReportForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private localOrderCode As String = ""
    Private isStateChanging As Boolean = False

    ' ========== 查找条件变量 ==========
    Private searchStartDate As String = ""
    Private searchEndDate As String = ""
    Private searchShopFilter As String = ""
    Private searchSpecFilter As String = ""
    Private searchFactoryFilter As String = ""
    Private searchName As String = ""

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
        Me.Text = "商品退库报表"
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

        ' 品类选中事件 - 联动规格
        AddHandler chkCategory.ItemCheck, AddressOf chkCategory_ItemCheck
        ' 规格选中事件 - 更新规格ID列表
        AddHandler chkSpec.ItemCheck, AddressOf chkSpec_ItemCheck
    End Sub

    Private Sub chkSpec_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        Me.BeginInvoke(New Action(Sub()
            UpdateSpecIdList()
        End Sub))
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
        UpdateSpecIdList()
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
            If chkShop.Items.Count <= 1 Then
                chkShop.Enabled = False
            End If
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

    ' ========== 品类选中联动 - 更新规格列表 ==========
    Private Sub chkCategory_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        ' 延迟执行以获取最新选中状态
        Me.BeginInvoke(New Action(Sub()
            UpdateSpecListByCategory()
            UpdateSpecIdList()
        End Sub))
    End Sub

    ' ========== 根据选中品类筛选规格列表 ==========
    Private Sub UpdateSpecListByCategory()
        Try
            Dim categoryFilter As String = GetCheckedCategoryFilter()
            Dim sql As String
            If String.IsNullOrEmpty(categoryFilter) Then
                sql = "SELECT guige.id as id,guige.title as title FROM (SELECT id,category_id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title UNION ALL SELECT 0 as id,0 as category_id,'未匹配' as title ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id) as guige ORDER BY guige.id asc"
            Else
                sql = $"SELECT guige.id as id,guige.title as title FROM (SELECT id,category_id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title UNION ALL SELECT 0 as id,0 as category_id,'未匹配' as title ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id) as guige WHERE guige.category_id in ({categoryFilter}) ORDER BY guige.id asc"
            End If
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            chkSpec.Items.Clear()
            If dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    Dim title As String = SafeString(dt.Rows(i)("title"))
                    chkSpec.Items.Add(title, False)
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 更新规格ID列表 ==========
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

            _specIdList.Clear()
            For Each row As DataRow In dt.Rows
                _specIdList.Add(SafeString(row("id")))
            Next
        Catch ex As Exception
        End Try
    End Sub

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
        If ids.Count = 0 Then
            ' 未选中时使用全局店铺权限
            Return $" and a.ykufang in ({UserShopPermission})"
        End If
        Return " and a.ykufang in (" & String.Join(",", ids) & ")"
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
                AddColumns({"序号", "退库单号", "退库时间", "数量", "金重", "总重", "成本工费", "成本附加费", "成本价", "状态", "库房", "备注", "操作账户"},
                           {45, 120, 140, 85, 85, 85, 85, 85, 85, 85, 85, -1, 140})
            Else
                ' 订单视图（有订单编码 - 显示详情列）
                AddColumns({"序号", "商品编码", "退库时间", "退库单号", "品类", "规格", "商品名称", "款号", "原库房", "新库房", "单件重", "退库数量", "退库金重", "退库重量", "成本工费", "成本附加费", "成本价", "备注"},
                           {45, 140, 130, 130, 100, 120, 60, 140, 120, 75, 75, 75, 75, 75, 75, 75, 75, 300})
            End If
        ElseIf radioDetail.Checked Then
            AddColumns({"序号", "商品编码", "退库时间", "退库单号", "品类", "规格", "商品名称", "款号", "原库房", "新库房", "单件重", "退库数量", "退库金重", "退库重量", "成本工费", "成本附加费", "成本价", "备注"},
                       {45, 140, 130, 130, 100, 120, 60, 140, 120, 75, 75, 75, 75, 75, 75, 75, 75, 300})
        ElseIf radioDay.Checked OrElse radioMonth.Checked OrElse radioYear.Checked Then
            AddColumns({"序号", "日期", "退库总数", "退库金重", "退库总重", "成本工费", "成本附加费", "成本价"},
                       {45, 100, -1, -1, -1, -1, -1, -1})
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

    ' ========== 构建内层SQL子查询（订单/明细模式） ==========
    Private Function BuildInnerSQL(startDate As String, endDate As String, shopFilter As String, factoryFilter As String, searchName As String) As String
        Dim searchClause As String = ""
        If String.IsNullOrEmpty(searchName) Then
            searchClause = "WHERE (b.poduct_code LIKE '%%' OR b.fu_code LIKE '%%' OR b.product_name LIKE '%%' OR b.item_number LIKE '%%' OR b.caizhi LIKE '%%' OR g.title LIKE '%%' OR g.jianxie LIKE '%%')"
        Else
            Dim safeName As String = SafeSQL(searchName)
            searchClause = $"WHERE (b.poduct_code LIKE '%{safeName}%' OR b.fu_code LIKE '%{safeName}%' OR b.product_name LIKE '%{safeName}%' OR b.item_number LIKE '%{safeName}%' OR b.caizhi LIKE '%{safeName}%' OR g.title LIKE '%{safeName}%' OR g.jianxie LIKE '%{safeName}%')"
        End If

        Return "SELECT a.creationtime AS tuikutime, i.tuiku_umber AS tuikudh, i.state as zhuangtai, i.remarks as beizhu, " &
               "CONCAT(a.cjuser, '(', j.NAME, ')') AS ccjuser, " &
               "COALESCE(e1.id, e2.id, '0') AS guigeid, " &
               "a.poduct_code AS bianma, " &
               "CASE WHEN COALESCE(m.title, '') = '' THEN '未匹配' ELSE m.title END AS pinlei, " &
               "COALESCE(e1.title, e2.title, '未匹配') AS guige, " &
               "b.product_name as mingcheng, b.item_number as kuanhoa, " &
               "CASE WHEN a.ykufang = '0' THEN '总库' ELSE h.title END AS tkkufang, " &
               "CASE WHEN a.xkufang = '0' THEN '总库' ELSE k.title END AS jskufang, " &
               "b.single as danjian, a.quantity as tuikushu, a.jinzhong as tuikujin, " &
               "CASE WHEN COALESCE(c.lingxiao, '') = '是' THEN a.jinzhong ELSE b.single*a.quantity END AS zongzhong, " &
               "e.basic_cost as chengbengf, e.company_surcharge as chengbenfj, " &
               "CASE WHEN COALESCE(c.lingxiao, '') = '是' THEN e.cost_price / b.jin_zhong * a.jinzhong ELSE e.cost_price * a.quantity END AS chengben, " &
               "a.remarks as xqbeizhu " &
               "FROM xipunum_erp_tuiku AS a " &
               "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
               "INNER JOIN xipunum_erp_store AS e ON e.poduct_code = b.poduct_code " &
               "INNER JOIN xipunum_erp_store_order AS f ON f.id = e.order_id " &
               "INNER JOIN xipunum_erp_about AS g ON g.id = f.factory " &
               "LEFT JOIN xipunum_erp_type AS h ON h.id = a.ykufang " &
               "INNER JOIN xipunum_erp_tuiku_order AS i ON i.id = a.order_id " &
               "INNER JOIN xipunum_erp_user AS j ON j.USER = i.cjuser " &
               "LEFT JOIN xipunum_erp_type AS k ON k.id = a.xkufang " &
               "LEFT JOIN xipunum_erp_ksiamges AS c ON c.kuanhao = b.item_number AND b.item_number != '' " &
               "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = c.specification_id AND b.item_number IS NOT NULL AND b.item_number != '' " &
               "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id AND b.specification_id IS NOT NULL AND b.specification_id != '' " &
               "LEFT JOIN xipunum_erp_category AS m ON m.id = COALESCE(e1.category_id, e2.category_id) AND COALESCE(e1.category_id, e2.category_id) IS NOT NULL " &
               $"{searchClause} {shopFilter} {factoryFilter} " &
               $"AND a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' " &
               "ORDER BY a.creationtime DESC"
    End Function

    ' ========== 构建内层SQL子查询（日期模式 - chengbengf计算不同） ==========
    Private Function BuildDateInnerSQL(startDate As String, endDate As String, shopFilter As String, factoryFilter As String, searchName As String) As String
        Dim searchClause As String = ""
        If String.IsNullOrEmpty(searchName) Then
            searchClause = "WHERE (b.poduct_code LIKE '%%' OR b.fu_code LIKE '%%' OR b.product_name LIKE '%%' OR b.item_number LIKE '%%' OR b.caizhi LIKE '%%' OR g.title LIKE '%%' OR g.jianxie LIKE '%%')"
        Else
            Dim safeName As String = SafeSQL(searchName)
            searchClause = $"WHERE (b.poduct_code LIKE '%{safeName}%' OR b.fu_code LIKE '%{safeName}%' OR b.product_name LIKE '%{safeName}%' OR b.item_number LIKE '%{safeName}%' OR b.caizhi LIKE '%{safeName}%' OR g.title LIKE '%{safeName}%' OR g.jianxie LIKE '%{safeName}%')"
        End If

        Return "SELECT a.creationtime AS tuikutime, i.tuiku_umber AS tuikudh, i.state as zhuangtai, i.remarks as beizhu, " &
               "CONCAT(a.cjuser, '(', j.NAME, ')') AS ccjuser, " &
               "COALESCE(e1.id, e2.id, '0') AS guigeid, " &
               "a.poduct_code AS bianma, " &
               "CASE WHEN COALESCE(m.title, '') = '' THEN '未匹配' ELSE m.title END AS pinlei, " &
               "COALESCE(e1.title, e2.title, '未匹配') AS guige, " &
               "b.product_name as mingcheng, b.item_number as kuanhoa, " &
               "CASE WHEN a.ykufang = '0' THEN '总库' ELSE h.title END AS tkkufang, " &
               "CASE WHEN a.xkufang = '0' THEN '总库' ELSE k.title END AS jskufang, " &
               "b.single as danjian, a.quantity as tuikushu, a.jinzhong as tuikujin, " &
               "CASE WHEN COALESCE(c.lingxiao, '') = '是' THEN a.jinzhong ELSE b.single*a.quantity END AS zongzhong, " &
               "e.basic_cost*a.jinzhong as chengbengf, e.company_surcharge as chengbenfj, " &
               "CASE WHEN COALESCE(c.lingxiao, '') = '是' THEN e.cost_price / b.jin_zhong * a.jinzhong ELSE e.cost_price * a.quantity END AS chengben, " &
               "a.remarks as xqbeizhu " &
               "FROM xipunum_erp_tuiku AS a " &
               "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
               "INNER JOIN xipunum_erp_store AS e ON e.poduct_code = b.poduct_code " &
               "INNER JOIN xipunum_erp_store_order AS f ON f.id = e.order_id " &
               "INNER JOIN xipunum_erp_about AS g ON g.id = f.factory " &
               "LEFT JOIN xipunum_erp_type AS h ON h.id = a.ykufang " &
               "INNER JOIN xipunum_erp_tuiku_order AS i ON i.id = a.order_id " &
               "INNER JOIN xipunum_erp_user AS j ON j.USER = i.cjuser " &
               "LEFT JOIN xipunum_erp_type AS k ON k.id = a.xkufang " &
               "LEFT JOIN xipunum_erp_ksiamges AS c ON c.kuanhao = b.item_number AND b.item_number != '' " &
               "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = c.specification_id AND b.item_number IS NOT NULL AND b.item_number != '' " &
               "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id AND b.specification_id IS NOT NULL AND b.specification_id != '' " &
               "LEFT JOIN xipunum_erp_category AS m ON m.id = COALESCE(e1.category_id, e2.category_id) AND COALESCE(e1.category_id, e2.category_id) IS NOT NULL " &
               $"{searchClause} {shopFilter} {factoryFilter} " &
               $"AND a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' " &
               "ORDER BY a.creationtime DESC"
    End Function

    ' ========== 构建订单详情内层SQL（指定订单编码） ==========
    Private Function BuildOrderDetailInnerSQL(orderCode As String) As String
        Return "SELECT a.creationtime AS tuikutime, i.tuiku_umber AS tuikudh, i.state as zhuangtai, i.remarks as beizhu, " &
               "CONCAT(a.cjuser, '(', j.NAME, ')') AS ccjuser, " &
               "COALESCE(e1.id, e2.id, '0') AS guigeid, " &
               "a.poduct_code AS bianma, " &
               "CASE WHEN COALESCE(m.title, '') = '' THEN '未匹配' ELSE m.title END AS pinlei, " &
               "COALESCE(e1.title, e2.title, '未匹配') AS guige, " &
               "b.product_name as mingcheng, b.item_number as kuanhoa, " &
               "CASE WHEN a.ykufang = '0' THEN '总库' ELSE h.title END AS tkkufang, " &
               "CASE WHEN a.xkufang = '0' THEN '总库' ELSE k.title END AS jskufang, " &
               "b.single as danjian, a.quantity as tuikushu, a.jinzhong as tuikujin, " &
               "CASE WHEN COALESCE(c.lingxiao, '') = '是' THEN a.jinzhong ELSE b.single*a.quantity END AS zongzhong, " &
               "e.basic_cost as chengbengf, e.company_surcharge as chengbenfj, " &
               "CASE WHEN COALESCE(c.lingxiao, '') = '是' THEN e.cost_price / b.jin_zhong * a.jinzhong ELSE e.cost_price * a.quantity END AS chengben, " &
               "a.remarks as xqbeizhu " &
               "FROM xipunum_erp_tuiku AS a " &
               "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
               "INNER JOIN xipunum_erp_store AS e ON e.poduct_code = b.poduct_code " &
               "INNER JOIN xipunum_erp_store_order AS f ON f.id = e.order_id " &
               "INNER JOIN xipunum_erp_about AS g ON g.id = f.factory " &
               "LEFT JOIN xipunum_erp_type AS h ON h.id = a.ykufang " &
               "INNER JOIN xipunum_erp_tuiku_order AS i ON i.id = a.order_id " &
               "INNER JOIN xipunum_erp_user AS j ON j.USER = i.cjuser " &
               "LEFT JOIN xipunum_erp_type AS k ON k.id = a.xkufang " &
               "LEFT JOIN xipunum_erp_ksiamges AS c ON c.kuanhao = b.item_number AND b.item_number != '' " &
               "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = c.specification_id AND b.item_number IS NOT NULL AND b.item_number != '' " &
               "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id AND b.specification_id IS NOT NULL AND b.specification_id != '' " &
               "LEFT JOIN xipunum_erp_category AS m ON m.id = COALESCE(e1.category_id, e2.category_id) AND COALESCE(e1.category_id, e2.category_id) IS NOT NULL " &
               $"WHERE i.tuiku_umber ='{SafeSQL(orderCode)}' " &
               "ORDER BY a.creationtime DESC"
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
            searchFactoryFilter = GetCheckedFactoryFilter()
            searchSpecFilter = GetSpecIdFilter()

            If txtSearch.Text = "" OrElse txtSearch.Text = "请输入查找信息" Then
                searchName = ""
            Else
                searchName = txtSearch.Text
            End If

            Me.Text = $"商品退库报表 截止时间:{dtpStart.Value:yyyy-MM-dd}"

            dgvReport.Rows.Clear()

            If radioOrder.Checked Then
                QueryOrderReport(startDate, endDate)
            ElseIf radioDetail.Checked Then
                QueryDetailReport(startDate, endDate)
            ElseIf radioDay.Checked OrElse radioMonth.Checked OrElse radioYear.Checked Then
                QueryDateReport()
            End If

        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        Finally
            btnQuery.Enabled = True
            btnReset.Enabled = True
            btnExport.Enabled = True
        End Try
    End Sub

    ' ========== 退库报表订单视图 ==========
    Private Sub QueryOrderReport(startDate As String, endDate As String)
        If String.IsNullOrEmpty(localOrderCode) Then
            ' 订单汇总视图
            Dim innerSQL As String = BuildInnerSQL(startDate, endDate, searchShopFilter, searchFactoryFilter, searchName)

            Dim sql As String = "SELECT tol.tuikudh AS tuiku_umber, tol.tuikutime AS tuikutime, " &
                "CAST(ROUND(COALESCE(sum(tol.tuikushu), 0),3) AS DECIMAL(30,3)) AS tuikushu, " &
                "CAST(ROUND(COALESCE(sum(tol.tuikujin), 0),3) AS DECIMAL(30,3)) AS tuikujin, " &
                "CAST(ROUND(COALESCE(sum(tol.zongzhong), 0),3) AS DECIMAL(30,3)) AS zongzhong, " &
                "CAST(ROUND(COALESCE(sum(tol.chengbengf*tol.tuikujin), 0),3) AS DECIMAL(30,3)) AS chengbengf, " &
                "CAST(ROUND(COALESCE(sum(tol.chengbenfj), 0),3) AS DECIMAL(30,3)) AS chengbenfj, " &
                "CAST(ROUND(COALESCE(sum(tol.chengben), 0),3) AS DECIMAL(30,3)) AS chengben, " &
                "tol.zhuangtai AS zhuangtai, tol.tkkufang AS tkkufang, tol.beizhu AS beizhu, tol.ccjuser AS ccjuser " &
                $"FROM ({innerSQL}) AS tol GROUP BY tol.tuikutime, tol.tuikudh ORDER BY tol.tuikutime DESC"

            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim row As DataRow = dt.Rows(i)
                dgvReport.Rows.Add(
                    Right("000" & (i + 1).ToString(), dt.Rows.Count.ToString().Length),
                    SafeString(row("tuiku_umber")),
                    SafeString(row("tuikutime")),
                    SafeString(row("tuikushu")),
                    SafeString(row("tuikujin")),
                    SafeString(row("zongzhong")),
                    SafeString(row("chengbengf")),
                    SafeString(row("chengbenfj")),
                    SafeString(row("chengben")),
                    SafeString(row("zhuangtai")),
                    SafeString(row("tkkufang")),
                    SafeString(row("beizhu")),
                    SafeString(row("ccjuser"))
                )
            Next

            ' 合计行
            Dim sumSQL As String = "SELECT " &
                "CAST(ROUND(COALESCE(sum(tol.tuikushu), 0),3) AS DECIMAL(30,3)) AS tuikushu, " &
                "CAST(ROUND(COALESCE(sum(tol.tuikujin), 0),3) AS DECIMAL(30,3)) AS tuikujin, " &
                "CAST(ROUND(COALESCE(sum(tol.zongzhong), 0),3) AS DECIMAL(30,3)) AS zongzhong, " &
                "CAST(ROUND(COALESCE(sum(tol.chengbengf*tol.tuikujin), 0),3) AS DECIMAL(30,3)) AS chengbengf, " &
                "CAST(ROUND(COALESCE(sum(tol.chengbenfj), 0),3) AS DECIMAL(30,3)) AS chengbenfj, " &
                "CAST(ROUND(COALESCE(sum(tol.chengben), 0),3) AS DECIMAL(30,3)) AS chengben " &
                $"FROM ({innerSQL}) AS tol"
            Dim sumDt As DataTable = ExecuteQuery(sumSQL, MySQL_ReadReport)
            If sumDt.Rows.Count > 0 Then
                dgvReport.Rows.Add("合计", "", "",
                    SafeString(sumDt.Rows(0)("tuikushu")),
                    SafeString(sumDt.Rows(0)("tuikujin")),
                    SafeString(sumDt.Rows(0)("zongzhong")),
                    SafeString(sumDt.Rows(0)("chengbengf")),
                    SafeString(sumDt.Rows(0)("chengbenfj")),
                    SafeString(sumDt.Rows(0)("chengben")),
                    "", "", "", "")
            End If

        Else
            ' 订单详情视图（指定订单编码）
            Dim innerSQL As String = BuildOrderDetailInnerSQL(localOrderCode)

            Dim sql As String = "SELECT tol.bianma as bianma, tol.tuikutime as tuikutime, tol.tuikudh as tuikudh, " &
                "tol.pinlei as pinlei, tol.guige as guige, tol.mingcheng as mingcheng, tol.kuanhoa as kuanhoa, " &
                "tol.tkkufang as tkkufang, tol.jskufang as jskufang, " &
                "CAST(ROUND(COALESCE(tol.danjian, 0),3) AS DECIMAL(30,3)) as danjian, " &
                "CAST(ROUND(COALESCE(tol.tuikushu, 0),3) AS DECIMAL(30,3)) as tuikushu, " &
                "CAST(ROUND(COALESCE(tol.tuikujin, 0),3) AS DECIMAL(30,3)) as tuikujin, " &
                "CAST(ROUND(COALESCE(tol.zongzhong, 0),3) AS DECIMAL(30,3)) as zongzhong, " &
                "CAST(ROUND(COALESCE(tol.chengbengf, 0),3) AS DECIMAL(30,3)) as chengbengf, " &
                "CAST(ROUND(COALESCE(tol.chengbenfj, 0),3) AS DECIMAL(30,3)) as chengbenfj, " &
                "CAST(ROUND(COALESCE(tol.chengben, 0),3) AS DECIMAL(30,3)) as chengben, " &
                "tol.xqbeizhu as xqbeizhu " &
                $"FROM ({innerSQL}) AS tol GROUP BY tol.tuikutime, tol.tuikudh, tol.guigeid, tol.bianma ORDER BY tol.tuikutime DESC"

            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim row As DataRow = dt.Rows(i)
                dgvReport.Rows.Add(
                    Right("000" & (i + 1).ToString(), dt.Rows.Count.ToString().Length),
                    SafeString(row("bianma")),
                    SafeString(row("tuikutime")),
                    SafeString(row("tuikudh")),
                    SafeString(row("pinlei")),
                    SafeString(row("guige")),
                    SafeString(row("mingcheng")),
                    SafeString(row("kuanhoa")),
                    SafeString(row("jskufang")),
                    SafeString(row("guige")),
                    SafeString(row("danjian")),
                    SafeString(row("tuikushu")),
                    SafeString(row("tuikujin")),
                    SafeString(row("zongzhong")),
                    SafeString(row("chengbengf")),
                    SafeString(row("chengbenfj")),
                    SafeString(row("chengben")),
                    SafeString(row("xqbeizhu"))
                )
            Next

            ' 合计行
            Dim sumSQL As String = "SELECT " &
                "CAST(ROUND(COALESCE(sum(tol.tuikushu), 0),3) AS DECIMAL(30,3)) as tuikushu, " &
                "CAST(ROUND(COALESCE(sum(tol.tuikujin), 0),3) AS DECIMAL(30,3)) as tuikujin, " &
                "CAST(ROUND(COALESCE(sum(tol.zongzhong), 0),3) AS DECIMAL(30,3)) as zongzhong " &
                $"FROM ({innerSQL}) AS tol"
            Dim sumDt As DataTable = ExecuteQuery(sumSQL, MySQL_ReadReport)
            If sumDt.Rows.Count > 0 Then
                dgvReport.Rows.Add("合计", "", "", "", "", "", "", "", "", "",
                    SafeString(sumDt.Rows(0)("tuikushu")),
                    SafeString(sumDt.Rows(0)("tuikujin")),
                    SafeString(sumDt.Rows(0)("zongzhong")),
                    "", "", "", "")
            End If
        End If
    End Sub

    ' ========== 退库报表明细视图 ==========
    Private Sub QueryDetailReport(startDate As String, endDate As String)
        Dim innerSQL As String = BuildInnerSQL(startDate, endDate, searchShopFilter, searchFactoryFilter, searchName)

        Dim sql As String = "SELECT tol.bianma as bianma, tol.tuikutime as tuikutime, tol.tuikudh as tuikudh, " &
            "tol.pinlei as pinlei, tol.guige as guige, tol.mingcheng as mingcheng, tol.kuanhoa as kuanhoa, " &
            "tol.tkkufang as tkkufang, tol.jskufang as jskufang, " &
            "CAST(ROUND(COALESCE(tol.danjian, 0),3) AS DECIMAL(30,3)) as danjian, " &
            "CAST(ROUND(COALESCE(tol.tuikushu, 0),3) AS DECIMAL(30,3)) as tuikushu, " &
            "CAST(ROUND(COALESCE(tol.tuikujin, 0),3) AS DECIMAL(30,3)) as tuikujin, " &
            "CAST(ROUND(COALESCE(tol.zongzhong, 0),3) AS DECIMAL(30,3)) as zongzhong, " &
            "CAST(ROUND(COALESCE(tol.chengbengf, 0),3) AS DECIMAL(30,3)) as chengbengf, " &
            "CAST(ROUND(COALESCE(tol.chengbenfj, 0),3) AS DECIMAL(30,3)) as chengbenfj, " &
            "CAST(ROUND(COALESCE(tol.chengben, 0),3) AS DECIMAL(30,3)) as chengben, " &
            "tol.xqbeizhu as xqbeizhu " &
            $"FROM ({innerSQL}) AS tol {searchSpecFilter} GROUP BY tol.tuikutime, tol.tuikudh, tol.guigeid, tol.bianma ORDER BY tol.tuikutime DESC"

        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            dgvReport.Rows.Add(
                Right("000" & (i + 1).ToString(), dt.Rows.Count.ToString().Length),
                SafeString(row("bianma")),
                SafeString(row("tuikutime")),
                SafeString(row("tuikudh")),
                SafeString(row("pinlei")),
                SafeString(row("guige")),
                SafeString(row("mingcheng")),
                SafeString(row("kuanhoa")),
                SafeString(row("jskufang")),
                SafeString(row("guige")),
                SafeString(row("danjian")),
                SafeString(row("tuikushu")),
                SafeString(row("tuikujin")),
                SafeString(row("zongzhong")),
                SafeString(row("chengbengf")),
                SafeString(row("chengbenfj")),
                SafeString(row("chengben")),
                SafeString(row("xqbeizhu"))
            )
        Next

        ' 合计行
        Dim sumSQL As String = "SELECT " &
            "CAST(ROUND(COALESCE(sum(tol.tuikushu), 0),3) AS DECIMAL(30,3)) as tuikushu, " &
            "CAST(ROUND(COALESCE(sum(tol.tuikujin), 0),3) AS DECIMAL(30,3)) as tuikujin, " &
            "CAST(ROUND(COALESCE(sum(tol.zongzhong), 0),3) AS DECIMAL(30,3)) as zongzhong " &
            $"FROM ({innerSQL}) AS tol {searchSpecFilter}"
        Dim sumDt As DataTable = ExecuteQuery(sumSQL, MySQL_ReadReport)
        If sumDt.Rows.Count > 0 Then
            dgvReport.Rows.Add("合计", "", "", "", "", "", "", "", "", "",
                SafeString(sumDt.Rows(0)("tuikushu")),
                SafeString(sumDt.Rows(0)("tuikujin")),
                SafeString(sumDt.Rows(0)("zongzhong")),
                "", "", "", "")
        End If
    End Sub

    ' ========== 退库报表日期视图 ==========
    Private Sub QueryDateReport()
        Dim reportStartDate As String = ""
        Dim reportEndDate As String = ""
        Dim loopCount As Integer = 0
        Dim baseDate As DateTime

        If radioDay.Checked Then
            reportStartDate = dtpStart.Value.ToString("yyyy-MM-dd")
            reportEndDate = dtpEnd.Value.ToString("yyyy-MM-dd")
            baseDate = DateTime.Parse(reportStartDate)
            loopCount = CInt((DateTime.Parse(reportEndDate) - DateTime.Parse(reportStartDate)).TotalDays) + 1
        ElseIf radioMonth.Checked Then
            reportStartDate = dtpStart.Value.ToString("yyyy-MM") & "-01"
            reportEndDate = dtpEnd.Value.ToString("yyyy-MM") & "-01"
            baseDate = DateTime.Parse(reportStartDate)
            loopCount = (DateTime.Parse(reportEndDate).Year - DateTime.Parse(reportStartDate).Year) * 12 + 
                        DateTime.Parse(reportEndDate).Month - DateTime.Parse(reportStartDate).Month + 1
        ElseIf radioYear.Checked Then
            reportStartDate = dtpStart.Value.ToString("yyyy") & "-01-01"
            reportEndDate = dtpEnd.Value.ToString("yyyy") & "-01-01"
            baseDate = DateTime.Parse(reportStartDate)
            loopCount = DateTime.Parse(reportEndDate).Year - DateTime.Parse(reportStartDate).Year + 1
        End If

        If loopCount <= 0 Then Return

        Dim totalTuikushu As Decimal = 0
        Dim totalTuikujin As Decimal = 0
        Dim totalZongzhong As Decimal = 0
        Dim totalChengbengf As Decimal = 0
        Dim totalChengbenfj As Decimal = 0
        Dim totalChengben As Decimal = 0

        For i As Integer = 0 To loopCount - 1
            Dim loopDate As String = ""
            Dim loopStart As String = ""
            Dim loopEnd As String = ""

            If radioDay.Checked Then
                Dim d As DateTime = baseDate.AddDays(i)
                loopDate = d.ToString("yyyy-MM-dd")
                loopStart = loopDate & " 00:00:00"
                loopEnd = d.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"
            ElseIf radioMonth.Checked Then
                Dim d As DateTime = baseDate.AddMonths(i)
                loopDate = d.ToString("yyyy-MM")
                loopStart = d.ToString("yyyy-MM-dd") & " 00:00:00"
                loopEnd = d.AddMonths(1).ToString("yyyy-MM-dd") & " 00:00:00"
            ElseIf radioYear.Checked Then
                Dim d As DateTime = baseDate.AddYears(i)
                loopDate = d.ToString("yyyy")
                loopStart = d.ToString("yyyy-MM-dd") & " 00:00:00"
                loopEnd = d.AddYears(1).ToString("yyyy-MM-dd") & " 00:00:00"
            End If

            Dim innerSQL As String = BuildDateInnerSQL(loopStart, loopEnd, searchShopFilter, searchFactoryFilter, searchName)

            Dim sql As String = "SELECT " &
                "CAST(ROUND(COALESCE(sum(tol.tuikushu), 0),2) AS DECIMAL(30,2)) as tuikushu, " &
                "CAST(ROUND(COALESCE(sum(tol.tuikujin), 0),3) AS DECIMAL(30,3)) as tuikujin, " &
                "CAST(ROUND(COALESCE(sum(tol.zongzhong), 0),3) AS DECIMAL(30,3)) as zongzhong, " &
                "CAST(ROUND(COALESCE(sum(tol.chengbengf), 0),2) AS DECIMAL(30,2)) as chengbengf, " &
                "CAST(ROUND(COALESCE(sum(tol.chengbenfj), 0),2) AS DECIMAL(30,2)) as chengbenfj, " &
                "CAST(ROUND(COALESCE(sum(tol.chengben), 0),2) AS DECIMAL(30,2)) as chengben " &
                $"FROM ({innerSQL}) AS tol {searchSpecFilter}"

            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                Dim tuikushu As Decimal = SafeDecimal(row("tuikushu"))
                Dim tuikujin As Decimal = SafeDecimal(row("tuikujin"))
                Dim zongzhong As Decimal = SafeDecimal(row("zongzhong"))
                Dim chengbengf As Decimal = SafeDecimal(row("chengbengf"))
                Dim chengbenfj As Decimal = SafeDecimal(row("chengbenfj"))
                Dim chengben As Decimal = SafeDecimal(row("chengben"))

                dgvReport.Rows.Add(
                    Right("000" & (i + 1).ToString(), loopCount.ToString().Length),
                    loopDate,
                    tuikushu.ToString(),
                    tuikujin.ToString(),
                    zongzhong.ToString(),
                    chengbengf.ToString(),
                    chengbenfj.ToString(),
                    chengben.ToString()
                )

                totalTuikushu += tuikushu
                totalTuikujin += tuikujin
                totalZongzhong += zongzhong
                totalChengbengf += chengbengf
                totalChengbenfj += chengbenfj
                totalChengben += chengben
            End If
        Next

        ' 合计行
        dgvReport.Rows.Add("合计", "",
            totalTuikushu.ToString(),
            totalTuikujin.ToString(),
            totalZongzhong.ToString(),
            totalChengbengf.ToString(),
            totalChengbenfj.ToString(),
            totalChengben.ToString())
    End Sub

    ' ========== 单选框状态改变 ==========
    Private Sub RadioChanged(sender As Object, e As EventArgs)
        If isStateChanging Then Return
        isStateChanging = True

        Dim radio As RadioButton = DirectCast(sender, RadioButton)
        If Not radio.Checked Then
            isStateChanging = False
            Return
        End If

        ' 取消其他选中
        radioOrder.Checked = (radio Is radioOrder)
        radioDetail.Checked = (radio Is radioDetail)
        radioDay.Checked = (radio Is radioDay)
        radioMonth.Checked = (radio Is radioMonth)
        radioYear.Checked = (radio Is radioYear)

        ' 控制筛选条件启用/禁用
        Dim enableFilters As Boolean = Not radioOrder.Checked
        chkCategory.Enabled = enableFilters
        chkSpec.Enabled = enableFilters
        chkFactory.Enabled = enableFilters
        txtSearch.Enabled = enableFilters
        txtFactorySearch.Enabled = enableFilters

        ' 重置订单编码
        localOrderCode = ""

        ' 设置日期格式
        If radioOrder.Checked OrElse radioDetail.Checked OrElse radioDay.Checked Then
            dtpStart.Format = DateTimePickerFormat.Short
            dtpEnd.Format = DateTimePickerFormat.Short
            dtpStart.Value = DateTime.Now
            dtpEnd.Value = DateTime.Now
        ElseIf radioMonth.Checked Then
            dtpStart.Format = DateTimePickerFormat.Custom
            dtpStart.CustomFormat = "yyyy-MM"
            dtpEnd.Format = DateTimePickerFormat.Custom
            dtpEnd.CustomFormat = "yyyy-MM"
            dtpStart.Value = DateTime.Now
            dtpEnd.Value = DateTime.Now
        ElseIf radioYear.Checked Then
            dtpStart.Format = DateTimePickerFormat.Custom
            dtpStart.CustomFormat = "yyyy"
            dtpEnd.Format = DateTimePickerFormat.Custom
            dtpEnd.CustomFormat = "yyyy"
            dtpStart.Value = DateTime.Now
            dtpEnd.Value = DateTime.Now
        End If

        LoadTableHeader()
        ExecuteMainQuery()

        isStateChanging = False
    End Sub

    ' ========== 工厂查找 ==========
    Private Sub txtFactorySearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtFactorySearch.Text) Then
                ShowWarning("查找工厂不能为空！")
                Return
            End If
            SearchFactory(txtFactorySearch.Text)
            txtFactorySearch.Text = ""
        End If
    End Sub

    Private Sub SearchFactory(searchText As String)
        Try
            Dim safeText As String = SafeSQL(searchText)
            Dim sql As String = $"SELECT id, title FROM xipunum_erp_about WHERE title LIKE '%{safeText}%' OR jianxie LIKE '%{safeText}%' OR name LIKE '%{safeText}%' ORDER BY id asc"
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
            chkFactory.Items.Clear()
            If dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    Dim id As String = SafeString(dt.Rows(i)("id"))
                    Dim title As String = SafeString(dt.Rows(i)("title"))
                    chkFactory.Items.Add(title & "|" & id, False)
                Next
            Else
                ShowWarning("查询无此信息数据！")
            End If
        Catch ex As Exception
            ShowError("工厂查找失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 搜索信息回车 ==========
    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtSearch.Text) Then
                ShowWarning("查找信息不能为空！")
                Return
            End If
            ExecuteMainQuery()
            txtSearch.Text = ""
        End If
    End Sub

    ' ========== 右键菜单事件 ==========
    Private Sub dgvReport_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs)
        If e.Button = MouseButtons.Right AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If radioOrder.Checked Then
                Dim orderCode As String = SafeString(dgvReport.Rows(e.RowIndex).Cells(1).Value)
                If orderCode <> "" Then
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
        End If
    End Sub

    Private Sub menuItemDetail_Click(sender As Object, e As EventArgs)
        If dgvReport.CurrentRow IsNot Nothing Then
            localOrderCode = SafeString(dgvReport.CurrentRow.Cells(1).Value)
            LoadTableHeader()
            ExecuteMainQuery()
        End If
    End Sub

    Private Sub menuItemCopy_Click(sender As Object, e As EventArgs)
        If dgvReport.CurrentRow IsNot Nothing AndAlso dgvReport.CurrentCell IsNot Nothing Then
            Dim cellValue As String = SafeString(dgvReport.CurrentCell.Value)
            Clipboard.SetText(cellValue)
            ShowInfo($"复制:{cellValue} 到剪切板成功！")
        End If
    End Sub

    Private Sub menuItemReturn_Click(sender As Object, e As EventArgs)
        localOrderCode = ""
        LoadTableHeader()
        ExecuteMainQuery()
    End Sub

    ' ========== 按钮事件 ==========
    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
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
            ExportToExcel(dt, "商品退库报表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        ' 重置到初始状态
        localOrderCode = ""
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        dtpStart.Format = DateTimePickerFormat.Short
        dtpEnd.Format = DateTimePickerFormat.Short
        isStateChanging = True
        radioOrder.Checked = True
        radioDetail.Checked = False
        radioDay.Checked = False
        radioMonth.Checked = False
        radioYear.Checked = False
        chkCategory.Enabled = False
        chkSpec.Enabled = False
        chkFactory.Enabled = False
        txtSearch.Enabled = False
        txtFactorySearch.Enabled = False
        isStateChanging = False
        LoadShopList()
        LoadCategoryList()
        LoadSpecList()
        UpdateSpecIdList()
        LoadTableHeader()
        ExecuteMainQuery()
    End Sub
End Class

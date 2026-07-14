' ============================================================================
' 商品查询报表窗口
' 功能: 综合商品查询，支持多维度筛选（店铺/品类/规格/工厂/款号/金重/总重/圈号）
' 对应易语言: 窗口程序集_窗口_商品查询报表
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class ProductQueryReportForm
    Inherits System.Windows.Forms.Form

    ' ========== 查找条件变量 ==========
    Private searchStartDate As String = ""
    Private searchEndDate As String = ""
    Private searchShopFilter As String = ""
    Private searchSpecFilter As String = ""
    Private searchFactoryFilter As String = ""
    Private searchName As String = ""
    Private searchItemNoFilter As String = " "
    Private searchJinzhongFilter As String = ""
    Private searchWeightFilter As String = ""
    Private searchQuanduFilter As String = ""

    ' ========== 规格ID列表 ==========
    Private _specIdList As New List(Of String)()

    ' ========== 控件声明 ==========
    Private dgvReport As New DataGridView()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private chkShop As New CheckedListBox()
    Private chkCategory As New CheckedListBox()
    Private chkSpec As New CheckedListBox()
    Private chkFactory As New CheckedListBox()
    Private txtSearch As New TextBox()
    Private txtFactorySearch As New TextBox()
    Private txtJinzhong1 As New TextBox()
    Private txtJinzhong2 As New TextBox()
    Private txtWeight1 As New TextBox()
    Private txtWeight2 As New TextBox()
    Private txtQuandu1 As New TextBox()
    Private txtQuandu2 As New TextBox()
    Private chkJinzhong As New CheckBox()
    Private chkWeight As New CheckBox()
    Private chkQuandu As New CheckBox()
    Private radioAll As New RadioButton()
    Private radioEmpty As New RadioButton()
    Private radioNotEmpty As New RadioButton()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnReset As New Button()
    Private chkShopAll As New CheckBox()
    Private chkCategoryAll As New CheckBox()
    Private chkSpecAll As New CheckBox()
    Private chkFactoryAll As New CheckBox()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    ' ========== UI初始化 ==========
    Private Sub InitializeUI()
        Me.Text = "商品查询报表"
        Me.Size = New Drawing.Size(1400, 800)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' ========== 顶部筛选面板 ==========
        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 160
        panelTop.AutoScroll = True
        Me.Controls.Add(panelTop)

        ' 日期范围
        AddLabel(panelTop, "开始：", 20, 12)
        dtpStart.Location = New Drawing.Point(60, 9)
        dtpStart.Size = New Drawing.Size(130, 25)
        dtpStart.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpStart)

        AddLabel(panelTop, "结束：", 200, 12)
        dtpEnd.Location = New Drawing.Point(240, 9)
        dtpEnd.Size = New Drawing.Size(130, 25)
        dtpEnd.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpEnd)

        ' 查找信息
        AddLabel(panelTop, "查找：", 380, 12)
        txtSearch.Location = New Drawing.Point(420, 9)
        txtSearch.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtSearch)

        ' 款号单选
        AddLabel(panelTop, "款号：", 580, 12)
        radioAll.Text = "全部"
        radioAll.Location = New Drawing.Point(620, 10)
        radioAll.AutoSize = True
        panelTop.Controls.Add(radioAll)

        radioEmpty.Text = "为空"
        radioEmpty.Location = New Drawing.Point(680, 10)
        radioEmpty.AutoSize = True
        panelTop.Controls.Add(radioEmpty)

        radioNotEmpty.Text = "不为空"
        radioNotEmpty.Location = New Drawing.Point(740, 10)
        radioNotEmpty.AutoSize = True
        panelTop.Controls.Add(radioNotEmpty)

        ' 店铺多选
        AddLabel(panelTop, "店铺：", 20, 45)
        chkShop.Location = New Drawing.Point(60, 42)
        chkShop.Size = New Drawing.Size(160, 80)
        chkShop.CheckOnClick = True
        panelTop.Controls.Add(chkShop)
        chkShopAll.Text = "全选"
        chkShopAll.Location = New Drawing.Point(225, 45)
        chkShopAll.AutoSize = True
        panelTop.Controls.Add(chkShopAll)

        ' 品类多选
        AddLabel(panelTop, "品类：", 290, 45)
        chkCategory.Location = New Drawing.Point(330, 42)
        chkCategory.Size = New Drawing.Size(160, 80)
        chkCategory.CheckOnClick = True
        panelTop.Controls.Add(chkCategory)
        chkCategoryAll.Text = "全选"
        chkCategoryAll.Location = New Drawing.Point(495, 45)
        chkCategoryAll.AutoSize = True
        panelTop.Controls.Add(chkCategoryAll)

        ' 规格多选
        AddLabel(panelTop, "规格：", 560, 45)
        chkSpec.Location = New Drawing.Point(600, 42)
        chkSpec.Size = New Drawing.Size(160, 80)
        chkSpec.CheckOnClick = True
        panelTop.Controls.Add(chkSpec)
        chkSpecAll.Text = "全选"
        chkSpecAll.Location = New Drawing.Point(765, 45)
        chkSpecAll.AutoSize = True
        panelTop.Controls.Add(chkSpecAll)

        ' 工厂多选
        AddLabel(panelTop, "工厂：", 830, 45)
        chkFactory.Location = New Drawing.Point(870, 42)
        chkFactory.Size = New Drawing.Size(160, 80)
        chkFactory.CheckOnClick = True
        panelTop.Controls.Add(chkFactory)
        chkFactoryAll.Text = "全选"
        chkFactoryAll.Location = New Drawing.Point(1035, 45)
        chkFactoryAll.AutoSize = True
        panelTop.Controls.Add(chkFactoryAll)

        ' 工厂查找
        AddLabel(panelTop, "工厂查找：", 830, 128)
        txtFactorySearch.Location = New Drawing.Point(890, 125)
        txtFactorySearch.Size = New Drawing.Size(120, 25)
        panelTop.Controls.Add(txtFactorySearch)

        ' 金重范围
        chkJinzhong.Text = "金重"
        chkJinzhong.Location = New Drawing.Point(20, 128)
        chkJinzhong.AutoSize = True
        panelTop.Controls.Add(chkJinzhong)
        txtJinzhong1.Location = New Drawing.Point(75, 126)
        txtJinzhong1.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(txtJinzhong1)
        AddLabel(panelTop, "-", 140, 128)
        txtJinzhong2.Location = New Drawing.Point(150, 126)
        txtJinzhong2.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(txtJinzhong2)

        ' 总重范围
        chkWeight.Text = "总重"
        chkWeight.Location = New Drawing.Point(225, 128)
        chkWeight.AutoSize = True
        panelTop.Controls.Add(chkWeight)
        txtWeight1.Location = New Drawing.Point(280, 126)
        txtWeight1.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(txtWeight1)
        AddLabel(panelTop, "-", 345, 128)
        txtWeight2.Location = New Drawing.Point(355, 126)
        txtWeight2.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(txtWeight2)

        ' 圈号范围
        chkQuandu.Text = "圈号"
        chkQuandu.Location = New Drawing.Point(430, 128)
        chkQuandu.AutoSize = True
        panelTop.Controls.Add(chkQuandu)
        txtQuandu1.Location = New Drawing.Point(485, 126)
        txtQuandu1.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(txtQuandu1)
        AddLabel(panelTop, "-", 550, 128)
        txtQuandu2.Location = New Drawing.Point(560, 126)
        txtQuandu2.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(txtQuandu2)

        ' 按钮
        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(640, 126)
        btnQuery.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnQuery)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(730, 126)
        btnExport.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnExport)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(820, 126)
        btnReset.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnReset)

        ' DataGridView
        dgvReport.Dock = DockStyle.Fill
        dgvReport.ReadOnly = True
        dgvReport.AllowUserToAddRows = False
        dgvReport.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvReport.RowHeadersWidth = 40
        Me.Controls.Add(dgvReport)
        dgvReport.BringToFront()

        ' 事件绑定
        AddHandler chkShopAll.CheckedChanged, AddressOf chkShopAll_CheckedChanged
        AddHandler chkCategoryAll.CheckedChanged, AddressOf chkCategoryAll_CheckedChanged
        AddHandler chkSpecAll.CheckedChanged, AddressOf chkSpecAll_CheckedChanged
        AddHandler chkFactoryAll.CheckedChanged, AddressOf chkFactoryAll_CheckedChanged
        AddHandler chkCategory.ItemCheck, AddressOf chkCategory_ItemCheck
        AddHandler chkSpec.ItemCheck, AddressOf chkSpec_ItemCheck
        AddHandler txtFactorySearch.KeyDown, AddressOf txtFactorySearch_KeyDown
        AddHandler txtSearch.KeyDown, AddressOf txtSearch_KeyDown
        AddHandler radioAll.CheckedChanged, AddressOf radioAll_CheckedChanged
        AddHandler radioEmpty.CheckedChanged, AddressOf radioEmpty_CheckedChanged
        AddHandler radioNotEmpty.CheckedChanged, AddressOf radioNotEmpty_CheckedChanged
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
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        radioAll.Checked = True
        LoadShopList()
        LoadCategoryList()
        LoadSpecList()
        UpdateSpecIdList()
        LoadHeaders()
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
        Dim sql As String = "SELECT id, title FROM xipunum_erp_category WHERE 1=1"
        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
        For Each row As DataRow In dt.Rows
            chkCategory.Items.Add(New With {.ID = SafeString(row("id")), .Title = SafeString(row("title"))})
        Next
        For i As Integer = 0 To chkCategory.Items.Count - 1
            chkCategory.SetItemChecked(i, True)
        Next
    End Sub

    ' ========== 加载规格列表 ==========
    Private Sub LoadSpecList()
        chkSpec.Items.Clear()
        Dim sql As String = "SELECT title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title ORDER BY id asc"
        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
        For Each row As DataRow In dt.Rows
            chkSpec.Items.Add(SafeString(row("title")))
        Next
    End Sub

    ' ========== 品类选中变化时刷新规格列表 ==========
    Private Sub ReloadSpecList()
        Dim categoryIds As New List(Of String)()
        For i As Integer = 0 To chkCategory.Items.Count - 1
            If chkCategory.GetItemChecked(i) Then
                Dim item = chkCategory.Items(i)
                categoryIds.Add(SafeString(item.ID))
            End If
        Next

        chkSpec.Items.Clear()
        Dim sql As String
        If categoryIds.Count = 0 Then
            sql = "SELECT title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title ORDER BY id asc"
        Else
            sql = "SELECT title FROM xipunum_erp_specs WHERE category_id in (" & String.Join(",", categoryIds) & ") GROUP BY title ORDER BY id asc"
        End If
        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
        For Each row As DataRow In dt.Rows
            chkSpec.Items.Add(SafeString(row("title")))
        Next
    End Sub

    ' ========== 更新规格ID列表 ==========
    Private Sub UpdateSpecIdList()
        _specIdList.Clear()

        Dim categoryIds As New List(Of String)()
        For i As Integer = 0 To chkCategory.Items.Count - 1
            If chkCategory.GetItemChecked(i) Then
                Dim item = chkCategory.Items(i)
                categoryIds.Add("'" & SafeString(item.ID) & "'")
            End If
        Next

        Dim specTitles As New List(Of String)()
        For i As Integer = 0 To chkSpec.Items.Count - 1
            If chkSpec.GetItemChecked(i) Then
                specTitles.Add("'" & SafeString(chkSpec.Items(i)) & "'")
            End If
        Next

        Dim catFilter As String = ""
        If categoryIds.Count > 0 Then
            catFilter = " and category_id in (" & String.Join(",", categoryIds) & ")"
        End If

        Dim specFilter As String = ""
        If specTitles.Count > 0 Then
            specFilter = " and title in (" & String.Join(",", specTitles) & ")"
        End If

        Dim sql As String = "SELECT id, title FROM xipunum_erp_specs WHERE 1=1 " & catFilter & specFilter & " GROUP BY id ORDER BY id asc"
        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
        For Each row As DataRow In dt.Rows
            _specIdList.Add(SafeString(row("id")))
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
        ReloadSpecList()
        UpdateSpecIdList()
    End Sub

    Private Sub chkSpecAll_CheckedChanged(sender As Object, e As EventArgs)
        For i As Integer = 0 To chkSpec.Items.Count - 1
            chkSpec.SetItemChecked(i, chkSpecAll.Checked)
        Next
        UpdateSpecIdList()
    End Sub

    Private Sub chkFactoryAll_CheckedChanged(sender As Object, e As EventArgs)
        For i As Integer = 0 To chkFactory.Items.Count - 1
            chkFactory.SetItemChecked(i, chkFactoryAll.Checked)
        Next
    End Sub

    ' ========== 品类/规格选中变化 ==========
    Private Sub chkCategory_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        BeginInvoke(New Action(Sub()
            ReloadSpecList()
            UpdateSpecIdList()
        End Sub))
    End Sub

    Private Sub chkSpec_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        BeginInvoke(New Action(Sub()
            UpdateSpecIdList()
        End Sub))
    End Sub

    ' ========== 款号单选 ==========
    Private Sub radioAll_CheckedChanged(sender As Object, e As EventArgs)
        If radioAll.Checked Then
            radioEmpty.Checked = False
            radioNotEmpty.Checked = False
        End If
    End Sub

    Private Sub radioEmpty_CheckedChanged(sender As Object, e As EventArgs)
        If radioEmpty.Checked Then
            radioAll.Checked = False
            radioNotEmpty.Checked = False
        End If
    End Sub

    Private Sub radioNotEmpty_CheckedChanged(sender As Object, e As EventArgs)
        If radioNotEmpty.Checked Then
            radioAll.Checked = False
            radioEmpty.Checked = False
        End If
    End Sub

    ' ========== 工厂查找 ==========
    Private Sub txtFactorySearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtFactorySearch.Text) Then
                ShowWarning("查找工厂不能为空！")
                Return
            End If
            SearchFactory()
            txtFactorySearch.Text = ""
        End If
    End Sub

    Private Sub SearchFactory()
        Dim searchText As String = txtFactorySearch.Text
        chkFactory.Items.Clear()
        Dim sql As String = "SELECT id, title FROM xipunum_erp_about WHERE title like '%" & searchText & "%' OR jianxie like '%" & searchText & "%' OR name like '%" & searchText & "%' ORDER BY id asc"
        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
        For Each row As DataRow In dt.Rows
            chkFactory.Items.Add(New With {.ID = SafeString(row("id")), .Title = SafeString(row("title"))})
        Next
        If dt.Rows.Count = 0 Then
            ShowWarning("查询无此信息数据！")
        End If
    End Sub

    ' ========== 查找信息回车 ==========
    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtSearch.Text) Then
                ShowWarning("查找信息不能为空！")
                Return
            End If
            btnQuery_Click(Nothing, Nothing)
            txtSearch.Text = ""
        End If
    End Sub

    ' ========== 加载表头 ==========
    Private Sub LoadHeaders()
        dgvReport.Columns.Clear()
        dgvReport.Rows.Clear()

        Dim headers As String() = {"序号", "商品编码", "原始编码", "入库单号", "入库时间", "品类名称", "规格", "材质",
                                   "商品名称", "款号", "是否镶嵌", "工厂", "单件重", "库存数量", "库存金重", "库存总重",
                                   "库房", "成本价", "预售价", "成本工费", "成本附加费", "参考工费", "销售工费", "销售附加费", "原料克价"}
        Dim widths As Integer() = {45, 100, 100, 130, 130, 80, 80, 65, 200, 120, 75, 100, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75}
        For i As Integer = 0 To headers.Length - 1
            dgvReport.Columns.Add("col" & i, headers(i))
            dgvReport.Columns(i).Width = widths(i)
        Next
    End Sub

    ' ========== 查询按钮 ==========
    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Try
            btnQuery.Enabled = False
            btnReset.Enabled = False
            btnExport.Enabled = False

            ' 查找信息
            If String.IsNullOrEmpty(txtSearch.Text) Then
                searchName = ""
            Else
                searchName = txtSearch.Text
            End If

            ' 店铺过滤（必选）
            Dim shopIds As New List(Of String)()
            For i As Integer = 0 To chkShop.Items.Count - 1
                If chkShop.GetItemChecked(i) Then
                    Dim item = chkShop.Items(i)
                    shopIds.Add("'" & SafeString(item.ID) & "'")
                End If
            Next

            If shopIds.Count = 0 Then
                ShowWarning("请选择需要查询的店铺！")
                btnQuery.Enabled = True
                btnReset.Enabled = True
                btnExport.Enabled = True
                Return
            End If
            searchShopFilter = " and b.kufang in (" & String.Join(",", shopIds) & ")"

            ' 规格过滤
            UpdateSpecIdList()
            If _specIdList.Count > 0 Then
                searchSpecFilter = " and COALESCE(e1.id, e2.id, '0') in ('" & String.Join("','", _specIdList) & "')"
            Else
                searchSpecFilter = ""
            End If

            ' 工厂过滤
            Dim factoryIds As New List(Of String)()
            For i As Integer = 0 To chkFactory.Items.Count - 1
                If chkFactory.GetItemChecked(i) Then
                    Dim item = chkFactory.Items(i)
                    factoryIds.Add("'" & SafeString(item.ID) & "'")
                End If
            Next
            If factoryIds.Count > 0 Then
                searchFactoryFilter = " and f.factory in (" & String.Join(",", factoryIds) & ")"
            Else
                searchFactoryFilter = ""
            End If

            ' 日期
            searchStartDate = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim endDay As DateTime = dtpEnd.Value.AddDays(1)
            searchEndDate = endDay.ToString("yyyy-MM-dd") & " 00:00:00"

            ' 款号过滤
            If radioAll.Checked Then
                searchItemNoFilter = " "
            ElseIf radioEmpty.Checked Then
                searchItemNoFilter = " and (a.item_number = '' or a.item_number IS NULL)"
            ElseIf radioNotEmpty.Checked Then
                searchItemNoFilter = " and a.item_number <> ''"
            End If

            ' 金重范围
            If chkJinzhong.Checked Then
                searchJinzhongFilter = " and b.jinzhong >= '" & txtJinzhong1.Text & "' and b.jinzhong <= '" & txtJinzhong2.Text & "'"
            Else
                searchJinzhongFilter = ""
            End If

            ' 总重范围
            If chkWeight.Checked Then
                searchWeightFilter = " and a.weight >= '" & txtWeight1.Text & "' and a.weight <= '" & txtWeight2.Text & "'"
            Else
                searchWeightFilter = ""
            End If

            ' 圈号范围
            If chkQuandu.Checked Then
                searchQuanduFilter = " and a.quandu >= '" & txtQuandu1.Text & "' and a.quandu <= '" & txtQuandu2.Text & "'"
            Else
                searchQuanduFilter = ""
            End If

            ' 执行查询
            QueryData()

        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        Finally
            btnQuery.Enabled = True
            btnReset.Enabled = True
            btnExport.Enabled = True
        End Try
    End Sub

    ' ========== 构建内层SQL ==========
    Private Function BuildInnerSQL() As String
        Dim sql As String = "SELECT b.poduct_code AS bianma, a.fu_code AS afu_code, f.odd_numbers as rukudanhao, " &
                            "c.creationtime AS rukutime, " &
                            "CASE WHEN COALESCE(d.title, '') = '' THEN '未匹配' ELSE d.title END AS pinlei, " &
                            "COALESCE(e1.title, e2.title, '未匹配') AS guige, " &
                            "a.caizhi AS acaizhi, a.product_name AS mingcheng, a.item_number AS kuanhao, " &
                            "CASE WHEN COALESCE(i.xiangqian, a.xiangqian) = '' THEN '' ELSE COALESCE(i.xiangqian, a.xiangqian) END AS xiangqian, " &
                            "CASE WHEN f.factory = '' THEN '' ELSE g.title END AS gongchang, " &
                            "CAST(ROUND(a.single, 3) AS DECIMAL(10,3)) AS danjian, " &
                            "CAST(ROUND(b.quantity, 2) AS DECIMAL(10,2)) AS kucun, " &
                            "CAST(ROUND(b.jinzhong, 3) AS DECIMAL(10,3)) AS jinzhong, " &
                            "CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN b.jinzhong ELSE CAST(ROUND(a.weight / a.quantity * b.quantity, 3) AS DECIMAL(10,3)) END AS zongzhong, " &
                            "CASE WHEN b.kufang = '0' THEN '总库' ELSE h.title END AS kufang, " &
                            "CAST(ROUND(CASE WHEN a.quantity = '0' THEN c.cost_price*b.jinzhong/a.jin_zhong ELSE c.cost_price*b.quantity END, 2) AS DECIMAL(20,2)) AS chengben, " &
                            "CAST(ROUND(c.sales_price, 2) AS DECIMAL(20,2)) AS yushou, " &
                            "CAST(ROUND(c.basic_cost, 2) AS DECIMAL(20,2)) AS chengbengf, " &
                            "CAST(ROUND(c.company_surcharge, 2) AS DECIMAL(20,2)) AS chengbenfj, " &
                            "CAST(ROUND(c.premium_cost, 2) AS DECIMAL(20,2)) AS cankaogf, " &
                            "CAST(ROUND(c.sales_cost, 2) AS DECIMAL(20,2)) AS xiaoshougf, " &
                            "CAST(ROUND(c.sales_surcharge, 2) AS DECIMAL(20,2)) AS xiaoshoufj, " &
                            "f.gold_price as ylkejia " &
                            "FROM xipunum_erp_shop_kucun AS b " &
                            "INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code " &
                            "INNER JOIN xipunum_erp_store AS c ON c.poduct_code = b.poduct_code " &
                            "INNER JOIN xipunum_erp_store_order AS f ON f.id = c.order_id " &
                            "LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang " &
                            "LEFT JOIN xipunum_erp_about AS g ON g.id = f.factory " &
                            searchFactoryFilter &
                            " LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = a.item_number AND a.item_number != '' " &
                            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id AND a.item_number IS NOT NULL AND a.item_number != '' " &
                            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id AND a.specification_id IS NOT NULL AND a.specification_id != '' " &
                            "LEFT JOIN xipunum_erp_category AS d ON d.id = COALESCE(e1.category_id, e2.category_id) AND COALESCE(e1.category_id, e2.category_id) IS NOT NULL " &
                            "WHERE (a.poduct_code LIKE '%" & searchName & "%' OR a.fu_code LIKE '%" & searchName & "%' OR a.product_name LIKE '%" & searchName & "%' OR a.item_number LIKE '%" & searchName & "%' OR a.caizhi LIKE '%" & searchName & "%') " &
                            "AND c.creationtime >= '" & searchStartDate & "' AND c.creationtime < '" & searchEndDate & "' " &
                            searchShopFilter & searchItemNoFilter & searchSpecFilter & searchJinzhongFilter & searchWeightFilter & searchQuanduFilter &
                            " ORDER BY c.creationtime DESC"
        Return sql
    End Function

    ' ========== 数据查询 ==========
    Private Sub QueryData()
        dgvReport.Rows.Clear()

        Dim innerSql As String = BuildInnerSQL()
        Dim dt As DataTable = ExecuteQuery(innerSql, MySQL_ReadReport)

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            Dim seq As String = (i + 1).ToString().PadLeft(dt.Rows.Count.ToString().Length, "0"c)
            dgvReport.Rows.Add(
                seq,
                SafeString(row("bianma")),
                SafeString(row("afu_code")),
                SafeString(row("rukudanhao")),
                SafeString(row("rukutime")),
                SafeString(row("pinlei")),
                SafeString(row("guige")),
                SafeString(row("acaizhi")),
                SafeString(row("mingcheng")),
                SafeString(row("kuanhao")),
                SafeString(row("xiangqian")),
                SafeString(row("gongchang")),
                SafeString(row("danjian")),
                SafeString(row("kucun")),
                SafeString(row("jinzhong")),
                SafeString(row("zongzhong")),
                SafeString(row("kufang")),
                SafeString(row("chengben")),
                SafeString(row("yushou")),
                SafeString(row("chengbengf")),
                SafeString(row("chengbenfj")),
                SafeString(row("cankaogf")),
                SafeString(row("xiaoshougf")),
                SafeString(row("xiaoshoufj")),
                SafeString(row("ylkejia"))
            )
        Next

        ' 合计行
        If dt.Rows.Count > 0 Then
            Dim sumSql As String = "SELECT CAST(ROUND(sum(sum.kucun), 3) AS DECIMAL(20,3)) AS kucun, " &
                                   "CAST(ROUND(sum(sum.jinzhong), 3) AS DECIMAL(20,3)) AS jinzhong, " &
                                   "CAST(ROUND(sum(sum.zongzhong), 3) AS DECIMAL(20,3)) AS zongzhong, " &
                                   "CAST(ROUND(sum(sum.chengben), 2) AS DECIMAL(20,2)) AS chengben, " &
                                   "CAST(ROUND(sum(sum.yushou), 2) AS DECIMAL(20,2)) AS yushou, " &
                                   "CAST(ROUND(sum(sum.chengbenfj), 2) AS DECIMAL(20,2)) AS chengbenfj, " &
                                   "CAST(ROUND(sum(sum.xiaoshoufj), 2) AS DECIMAL(20,2)) AS xiaoshoufj " &
                                   "FROM (" & innerSql & ") as sum"
            Dim sumDt As DataTable = ExecuteQuery(sumSql, MySQL_ReadReport)
            If sumDt.Rows.Count > 0 Then
                Dim sumRow As DataRow = sumDt.Rows(0)
                Dim totalIdx As Integer = dgvReport.Rows.Add()
                dgvReport.Rows(totalIdx).Cells(1).Value = "合计"
                dgvReport.Rows(totalIdx).Cells(13).Value = SafeString(sumRow("kucun"))
                dgvReport.Rows(totalIdx).Cells(14).Value = SafeString(sumRow("jinzhong"))
                dgvReport.Rows(totalIdx).Cells(15).Value = SafeString(sumRow("zongzhong"))
                dgvReport.Rows(totalIdx).Cells(17).Value = SafeString(sumRow("chengben"))
                dgvReport.Rows(totalIdx).Cells(18).Value = SafeString(sumRow("yushou"))
                dgvReport.Rows(totalIdx).Cells(20).Value = SafeString(sumRow("chengbenfj"))
                dgvReport.Rows(totalIdx).Cells(23).Value = SafeString(sumRow("xiaoshoufj"))
            End If
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
            ExportToExcel(dt, "商品查询报表表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 重置按钮 ==========
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        txtSearch.Text = ""
        txtFactorySearch.Text = ""
        txtJinzhong1.Text = ""
        txtJinzhong2.Text = ""
        txtWeight1.Text = ""
        txtWeight2.Text = ""
        txtQuandu1.Text = ""
        txtQuandu2.Text = ""
        chkJinzhong.Checked = False
        chkWeight.Checked = False
        chkQuandu.Checked = False
        radioAll.Checked = True
        For i As Integer = 0 To chkShop.Items.Count - 1
            chkShop.SetItemChecked(i, False)
        Next
        If chkShop.Items.Count > 0 Then
            chkShop.SetItemChecked(0, True)
        End If
        For i As Integer = 0 To chkCategory.Items.Count - 1
            chkCategory.SetItemChecked(i, True)
        Next
        chkFactory.Items.Clear()
        ReloadSpecList()
        UpdateSpecIdList()
        dgvReport.Rows.Clear()
    End Sub
End Class

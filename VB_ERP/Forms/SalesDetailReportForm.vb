' ============================================================================
' 销售详情报表窗口
' 功能: 销售明细查询，支持品类/自定义两种分组模式，三级层次化展示（店铺→品类→明细）
' 对应易语言: 窗口程序集_窗口_销售详情报表
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SalesDetailReportForm
    Inherits System.Windows.Forms.Form

    ' ========== 查找条件变量 ==========
    Private searchStartDate As String = ""
    Private searchEndDate As String = ""
    Private searchShopFilter As String = ""
    Private searchMaterialFilter As String = ""
    Private searchPilingFilter As String = " "
    Private searchJinzhongFilter As String = ""

    ' ========== 控件声明 ==========
    Private dgvReport As New DataGridView()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private chkShop As New CheckedListBox()
    Private chkMaterial As New CheckedListBox()
    Private chkShopAll As New CheckBox()
    Private chkMaterialAll As New CheckBox()
    Private radioAll As New RadioButton()
    Private radioPifa As New RadioButton()
    Private radioLing As New RadioButton()
    Private radioCategory As New RadioButton()
    Private radioCustom As New RadioButton()
    Private chkJinzhong As New CheckBox()
    Private txtJinzhong1 As New TextBox()
    Private txtJinzhong2 As New TextBox()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnReset As New Button()
    Private contextMenuReport As New ContextMenuStrip()
    Private menuItemCopy As New ToolStripMenuItem("复制单元格")

    ' ========== 颜色 ==========
    Private ReadOnly RedColor As Drawing.Color = Drawing.ColorTranslator.FromHtml("#E0422F")
    Private ReadOnly BlueColor As Drawing.Color = Drawing.ColorTranslator.FromHtml("#0A8A6E")

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    ' ========== UI初始化 ==========
    Private Sub InitializeUI()
        Me.Text = "销售详情报表"
        Me.Size = New Drawing.Size(1400, 800)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 100
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

        ' 批零单选
        AddLabel(panelTop, "批零：", 380, 12)
        radioAll.Text = "全部"
        radioAll.Location = New Drawing.Point(420, 10)
        radioAll.AutoSize = True
        panelTop.Controls.Add(radioAll)

        radioPifa.Text = "批发"
        radioPifa.Location = New Drawing.Point(480, 10)
        radioPifa.AutoSize = True
        panelTop.Controls.Add(radioPifa)

        radioLing.Text = "零售"
        radioLing.Location = New Drawing.Point(540, 10)
        radioLing.AutoSize = True
        panelTop.Controls.Add(radioLing)

        ' 显示模式单选
        AddLabel(panelTop, "模式：", 620, 12)
        radioCategory.Text = "品类"
        radioCategory.Location = New Drawing.Point(660, 10)
        radioCategory.AutoSize = True
        panelTop.Controls.Add(radioCategory)

        radioCustom.Text = "自定义"
        radioCustom.Location = New Drawing.Point(720, 10)
        radioCustom.AutoSize = True
        panelTop.Controls.Add(radioCustom)

        ' 金重范围
        chkJinzhong.Text = "金重"
        chkJinzhong.Location = New Drawing.Point(820, 10)
        chkJinzhong.AutoSize = True
        panelTop.Controls.Add(chkJinzhong)
        txtJinzhong1.Location = New Drawing.Point(875, 8)
        txtJinzhong1.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(txtJinzhong1)
        AddLabel(panelTop, "-", 940, 12)
        txtJinzhong2.Location = New Drawing.Point(950, 8)
        txtJinzhong2.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(txtJinzhong2)

        ' 店铺多选
        AddLabel(panelTop, "店铺：", 20, 42)
        chkShop.Location = New Drawing.Point(60, 39)
        chkShop.Size = New Drawing.Size(200, 55)
        chkShop.CheckOnClick = True
        chkShop.DisplayMember = "Title"
        panelTop.Controls.Add(chkShop)
        chkShopAll.Text = "全选"
        chkShopAll.Location = New Drawing.Point(265, 42)
        chkShopAll.AutoSize = True
        panelTop.Controls.Add(chkShopAll)

        ' 材质多选
        AddLabel(panelTop, "材质：", 360, 42)
        chkMaterial.Location = New Drawing.Point(400, 39)
        chkMaterial.Size = New Drawing.Size(200, 55)
        chkMaterial.CheckOnClick = True
        panelTop.Controls.Add(chkMaterial)
        chkMaterialAll.Text = "全选"
        chkMaterialAll.Location = New Drawing.Point(605, 42)
        chkMaterialAll.AutoSize = True
        panelTop.Controls.Add(chkMaterialAll)

        ' 按钮
        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(820, 37)
        btnQuery.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnQuery)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(910, 37)
        btnExport.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnExport)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(1000, 37)
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

        ' 右键菜单
        contextMenuReport.Items.Add(menuItemCopy)

        ' 事件绑定
        AddHandler chkShopAll.CheckedChanged, AddressOf chkShopAll_CheckedChanged
        AddHandler chkMaterialAll.CheckedChanged, AddressOf chkMaterialAll_CheckedChanged
        AddHandler radioAll.CheckedChanged, AddressOf radioPiling_CheckedChanged
        AddHandler radioPifa.CheckedChanged, AddressOf radioPiling_CheckedChanged
        AddHandler radioLing.CheckedChanged, AddressOf radioPiling_CheckedChanged
        AddHandler radioCategory.CheckedChanged, AddressOf radioMode_CheckedChanged
        AddHandler radioCustom.CheckedChanged, AddressOf radioMode_CheckedChanged
        AddHandler menuItemCopy.Click, AddressOf menuItemCopy_Click
        AddHandler dgvReport.CellMouseDown, AddressOf dgvReport_CellMouseDown

        InitGrid()
    End Sub

    Private Sub AddLabel(parent As Control, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    ' ========== 表头初始化 ==========
    Private Sub InitGrid()
        dgvReport.Columns.Clear()
        Dim headers() As String = {"序号", "出库单号", "商品编码", "销售时间", "商品名称", "款号", "库房", "品类", "规格", "材质", "圈口/长度", "成色", "单件重", "金重", "重量", "单位", "料价", "成本工费", "参考工费", "成本附加费", "成本价", "销售单价", "销售金额", "数量", "原附加费", "销售克价", "销售工费", "销售附加费", "折扣", "应收金额", "工费利润", "成本工费合计", "销售工费合计", "批零", "状态"}
        Dim widths() As Integer = {45, 150, 100, 100, 140, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvReport.Columns.Add(col)
        Next
    End Sub

    ' ========== 窗体加载 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        Me.Text = "销售详情报表 时间:" & dtpStart.Value.ToString("yyyy-MM-dd") & "至" & dtpEnd.Value.ToString("yyyy-MM-dd")
        radioAll.Checked = True
        radioCategory.Checked = True
        LoadShopList()
        LoadMaterialList()
        btnQuery_Click(Nothing, Nothing)
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

    ' ========== 加载材质列表 ==========
    Private Sub LoadMaterialList()
        chkMaterial.Items.Clear()
        Dim sql As String = "SELECT caizhi FROM xipunum_erp_shop WHERE 1=1 GROUP BY caizhi ORDER BY id ASC"
        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
        For Each row As DataRow In dt.Rows
            chkMaterial.Items.Add(SafeString(row("caizhi")))
        Next
        For i As Integer = 0 To chkMaterial.Items.Count - 1
            chkMaterial.SetItemChecked(i, True)
        Next
    End Sub

    ' ========== 全选 ==========
    Private Sub chkShopAll_CheckedChanged(sender As Object, e As EventArgs)
        For i As Integer = 0 To chkShop.Items.Count - 1
            chkShop.SetItemChecked(i, chkShopAll.Checked)
        Next
    End Sub

    Private Sub chkMaterialAll_CheckedChanged(sender As Object, e As EventArgs)
        For i As Integer = 0 To chkMaterial.Items.Count - 1
            chkMaterial.SetItemChecked(i, chkMaterialAll.Checked)
        Next
    End Sub

    ' ========== 批零单选 ==========
    Private Sub radioPiling_CheckedChanged(sender As Object, e As EventArgs)
        If radioAll.Checked Then
            searchPilingFilter = " "
        ElseIf radioPifa.Checked Then
            searchPilingFilter = " and a.pling =  '批发'"
        ElseIf radioLing.Checked Then
            searchPilingFilter = " and a.pling =  '零售'"
        End If
    End Sub

    ' ========== 模式单选 ==========
    Private Sub radioMode_CheckedChanged(sender As Object, e As EventArgs)
        If sender Is radioCategory Then
            radioCustom.Checked = Not radioCategory.Checked
        ElseIf sender Is radioCustom Then
            radioCategory.Checked = Not radioCustom.Checked
        End If
    End Sub

    ' ========== 查询按钮 ==========
    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        btnQuery.Enabled = False
        btnReset.Enabled = False
        btnExport.Enabled = False
        Me.Text = "销售详情报表 时间:" & dtpStart.Value.ToString("yyyy-MM-dd") & "至" & dtpEnd.Value.ToString("yyyy-MM-dd")

        ' 店铺筛选
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
        searchShopFilter = " and a.kufang in (" & String.Join(",", shopIds) & ")"

        ' 材质筛选
        Dim materialNames As New List(Of String)()
        For i As Integer = 0 To chkMaterial.Items.Count - 1
            If chkMaterial.GetItemChecked(i) Then
                materialNames.Add("'" & SafeString(chkMaterial.Items(i)) & "'")
            End If
        Next
        If materialNames.Count > 0 Then
            searchMaterialFilter = " and e.caizhi in (" & String.Join(",", materialNames) & ")"
        Else
            searchMaterialFilter = ""
        End If

        ' 金重范围
        If chkJinzhong.Checked Then
            searchJinzhongFilter = " and a.net_weight >= '" & txtJinzhong1.Text & "' and b.net_weight <= '" & txtJinzhong2.Text & "'"
        Else
            searchJinzhongFilter = ""
        End If

        ' 日期
        searchStartDate = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
        searchEndDate = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"

        ' 执行查询
        If radioCategory.Checked Then
            QueryDataCategory()
        Else
            QueryDataCustom()
        End If

        btnQuery.Enabled = True
        btnReset.Enabled = True
        btnExport.Enabled = True
    End Sub

    ' ========== 构建内层SQL（汇总用）==========
    Private Function BuildInnerSQL(Optional shopOverride As String = Nothing) As String
        Dim shopFilter As String = searchShopFilter
        If shopOverride IsNot Nothing Then
            shopFilter = " and a.kufang='" & shopOverride & "'"
        End If
        Return "SELECT CASE WHEN a.kufang = '0' THEN '0' ELSE a.kufang END AS kufangid, " &
            "CASE WHEN COALESCE(f.id, '') = '' THEN '0' ELSE f.id END AS pinleiid, " &
            "COALESCE(e1.category_id, e2.category_id, '0') AS guigeid, " &
            "CASE WHEN a.kufang = '0' THEN '总库' ELSE h.title END AS kufang, " &
            "b.settlement_number as fsettlement_number, a.poduct_code as outbound_product_code, " &
            "e.product_name as shop_product_name, e.item_number as shop_item_number, " &
            "CASE WHEN COALESCE(f.title, '') = '' THEN '未匹配' ELSE f.title END AS pinlei, " &
            "COALESCE(e1.title, e2.title, '未匹配') AS specification_title, " &
            "e.caizhi as material, e.quandu as quandu, c.factory_condition as company_condition, " &
            "e.single as danjian, a.net_weight as jinzhong, " &
            "CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN a.net_weight ELSE CAST(ROUND(e.weight / e.quantity * a.quantity, 3) AS DECIMAL(10, 3)) END AS zhongliang, " &
            "e.sales_unit as danwei, " &
            "CASE WHEN a.sales_return = '0' THEN c.basic_cost ELSE -c.basic_cost END AS chengbengf, " &
            "CASE WHEN a.sales_return = '0' THEN c.premium_cost ELSE -c.premium_cost END AS cankaogf, " &
            "CASE WHEN a.sales_return = '0' THEN c.company_surcharge ELSE -c.company_surcharge END AS chengbenfj, " &
            "CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN CAST(ROUND(c.cost_price / e.jin_zhong * a.net_weight, 2) AS DECIMAL(30, 2)) " &
            "ELSE CAST(ROUND(c.cost_price * a.quantity, 2) AS DECIMAL(30, 2)) END AS chengben, " &
            "a.xiaodan_amount as xiaodan_amount, a.xiao_amount as xiao_amount, a.quantity as xsshuliang, " &
            "CASE WHEN a.sales_return = '0' THEN c.sales_surcharge ELSE -c.sales_surcharge END AS yxiaoshoufj, " &
            "a.gold_price AS kejia, a.sales_cost AS xiaoshougf, a.sales_surcharge AS xiaoshoufj, " &
            "a.zhekou AS zhekou, a.settlement AS shishou, a.pling AS piling, " &
            "CASE WHEN a.sales_return = '0' THEN '客销' ELSE '客退' END AS sales_return, " &
            "CAST(ROUND(CASE WHEN a.sales_return = 1 THEN -a.sales_cost * a.net_weight + a.sales_surcharge ELSE a.sales_cost * a.net_weight + a.sales_surcharge END - " &
            "CASE WHEN a.sales_return = 1 THEN c.basic_cost * a.net_weight - c.company_surcharge ELSE c.basic_cost * a.net_weight + c.company_surcharge END, 2) AS DECIMAL(30, 2)) AS gongfeilr, " &
            "CAST(ROUND(CASE WHEN a.sales_return = 1 THEN c.basic_cost * a.net_weight - c.company_surcharge ELSE c.basic_cost * a.net_weight + c.company_surcharge END, 2) AS DECIMAL(20, 2)) AS chegbengfhj, " &
            "CAST(ROUND(CASE WHEN a.sales_return = 1 THEN -a.sales_cost * a.net_weight + a.sales_surcharge ELSE a.sales_cost * a.net_weight + a.sales_surcharge END, 2) AS DECIMAL(20, 2)) AS xiaoshougfhj " &
            "FROM xipunum_erp_outbound AS a " &
            "INNER JOIN xipunum_erp_outbound_order as b ON b.id=a.order_id " &
            "INNER JOIN xipunum_erp_shop AS e ON e.poduct_code = a.poduct_code " & searchMaterialFilter &
            " INNER JOIN xipunum_erp_store AS c ON c.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = e.item_number AND e.item_number != '' " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = a.kufang " &
            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id AND e.item_number IS NOT NULL AND e.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = e.specification_id AND e.specification_id IS NOT NULL AND e.specification_id != '' " &
            "LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) AND COALESCE(e1.category_id, e2.category_id) IS NOT NULL " &
            "WHERE a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "' " &
            shopFilter & searchPilingFilter & searchJinzhongFilter & " ORDER BY a.id DESC"
    End Function

    ' ========== 构建明细SQL ==========
    Private Function BuildDetailSQL(shopId As String, catFilter As String, Optional orderBy As String = " ORDER BY a.id DESC") As String
        Return "SELECT LEFT(a.creationtime, 10) as acreationtime, " &
            "CASE WHEN a.kufang = '0' THEN '0' ELSE a.kufang END AS kufangid, " &
            "CASE WHEN COALESCE(f.id, '') = '' THEN '0' ELSE f.id END AS pinleiid, " &
            "COALESCE(e1.category_id, e2.category_id, '0') AS guigeid, " &
            "CASE WHEN a.kufang = '0' THEN '总库' ELSE h.title END AS kufang, " &
            "b.settlement_number as fsettlement_number, a.poduct_code as outbound_product_code, " &
            "e.product_name as shop_product_name, e.item_number as shop_item_number, " &
            "CASE WHEN COALESCE(f.title, '') = '' THEN '未匹配' ELSE f.title END AS pinlei, " &
            "COALESCE(e1.title, e2.title, '未匹配') AS specification_title, " &
            "e.caizhi as material, e.quandu as quandu, c.factory_condition as company_condition, " &
            "e.single as danjian, a.net_weight as jinzhong, " &
            "CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN a.net_weight ELSE CAST(ROUND(e.weight / e.quantity * a.quantity, 3) AS DECIMAL(10, 3)) END AS zhongliang, " &
            "e.sales_unit as danwei, " &
            "CASE WHEN a.sales_return = '0' THEN c.basic_cost ELSE -c.basic_cost END AS chengbengf, " &
            "CASE WHEN a.sales_return = '0' THEN c.premium_cost ELSE -c.premium_cost END AS cankaogf, " &
            "CASE WHEN a.sales_return = '0' THEN c.company_surcharge ELSE -c.company_surcharge END AS chengbenfj, " &
            "CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN CAST(ROUND(c.cost_price / e.jin_zhong * a.net_weight, 2) AS DECIMAL(30, 2)) " &
            "ELSE CAST(ROUND(c.cost_price * a.quantity, 2) AS DECIMAL(30, 2)) END AS chengben, " &
            "a.xiaodan_amount as xiaodan_amount, a.xiao_amount as xiao_amount, a.quantity as xsshuliang, " &
            "CASE WHEN a.sales_return = '0' THEN c.sales_surcharge ELSE -c.sales_surcharge END AS yxiaoshoufj, " &
            "a.gold_price AS kejia, a.sales_cost AS xiaoshougf, a.sales_surcharge AS xiaoshoufj, " &
            "a.zhekou AS zhekou, a.settlement AS shishou, a.pling AS piling, " &
            "CASE WHEN a.sales_return = '0' THEN '客销' ELSE '客退' END AS sales_return, " &
            "CAST(ROUND(CASE WHEN a.sales_return = 1 THEN -a.sales_cost * a.net_weight + a.sales_surcharge ELSE a.sales_cost * a.net_weight + a.sales_surcharge END - " &
            "CASE WHEN a.sales_return = 1 THEN c.basic_cost * a.net_weight - c.company_surcharge ELSE c.basic_cost * a.net_weight + c.company_surcharge END, 2) AS DECIMAL(30, 2)) AS gongfeilr, " &
            "CAST(ROUND(CASE WHEN a.sales_return = 1 THEN c.basic_cost * a.net_weight - c.company_surcharge ELSE c.basic_cost * a.net_weight + c.company_surcharge END, 2) AS DECIMAL(20, 2)) AS chegbengfhj, " &
            "CAST(ROUND(CASE WHEN a.sales_return = 1 THEN -a.sales_cost * a.net_weight + a.sales_surcharge ELSE a.sales_cost * a.net_weight + a.sales_surcharge END, 2) AS DECIMAL(20, 2)) AS xiaoshougfhj, " &
            "x.gold_price as gold_price " &
            "FROM xipunum_erp_outbound AS a " &
            "INNER JOIN xipunum_erp_outbound_order as b ON b.id=a.order_id " &
            "INNER JOIN xipunum_erp_shop AS e ON e.poduct_code = a.poduct_code " & searchMaterialFilter &
            " INNER JOIN xipunum_erp_store AS c ON c.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = e.item_number AND e.item_number != '' " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = a.kufang " &
            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id AND e.item_number IS NOT NULL AND e.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = e.specification_id AND e.specification_id IS NOT NULL AND e.specification_id != '' " &
            "LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) AND COALESCE(e1.category_id, e2.category_id) IS NOT NULL " &
            "INNER JOIN xipunum_erp_store_order AS x ON x.id = c.order_id " &
            "WHERE a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "' " &
            "AND a.kufang='" & shopId & "'" & searchPilingFilter & searchJinzhongFilter & catFilter & orderBy
    End Function

    ' ========== 汇总外层字段 ==========
    Private ReadOnly SummaryOuterFields As String =
        "CAST(ROUND(sum(sum.jinzhong), 3) AS DECIMAL(20, 3)) AS jinzhong, " &
        "CAST(ROUND(sum(sum.xiao_amount), 2) AS DECIMAL(20, 2)) AS xiao_amount, " &
        "CAST(ROUND(sum(sum.zhongliang), 3) AS DECIMAL(20, 3)) AS zhongliang, " &
        "CAST(ROUND(sum(sum.xsshuliang), 2) AS DECIMAL(20, 2)) AS xsshuliang, " &
        "CAST(ROUND(sum(sum.shishou), 2) AS DECIMAL(20, 2)) AS shishou, " &
        "CAST(ROUND(sum(sum.chengben), 2) AS DECIMAL(20, 2)) AS chengben, " &
        "CAST(ROUND(sum(sum.gongfeilr), 2) AS DECIMAL(20, 2)) AS gongfeilr, " &
        "CAST(ROUND(sum(sum.chegbengfhj), 2) AS DECIMAL(20, 2)) AS chegbengfhj, " &
        "CAST(ROUND(sum(sum.xiaoshougfhj), 2) AS DECIMAL(20, 2)) AS xiaoshougfhj"

    ' ========== 品类模式查询 ==========
    Private Sub QueryDataCategory()
        dgvReport.Rows.Clear()
        Dim innerSQL As String = BuildInnerSQL()
        Dim rowIndex As Integer = 0

        ' 店铺级汇总
        Dim shopSQL As String = "SELECT sum.kufangid AS kufangid, sum.kufang AS kufang, " & SummaryOuterFields &
            " FROM (" & innerSQL & ") AS sum GROUP BY sum.kufangid ORDER BY sum.kufangid asc"
        Dim shopDT As DataTable = ExecuteQuery(shopSQL, MySQL_ReadReport)

        For Each shopRow As DataRow In shopDT.Rows
            Dim shopId As String = SafeString(shopRow("kufangid"))
            Dim shopName As String = SafeString(shopRow("kufang"))
            Dim shopJinzhong As String = SafeString(shopRow("jinzhong"))
            Dim shopZhongliang As String = SafeString(shopRow("zhongliang"))
            Dim shopXsshuliang As String = SafeString(shopRow("xsshuliang"))
            Dim shopXiaoAmount As String = SafeString(shopRow("xiao_amount"))
            Dim shopShishou As String = SafeString(shopRow("shishou"))
            Dim shopChengben As String = SafeString(shopRow("chengben"))
            Dim shopGongfeilr As String = SafeString(shopRow("gongfeilr"))
            Dim shopChegbengfhj As String = SafeString(shopRow("chegbengfhj"))
            Dim shopXiaoshougfhj As String = SafeString(shopRow("xiaoshougfhj"))

            ' 品类级汇总
            Dim catSQL As String = "SELECT sum.kufangid AS kufangid, sum.kufang AS kufang, sum.pinleiid AS pinleiid, sum.pinlei AS pinlei, " & SummaryOuterFields &
                " FROM (" & innerSQL & ") AS sum WHERE sum.kufangid='" & shopId & "' GROUP BY sum.kufangid, sum.pinleiid ORDER BY sum.kufangid asc"
            Dim catDT As DataTable = ExecuteQuery(catSQL, MySQL_ReadReport)

            For Each catRow As DataRow In catDT.Rows
                Dim catId As String = SafeString(catRow("pinleiid"))
                Dim catName As String = SafeString(catRow("pinlei"))
                Dim catJinzhong As String = SafeString(catRow("jinzhong"))
                Dim catZhongliang As String = SafeString(catRow("zhongliang"))
                Dim catXsshuliang As String = SafeString(catRow("xsshuliang"))
                Dim catXiaoAmount As String = SafeString(catRow("xiao_amount"))
                Dim catShishou As String = SafeString(catRow("shishou"))
                Dim catChengben As String = SafeString(catRow("chengben"))
                Dim catGongfeilr As String = SafeString(catRow("gongfeilr"))
                Dim catChegbengfhj As String = SafeString(catRow("chegbengfhj"))
                Dim catXiaoshougfhj As String = SafeString(catRow("xiaoshougfhj"))

                ' 明细数据
                Dim detailSQL As String = BuildDetailSQL(shopId, " and COALESCE(e1.category_id, e2.category_id, '0')='" & catId & "'")
                Dim detailDT As DataTable = ExecuteQuery(detailSQL, MySQL_ReadReport)

                Dim detailIndex As Integer = 0
                For Each detailRow As DataRow In detailDT.Rows
                    detailIndex += 1
                    rowIndex += 1
                    dgvReport.Rows.Add(rowIndex,
                        SafeString(detailRow("fsettlement_number")), SafeString(detailRow("outbound_product_code")),
                        SafeString(detailRow("acreationtime")), SafeString(detailRow("shop_product_name")),
                        SafeString(detailRow("shop_item_number")), SafeString(detailRow("kufang")),
                        SafeString(detailRow("pinlei")), SafeString(detailRow("specification_title")),
                        SafeString(detailRow("material")), SafeString(detailRow("quandu")),
                        SafeString(detailRow("company_condition")), SafeString(detailRow("danjian")),
                        SafeString(detailRow("jinzhong")), SafeString(detailRow("zhongliang")),
                        SafeString(detailRow("danwei")), SafeString(detailRow("gold_price")),
                        SafeString(detailRow("chengbengf")), SafeString(detailRow("cankaogf")),
                        SafeString(detailRow("chengbenfj")), SafeString(detailRow("chengben")),
                        SafeString(detailRow("xiaodan_amount")), SafeString(detailRow("xiao_amount")),
                        SafeString(detailRow("xsshuliang")), SafeString(detailRow("yxiaoshoufj")),
                        SafeString(detailRow("kejia")), SafeString(detailRow("xiaoshougf")),
                        SafeString(detailRow("xiaoshoufj")), SafeString(detailRow("zhekou")),
                        SafeString(detailRow("shishou")), SafeString(detailRow("gongfeilr")),
                        SafeString(detailRow("chegbengfhj")), SafeString(detailRow("xiaoshougfhj")),
                        SafeString(detailRow("piling")), SafeString(detailRow("sales_return")))
                Next

                ' 品类合计行
                rowIndex += 1
                Dim catTotalIdx As Integer = dgvReport.Rows.Add(rowIndex, "", "", "", "", "", "", "", "品类合计", "", "", "", "",
                    catJinzhong, catZhongliang, "", "", "", "", "", catChengben, "", catXiaoAmount, catXsshuliang,
                    "", "", "", "", "", catShishou, catGongfeilr, catChegbengfhj, catXiaoshougfhj, "", "")
                SetRowColor(catTotalIdx, BlueColor, 0, 30)
            Next

            ' 店铺合计行
            rowIndex += 1
            Dim shopTotalIdx As Integer = dgvReport.Rows.Add(rowIndex, "", "", "", "", "", shopName, "", "合计", "", "", "", "",
                shopJinzhong, shopZhongliang, "", "", "", "", "", shopChengben, "", shopXiaoAmount, shopXsshuliang,
                "", "", "", "", "", shopShishou, shopGongfeilr, shopChegbengfhj, shopXiaoshougfhj, "", "")
            SetRowColor(shopTotalIdx, RedColor, 0, 30)
        Next
    End Sub

    ' ========== 自定义模式查询 ==========
    Private Sub QueryDataCustom()
        dgvReport.Rows.Clear()
        Dim innerSQL As String = BuildInnerSQL()
        Dim rowIndex As Integer = 0

        ' 店铺级汇总（同品类模式）
        Dim shopSQL As String = "SELECT sum.kufangid AS kufangid, sum.kufang AS kufang, " & SummaryOuterFields &
            " FROM (" & innerSQL & ") AS sum GROUP BY sum.kufangid ORDER BY sum.kufangid asc"
        Dim shopDT As DataTable = ExecuteQuery(shopSQL, MySQL_ReadReport)

        ' 加载自定义品类配置
        Dim configSQL As String = "SELECT title, category_list FROM xipunum_erp_category_stat_config WHERE id in ('2','4','9','13') ORDER BY id asc"
        Dim configDT As DataTable = ExecuteQuery(configSQL, MySQL_ReadReport)
        Dim configList As New List(Of (title As String, catIds As String))()
        For Each configRow As DataRow In configDT.Rows
            Dim title As String = SafeString(configRow("title"))
            Dim catList As String = SafeString(configRow("category_list"))
            If catList <> "" Then
                catList = catList.Replace("[", "'").Replace("]", "'").Replace(",", "','")
            End If
            configList.Add((title, catList))
        Next

        For Each shopRow As DataRow In shopDT.Rows
            Dim shopId As String = SafeString(shopRow("kufangid"))
            Dim shopName As String = SafeString(shopRow("kufang"))

            ' 合并所有自定义品类ID
            Dim allCatIds As String = ""
            For Each config In configList
                If config.catIds <> "" Then
                    If allCatIds <> "" Then allCatIds &= ","
                    allCatIds &= config.catIds
                End If
            Next
            If allCatIds = "" Then allCatIds = "''"

            ' 店铺级自定义汇总
            Dim shopCustomSQL As String = "SELECT " & SummaryOuterFields &
                " FROM (" & innerSQL & ") AS sum WHERE sum.kufangid='" & shopId & "' AND sum.pinleiid in (" & allCatIds & ")"
            Dim shopCustomDT As DataTable = ExecuteQuery(shopCustomSQL, MySQL_ReadReport)
            Dim shopJinzhong As String = "0"
            Dim shopZhongliang As String = "0"
            Dim shopXsshuliang As String = "0"
            Dim shopXiaoAmount As String = "0"
            Dim shopShishou As String = "0"
            Dim shopChengben As String = "0"
            Dim shopGongfeilr As String = "0"
            Dim shopChegbengfhj As String = "0"
            Dim shopXiaoshougfhj As String = "0"
            If shopCustomDT.Rows.Count > 0 Then
                shopJinzhong = SafeString(shopCustomDT.Rows(0)("jinzhong"))
                shopZhongliang = SafeString(shopCustomDT.Rows(0)("zhongliang"))
                shopXsshuliang = SafeString(shopCustomDT.Rows(0)("xsshuliang"))
                shopXiaoAmount = SafeString(shopCustomDT.Rows(0)("xiao_amount"))
                shopShishou = SafeString(shopCustomDT.Rows(0)("shishou"))
                shopChengben = SafeString(shopCustomDT.Rows(0)("chengben"))
                shopGongfeilr = SafeString(shopCustomDT.Rows(0)("gongfeilr"))
                shopChegbengfhj = SafeString(shopCustomDT.Rows(0)("chegbengfhj"))
                shopXiaoshougfhj = SafeString(shopCustomDT.Rows(0)("xiaoshougfhj"))
            End If

            ' 每个自定义品类
            For Each config In configList
                If config.catIds = "" Then Continue For

                ' 品类级汇总
                Dim catSQL As String = "SELECT " & SummaryOuterFields &
                    " FROM (" & innerSQL & ") AS sum WHERE sum.kufangid='" & shopId & "' AND sum.pinleiid in (" & config.catIds & ")"
                Dim catDT As DataTable = ExecuteQuery(catSQL, MySQL_ReadReport)
                Dim catJinzhong As String = "0"
                Dim catZhongliang As String = "0"
                Dim catXsshuliang As String = "0"
                Dim catXiaoAmount As String = "0"
                Dim catShishou As String = "0"
                Dim catChengben As String = "0"
                Dim catGongfeilr As String = "0"
                Dim catChegbengfhj As String = "0"
                Dim catXiaoshougfhj As String = "0"
                If catDT.Rows.Count > 0 Then
                    catJinzhong = SafeString(catDT.Rows(0)("jinzhong"))
                    catZhongliang = SafeString(catDT.Rows(0)("zhongliang"))
                    catXsshuliang = SafeString(catDT.Rows(0)("xsshuliang"))
                    catXiaoAmount = SafeString(catDT.Rows(0)("xiao_amount"))
                    catShishou = SafeString(catDT.Rows(0)("shishou"))
                    catChengben = SafeString(catDT.Rows(0)("chengben"))
                    catGongfeilr = SafeString(catDT.Rows(0)("gongfeilr"))
                    catChegbengfhj = SafeString(catDT.Rows(0)("chegbengfhj"))
                    catXiaoshougfhj = SafeString(catDT.Rows(0)("xiaoshougfhj"))
                End If

                ' 明细数据
                Dim detailSQL As String = BuildDetailSQL(shopId, " and COALESCE(e1.category_id, e2.category_id, '0') in (" & config.catIds & ")", " ORDER BY COALESCE(e1.category_id, e2.category_id, '0') asc")
                Dim detailDT As DataTable = ExecuteQuery(detailSQL, MySQL_ReadReport)

                Dim detailIndex As Integer = 0
                For Each detailRow As DataRow In detailDT.Rows
                    detailIndex += 1
                    rowIndex += 1
                    dgvReport.Rows.Add(detailIndex,
                        SafeString(detailRow("fsettlement_number")), SafeString(detailRow("outbound_product_code")),
                        SafeString(detailRow("acreationtime")), SafeString(detailRow("shop_product_name")),
                        SafeString(detailRow("shop_item_number")), SafeString(detailRow("kufang")),
                        SafeString(detailRow("pinlei")), SafeString(detailRow("specification_title")),
                        SafeString(detailRow("material")), SafeString(detailRow("quandu")),
                        SafeString(detailRow("company_condition")), SafeString(detailRow("danjian")),
                        SafeString(detailRow("jinzhong")), SafeString(detailRow("zhongliang")),
                        SafeString(detailRow("danwei")), SafeString(detailRow("gold_price")),
                        SafeString(detailRow("chengbengf")), SafeString(detailRow("cankaogf")),
                        SafeString(detailRow("chengbenfj")), SafeString(detailRow("chengben")),
                        SafeString(detailRow("xiaodan_amount")), SafeString(detailRow("xiao_amount")),
                        SafeString(detailRow("xsshuliang")), SafeString(detailRow("yxiaoshoufj")),
                        SafeString(detailRow("kejia")), SafeString(detailRow("xiaoshougf")),
                        SafeString(detailRow("xiaoshoufj")), SafeString(detailRow("zhekou")),
                        SafeString(detailRow("shishou")), SafeString(detailRow("gongfeilr")),
                        SafeString(detailRow("chegbengfhj")), SafeString(detailRow("xiaoshougfhj")),
                        SafeString(detailRow("piling")), SafeString(detailRow("sales_return")))
                Next

                ' 自定义品类合计行（仅当有数据时）
                If catJinzhong <> "" Or catZhongliang <> "" Or catChengben <> "" Or catXiaoAmount <> "" Or
                   catXsshuliang <> "" Or catShishou <> "" Or catGongfeilr <> "" Or catChegbengfhj <> "" Or catXiaoshougfhj <> "" Then
                    rowIndex += 1
                    dgvReport.Rows.Add("", "", "", "", "", "", "", "", config.title & "合计", "", "", "", "",
                        catJinzhong, catZhongliang, "", "", "", "", "", catChengben, "", catXiaoAmount, catXsshuliang,
                        "", "", "", "", "", catShishou, catGongfeilr, catChegbengfhj, catXiaoshougfhj, "", "")
                End If
            Next

            ' 店铺合计行
            rowIndex += 1
            Dim shopTotalIdx As Integer = dgvReport.Rows.Add("", "", "", "", "", "", shopName, "", "合计", "", "", "", "",
                shopJinzhong, shopZhongliang, "", "", "", "", "", shopChengben, "", shopXiaoAmount, shopXsshuliang,
                "", "", "", "", "", shopShishou, shopGongfeilr, shopChegbengfhj, shopXiaoshougfhj, "", "")
            SetRowColor(shopTotalIdx, RedColor, 0, 30)
        Next
    End Sub

    ' ========== 设置行颜色 ==========
    Private Sub SetRowColor(rowIndex As Integer, color As Drawing.Color, startCol As Integer, endCol As Integer)
        For col As Integer = startCol To Math.Min(endCol, dgvReport.Columns.Count - 1)
            dgvReport.Rows(rowIndex).Cells(col).Style.ForeColor = color
        Next
    End Sub

    ' ========== 右键菜单 - 复制单元格 ==========
    Private Sub dgvReport_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs)
        If e.Button = MouseButtons.Right AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim cellValue As String = SafeString(dgvReport.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
            If cellValue <> "" Then
                dgvReport.Rows(e.RowIndex).Selected = True
                contextMenuReport.Show(dgvReport, e.Location)
            End If
        End If
    End Sub

    Private Sub menuItemCopy_Click(sender As Object, e As EventArgs)
        If dgvReport.CurrentCell IsNot Nothing Then
            Clipboard.SetText(SafeString(dgvReport.CurrentCell.Value))
            ShowInfo("复制:" & SafeString(dgvReport.CurrentCell.Value) & " 到剪切板成功！")
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
            ExportToExcel(dt, "销售详情报表表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 重置按钮 ==========
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        chkJinzhong.Checked = False
        txtJinzhong1.Text = ""
        txtJinzhong2.Text = ""
        radioAll.Checked = True
        radioCategory.Checked = True
        LoadShopList()
        LoadMaterialList()
        dgvReport.Rows.Clear()
        btnQuery_Click(Nothing, Nothing)
    End Sub

End Class

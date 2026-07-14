' ============================================================================
' 销售查询报表窗口
' 功能: 三级层次化销售数据统计（店铺→品类→规格），含回收对比、黄金应结、优惠汇总
' 对应易语言: 窗口程序集_窗口_销售查询报表
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SalesQueryReportForm
    Inherits System.Windows.Forms.Form

    ' ========== 查找条件变量 ==========
    Private searchStartDate As String = ""
    Private searchEndDate As String = ""
    Private searchShopFilter As String = ""
    Private searchMaterialFilter As String = ""
    Private materialTotalCount As Integer = 0

    ' ========== 控件声明 ==========
    Private dgvReport As New DataGridView()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private chkShop As New CheckedListBox()
    Private chkMaterial As New CheckedListBox()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnReset As New Button()
    Private chkShopAll As New CheckBox()
    Private chkMaterialAll As New CheckBox()

    ' ========== 红色字体 ==========
    Private ReadOnly RedColor As Drawing.Color = Drawing.ColorTranslator.FromHtml("#E0422F")

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    ' ========== UI初始化 ==========
    Private Sub InitializeUI()
        Me.Text = "销售查询报表"
        Me.Size = New Drawing.Size(1400, 800)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' ========== 顶部筛选面板 ==========
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
        btnQuery.Location = New Drawing.Point(820, 9)
        btnQuery.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnQuery)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(910, 9)
        btnExport.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnExport)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(1000, 9)
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
        AddHandler chkMaterialAll.CheckedChanged, AddressOf chkMaterialAll_CheckedChanged

        ' 初始化表头
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
        Dim headers() As String = {"序号", "店铺名称", "品类", "规格", "实际数量", "实际金重", "实际总重", "实际金额", "", "销售数量", "销售金重", "销售总重", "销售金额", "实销金额", "", "客退数量", "客退金重", "客退总重", "客退金额", "实退金额"}
        Dim widths() As Integer = {45, 80, 75, 75, -1, 100, 100, 115, 5, -1, 100, 100, 115, 115, 5, -1, -1, -1, -1, -1}

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

    ' ========== 窗体加载 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        Me.Text = "销售查询报表 时间:" & dtpStart.Value.ToString("yyyy-MM-dd") & "至" & dtpEnd.Value.ToString("yyyy-MM-dd")
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
        materialTotalCount = chkMaterial.Items.Count
        For i As Integer = 0 To chkMaterial.Items.Count - 1
            chkMaterial.SetItemChecked(i, True)
        Next
    End Sub

    ' ========== 全选/取消全选 ==========
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

    ' ========== 查询按钮 ==========
    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        btnQuery.Enabled = False
        btnReset.Enabled = False
        btnExport.Enabled = False
        Me.Text = "销售查询报表 时间:" & dtpStart.Value.ToString("yyyy-MM-dd") & "至" & dtpEnd.Value.ToString("yyyy-MM-dd")

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

        ' 日期
        searchStartDate = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
        searchEndDate = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"

        ' 执行查询
        QueryData()

        btnQuery.Enabled = True
        btnReset.Enabled = True
        btnExport.Enabled = True
    End Sub

    ' ========== 构建内层SQL ==========
    Private Function BuildInnerSQL() As String
        Dim sql As String = "SELECT CASE WHEN a.kufang = '0' THEN '0' ELSE a.kufang END AS kufangid, " &
            "CASE WHEN COALESCE(f.id, '') = '' THEN '0' ELSE f.id END AS pinleiid, " &
            "COALESCE(e1.id, e2.id, '0') AS guigeid, " &
            "CASE WHEN a.kufang = '0' THEN '总库' ELSE h.title END AS kufang, " &
            "CASE WHEN COALESCE(f.title, '') = '' THEN '未匹配' ELSE f.title END AS pinlei, " &
            "COALESCE(e1.title, e2.title, '未匹配') AS guige, " &
            "CAST(ROUND(a.quantity, 2) AS DECIMAL(30, 2)) AS sjquantity, " &
            "CAST(ROUND(a.net_weight, 3) AS DECIMAL(30, 3)) AS sjjinzhong, " &
            "CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN CAST(ROUND(a.net_weight, 3) AS DECIMAL(20, 3)) " &
            "ELSE CAST(ROUND(e.weight / e.quantity * a.quantity, 3) AS DECIMAL(30, 3)) END AS sjzongzhong, " &
            "CAST(ROUND(a.settlement, 2) AS DECIMAL(30, 2)) AS sjjine, " &
            "CASE WHEN sales_return = 0 THEN CAST(ROUND(a.quantity, 2) AS DECIMAL(30, 2)) ELSE '0.00' END AS xsquantity, " &
            "CASE WHEN sales_return = 0 THEN CAST(ROUND(a.net_weight, 3) AS DECIMAL(30, 3)) ELSE '0.000' END AS xsjinzhong, " &
            "CASE WHEN sales_return = 0 THEN CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN CAST(ROUND(a.net_weight, 3) AS DECIMAL(20, 3)) " &
            "ELSE CAST(ROUND(e.weight / e.quantity * a.quantity, 3) AS DECIMAL(30, 3)) END " &
            "ELSE CAST(0.000 AS DECIMAL(20, 3)) END AS xszongzhong, " &
            "CASE WHEN sales_return = 0 THEN CAST(ROUND(a.xiao_amount, 2) AS DECIMAL(30, 2)) ELSE '0.00' END AS xsjine, " &
            "CASE WHEN sales_return = 0 THEN CAST(ROUND(a.settlement, 2) AS DECIMAL(30, 2)) ELSE '0.00' END AS sxjine, " &
            "CASE WHEN sales_return = 1 THEN CAST(ROUND(a.quantity, 2) AS DECIMAL(30, 2)) ELSE '0.00' END AS ktquantity, " &
            "CASE WHEN sales_return = 1 THEN CAST(ROUND(a.net_weight, 3) AS DECIMAL(30, 3)) ELSE '0.000' END AS ktjinzhong, " &
            "CASE WHEN sales_return = 1 THEN CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN CAST(ROUND(a.net_weight, 3) AS DECIMAL(20, 3)) " &
            "ELSE CAST(ROUND(e.weight / e.quantity * a.quantity, 3) AS DECIMAL(30, 3)) END " &
            "ELSE CAST(0.000 AS DECIMAL(20, 3)) END AS ktzongzhong, " &
            "CASE WHEN sales_return = 1 THEN CAST(ROUND(a.xiao_amount, 2) AS DECIMAL(30, 2)) ELSE '0.00' END AS ktjine, " &
            "CASE WHEN sales_return = 1 THEN CAST(ROUND(a.settlement, 2) AS DECIMAL(30, 2)) ELSE '0.00' END AS stjine " &
            "FROM xipunum_erp_outbound AS a " &
            "INNER JOIN xipunum_erp_shop AS e ON e.poduct_code = a.poduct_code " & searchMaterialFilter &
            " LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = e.item_number AND e.item_number != '' " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = a.kufang " &
            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id AND e.item_number IS NOT NULL AND e.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = e.specification_id AND e.specification_id IS NOT NULL AND e.specification_id != '' " &
            "LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) AND COALESCE(e1.category_id, e2.category_id) IS NOT NULL " &
            "WHERE a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "'" & searchShopFilter &
            " ORDER BY a.id DESC"
        Return sql
    End Function

    ' ========== 外层聚合字段 ==========
    Private ReadOnly OuterFields As String =
        "CAST(ROUND(sum(sum.sjquantity), 2) AS DECIMAL(20, 2)) AS sjquantity, " &
        "CAST(ROUND(sum(sum.sjjinzhong), 3) AS DECIMAL(20, 3)) AS sjjinzhong, " &
        "CAST(ROUND(sum(sum.sjzongzhong), 3) AS DECIMAL(20, 3)) AS sjzongzhong, " &
        "CAST(ROUND(sum(sum.sjjine), 2) AS DECIMAL(20, 2)) AS sjjine, " &
        "CAST(ROUND(sum(sum.xsquantity), 2) AS DECIMAL(20, 2)) AS xsquantity, " &
        "CAST(ROUND(sum(sum.xsjinzhong), 3) AS DECIMAL(20, 3)) AS xsjinzhong, " &
        "CAST(ROUND(sum(sum.xszongzhong), 3) AS DECIMAL(20, 3)) AS xszongzhong, " &
        "CAST(ROUND(sum(sum.xsjine), 2) AS DECIMAL(20, 2)) AS xsjine, " &
        "CAST(ROUND(sum(sum.sxjine), 2) AS DECIMAL(20, 2)) AS sxjine, " &
        "CAST(ROUND(sum(sum.ktquantity), 2) AS DECIMAL(20, 2)) AS ktquantity, " &
        "CAST(ROUND(sum(sum.ktjinzhong), 3) AS DECIMAL(20, 3)) AS ktjinzhong, " &
        "CAST(ROUND(sum(sum.ktzongzhong), 3) AS DECIMAL(20, 3)) AS ktzongzhong, " &
        "CAST(ROUND(sum(sum.ktjine), 2) AS DECIMAL(20, 2)) AS ktjine, " &
        "CAST(ROUND(sum(sum.stjine), 2) AS DECIMAL(20, 2)) AS stjine"

    ' ========== 主查询：三级层次化数据 ==========
    Private Sub QueryData()
        dgvReport.Rows.Clear()
        Dim innerSQL As String = BuildInnerSQL()
        Dim rowIndex As Integer = 0

        ' === 第一级：店铺汇总 ===
        Dim shopSQL As String = "SELECT sum.kufangid as kufangid, sum.kufang as kufang, " & OuterFields &
            " FROM (" & innerSQL & ") AS sum GROUP BY sum.kufangid ORDER BY CAST(sum.kufangid AS UNSIGNED) asc"
        Dim shopDT As DataTable = ExecuteQuery(shopSQL, MySQL_ReadReport)

        For Each shopRow As DataRow In shopDT.Rows
            Dim shopId As String = SafeString(shopRow("kufangid"))
            Dim shopName As String = SafeString(shopRow("kufang"))
            Dim shopSJQ As String = SafeString(shopRow("sjquantity"))
            Dim shopSJJ As String = SafeString(shopRow("sjjinzhong"))
            Dim shopSJZ As String = SafeString(shopRow("sjzongzhong"))
            Dim shopSJJE As String = SafeString(shopRow("sjjine"))
            Dim shopXSQ As String = SafeString(shopRow("xsquantity"))
            Dim shopXSJ As String = SafeString(shopRow("xsjinzhong"))
            Dim shopXSZ As String = SafeString(shopRow("xszongzhong"))
            Dim shopXSJE As String = SafeString(shopRow("xsjine"))
            Dim shopSXJE As String = SafeString(shopRow("sxjine"))
            Dim shopKTQ As String = SafeString(shopRow("ktquantity"))
            Dim shopKTJ As String = SafeString(shopRow("ktjinzhong"))
            Dim shopKTZ As String = SafeString(shopRow("ktzongzhong"))
            Dim shopKTJE As String = SafeString(shopRow("ktjine"))
            Dim shopSTJE As String = SafeString(shopRow("stjine"))

            ' === 第二级：品类汇总 ===
            Dim catSQL As String = "SELECT sum.pinleiid as pinleiid, sum.pinlei as pinlei, " & OuterFields &
                " FROM (" & innerSQL & ") AS sum WHERE sum.kufangid='" & shopId & "' GROUP BY sum.kufangid, sum.pinleiid ORDER BY sum.pinleiid asc"
            Dim catDT As DataTable = ExecuteQuery(catSQL, MySQL_ReadReport)

            For Each catRow As DataRow In catDT.Rows
                Dim catId As String = SafeString(catRow("pinleiid"))
                Dim catName As String = SafeString(catRow("pinlei"))
                Dim catSJQ As String = SafeString(catRow("sjquantity"))
                Dim catSJJ As String = SafeString(catRow("sjjinzhong"))
                Dim catSJZ As String = SafeString(catRow("sjzongzhong"))
                Dim catSJJE As String = SafeString(catRow("sjjine"))
                Dim catXSQ As String = SafeString(catRow("xsquantity"))
                Dim catXSJ As String = SafeString(catRow("xsjinzhong"))
                Dim catXSZ As String = SafeString(catRow("xszongzhong"))
                Dim catXSJE As String = SafeString(catRow("xsjine"))
                Dim catSXJE As String = SafeString(catRow("sxjine"))
                Dim catKTQ As String = SafeString(catRow("ktquantity"))
                Dim catKTJ As String = SafeString(catRow("ktjinzhong"))
                Dim catKTZ As String = SafeString(catRow("ktzongzhong"))
                Dim catKTJE As String = SafeString(catRow("ktjine"))
                Dim catSTJE As String = SafeString(catRow("stjine"))

                ' === 第三级：规格明细 ===
                Dim specSQL As String = "SELECT sum.guigeid as guigeid, sum.guige as guige, " & OuterFields &
                    " FROM (" & innerSQL & ") AS sum WHERE sum.kufangid='" & shopId & "' and sum.pinleiid='" & catId & "' GROUP BY sum.kufangid, sum.pinleiid, sum.guigeid ORDER BY sum.pinleiid asc"
                Dim specDT As DataTable = ExecuteQuery(specSQL, MySQL_ReadReport)

                Dim specIndex As Integer = 0
                For Each specRow As DataRow In specDT.Rows
                    specIndex += 1
                    Dim specName As String = SafeString(specRow("guige"))
                    rowIndex += 1
                    dgvReport.Rows.Add(specIndex, "", "", specName,
                        SafeString(specRow("sjquantity")), SafeString(specRow("sjjinzhong")), SafeString(specRow("sjzongzhong")), SafeString(specRow("sjjine")), "",
                        SafeString(specRow("xsquantity")), SafeString(specRow("xsjinzhong")), SafeString(specRow("xszongzhong")), SafeString(specRow("xsjine")), SafeString(specRow("sxjine")), "",
                        SafeString(specRow("ktquantity")), SafeString(specRow("ktjinzhong")), SafeString(specRow("ktzongzhong")), SafeString(specRow("ktjine")), SafeString(specRow("stjine")))
                Next

                ' 品类合计行
                rowIndex += 1
                dgvReport.Rows.Add("", "", catName, "品类合计",
                    catSJQ, catSJJ, catSJZ, catSJJE, "",
                    catXSQ, catXSJ, catXSZ, catXSJE, catSXJE, "",
                    catKTQ, catKTJ, catKTZ, catKTJE, catSTJE)
            Next

            ' 店铺合计行
            rowIndex += 1
            Dim shopTotalRowIdx As Integer = dgvReport.Rows.Add("", "店铺合计", "", "",
                shopSJQ, shopSJJ, shopSJZ, shopSJJE, "",
                shopXSQ, shopXSJ, shopXSZ, shopXSJE, shopSXJE, "",
                shopKTQ, shopKTJ, shopKTZ, shopKTJE, shopSTJE)
            SetRowRedFont(shopTotalRowIdx)

            ' === 回收分类明细 ===
            Dim recoveryCatSQL As String = "SELECT c.title AS product_name, " &
                "CAST(ROUND(sum(a.jin_zhong), 3) AS DECIMAL(30, 3)) AS jin_zhong, " &
                "CAST(ROUND(sum(a.total), 3) AS DECIMAL(30, 3)) AS total, " &
                "CAST(ROUND(sum(a.retreat_amount), 2) AS DECIMAL(30, 2)) AS retreat_amount " &
                "FROM xipunum_erp_retreat AS a " &
                "INNER JOIN xipunum_erp_retreat_title AS b ON b.id = a.product_name " &
                "INNER JOIN xipunum_erp_category AS c ON c.id = b.category_id " &
                "WHERE a.shopping_guide IN (SELECT USER FROM xipunum_erp_user WHERE department = '" & shopId & "' GROUP BY USER) " &
                "AND a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "' GROUP BY c.id"
            Dim recoveryCatDT As DataTable = ExecuteQuery(recoveryCatSQL, MySQL_ReadReport)

            Dim recoveryIndex As Integer = 0
            For Each recRow As DataRow In recoveryCatDT.Rows
                recoveryIndex += 1
                rowIndex += 1
                dgvReport.Rows.Add(recoveryIndex, "", "", SafeString(recRow("product_name")), "",
                    SafeString(recRow("jin_zhong")), SafeString(recRow("total")), SafeString(recRow("retreat_amount")), "", "", "", "", "", "", "", "", "", "", "", "")
            Next

            ' 回收合计
            Dim recoveryTotalSQL As String = "SELECT CAST(ROUND(sum(jin_zhong), 3) AS DECIMAL(30, 3)) as jin_zhong, " &
                "CAST(ROUND(sum(total), 3) AS DECIMAL(30, 3)) as total, " &
                "CAST(ROUND(sum(retreat_amount), 2) AS DECIMAL(30, 2)) as retreat_amount " &
                "FROM xipunum_erp_retreat WHERE cjuser in (SELECT user FROM xipunum_erp_user WHERE department='" & shopId & "' GROUP BY user) " &
                "and creationtime>='" & searchStartDate & "' and creationtime<='" & searchEndDate & "'"
            Dim recoveryTotalDT As DataTable = ExecuteQuery(recoveryTotalSQL, MySQL_ReadReport)
            Dim recoveryTotalJinzhong As String = "0"
            Dim recoveryTotalTotal As String = "0"
            Dim recoveryTotalAmount As String = "0"
            If recoveryTotalDT.Rows.Count > 0 Then
                recoveryTotalJinzhong = SafeString(recoveryTotalDT.Rows(0)("jin_zhong"))
                recoveryTotalTotal = SafeString(recoveryTotalDT.Rows(0)("total"))
                recoveryTotalAmount = SafeString(recoveryTotalDT.Rows(0)("retreat_amount"))
            End If

            rowIndex += 1
            dgvReport.Rows.Add("", "回收合计", "", "", "",
                recoveryTotalJinzhong, recoveryTotalTotal, recoveryTotalAmount, "", "", "", "", "", "", "", "", "", "", "", "")

            ' === 黄金汇总（销售侧）===
            Dim goldSQL As String = "SELECT sum.guigeid as guigeid, sum.guige as guige, " & OuterFields &
                " FROM (" & innerSQL & ") AS sum WHERE sum.kufangid='" & shopId & "' and sum.pinleiid in ('1','3','6','9') GROUP BY sum.kufangid ORDER BY sum.pinleiid asc"
            Dim goldDT As DataTable = ExecuteQuery(goldSQL, MySQL_ReadReport)
            Dim goldSJinzhong As String = "0"
            Dim goldSJine As String = "0"
            If goldDT.Rows.Count > 0 Then
                goldSJinzhong = SafeString(goldDT.Rows(0)("sjjinzhong"))
                goldSJine = SafeString(goldDT.Rows(0)("sjjine"))
            End If

            ' === 黄金回收 ===
            Dim goldRecoverySQL As String = "SELECT CAST(ROUND(sum(a.jin_zhong), 3) AS DECIMAL(30, 3)) AS jin_zhong, " &
                "CAST(ROUND(sum(a.retreat_amount), 2) AS DECIMAL(30, 2)) AS retreat_amount " &
                "FROM xipunum_erp_retreat AS a " &
                "INNER JOIN xipunum_erp_retreat_title AS b ON b.id = a.product_name and b.category_id in ('1','3','6','9') " &
                "INNER JOIN xipunum_erp_category AS c ON c.id = b.category_id " &
                "WHERE a.shopping_guide IN (SELECT USER FROM xipunum_erp_user WHERE department = '" & shopId & "' GROUP BY USER) " &
                "AND a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "'"
            Dim goldRecoveryDT As DataTable = ExecuteQuery(goldRecoverySQL, MySQL_ReadReport)
            Dim goldRecoveryJinzhong As String = "0"
            Dim goldRecoveryAmount As String = "0"
            If goldRecoveryDT.Rows.Count > 0 Then
                goldRecoveryJinzhong = SafeString(goldRecoveryDT.Rows(0)("jin_zhong"))
                goldRecoveryAmount = SafeString(goldRecoveryDT.Rows(0)("retreat_amount"))
            End If

            ' 黄金差异（黄金应结）
            Dim goldDiffJinzhong As String = FormatDecimal3(SafeDecimal(goldSJinzhong) - SafeDecimal(goldRecoveryJinzhong))
            Dim goldDiffAmount As String = FormatDecimal2(SafeDecimal(goldSJine) - SafeDecimal(goldRecoveryAmount))

            rowIndex += 1
            Dim goldRowIdx As Integer = dgvReport.Rows.Add("", "", "黄金应结", "合计", "",
                goldDiffJinzhong, "", goldDiffAmount, "", "", "", "", "", "", "", "", "", "", "", "")
            SetRowRedFont(goldRowIdx)

            ' === 店铺差异合计 ===
            Dim diffJinzhong As String = FormatDecimal3(SafeDecimal(shopSJJ) - SafeDecimal(recoveryTotalJinzhong))
            Dim diffAmount As String = FormatDecimal2(SafeDecimal(shopSJJE) - SafeDecimal(recoveryTotalAmount))

            rowIndex += 1
            Dim diffRowIdx As Integer = dgvReport.Rows.Add("", shopName, "", "合计", "",
                diffJinzhong, "", diffAmount, "", "", "", "", "", "", "", "", "", "", "", "")

            ' === 优惠总金额（仅所有材质都选中时显示）===
            Dim checkedMaterialCount As Integer = 0
            For i As Integer = 0 To chkMaterial.Items.Count - 1
                If chkMaterial.GetItemChecked(i) Then
                    checkedMaterialCount += 1
                End If
            Next
            If checkedMaterialCount = materialTotalCount Then
                Dim discountSQL As String = "SELECT CAST(ROUND(sum(tol.youhui), 2) AS DECIMAL(30, 2)) AS youhui " &
                    "FROM (SELECT CASE WHEN f.sales_return = 1 THEN -f.youhui ELSE f.youhui END AS youhui " &
                    "FROM xipunum_erp_outbound AS a " &
                    "INNER JOIN xipunum_erp_outbound_order AS f ON f.id = a.order_id " &
                    "WHERE a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "' " &
                    "and a.kufang='" & shopId & "' GROUP BY a.order_id) AS tol"
                Dim discountDT As DataTable = ExecuteQuery(discountSQL, MySQL_ReadReport)
                Dim discountAmount As String = "0.00"
                If discountDT.Rows.Count > 0 Then
                    discountAmount = SafeString(discountDT.Rows(0)("youhui"))
                End If
                dgvReport.Rows(diffRowIdx).Cells(11).Value = "优惠总金额"
                dgvReport.Rows(diffRowIdx).Cells(12).Value = discountAmount
            End If

            SetRowRedFont(diffRowIdx)
        Next
    End Sub

    ' ========== 设置行红色字体 ==========
    Private Sub SetRowRedFont(rowIndex As Integer)
        For col As Integer = 0 To dgvReport.Columns.Count - 1
            dgvReport.Rows(rowIndex).Cells(col).Style.ForeColor = RedColor
        Next
    End Sub

    ' ========== 三位小数格式化 ==========
    Private Function FormatDecimal3(value As Decimal) As String
        Dim rounded As Decimal = Math.Round(value, 3)
        Dim text As String = rounded.ToString()
        If text = "" Then Return "0.000"
        If text.Contains(".") Then
            Dim parts() As String = text.Split("."c)
            If parts(1).Length = 1 Then
                text = text & "00"
            ElseIf parts(1).Length = 2 Then
                text = text & "0"
            End If
        Else
            text = text & ".000"
        End If
        Return text
    End Function

    ' ========== 两位小数格式化 ==========
    Private Function FormatDecimal2(value As Decimal) As String
        Dim rounded As Decimal = Math.Round(value, 2)
        Dim text As String = rounded.ToString()
        If text = "" Then Return "0.00"
        If text.Contains(".") Then
            Dim parts() As String = text.Split("."c)
            If parts(1).Length = 1 Then
                text = text & "0"
            End If
        Else
            text = text & ".00"
        End If
        Return text
    End Function

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
            ExportToExcel(dt, "销售查询报表表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 重置按钮 ==========
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        LoadShopList()
        LoadMaterialList()
        dgvReport.Rows.Clear()
        btnQuery_Click(Nothing, Nothing)
    End Sub

End Class

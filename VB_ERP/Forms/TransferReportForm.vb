' ============================================================================
' 商品调拨报表窗口
' 功能: 调拨数据统计，支持多维度筛选（日期/店铺 × 订单/明细/天/月/年 × 调出/调入）
' 对应易语言: 窗口程序集_窗口_商品调拨报表
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class TransferReportForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private localOrderCode As String = ""
    Private isStateChanging As Boolean = False
    Private summaryMode As String = "日期"

    ' ========== 查找条件变量 ==========
    Private searchStartDate As String = ""
    Private searchEndDate As String = ""
    Private searchShopFilter As String = ""
    Private searchDateShopFilter As String = ""
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
    Private radioDate As New RadioButton()
    Private radioShopMode As New RadioButton()
    Private radioOrder As New RadioButton()
    Private radioDetail As New RadioButton()
    Private radioDay As New RadioButton()
    Private radioMonth As New RadioButton()
    Private radioYear As New RadioButton()
    Private radioTransferOut As New RadioButton()
    Private radioTransferIn As New RadioButton()
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
        Me.Text = "商品调拨报表"
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
        dtpStart.Format = DateTimePickerFormat.Custom
        dtpStart.CustomFormat = "yyyy-MM-dd"
        panelTop.Controls.Add(dtpStart)

        AddLabel(panelTop, "结束：", 200, 15)
        dtpEnd.Location = New Drawing.Point(240, 12)
        dtpEnd.Size = New Drawing.Size(130, 25)
        dtpEnd.Format = DateTimePickerFormat.Custom
        dtpEnd.CustomFormat = "yyyy-MM-dd"
        panelTop.Controls.Add(dtpEnd)

        ' 汇总维度
        AddLabel(panelTop, "汇总：", 20, 45)
        radioDate.Text = "日期"
        radioDate.Location = New Drawing.Point(60, 42)
        radioDate.Size = New Drawing.Size(50, 25)
        panelTop.Controls.Add(radioDate)

        radioShopMode.Text = "店铺"
        radioShopMode.Location = New Drawing.Point(115, 42)
        radioShopMode.Size = New Drawing.Size(50, 25)
        panelTop.Controls.Add(radioShopMode)

        ' 报表模式
        AddLabel(panelTop, "模式：", 175, 45)
        radioOrder.Text = "订单"
        radioOrder.Location = New Drawing.Point(215, 42)
        radioOrder.Size = New Drawing.Size(50, 25)
        panelTop.Controls.Add(radioOrder)

        radioDetail.Text = "明细"
        radioDetail.Location = New Drawing.Point(270, 42)
        radioDetail.Size = New Drawing.Size(50, 25)
        panelTop.Controls.Add(radioDetail)

        radioDay.Text = "天"
        radioDay.Location = New Drawing.Point(325, 42)
        radioDay.Size = New Drawing.Size(35, 25)
        panelTop.Controls.Add(radioDay)

        radioMonth.Text = "月"
        radioMonth.Location = New Drawing.Point(365, 42)
        radioMonth.Size = New Drawing.Size(35, 25)
        panelTop.Controls.Add(radioMonth)

        radioYear.Text = "年"
        radioYear.Location = New Drawing.Point(405, 42)
        radioYear.Size = New Drawing.Size(35, 25)
        panelTop.Controls.Add(radioYear)

        ' 调拨方向
        AddLabel(panelTop, "方向：", 450, 45)
        radioTransferOut.Text = "调出"
        radioTransferOut.Location = New Drawing.Point(490, 42)
        radioTransferOut.Size = New Drawing.Size(50, 25)
        panelTop.Controls.Add(radioTransferOut)

        radioTransferIn.Text = "调入"
        radioTransferIn.Location = New Drawing.Point(545, 42)
        radioTransferIn.Size = New Drawing.Size(50, 25)
        panelTop.Controls.Add(radioTransferIn)

        ' 查找信息
        AddLabel(panelTop, "查找：", 20, 75)
        txtSearch.Location = New Drawing.Point(60, 72)
        txtSearch.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtSearch)

        ' 工厂查找
        AddLabel(panelTop, "工厂查找：", 220, 75)
        txtFactorySearch.Location = New Drawing.Point(285, 72)
        txtFactorySearch.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtFactorySearch)

        ' 店铺多选
        AddLabel(panelTop, "店铺：", 20, 105)
        chkShop.Location = New Drawing.Point(60, 102)
        chkShop.Size = New Drawing.Size(120, 55)
        chkShop.CheckOnClick = True
        panelTop.Controls.Add(chkShop)

        ' 品类多选
        AddLabel(panelTop, "品类：", 190, 105)
        chkCategory.Location = New Drawing.Point(230, 102)
        chkCategory.Size = New Drawing.Size(120, 55)
        chkCategory.CheckOnClick = True
        panelTop.Controls.Add(chkCategory)

        ' 规格多选
        AddLabel(panelTop, "规格：", 360, 105)
        chkSpec.Location = New Drawing.Point(400, 102)
        chkSpec.Size = New Drawing.Size(120, 55)
        chkSpec.CheckOnClick = True
        panelTop.Controls.Add(chkSpec)

        ' 工厂多选
        AddLabel(panelTop, "工厂：", 530, 105)
        chkFactory.Location = New Drawing.Point(570, 102)
        chkFactory.Size = New Drawing.Size(120, 55)
        chkFactory.CheckOnClick = True
        panelTop.Controls.Add(chkFactory)

        ' 按钮
        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(710, 12)
        btnQuery.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnQuery)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(800, 12)
        btnExport.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnExport)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(890, 12)
        btnReset.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnReset)

        ' DataGridView
        dgvReport.Dock = DockStyle.Fill
        dgvReport.ReadOnly = True
        dgvReport.AllowUserToAddRows = False
        dgvReport.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvReport.AllowUserToResizeRows = False
        dgvReport.RowHeadersVisible = False
        Me.Controls.Add(dgvReport)
        dgvReport.BringToFront()

        ' ContextMenu
        contextMenuReport.Items.AddRange({menuItemDetail, menuItemCopy, menuItemReturn})

        ' 事件绑定
        AddHandler radioDate.CheckedChanged, AddressOf radioDate_CheckedChanged
        AddHandler radioShopMode.CheckedChanged, AddressOf radioShopMode_CheckedChanged
        AddHandler radioOrder.CheckedChanged, AddressOf radioOrder_CheckedChanged
        AddHandler radioDetail.CheckedChanged, AddressOf radioDetail_CheckedChanged
        AddHandler radioDay.CheckedChanged, AddressOf radioDay_CheckedChanged
        AddHandler radioMonth.CheckedChanged, AddressOf radioMonth_CheckedChanged
        AddHandler radioYear.CheckedChanged, AddressOf radioYear_CheckedChanged
        AddHandler radioTransferOut.CheckedChanged, AddressOf radioTransferOut_CheckedChanged
        AddHandler radioTransferIn.CheckedChanged, AddressOf radioTransferIn_CheckedChanged
        AddHandler chkCategory.ItemCheck, AddressOf chkCategory_ItemCheck
        AddHandler chkSpec.ItemCheck, AddressOf chkSpec_ItemCheck
        AddHandler dgvReport.CellMouseUp, AddressOf dgvReport_CellMouseUp
        AddHandler menuItemDetail.Click, AddressOf menuItemDetail_Click
        AddHandler menuItemCopy.Click, AddressOf menuItemCopy_Click
        AddHandler menuItemReturn.Click, AddressOf menuItemReturn_Click
        AddHandler txtFactorySearch.KeyDown, AddressOf txtFactorySearch_KeyDown
        AddHandler txtSearch.KeyDown, AddressOf txtSearch_KeyDown
    End Sub

    Private Sub AddLabel(parent As Control, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        isStateChanging = True
        summaryMode = "日期"
        radioDate.Checked = True
        radioOrder.Checked = True
        radioTransferOut.Checked = True
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        isStateChanging = False

        InitQueryConditions()
        LoadHeaders()
        btnQuery_Click(Nothing, Nothing)
    End Sub

    ' ========== 表头加载 ==========
    Private Sub LoadHeaders()
        dgvReport.Columns.Clear()
        dgvReport.Rows.Clear()

        Dim headers As String() = Nothing

        If summaryMode = "日期" Then
            If radioOrder.Checked Then
                If localOrderCode = "" Then
                    headers = {"序号", "调拨单号", "调拨时间", "类型", "原库房", "新库房", "数量", "金重", "总重", "成本工费", "成本附加费", "成本价", "销售附加费", "备注", "操作账户"}
                Else
                    headers = {"序号", "商品编码", "原编码", "调拨时间", "调拨单号", "调拨类型", "商品名称", "品类", "规格", "款号", "调出库房", "调入库房", "单件重", "调拨数量", "调拨金重", "调拨重量", "每克工费", "成本附加费", "成本价", "销售附加费", "备注", "操作账户"}
                End If
            ElseIf radioDetail.Checked Then
                headers = {"序号", "商品编码", "原编码", "调拨时间", "调拨单号", "调拨类型", "商品名称", "品类", "规格", "款号", "调出库房", "调入库房", "单件重", "调拨数量", "调拨金重", "调拨重量", "每克工费", "成本附加费", "成本价", "销售附加费", "备注", "操作账户"}
            ElseIf radioDay.Checked OrElse radioMonth.Checked OrElse radioYear.Checked Then
                headers = {"序号", "日期", "调出数量", "调出金重", "调出重量", "调出工费", "调出附加费", "调出成本价", "", "调入数量", "调入金重", "调入数量", "调入工费", "调入附加费", "调入成本价"}
            End If
        ElseIf summaryMode = "店铺" Then
            If radioTransferOut.Checked Then
                headers = {"序号", "调出店铺", "调入店铺", "调出数量", "调出金重", "调出重量", "调出工费", "调出附加费", "调出成本价"}
            Else
                headers = {"序号", "调出店铺", "调入店铺", "调入数量", "调入金重", "调入数量", "调入工费", "调入附加费", "调入成本价"}
            End If
        End If

        If headers IsNot Nothing Then
            For Each h As String In headers
                dgvReport.Columns.Add(h, h)
            Next
        End If
    End Sub

    ' ========== 查询条件初始化 ==========
    Private Sub InitQueryConditions()
        localOrderCode = ""
        summaryMode = "日期"

        ' 店铺列表
        chkShop.Items.Clear()
        Dim shopSQL As String = "SELECT id AS akufang,CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in (" & GlobalVariables.UserShopPermission & ") UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN (" & GlobalVariables.UserShopPermission & ") ORDER BY akufang = '0' DESC, akufang"
        Dim shopDt As DataTable = DatabaseModule.ExecuteQuery(shopSQL, DatabaseModule.MySQL_ReadReport)
        For i As Integer = 0 To shopDt.Rows.Count - 1
            Dim item As New With {.ID = SafeString(shopDt.Rows(i)("akufang")), .Title = SafeString(shopDt.Rows(i)("btitle"))}
            chkShop.Items.Add(item, i = 0)
        Next

        ' 品类列表
        chkCategory.Items.Clear()
        Dim catSQL As String = "SELECT id, title FROM xipunum_erp_category WHERE 1=1 UNION ALL SELECT 0 as id, '未匹配' as title ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id"
        Dim catDt As DataTable = DatabaseModule.ExecuteQuery(catSQL, DatabaseModule.MySQL_ReadReport)
        For i As Integer = 0 To catDt.Rows.Count - 1
            Dim item As New With {.ID = SafeString(catDt.Rows(i)("id")), .Title = SafeString(catDt.Rows(i)("title"))}
            chkCategory.Items.Add(item, True)
        Next

        ' 规格列表
        chkSpec.Items.Clear()
        Dim specSQL As String = "SELECT guige.id as id,guige.title as title FROM (SELECT id,category_id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title UNION ALL SELECT 0 as id,0 as category_id,'未匹配' as title ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id) as guige"
        Dim specDt As DataTable = DatabaseModule.ExecuteQuery(specSQL, DatabaseModule.MySQL_ReadReport)
        For i As Integer = 0 To specDt.Rows.Count - 1
            Dim item As New With {.ID = SafeString(specDt.Rows(i)("id")), .Title = SafeString(specDt.Rows(i)("title"))}
            chkSpec.Items.Add(item, False)
        Next

        ' 工厂列表
        chkFactory.Items.Clear()

        ' 初始化规格ID列表
        UpdateSpecIdList()

        ' 初始控件状态
        UpdateControlState()
    End Sub

    ' ========== 规格ID列表更新 ==========
    Private Sub UpdateSpecIdList()
        _specIdList.Clear()

        ' 获取选中的品类
        Dim categoryIds As String = ""
        Dim catCount As Integer = 0
        For i As Integer = 0 To chkCategory.Items.Count - 1
            If chkCategory.GetItemChecked(i) Then
                Dim item = chkCategory.Items(i)
                categoryIds &= "'" & item.ID & "',"
                catCount += 1
            End If
        Next

        ' 获取选中的规格名称
        Dim specTitles As String = ""
        Dim specCount As Integer = 0
        For i As Integer = 0 To chkSpec.Items.Count - 1
            If chkSpec.GetItemChecked(i) Then
                Dim item = chkSpec.Items(i)
                specTitles &= "'" & item.Title & "',"
                specCount += 1
            End If
        Next

        Dim catFilter As String = ""
        If catCount > 0 Then
            catFilter = " and guige.category_id in (" & categoryIds.Substring(0, categoryIds.Length - 1) & ")"
        End If

        Dim specFilter As String = ""
        If specCount > 0 Then
            specFilter = " and guige.title in (" & specTitles.Substring(0, specTitles.Length - 1) & ")"
        End If

        Dim sql As String = "SELECT guige.id as id,guige.title as title FROM (SELECT id,category_id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY id UNION ALL SELECT 0 as id,0 as category_id,'未匹配' as title ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id) as guige WHERE 1=1 " & catFilter & specFilter & " ORDER BY guige.id asc"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, DatabaseModule.MySQL_ReadReport)
        For Each row As DataRow In dt.Rows
            _specIdList.Add(SafeString(row("id")))
        Next
    End Sub

    ' ========== 品类选中事件 → 刷新规格列表 ==========
    Private Sub chkCategory_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        If isStateChanging Then Return
        Me.BeginInvoke(New Action(Sub()
            ReloadSpecList()
            UpdateSpecIdList()
        End Sub))
    End Sub

    Private Sub ReloadSpecList()
        Dim categoryIds As String = ""
        Dim catCount As Integer = 0
        For i As Integer = 0 To chkCategory.Items.Count - 1
            If chkCategory.GetItemChecked(i) Then
                Dim item = chkCategory.Items(i)
                categoryIds &= item.ID & ","
                catCount += 1
            End If
        Next

        Dim sql As String
        If catCount = 0 Then
            sql = "SELECT guige.id as id,guige.title as title FROM (SELECT id,category_id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title UNION ALL SELECT 0 as id,0 as category_id,'未匹配' as title ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id) as guige ORDER BY guige.id asc"
        Else
            sql = "SELECT guige.id as id,guige.title as title FROM (SELECT id,category_id,title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title UNION ALL SELECT 0 as id,0 as category_id,'未匹配' as title ORDER BY CASE WHEN id = 0 THEN 0 ELSE 1 END, id) as guige WHERE guige.category_id in (" & categoryIds.Substring(0, categoryIds.Length - 1) & ")  ORDER BY guige.id asc"
        End If

        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, DatabaseModule.MySQL_ReadReport)
        chkSpec.Items.Clear()
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim item As New With {.ID = SafeString(dt.Rows(i)("id")), .Title = SafeString(dt.Rows(i)("title"))}
            chkSpec.Items.Add(item, False)
        Next
    End Sub

    Private Sub chkSpec_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        If isStateChanging Then Return
        Me.BeginInvoke(New Action(Sub()
            UpdateSpecIdList()
        End Sub))
    End Sub

    ' ========== 工厂查找 ==========
    Private Sub txtFactorySearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtFactorySearch.Text) Then
                ShowWarning("查找工厂不能为空！")
                txtFactorySearch.Text = ""
                Return
            End If
            SearchFactory(txtFactorySearch.Text)
            txtFactorySearch.Text = ""
        End If
    End Sub

    Private Sub SearchFactory(searchText As String)
        Dim sql As String = "SELECT id,title FROM xipunum_erp_about WHERE title like '%" & searchText & "%' OR jianxie like '%" & searchText & "%' OR name like '%" & searchText & "%' ORDER BY id asc"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, DatabaseModule.MySQL_ReadReport)
        chkFactory.Items.Clear()
        If dt.Rows.Count = 0 Then
            ShowWarning("查询无此信息数据！")
            Return
        End If
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim item As New With {.ID = SafeString(dt.Rows(i)("id")), .Title = SafeString(dt.Rows(i)("title"))}
            chkFactory.Items.Add(item, False)
        Next
    End Sub

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

            ' 获取店铺筛选
            Dim shopIds As String = ""
            Dim shopCount As Integer = 0
            For i As Integer = 0 To chkShop.Items.Count - 1
                If chkShop.GetItemChecked(i) Then
                    Dim item = chkShop.Items(i)
                    shopIds &= "'" & item.ID & "',"
                    shopCount += 1
                End If
            Next

            ' 查找日期库房 (无引号，逗号分隔)
            If shopCount > 0 Then
                searchDateShopFilter = shopIds.Substring(0, shopIds.Length - 1)
            Else
                searchDateShopFilter = GlobalVariables.UserShopPermission
            End If

            ' 查找信息库房 (带 and a.ykufang/xkufang in)
            If shopCount > 0 Then
                Dim idList As String = shopIds.Substring(0, shopIds.Length - 1)
                If radioTransferOut.Checked Then
                    searchShopFilter = " and a.ykufang in (" & idList & ")"
                End If
                If radioTransferIn.Checked Then
                    searchShopFilter = " and a.xkufang in (" & idList & ")"
                End If
            Else
                If radioTransferOut.Checked Then
                    searchShopFilter = " and a.ykufang in (" & GlobalVariables.UserShopPermission & ")"
                End If
                If radioTransferIn.Checked Then
                    searchShopFilter = " and a.xkufang in (" & GlobalVariables.UserShopPermission & ")"
                End If
            End If

            ' 规格筛选
            If _specIdList.Count > 0 Then
                searchSpecFilter = " WHERE tol.guigeid in (" & String.Join(",", _specIdList.Select(Function(x) "'" & x & "'")) & ")"
            Else
                searchSpecFilter = ""
            End If

            ' 工厂筛选
            Dim factoryIds As String = ""
            Dim factoryCount As Integer = 0
            For i As Integer = 0 To chkFactory.Items.Count - 1
                If chkFactory.GetItemChecked(i) Then
                    Dim item = chkFactory.Items(i)
                    factoryIds &= "'" & item.ID & "',"
                    factoryCount += 1
                End If
            Next
            If factoryCount > 0 Then
                searchFactoryFilter = " and f.factory in (" & factoryIds.Substring(0, factoryIds.Length - 1) & ")"
            Else
                searchFactoryFilter = ""
            End If

            ' 日期
            searchStartDate = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
            searchEndDate = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"

            ' 查找信息
            If String.IsNullOrEmpty(txtSearch.Text) Then
                searchName = ""
            Else
                searchName = txtSearch.Text
            End If

            ' 按模式查询
            If summaryMode = "日期" Then
                If radioOrder.Checked Then
                    QueryOrder()
                ElseIf radioDetail.Checked Then
                    QueryDetail()
                ElseIf radioDay.Checked OrElse radioMonth.Checked OrElse radioYear.Checked Then
                    QueryDate()
                End If
            ElseIf summaryMode = "店铺" Then
                QueryShop()
            End If

        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        Finally
            btnQuery.Enabled = True
            btnReset.Enabled = True
            btnExport.Enabled = True
        End Try
    End Sub

    ' ========== 内部SQL构建 ==========
    Private Function BuildInnerSQL(whereClause As String) As String
        Return "SELECT COALESCE ( e1.id, e2.id, '0' ) AS guigeid," &
               "j.transfer_umber AS dbdanhao,j.creationtime AS dbshijian,j.type as leixing,j.remarks as ddbaizhu," &
               "a.poduct_code AS bianma,b.fu_code AS fu_code,b.product_name as mingcheng," &
               "CASE WHEN COALESCE(d.title, '' ) = '' THEN '未匹配' ELSE d.title END AS pinlei," &
               "COALESCE (e1.title,e2.title, '未匹配' ) AS guige,b.item_number as kuanhao," &
               "CASE WHEN a.ykufang = '0' THEN '总库' ELSE h.title END AS ykufang," &
               "CASE WHEN a.xkufang = '0' THEN '总库' ELSE k.title END AS xkufang," &
               "COALESCE(b.single,0) as danjian,COALESCE(a.quantity,0) as shuliang,COALESCE(a.jinzhong,0) as jinzhong," &
               "COALESCE(CASE WHEN COALESCE ( i.lingxiao, '' ) = '是' THEN a.jinzhong ELSE b.single * a.quantity END,0) AS zongzhong," &
               "COALESCE(c.basic_cost,0) as chengbengf,COALESCE(c.company_surcharge,0) as cbfujia," &
               "COALESCE(c.cost_price * a.quantity,0) as chengben,COALESCE(c.sales_surcharge, 0) AS sales_surcharge," &
               "a.remarks as xqbeizhu,CONCAT( a.cjuser, '(', g.NAME, ')' ) AS ccjuser " &
               "FROM xipunum_erp_transfer AS a " &
               "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
               "INNER JOIN xipunum_erp_store AS c ON c.poduct_code = a.poduct_code " &
               "INNER JOIN xipunum_erp_store_order AS f ON f.id = c.order_id " &
               searchFactoryFilter &
               " LEFT JOIN xipunum_erp_type AS h ON h.id = a.ykufang" &
               " LEFT JOIN xipunum_erp_type AS k ON k.id = a.xkufang" &
               " LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = b.item_number AND b.item_number != ''" &
               " LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id AND b.item_number IS NOT NULL AND b.item_number != ''" &
               " LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id AND b.specification_id IS NOT NULL AND b.specification_id != ''" &
               " LEFT JOIN xipunum_erp_category AS d ON d.id = COALESCE ( e1.category_id, e2.category_id ) AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL " &
               " INNER JOIN xipunum_erp_user AS g ON g.USER = a.cjuser" &
               " INNER JOIN xipunum_erp_transfer_order AS j ON j.id = a.order_id" &
               " WHERE " & whereClause & " ORDER BY a.creationtime DESC"
    End Function

    Private Function BuildNormalWhereClause() As String
        Return "(b.poduct_code LIKE '%" & searchName & "%' OR b.fu_code LIKE '%" & searchName & "%' OR b.product_name LIKE '%" & searchName & "%' OR b.item_number LIKE '%" & searchName & "%' OR b.caizhi LIKE '%" & searchName & "%') AND a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "' " & searchShopFilter
    End Function

    ' ========== 订单模式查询 ==========
    Private Sub QueryOrder()
        dgvReport.Rows.Clear()

        If localOrderCode = "" Then
            ' 订单汇总
            Dim innerSQL As String = BuildInnerSQL(BuildNormalWhereClause())
            Dim outerSQL As String = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,CAST(ROUND(sum(tol.sales_surcharge), 2) AS DECIMAL (30, 2)) AS sales_surcharge,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & innerSQL & ") AS tol GROUP BY tol.dbdanhao,tol.dbshijian ORDER BY tol.dbshijian desc"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(outerSQL, DatabaseModule.MySQL_ReadReport)

            For i As Integer = 0 To dt.Rows.Count - 1
                Dim row As DataRow = dt.Rows(i)
                Dim idx As Integer = dgvReport.Rows.Add()
                dgvReport.Rows(idx).Cells(0).Value = (i + 1).ToString()
                dgvReport.Rows(idx).Cells(1).Value = SafeString(row("dbdanhao"))
                dgvReport.Rows(idx).Cells(2).Value = SafeString(row("dbshijian"))
                dgvReport.Rows(idx).Cells(3).Value = SafeString(row("leixing"))
                dgvReport.Rows(idx).Cells(4).Value = SafeString(row("ykufang"))
                dgvReport.Rows(idx).Cells(5).Value = SafeString(row("xkufang"))
                dgvReport.Rows(idx).Cells(6).Value = SafeString(row("shuliang"))
                dgvReport.Rows(idx).Cells(7).Value = SafeString(row("jinzhong"))
                dgvReport.Rows(idx).Cells(8).Value = SafeString(row("zongzhong"))
                dgvReport.Rows(idx).Cells(9).Value = SafeString(row("chengbengf"))
                dgvReport.Rows(idx).Cells(10).Value = SafeString(row("cbfujia"))
                dgvReport.Rows(idx).Cells(11).Value = SafeString(row("chengben"))
                dgvReport.Rows(idx).Cells(12).Value = SafeString(row("sales_surcharge"))
                dgvReport.Rows(idx).Cells(13).Value = SafeString(row("ddbaizhu"))
                dgvReport.Rows(idx).Cells(14).Value = SafeString(row("ccjuser"))
            Next

            ' 合计行
            Dim totalSQL As String = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,CAST(ROUND(sum(tol.sales_surcharge), 2) AS DECIMAL (30, 2)) AS sales_surcharge,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & innerSQL & ") AS tol"
            Dim totalDt As DataTable = DatabaseModule.ExecuteQuery(totalSQL, DatabaseModule.MySQL_ReadReport)
            If totalDt.Rows.Count > 0 Then
                Dim row As DataRow = totalDt.Rows(0)
                Dim idx As Integer = dgvReport.Rows.Add()
                dgvReport.Rows(idx).Cells(0).Value = "合计"
                dgvReport.Rows(idx).Cells(6).Value = SafeString(row("shuliang"))
                dgvReport.Rows(idx).Cells(7).Value = SafeString(row("jinzhong"))
                dgvReport.Rows(idx).Cells(8).Value = SafeString(row("zongzhong"))
                dgvReport.Rows(idx).Cells(9).Value = SafeString(row("chengbengf"))
                dgvReport.Rows(idx).Cells(10).Value = SafeString(row("cbfujia"))
                dgvReport.Rows(idx).Cells(11).Value = SafeString(row("chengben"))
                dgvReport.Rows(idx).Cells(12).Value = SafeString(row("sales_surcharge"))
            End If
        Else
            ' 订单明细 (drill-down)
            Dim drillWhere As String = "j.transfer_umber = '" & localOrderCode & "'"
            Dim drillInner As String = BuildInnerSQL(drillWhere)
            ' drill-down 不含 factoryFilter
            drillInner = drillInner.Replace(searchFactoryFilter, "")
            Dim outerSQL As String = "SELECT tol.bianma as bianma,tol.fu_code as fu_code,tol.dbshijian as dbshijian,tol.dbdanhao as dbdanhao,tol.leixing as leixing,tol.mingcheng as mingcheng,tol.pinlei as pinlei,tol.guige as guige,tol.kuanhao as kuanhao,tol.ykufang as ykufang,tol.xkufang as xkufang,tol.danjian as danjian,CAST(ROUND(tol.shuliang, 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(tol.jinzhong, 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(tol.zongzhong, 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(tol.chengbengf, 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(tol.cbfujia, 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(tol.chengben, 2) AS DECIMAL (30, 2)) as chengben,CAST(ROUND(tol.sales_surcharge, 2) AS DECIMAL (30, 2)) AS sales_surcharge,tol.ddbaizhu as xqbeizhu,tol.ccjuser as ccjuser FROM (" & drillInner & ") AS tol ORDER BY tol.dbshijian desc"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(outerSQL, DatabaseModule.MySQL_ReadReport)

            For i As Integer = 0 To dt.Rows.Count - 1
                Dim row As DataRow = dt.Rows(i)
                Dim idx As Integer = dgvReport.Rows.Add()
                dgvReport.Rows(idx).Cells(0).Value = (i + 1).ToString()
                dgvReport.Rows(idx).Cells(1).Value = SafeString(row("bianma"))
                dgvReport.Rows(idx).Cells(2).Value = SafeString(row("fu_code"))
                dgvReport.Rows(idx).Cells(3).Value = SafeString(row("dbshijian"))
                dgvReport.Rows(idx).Cells(4).Value = SafeString(row("dbdanhao"))
                dgvReport.Rows(idx).Cells(5).Value = SafeString(row("leixing"))
                dgvReport.Rows(idx).Cells(6).Value = SafeString(row("mingcheng"))
                dgvReport.Rows(idx).Cells(7).Value = SafeString(row("pinlei"))
                dgvReport.Rows(idx).Cells(8).Value = SafeString(row("guige"))
                dgvReport.Rows(idx).Cells(9).Value = SafeString(row("kuanhao"))
                dgvReport.Rows(idx).Cells(10).Value = SafeString(row("ykufang"))
                dgvReport.Rows(idx).Cells(11).Value = SafeString(row("xkufang"))
                dgvReport.Rows(idx).Cells(12).Value = SafeString(row("danjian"))
                dgvReport.Rows(idx).Cells(13).Value = SafeString(row("shuliang"))
                dgvReport.Rows(idx).Cells(14).Value = SafeString(row("jinzhong"))
                dgvReport.Rows(idx).Cells(15).Value = SafeString(row("zongzhong"))
                dgvReport.Rows(idx).Cells(16).Value = SafeString(row("chengbengf"))
                dgvReport.Rows(idx).Cells(17).Value = SafeString(row("cbfujia"))
                dgvReport.Rows(idx).Cells(18).Value = SafeString(row("chengben"))
                dgvReport.Rows(idx).Cells(19).Value = SafeString(row("sales_surcharge"))
                dgvReport.Rows(idx).Cells(20).Value = SafeString(row("xqbeizhu"))
                dgvReport.Rows(idx).Cells(21).Value = SafeString(row("ccjuser"))
            Next

            ' 明细合计 (使用订单汇总SQL with all filters)
            Dim innerSQL As String = BuildInnerSQL(BuildNormalWhereClause())
            Dim totalSQL As String = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,CAST(ROUND(sum(tol.sales_surcharge), 2) AS DECIMAL (30, 2)) AS sales_surcharge,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & innerSQL & ") AS tol"
            Dim totalDt As DataTable = DatabaseModule.ExecuteQuery(totalSQL, DatabaseModule.MySQL_ReadReport)
            If totalDt.Rows.Count > 0 Then
                Dim row As DataRow = totalDt.Rows(0)
                Dim idx As Integer = dgvReport.Rows.Add()
                dgvReport.Rows(idx).Cells(0).Value = "合计"
                dgvReport.Rows(idx).Cells(13).Value = SafeString(row("shuliang"))
                dgvReport.Rows(idx).Cells(14).Value = SafeString(row("jinzhong"))
                dgvReport.Rows(idx).Cells(15).Value = SafeString(row("zongzhong"))
                dgvReport.Rows(idx).Cells(16).Value = SafeString(row("chengbengf"))
                dgvReport.Rows(idx).Cells(17).Value = SafeString(row("cbfujia"))
                dgvReport.Rows(idx).Cells(18).Value = SafeString(row("chengben"))
                dgvReport.Rows(idx).Cells(19).Value = SafeString(row("sales_surcharge"))
            End If
        End If
    End Sub

    ' ========== 明细模式查询 ==========
    Private Sub QueryDetail()
        dgvReport.Rows.Clear()

        Dim innerSQL As String = BuildInnerSQL(BuildNormalWhereClause())
        Dim outerSQL As String = "SELECT tol.bianma as bianma,tol.fu_code as fu_code,tol.dbshijian as dbshijian,tol.dbdanhao as dbdanhao,tol.leixing as leixing,tol.mingcheng as mingcheng,tol.pinlei as pinlei,tol.guige as guige,tol.kuanhao as kuanhao,tol.ykufang as ykufang,tol.xkufang as xkufang,tol.danjian as danjian,CAST(ROUND(tol.shuliang, 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(tol.jinzhong, 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(tol.zongzhong, 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(tol.chengbengf, 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(tol.cbfujia, 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(tol.chengben, 2) AS DECIMAL (30, 2)) as chengben,CAST(ROUND(tol.sales_surcharge, 2) AS DECIMAL (30, 2)) AS sales_surcharge,tol.ddbaizhu as xqbeizhu,tol.ccjuser as ccjuser FROM (" & innerSQL & ") AS tol " & searchSpecFilter & " ORDER BY tol.dbshijian desc"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(outerSQL, DatabaseModule.MySQL_ReadReport)

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            Dim idx As Integer = dgvReport.Rows.Add()
            dgvReport.Rows(idx).Cells(0).Value = (i + 1).ToString()
            dgvReport.Rows(idx).Cells(1).Value = SafeString(row("bianma"))
            dgvReport.Rows(idx).Cells(2).Value = SafeString(row("fu_code"))
            dgvReport.Rows(idx).Cells(3).Value = SafeString(row("dbshijian"))
            dgvReport.Rows(idx).Cells(4).Value = SafeString(row("dbdanhao"))
            dgvReport.Rows(idx).Cells(5).Value = SafeString(row("leixing"))
            dgvReport.Rows(idx).Cells(6).Value = SafeString(row("mingcheng"))
            dgvReport.Rows(idx).Cells(7).Value = SafeString(row("pinlei"))
            dgvReport.Rows(idx).Cells(8).Value = SafeString(row("guige"))
            dgvReport.Rows(idx).Cells(9).Value = SafeString(row("kuanhao"))
            dgvReport.Rows(idx).Cells(10).Value = SafeString(row("ykufang"))
            dgvReport.Rows(idx).Cells(11).Value = SafeString(row("xkufang"))
            dgvReport.Rows(idx).Cells(12).Value = SafeString(row("danjian"))
            dgvReport.Rows(idx).Cells(13).Value = SafeString(row("shuliang"))
            dgvReport.Rows(idx).Cells(14).Value = SafeString(row("jinzhong"))
            dgvReport.Rows(idx).Cells(15).Value = SafeString(row("zongzhong"))
            dgvReport.Rows(idx).Cells(16).Value = SafeString(row("chengbengf"))
            dgvReport.Rows(idx).Cells(17).Value = SafeString(row("cbfujia"))
            dgvReport.Rows(idx).Cells(18).Value = SafeString(row("chengben"))
            dgvReport.Rows(idx).Cells(19).Value = SafeString(row("sales_surcharge"))
            dgvReport.Rows(idx).Cells(20).Value = SafeString(row("xqbeizhu"))
            dgvReport.Rows(idx).Cells(21).Value = SafeString(row("ccjuser"))
        Next

        ' 合计行
        Dim totalSQL As String = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,CAST(ROUND(sum(tol.sales_surcharge), 2) AS DECIMAL (30, 2)) AS sales_surcharge,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & innerSQL & ") AS tol"
        Dim totalDt As DataTable = DatabaseModule.ExecuteQuery(totalSQL, DatabaseModule.MySQL_ReadReport)
        If totalDt.Rows.Count > 0 Then
            Dim row As DataRow = totalDt.Rows(0)
            Dim idx As Integer = dgvReport.Rows.Add()
            dgvReport.Rows(idx).Cells(0).Value = "合计"
            dgvReport.Rows(idx).Cells(13).Value = SafeString(row("shuliang"))
            dgvReport.Rows(idx).Cells(14).Value = SafeString(row("jinzhong"))
            dgvReport.Rows(idx).Cells(15).Value = SafeString(row("zongzhong"))
            dgvReport.Rows(idx).Cells(16).Value = SafeString(row("chengbengf"))
            dgvReport.Rows(idx).Cells(17).Value = SafeString(row("cbfujia"))
            dgvReport.Rows(idx).Cells(18).Value = SafeString(row("chengben"))
            dgvReport.Rows(idx).Cells(19).Value = SafeString(row("sales_surcharge"))
        End If
    End Sub

    ' ========== 日期模式查询 (天/月/年) ==========
    Private Sub QueryDate()
        dgvReport.Rows.Clear()

        Dim reportStartStr As String = dtpStart.Value.ToString("yyyy-MM-dd")
        Dim reportEndStr As String = dtpEnd.Value.ToString("yyyy-MM-dd")

        Dim searchStart1Date As DateTime
        Dim searchEnd1Date As DateTime
        Dim loopCount As Integer

        If radioDay.Checked Then
            searchStart1Date = DateTime.Parse(reportStartStr)
            searchEnd1Date = DateTime.Parse(reportEndStr).AddDays(1)
            loopCount = CInt((DateTime.Parse(reportEndStr) - DateTime.Parse(reportStartStr)).TotalDays) + 1
        ElseIf radioMonth.Checked Then
            searchStart1Date = DateTime.Parse(dtpStart.Value.ToString("yyyy-MM") & "-01")
            searchEnd1Date = searchStart1Date.AddMonths(1)
            Dim endMonth As DateTime = DateTime.Parse(dtpEnd.Value.ToString("yyyy-MM") & "-01")
            loopCount = (endMonth.Year - searchStart1Date.Year) * 12 + (endMonth.Month - searchStart1Date.Month) + 1
        ElseIf radioYear.Checked Then
            searchStart1Date = DateTime.Parse(dtpStart.Value.ToString("yyyy") & "-01-01")
            searchEnd1Date = searchStart1Date.AddYears(1)
            Dim endYear As DateTime = DateTime.Parse(dtpEnd.Value.ToString("yyyy") & "-01-01")
            loopCount = endYear.Year - searchStart1Date.Year + 1
        Else
            Return
        End If

        Dim start1Str As String = searchStart1Date.ToString("yyyy-MM-dd") & " 00:00:00"
        Dim end1Str As String = searchEnd1Date.ToString("yyyy-MM-dd") & " 00:00:00"

        For i As Integer = 1 To loopCount
            Dim loopDateStr As String = ""
            Dim periodStart As DateTime
            Dim periodEnd As DateTime

            If radioDay.Checked Then
                Dim loopDate As DateTime = searchStart1Date.AddDays(i - 1)
                loopDateStr = loopDate.ToString("yyyy-MM-dd")
                periodStart = loopDate
                periodEnd = loopDate.AddDays(1)
            ElseIf radioMonth.Checked Then
                Dim loopDate As DateTime = searchStart1Date.AddMonths(i - 1)
                loopDateStr = loopDate.ToString("yyyy-MM")
                periodStart = DateTime.Parse(loopDateStr & "-01")
                periodEnd = periodStart.AddMonths(1)
            ElseIf radioYear.Checked Then
                Dim loopDate As DateTime = searchStart1Date.AddYears(i - 1)
                loopDateStr = loopDate.ToString("yyyy")
                periodStart = DateTime.Parse(loopDateStr & "-01-01")
                periodEnd = periodStart.AddYears(1)
            End If

            Dim periodStartStr As String = periodStart.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim periodEndStr As String = periodEnd.ToString("yyyy-MM-dd") & " 00:00:00"

            ' 调出数据
            Dim outWhere As String = "(b.poduct_code LIKE '%" & searchName & "%' OR b.fu_code LIKE '%" & searchName & "%' OR b.product_name LIKE '%" & searchName & "%' OR b.item_number LIKE '%" & searchName & "%' OR b.caizhi LIKE '%" & searchName & "%') AND a.creationtime >= '" & periodStartStr & "' AND a.creationtime < '" & periodEndStr & "' and a.ykufang in (" & searchDateShopFilter & ")"
            Dim outInner As String = BuildInnerSQL(outWhere)
            Dim outSQL As String = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & outInner & ") AS tol " & searchSpecFilter & " ORDER BY tol.dbshijian desc"
            Dim outDt As DataTable = DatabaseModule.ExecuteQuery(outSQL, DatabaseModule.MySQL_ReadReport)

            Dim outShuliang As String = ""
            Dim outJinzhong As String = ""
            Dim outZongzhong As String = ""
            Dim outChengbengf As String = ""
            Dim outCbfujia As String = ""
            Dim outChengben As String = ""
            If outDt.Rows.Count > 0 Then
                outShuliang = SafeString(outDt.Rows(0)("shuliang"))
                outJinzhong = SafeString(outDt.Rows(0)("jinzhong"))
                outZongzhong = SafeString(outDt.Rows(0)("zongzhong"))
                outChengbengf = SafeString(outDt.Rows(0)("chengbengf"))
                outCbfujia = SafeString(outDt.Rows(0)("cbfujia"))
                outChengben = SafeString(outDt.Rows(0)("chengben"))
            End If

            ' 调入数据
            Dim inWhere As String = "(b.poduct_code LIKE '%" & searchName & "%' OR b.fu_code LIKE '%" & searchName & "%' OR b.product_name LIKE '%" & searchName & "%' OR b.item_number LIKE '%" & searchName & "%' OR b.caizhi LIKE '%" & searchName & "%') AND a.creationtime >= '" & periodStartStr & "' AND a.creationtime < '" & periodEndStr & "' and a.xkufang in (" & searchDateShopFilter & ")"
            Dim inInner As String = BuildInnerSQL(inWhere)
            Dim inSQL As String = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & inInner & ") AS tol " & searchSpecFilter & " ORDER BY tol.dbshijian desc"
            Dim inDt As DataTable = DatabaseModule.ExecuteQuery(inSQL, DatabaseModule.MySQL_ReadReport)

            Dim inShuliang As String = ""
            Dim inJinzhong As String = ""
            Dim inZhongliang As String = ""
            Dim inCbgongfei As String = ""
            Dim inCbfujia As String = ""
            Dim inChengben As String = ""
            If inDt.Rows.Count > 0 Then
                Dim inRow As DataRow = inDt.Rows(0)
                inShuliang = GetFieldSafe(inRow, "shuliang")
                inJinzhong = GetFieldSafe(inRow, "jinzhong")
                inZhongliang = GetFieldSafe(inRow, "zhongliang")
                inCbgongfei = GetFieldSafe(inRow, "cbgongfei")
                inCbfujia = GetFieldSafe(inRow, "cbfujia")
                inChengben = GetFieldSafe(inRow, "chengben")
            End If

            ' 写入行
            Dim idx As Integer = dgvReport.Rows.Add()
            dgvReport.Rows(idx).Cells(0).Value = i.ToString()
            dgvReport.Rows(idx).Cells(1).Value = loopDateStr
            dgvReport.Rows(idx).Cells(2).Value = outShuliang
            dgvReport.Rows(idx).Cells(3).Value = outJinzhong
            dgvReport.Rows(idx).Cells(4).Value = outZongzhong
            dgvReport.Rows(idx).Cells(5).Value = outChengbengf
            dgvReport.Rows(idx).Cells(6).Value = outCbfujia
            dgvReport.Rows(idx).Cells(7).Value = outChengben
            ' col 8 is empty separator
            dgvReport.Rows(idx).Cells(9).Value = inShuliang
            dgvReport.Rows(idx).Cells(10).Value = inJinzhong
            dgvReport.Rows(idx).Cells(11).Value = inZhongliang
            dgvReport.Rows(idx).Cells(12).Value = inCbgongfei
            dgvReport.Rows(idx).Cells(13).Value = inCbfujia
            dgvReport.Rows(idx).Cells(14).Value = inChengben
        Next

        ' 合计行 - 调出
        Dim totalOutWhere As String = "(b.poduct_code LIKE '%" & searchName & "%' OR b.fu_code LIKE '%" & searchName & "%' OR b.product_name LIKE '%" & searchName & "%' OR b.item_number LIKE '%" & searchName & "%' OR b.caizhi LIKE '%" & searchName & "%') AND a.creationtime >= '" & start1Str & "' AND a.creationtime < '" & end1Str & "' and a.ykufang in (" & searchDateShopFilter & ")"
        Dim totalOutInner As String = BuildInnerSQL(totalOutWhere)
        Dim totalOutSQL As String = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & totalOutInner & ") AS tol " & searchSpecFilter & " ORDER BY tol.dbshijian desc"
        Dim totalOutDt As DataTable = DatabaseModule.ExecuteQuery(totalOutSQL, DatabaseModule.MySQL_ReadReport)

        Dim totOutShuliang As String = ""
        Dim totOutJinzhong As String = ""
        Dim totOutZongzhong As String = ""
        Dim totOutChengbengf As String = ""
        Dim totOutCbfujia As String = ""
        Dim totOutChengben As String = ""
        If totalOutDt.Rows.Count > 0 Then
            totOutShuliang = SafeString(totalOutDt.Rows(0)("shuliang"))
            totOutJinzhong = SafeString(totalOutDt.Rows(0)("jinzhong"))
            totOutZongzhong = SafeString(totalOutDt.Rows(0)("zongzhong"))
            totOutChengbengf = SafeString(totalOutDt.Rows(0)("chengbengf"))
            totOutCbfujia = SafeString(totalOutDt.Rows(0)("cbfujia"))
            totOutChengben = SafeString(totalOutDt.Rows(0)("chengben"))
        End If

        ' 合计行 - 调入
        Dim totalInWhere As String = "(b.poduct_code LIKE '%" & searchName & "%' OR b.fu_code LIKE '%" & searchName & "%' OR b.product_name LIKE '%" & searchName & "%' OR b.item_number LIKE '%" & searchName & "%' OR b.caizhi LIKE '%" & searchName & "%') AND a.creationtime >= '" & start1Str & "' AND a.creationtime < '" & end1Str & "' and a.xkufang in (" & searchDateShopFilter & ")"
        Dim totalInInner As String = BuildInnerSQL(totalInWhere)
        Dim totalInSQL As String = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & totalInInner & ") AS tol " & searchSpecFilter & " ORDER BY tol.dbshijian desc"
        Dim totalInDt As DataTable = DatabaseModule.ExecuteQuery(totalInSQL, DatabaseModule.MySQL_ReadReport)

        Dim totInShuliang As String = ""
        Dim totInJinzhong As String = ""
        Dim totInZhongliang As String = ""
        Dim totInCbgongfei As String = ""
        Dim totInCbfujia As String = ""
        Dim totInChengben As String = ""
        If totalInDt.Rows.Count > 0 Then
            Dim inRow As DataRow = totalInDt.Rows(0)
            totInShuliang = GetFieldSafe(inRow, "shuliang")
            totInJinzhong = GetFieldSafe(inRow, "jinzhong")
            totInZhongliang = GetFieldSafe(inRow, "zhongliang")
            totInCbgongfei = GetFieldSafe(inRow, "cbgongfei")
            totInCbfujia = GetFieldSafe(inRow, "cbfujia")
            totInChengben = GetFieldSafe(inRow, "chengben")
        End If

        Dim totalIdx As Integer = dgvReport.Rows.Add()
        dgvReport.Rows(totalIdx).Cells(0).Value = "合计"
        dgvReport.Rows(totalIdx).Cells(1).Value = ""
        dgvReport.Rows(totalIdx).Cells(2).Value = totOutShuliang
        dgvReport.Rows(totalIdx).Cells(3).Value = totOutJinzhong
        dgvReport.Rows(totalIdx).Cells(4).Value = totOutZongzhong
        dgvReport.Rows(totalIdx).Cells(5).Value = totOutChengbengf
        dgvReport.Rows(totalIdx).Cells(6).Value = totOutCbfujia
        dgvReport.Rows(totalIdx).Cells(7).Value = totOutChengben
        dgvReport.Rows(totalIdx).Cells(9).Value = totInShuliang
        dgvReport.Rows(totalIdx).Cells(10).Value = totInJinzhong
        dgvReport.Rows(totalIdx).Cells(11).Value = totInZhongliang
        dgvReport.Rows(totalIdx).Cells(12).Value = totInCbgongfei
        dgvReport.Rows(totalIdx).Cells(13).Value = totInCbfujia
        dgvReport.Rows(totalIdx).Cells(14).Value = totInChengben
    End Sub

    ' ========== 店铺模式查询 ==========
    Private Sub QueryShop()
        dgvReport.Rows.Clear()

        Dim dataSQL As String
        If radioTransferOut.Checked Then
            ' 调出: filter by ykufang, group by xkufang
            Dim whereClause As String = "(b.poduct_code LIKE '%" & searchName & "%' OR b.fu_code LIKE '%" & searchName & "%' OR b.product_name LIKE '%" & searchName & "%' OR b.item_number LIKE '%" & searchName & "%' OR b.caizhi LIKE '%" & searchName & "%') AND a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "' and a.ykufang in (" & searchDateShopFilter & ")"
            Dim innerSQL As String = BuildInnerSQL(whereClause)
            dataSQL = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & innerSQL & ") AS tol " & searchSpecFilter & " GROUP BY tol.xkufang ORDER BY tol.dbshijian desc"
        Else
            ' 调入: filter by xkufang, group by ykufang
            Dim whereClause As String = "(b.poduct_code LIKE '%" & searchName & "%' OR b.fu_code LIKE '%" & searchName & "%' OR b.product_name LIKE '%" & searchName & "%' OR b.item_number LIKE '%" & searchName & "%' OR b.caizhi LIKE '%" & searchName & "%') AND a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "' and a.xkufang in (" & searchDateShopFilter & ")"
            Dim innerSQL As String = BuildInnerSQL(whereClause)
            dataSQL = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & innerSQL & ") AS tol " & searchSpecFilter & " GROUP BY tol.ykufang ORDER BY tol.dbshijian desc"
        End If

        Dim dt As DataTable = DatabaseModule.ExecuteQuery(dataSQL, DatabaseModule.MySQL_ReadReport)

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)
            Dim idx As Integer = dgvReport.Rows.Add()
            dgvReport.Rows(idx).Cells(0).Value = (i + 1).ToString()
            dgvReport.Rows(idx).Cells(1).Value = SafeString(row("ykufang"))
            dgvReport.Rows(idx).Cells(2).Value = SafeString(row("xkufang"))
            dgvReport.Rows(idx).Cells(3).Value = SafeString(row("shuliang"))
            dgvReport.Rows(idx).Cells(4).Value = SafeString(row("jinzhong"))
            dgvReport.Rows(idx).Cells(5).Value = SafeString(row("zongzhong"))
            dgvReport.Rows(idx).Cells(6).Value = SafeString(row("chengbengf"))
            dgvReport.Rows(idx).Cells(7).Value = SafeString(row("cbfujia"))
            dgvReport.Rows(idx).Cells(8).Value = SafeString(row("chengben"))
        Next

        ' 合计行 (无 GROUP BY)
        Dim totalSQL As String
        If radioTransferOut.Checked Then
            Dim whereClause As String = "(b.poduct_code LIKE '%" & searchName & "%' OR b.fu_code LIKE '%" & searchName & "%' OR b.product_name LIKE '%" & searchName & "%' OR b.item_number LIKE '%" & searchName & "%' OR b.caizhi LIKE '%" & searchName & "%') AND a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "' and a.ykufang in (" & searchDateShopFilter & ")"
            Dim innerSQL As String = BuildInnerSQL(whereClause)
            totalSQL = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & innerSQL & ") AS tol " & searchSpecFilter & " ORDER BY tol.dbshijian desc"
        Else
            Dim whereClause As String = "(b.poduct_code LIKE '%" & searchName & "%' OR b.fu_code LIKE '%" & searchName & "%' OR b.product_name LIKE '%" & searchName & "%' OR b.item_number LIKE '%" & searchName & "%' OR b.caizhi LIKE '%" & searchName & "%') AND a.creationtime >= '" & searchStartDate & "' AND a.creationtime < '" & searchEndDate & "' and a.xkufang in (" & searchDateShopFilter & ")"
            Dim innerSQL As String = BuildInnerSQL(whereClause)
            totalSQL = "SELECT tol.dbdanhao as dbdanhao,tol.dbshijian as dbshijian,tol.leixing as leixing,tol.ykufang as ykufang,tol.xkufang as xkufang,CAST(ROUND(sum(tol.shuliang), 2) AS DECIMAL (30, 2)) as shuliang,CAST(ROUND(sum(tol.jinzhong), 3) AS DECIMAL (30, 3)) as jinzhong,CAST(ROUND(sum(tol.zongzhong), 3) AS DECIMAL (30, 3)) as zongzhong,CAST(ROUND(sum(tol.chengbengf*tol.jinzhong), 2) AS DECIMAL (30, 2)) as chengbengf,CAST(ROUND(sum(tol.cbfujia), 2) AS DECIMAL (30, 2)) as cbfujia,CAST(ROUND(sum(tol.chengben), 2) AS DECIMAL (30, 2)) as chengben,tol.ddbaizhu as ddbaizhu,tol.ccjuser as ccjuser FROM (" & innerSQL & ") AS tol " & searchSpecFilter & " ORDER BY tol.dbshijian desc"
        End If

        Dim totalDt As DataTable = DatabaseModule.ExecuteQuery(totalSQL, DatabaseModule.MySQL_ReadReport)
        If totalDt.Rows.Count > 0 Then
            Dim row As DataRow = totalDt.Rows(0)
            Dim idx As Integer = dgvReport.Rows.Add()
            dgvReport.Rows(idx).Cells(0).Value = "合计"
            dgvReport.Rows(idx).Cells(3).Value = SafeString(row("shuliang"))
            dgvReport.Rows(idx).Cells(4).Value = SafeString(row("jinzhong"))
            dgvReport.Rows(idx).Cells(5).Value = SafeString(row("zongzhong"))
            dgvReport.Rows(idx).Cells(6).Value = SafeString(row("chengbengf"))
            dgvReport.Rows(idx).Cells(7).Value = SafeString(row("cbfujia"))
            dgvReport.Rows(idx).Cells(8).Value = SafeString(row("chengben"))
        End If
    End Sub

    ' ========== 导出 ==========
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
                    newRow(i) = If(row.Cells(i).Value, "")
                Next
                dt.Rows.Add(newRow)
            Next
            ExportToExcel(dt, "商品调拨报表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 重置 ==========
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        isStateChanging = True
        localOrderCode = ""
        summaryMode = "日期"
        radioDate.Checked = True
        radioOrder.Checked = True
        radioTransferOut.Checked = True
        dtpStart.Format = DateTimePickerFormat.Custom
        dtpStart.CustomFormat = "yyyy-MM-dd"
        dtpEnd.Format = DateTimePickerFormat.Custom
        dtpEnd.CustomFormat = "yyyy-MM-dd"
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        isStateChanging = False
        InitQueryConditions()
        LoadHeaders()
        btnQuery_Click(Nothing, Nothing)
    End Sub

    ' ========== 右键菜单 ==========
    Private Sub dgvReport_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs)
        If e.Button = MouseButtons.Right AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If radioOrder.Checked Then
                If dgvReport.Rows(e.RowIndex).Cells(1).Value IsNot Nothing AndAlso dgvReport.Rows(e.RowIndex).Cells(1).Value.ToString() <> "" Then
                    dgvReport.Rows(e.RowIndex).Selected = True
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
    End Sub

    Private Sub menuItemDetail_Click(sender As Object, e As EventArgs)
        If dgvReport.CurrentRow IsNot Nothing Then
            localOrderCode = SafeString(dgvReport.CurrentRow.Cells(1).Value)
            LoadHeaders()
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub menuItemCopy_Click(sender As Object, e As EventArgs)
        If dgvReport.CurrentCell IsNot Nothing Then
            Dim cellValue As String = SafeString(dgvReport.CurrentCell.Value)
            Clipboard.SetText(cellValue)
            ShowInfo("复制:" & cellValue & " 到剪切板成功！")
        End If
    End Sub

    Private Sub menuItemReturn_Click(sender As Object, e As EventArgs)
        localOrderCode = ""
        LoadHeaders()
        btnQuery_Click(Nothing, Nothing)
    End Sub

    ' ========== 单选框事件 ==========
    Private Sub radioDate_CheckedChanged(sender As Object, e As EventArgs)
        If isStateChanging Then Return
        If radioDate.Checked Then
            isStateChanging = True
            summaryMode = "日期"
            radioDay.Enabled = True
            radioMonth.Enabled = True
            radioYear.Enabled = True
            radioDetail.Enabled = True
            radioOrder.Enabled = True
            radioTransferOut.Enabled = True
            radioTransferIn.Enabled = True
            If radioOrder.Checked Then
                chkCategory.Enabled = False
                chkSpec.Enabled = False
                chkFactory.Enabled = False
                txtFactorySearch.Enabled = False
            Else
                chkCategory.Enabled = True
                chkSpec.Enabled = True
                chkFactory.Enabled = True
                txtFactorySearch.Enabled = True
            End If
            UpdateDateFormat()
            isStateChanging = False
            LoadHeaders()
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub radioShopMode_CheckedChanged(sender As Object, e As EventArgs)
        If isStateChanging Then Return
        If radioShopMode.Checked Then
            isStateChanging = True
            summaryMode = "店铺"
            radioDay.Enabled = False
            radioMonth.Enabled = False
            radioYear.Enabled = False
            radioDetail.Enabled = False
            radioOrder.Enabled = False
            radioTransferOut.Enabled = True
            radioTransferIn.Enabled = True
            chkCategory.Enabled = True
            chkSpec.Enabled = True
            chkFactory.Enabled = True
            txtFactorySearch.Enabled = True
            dtpStart.Format = DateTimePickerFormat.Custom
            dtpStart.CustomFormat = "yyyy-MM-dd"
            dtpEnd.Format = DateTimePickerFormat.Custom
            dtpEnd.CustomFormat = "yyyy-MM-dd"
            dtpStart.Value = DateTime.Now
            dtpEnd.Value = DateTime.Now
            isStateChanging = False
            LoadHeaders()
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub radioOrder_CheckedChanged(sender As Object, e As EventArgs)
        If isStateChanging Then Return
        If radioOrder.Checked Then
            isStateChanging = True
            radioDetail.Checked = False
            radioDay.Checked = False
            radioMonth.Checked = False
            radioYear.Checked = False
            chkCategory.Enabled = False
            chkSpec.Enabled = False
            chkFactory.Enabled = False
            txtFactorySearch.Enabled = False
            radioTransferOut.Enabled = True
            radioTransferIn.Enabled = True
            localOrderCode = ""
            dtpStart.Format = DateTimePickerFormat.Custom
            dtpStart.CustomFormat = "yyyy-MM-dd"
            dtpEnd.Format = DateTimePickerFormat.Custom
            dtpEnd.CustomFormat = "yyyy-MM-dd"
            dtpStart.Value = DateTime.Now
            dtpEnd.Value = DateTime.Now
            isStateChanging = False
            LoadHeaders()
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
            chkSpec.Enabled = True
            chkFactory.Enabled = True
            txtFactorySearch.Enabled = True
            radioTransferOut.Enabled = True
            radioTransferIn.Enabled = True
            localOrderCode = ""
            dtpStart.Format = DateTimePickerFormat.Custom
            dtpStart.CustomFormat = "yyyy-MM-dd"
            dtpEnd.Format = DateTimePickerFormat.Custom
            dtpEnd.CustomFormat = "yyyy-MM-dd"
            dtpStart.Value = DateTime.Now
            dtpEnd.Value = DateTime.Now
            isStateChanging = False
            LoadHeaders()
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
            chkSpec.Enabled = True
            chkFactory.Enabled = True
            txtFactorySearch.Enabled = True
            radioTransferOut.Enabled = False
            radioTransferIn.Enabled = False
            localOrderCode = ""
            dtpStart.Format = DateTimePickerFormat.Custom
            dtpStart.CustomFormat = "yyyy-MM-dd"
            dtpEnd.Format = DateTimePickerFormat.Custom
            dtpEnd.CustomFormat = "yyyy-MM-dd"
            dtpStart.Value = DateTime.Now
            dtpEnd.Value = DateTime.Now
            isStateChanging = False
            LoadHeaders()
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
            chkSpec.Enabled = True
            chkFactory.Enabled = True
            txtFactorySearch.Enabled = True
            radioTransferOut.Enabled = False
            radioTransferIn.Enabled = False
            localOrderCode = ""
            dtpStart.Format = DateTimePickerFormat.Custom
            dtpStart.CustomFormat = "yyyy-MM"
            dtpEnd.Format = DateTimePickerFormat.Custom
            dtpEnd.CustomFormat = "yyyy-MM"
            dtpStart.Value = DateTime.Now
            dtpEnd.Value = DateTime.Now
            isStateChanging = False
            LoadHeaders()
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
            chkSpec.Enabled = True
            chkFactory.Enabled = True
            txtFactorySearch.Enabled = True
            radioTransferOut.Enabled = False
            radioTransferIn.Enabled = False
            localOrderCode = ""
            dtpStart.Format = DateTimePickerFormat.Custom
            dtpStart.CustomFormat = "yyyy"
            dtpEnd.Format = DateTimePickerFormat.Custom
            dtpEnd.CustomFormat = "yyyy"
            dtpStart.Value = DateTime.Now
            dtpEnd.Value = DateTime.Now
            isStateChanging = False
            LoadHeaders()
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub radioTransferOut_CheckedChanged(sender As Object, e As EventArgs)
        If isStateChanging Then Return
        If radioTransferOut.Checked Then
            isStateChanging = True
            radioTransferIn.Checked = False
            isStateChanging = False
            LoadHeaders()
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub radioTransferIn_CheckedChanged(sender As Object, e As EventArgs)
        If isStateChanging Then Return
        If radioTransferIn.Checked Then
            isStateChanging = True
            radioTransferOut.Checked = False
            isStateChanging = False
            LoadHeaders()
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    ' ========== 辅助方法 ==========
    Private Sub UpdateControlState()
        If summaryMode = "日期" Then
            If radioOrder.Checked Then
                chkCategory.Enabled = False
                chkSpec.Enabled = False
                chkFactory.Enabled = False
                txtFactorySearch.Enabled = False
            Else
                chkCategory.Enabled = True
                chkSpec.Enabled = True
                chkFactory.Enabled = True
                txtFactorySearch.Enabled = True
            End If
            If radioDay.Checked OrElse radioMonth.Checked OrElse radioYear.Checked Then
                radioTransferOut.Enabled = False
                radioTransferIn.Enabled = False
            End If
        End If
    End Sub

    Private Sub UpdateDateFormat()
        If radioDay.Checked OrElse radioOrder.Checked OrElse radioDetail.Checked Then
            dtpStart.Format = DateTimePickerFormat.Custom
            dtpStart.CustomFormat = "yyyy-MM-dd"
            dtpEnd.Format = DateTimePickerFormat.Custom
            dtpEnd.CustomFormat = "yyyy-MM-dd"
        ElseIf radioMonth.Checked Then
            dtpStart.Format = DateTimePickerFormat.Custom
            dtpStart.CustomFormat = "yyyy-MM"
            dtpEnd.Format = DateTimePickerFormat.Custom
            dtpEnd.CustomFormat = "yyyy-MM"
        ElseIf radioYear.Checked Then
            dtpStart.Format = DateTimePickerFormat.Custom
            dtpStart.CustomFormat = "yyyy"
            dtpEnd.Format = DateTimePickerFormat.Custom
            dtpEnd.CustomFormat = "yyyy"
        End If
    End Sub

    Private Function GetFieldSafe(row As DataRow, fieldName As String) As String
        If row.Table.Columns.Contains(fieldName) Then
            Return SafeString(row(fieldName))
        Else
            Return ""
        End If
    End Function
End Class

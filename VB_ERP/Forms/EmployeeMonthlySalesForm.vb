' ============================================================================
' 报表员工月销售统计窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_报表员工月销售统计.form.e.txt
' 功能: 员工月销售统计，支持重量/金额模式、所有/黄金/品类筛选、店铺/品类/规格/工厂多选
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class EmployeeMonthlySalesForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private 局部_查找类型文本 As String = ""                 ' 查找类型（员工月销售统计/导购数据明细）
    Private 查找开始日期 As String = ""                      ' 查找开始日期
    Private 查找结束日期 As String = ""                      ' 查找结束日期
    Private 查找信息库房 As String = ""                       ' 查找信息库房SQL片段
    Private 查找信息规格 As String = ""                       ' 查找信息规格SQL片段
    Private 查找信息工厂 As String = ""                       ' 查找信息工厂SQL片段
    Private 局部_查找导购批零 As String = ""                   ' 查找导购批零（零售/批发）
    Private 局部_查找导购账户 As String = ""                   ' 查找导购账户
    Private 查找结算品类 As String = ""                       ' 查找结算品类

    ' ========== 控件声明 ==========
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnReset As New Button()
    Private txtSearch As New TextBox()                        ' 编辑框_包含信息（工厂查找）
    Private rbWeight As New RadioButton()                     ' 单选框_重量
    Private rbAmount As New RadioButton()                     ' 单选框_金额
    Private rbAll As New RadioButton()                       ' 单选框_所有
    Private rbGold As New RadioButton()                      ' 单选框_黄金
    Private rbCategory As New RadioButton()                   ' 单选框_品类
    Private clbShop As New CheckedListBox()                  ' 店铺名称_超级列表框EX
    Private clbCategory As New CheckedListBox()             ' 品类名称_超级列表框EX
    Private clbSpec As New CheckedListBox()                 ' 规格名称_超级列表框EX
    Private clbFactory As New CheckedListBox()              ' 工厂名称_超级列表框EX
    Private clbSpecID As New CheckedListBox()               ' 规格名称id_超级列表框EX（隐藏）
    Private dgvReport As New DataGridView()                 ' 报表数据_高级表格Ex
    Private cmsMenu As New ContextMenuStrip()                ' 右键菜单
    Private tsmiDetail As New ToolStripMenuItem()           ' 数据明细
    Private tsmiReturn As New ToolStripMenuItem()           ' 返回列表

    ' ========== 初始化 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "员工月销售统计"
        Me.Size = New Drawing.Size(1366, 750)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' ========== 顶部查询条件面板 ==========
        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 35
        Me.Controls.Add(panelTop)

        Dim lblStart As New Label()
        lblStart.Text = "起始时间:"
        lblStart.Location = New Drawing.Point(10, 8)
        lblStart.AutoSize = True
        panelTop.Controls.Add(lblStart)

        dtpStart.Location = New Drawing.Point(70, 5)
        dtpStart.Size = New Drawing.Size(130, 25)
        dtpStart.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpStart)

        Dim lblEnd As New Label()
        lblEnd.Text = "结束时间:"
        lblEnd.Location = New Drawing.Point(210, 8)
        lblEnd.AutoSize = True
        panelTop.Controls.Add(lblEnd)

        dtpEnd.Location = New Drawing.Point(270, 5)
        dtpEnd.Size = New Drawing.Size(130, 25)
        dtpEnd.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpEnd)

        ' 重量/金额 单选框
        Dim grpMode As New GroupBox()
        grpMode.Text = ""
        grpMode.Location = New Drawing.Point(410, 1)
        grpMode.Size = New Drawing.Size(130, 32)
        panelTop.Controls.Add(grpMode)

        rbWeight.Text = "重量"
        rbWeight.Location = New Drawing.Point(5, 10)
        rbWeight.AutoSize = True
        rbWeight.Checked = True
        AddHandler rbWeight.CheckedChanged, AddressOf rbWeight_CheckedChanged
        grpMode.Controls.Add(rbWeight)

        rbAmount.Text = "金额"
        rbAmount.Location = New Drawing.Point(65, 10)
        rbAmount.AutoSize = True
        AddHandler rbAmount.CheckedChanged, AddressOf rbAmount_CheckedChanged
        grpMode.Controls.Add(rbAmount)

        ' 所有/黄金/品类 单选框
        Dim grpCategory As New GroupBox()
        grpCategory.Text = ""
        grpCategory.Location = New Drawing.Point(545, 1)
        grpCategory.Size = New Drawing.Size(190, 32)
        panelTop.Controls.Add(grpCategory)

        rbAll.Text = "所有"
        rbAll.Location = New Drawing.Point(5, 10)
        rbAll.AutoSize = True
        rbAll.Checked = True
        AddHandler rbAll.CheckedChanged, AddressOf rbAll_CheckedChanged
        grpCategory.Controls.Add(rbAll)

        rbGold.Text = "黄金"
        rbGold.Location = New Drawing.Point(65, 10)
        rbGold.AutoSize = True
        AddHandler rbGold.CheckedChanged, AddressOf rbGold_CheckedChanged
        grpCategory.Controls.Add(rbGold)

        rbCategory.Text = "品类"
        rbCategory.Location = New Drawing.Point(125, 10)
        rbCategory.AutoSize = True
        AddHandler rbCategory.CheckedChanged, AddressOf rbCategory_CheckedChanged
        grpCategory.Controls.Add(rbCategory)

        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(745, 4)
        btnQuery.Size = New Drawing.Size(70, 28)
        panelTop.Controls.Add(btnQuery)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(820, 4)
        btnReset.Size = New Drawing.Size(70, 28)
        panelTop.Controls.Add(btnReset)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(895, 4)
        btnExport.Size = New Drawing.Size(70, 28)
        panelTop.Controls.Add(btnExport)

        ' 工厂查找输入框
        txtSearch.Location = New Drawing.Point(970, 5)
        txtSearch.Size = New Drawing.Size(120, 25)
        txtSearch.Text = "请输入查找信息"
        AddHandler txtSearch.Enter, AddressOf txtSearch_Enter
        AddHandler txtSearch.Leave, AddressOf txtSearch_Leave
        AddHandler txtSearch.KeyDown, AddressOf txtSearch_KeyDown
        panelTop.Controls.Add(txtSearch)

        ' ========== 左侧筛选面板 ==========
        Dim panelLeft As New Panel()
        panelLeft.Dock = DockStyle.Left
        panelLeft.Width = 200
        panelTop.SendToBack()
        Me.Controls.Add(panelLeft)

        Dim y As Integer = 5

        ' 店铺名称
        Dim lblShop As New Label()
        lblShop.Text = "店铺名称"
        lblShop.Location = New Drawing.Point(5, y)
        lblShop.AutoSize = True
        panelLeft.Controls.Add(lblShop)
        y += 20

        clbShop.Location = New Drawing.Point(5, y)
        clbShop.Size = New Drawing.Size(190, 120)
        clbShop.CheckOnClick = True
        AddHandler clbShop.ItemCheck, AddressOf clbShop_ItemCheck
        panelLeft.Controls.Add(clbShop)
        y += 125

        ' 品类名称
        Dim lblCategory As New Label()
        lblCategory.Text = "品类名称"
        lblCategory.Location = New Drawing.Point(5, y)
        lblCategory.AutoSize = True
        panelLeft.Controls.Add(lblCategory)
        y += 20

        clbCategory.Location = New Drawing.Point(5, y)
        clbCategory.Size = New Drawing.Size(190, 100)
        clbCategory.CheckOnClick = True
        AddHandler clbCategory.ItemCheck, AddressOf clbCategory_ItemCheck
        panelLeft.Controls.Add(clbCategory)
        y += 105

        ' 规格名称
        Dim lblSpec As New Label()
        lblSpec.Text = "规格名称"
        lblSpec.Location = New Drawing.Point(5, y)
        lblSpec.AutoSize = True
        panelLeft.Controls.Add(lblSpec)
        y += 20

        clbSpec.Location = New Drawing.Point(5, y)
        clbSpec.Size = New Drawing.Size(190, 100)
        clbSpec.CheckOnClick = True
        AddHandler clbSpec.ItemCheck, AddressOf clbSpec_ItemCheck
        panelLeft.Controls.Add(clbSpec)
        y += 105

        ' 工厂名称
        Dim lblFactory As New Label()
        lblFactory.Text = "工厂名称"
        lblFactory.Location = New Drawing.Point(5, y)
        lblFactory.AutoSize = True
        panelLeft.Controls.Add(lblFactory)
        y += 20

        clbFactory.Location = New Drawing.Point(5, y)
        clbFactory.Size = New Drawing.Size(190, 100)
        clbFactory.CheckOnClick = True
        panelLeft.Controls.Add(clbFactory)

        ' 规格名称id（隐藏）
        clbSpecID.Visible = False

        ' ========== DataGridView ==========
        dgvReport.Dock = DockStyle.Fill
        dgvReport.ReadOnly = True
        dgvReport.AllowUserToAddRows = False
        dgvReport.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvReport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        AddHandler dgvReport.CellMouseUp, AddressOf dgvReport_CellMouseUp
        Me.Controls.Add(dgvReport)
        dgvReport.BringToFront()

        ' ========== 右键菜单 ==========
        tsmiDetail.Text = "数据明细"
        AddHandler tsmiDetail.Click, AddressOf tsmiDetail_Click
        cmsMenu.Items.Add(tsmiDetail)

        tsmiReturn.Text = "返回列表"
        AddHandler tsmiReturn.Click, AddressOf tsmiReturn_Click
        cmsMenu.Items.Add(tsmiReturn)
    End Sub

    ' ========== 窗口创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        局部_查找类型文本 = "员工月销售统计"
        dtpStart.Value = DateTime.Now
        dtpEnd.Value = DateTime.Now
        txtSearch.Text = "请输入查找信息"

        _查询条件_初始化()
        _子程序_加载表头()
        btnQuery_Click(Nothing, Nothing)
    End Sub

    ' ========== 子程序：加载表头 ==========
    Private Sub _子程序_加载表头()
        dgvReport.Columns.Clear()
        dgvReport.Rows.Clear()

        If 局部_查找类型文本 = "员工月销售统计" Then
            If rbWeight.Checked Then
                ' 重量模式表头
                Dim headers As String() = {"序号", "账户", "导购员", "品类", "销售数量", "销售金重", "", "退货数量", "退货金重", "", "实际数量", "实际金重", "", "所在店铺", "零批"}
                Dim widths As Integer() = {45, 100, 100, 100, 80, 80, 5, 80, 80, 5, 80, 80, 5, 100, 65}
                For i As Integer = 0 To headers.Length - 1
                    Dim col As New DataGridViewTextBoxColumn()
                    col.HeaderText = headers(i)
                    col.Name = "col" & i.ToString()
                    col.Width = widths(i)
                    If headers(i) = "" Then col.Visible = False
                    dgvReport.Columns.Add(col)
                Next
            Else
                ' 金额模式表头
                Dim headers As String() = {"序号", "账户", "导购员", "品类", "销售数量", "销售金额", "实销金额", "", "退货数量", "退货金额", "实退金额", "", "实际数量", "实际金额", "", "所在店铺", "零批"}
                Dim widths As Integer() = {45, 100, 100, 100, 80, 80, 80, 5, 80, 80, 80, 5, 80, 80, 5, 100, 65}
                For i As Integer = 0 To headers.Length - 1
                    Dim col As New DataGridViewTextBoxColumn()
                    col.HeaderText = headers(i)
                    col.Name = "col" & i.ToString()
                    col.Width = widths(i)
                    If headers(i) = "" Then col.Visible = False
                    dgvReport.Columns.Add(col)
                Next
            End If
        Else
            ' 导购数据明细表头
            Dim headers As String() = {"序号", "商品编码", "商品名称", "款号", "品类", "规格", "材质", "圈口/长度", "成色", "数量", "净重", "成本工费", "参考工费", "成本附加费", "销售单价", "销售金额", "原附加费", "销售克价", "销售工费", "销售附加费", "折扣", "实收金额", "导购员"}
            Dim widths As Integer() = {45, 100, 140, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70}
            For i As Integer = 0 To headers.Length - 1
                Dim col As New DataGridViewTextBoxColumn()
                col.HeaderText = headers(i)
                col.Name = "col" & i.ToString()
                col.Width = widths(i)
                dgvReport.Columns.Add(col)
            Next
        End If
    End Sub

    ' ========== 查询条件初始化 ==========
    Private Sub _查询条件_初始化()
        ' 加载店铺列表
        clbShop.Items.Clear()
        Dim sqlShop As String = "SELECT id AS akufang,CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in (" & UserShopPermission & ") UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN (" & UserShopPermission & ") ORDER BY akufang = '0' DESC, akufang"
        Dim dtShop As DataTable = ExecuteQuery(sqlShop, MySQL_ReadReport)
        For i As Integer = 0 To dtShop.Rows.Count - 1
            Dim id As String = SafeString(dtShop.Rows(i)("akufang"))
            Dim name As String = SafeString(dtShop.Rows(i)("btitle"))
            clbShop.Items.Add(name, If(i = 0, CheckState.Checked, CheckState.Unchecked))
            clbShop.Items(i).Tag = id
        Next
        If dtShop.Rows.Count <= 1 Then clbShop.Enabled = False

        ' 加载品类列表
        clbCategory.Items.Clear()
        Dim sqlCat As String = "SELECT id,title FROM xipunum_erp_category WHERE 1=1"
        Dim dtCat As DataTable = ExecuteQuery(sqlCat, MySQL_ReadReport)
        For i As Integer = 0 To dtCat.Rows.Count - 1
            Dim id As String = SafeString(dtCat.Rows(i)("id"))
            Dim title As String = SafeString(dtCat.Rows(i)("title"))
            clbCategory.Items.Add(title, CheckState.Checked)
            clbCategory.Items(i).Tag = id
        Next

        ' 加载规格列表
        clbSpec.Items.Clear()
        Dim sqlSpec As String = "SELECT title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title ORDER BY id asc"
        Dim dtSpec As DataTable = ExecuteQuery(sqlSpec, MySQL_ReadReport)
        For i As Integer = 0 To dtSpec.Rows.Count - 1
            Dim title As String = SafeString(dtSpec.Rows(i)("title"))
            clbSpec.Items.Add(title, CheckState.Checked)
            clbSpec.Items(i).Tag = (i + 1).ToString()
        Next

        ' 工厂列表初始为空
        clbFactory.Items.Clear()

        ' 选中规格id数据初始化
        _选中规格id数据初始化()
    End Sub

    ' ========== 品类选中_id数据初始化 ==========
    Private Sub _品类选中_id数据初始化()
        ' 获取选中的品类
        Dim 品类查看权限 As String = ""
        Dim 品类查看数量 As Integer = 0
        For i As Integer = 0 To clbCategory.Items.Count - 1
            If clbCategory.GetItemChecked(i) Then
                品类查看权限 &= clbCategory.Items(i).Tag.ToString() & ","
                品类查看数量 += 1
            End If
        Next

        clbSpec.Items.Clear()
        Dim sqlSpec As String = ""
        If 品类查看数量 = 0 Then
            sqlSpec = "SELECT title FROM xipunum_erp_specs WHERE 1=1 GROUP BY title ORDER BY id asc"
        Else
            sqlSpec = "SELECT title FROM xipunum_erp_specs WHERE category_id in (" & 品类查看权限.TrimEnd(","c) & ") GROUP BY title ORDER BY id asc"
        End If
        Dim dtSpec As DataTable = ExecuteQuery(sqlSpec, MySQL_ReadReport)
        For i As Integer = 0 To dtSpec.Rows.Count - 1
            Dim title As String = SafeString(dtSpec.Rows(i)("title"))
            clbSpec.Items.Add(title, CheckState.Unchecked)
            clbSpec.Items(i).Tag = (i + 1).ToString()
        Next

        _选中规格id数据初始化()
    End Sub

    ' ========== 选中规格id数据初始化 ==========
    Private Sub _选中规格id数据初始化()
        ' 获取选中的品类
        Dim 品类查看权限 As String = ""
        Dim 品类查看数量 As Integer = 0
        For i As Integer = 0 To clbCategory.Items.Count - 1
            If clbCategory.GetItemChecked(i) Then
                品类查看权限 &= "'" & clbCategory.Items(i).Tag.ToString() & "',"
                品类查看数量 += 1
            End If
        Next

        ' 获取选中的规格
        Dim 规格查看权限 As String = ""
        Dim 规格查看数量 As Integer = 0
        For i As Integer = 0 To clbSpec.Items.Count - 1
            If clbSpec.GetItemChecked(i) Then
                规格查看权限 &= "'" & clbSpec.Items(i).ToString() & "',"
                规格查看数量 += 1
            End If
        Next

        Dim 查找品类信息 As String = ""
        Dim 查找规格信息 As String = ""
        If 品类查看数量 > 0 Then
            查找品类信息 = " and category_id in (" & 品类查看权限.TrimEnd(","c) & ")"
        End If
        If 规格查看数量 > 0 Then
            查找规格信息 = " and title in (" & 规格查看权限.TrimEnd(","c) & ")"
        End If

        clbSpecID.Items.Clear()
        Dim sqlSpecID As String = "SELECT id,title FROM xipunum_erp_specs WHERE 1=1 " & 查找品类信息 & 查找规格信息 & " GROUP BY id ORDER BY id asc"
        Dim dtSpecID As DataTable = ExecuteQuery(sqlSpecID, MySQL_ReadReport)
        For i As Integer = 0 To dtSpecID.Rows.Count - 1
            Dim id As String = SafeString(dtSpecID.Rows(i)("id"))
            Dim title As String = SafeString(dtSpecID.Rows(i)("title"))
            clbSpecID.Items.Add(title & "|" & id)
            clbSpecID.Items(i).Tag = id
        Next
    End Sub

    ' ========== 工厂查找_数据初始化 ==========
    Private Sub _工厂查找_数据初始化()
        Dim 查找工厂名称信息 As String = txtSearch.Text.Trim()
        If String.IsNullOrEmpty(查找工厂名称信息) OrElse 查找工厂名称信息 = "请输入查找信息" Then Return

        clbFactory.Items.Clear()
        Dim sql As String = "SELECT id,title FROM xipunum_erp_about WHERE title like '%" & 查找工厂名称信息 & "%' OR jianxie like '%" & 查找工厂名称信息 & "%' OR name like '%" & 查找工厂名称信息 & "%' ORDER BY id asc"
        Dim dt As DataTable = ExecuteQuery(sql, MySQL_ReadReport)
        If dt.Rows.Count = 0 Then
            ShowWarning("查询无此信息数据！")
            txtSearch.Text = ""
            Return
        End If
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim id As String = SafeString(dt.Rows(i)("id"))
            Dim title As String = SafeString(dt.Rows(i)("title"))
            clbFactory.Items.Add(title, CheckState.Unchecked)
            clbFactory.Items(i).Tag = id
        Next
        txtSearch.Text = ""
    End Sub

    ' ========== 查询按钮 ==========
    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        btnQuery.Enabled = False
        btnReset.Enabled = False
        btnExport.Enabled = False
        Me.Text = "员工月销售统计 时间:" & dtpStart.Value.ToString("yyyy-MM-dd") & "至" & dtpEnd.Value.ToString("yyyy-MM-dd")

        ' 获取选中的店铺
        Dim 店铺名称权限 As String = ""
        Dim 店铺名称数量 As Integer = 0
        For i As Integer = 0 To clbShop.Items.Count - 1
            If clbShop.GetItemChecked(i) Then
                店铺名称权限 &= "'" & clbShop.Items(i).Tag.ToString() & "',"
                店铺名称数量 += 1
            End If
        Next

        ' 获取选中的规格id
        Dim 规格查看权限 As String = ""
        For i As Integer = 0 To clbSpecID.Items.Count - 1
            规格查看权限 &= "'" & clbSpecID.Items(i).Tag.ToString() & "',"
        Next

        ' 获取选中的工厂
        Dim 工厂查看权限 As String = ""
        Dim 工厂查看数量 As Integer = 0
        For i As Integer = 0 To clbFactory.Items.Count - 1
            If clbFactory.GetItemChecked(i) Then
                工厂查看权限 &= "'" & clbFactory.Items(i).Tag.ToString() & "',"
                工厂查看数量 += 1
            End If
        Next

        查找开始日期 = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
        查找结束日期 = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"
        查找结算品类 = SettlementCategoryList

        If 店铺名称数量 > 0 Then
            查找信息库房 = " and a.kufang in (" & 店铺名称权限.TrimEnd(","c) & ")"
        Else
            查找信息库房 = ""
        End If

        If clbSpecID.Items.Count > 0 Then
            查找信息规格 = " and COALESCE ( e1.id, e2.id, '0' ) in (" & 规格查看权限.TrimEnd(","c) & ")"
        Else
            查找信息规格 = ""
        End If

        If 工厂查看数量 > 0 Then
            查找信息工厂 = " and d.factory in (" & 工厂查看权限.TrimEnd(","c) & ")"
        Else
            查找信息工厂 = ""
        End If

        If rbWeight.Checked Then
            If 局部_查找类型文本 = "员工月销售统计" Then
                _子程序_导购业绩数据列表()
            Else
                _子程序_导购业绩数据详情()
            End If
        Else
            _子程序_导购业绩数据列表()
        End If

        btnQuery.Enabled = True
        btnReset.Enabled = True
        btnExport.Enabled = True
    End Sub

    ' ========== 子程序：导购业绩数据列表 ==========
    Private Sub _子程序_导购业绩数据列表()
        dgvReport.Rows.Clear()

        ' 构建主查询SQL - 获取每个导购+品类+批零的汇总数据
        Dim sqlBase As String = "SELECT CASE WHEN COALESCE(g.id, '' ) = '' THEN '0' ELSE g.id END AS pinleiid,CASE WHEN COALESCE(g.title, '' ) = '' THEN '未匹配' ELSE g.title END AS pinlei,a.shopping_guide as zhanghu,b.NAME AS daogou," &
            "SUM(CASE WHEN a.sales_return= 0 THEN a.quantity ELSE 0 END) AS positive_quantity," &
            "SUM(CASE WHEN a.sales_return= 0 THEN a.net_weight ELSE 0 END) AS positive_net_weight," &
            "SUM(CASE WHEN a.sales_return= 0 THEN a.xiao_amount ELSE 0 END) AS positive_xiao_amount," &
            "SUM(CASE WHEN a.sales_return= 0 THEN a.settlement ELSE 0 END) AS positive_settlement," &
            "SUM(CASE WHEN a.sales_return= 1 THEN a.quantity ELSE 0 END) AS negative_quantity," &
            "SUM(CASE WHEN a.sales_return= 1 THEN a.net_weight ELSE 0 END) AS negative_net_weight," &
            "SUM(CASE WHEN a.sales_return= 1 THEN a.xiao_amount ELSE 0 END) AS negative_xiao_amount," &
            "SUM(CASE WHEN a.sales_return= 1 THEN a.settlement ELSE 0 END) AS negative_settlement," &
            "SUM(a.quantity) AS total_quantity,SUM(a.net_weight) AS total_net_weight,SUM(a.settlement) AS total_settlement," &
            "f.title as kftitle,a.pling as apiling " &
            "FROM xipunum_erp_outbound AS a " &
            "INNER JOIN xipunum_erp_user AS b ON b.USER = a.shopping_guide " &
            "INNER JOIN xipunum_erp_store AS c ON c.poduct_code = a.poduct_code " &
            "INNER JOIN xipunum_erp_store_order AS d ON d.id = c.order_id " &
            查找信息工厂 &
            " INNER JOIN xipunum_erp_shop AS e ON e.poduct_code = a.poduct_code " &
            " LEFT JOIN xipunum_erp_type AS f ON f.id = b.department " &
            " LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = e.item_number AND e.item_number != ''" &
            " LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id AND e.item_number IS NOT NULL AND e.item_number != ''" &
            " LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = e.specification_id AND e.specification_id IS NOT NULL AND e.specification_id != ''" &
            " LEFT JOIN xipunum_erp_category AS g ON g.id = COALESCE ( e1.category_id, e2.category_id ) AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL" &
            " WHERE a.creationtime >= '" & 查找开始日期 & "'  AND a.creationtime < '" & 查找结束日期 & "'"

        ' 黄金模式追加品类筛选
        Dim sqlGoldFilter As String = ""
        If rbGold.Checked Then
            sqlGoldFilter = " and CASE WHEN COALESCE(g.id, '') = '' THEN '0' ELSE g.id END in (" & 查找结算品类 & ")"
        End If

        Dim sqlWhere As String = sqlGoldFilter & 查找信息库房 & 查找信息规格 &
            " GROUP BY a.shopping_guide,a.pling,CASE WHEN COALESCE(g.id, '' ) = '' THEN '0' ELSE g.id END" &
            " ORDER BY a.pling,a.shopping_guide,CASE WHEN COALESCE(g.title, '' ) = '' THEN '未匹配' ELSE g.title END DESC"

        Dim sqlMain As String = sqlBase & sqlWhere
        Dim dtMain As DataTable = ExecuteQuery(sqlMain, MySQL_ReadReport)

        ' 按导购分组，每个导购下按品类显示明细行，然后是小计行
        Dim 高级表格行数 As Integer = -1
        Dim currentAccount As String = ""
        Dim isFirst As Boolean = True

        ' 用于小计的变量
        Dim subSalesQty As Decimal = 0, subSalesWt As Decimal = 0, subSalesAmt As Decimal = 0, subRealAmt As Decimal = 0
        Dim subReturnQty As Decimal = 0, subReturnWt As Decimal = 0, subReturnAmt As Decimal = 0, subReturnRealAmt As Decimal = 0
        Dim subTotalQty As Decimal = 0, subTotalWt As Decimal = 0, subTotalAmt As Decimal = 0

        For i As Integer = 0 To dtMain.Rows.Count - 1
            Dim zhanghu As String = SafeString(dtMain.Rows(i)("zhanghu"))
            Dim daogou As String = SafeString(dtMain.Rows(i)("daogou"))
            Dim pinlei As String = SafeString(dtMain.Rows(i)("pinlei"))
            Dim pinleiid As String = SafeString(dtMain.Rows(i)("pinleiid"))
            Dim posQty As Decimal = SafeDecimal(dtMain.Rows(i)("positive_quantity"))
            Dim posWt As Decimal = SafeDecimal(dtMain.Rows(i)("positive_net_weight"))
            Dim posAmt As Decimal = SafeDecimal(dtMain.Rows(i)("positive_xiao_amount"))
            Dim posSet As Decimal = SafeDecimal(dtMain.Rows(i)("positive_settlement"))
            Dim negQty As Decimal = SafeDecimal(dtMain.Rows(i)("negative_quantity"))
            Dim negWt As Decimal = SafeDecimal(dtMain.Rows(i)("negative_net_weight"))
            Dim negAmt As Decimal = SafeDecimal(dtMain.Rows(i)("negative_xiao_amount"))
            Dim negSet As Decimal = SafeDecimal(dtMain.Rows(i)("negative_settlement"))
            Dim totQty As Decimal = SafeDecimal(dtMain.Rows(i)("total_quantity"))
            Dim totWt As Decimal = SafeDecimal(dtMain.Rows(i)("total_net_weight"))
            Dim totSet As Decimal = SafeDecimal(dtMain.Rows(i)("total_settlement"))
            Dim kftitle As String = SafeString(dtMain.Rows(i)("kftitle"))
            Dim apiling As String = SafeString(dtMain.Rows(i)("apiling"))

            ' 检测导购变化 - 写入小计行
            If zhanghu <> currentAccount Then
                If Not isFirst Then
                    ' 写入小计行
                    高级表格行数 += 1
                    dgvReport.Rows.Add()
                    dgvReport.Rows(高级表格行数).Cells(0).Value = ""
                    dgvReport.Rows(高级表格行数).Cells(1).Value = currentAccount
                    dgvReport.Rows(高级表格行数).Cells(2).Value = dtMain.Rows(i - 1)("daogou").ToString()
                    dgvReport.Rows(高级表格行数).Cells(3).Value = "小计"
                    Dim subRow As DataGridViewRow = dgvReport.Rows(高级表格行数)
                    If rbWeight.Checked Then
                        subRow.Cells(4).Value = subSalesQty.ToString("F2")
                        subRow.Cells(5).Value = subSalesWt.ToString("F3")
                        subRow.Cells(7).Value = subReturnQty.ToString("F2")
                        subRow.Cells(8).Value = subReturnWt.ToString("F3")
                        subRow.Cells(10).Value = subTotalQty.ToString("F2")
                        subRow.Cells(11).Value = subTotalWt.ToString("F3")
                        subRow.Cells(13).Value = kftitle
                        subRow.Cells(14).Value = apiling
                    Else
                        subRow.Cells(4).Value = subSalesQty.ToString("F2")
                        subRow.Cells(5).Value = subSalesAmt.ToString("F2")
                        subRow.Cells(6).Value = subRealAmt.ToString("F2")
                        subRow.Cells(8).Value = subReturnQty.ToString("F2")
                        subRow.Cells(9).Value = subReturnAmt.ToString("F2")
                        subRow.Cells(10).Value = subReturnRealAmt.ToString("F2")
                        subRow.Cells(12).Value = subTotalQty.ToString("F2")
                        subRow.Cells(13).Value = subTotalAmt.ToString("F2")
                        subRow.Cells(15).Value = kftitle
                        subRow.Cells(16).Value = apiling
                    End If
                    ' 重置小计
                    subSalesQty = 0 : subSalesWt = 0 : subSalesAmt = 0 : subRealAmt = 0
                    subReturnQty = 0 : subReturnWt = 0 : subReturnAmt = 0 : subReturnRealAmt = 0
                    subTotalQty = 0 : subTotalWt = 0 : subTotalAmt = 0
                End If
                currentAccount = zhanghu
                isFirst = False
            End If

            ' 累加小计
            subSalesQty += posQty : subSalesWt += posWt : subSalesAmt += posAmt : subRealAmt += posSet
            subReturnQty += negQty : subReturnWt += negWt : subReturnAmt += negAmt : subReturnRealAmt += negSet
            subTotalQty += totQty : subTotalWt += totWt : subTotalAmt += totSet

            ' 写入明细行
            高级表格行数 += 1
            dgvReport.Rows.Add()
            Dim row As DataGridViewRow = dgvReport.Rows(高级表格行数)
            row.Cells(0).Value = (高级表格行数 + 1).ToString()
            row.Cells(1).Value = ""
            row.Cells(2).Value = ""
            row.Cells(3).Value = pinlei
            If rbWeight.Checked Then
                row.Cells(4).Value = posQty.ToString("F2")
                row.Cells(5).Value = posWt.ToString("F3")
                row.Cells(7).Value = negQty.ToString("F2")
                row.Cells(8).Value = negWt.ToString("F3")
                row.Cells(10).Value = totQty.ToString("F2")
                row.Cells(11).Value = totWt.ToString("F3")
                row.Cells(13).Value = kftitle
                row.Cells(14).Value = apiling
            Else
                row.Cells(4).Value = posQty.ToString("F2")
                row.Cells(5).Value = posAmt.ToString("F2")
                row.Cells(6).Value = posSet.ToString("F2")
                row.Cells(8).Value = negQty.ToString("F2")
                row.Cells(9).Value = negAmt.ToString("F2")
                row.Cells(10).Value = negSet.ToString("F2")
                row.Cells(12).Value = totQty.ToString("F2")
                row.Cells(13).Value = totSet.ToString("F2")
                row.Cells(15).Value = kftitle
                row.Cells(16).Value = apiling
            End If
        Next

        ' 写入最后一组的小计
        If dtMain.Rows.Count > 0 Then
            高级表格行数 += 1
            dgvReport.Rows.Add()
            Dim lastIdx As Integer = dtMain.Rows.Count - 1
            dgvReport.Rows(高级表格行数).Cells(0).Value = ""
            dgvReport.Rows(高级表格行数).Cells(1).Value = currentAccount
            dgvReport.Rows(高级表格行数).Cells(2).Value = SafeString(dtMain.Rows(lastIdx)("daogou"))
            dgvReport.Rows(高级表格行数).Cells(3).Value = "小计"
            Dim subRow2 As DataGridViewRow = dgvReport.Rows(高级表格行数)
            If rbWeight.Checked Then
                subRow2.Cells(4).Value = subSalesQty.ToString("F2")
                subRow2.Cells(5).Value = subSalesWt.ToString("F3")
                subRow2.Cells(7).Value = subReturnQty.ToString("F2")
                subRow2.Cells(8).Value = subReturnWt.ToString("F3")
                subRow2.Cells(10).Value = subTotalQty.ToString("F2")
                subRow2.Cells(11).Value = subTotalWt.ToString("F3")
            Else
                subRow2.Cells(4).Value = subSalesQty.ToString("F2")
                subRow2.Cells(5).Value = subSalesAmt.ToString("F2")
                subRow2.Cells(6).Value = subRealAmt.ToString("F2")
                subRow2.Cells(8).Value = subReturnQty.ToString("F2")
                subRow2.Cells(9).Value = subReturnAmt.ToString("F2")
                subRow2.Cells(10).Value = subReturnRealAmt.ToString("F2")
                subRow2.Cells(12).Value = subTotalQty.ToString("F2")
                subRow2.Cells(13).Value = subTotalAmt.ToString("F2")
            End If
        End If

        ' ========== 零售合计 ==========
        Dim sqlSumBase As String = "SELECT " &
            "SUM(CASE WHEN a.sales_return= 0 THEN a.quantity ELSE 0 END) AS positive_quantity," &
            "SUM(CASE WHEN a.sales_return= 0 THEN a.net_weight ELSE 0 END) AS positive_net_weight," &
            "SUM(CASE WHEN a.sales_return= 0 THEN a.xiao_amount ELSE 0 END) AS positive_xiao_amount," &
            "SUM(CASE WHEN a.sales_return= 0 THEN a.settlement ELSE 0 END) AS positive_settlement," &
            "SUM(CASE WHEN a.sales_return= 1 THEN a.quantity ELSE 0 END) AS negative_quantity," &
            "SUM(CASE WHEN a.sales_return= 1 THEN a.net_weight ELSE 0 END) AS negative_net_weight," &
            "SUM(CASE WHEN a.sales_return= 1 THEN a.xiao_amount ELSE 0 END) AS negative_xiao_amount," &
            "SUM(CASE WHEN a.sales_return= 1 THEN a.settlement ELSE 0 END) AS negative_settlement," &
            "SUM(a.quantity) AS total_quantity,SUM(a.net_weight) AS total_net_weight,SUM(a.settlement) AS total_settlement " &
            "FROM xipunum_erp_outbound AS a " &
            "INNER JOIN xipunum_erp_user AS b ON b.USER = a.shopping_guide " &
            "INNER JOIN xipunum_erp_store AS c ON c.poduct_code = a.poduct_code " &
            "INNER JOIN xipunum_erp_store_order AS d ON d.id = c.order_id " &
            查找信息工厂 &
            " INNER JOIN xipunum_erp_shop AS e ON e.poduct_code = a.poduct_code " &
            " LEFT JOIN xipunum_erp_type AS f ON f.id = b.department " &
            " LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = e.item_number AND e.item_number != ''" &
            " LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id AND e.item_number IS NOT NULL AND e.item_number != ''" &
            " LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = e.specification_id AND e.specification_id IS NOT NULL AND e.specification_id != ''" &
            " LEFT JOIN xipunum_erp_category AS g ON g.id = COALESCE ( e1.category_id, e2.category_id ) AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL" &
            " WHERE a.creationtime >= '" & 查找开始日期 & "'  AND a.creationtime < '" & 查找结束日期 & "'"

        Dim sqlRetail As String = sqlSumBase & sqlGoldFilter & " and a.pling='零售' " & 查找信息库房 & 查找信息规格 &
            " ORDER BY a.pling,a.shopping_guide,CASE WHEN COALESCE(g.title, '' ) = '' THEN '未匹配' ELSE g.title END DESC"
        Dim dtRetail As DataTable = ExecuteQuery(sqlRetail, MySQL_ReadReport)

        高级表格行数 += 1
        dgvReport.Rows.Add()
        dgvReport.Rows(高级表格行数).Cells(1).Value = "零售合计"
        If dtRetail.Rows.Count > 0 Then
            Dim r As DataGridViewRow = dgvReport.Rows(高级表格行数)
            If rbWeight.Checked Then
                r.Cells(4).Value = SafeDecimal(dtRetail.Rows(0)("positive_quantity")).ToString("F2")
                r.Cells(5).Value = SafeDecimal(dtRetail.Rows(0)("positive_net_weight")).ToString("F3")
                r.Cells(7).Value = SafeDecimal(dtRetail.Rows(0)("negative_quantity")).ToString("F2")
                r.Cells(8).Value = SafeDecimal(dtRetail.Rows(0)("negative_net_weight")).ToString("F3")
                r.Cells(10).Value = SafeDecimal(dtRetail.Rows(0)("total_quantity")).ToString("F2")
                r.Cells(11).Value = SafeDecimal(dtRetail.Rows(0)("total_net_weight")).ToString("F3")
            Else
                r.Cells(4).Value = SafeDecimal(dtRetail.Rows(0)("positive_quantity")).ToString("F2")
                r.Cells(5).Value = SafeDecimal(dtRetail.Rows(0)("positive_xiao_amount")).ToString("F2")
                r.Cells(6).Value = SafeDecimal(dtRetail.Rows(0)("positive_settlement")).ToString("F2")
                r.Cells(8).Value = SafeDecimal(dtRetail.Rows(0)("negative_quantity")).ToString("F2")
                r.Cells(9).Value = SafeDecimal(dtRetail.Rows(0)("negative_xiao_amount")).ToString("F2")
                r.Cells(10).Value = SafeDecimal(dtRetail.Rows(0)("negative_settlement")).ToString("F2")
                r.Cells(12).Value = SafeDecimal(dtRetail.Rows(0)("total_quantity")).ToString("F2")
                r.Cells(13).Value = SafeDecimal(dtRetail.Rows(0)("total_settlement")).ToString("F2")
            End If
        End If

        ' ========== 批发合计 ==========
        Dim sqlWholesale As String = sqlSumBase & sqlGoldFilter & " and a.pling='批发' " & 查找信息库房 & 查找信息规格 &
            " ORDER BY a.pling,a.shopping_guide,CASE WHEN COALESCE(g.title, '' ) = '' THEN '未匹配' ELSE g.title END DESC"
        Dim dtWholesale As DataTable = ExecuteQuery(sqlWholesale, MySQL_ReadReport)

        高级表格行数 += 1
        dgvReport.Rows.Add()
        dgvReport.Rows(高级表格行数).Cells(1).Value = "批发合计"
        If dtWholesale.Rows.Count > 0 Then
            Dim r As DataGridViewRow = dgvReport.Rows(高级表格行数)
            If rbWeight.Checked Then
                r.Cells(4).Value = SafeDecimal(dtWholesale.Rows(0)("positive_quantity")).ToString("F2")
                r.Cells(5).Value = SafeDecimal(dtWholesale.Rows(0)("positive_net_weight")).ToString("F3")
                r.Cells(7).Value = SafeDecimal(dtWholesale.Rows(0)("negative_quantity")).ToString("F2")
                r.Cells(8).Value = SafeDecimal(dtWholesale.Rows(0)("negative_net_weight")).ToString("F3")
                r.Cells(10).Value = SafeDecimal(dtWholesale.Rows(0)("total_quantity")).ToString("F2")
                r.Cells(11).Value = SafeDecimal(dtWholesale.Rows(0)("total_net_weight")).ToString("F3")
            Else
                r.Cells(4).Value = SafeDecimal(dtWholesale.Rows(0)("positive_quantity")).ToString("F2")
                r.Cells(5).Value = SafeDecimal(dtWholesale.Rows(0)("positive_xiao_amount")).ToString("F2")
                r.Cells(6).Value = SafeDecimal(dtWholesale.Rows(0)("positive_settlement")).ToString("F2")
                r.Cells(8).Value = SafeDecimal(dtWholesale.Rows(0)("negative_quantity")).ToString("F2")
                r.Cells(9).Value = SafeDecimal(dtWholesale.Rows(0)("negative_xiao_amount")).ToString("F2")
                r.Cells(10).Value = SafeDecimal(dtWholesale.Rows(0)("negative_settlement")).ToString("F2")
                r.Cells(12).Value = SafeDecimal(dtWholesale.Rows(0)("total_quantity")).ToString("F2")
                r.Cells(13).Value = SafeDecimal(dtWholesale.Rows(0)("total_settlement")).ToString("F2")
            End If
        End If

        ' ========== 合计 ==========
        Dim sqlTotal As String = sqlSumBase & sqlGoldFilter & " " & 查找信息库房 & 查找信息规格 &
            " ORDER BY a.pling,a.shopping_guide,CASE WHEN COALESCE(g.title, '' ) = '' THEN '未匹配' ELSE g.title END DESC"
        Dim dtTotal As DataTable = ExecuteQuery(sqlTotal, MySQL_ReadReport)

        高级表格行数 += 1
        dgvReport.Rows.Add()
        dgvReport.Rows(高级表格行数).Cells(1).Value = "合计"
        If dtTotal.Rows.Count > 0 Then
            Dim r As DataGridViewRow = dgvReport.Rows(高级表格行数)
            If rbWeight.Checked Then
                r.Cells(4).Value = SafeDecimal(dtTotal.Rows(0)("positive_quantity")).ToString("F2")
                r.Cells(5).Value = SafeDecimal(dtTotal.Rows(0)("positive_net_weight")).ToString("F3")
                r.Cells(7).Value = SafeDecimal(dtTotal.Rows(0)("negative_quantity")).ToString("F2")
                r.Cells(8).Value = SafeDecimal(dtTotal.Rows(0)("negative_net_weight")).ToString("F3")
                r.Cells(10).Value = SafeDecimal(dtTotal.Rows(0)("total_quantity")).ToString("F2")
                r.Cells(11).Value = SafeDecimal(dtTotal.Rows(0)("total_net_weight")).ToString("F3")
            Else
                r.Cells(4).Value = SafeDecimal(dtTotal.Rows(0)("positive_quantity")).ToString("F2")
                r.Cells(5).Value = SafeDecimal(dtTotal.Rows(0)("positive_xiao_amount")).ToString("F2")
                r.Cells(6).Value = SafeDecimal(dtTotal.Rows(0)("positive_settlement")).ToString("F2")
                r.Cells(8).Value = SafeDecimal(dtTotal.Rows(0)("negative_quantity")).ToString("F2")
                r.Cells(9).Value = SafeDecimal(dtTotal.Rows(0)("negative_xiao_amount")).ToString("F2")
                r.Cells(10).Value = SafeDecimal(dtTotal.Rows(0)("negative_settlement")).ToString("F2")
                r.Cells(12).Value = SafeDecimal(dtTotal.Rows(0)("total_quantity")).ToString("F2")
                r.Cells(13).Value = SafeDecimal(dtTotal.Rows(0)("total_settlement")).ToString("F2")
            End If
        End If
    End Sub

    ' ========== 子程序：导购业绩数据详情 ==========
    Private Sub _子程序_导购业绩数据详情()
        dgvReport.Rows.Clear()

        Dim sqlDetail As String = "SELECT a.poduct_code AS bianma,e.product_name as mingcheng,e.item_number as kuanhao," &
            "CASE WHEN COALESCE ( g.title, '' ) = '' THEN '未匹配' ELSE g.title END AS pinlei," &
            "COALESCE (e1.title,e2.title, '未匹配' ) AS guige,e.caizhi as caizhi,e.quandu as quankou,c.factory_condition as chengse," &
            "a.quantity as xsnum,a.net_weight as jingzhong,c.basic_cost as chengbengf,c.premium_cost as cankaogf,c.company_surcharge as chengbenfj," &
            "a.xiaodan_amount as danjia,a.xiao_amount as xiaoshou,c.sales_surcharge as yuanfujia,a.gold_price as kejia," &
            "a.sales_cost as xiaoshoukj,a.sales_surcharge as xsfujia,a.zhekou as zhekou,a.settlement as shishou,b.NAME AS daogou " &
            "FROM xipunum_erp_outbound AS a " &
            "INNER JOIN xipunum_erp_user AS b ON b.USER = a.shopping_guide " &
            "INNER JOIN xipunum_erp_store AS c ON c.poduct_code = a.poduct_code " &
            "INNER JOIN xipunum_erp_store_order AS d ON d.id = c.order_id " &
            查找信息工厂 &
            " INNER JOIN xipunum_erp_shop AS e ON e.poduct_code = a.poduct_code" &
            " LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = e.item_number AND e.item_number != ''" &
            " LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id AND e.item_number IS NOT NULL AND e.item_number != ''" &
            " LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = e.specification_id AND e.specification_id IS NOT NULL AND e.specification_id != ''" &
            " LEFT JOIN xipunum_erp_category AS g ON g.id = COALESCE ( e1.category_id, e2.category_id ) AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL" &
            " WHERE a.creationtime >= '" & 查找开始日期 & "' AND a.creationtime < '" & 查找结束日期 & "'" &
            " and a.shopping_guide='" & 局部_查找导购账户 & "' and a.pling='" & 局部_查找导购批零 & "'" &
            查找信息库房 & 查找信息规格 & " ORDER BY a.id DESC"

        Dim dtDetail As DataTable = ExecuteQuery(sqlDetail, MySQL_ReadReport)

        For i As Integer = 0 To dtDetail.Rows.Count - 1
            dgvReport.Rows.Add()
            Dim row As DataGridViewRow = dgvReport.Rows(i)
            row.Cells(0).Value = (i + 1).ToString()
            row.Cells(1).Value = SafeString(dtDetail.Rows(i)("bianma"))
            row.Cells(2).Value = SafeString(dtDetail.Rows(i)("mingcheng"))
            row.Cells(3).Value = SafeString(dtDetail.Rows(i)("kuanhao"))
            row.Cells(4).Value = SafeString(dtDetail.Rows(i)("pinlei"))
            row.Cells(5).Value = SafeString(dtDetail.Rows(i)("guige"))
            row.Cells(6).Value = SafeString(dtDetail.Rows(i)("caizhi"))
            row.Cells(7).Value = SafeString(dtDetail.Rows(i)("quankou"))
            row.Cells(8).Value = SafeString(dtDetail.Rows(i)("chengse"))
            row.Cells(9).Value = SafeDecimal(dtDetail.Rows(i)("xsnum")).ToString("F2")
            row.Cells(10).Value = SafeDecimal(dtDetail.Rows(i)("jingzhong")).ToString("F3")
            row.Cells(11).Value = SafeDecimal(dtDetail.Rows(i)("chengbengf")).ToString("F2")
            row.Cells(12).Value = SafeDecimal(dtDetail.Rows(i)("cankaogf")).ToString("F2")
            row.Cells(13).Value = SafeDecimal(dtDetail.Rows(i)("chengbenfj")).ToString("F2")
            row.Cells(14).Value = SafeDecimal(dtDetail.Rows(i)("danjia")).ToString("F2")
            row.Cells(15).Value = SafeDecimal(dtDetail.Rows(i)("xiaoshou")).ToString("F2")
            row.Cells(16).Value = SafeDecimal(dtDetail.Rows(i)("yuanfujia")).ToString("F2")
            row.Cells(17).Value = SafeDecimal(dtDetail.Rows(i)("kejia")).ToString("F2")
            row.Cells(18).Value = SafeDecimal(dtDetail.Rows(i)("xiaoshoukj")).ToString("F2")
            row.Cells(19).Value = SafeDecimal(dtDetail.Rows(i)("xsfujia")).ToString("F2")
            row.Cells(20).Value = SafeDecimal(dtDetail.Rows(i)("zhekou")).ToString("F2")
            row.Cells(21).Value = SafeDecimal(dtDetail.Rows(i)("shishou")).ToString("F2")
            row.Cells(22).Value = SafeString(dtDetail.Rows(i)("daogou"))
        Next

        ' 合计行
        Dim sqlSum As String = "SELECT CAST(ROUND(sum(a.quantity), 2) AS DECIMAL (20, 2)) AS xsnum," &
            "CAST(ROUND(sum(a.net_weight), 2) AS DECIMAL (20, 2)) AS jingzhong," &
            "CAST(ROUND(sum(a.xiao_amount), 2) AS DECIMAL (20, 2)) AS xiaoshou," &
            "CAST(ROUND(sum(a.settlement), 2) AS DECIMAL (20, 2)) AS shishou " &
            "FROM xipunum_erp_outbound AS a " &
            "INNER JOIN xipunum_erp_user AS b ON b.USER = a.shopping_guide " &
            "INNER JOIN xipunum_erp_store AS c ON c.poduct_code = a.poduct_code " &
            "INNER JOIN xipunum_erp_store_order AS d ON d.id = c.order_id " &
            查找信息工厂 &
            " INNER JOIN xipunum_erp_shop AS e ON e.poduct_code = a.poduct_code " &
            查找信息规格 &
            " LEFT JOIN xipunum_erp_category as f on f.id=e.category_id" &
            " LEFT JOIN xipunum_erp_specs as g on g.id=e.specification_id" &
            " WHERE a.creationtime >= '" & 查找开始日期 & "' AND a.creationtime < '" & 查找结束日期 & "'" &
            " and a.shopping_guide='" & 局部_查找导购账户 & "' and a.pling='" & 局部_查找导购批零 & "'" &
            查找信息库房 & " ORDER BY a.id DESC"
        Dim dtSum As DataTable = ExecuteQuery(sqlSum, MySQL_ReadReport)

        Dim sumRowIndex As Integer = dgvReport.Rows.Add()
        dgvReport.Rows(sumRowIndex).Cells(1).Value = "合计"
        If dtSum.Rows.Count > 0 Then
            dgvReport.Rows(sumRowIndex).Cells(9).Value = SafeDecimal(dtSum.Rows(0)("xsnum")).ToString("F2")
            dgvReport.Rows(sumRowIndex).Cells(10).Value = SafeDecimal(dtSum.Rows(0)("jingzhong")).ToString("F3")
            dgvReport.Rows(sumRowIndex).Cells(15).Value = SafeDecimal(dtSum.Rows(0)("xiaoshou")).ToString("F2")
            dgvReport.Rows(sumRowIndex).Cells(21).Value = SafeDecimal(dtSum.Rows(0)("shishou")).ToString("F2")
        End If
    End Sub

    ' ========== 导出按钮 ==========
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If dgvReport.Rows.Count = 0 Then
            ShowWarning("没有数据可导出！")
            Return
        End If
        Try
            ' 将DataGridView数据转为DataTable
            Dim dt As New DataTable()
            For Each col As DataGridViewColumn In dgvReport.Columns
                If col.Visible Then
                    dt.Columns.Add(col.HeaderText)
                End If
            Next
            For Each row As DataGridViewRow In dgvReport.Rows
                Dim dr As DataRow = dt.NewRow()
                Dim colIdx As Integer = 0
                For Each col As DataGridViewColumn In dgvReport.Columns
                    If col.Visible Then
                        dr(colIdx) = If(row.Cells(col.Index).Value, "")
                        colIdx += 1
                    End If
                Next
                dt.Rows.Add(dr)
            Next
            ExportToExcel(dt, "报表员工月销售统计表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 重置按钮 ==========
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Form_Load(Nothing, Nothing)
    End Sub

    ' ========== 单选框事件 ==========
    Private Sub rbWeight_CheckedChanged(sender As Object, e As EventArgs)
        If rbWeight.Checked Then
            rbAmount.Checked = False
            _子程序_加载表头()
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub rbAmount_CheckedChanged(sender As Object, e As EventArgs)
        If rbAmount.Checked Then
            rbWeight.Checked = False
            _子程序_加载表头()
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub rbAll_CheckedChanged(sender As Object, e As EventArgs)
        If rbAll.Checked Then
            rbGold.Checked = False
            rbCategory.Checked = False
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub rbGold_CheckedChanged(sender As Object, e As EventArgs)
        If rbGold.Checked Then
            rbAll.Checked = False
            rbCategory.Checked = False
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub rbCategory_CheckedChanged(sender As Object, e As EventArgs)
        If rbCategory.Checked Then
            rbAll.Checked = False
            rbGold.Checked = False
            btnQuery_Click(Nothing, Nothing)
        End If
    End Sub

    ' ========== 列表框事件 ==========
    Private Sub clbShop_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        ' 店铺选择变化时不需要特殊处理
    End Sub

    Private Sub clbCategory_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        ' 品类变化时刷新规格列表
        Me.BeginInvoke(New Action(Sub()
                                      _品类选中_id数据初始化()
                                  End Sub))
    End Sub

    Private Sub clbSpec_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        ' 规格变化时刷新规格id列表
        Me.BeginInvoke(New Action(Sub()
                                      _选中规格id数据初始化()
                                  End Sub))
    End Sub

    ' ========== 搜索框事件 ==========
    Private Sub txtSearch_Enter(sender As Object, e As EventArgs)
        If txtSearch.Text = "请输入查找信息" Then
            txtSearch.Text = ""
        End If
    End Sub

    Private Sub txtSearch_Leave(sender As Object, e As EventArgs)
        If txtSearch.Text = "" Then
            txtSearch.Text = "请输入查找信息"
        End If
    End Sub

    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtSearch.Text.Trim()) OrElse txtSearch.Text = "请输入查找信息" Then
                ShowWarning("查找工厂不能为空！")
                txtSearch.Text = ""
                Return
            End If
            _工厂查找_数据初始化()
            txtSearch.Text = ""
        End If
    End Sub

    ' ========== 右键菜单事件 ==========
    Private Sub dgvReport_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs)
        If e.Button = MouseButtons.Right AndAlso e.RowIndex > 0 AndAlso e.ColumnIndex > 0 Then
            Dim clickRow As Integer = If(局部_查找类型文本 = "导购数据明细", dgvReport.Rows.Count - 1, dgvReport.Rows.Count - 3)
            If e.RowIndex < clickRow Then
                局部_查找导购账户 = dgvReport.Rows(e.RowIndex).Cells(1).Value?.ToString()
                If rbWeight.Checked Then
                    局部_查找导购批零 = If(dgvReport.Columns.Count > 14, dgvReport.Rows(e.RowIndex).Cells(14).Value?.ToString(), "")
                Else
                    局部_查找导购批零 = If(dgvReport.Columns.Count > 16, dgvReport.Rows(e.RowIndex).Cells(16).Value?.ToString(), "")
                End If
                If Not String.IsNullOrEmpty(局部_查找导购账户) Then
                    dgvReport.Rows(e.RowIndex).Selected = True
                    cmsMenu.Show(dgvReport, e.Location)
                End If
            End If
        End If
    End Sub

    Private Sub tsmiDetail_Click(sender As Object, e As EventArgs)
        局部_查找类型文本 = "导购数据明细"
        _子程序_加载表头()
        btnQuery_Click(Nothing, Nothing)
    End Sub

    Private Sub tsmiReturn_Click(sender As Object, e As EventArgs)
        局部_查找类型文本 = "员工月销售统计"
        局部_查找导购账户 = ""
        _子程序_加载表头()
        btnQuery_Click(Nothing, Nothing)
    End Sub

End Class

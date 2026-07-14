' ============================================================================
' 报表员工绩效窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_报表员工绩效.form.e.txt
' 功能: 员工绩效统计，包含销售/回收数据加载、5档提成计算、绩效合计
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class EmployeePerformanceReportForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private 查找开始日期 As String = ""
    Private 查找结束日期 As String = ""
    Private 查找库房名称 As String = ""

    ' ========== 控件声明 ==========
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnReset As New Button()
    Private clbShop As New CheckedListBox()
    Private dgvReport As New DataGridView()

    ' ========== 内部数据表（隐藏的销售/回收数据表格） ==========
    Private dtSales As New DataTable()
    Private dtRecovery As New DataTable()

    ' ========== 初始化 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "员工绩效报表"
        Me.Size = New Drawing.Size(1366, 750)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' ========== 顶部面板 ==========
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

        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(410, 4)
        btnQuery.Size = New Drawing.Size(70, 28)
        panelTop.Controls.Add(btnQuery)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(485, 4)
        btnReset.Size = New Drawing.Size(70, 28)
        panelTop.Controls.Add(btnReset)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(560, 4)
        btnExport.Size = New Drawing.Size(70, 28)
        panelTop.Controls.Add(btnExport)

        ' ========== 左侧店铺选择 ==========
        Dim panelLeft As New Panel()
        panelLeft.Dock = DockStyle.Left
        panelLeft.Width = 200
        Me.Controls.Add(panelLeft)

        Dim lblShop As New Label()
        lblShop.Text = "店铺名称"
        lblShop.Location = New Drawing.Point(5, 5)
        lblShop.AutoSize = True
        panelLeft.Controls.Add(lblShop)

        clbShop.Location = New Drawing.Point(5, 25)
        clbShop.Size = New Drawing.Size(190, 300)
        clbShop.CheckOnClick = True
        panelLeft.Controls.Add(clbShop)

        ' ========== DataGridView ==========
        dgvReport.Dock = DockStyle.Fill
        dgvReport.ReadOnly = True
        dgvReport.AllowUserToAddRows = False
        dgvReport.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvReport)
        dgvReport.BringToFront()

        ' ========== 初始化内部数据表 ==========
        dtSales.Columns.Add("kufangid", GetType(String))
        dtSales.Columns.Add("gangweiid", GetType(String))
        dtSales.Columns.Add("pinleiid", GetType(String))
        dtSales.Columns.Add("daogou", GetType(String))
        dtSales.Columns.Add("pingling", GetType(String))
        dtSales.Columns.Add("xsshu", GetType(String))
        dtSales.Columns.Add("xszhong", GetType(String))
        dtSales.Columns.Add("xsjine", GetType(String))

        dtRecovery.Columns.Add("kufangid", GetType(String))
        dtRecovery.Columns.Add("gangweiid", GetType(String))
        dtRecovery.Columns.Add("pinleiid", GetType(String))
        dtRecovery.Columns.Add("daogou", GetType(String))
        dtRecovery.Columns.Add("pingling", GetType(String))
        dtRecovery.Columns.Add("hsshu", GetType(String))
        dtRecovery.Columns.Add("hszhong", GetType(String))
        dtRecovery.Columns.Add("hsjine", GetType(String))
    End Sub

    ' ========== 窗口创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        dtpStart.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        dtpEnd.Value = DateTime.Now
        _查询条件_初始化()
        _子程序_加载表头()
        _报表数据详情_基础()
        子程序_删除表格()
    End Sub

    ' ========== 子程序：加载表头 ==========
    Private Sub _子程序_加载表头()
        dgvReport.Columns.Clear()
        dgvReport.Rows.Clear()

        Dim headers As String() = {"序号", "店铺", "岗位", "员工", "批零", "库房id", "岗位id", "员工账户", "销售数量", "回收数量", "销售重量", "回收重量", "销售金额", "回收金额", "销售绩效", "回收绩效", "绩效合计"}
        Dim widths As Integer() = {65, 65, 65, 120, 65, 0, 0, 0, 80, 80, 80, 80, 80, 80, 80, 80, 80}
        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i.ToString()
            col.Width = widths(i)
            If widths(i) = 0 Then col.Visible = False
            dgvReport.Columns.Add(col)
        Next
    End Sub

    ' ========== 查询条件初始化 ==========
    Private Sub _查询条件_初始化()
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
    End Sub

    ' ========== 子程序：删除表格（清空内部数据表） ==========
    Private Sub 子程序_删除表格()
        dtSales.Clear()
        dtRecovery.Clear()
    End Sub

    ' ========== 查询按钮 ==========
    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        btnQuery.Enabled = False
        btnReset.Enabled = False
        btnExport.Enabled = False
        Me.Text = "员工绩效报表 时间:" & dtpStart.Value.ToString("yyyy-MM-dd") & "至" & dtpEnd.Value.ToString("yyyy-MM-dd")

        _报表数据详情_基础()
        _报表数据详情_展示()

        btnQuery.Enabled = True
        btnReset.Enabled = True
        btnExport.Enabled = True
    End Sub

    ' ========== 报表数据详情_基础 ==========
    Private Sub _报表数据详情_基础()
        ' 获取选中的店铺
        Dim 店铺名称权限 As String = ""
        Dim 店铺名称数量 As Integer = 0
        For i As Integer = 0 To clbShop.Items.Count - 1
            If clbShop.GetItemChecked(i) Then
                店铺名称权限 &= "'" & clbShop.Items(i).Tag.ToString() & "',"
                店铺名称数量 += 1
            End If
        Next
        If 店铺名称数量 = 0 Then
            ShowWarning("查找库房不能为空！")
            Return
        End If

        查找开始日期 = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
        查找结束日期 = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"
        查找库房名称 = 店铺名称权限.TrimEnd(","c)

        dgvReport.Rows.Clear()

        ' 查询员工信息
        Dim sqlEmp As String = "SELECT b.title AS kufang, c.title AS gangwei, a.USER AS USER, a.NAME AS NAME, a.department as kufangid, a.post as gangweiid FROM xipunum_erp_user AS a  INNER JOIN xipunum_erp_type AS b ON b.id = a.department INNER JOIN xipunum_erp_type AS c ON c.id = a.post WHERE a.department IN (" & 查找库房名称 & ") and a.state='0' ORDER BY a.department,a.post asc;"
        Dim dtEmp As DataTable = ExecuteQuery(sqlEmp, MySQL_ReadReport)

        For i As Integer = 0 To dtEmp.Rows.Count - 1
            Dim kufang As String = SafeString(dtEmp.Rows(i)("kufang"))
            Dim gangwei As String = SafeString(dtEmp.Rows(i)("gangwei"))
            Dim user As String = SafeString(dtEmp.Rows(i)("USER"))
            Dim name As String = SafeString(dtEmp.Rows(i)("NAME"))
            Dim kufangid As String = SafeString(dtEmp.Rows(i)("kufangid"))
            Dim gangweiid As String = SafeString(dtEmp.Rows(i)("gangweiid"))

            ' 零售行
            Dim rowIdx1 As Integer = dgvReport.Rows.Add()
            dgvReport.Rows(rowIdx1).Cells(0).Value = (i + 1).ToString()
            dgvReport.Rows(rowIdx1).Cells(1).Value = kufang
            dgvReport.Rows(rowIdx1).Cells(2).Value = gangwei
            dgvReport.Rows(rowIdx1).Cells(3).Value = user & "(" & name & ")"
            dgvReport.Rows(rowIdx1).Cells(4).Value = "零售"
            dgvReport.Rows(rowIdx1).Cells(5).Value = kufangid
            dgvReport.Rows(rowIdx1).Cells(6).Value = gangweiid
            dgvReport.Rows(rowIdx1).Cells(7).Value = user
            dgvReport.Rows(rowIdx1).Cells(8).Value = "0.00"
            dgvReport.Rows(rowIdx1).Cells(9).Value = "0.00"
            dgvReport.Rows(rowIdx1).Cells(10).Value = "0.000"
            dgvReport.Rows(rowIdx1).Cells(11).Value = "0.000"
            dgvReport.Rows(rowIdx1).Cells(12).Value = "0.00"
            dgvReport.Rows(rowIdx1).Cells(13).Value = "0.00"
            dgvReport.Rows(rowIdx1).Cells(14).Value = "0.00"
            dgvReport.Rows(rowIdx1).Cells(15).Value = "0.00"
            dgvReport.Rows(rowIdx1).Cells(16).Value = "0.00"

            ' 批发行
            Dim rowIdx2 As Integer = dgvReport.Rows.Add()
            dgvReport.Rows(rowIdx2).Cells(4).Value = "批发"
            dgvReport.Rows(rowIdx2).Cells(5).Value = kufangid
            dgvReport.Rows(rowIdx2).Cells(6).Value = gangweiid
            dgvReport.Rows(rowIdx2).Cells(7).Value = user
            dgvReport.Rows(rowIdx2).Cells(8).Value = "0.00"
            dgvReport.Rows(rowIdx2).Cells(9).Value = "0.00"
            dgvReport.Rows(rowIdx2).Cells(10).Value = "0.000"
            dgvReport.Rows(rowIdx2).Cells(11).Value = "0.000"
            dgvReport.Rows(rowIdx2).Cells(12).Value = "0.00"
            dgvReport.Rows(rowIdx2).Cells(13).Value = "0.00"
            dgvReport.Rows(rowIdx2).Cells(14).Value = "0.00"
            dgvReport.Rows(rowIdx2).Cells(15).Value = "0.00"
            dgvReport.Rows(rowIdx2).Cells(16).Value = "0.00"
        Next
    End Sub

    ' ========== 报表数据详情_展示 ==========
    Private Sub _报表数据详情_展示()
        子程序_删除表格()

        ' ========== 加载销售数据 ==========
        Dim sqlSales As String = "SELECT a.kufang as dianpuid, h.post as gangweiid, CASE WHEN COALESCE ( d.id, '' ) = '' THEN '0' ELSE d.id END AS pinleiid, a.shopping_guide AS daogou, a.pling as pingling, sum(COALESCE(a.quantity, 0)) as xsshu, sum(COALESCE(a.net_weight, 0)) as xszhong, sum(COALESCE(a.settlement, 0)) as xsjine FROM xipunum_erp_outbound AS a  INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_ksiamges AS c ON c.kuanhao = b.item_number AND b.item_number != '' LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = c.specification_id AND b.item_number IS NOT NULL AND b.item_number != '' LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id AND b.specification_id IS NOT NULL AND b.specification_id != '' LEFT JOIN xipunum_erp_category AS d ON d.id = COALESCE ( e1.category_id, e2.category_id ) AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL  INNER JOIN xipunum_erp_shop AS e ON e.poduct_code = a.poduct_code INNER JOIN xipunum_erp_user AS h ON h.user = a.shopping_guide INNER JOIN xipunum_erp_type AS i ON i.id = a.kufang INNER JOIN xipunum_erp_type AS j ON j.id = h.post WHERE a.creationtime >= '" & 查找开始日期 & "'  AND a.creationtime < '" & 查找结束日期 & "' and a.kufang in (" & 查找库房名称 & ") GROUP BY a.kufang,h.post,CASE WHEN COALESCE ( d.id, '' ) = '' THEN '0' ELSE d.id END,a.shopping_guide,a.pling ORDER BY a.kufang,h.post,a.pling asc;"
        Dim dtSalesData As DataTable = ExecuteQuery(sqlSales, MySQL_ReadReport)
        For Each dr As DataRow In dtSalesData.Rows
            Dim newRow As DataRow = dtSales.NewRow()
            newRow("kufangid") = SafeString(dr("dianpuid"))
            newRow("gangweiid") = SafeString(dr("gangweiid"))
            newRow("pinleiid") = SafeString(dr("pinleiid"))
            newRow("daogou") = SafeString(dr("daogou"))
            newRow("pingling") = SafeString(dr("pingling"))
            newRow("xsshu") = SafeDecimal(dr("xsshu")).ToString("F2")
            newRow("xszhong") = SafeDecimal(dr("xszhong")).ToString("F3")
            newRow("xsjine") = SafeDecimal(dr("xsjine")).ToString("F2")
            dtSales.Rows.Add(newRow)
        Next

        ' ========== 加载回收数据 ==========
        Dim sqlRecovery As String = "SELECT h.department AS kufangid, h.post AS gangwei, b.category_id as pinleiid, a.shopping_guide as daogou, CASE WHEN d.pling = '批发' THEN d.pling ELSE '零售' END AS pingling, sum(COALESCE(a.quantity, 0)) as hsshu, sum(COALESCE(a.jin_zhong, 0)) as hszhong, sum(COALESCE(a.retreat_amount, 0)) as hsjine FROM xipunum_erp_retreat AS a INNER JOIN xipunum_erp_retreat_title AS b ON b.id = a.product_name INNER JOIN xipunum_erp_retreat_order AS c ON c.id = a.order_id left JOIN xipunum_erp_outbound_order AS d ON d.retrea_umber = c.retrea_umber INNER JOIN xipunum_erp_user AS h ON h.USER = a.shopping_guide  WHERE a.creationtime >= '" & 查找开始日期 & "' AND a.creationtime < '" & 查找结束日期 & "' AND h.department IN (" & 查找库房名称 & ") GROUP BY h.department,h.post,b.category_id,a.shopping_guide,CASE WHEN d.pling = '批发' THEN d.pling ELSE '零售' END ORDER BY h.department,h.post,CASE WHEN d.pling = '批发' THEN d.pling ELSE '零售' END asc;"
        Dim dtRecoveryData As DataTable = ExecuteQuery(sqlRecovery, MySQL_ReadReport)
        For Each dr As DataRow In dtRecoveryData.Rows
            Dim newRow As DataRow = dtRecovery.NewRow()
            newRow("kufangid") = SafeString(dr("kufangid"))
            newRow("gangweiid") = SafeString(dr("gangwei"))
            newRow("pinleiid") = SafeString(dr("pinleiid"))
            newRow("daogou") = SafeString(dr("daogou"))
            newRow("pingling") = SafeString(dr("pingling"))
            newRow("hsshu") = SafeDecimal(dr("hsshu")).ToString("F2")
            newRow("hszhong") = SafeDecimal(dr("hszhong")).ToString("F3")
            newRow("hsjine") = SafeDecimal(dr("hsjine")).ToString("F2")
            dtRecovery.Rows.Add(newRow)
        Next

        ' ========== 计算员工绩效 ==========
        Dim 计算员工数量 As Integer = dgvReport.Rows.Count
        For 计算员工计次 As Integer = 0 To 计算员工数量 - 1
            Dim 计算员工店铺id As String = SafeString(dgvReport.Rows(计算员工计次).Cells(5).Value)
            Dim 计算员工岗位id As String = SafeString(dgvReport.Rows(计算员工计次).Cells(6).Value)
            Dim 计算员工员工账户 As String = SafeString(dgvReport.Rows(计算员工计次).Cells(7).Value)
            Dim 计算员工批零 As String = SafeString(dgvReport.Rows(计算员工计次).Cells(4).Value)
            Dim 计算员工批零id As String = If(计算员工批零 = "批发", "1", "0")

            Dim 销售提成合计销数量 As Decimal = 0
            Dim 销售提成合计销重量 As Decimal = 0
            Dim 销售提成合计销金额 As Decimal = 0
            Dim 回收提成合计销数量 As Decimal = 0
            Dim 回收提成合计销重量 As Decimal = 0
            Dim 回收提成合计销金额 As Decimal = 0
            Dim 销售提成合计绩效金额 As Decimal = 0
            Dim 回收提成合计绩效金额 As Decimal = 0

            ' 查询岗位计算范围
            Dim 岗位计算范围 As String = ""
            Dim sqlRange As String = "SELECT fanwei FROM xipunum_erp_category_score WHERE piling='" & 计算员工批零id & "' and gangweiid='" & 计算员工岗位id & "' GROUP BY fanwei ORDER BY id asc LIMIT 1;"
            Dim dtRange As DataTable = ExecuteQuery(sqlRange, MySQL_ReadReport)
            If dtRange.Rows.Count > 0 Then
                岗位计算范围 = SafeString(dtRange.Rows(0)("fanwei"))
            End If

            ' ========== 销售数据计算 ==========
            For Each drSale As DataRow In dtSales.Rows
                Dim 销售数据计算库房id As String = SafeString(drSale("kufangid"))
                Dim 销售数据计算岗位id As String = SafeString(drSale("gangweiid"))
                Dim 销售数据计算品类id As String = SafeString(drSale("pinleiid"))
                Dim 销售数据计算账户 As String = SafeString(drSale("daogou"))
                Dim 销售数据计算批零 As String = SafeString(drSale("pingling"))
                Dim 销售数据提成销数量 As Decimal = SafeDecimal(drSale("xsshu"))
                Dim 销售数据提成销重量 As Decimal = SafeDecimal(drSale("xszhong"))
                Dim 销售数据提成销金额 As Decimal = SafeDecimal(drSale("xsjine"))
                Dim 销售数据提成绩效 As Decimal = 0

                ' 累加销售合计
                If 计算员工店铺id = 销售数据计算库房id AndAlso 计算员工岗位id = 销售数据计算岗位id AndAlso 计算员工员工账户 = 销售数据计算账户 AndAlso 计算员工批零 = 销售数据计算批零 Then
                    销售提成合计销数量 += 销售数据提成销数量
                    销售提成合计销重量 += 销售数据提成销重量
                    销售提成合计销金额 += 销售数据提成销金额
                End If

                ' 查询岗位品类提成比例
                Dim sqlScore As String = "SELECT categoryid AS categoryid, fanwei AS fanwei, danwei AS danwei, data1 AS data1, data1num AS data1num, CASE WHEN data2 = '0.000' THEN data1 ELSE data2 END AS data2, CASE WHEN data2 = '0.000' THEN data1num ELSE data2num END AS data2num, CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1 ELSE data2 END ELSE data3 END AS data3, CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1num ELSE data2num END ELSE data3num END AS data3num, CASE WHEN data4 = '0.000' THEN CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1 ELSE data2 END ELSE data3 END ELSE data4 END AS data4, CASE WHEN data4 = '0.000' THEN CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1num ELSE data2num END ELSE data3num END ELSE data4num END AS data4num, CASE WHEN data5 = '0.000' THEN CASE WHEN data4 = '0.000' THEN CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1 ELSE data2 END ELSE data3 END ELSE data4 END ELSE data5 END AS data5, CASE WHEN data5 = '0.000' THEN CASE WHEN data4 = '0.000' THEN CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1num ELSE data2num END ELSE data3num END ELSE data4num END ELSE data5num END AS data5num FROM xipunum_erp_category_score WHERE piling='" & 计算员工批零id & "' and gangweiid='" & 计算员工岗位id & "' and categoryid='" & 销售数据计算品类id & "' GROUP BY categoryid ORDER BY id asc;"
                Dim dtScore As DataTable = ExecuteQuery(sqlScore, MySQL_ReadReport)
                Dim 岗位提成计算字段 As String = ""
                Dim 档位(4) As Decimal
                Dim 档位值(4) As Decimal
                If dtScore.Rows.Count > 0 Then
                    岗位提成计算字段 = SafeString(dtScore.Rows(0)("danwei"))
                    档位(0) = SafeDecimal(dtScore.Rows(0)("data1"))
                    档位值(0) = SafeDecimal(dtScore.Rows(0)("data1num"))
                    档位(1) = SafeDecimal(dtScore.Rows(0)("data2"))
                    档位值(1) = SafeDecimal(dtScore.Rows(0)("data2num"))
                    档位(2) = SafeDecimal(dtScore.Rows(0)("data3"))
                    档位值(2) = SafeDecimal(dtScore.Rows(0)("data3num"))
                    档位(3) = SafeDecimal(dtScore.Rows(0)("data4"))
                    档位值(3) = SafeDecimal(dtScore.Rows(0)("data4num"))
                    档位(4) = SafeDecimal(dtScore.Rows(0)("data5"))
                    档位值(4) = SafeDecimal(dtScore.Rows(0)("data5num"))
                End If

                ' 确定提成计算字段
                Select Case 岗位提成计算字段
                    Case "0" : 销售数据提成绩效 = 销售数据提成销数量
                    Case "1" : 销售数据提成绩效 = 销售数据提成销重量
                    Case "2" : 销售数据提成绩效 = 销售数据提成销金额
                End Select

                ' 计算5档绩效金额
                Dim 绩效金额 As Decimal = CalculateTieredCommission(销售数据提成绩效, 岗位提成计算字段, 档位, 档位值)

                ' 根据计算范围累加绩效
                Dim match As Boolean = False
                Select Case 岗位计算范围
                    Case "0"
                        match = (计算员工店铺id = 销售数据计算库房id AndAlso 计算员工批零 = 销售数据计算批零)
                    Case "1"
                        match = (计算员工店铺id = 销售数据计算库房id AndAlso 计算员工岗位id = 销售数据计算岗位id AndAlso 计算员工批零 = 销售数据计算批零)
                    Case "2"
                        match = (计算员工店铺id = 销售数据计算库房id AndAlso 计算员工岗位id = 销售数据计算岗位id AndAlso 计算员工员工账户 = 销售数据计算账户 AndAlso 计算员工批零 = 销售数据计算批零)
                End Select
                If match Then
                    销售提成合计绩效金额 += 绩效金额
                End If
            Next

            ' ========== 回收数据计算 ==========
            For Each drRec As DataRow In dtRecovery.Rows
                Dim 回收数据计算库房id As String = SafeString(drRec("kufangid"))
                Dim 回收数据计算岗位id As String = SafeString(drRec("gangweiid"))
                Dim 回收数据计算品类id As String = SafeString(drRec("pinleiid"))
                Dim 回收数据计算账户 As String = SafeString(drRec("daogou"))
                Dim 回收数据计算批零 As String = SafeString(drRec("pingling"))
                Dim 回收数据提成销数量 As Decimal = SafeDecimal(drRec("hsshu"))
                Dim 回收数据提成销重量 As Decimal = SafeDecimal(drRec("hszhong"))
                Dim 回收数据提成销金额 As Decimal = SafeDecimal(drRec("hsjine"))
                Dim 回收数据提成绩效 As Decimal = 0

                ' 累加回收合计
                If 计算员工店铺id = 回收数据计算库房id AndAlso 计算员工岗位id = 回收数据计算岗位id AndAlso 计算员工员工账户 = 回收数据计算账户 AndAlso 计算员工批零 = 回收数据计算批零 Then
                    回收提成合计销数量 += 回收数据提成销数量
                    回收提成合计销重量 += 回收数据提成销重量
                    回收提成合计销金额 += 回收数据提成销金额
                End If

                ' 查询岗位品类提成比例
                Dim sqlScore As String = "SELECT categoryid AS categoryid, fanwei AS fanwei, danwei AS danwei, data1 AS data1, data1num AS data1num, CASE WHEN data2 = '0.000' THEN data1 ELSE data2 END AS data2, CASE WHEN data2 = '0.000' THEN data1num ELSE data2num END AS data2num, CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1 ELSE data2 END ELSE data3 END AS data3, CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1num ELSE data2num END ELSE data3num END AS data3num, CASE WHEN data4 = '0.000' THEN CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1 ELSE data2 END ELSE data3 END ELSE data4 END AS data4, CASE WHEN data4 = '0.000' THEN CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1num ELSE data2num END ELSE data3num END ELSE data4num END AS data4num, CASE WHEN data5 = '0.000' THEN CASE WHEN data4 = '0.000' THEN CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1 ELSE data2 END ELSE data3 END ELSE data4 END ELSE data5 END AS data5, CASE WHEN data5 = '0.000' THEN CASE WHEN data4 = '0.000' THEN CASE WHEN data3 = '0.000' THEN CASE WHEN data2 = '0.000' THEN data1num ELSE data2num END ELSE data3num END ELSE data4num END ELSE data5num END AS data5num FROM xipunum_erp_category_score WHERE piling='" & 计算员工批零id & "' and gangweiid='" & 计算员工岗位id & "' and categoryid='" & 回收数据计算品类id & "' GROUP BY categoryid ORDER BY id asc;"
                Dim dtScore As DataTable = ExecuteQuery(sqlScore, MySQL_ReadReport)
                Dim 岗位提成计算字段 As String = ""
                Dim 档位(4) As Decimal
                Dim 档位值(4) As Decimal
                If dtScore.Rows.Count > 0 Then
                    岗位提成计算字段 = SafeString(dtScore.Rows(0)("danwei"))
                    档位(0) = SafeDecimal(dtScore.Rows(0)("data1"))
                    档位值(0) = SafeDecimal(dtScore.Rows(0)("data1num"))
                    档位(1) = SafeDecimal(dtScore.Rows(0)("data2"))
                    档位值(1) = SafeDecimal(dtScore.Rows(0)("data2num"))
                    档位(2) = SafeDecimal(dtScore.Rows(0)("data3"))
                    档位值(2) = SafeDecimal(dtScore.Rows(0)("data3num"))
                    档位(3) = SafeDecimal(dtScore.Rows(0)("data4"))
                    档位值(3) = SafeDecimal(dtScore.Rows(0)("data4num"))
                    档位(4) = SafeDecimal(dtScore.Rows(0)("data5"))
                    档位值(4) = SafeDecimal(dtScore.Rows(0)("data5num"))
                End If

                ' 确定提成计算字段
                Select Case 岗位提成计算字段
                    Case "0" : 回收数据提成绩效 = 回收数据提成销数量
                    Case "1" : 回收数据提成绩效 = 回收数据提成销重量
                    Case "2" : 回收数据提成绩效 = 回收数据提成销金额
                End Select

                ' 计算5档绩效金额
                Dim 绩效金额 As Decimal = CalculateTieredCommission(回收数据提成绩效, 岗位提成计算字段, 档位, 档位值)

                ' 根据计算范围累加绩效
                Dim match As Boolean = False
                Select Case 岗位计算范围
                    Case "0"
                        match = (计算员工店铺id = 回收数据计算库房id AndAlso 计算员工批零 = 回收数据计算批零)
                    Case "1"
                        match = (计算员工店铺id = 回收数据计算库房id AndAlso 计算员工岗位id = 回收数据计算岗位id AndAlso 计算员工批零 = 回收数据计算批零)
                    Case "2"
                        match = (计算员工店铺id = 回收数据计算库房id AndAlso 计算员工岗位id = 回收数据计算岗位id AndAlso 计算员工员工账户 = 回收数据计算账户 AndAlso 计算员工批零 = 回收数据计算批零)
                End Select
                If match Then
                    回收提成合计绩效金额 += 绩效金额
                End If
            Next

            ' ========== 写入结果 ==========
            Dim 提成合计绩效金额 As Decimal = 销售提成合计绩效金额 + 回收提成合计绩效金额

            dgvReport.Rows(计算员工计次).Cells(8).Value = 销售提成合计销数量.ToString("F2")
            dgvReport.Rows(计算员工计次).Cells(9).Value = 回收提成合计销数量.ToString("F2")
            dgvReport.Rows(计算员工计次).Cells(10).Value = 销售提成合计销重量.ToString("F3")
            dgvReport.Rows(计算员工计次).Cells(11).Value = 回收提成合计销重量.ToString("F3")
            dgvReport.Rows(计算员工计次).Cells(12).Value = 销售提成合计销金额.ToString("F2")
            dgvReport.Rows(计算员工计次).Cells(13).Value = 回收提成合计销金额.ToString("F2")
            dgvReport.Rows(计算员工计次).Cells(14).Value = 销售提成合计绩效金额.ToString("F2")
            dgvReport.Rows(计算员工计次).Cells(15).Value = 回收提成合计绩效金额.ToString("F2")
            dgvReport.Rows(计算员工计次).Cells(16).Value = 提成合计绩效金额.ToString("F2")
        Next
    End Sub

    ' ========== 计算5档阶梯提成 ==========
    Private Function CalculateTieredCommission(绩效值 As Decimal, 计算字段 As String, 档位() As Decimal, 档位值() As Decimal) As Decimal
        Dim 档1 As Decimal = 0, 档2 As Decimal = 0, 档3 As Decimal = 0, 档4 As Decimal = 0, 档5 As Decimal = 0
        Dim isPercent As Boolean = (计算字段 = "2")

        ' 第1档：绩效值 > 档位(0) 且 绩效值 <= 档位(1)
        If 绩效值 <= 档位(1) AndAlso 绩效值 > 档位(0) Then
            If isPercent Then
                档1 = (绩效值 - 档位(0)) * 档位值(0) / 100
            Else
                档1 = (绩效值 - 档位(0)) * 档位值(0)
            End If
        End If

        ' 第2档
        If 绩效值 <= 档位(2) AndAlso 绩效值 > 档位(1) Then
            If isPercent Then
                档1 = (档位(1) - 档位(0)) * 档位值(0) / 100
                档2 = (绩效值 - 档位(1)) * 档位值(1) / 100
            Else
                档1 = (档位(1) - 档位(0)) * 档位值(0)
                档2 = (绩效值 - 档位(1)) * 档位值(1)
            End If
        End If

        ' 第3档
        If 绩效值 <= 档位(3) AndAlso 绩效值 > 档位(2) Then
            If isPercent Then
                档1 = (档位(1) - 档位(0)) * 档位值(0) / 100
                档2 = (档位(2) - 档位(1)) * 档位值(1) / 100
                档3 = (绩效值 - 档位(2)) * 档位值(2) / 100
            Else
                档1 = (档位(1) - 档位(0)) * 档位值(0)
                档2 = (档位(2) - 档位(1)) * 档位值(1)
                档3 = (绩效值 - 档位(2)) * 档位值(2)
            End If
        End If

        ' 第4档
        If 绩效值 <= 档位(4) AndAlso 绩效值 > 档位(3) Then
            If isPercent Then
                档1 = (档位(1) - 档位(0)) * 档位值(0) / 100
                档2 = (档位(2) - 档位(1)) * 档位值(1) / 100
                档3 = (档位(3) - 档位(2)) * 档位值(2) / 100
                档4 = (绩效值 - 档位(3)) * 档位值(3) / 100
            Else
                档1 = (档位(1) - 档位(0)) * 档位值(0)
                档2 = (档位(2) - 档位(1)) * 档位值(1)
                档3 = (档位(3) - 档位(2)) * 档位值(2)
                档4 = (绩效值 - 档位(3)) * 档位值(3)
            End If
        End If

        ' 第5档：绩效值 > 档位(4)
        If 绩效值 > 档位(4) Then
            If isPercent Then
                档1 = (档位(1) - 档位(0)) * 档位值(0) / 100
                档2 = (档位(2) - 档位(1)) * 档位值(1) / 100
                档3 = (档位(3) - 档位(2)) * 档位值(2) / 100
                档4 = (档位(4) - 档位(3)) * 档位值(3) / 100
                档5 = (绩效值 - 档位(4)) * 档位值(4) / 100
            Else
                档1 = (档位(1) - 档位(0)) * 档位值(0)
                档2 = (档位(2) - 档位(1)) * 档位值(1)
                档3 = (档位(3) - 档位(2)) * 档位值(2)
                档4 = (档位(4) - 档位(3)) * 档位值(3)
                档5 = (绩效值 - 档位(4)) * 档位值(4)
            End If
        End If

        Return 档1 + 档2 + 档3 + 档4 + 档5
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
            ExportToExcel(dt, "销售查询报表表.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 重置按钮 ==========
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Form_Load(Nothing, Nothing)
    End Sub

End Class

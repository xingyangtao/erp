' ============================================================================
' 商品入库详情窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品入库详情.form.e.txt
' 包含所有6个程序集变量、15个子程序、复杂5表联查SQL、事务保存、标签打印
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.IO

Public Class InboundDetailForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（6个） ==========
    Private currentRow As Integer = -1              ' 集_行号
    Private currentCol As Integer = -1              ' 集_列号
    Private orderStatus As String = ""              ' 局部_入库订单状态
    Private deletedCodes As String = ""             ' 局部_删除编码文本
    Private productCount As Integer = 0             ' 局部_商品入库数量
    Private deleteBtn As New Button()              ' 删除按钮

    ' ========== 控件声明 ==========
    Private dgvDetail As New DataGridView()         ' 高级表格1
    Private txtOrderNumber As New TextBox()         ' 单据号_编辑框
    Private txtDeliveryNumber As New TextBox()      ' 送货单号_编辑框
    Private txtRemarks As New TextBox()             ' 备注_编辑框
    Private txtGoldPrice As New TextBox()           ' 原料价_编辑框
    Private txtTotalWeight As New TextBox()         ' 总重_编辑框
    Private cmbFactory As New ComboBox()            ' 工厂名称组合框
    Private cmbSource As New ComboBox()             ' 来源组合框
    Private cmbSettlement As New ComboBox()         ' 结算方式组合框
    Private cmbLabelStyle As New ComboBox()         ' 组合框标签样式
    Private rbInlaid As New RadioButton()           ' 单选框_镶嵌
    Private rbNotInlaid As New RadioButton()        ' 单选框_非镶嵌
    Private toolStrip As New ToolStrip()            ' 工具条_通用
    Private panelTop As New Panel()                 ' 外形框_头部

    ' 表格列索引常量（待审36列，已审35列）
    Private Const COL_SEQ As Integer = 0            ' 序号/选择
    Private Const COL_CODE As Integer = 1           ' 商品编码
    Private Const COL_NAME As Integer = 2           ' 商品名称
    Private Const COL_CATEGORY As Integer = 3       ' 商品品类
    Private Const COL_SPEC As Integer = 4           ' 规格
    Private Const COL_KUANHAO As Integer = 5        ' 款号
    Private Const COL_CAIZHI As Integer = 6         ' 材质
    Private Const COL_QUANDU As Integer = 7         ' 圈口长度
    Private Const COL_WIDTH As Integer = 8          ' 面宽
    Private Const COL_THICKNESS As Integer = 9      ' 厚度
    Private Const COL_FACTORY_COND As Integer = 10  ' 工厂成色
    Private Const COL_COMPANY_COND As Integer = 11  ' 公司成色
    Private Const COL_SINGLE As Integer = 12        ' 单件重
    Private Const COL_QTY As Integer = 13           ' 数量
    Private Const COL_STOCK As Integer = 14         ' 库存
    Private Const COL_JINZHONG As Integer = 15      ' 金重
    Private Const COL_LOSS As Integer = 16          ' 损耗
    Private Const COL_INCLUDING As Integer = 17     ' 含耗重
    Private Const COL_WEIGHT As Integer = 18        ' 重量
    Private Const COL_UNIT As Integer = 19          ' 单位
    Private Const COL_SHITOU As Integer = 20        ' 石重
    Private Const COL_STNUM As Integer = 21         ' 石头数
    Private Const COL_SHITOU1 As Integer = 22       ' 副石重
    Private Const COL_SHNUM1 As Integer = 23        ' 副石头数
    Private Const COL_COST_PRICE As Integer = 24    ' 成本单价
    Private Const COL_COEFFICIENT As Integer = 25   ' 系数
    Private Const COL_BASIC_COST As Integer = 26    ' 成本工费
    Private Const COL_PREMIUM_COST As Integer = 27  ' 参考工费
    Private Const COL_SALES_COST As Integer = 28    ' 销售工费
    Private Const COL_COMPANY_SUR As Integer = 29   ' 成本附加费
    Private Const COL_SALES_SUR As Integer = 30     ' 销售附加费
    Private Const COL_SALES_PRICE As Integer = 31   ' 销售价
    Private Const COL_KUFANG As Integer = 32        ' 库房
    Private Const COL_REMARKS As Integer = 33       ' 备注
    Private Const COL_ZHUSE As Integer = 34         ' 主石色
    Private Const COL_DELETE As Integer = 35        ' 操作（待审模式才有）

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler dgvDetail.CellEndEdit, AddressOf DgvDetail_CellEndEdit
        AddHandler dgvDetail.CellClick, AddressOf DgvDetail_CellClick
        AddHandler dgvDetail.SelectionChanged, AddressOf DgvDetail_SelectionChanged
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "商品入库详情"
        Me.Size = New Drawing.Size(1427, 664)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' 工具条
        toolStrip.Dock = DockStyle.Top
        toolStrip.Height = 45
        Me.Controls.Add(toolStrip)

        ' 顶部信息区
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 145
        Me.Controls.Add(panelTop)

        ' 单据号
        Dim lblOrder As New Label()
        lblOrder.Text = "单据号："
        lblOrder.Location = New Drawing.Point(10, 12)
        lblOrder.AutoSize = True
        panelTop.Controls.Add(lblOrder)
        txtOrderNumber.Location = New Drawing.Point(60, 9)
        txtOrderNumber.Size = New Drawing.Size(150, 25)
        txtOrderNumber.ReadOnly = True
        panelTop.Controls.Add(txtOrderNumber)

        ' 送货单号
        Dim lblDelivery As New Label()
        lblDelivery.Text = "送货单号："
        lblDelivery.Location = New Drawing.Point(220, 12)
        lblDelivery.AutoSize = True
        panelTop.Controls.Add(lblDelivery)
        txtDeliveryNumber.Location = New Drawing.Point(290, 9)
        txtDeliveryNumber.Size = New Drawing.Size(150, 25)
        txtDeliveryNumber.ReadOnly = True
        panelTop.Controls.Add(txtDeliveryNumber)

        ' 原料价
        Dim lblGoldPrice As New Label()
        lblGoldPrice.Text = "原料价："
        lblGoldPrice.Location = New Drawing.Point(450, 12)
        lblGoldPrice.AutoSize = True
        panelTop.Controls.Add(lblGoldPrice)
        txtGoldPrice.Location = New Drawing.Point(500, 9)
        txtGoldPrice.Size = New Drawing.Size(80, 25)
        txtGoldPrice.ReadOnly = True
        panelTop.Controls.Add(txtGoldPrice)

        ' 总重
        Dim lblTotal As New Label()
        lblTotal.Text = "总重："
        lblTotal.Location = New Drawing.Point(590, 12)
        lblTotal.AutoSize = True
        panelTop.Controls.Add(lblTotal)
        txtTotalWeight.Location = New Drawing.Point(630, 9)
        txtTotalWeight.Size = New Drawing.Size(80, 25)
        txtTotalWeight.ReadOnly = True
        panelTop.Controls.Add(txtTotalWeight)

        ' 标签样式
        Dim lblLabel As New Label()
        lblLabel.Text = "标签样式："
        lblLabel.Location = New Drawing.Point(720, 12)
        lblLabel.AutoSize = True
        panelTop.Controls.Add(lblLabel)
        cmbLabelStyle.Location = New Drawing.Point(780, 9)
        cmbLabelStyle.Size = New Drawing.Size(126, 25)
        panelTop.Controls.Add(cmbLabelStyle)

        ' 工厂
        Dim lblFactory As New Label()
        lblFactory.Text = "工厂："
        lblFactory.Location = New Drawing.Point(10, 45)
        lblFactory.AutoSize = True
        panelTop.Controls.Add(lblFactory)
        cmbFactory.Location = New Drawing.Point(60, 42)
        cmbFactory.Size = New Drawing.Size(200, 25)
        cmbFactory.DropDownStyle = ComboBoxStyle.DropDownList
        panelTop.Controls.Add(cmbFactory)

        ' 来源
        Dim lblSource As New Label()
        lblSource.Text = "来源："
        lblSource.Location = New Drawing.Point(270, 45)
        lblSource.AutoSize = True
        panelTop.Controls.Add(lblSource)
        cmbSource.Location = New Drawing.Point(310, 42)
        cmbSource.Size = New Drawing.Size(120, 25)
        cmbSource.DropDownStyle = ComboBoxStyle.DropDownList
        panelTop.Controls.Add(cmbSource)

        ' 结算方式
        Dim lblSettlement As New Label()
        lblSettlement.Text = "结算方式："
        lblSettlement.Location = New Drawing.Point(440, 45)
        lblSettlement.AutoSize = True
        panelTop.Controls.Add(lblSettlement)
        cmbSettlement.Location = New Drawing.Point(510, 42)
        cmbSettlement.Size = New Drawing.Size(120, 25)
        cmbSettlement.DropDownStyle = ComboBoxStyle.DropDownList
        panelTop.Controls.Add(cmbSettlement)

        ' 镶嵌/非镶嵌
        rbInlaid.Text = "镶嵌"
        rbInlaid.Location = New Drawing.Point(640, 42)
        rbInlaid.AutoSize = True
        panelTop.Controls.Add(rbInlaid)
        rbNotInlaid.Text = "非镶嵌"
        rbNotInlaid.Location = New Drawing.Point(700, 42)
        rbNotInlaid.AutoSize = True
        panelTop.Controls.Add(rbNotInlaid)

        ' 备注
        Dim lblRemarks As New Label()
        lblRemarks.Text = "备注："
        lblRemarks.Location = New Drawing.Point(10, 80)
        lblRemarks.AutoSize = True
        panelTop.Controls.Add(lblRemarks)
        txtRemarks.Location = New Drawing.Point(60, 77)
        txtRemarks.Size = New Drawing.Size(400, 25)
        panelTop.Controls.Add(txtRemarks)

        ' DataGridView
        dgvDetail.Dock = DockStyle.Fill
        dgvDetail.AllowUserToAddRows = False
        dgvDetail.RowHeadersVisible = False
        dgvDetail.BackgroundColor = System.Drawing.Color.White
        Me.Controls.Add(dgvDetail)
    End Sub

    ' ========== _窗口_商品入库详情_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        deletedCodes = ""
        productCount = 0

        If InboundOrderNumber <> "" Then
            txtOrderNumber.Text = InboundOrderNumber
        End If

        ' 加载标签模板文件列表
        cmbLabelStyle.Items.Clear()
        Dim labelDir As String = Path.Combine(Application.StartupPath, "voucher\biaoqian")
        If Directory.Exists(labelDir) Then
            Dim files() As String = Directory.GetFiles(labelDir, "*.qdf")
            For Each f As String In files
                cmbLabelStyle.Items.Add(Path.GetFileName(f))
            Next
        End If
        If cmbLabelStyle.Items.Count > 0 Then
            cmbLabelStyle.SelectedIndex = 0
        End If

        ' 获取订单ID
        Dim orderId As String = GetOrderID()
        If orderId = "" Then
            ShowError("未找到入库订单数据！")
            Return
        End If

        ' 读取订单状态
        orderStatus = ReadOrderStatus(orderId)

        ' 加载表头和数据
        LoadTableHeader()
        LoadTableData()

        ' 根据状态设置工具栏和表格
        SetupUIByStatus()
    End Sub

    ' ========== 根据订单状态设置UI ==========
    Private Sub SetupUIByStatus()
        toolStrip.Items.Clear()
        If orderStatus = "待审" Then
            ' 待审：保存、标签打印、提取编码、批量修改
            AddToolStripButton("保存")
            AddToolStripButton("标签打印")
            AddToolStripButton("提取编码")
            AddToolStripButton("批量修改")
            dgvDetail.SelectionMode = DataGridViewSelectionMode.CellSelect
            txtRemarks.ReadOnly = False
        Else
            ' 已审：标签打印、提取编码、全选、反选
            AddToolStripButton("标签打印")
            AddToolStripButton("提取编码")
            AddToolStripButton("全选")
            AddToolStripButton("反选")
            dgvDetail.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            txtRemarks.ReadOnly = True
        End If
    End Sub

    Private Sub AddToolStripButton(text As String)
        Dim btn As New ToolStripButton(text)
        AddHandler btn.Click, AddressOf ToolStripButton_Click
        toolStrip.Items.Add(btn)
    End Sub

    ' ========== _工具条_通用_被单击 ==========
    Private Sub ToolStripButton_Click(sender As Object, e As EventArgs)
        Dim btnName As String = DirectCast(sender, ToolStripButton).Text
        If btnName = "" Then Return

        Select Case btnName
            Case "保存"
                If orderStatus <> "待审" Then
                    ShowWarning("当前订单状态不允许保存修改！")
                    Return
                End If
                SaveEditDetails()
            Case "标签打印"
                PrintLabels()
            Case "提取编码"
                ExtractProductCodes()
            Case "全选"
                SelectAll()
            Case "反选"
                InvertSelection()
            Case "批量修改"
                If UserOperation = "" OrElse Not UserOperation.Contains(",15商品入库修改,") Then
                    ShowWarning("无权操作！")
                    Return
                End If
                If orderStatus <> "待审" Then
                    ShowWarning("当前订单状态不允许批量修改！")
                    Return
                End If
                Dim batchForm As New InboundBatchEditForm()
                batchForm.SetParentForm(Me)
                batchForm.ShowDialog()
        End Select
    End Sub

    ' ========== _窗口_商品入库详情_尺寸被改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        Dim nWidth As Integer = Me.ClientRectangle.Width
        Dim nHeight As Integer = Me.ClientRectangle.Height
        If nWidth < 900 Then nWidth = 900
        If nHeight < 450 Then nHeight = 450

        toolStrip.Width = nWidth
        toolStrip.Left = 0
        toolStrip.Top = 0
        panelTop.Width = nWidth
        panelTop.Left = 0
        panelTop.Top = 52
    End Sub

    ' ========== 子程序_删除表格1 ==========
    Private Sub ClearTable()
        dgvDetail.Columns.Clear()
        dgvDetail.Rows.Clear()
    End Sub

    ' ========== _高级表格1_光标位置改变 ==========
    Private Sub DgvDetail_SelectionChanged(sender As Object, e As EventArgs)
        If dgvDetail.CurrentCell IsNot Nothing Then
            currentRow = dgvDetail.CurrentCell.RowIndex
            currentCol = dgvDetail.CurrentCell.ColumnIndex
        Else
            currentRow = -1
            currentCol = -1
        End If
    End Sub

    ' ========== _高级表格1_加载表头 ==========
    Private Sub LoadTableHeader()
        dgvDetail.Columns.Clear()

        Dim headers() As String
        Dim widths() As Integer

        If orderStatus = "待审" Then
            headers = {"序号", "商品编码", "商品名称", "商品品类", "规格", "款号", "材质", "圈口长度", "面宽", "厚度",
                       "工厂成色", "公司成色", "单件重", "数量", "库存", "金重", "损耗", "含耗重", "重量", "单位",
                       "石重", "石头数", "副石重", "副石头数", "成本单价", "系数", "成本工费", "参考工费", "销售工费",
                       "成本附加费", "销售附加费", "销售价", "库房", "备注", "主石色", "操作"}
            widths = {45, 100, 150, 75, 75, 110, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60,
                      60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 80, 60, 60, 60}
        Else
            headers = {"序号", "商品编码", "商品名称", "商品品类", "规格", "款号", "材质", "圈口长度", "面宽", "厚度",
                       "工厂成色", "公司成色", "单件重", "数量", "库存", "金重", "损耗", "含耗重", "重量", "单位",
                       "石重", "石头数", "副石重", "副石头数", "成本单价", "系数", "成本工费", "参考工费", "销售工费",
                       "成本附加费", "销售附加费", "销售价", "库房", "备注", "主石色"}
            widths = {45, 100, 150, 75, 75, 110, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60,
                      60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 80, 60, 60}
        End If

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Width = widths(i)
            col.SortMode = DataGridViewColumnSortMode.NotSortable
            dgvDetail.Columns.Add(col)
        Next

        ' 待审模式：第0列为文本型序号，最后添加操作列（按钮）
        ' 已审模式：第0列为选择型（CheckBox）
        If orderStatus = "待审" Then
            ' 第0列显示序号文本
            dgvDetail.Columns(COL_SEQ).ReadOnly = True
            ' 操作列使用按钮
            Dim deleteCol As New DataGridViewButtonColumn()
            deleteCol.HeaderText = "操作"
            deleteCol.Text = "删除"
            deleteCol.UseColumnTextForButtonValue = True
            deleteCol.Width = 60
            ' 替换最后一列
            dgvDetail.Columns.RemoveAt(COL_DELETE)
            dgvDetail.Columns.Add(deleteCol)
        Else
            ' 已审模式：第0列为CheckBox选择列
            Dim checkCol As New DataGridViewCheckBoxColumn()
            checkCol.HeaderText = "序号"
            checkCol.Width = 45
            dgvDetail.Columns.RemoveAt(COL_SEQ)
            dgvDetail.Columns.Insert(COL_SEQ, checkCol)
        End If
    End Sub

    ' ========== _高级表格1_加载数据 ==========
    Private Sub LoadTableData()
        If MySQL_Read Is Nothing Then
            ShowError("数据库读取连接无效！")
            Return
        End If

        Dim orderId As String = GetOrderID()
        If orderId = "" Then
            ShowError("未找到入库订单ID！")
            Return
        End If

        ' 读取订单基本信息
        Dim orderSql As String = "SELECT a.gold_price AS gold_price,a.odd_numbers AS aodd_numbers,a.xiangqian AS axiangqian," &
            "a.delivery AS adelivery,a.source AS asource,a.settlement AS asettlement,a.total AS atotal," &
            "b.title AS btitle,a.remarks AS aremarks,a.state AS astatus " &
            "FROM xipunum_erp_store_order AS a LEFT JOIN xipunum_erp_about AS b ON b.id=a.factory " &
            "WHERE a.id='" & SafeSQL(orderId) & "' ORDER BY a.id DESC LIMIT 1"
        Dim orderDt As DataTable = ExecuteQuery(orderSql, MySQL_Read)
        If orderDt Is Nothing OrElse orderDt.Rows.Count = 0 Then
            ShowError("入库订单不存在！")
            Return
        End If

        Dim dr As DataRow = orderDt.Rows(0)
        Dim orderNumber As String = SafeString(dr("aodd_numbers"))
        Dim xiangqian As String = SafeString(dr("axiangqian"))
        Dim deliveryNo As String = SafeString(dr("adelivery"))
        Dim source As String = SafeString(dr("asource"))
        Dim settlement As String = SafeString(dr("asettlement"))
        Dim total As String = SafeString(dr("atotal"))
        Dim factoryName As String = SafeString(dr("btitle"))
        Dim remarks As String = SafeString(dr("aremarks"))
        Dim goldPrice As String = SafeString(dr("gold_price"))
        orderStatus = SafeString(dr("astatus"))

        txtOrderNumber.Text = orderNumber
        txtDeliveryNumber.Text = deliveryNo
        txtRemarks.Text = remarks
        txtGoldPrice.Text = goldPrice

        If xiangqian = "镶嵌" Then
            rbNotInlaid.Checked = False
            rbInlaid.Checked = True
        Else
            rbNotInlaid.Checked = True
            rbInlaid.Checked = False
        End If

        cmbSource.Text = source
        cmbSettlement.Text = settlement
        cmbFactory.Text = factoryName

        ' 读取入库明细数据（5表联查）
        Dim detailSql As String = "SELECT a.poduct_code AS apoduct_code,b.product_name AS bproduct_name," &
            "CASE WHEN COALESCE(d.title,'')='' THEN '未匹配' ELSE d.title END AS ctitle," &
            "COALESCE(e1.title,e2.title,'未匹配') AS dtitle," &
            "b.item_number AS bitem_number,b.caizhi AS bcaizhi,b.quandu AS bquandu,b.wide AS bwide," &
            "b.thickness AS bthickness,a.factory_condition AS afactory_condition," &
            "a.company_condition AS acompany_condition,b.single AS bsingle,b.quantity AS bquantity," &
            "b.jin_zhong AS bjin_zhong,b.loss AS bloss,b.including AS bincluding,b.weight AS bweight," &
            "b.sales_unit AS bsales_unit,a.cost_price AS acost_price,a.coefficient as acoefficient," &
            "a.basic_cost AS abasic_cost,a.premium_cost AS apremium_cost,a.sales_cost AS asales_cost," &
            "a.company_surcharge AS acompany_surcharge,a.sales_surcharge AS asales_surcharge," &
            "a.sales_price AS asales_price," &
            "CASE WHEN b.kufang=0 OR b.kufang='0' THEN '总库' ELSE e.title END AS etitle," &
            "a.remarks AS aremarks,b.kufang AS bkufang " &
            "FROM xipunum_erp_store AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code=a.poduct_code " &
            "LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao=b.item_number AND b.item_number<>'' " &
            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id=i.specification_id AND b.item_number IS NOT NULL AND b.item_number<>'' " &
            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id=b.specification_id AND b.specification_id IS NOT NULL AND b.specification_id<>'' " &
            "LEFT JOIN xipunum_erp_category AS d ON d.id=COALESCE(e1.category_id,e2.category_id) " &
            "LEFT JOIN xipunum_erp_type AS e ON e.id=b.kufang " &
            "WHERE a.order_id='" & SafeSQL(orderId) & "' ORDER BY a.id ASC"
        Dim detailDt As DataTable = ExecuteQuery(detailSql, MySQL_Read)
        If detailDt Is Nothing Then
            ShowError("读取入库明细失败！")
            Return
        End If

        Dim dataCount As Integer = detailDt.Rows.Count
        productCount = dataCount

        ' 插入数据行（至少26行，含合计行）
        Dim totalRows As Integer
        If dataCount >= 25 Then
            totalRows = dataCount + 1
        Else
            totalRows = 26
        End If

        ' 添加数据行（含合计行）
        For i As Integer = 0 To totalRows - 1
            dgvDetail.Rows.Add()
        Next

        If dataCount = 0 Then
            CalculateStatistics()
            Return
        End If

        ' 填充数据
        For i As Integer = 0 To dataCount - 1
            Dim row As DataRow = detailDt.Rows(i)
            Dim productCode As String = SafeString(row("apoduct_code"))
            Dim productName As String = SafeString(row("bproduct_name"))
            Dim category As String = SafeString(row("ctitle"))
            Dim spec As String = SafeString(row("dtitle"))
            Dim kuanhao As String = SafeString(row("bitem_number"))
            Dim caizhi As String = SafeString(row("bcaizhi"))
            Dim quandu As String = SafeString(row("bquandu"))
            Dim width As String = SafeString(row("bwide"))
            Dim thickness As String = SafeString(row("bthickness"))
            Dim factoryCond As String = SafeString(row("afactory_condition"))
            Dim companyCond As String = SafeString(row("acompany_condition"))
            Dim single As String = SafeString(row("bsingle"))
            If single = "" Then single = "0.000"
            Dim quantity As String = SafeString(row("bquantity"))
            If quantity = "" Then quantity = "0"
            Dim jinZhong As String = SafeString(row("bjin_zhong"))
            If jinZhong = "" Then jinZhong = "0.000"
            Dim loss As String = SafeString(row("bloss"))
            If loss = "" Then loss = "0.000"
            Dim including As String = SafeString(row("bincluding"))
            If including = "" Then including = "0.000"
            Dim weight As String = SafeString(row("bweight"))
            If weight = "" Then weight = "0.000"
            Dim unit As String = SafeString(row("bsales_unit"))
            Dim costPrice As String = SafeString(row("acost_price"))
            If costPrice = "" Then costPrice = "0.00"
            Dim coefficient As String = SafeString(row("acoefficient"))
            If coefficient = "" Then coefficient = "0"
            Dim basicCost As String = SafeString(row("abasic_cost"))
            If basicCost = "" Then basicCost = "0.00"
            Dim premiumCost As String = SafeString(row("apremium_cost"))
            If premiumCost = "" Then premiumCost = "0.00"
            Dim salesCost As String = SafeString(row("asales_cost"))
            If salesCost = "" Then salesCost = "0.00"
            Dim companySur As String = SafeString(row("acompany_surcharge"))
            If companySur = "" Then companySur = "0.00"
            Dim salesSur As String = SafeString(row("asales_surcharge"))
            If salesSur = "" Then salesSur = "0.00"
            Dim salesPrice As String = SafeString(row("asales_price"))
            If salesPrice = "" Then salesPrice = "0.00"
            Dim kufang As String = SafeString(row("etitle"))
            Dim remarksRow As String = SafeString(row("aremarks"))
            Dim kufangId As String = SafeString(row("bkufang"))

            Dim shitou As String = "0.000"
            Dim stnum As String = "0"
            Dim shitou1 As String = "0.000"
            Dim shnum1 As String = "0"
            Dim zhuse As String = "白"

            Dim stock As String = "0"

            ' 镶嵌商品：查询镶嵌数据
            If xiangqian = "镶嵌" AndAlso productCode <> "" Then
                Dim xqSql As String = "SELECT shitou,stnum,shitou1,shnum1,zhuse FROM xipunum_erp_shop_xiangqian " &
                    "WHERE poduct_code='" & SafeSQL(productCode) & "' ORDER BY id ASC LIMIT 1"
                Dim xqDt As DataTable = ExecuteQuery(xqSql, MySQL_Read)
                If xqDt IsNot Nothing AndAlso xqDt.Rows.Count > 0 Then
                    shitou = SafeString(xqDt.Rows(0)("shitou"))
                    If shitou = "" Then shitou = "0.000"
                    stnum = SafeString(xqDt.Rows(0)("stnum"))
                    If stnum = "" Then stnum = "0"
                    shitou1 = SafeString(xqDt.Rows(0)("shitou1"))
                    If shitou1 = "" Then shitou1 = "0.000"
                    shnum1 = SafeString(xqDt.Rows(0)("shnum1"))
                    If shnum1 = "" Then shnum1 = "0"
                    zhuse = SafeString(xqDt.Rows(0)("zhuse"))
                    If zhuse = "" Then zhuse = "白"
                End If
            End If

            ' 查询库存数量
            If productCode <> "" Then
                Dim stockSql As String = "SELECT IFNULL(SUM(quantity),0) AS shuliang FROM xipunum_erp_shop_kucun " &
                    "WHERE poduct_code='" & SafeSQL(productCode) & "'"
                Dim stockDt As DataTable = ExecuteQuery(stockSql, MySQL_Read)
                If stockDt IsNot Nothing AndAlso stockDt.Rows.Count > 0 Then
                    stock = SafeString(stockDt.Rows(0)("shuliang"))
                End If
            End If
            If stock = "" Then stock = "0"

            ' 填充表格行
            Dim dgvRow As DataGridViewRow = dgvDetail.Rows(i)

            If orderStatus = "待审" Then
                ' 序号
                dgvRow.Cells(COL_SEQ).Value = i.ToString().PadLeft(dataCount.ToString().Length, "0"c)
            Else
                ' 已审模式：根据库存是否>0设置选中状态
                dgvRow.Cells(COL_SEQ).Value = (SafeDecimal(stock) > 0)
            End If

            dgvRow.Cells(COL_CODE).Value = productCode
            dgvRow.Cells(COL_NAME).Value = productName
            dgvRow.Cells(COL_CATEGORY).Value = category
            dgvRow.Cells(COL_SPEC).Value = spec
            dgvRow.Cells(COL_KUANHAO).Value = kuanhao
            dgvRow.Cells(COL_CAIZHI).Value = caizhi
            dgvRow.Cells(COL_QUANDU).Value = quandu
            dgvRow.Cells(COL_WIDTH).Value = width
            dgvRow.Cells(COL_THICKNESS).Value = thickness
            dgvRow.Cells(COL_FACTORY_COND).Value = factoryCond
            dgvRow.Cells(COL_COMPANY_COND).Value = companyCond
            dgvRow.Cells(COL_SINGLE).Value = single
            dgvRow.Cells(COL_QTY).Value = quantity
            dgvRow.Cells(COL_STOCK).Value = stock
            dgvRow.Cells(COL_JINZHONG).Value = jinZhong
            dgvRow.Cells(COL_LOSS).Value = loss
            dgvRow.Cells(COL_INCLUDING).Value = including
            dgvRow.Cells(COL_WEIGHT).Value = weight
            dgvRow.Cells(COL_UNIT).Value = unit
            dgvRow.Cells(COL_SHITOU).Value = shitou
            dgvRow.Cells(COL_STNUM).Value = stnum
            dgvRow.Cells(COL_SHITOU1).Value = shitou1
            dgvRow.Cells(COL_SHNUM1).Value = shnum1
            dgvRow.Cells(COL_COST_PRICE).Value = costPrice
            dgvRow.Cells(COL_COEFFICIENT).Value = coefficient
            dgvRow.Cells(COL_BASIC_COST).Value = basicCost
            dgvRow.Cells(COL_PREMIUM_COST).Value = premiumCost
            dgvRow.Cells(COL_SALES_COST).Value = salesCost
            dgvRow.Cells(COL_COMPANY_SUR).Value = companySur
            dgvRow.Cells(COL_SALES_SUR).Value = salesSur
            dgvRow.Cells(COL_SALES_PRICE).Value = salesPrice
            dgvRow.Cells(COL_KUFANG).Value = kufang
            dgvRow.Cells(COL_REMARKS).Value = remarksRow
            dgvRow.Cells(COL_ZHUSE).Value = zhuse
        Next

        ' 设置单元格只读和颜色
        SetupCellReadOnly()

        CalculateStatistics()
    End Sub

    ' ========== 设置单元格只读和颜色 ==========
    Private Sub SetupCellReadOnly()
        If dgvDetail.Rows.Count <= 1 Then Return
        Dim lastDataIndex As Integer = dgvDetail.Rows.Count - 2 ' 合计行前一行为最后一条数据

        If orderStatus <> "待审" Then
            ' 已审：全部只读
            For i As Integer = 0 To dgvDetail.Rows.Count - 1
                For j As Integer = 0 To dgvDetail.Columns.Count - 1
                    dgvDetail.Rows(i).Cells(j).ReadOnly = True
                    dgvDetail.Rows(i).Cells(j).Style.BackColor = System.Drawing.Color.Silver
                Next
            Next
        Else
            ' 待审模式：设置可编辑/只读列
            If lastDataIndex < 1 Then Return
            For i As Integer = 1 To lastDataIndex
                If SafeString(dgvDetail.Rows(i).Cells(COL_CODE).Value) <> "" Then
                    ' 只读列：0-5（序号、编码、名称、品类、规格、款号）
                    For j As Integer = 0 To 5
                        dgvDetail.Rows(i).Cells(j).ReadOnly = True
                        dgvDetail.Rows(i).Cells(j).Style.BackColor = System.Drawing.Color.LightGray
                    Next
                    ' 可编辑列：6-10（材质、圈口、面宽、厚度、工厂成色）
                    For j As Integer = 6 To 10
                        dgvDetail.Rows(i).Cells(j).ReadOnly = False
                        dgvDetail.Rows(i).Cells(j).Style.BackColor = System.Drawing.Color.White
                    Next
                    ' 只读列：11-12（公司成色、单件重）
                    For j As Integer = 11 To 12
                        dgvDetail.Rows(i).Cells(j).ReadOnly = True
                        dgvDetail.Rows(i).Cells(j).Style.BackColor = System.Drawing.Color.LightGray
                    Next
                    ' 可编辑列：13（数量）
                    dgvDetail.Rows(i).Cells(13).ReadOnly = False
                    dgvDetail.Rows(i).Cells(13).Style.BackColor = System.Drawing.Color.White
                    ' 只读列：14（库存）
                    dgvDetail.Rows(i).Cells(14).ReadOnly = True
                    dgvDetail.Rows(i).Cells(14).Style.BackColor = System.Drawing.Color.LightGray
                    ' 可编辑列：15-16（金重、损耗）
                    For j As Integer = 15 To 16
                        dgvDetail.Rows(i).Cells(j).ReadOnly = False
                        dgvDetail.Rows(i).Cells(j).Style.BackColor = System.Drawing.Color.White
                    Next
                    ' 只读列：17（含耗重）
                    dgvDetail.Rows(i).Cells(17).ReadOnly = True
                    dgvDetail.Rows(i).Cells(17).Style.BackColor = System.Drawing.Color.LightGray
                    ' 可编辑列：18（重量）
                    dgvDetail.Rows(i).Cells(18).ReadOnly = False
                    dgvDetail.Rows(i).Cells(18).Style.BackColor = System.Drawing.Color.White
                    ' 只读列：19（单位）
                    dgvDetail.Rows(i).Cells(19).ReadOnly = True
                    dgvDetail.Rows(i).Cells(19).Style.BackColor = System.Drawing.Color.LightGray
                    ' 镶嵌时可编辑：20-23（石重、石头数、副石重、副石头数）
                    If rbInlaid.Checked Then
                        For j As Integer = 20 To 23
                            dgvDetail.Rows(i).Cells(j).ReadOnly = False
                            dgvDetail.Rows(i).Cells(j).Style.BackColor = System.Drawing.Color.White
                        Next
                    Else
                        For j As Integer = 20 To 23
                            dgvDetail.Rows(i).Cells(j).ReadOnly = True
                            dgvDetail.Rows(i).Cells(j).Style.BackColor = System.Drawing.Color.LightGray
                        Next
                    End If
                    ' 只读列：24（成本单价）
                    dgvDetail.Rows(i).Cells(24).ReadOnly = True
                    dgvDetail.Rows(i).Cells(24).Style.BackColor = System.Drawing.Color.LightGray
                    ' 可编辑列：25-30（系数、成本工费、参考工费、销售工费、成本附加费、销售附加费）
                    For j As Integer = 25 To 30
                        dgvDetail.Rows(i).Cells(j).ReadOnly = False
                        dgvDetail.Rows(i).Cells(j).Style.BackColor = System.Drawing.Color.White
                    Next
                    ' 只读列：31-32（销售价、库房）
                    For j As Integer = 31 To 32
                        dgvDetail.Rows(i).Cells(j).ReadOnly = True
                        dgvDetail.Rows(i).Cells(j).Style.BackColor = System.Drawing.Color.LightGray
                    Next
                    ' 可编辑列：33-34（备注、主石色）
                    For j As Integer = 33 To 34
                        dgvDetail.Rows(i).Cells(j).ReadOnly = False
                        dgvDetail.Rows(i).Cells(j).Style.BackColor = System.Drawing.Color.White
                    Next
                End If
            Next

            ' 合计行只读
            Dim totalRowIndex As Integer = dgvDetail.Rows.Count - 1
            For j As Integer = 0 To dgvDetail.Columns.Count - 1
                dgvDetail.Rows(totalRowIndex).Cells(j).ReadOnly = True
                dgvDetail.Rows(totalRowIndex).Cells(j).Style.BackColor = System.Drawing.Color.LightGray
            Next
        End If
    End Sub

    ' ========== _高级表格1_数据统计 ==========
    Private Sub CalculateStatistics()
        If dgvDetail.Rows.Count <= 1 Then Return
        Dim totalRowIndex As Integer = dgvDetail.Rows.Count - 1

        ' 清空合计行
        For j As Integer = 0 To dgvDetail.Columns.Count - 1
            dgvDetail.Rows(totalRowIndex).Cells(j).Value = ""
        Next

        Dim sumQty As Decimal = 0
        Dim sumStock As Decimal = 0
        Dim sumJinZhong As Decimal = 0
        Dim sumWeight As Decimal = 0
        Dim sumShitou As Decimal = 0
        Dim sumStnum As Decimal = 0
        Dim sumShitou1 As Decimal = 0
        Dim sumShnum1 As Decimal = 0
        Dim sumCostPrice As Decimal = 0
        Dim sumBasicCost As Decimal = 0
        Dim sumPremiumCost As Decimal = 0
        Dim sumSalesCost As Decimal = 0
        Dim sumCompanySur As Decimal = 0
        Dim sumSalesSur As Decimal = 0

        Dim loopCount As Integer = totalRowIndex - 1
        If loopCount < 0 Then loopCount = 0

        For i As Integer = 0 To loopCount
            If SafeString(dgvDetail.Rows(i).Cells(COL_CODE).Value) <> "" Then
                sumQty += SafeDecimal(dgvDetail.Rows(i).Cells(COL_QTY).Value)
                sumStock += SafeDecimal(dgvDetail.Rows(i).Cells(COL_STOCK).Value)
                sumJinZhong += SafeDecimal(dgvDetail.Rows(i).Cells(COL_JINZHONG).Value)
                sumWeight += SafeDecimal(dgvDetail.Rows(i).Cells(COL_WEIGHT).Value)
                sumShitou += SafeDecimal(dgvDetail.Rows(i).Cells(COL_SHITOU).Value)
                sumStnum += SafeDecimal(dgvDetail.Rows(i).Cells(COL_STNUM).Value)
                sumShitou1 += SafeDecimal(dgvDetail.Rows(i).Cells(COL_SHITOU1).Value)
                sumShnum1 += SafeDecimal(dgvDetail.Rows(i).Cells(COL_SHNUM1).Value)
                sumCostPrice += SafeDecimal(dgvDetail.Rows(i).Cells(COL_COST_PRICE).Value) * SafeDecimal(dgvDetail.Rows(i).Cells(COL_QTY).Value)
                sumBasicCost += SafeDecimal(dgvDetail.Rows(i).Cells(COL_BASIC_COST).Value) * SafeDecimal(dgvDetail.Rows(i).Cells(COL_JINZHONG).Value)
                sumPremiumCost += SafeDecimal(dgvDetail.Rows(i).Cells(COL_PREMIUM_COST).Value) * SafeDecimal(dgvDetail.Rows(i).Cells(COL_JINZHONG).Value)
                sumSalesCost += SafeDecimal(dgvDetail.Rows(i).Cells(COL_SALES_COST).Value) * SafeDecimal(dgvDetail.Rows(i).Cells(COL_JINZHONG).Value)
                sumCompanySur += SafeDecimal(dgvDetail.Rows(i).Cells(COL_COMPANY_SUR).Value) * SafeDecimal(dgvDetail.Rows(i).Cells(COL_QTY).Value)
                sumSalesSur += SafeDecimal(dgvDetail.Rows(i).Cells(COL_SALES_SUR).Value)
            End If
        Next

        ' 写入合计行
        dgvDetail.Rows(totalRowIndex).Cells(COL_SEQ).Value = "合计"
        dgvDetail.Rows(totalRowIndex).Cells(COL_QTY).Value = sumQty.ToString()
        dgvDetail.Rows(totalRowIndex).Cells(COL_STOCK).Value = sumStock.ToString()
        dgvDetail.Rows(totalRowIndex).Cells(COL_JINZHONG).Value = FormatThreeDecimals(sumJinZhong)
        dgvDetail.Rows(totalRowIndex).Cells(COL_WEIGHT).Value = FormatThreeDecimals(sumWeight)
        dgvDetail.Rows(totalRowIndex).Cells(COL_SHITOU).Value = FormatThreeDecimals(sumShitou)
        dgvDetail.Rows(totalRowIndex).Cells(COL_STNUM).Value = sumStnum.ToString()
        dgvDetail.Rows(totalRowIndex).Cells(COL_SHITOU1).Value = FormatThreeDecimals(sumShitou1)
        dgvDetail.Rows(totalRowIndex).Cells(COL_SHNUM1).Value = sumShnum1.ToString()
        dgvDetail.Rows(totalRowIndex).Cells(COL_COST_PRICE).Value = FormatTwoDecimals(sumCostPrice)
        dgvDetail.Rows(totalRowIndex).Cells(COL_BASIC_COST).Value = FormatTwoDecimals(sumBasicCost)
        dgvDetail.Rows(totalRowIndex).Cells(COL_PREMIUM_COST).Value = FormatTwoDecimals(sumPremiumCost)
        dgvDetail.Rows(totalRowIndex).Cells(COL_SALES_COST).Value = FormatTwoDecimals(sumSalesCost)
        dgvDetail.Rows(totalRowIndex).Cells(COL_COMPANY_SUR).Value = FormatTwoDecimals(sumCompanySur)
        dgvDetail.Rows(totalRowIndex).Cells(COL_SALES_SUR).Value = FormatTwoDecimals(sumSalesSur)

        txtTotalWeight.Text = FormatThreeDecimals(sumWeight)
    End Sub

    ' ========== _超级按钮_编辑详情保存_被单击 ==========
    Private Sub SaveEditDetails()
        If MySQL_Write Is Nothing Then
            ShowError("数据库写入连接无效！")
            Return
        End If

        Dim totalRows As Integer = dgvDetail.Rows.Count
        If totalRows <= 2 Then
            ShowWarning("没有可保存的商品明细！")
            Return
        End If

        Dim totalRowIndex As Integer = totalRows - 1
        Dim saveCount As Integer = totalRows - 2
        Dim orderId As String = GetOrderID()
        If orderId = "" Then
            ShowError("未找到入库订单ID，无法保存！")
            Return
        End If

        Dim orderNumber As String = txtOrderNumber.Text
        Dim orderRemarks As String = txtRemarks.Text
        Dim operationDate As String = GetOperationDate()
        Dim operationAccount As String = GetOperationAccount()

        ' 开启事务
        Dim trans As MySqlTransaction = BeginTransaction(MySQL_Write)

        Try
            ' 删除已标记的商品
            If deletedCodes <> "" Then
                Dim codes() As String = deletedCodes.TrimEnd(","c).Split(","c)
                For Each code As String In codes
                    If code <> "" Then
                        ExecuteCommand("INSERT INTO xipunum_erp_shop_log(poduct_code,type,creationtime) VALUES('" & SafeSQL(code) & "','删除','" & operationDate & "')", MySQL_Write)
                        ExecuteCommand("DELETE FROM xipunum_erp_shop_lincun WHERE poduct_code='" & SafeSQL(code) & "'", MySQL_Write)
                        ExecuteCommand("DELETE FROM xipunum_erp_shop WHERE poduct_code='" & SafeSQL(code) & "'", MySQL_Write)
                        ExecuteCommand("DELETE FROM xipunum_erp_shop_xiangqian WHERE poduct_code='" & SafeSQL(code) & "'", MySQL_Write)
                        ExecuteCommand("DELETE FROM xipunum_erp_store WHERE poduct_code='" & SafeSQL(code) & "'", MySQL_Write)
                        ExecuteCommand("DELETE FROM xipunum_erp_zhengshu WHERE poduct_code='" & SafeSQL(code) & "'", MySQL_Write)
                        LogContent = "账户:" & UserAccount & " 删除入库商品编码：" & code & " 相关商品数据！"
                        ExecuteCommand("INSERT INTO xipunum_erp_xitong_log(type,title,conter,user,creationtime) VALUES('删除','入库删除','" & SafeSQL(LogContent) & "','" & SafeSQL(operationAccount) & "','" & operationDate & "')", MySQL_Write)
                    End If
                Next
            End If

            ' 更新订单主表
            Dim updateOrderSql As String = "UPDATE xipunum_erp_store_order SET " &
                "factory_zhezu='" & SafeSQL(SafeString(dgvDetail.Rows(totalRowIndex).Cells(COL_JINZHONG).Value)) & "'," &
                "jinzhong='" & SafeSQL(SafeString(dgvDetail.Rows(totalRowIndex).Cells(COL_JINZHONG).Value)) & "'," &
                "total='" & SafeSQL(SafeString(dgvDetail.Rows(totalRowIndex).Cells(COL_WEIGHT).Value)) & "'," &
                "total_weight='" & SafeSQL(SafeString(dgvDetail.Rows(totalRowIndex).Cells(COL_WEIGHT).Value)) & "'," &
                "jlzhong='" & SafeSQL(SafeString(dgvDetail.Rows(totalRowIndex).Cells(COL_JINZHONG).Value)) & "'," &
                "ljzhezu='" & SafeSQL(SafeString(dgvDetail.Rows(totalRowIndex).Cells(COL_JINZHONG).Value)) & "'," &
                "chengben='" & SafeSQL(SafeString(dgvDetail.Rows(totalRowIndex).Cells(COL_COST_PRICE).Value)) & "'," &
                "remarks='" & SafeSQL(orderRemarks) & "'," &
                "cjuser='" & SafeSQL(operationAccount) & "'," &
                "updatetime='" & operationDate & "' " &
                "WHERE odd_numbers='" & SafeSQL(orderNumber) & "' LIMIT 1"
            ExecuteCommand(updateOrderSql, MySQL_Write)

            ' 记录订单修改日志
            LogContent = "账户:" & UserAccount & " 修改入库订单：" & orderNumber & " 相关数据"
            ExecuteCommand("INSERT INTO xipunum_erp_xitong_log(type,title,conter,user,creationtime) VALUES('修改','修改入库订单','" & SafeSQL(LogContent) & "','" & SafeSQL(operationAccount) & "','" & operationDate & "')", MySQL_Write)
            ExecuteCommand("INSERT INTO xipunum_erp_store_log(order_id,user,conter,creationtime) VALUES('" & SafeSQL(orderId) & "','" & SafeSQL(operationAccount) & "','订单编辑','" & operationDate & "')", MySQL_Write)

            ' 遍历保存每个商品明细
            For i As Integer = 0 To saveCount - 1
                Dim productCode As String = SafeString(dgvDetail.Rows(i).Cells(COL_CODE).Value)
                If productCode <> "" Then
                    Dim caizhi As String = SafeString(dgvDetail.Rows(i).Cells(COL_CAIZHI).Value)
                    Dim quandu As String = SafeString(dgvDetail.Rows(i).Cells(COL_QUANDU).Value)
                    Dim wide As String = SafeString(dgvDetail.Rows(i).Cells(COL_WIDTH).Value)
                    Dim thickness As String = SafeString(dgvDetail.Rows(i).Cells(COL_THICKNESS).Value)
                    Dim factoryCond As String = SafeString(dgvDetail.Rows(i).Cells(COL_FACTORY_COND).Value)
                    Dim companyCond As String = SafeString(dgvDetail.Rows(i).Cells(COL_COMPANY_COND).Value)
                    Dim single As String = SafeString(dgvDetail.Rows(i).Cells(COL_SINGLE).Value)
                    Dim quantity As String = SafeString(dgvDetail.Rows(i).Cells(COL_QTY).Value)
                    Dim jinZhong As String = SafeString(dgvDetail.Rows(i).Cells(COL_JINZHONG).Value)
                    Dim loss As String = SafeString(dgvDetail.Rows(i).Cells(COL_LOSS).Value)
                    Dim including As String = SafeString(dgvDetail.Rows(i).Cells(COL_INCLUDING).Value)
                    Dim weight As String = SafeString(dgvDetail.Rows(i).Cells(COL_WEIGHT).Value)
                    Dim unit As String = SafeString(dgvDetail.Rows(i).Cells(COL_UNIT).Value)
                    Dim shitou As String = SafeString(dgvDetail.Rows(i).Cells(COL_SHITOU).Value)
                    Dim stnum As String = SafeString(dgvDetail.Rows(i).Cells(COL_STNUM).Value)
                    Dim shitou1 As String = SafeString(dgvDetail.Rows(i).Cells(COL_SHITOU1).Value)
                    Dim shnum1 As String = SafeString(dgvDetail.Rows(i).Cells(COL_SHNUM1).Value)
                    Dim costPrice As String = SafeString(dgvDetail.Rows(i).Cells(COL_COST_PRICE).Value)
                    Dim coefficient As String = SafeString(dgvDetail.Rows(i).Cells(COL_COEFFICIENT).Value)
                    Dim basicCost As String = SafeString(dgvDetail.Rows(i).Cells(COL_BASIC_COST).Value)
                    Dim premiumCost As String = SafeString(dgvDetail.Rows(i).Cells(COL_PREMIUM_COST).Value)
                    Dim salesCost As String = SafeString(dgvDetail.Rows(i).Cells(COL_SALES_COST).Value)
                    Dim companySur As String = SafeString(dgvDetail.Rows(i).Cells(COL_COMPANY_SUR).Value)
                    Dim salesSur As String = SafeString(dgvDetail.Rows(i).Cells(COL_SALES_SUR).Value)
                    Dim salesPrice As String = SafeString(dgvDetail.Rows(i).Cells(COL_SALES_PRICE).Value)
                    Dim remarksRow As String = SafeString(dgvDetail.Rows(i).Cells(COL_REMARKS).Value)
                    Dim zhuse As String = SafeString(dgvDetail.Rows(i).Cells(COL_ZHUSE).Value)

                    ' UPDATE xipunum_erp_shop
                    ExecuteCommand("UPDATE xipunum_erp_shop SET caizhi='" & SafeSQL(caizhi) & "',quandu='" & SafeSQL(quandu) & "',wide='" & SafeSQL(wide) & "',thickness='" & SafeSQL(thickness) & "',single='" & SafeSQL(single) & "',quantity='" & SafeSQL(quantity) & "',jin_zhong='" & SafeSQL(jinZhong) & "',loss='" & SafeSQL(loss) & "',including='" & SafeSQL(including) & "',weight='" & SafeSQL(weight) & "',sales_unit='" & SafeSQL(unit) & "',cjuser='" & SafeSQL(operationAccount) & "',updatetime='" & operationDate & "' WHERE poduct_code='" & SafeSQL(productCode) & "' LIMIT 1", MySQL_Write)

                    ' UPDATE xipunum_erp_shop_lincun
                    ExecuteCommand("UPDATE xipunum_erp_shop_lincun SET quantity='" & SafeSQL(quantity) & "',jinzhong='" & SafeSQL(jinZhong) & "',updatetime='" & operationDate & "' WHERE poduct_code='" & SafeSQL(productCode) & "' LIMIT 1", MySQL_Write)

                    ' UPDATE xipunum_erp_shopys
                    ExecuteCommand("UPDATE xipunum_erp_shopys SET single='" & SafeSQL(single) & "',quantity='" & SafeSQL(quantity) & "',jin_zhong='" & SafeSQL(jinZhong) & "',weight='" & SafeSQL(weight) & "',cjuser='" & SafeSQL(operationAccount) & "',updatetime='" & operationDate & "' WHERE poduct_code='" & SafeSQL(productCode) & "' LIMIT 1", MySQL_Write)

                    ' UPDATE xipunum_erp_store
                    ExecuteCommand("UPDATE xipunum_erp_store SET sales_price='" & SafeSQL(salesPrice) & "',factory_condition='" & SafeSQL(factoryCond) & "',company_condition='" & SafeSQL(companyCond) & "',cost_price='" & SafeSQL(costPrice) & "',coefficient='" & SafeSQL(coefficient) & "',basic_cost='" & SafeSQL(basicCost) & "',premium_cost='" & SafeSQL(premiumCost) & "',sales_cost='" & SafeSQL(salesCost) & "',company_surcharge='" & SafeSQL(companySur) & "',sales_surcharge='" & SafeSQL(salesSur) & "',remarks='" & SafeSQL(remarksRow) & "',cjuser='" & SafeSQL(operationAccount) & "',updatetime='" & operationDate & "' WHERE poduct_code='" & SafeSQL(productCode) & "' LIMIT 1", MySQL_Write)

                    ' INSERT xipunum_erp_history
                    ExecuteCommand("INSERT INTO xipunum_erp_history(poduct_code,updatetime,number,type,quantity,jinzhong,zhongliang,conter,cjuser) VALUES('" & SafeSQL(productCode) & "','" & operationDate & "','" & SafeSQL(orderNumber) & "','成品修改','" & SafeSQL(quantity) & "','" & SafeSQL(jinZhong) & "','" & SafeSQL(weight) & "','修改商品基本参数','" & SafeSQL(operationAccount) & "')", MySQL_Write)

                    ' INSERT xipunum_erp_shop_log
                    ExecuteCommand("INSERT INTO xipunum_erp_shop_log(poduct_code,type,creationtime) VALUES('" & SafeSQL(productCode) & "','编辑','" & operationDate & "')", MySQL_Write)

                    ' 镶嵌商品更新镶嵌数据
                    If rbInlaid.Checked Then
                        ExecuteCommand("UPDATE xipunum_erp_shop_xiangqian SET shitou='" & SafeSQL(shitou) & "',stnum='" & SafeSQL(stnum) & "',shitou1='" & SafeSQL(shitou1) & "',shnum1='" & SafeSQL(shnum1) & "',zhuse='" & SafeSQL(zhuse) & "',cjuser='" & SafeSQL(operationAccount) & "',updatetime='" & operationDate & "' WHERE poduct_code='" & SafeSQL(productCode) & "' LIMIT 1", MySQL_Write)
                    End If

                    ' 系统日志
                    LogContent = "账户:" & UserAccount & " 修改商品编码：" & productCode & " 相关数据"
                    ExecuteCommand("INSERT INTO xipunum_erp_xitong_log(type,title,conter,user,creationtime) VALUES('修改','修改商品信息','" & SafeSQL(LogContent) & "','" & SafeSQL(operationAccount) & "','" & operationDate & "')", MySQL_Write)
                End If
            Next

            ' 提交事务
            CommitTransaction(trans)
            ShowSuccess("商品编辑保存成功！")
            Me.Close()
        Catch ex As Exception
            If trans IsNot Nothing Then
                RollbackTransaction(trans)
            End If
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub

    ' ========== _超级按钮_标签打印_被单击 ==========
    Private Sub PrintLabels()
        Dim dataFile As String = Path.Combine(Application.StartupPath, "data\erpdata.mdb")
        Dim lockFile As String = Path.Combine(Application.StartupPath, "temp\lm_print.lock")

        If LabelPrinterConnection = "" OrElse LabelPrinterName = "" Then
            ShowWarning("标签打印机未配置！")
            Return
        End If
        If cmbLabelStyle.SelectedIndex = -1 Then
            ShowWarning("请选择标签模板！")
            Return
        End If

        Dim labelTemplateFile As String = Path.Combine(Application.StartupPath, "voucher\biaoqian\" & cmbLabelStyle.SelectedItem.ToString())
        If Not File.Exists(labelTemplateFile) Then
            ShowWarning("标签模板文件不存在！")
            Return
        End If
        If dgvDetail.Rows.Count <= 2 Then
            ShowWarning("没有可打印的商品数据！")
            Return
        End If

        ' 查询工厂简写
        Dim factoryJianxie As String = ""
        If cmbFactory.Text <> "" Then
            Dim factorySql As String = "SELECT jianxie FROM xipunum_erp_about WHERE title='" & SafeSQL(cmbFactory.Text) & "' ORDER BY id DESC LIMIT 1"
            Dim factoryDt As DataTable = ExecuteQuery(factorySql, MySQL_Read)
            If factoryDt IsNot Nothing AndAlso factoryDt.Rows.Count > 0 Then
                factoryJianxie = SafeString(factoryDt.Rows(0)("jianxie"))
            End If
        End If

        Dim printCount As Integer = dgvDetail.Rows.Count - 2
        Dim validCount As Integer = 0
        Dim actualInsert As Integer = 0

        ' 使用OleDb连接Access数据库
        Dim connStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dataFile
        Dim connAccess As New System.Data.OleDb.OleDbConnection(connStr)
        Try
            connAccess.Open()

            ' 清空打印数据表
            Dim cmdClear As New System.Data.OleDb.OleDbCommand("DELETE FROM biaoqian WHERE 1=1", connAccess)
            cmdClear.ExecuteNonQuery()

            ' 遍历表格行，插入打印数据
            For i As Integer = 0 To printCount - 1
                Dim shouldPrint As Boolean = False
                If orderStatus = "待审" Then
                    shouldPrint = True
                Else
                    shouldPrint = (dgvDetail.Rows(i).Cells(COL_SEQ).Value IsNot Nothing AndAlso CBool(dgvDetail.Rows(i).Cells(COL_SEQ).Value) = True)
                End If

                Dim productCode As String = SafeString(dgvDetail.Rows(i).Cells(COL_CODE).Value)
                If shouldPrint AndAlso productCode <> "" Then
                    Dim spec As String = SafeString(dgvDetail.Rows(i).Cells(COL_SPEC).Value)
                    Dim kuanhao As String = SafeString(dgvDetail.Rows(i).Cells(COL_KUANHAO).Value)
                    Dim caizhi As String = SafeString(dgvDetail.Rows(i).Cells(COL_CAIZHI).Value)
                    Dim quandu As String = SafeString(dgvDetail.Rows(i).Cells(COL_QUANDU).Value)
                    Dim companyCond As String = SafeString(dgvDetail.Rows(i).Cells(COL_COMPANY_COND).Value)
                    Dim single As String = SafeString(dgvDetail.Rows(i).Cells(COL_SINGLE).Value)
                    Dim jinZhong As String = SafeString(dgvDetail.Rows(i).Cells(COL_JINZHONG).Value)
                    Dim loss As String = SafeString(dgvDetail.Rows(i).Cells(COL_LOSS).Value)
                    Dim including As String = SafeString(dgvDetail.Rows(i).Cells(COL_INCLUDING).Value)
                    Dim weight As String = SafeString(dgvDetail.Rows(i).Cells(COL_WEIGHT).Value)
                    Dim shitou As String = SafeString(dgvDetail.Rows(i).Cells(COL_SHITOU).Value)
                    Dim stnum As String = SafeString(dgvDetail.Rows(i).Cells(COL_STNUM).Value)
                    Dim shitou1 As String = SafeString(dgvDetail.Rows(i).Cells(COL_SHITOU1).Value)
                    Dim shnum1 As String = SafeString(dgvDetail.Rows(i).Cells(COL_SHNUM1).Value)
                    Dim basicCost As String = SafeString(dgvDetail.Rows(i).Cells(COL_BASIC_COST).Value)
                    Dim premiumCost As String = SafeString(dgvDetail.Rows(i).Cells(COL_PREMIUM_COST).Value)
                    Dim salesCost As String = SafeString(dgvDetail.Rows(i).Cells(COL_SALES_COST).Value)
                    Dim companySur As String = SafeString(dgvDetail.Rows(i).Cells(COL_COMPANY_SUR).Value)
                    Dim salesSur As String = SafeString(dgvDetail.Rows(i).Cells(COL_SALES_SUR).Value)
                    Dim salesPrice As String = SafeString(dgvDetail.Rows(i).Cells(COL_SALES_PRICE).Value)
                    Dim kufang As String = SafeString(dgvDetail.Rows(i).Cells(COL_KUFANG).Value)

                    ' 查询库房信息
                    Dim shopName As String = ""
                    Dim shopJianxie As String = ""
                    Dim shopCompany As String = ""
                    Dim shopAddress As String = ""
                    If kufang <> "" Then
                        Dim shopSql As String = "SELECT title,data1,data2,data3 FROM xipunum_erp_type WHERE type='商铺' AND superior='0' AND title='" & SafeSQL(kufang) & "' ORDER BY id DESC LIMIT 1"
                        Dim shopDt As DataTable = ExecuteQuery(shopSql, MySQL_Read)
                        If shopDt IsNot Nothing AndAlso shopDt.Rows.Count > 0 Then
                            shopName = SafeString(shopDt.Rows(0)("title"))
                            shopJianxie = SafeString(shopDt.Rows(0)("data1"))
                            shopCompany = SafeString(shopDt.Rows(0)("data2"))
                            shopAddress = SafeString(shopDt.Rows(0)("data3"))
                        End If
                    End If

                    ' 查询商品基本信息
                    Dim preSalePrice As String = ""
                    Dim productName As String = ""
                    Dim certNo As String = ""
                    Dim color As String = ""
                    Dim infoSql As String = "SELECT c.sales_price AS sales_price,a.product_name AS mingcheng," &
                        "b.zsbianma AS zhengshu,b.yanse AS yanse " &
                        "FROM xipunum_erp_shop AS a LEFT JOIN xipunum_erp_zhengshu AS b ON b.poduct_code=a.poduct_code " &
                        "INNER JOIN xipunum_erp_store AS c ON c.poduct_code=a.poduct_code " &
                        "WHERE a.poduct_code='" & SafeSQL(productCode) & "' LIMIT 1"
                    Dim infoDt As DataTable = ExecuteQuery(infoSql, MySQL_Read)
                    If infoDt IsNot Nothing AndAlso infoDt.Rows.Count > 0 Then
                        preSalePrice = SafeString(infoDt.Rows(0)("sales_price"))
                        productName = SafeString(infoDt.Rows(0)("mingcheng"))
                        certNo = SafeString(infoDt.Rows(0)("zhengshu"))
                        color = SafeString(infoDt.Rows(0)("yanse"))
                    End If

                    ' 插入Access打印数据
                    Dim insertSql As String = "INSERT INTO biaoqian(工厂简写,工厂工费,成本工费,商铺简写,商品编码,商铺地址,商铺名称,材质,规格,圈号,净含量,净重,总重,销售工费,附加费,石重,石头数,副石重,副石头数,销售价,实销价,商品名称,其他名,证书号,标识码,颜色) VALUES(" &
                        "'" & SafeSQL(factoryJianxie) & "','" & SafeSQL(premiumCost) & "','" & SafeSQL(basicCost) & "','" & SafeSQL(shopJianxie) & "','" & SafeSQL(productCode) & "','" & SafeSQL(shopAddress) & "','" & SafeSQL(shopCompany) & "','" & SafeSQL(caizhi) & "','" & SafeSQL(spec) & "','" & SafeSQL(quandu) & "','" & SafeSQL(companyCond) & "','" & SafeSQL(jinZhong) & "','" & SafeSQL(weight) & "','" & SafeSQL(salesCost) & "','" & SafeSQL(salesSur) & "','" & SafeSQL(shitou) & "','" & SafeSQL(stnum) & "','" & SafeSQL(shitou1) & "','" & SafeSQL(shnum1) & "','" & SafeSQL(preSalePrice) & "','','" & SafeSQL(productName) & "','','" & SafeSQL(certNo) & "','','" & SafeSQL(color) & "')"
                    Dim cmdInsert As New System.Data.OleDb.OleDbCommand(insertSql, connAccess)
                    cmdInsert.ExecuteNonQuery()
                    validCount += 1
                    actualInsert += 1
                End If
            Next

            connAccess.Close()

            If validCount = 0 Then
                ShowWarning("没有选中可打印的商品！")
                Return
            End If

            ' 创建锁文件目录
            Dim tempDir As String = Path.Combine(Application.StartupPath, "temp")
            If Not Directory.Exists(tempDir) Then
                Directory.CreateDirectory(tempDir)
            End If
            If File.Exists(lockFile) Then
                File.Delete(lockFile)
            End If

            ' 拼接LMWPRINT命令行
            Dim cmdLine As String = """" & LabelPrinterConnection & """"
            cmdLine &= " /L=""" & labelTemplateFile & """"
            cmdLine &= " /C=" & actualInsert.ToString()
            cmdLine &= " /X=""" & LabelPrinterName & """"
            cmdLine &= " /N /Z=2 /FL=""" & lockFile & """"

            ' 运行打印命令
            Dim psi As New System.Diagnostics.ProcessStartInfo()
            psi.FileName = "cmd.exe"
            psi.Arguments = "/c " & cmdLine
            psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
            psi.CreateNoWindow = True
            System.Diagnostics.Process.Start(psi)

            System.Threading.Thread.Sleep(300)
        Catch ex As Exception
            If connAccess.State = ConnectionState.Open Then
                connAccess.Close()
            End If
            ShowError("标签打印失败：" & ex.Message)
        End Try
    End Sub

    ' ========== _超级按钮_提取编码数据_被单击 ==========
    Private Sub ExtractProductCodes()
        InboundProductCodeText = ""
        Dim codeText As String = ""
        Dim extractCount As Integer = 0

        If dgvDetail.Rows.Count <= 2 Then
            ShowWarning("没有可提取的商品编码！")
            Return
        End If

        Dim dataCount As Integer = dgvDetail.Rows.Count - 2
        For i As Integer = 0 To dataCount - 1
            Dim code As String = SafeString(dgvDetail.Rows(i).Cells(COL_CODE).Value)
            If code <> "" Then
                If orderStatus = "待审" Then
                    extractCount += 1
                    codeText &= code & vbCrLf
                Else
                    If dgvDetail.Rows(i).Cells(COL_SEQ).Value IsNot Nothing AndAlso CBool(dgvDetail.Rows(i).Cells(COL_SEQ).Value) = True Then
                        extractCount += 1
                        codeText &= code & vbCrLf
                    End If
                End If
            End If
        Next

        If extractCount = 0 Then
            ShowWarning("提取商品编码数量不能为0！")
            Return
        End If

        InboundProductCodeText = codeText
        ' 通知主窗口刷新
        HomePageQueryText = "商品入库"
        Me.Close()
    End Sub

    ' ========== _超级按钮_详情全选_被单击 ==========
    Private Sub SelectAll()
        If orderStatus = "待审" Then Return
        If dgvDetail.Rows.Count <= 2 Then Return

        Dim dataCount As Integer = dgvDetail.Rows.Count - 2
        For i As Integer = 0 To dataCount - 1
            If SafeString(dgvDetail.Rows(i).Cells(COL_CODE).Value) <> "" Then
                If SafeDecimal(dgvDetail.Rows(i).Cells(COL_STOCK).Value) > 0 Then
                    dgvDetail.Rows(i).Cells(COL_SEQ).Value = True
                End If
            End If
        Next
    End Sub

    ' ========== _超级按钮_详情反选_被单击 ==========
    Private Sub InvertSelection()
        If orderStatus = "待审" Then Return
        If dgvDetail.Rows.Count <= 2 Then Return

        Dim dataCount As Integer = dgvDetail.Rows.Count - 2
        For i As Integer = 0 To dataCount - 1
            If SafeString(dgvDetail.Rows(i).Cells(COL_CODE).Value) <> "" Then
                If SafeDecimal(dgvDetail.Rows(i).Cells(COL_STOCK).Value) > 0 Then
                    Dim currentVal As Boolean = False
                    If dgvDetail.Rows(i).Cells(COL_SEQ).Value IsNot Nothing Then
                        currentVal = CBool(dgvDetail.Rows(i).Cells(COL_SEQ).Value)
                    End If
                    dgvDetail.Rows(i).Cells(COL_SEQ).Value = Not currentVal
                End If
            End If
        Next
    End Sub

    ' ========== _高级表格1_结束编辑 ==========
    Private Sub DgvDetail_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex <= 0 Then Return
        If e.RowIndex >= dgvDetail.Rows.Count - 1 Then Return
        If e.ColumnIndex <= 0 OrElse e.ColumnIndex > 34 Then Return
        If SafeString(dgvDetail.Rows(e.RowIndex).Cells(COL_CODE).Value) = "" Then Return

        Dim rowIdx As Integer = e.RowIndex
        Dim colIdx As Integer = e.ColumnIndex

        ' 读取原料价
        Dim goldPrice As Decimal = SafeDecimal(txtGoldPrice.Text)
        ' 读取工厂成色
        Dim factoryCond As String = SafeString(dgvDetail.Rows(rowIdx).Cells(COL_FACTORY_COND).Value)

        Dim qty As Decimal = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_QTY).Value)
        Dim jinZhong As Decimal = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_JINZHONG).Value)
        Dim loss As Decimal = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_LOSS).Value)
        Dim weight As Decimal = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_WEIGHT).Value)
        Dim shitou As Decimal = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_SHITOU).Value)
        Dim shitou1 As Decimal = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_SHITOU1).Value)
        Dim coefficient As String = SafeString(dgvDetail.Rows(rowIdx).Cells(COL_COEFFICIENT).Value)
        Dim basicCost As Decimal = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_BASIC_COST).Value)
        Dim premiumCost As Decimal = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_PREMIUM_COST).Value)
        Dim salesCost As Decimal = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_SALES_COST).Value)
        Dim companySur As Decimal = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_COMPANY_SUR).Value)
        Dim salesSur As Decimal = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_SALES_SUR).Value)

        Dim single As Decimal = 0
        Dim including As Decimal = 0
        Dim costPrice As Decimal = 0
        Dim salesPrice As Decimal = 0

        ' 验证
        If qty <= 0 Then
            ShowWarning("入库数量必须大于0！")
            qty = 1
            dgvDetail.Rows(rowIdx).Cells(COL_QTY).Value = "1"
        End If
        If jinZhong < 0 Then
            jinZhong = 0
            dgvDetail.Rows(rowIdx).Cells(COL_JINZHONG).Value = "0"
        End If
        If weight < 0 Then
            weight = 0
            dgvDetail.Rows(rowIdx).Cells(COL_WEIGHT).Value = "0"
        End If
        If shitou < 0 Then
            shitou = 0
            dgvDetail.Rows(rowIdx).Cells(COL_SHITOU).Value = "0"
        End If
        If shitou1 < 0 Then
            shitou1 = 0
            dgvDetail.Rows(rowIdx).Cells(COL_SHITOU1).Value = "0"
        End If

        ' 根据编辑列进行计算
        Select Case colIdx
            Case COL_QTY ' 13: 数量
                single = weight / qty
            Case COL_JINZHONG ' 15: 金重
                If jinZhong > weight Then
                    ShowWarning("金重不可以大于重量！")
                    jinZhong = weight
                    dgvDetail.Rows(rowIdx).Cells(COL_JINZHONG).Value = jinZhong.ToString()
                End If
                weight = (shitou + shitou1) / 5 + jinZhong
                single = weight / qty
                costPrice = jinZhong * (basicCost + goldPrice) + companySur
                salesPrice = jinZhong * (salesCost + goldPrice) + salesSur
            Case COL_LOSS ' 16: 损耗
                including = loss + weight
            Case COL_WEIGHT ' 18: 重量
                If weight < jinZhong Then
                    ShowWarning("商品重量不能小于金重！")
                    weight = (shitou + shitou1) / 5 + jinZhong
                    dgvDetail.Rows(rowIdx).Cells(COL_WEIGHT).Value = weight.ToString()
                End If
                single = weight / qty
                including = loss + weight
                shitou = (weight - jinZhong) * 5 - shitou1
                If shitou < 0 Then shitou = 0
            Case COL_SHITOU ' 20: 石重
                shitou = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_SHITOU).Value)
                If shitou < 0 Then shitou = 0
                weight = (shitou + shitou1) / 5 + jinZhong
                single = weight / qty
                including = loss + weight
            Case COL_SHITOU1 ' 22: 副石重
                shitou1 = SafeDecimal(dgvDetail.Rows(rowIdx).Cells(COL_SHITOU1).Value)
                If shitou1 < 0 Then shitou1 = 0
                weight = (shitou + shitou1) / 5 + jinZhong
                single = weight / qty
                including = loss + weight
            Case COL_BASIC_COST, COL_COMPANY_SUR ' 26, 29: 成本工费/成本附加费
                costPrice = jinZhong * (basicCost + goldPrice) + companySur
            Case COL_PREMIUM_COST ' 27: 参考工费
                salesCost = premiumCost
                salesPrice = jinZhong * (salesCost + goldPrice) + salesSur
            Case COL_SALES_COST, COL_SALES_SUR ' 28, 30: 销售工费/销售附加费
                salesPrice = jinZhong * (salesCost + goldPrice) + salesSur
        End Select

        ' 格式化并写回单元格
        dgvDetail.Rows(rowIdx).Cells(COL_FACTORY_COND).Value = factoryCond
        dgvDetail.Rows(rowIdx).Cells(COL_COMPANY_COND).Value = factoryCond
        dgvDetail.Rows(rowIdx).Cells(COL_SINGLE).Value = FormatThreeDecimals(single)
        dgvDetail.Rows(rowIdx).Cells(COL_QTY).Value = qty.ToString()
        dgvDetail.Rows(rowIdx).Cells(COL_JINZHONG).Value = FormatThreeDecimals(jinZhong)
        dgvDetail.Rows(rowIdx).Cells(COL_LOSS).Value = FormatThreeDecimals(loss)
        dgvDetail.Rows(rowIdx).Cells(COL_INCLUDING).Value = FormatThreeDecimals(including)
        dgvDetail.Rows(rowIdx).Cells(COL_WEIGHT).Value = FormatThreeDecimals(weight)
        dgvDetail.Rows(rowIdx).Cells(COL_SHITOU).Value = FormatThreeDecimals(shitou)
        dgvDetail.Rows(rowIdx).Cells(COL_STNUM).Value = SafeString(dgvDetail.Rows(rowIdx).Cells(COL_STNUM).Value)
        dgvDetail.Rows(rowIdx).Cells(COL_SHITOU1).Value = FormatThreeDecimals(shitou1)
        dgvDetail.Rows(rowIdx).Cells(COL_SHNUM1).Value = SafeString(dgvDetail.Rows(rowIdx).Cells(COL_SHNUM1).Value)
        dgvDetail.Rows(rowIdx).Cells(COL_COEFFICIENT).Value = coefficient
        dgvDetail.Rows(rowIdx).Cells(COL_BASIC_COST).Value = FormatTwoDecimals(basicCost)
        dgvDetail.Rows(rowIdx).Cells(COL_PREMIUM_COST).Value = FormatTwoDecimals(premiumCost)
        dgvDetail.Rows(rowIdx).Cells(COL_SALES_COST).Value = FormatTwoDecimals(salesCost)
        dgvDetail.Rows(rowIdx).Cells(COL_COMPANY_SUR).Value = FormatTwoDecimals(companySur)
        dgvDetail.Rows(rowIdx).Cells(COL_SALES_SUR).Value = FormatTwoDecimals(salesSur)
        dgvDetail.Rows(rowIdx).Cells(COL_COST_PRICE).Value = FormatTwoDecimals(costPrice)
        dgvDetail.Rows(rowIdx).Cells(COL_SALES_PRICE).Value = FormatTwoDecimals(salesPrice)

        CalculateStatistics()
    End Sub

    ' ========== _高级表格1_按钮被点击 ==========
    Private Sub DgvDetail_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return

        ' 删除按钮点击（仅待审模式）
        If orderStatus <> "待审" Then Return
        If e.ColumnIndex <> dgvDetail.Columns.Count - 1 Then Return
        If e.RowIndex <= 0 OrElse e.RowIndex >= dgvDetail.Rows.Count - 1 Then Return

        Dim productCode As String = SafeString(dgvDetail.Rows(e.RowIndex).Cells(COL_CODE).Value)
        If productCode = "" Then Return

        If Not ConfirmAction("真的删除吗？") Then Return

        ' 检查商品是否已变更
        Dim changeSql As String = "SELECT COUNT(a.poduct_code) AS shuliang FROM xipunum_erp_store AS a " &
            "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
            "INNER JOIN xipunum_erp_shop_kucun AS c ON c.poduct_code = a.poduct_code " &
            "WHERE a.poduct_code ='" & SafeSQL(productCode) & "' AND (b.jin_zhong <> c.jinzhong OR b.kufang <> c.kufang)"
        Dim changeDt As DataTable = ExecuteQuery(changeSql, MySQL_Read)
        Dim changeCount As String = "0"
        If changeDt IsNot Nothing AndAlso changeDt.Rows.Count > 0 Then
            changeCount = SafeString(changeDt.Rows(0)("shuliang"))
        End If

        If SafeDecimal(changeCount) > 0 Then
            ShowWarning("此入库商品有变更，不可以删除！")
            Return
        End If

        ' 标记删除
        If productCount > 0 Then
            productCount -= 1
        End If
        deletedCodes &= productCode & ","

        ' 删除行
        dgvDetail.Rows.RemoveAt(e.RowIndex)

        ' 删除合计行（如果存在）
        If dgvDetail.Rows.Count > 1 Then
            If SafeString(dgvDetail.Rows(dgvDetail.Rows.Count - 1).Cells(COL_SEQ).Value) = "合计" Then
                dgvDetail.Rows.RemoveAt(dgvDetail.Rows.Count - 1)
            End If
        End If

        ' 补充空行至26行
        While dgvDetail.Rows.Count < 26
            dgvDetail.Rows.Add()
        End While

        ' 重新编号
        For i As Integer = 0 To productCount - 1
            dgvDetail.Rows(i).Cells(COL_SEQ).Value = i.ToString()
        Next

        CalculateStatistics()
    End Sub

    ' ========== _入库详情_取订单ID ==========
    Private Function GetOrderID() As String
        If MySQL_Read Is Nothing Then Return ""

        If InboundOrderNumber <> "" Then
            Dim sql As String = "SELECT id FROM xipunum_erp_store_order WHERE odd_numbers='" & SafeSQL(InboundOrderNumber) & "' ORDER BY id DESC LIMIT 1"
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_Read)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Return SafeString(dt.Rows(0)("id"))
            End If
        End If

        Return ""
    End Function

    ' ========== _入库详情_读取订单状态 ==========
    Private Function ReadOrderStatus(orderId As String) As String
        If MySQL_Read Is Nothing Then Return ""
        If orderId = "" Then Return ""

        Dim sql As String = "SELECT state FROM xipunum_erp_store_order WHERE id='" & SafeSQL(orderId) & "' LIMIT 1"
        Dim dt As DataTable = ExecuteQuery(sql, MySQL_Read)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Return SafeString(dt.Rows(0)("state"))
        End If
        Return ""
    End Function

    ' ========== 格式三位小数 ==========
    Private Function FormatThreeDecimals(value As Decimal) As String
        Dim rounded As Decimal = Math.Round(value, 3)
        Dim text As String = rounded.ToString()
        If text = "" Then Return "0.000"
        If text.Contains(".") Then
            Dim parts() As String = text.Split("."c)
            If parts.Length > 1 Then
                Select Case parts(1).Length
                    Case 1 : Return text & "00"
                    Case 2 : Return text & "0"
                    Case Else : Return text
                End Select
            End If
        Else
            Return text & ".000"
        End If
        Return text
    End Function

    ' ========== 格式二位小数 ==========
    Private Function FormatTwoDecimals(value As Decimal) As String
        Dim rounded As Decimal = Math.Round(value, 2)
        Dim text As String = rounded.ToString()
        If text = "" Then Return "0.00"
        If text.Contains(".") Then
            Dim parts() As String = text.Split("."c)
            If parts.Length > 1 Then
                If parts(1).Length = 1 Then
                    Return text & "0"
                Else
                    Return text
                End If
            End If
        Else
            Return text & ".00"
        End If
        Return text
    End Function

    ' ========== 复制单元格 ==========
    Public Sub CopyCellValue()
        If dgvDetail.CurrentCell Is Nothing Then Return
        If dgvDetail.CurrentCell.Value IsNot Nothing Then
            Clipboard.SetText(dgvDetail.CurrentCell.Value.ToString())
        End If
    End Sub

    ' ========== 供InboundBatchEditForm调用的公开方法 ==========
    Public Function GetDetailGrid() As DataGridView
        Return dgvDetail
    End Function

    Public Function IsNotInlaid() As Boolean
        Return rbNotInlaid.Checked
    End Function

    Public Sub RecalculateStatistics()
        CalculateStatistics()
    End Sub

End Class

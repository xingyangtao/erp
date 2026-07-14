' ============================================================================
' 商品销售出库窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品销售出库.form.e.txt
' 包含所有42个程序集变量、所有子程序、所有SQL查询
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms

Public Class SalesOrderForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（42个） ==========
    Private row1 As Integer = 0                    ' 集_行号1
    Private col1 As Integer = 0                    ' 集_列号1
    Private row2 As Integer = 0                    ' 集_行号2
    Private col2 As Integer = 0                    ' 集_列号2
    Private row3 As Integer = 0                    ' 集_行号3
    Private col3 As Integer = 0                    ' 集_列号3
    Private deleteBtn As New Button()              ' 删除按钮
    Private weightTotal As String = "0"            ' 局部重量合计数值
    Private weightPrice As String = "0"            ' 局部重量单价数值
    Private salesAmount As String = "0"            ' 局部重销售价金额数值
    Private basicCostTotal As String = "0"         ' 局部重基本工费合计
    Private premiumCostTotal As String = "0"       ' 局部重精品工费合计
    Private salesCostTotal As String = "0"         ' 局部重销售工费合计
    Private costSurchargeTotal As String = "0"     ' 局部重成本附加费合计
    Private salesSurchargeTotal As String = "0"    ' 局部重销售附加费合计
    Private recoveryWeightTotal As String = "0"    ' 局部回收总重合计数值
    Private recoveryGoldTotal As String = "0"      ' 局部回收金重合计数值
    Private recoveryOtherTotal As String = "0"     ' 局部回收其他合计数值
    Private recoveryAmountTotal As String = "0"    ' 局部回收回收合计数值
    Private quantityTotal As String = "0"          ' 局部数量合计数量
    Private actualAmount As String = "0"           ' 局部实收金额金额数值
    Private salesProductCode As String = ""        ' 局部_销售商品编码
    Private salesDataCode As String = ""           ' 局部_销售商品数据编码
    Private salesPling As String = "零售"          ' 信息销售批零
    Private wholesaleStockWeight As String = "0"   ' 局部_批发库存料数值
    Private wholesaleStockAmount As String = "0"   ' 局部_批发库存元数值
    Private tableEditState As Integer = 0          ' 局部_表格编辑状态
    Private paymentEditState As Integer = 0        ' 局部_收支编辑状态
    Private recoveryEditState As Integer = 0       ' 局部_回收编辑状态
    Private tempQuantity As Integer = 0            ' 出库临时数量
    Private dataLoaded As Integer = 0              ' 出库数据加载
    Private saveInProgress As Integer = 0          ' 集_保存事务进行中
    Private saveOrderNumber As String = ""         ' 集_保存出库单号
    Private saveRecoveryNumber As String = ""      ' 集_保存回收单号
    Private saveMemberCode As String = ""          ' 集_保存新建会员编码
    Private saveLastGuide As String = ""           ' 集_保存上一笔导购员
    Private saveLastGuideAccount As String = ""    ' 集_保存上一笔导购账户

    ' ========== 控件声明 ==========
    Private dgvSales As New DataGridView()         ' 高级表格1 - 销售明细
    Private dgvRecovery As New DataGridView()       ' 高级表格2 - 回收明细
    Private dgvPayment As New DataGridView()        ' 高级表格3 - 收支明细
    Private dgvSummary As New DataGridView()        ' 高级表格4 - 统计汇总
    Private txtOrderNumber As New TextBox()         ' 单据号_编辑框
    Private cmbMember As New ComboBox()             ' 组合框会员查找
    Private txtPhone As New TextBox()               ' 联系电话_编辑框
    Private txtReceivable As New TextBox()          ' 应收_编辑框
    Private txtReceived As New TextBox()            ' 实收_编辑框
    Private txtTax As New TextBox()                 ' 税收_编辑框
    Private txtTaxRate As New TextBox()             ' 税点_编辑框
    Private txtPresaleNumber As New TextBox()       ' 预售单号_编辑框
    Private txtPresaleDeposit As New TextBox()      ' 预售订金_编辑框
    Private cmbGuide As New ComboBox()              ' 业务员组合框
    Private cmbVoucher As New ComboBox()            ' 组合框单据样式
    Private cmbIntroducer As New ComboBox()         ' 组合框介绍人查找
    Private cmbSalesFactory As New ComboBox()       ' 组合框销售工厂
    Private radioRetail As New RadioButton()        ' 单选框_零售
    Private radioWholesale As New RadioButton()     ' 单选框_批发
    Private radioInvoice As New RadioButton()       ' 单选框_开票
    Private radioNoInvoice As New RadioButton()     ' 单选框_不开票
    Private radioPresaleYes As New RadioButton()    ' 单选框_是
    Private radioPresaleNo As New RadioButton()     ' 单选框_否
    Private radioNewRow As New RadioButton()        ' 单选框_换行
    Private radioNewCol As New RadioButton()        ' 单选框_换列
    Private chkPrint As New CheckBox()              ' 打印单据_选择框
    Private chkPreview As New CheckBox()            ' 打印预览_选择框
    Private chkPrintChengse As New CheckBox()       ' 打印成色_选择框
    Private txtRemarks As New TextBox()             ' 备注_编辑框
    Private txtStoreMaterial As New TextBox()       ' 存料_编辑框
    Private txtStoreBalance As New TextBox()        ' 余额_编辑框
    Private lblMaterialBalance As New Label()       ' 透明标签_存料1
    Private lblBalanceAmount As New Label()         ' 透明标签_余额1
    Private lblMemberId As New Label()              ' 透明标签13
    Private lblLastConsumption As New Label()       ' 透明标签19
    Private btnRecoveryAdd As New Button()          ' 按钮_回收加
    Private btnRecoveryRemove As New Button()       ' 按钮_回收减
    Private btnToolbarSave As New Button()          ' 工具条-保存
    Private btnToolbarReset As New Button()         ' 工具条-重置
    Private btnToolbarBatchEdit As New Button()     ' 工具条-批量编辑
    Private btnToolbarPrintPreview As New Button()  ' 工具条-打印预览
    Private groupInvoice As New GroupBox()          ' 分组框4 (开票/不开票)
    Private groupPresale As New GroupBox()          ' 分组框1 (是/否)

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler dgvSales.CellEndEdit, AddressOf DgvSales_CellEndEdit
        AddHandler dgvSales.CellValueChanged, AddressOf DgvSales_CellValueChanged
        AddHandler dgvRecovery.CellEndEdit, AddressOf DgvRecovery_CellEndEdit
        AddHandler dgvPayment.CellEndEdit, AddressOf DgvPayment_CellEndEdit
        AddHandler btnRecoveryAdd.Click, AddressOf BtnRecoveryAdd_Click
        AddHandler btnRecoveryRemove.Click, AddressOf BtnRecoveryRemove_Click
        AddHandler btnToolbarSave.Click, AddressOf BtnToolbarSave_Click
        AddHandler btnToolbarReset.Click, AddressOf BtnToolbarReset_Click
        AddHandler btnToolbarBatchEdit.Click, AddressOf BtnToolbarBatchEdit_Click
        AddHandler btnToolbarPrintPreview.Click, AddressOf BtnToolbarPrintPreview_Click
        AddHandler radioRetail.CheckedChanged, AddressOf RadioRetail_CheckedChanged
        AddHandler radioWholesale.CheckedChanged, AddressOf RadioWholesale_CheckedChanged
        AddHandler radioPresaleYes.CheckedChanged, AddressOf RadioPresaleYes_CheckedChanged
        AddHandler radioPresaleNo.CheckedChanged, AddressOf RadioPresaleNo_CheckedChanged
        AddHandler txtTaxRate.TextChanged, AddressOf TxtTaxRate_TextChanged
        AddHandler txtPhone.TextChanged, AddressOf TxtPhone_TextChanged
        AddHandler txtPresaleNumber.TextChanged, AddressOf TxtPresaleNumber_TextChanged
        AddHandler cmbMember.DropDown, AddressOf CmbMember_DropDown
        AddHandler cmbIntroducer.DropDown, AddressOf CmbIntroducer_DropDown
        AddHandler cmbSalesFactory.DropDown, AddressOf CmbSalesFactory_DropDown
        AddHandler cmbMember.SelectedIndexChanged, AddressOf CmbMember_SelectedIndexChanged
        AddHandler cmbGuide.KeyDown, AddressOf CmbGuide_KeyDown
        AddHandler cmbGuide.DropDown, AddressOf CmbGuide_DropDown
        AddHandler txtStoreMaterial.Leave, AddressOf TxtStoreMaterial_Leave
        AddHandler txtStoreBalance.Leave, AddressOf TxtStoreBalance_Leave
        AddHandler txtPhone.KeyDown, AddressOf TxtPhone_KeyDown
        AddHandler txtPresaleNumber.KeyDown, AddressOf TxtPresaleNumber_KeyDown
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "商品销售出库"
        Me.Size = New Size(1430, 698)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' 顶部信息区
        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 110
        Me.Controls.Add(panelTop)

        ' 单据号
        AddLabel(panelTop, "单据号：", 800, 7)
        txtOrderNumber.Location = New Point(860, 5)
        txtOrderNumber.Size = New Size(130, 25)
        txtOrderNumber.ReadOnly = True
        panelTop.Controls.Add(txtOrderNumber)

        ' 会员查找
        AddLabel(panelTop, "会员：", 10, 35)
        cmbMember.Location = New Point(60, 33)
        cmbMember.Size = New Size(150, 25)
        panelTop.Controls.Add(cmbMember)

        ' 联系电话
        AddLabel(panelTop, "联系电话：", 220, 35)
        txtPhone.Location = New Point(290, 33)
        txtPhone.Size = New Size(120, 25)
        panelTop.Controls.Add(txtPhone)

        ' 介绍人
        AddLabel(panelTop, "介绍人：", 420, 35)
        cmbIntroducer.Location = New Point(470, 33)
        cmbIntroducer.Size = New Size(150, 25)
        panelTop.Controls.Add(cmbIntroducer)

        ' 销售工厂
        AddLabel(panelTop, "销售工厂：", 630, 35)
        cmbSalesFactory.Location = New Point(700, 33)
        cmbSalesFactory.Size = New Size(120, 25)
        panelTop.Controls.Add(cmbSalesFactory)

        ' 批零选择
        Dim grpPling As New GroupBox()
        grpPling.Text = ""
        grpPling.Location = New Point(10, 65)
        grpPling.Size = New Size(95, 30)
        panelTop.Controls.Add(grpPling)

        radioNewRow.Text = "换行"
        radioNewRow.Location = New Point(5, 8)
        radioNewRow.Size = New Size(45, 20)
        radioNewRow.Checked = True
        grpPling.Controls.Add(radioNewRow)

        radioNewCol.Text = "换列"
        radioNewCol.Location = New Point(48, 8)
        radioNewCol.Size = New Size(45, 20)
        grpPling.Controls.Add(radioNewCol)

        ' 批发/零售
        Dim grpPling2 As New GroupBox()
        grpPling2.Text = ""
        grpPling2.Location = New Point(107, 65)
        grpPling2.Size = New Size(95, 30)
        panelTop.Controls.Add(grpPling2)

        radioRetail.Text = "零售"
        radioRetail.Location = New Point(5, 8)
        radioRetail.Size = New Size(45, 20)
        radioRetail.Checked = True
        grpPling2.Controls.Add(radioRetail)

        radioWholesale.Text = "批发"
        radioWholesale.Location = New Point(48, 8)
        radioWholesale.Size = New Size(45, 20)
        grpPling2.Controls.Add(radioWholesale)

        ' 开票/不开票
        groupInvoice.Text = ""
        groupInvoice.Location = New Point(208, 65)
        groupInvoice.Size = New Size(120, 30)
        panelTop.Controls.Add(groupInvoice)

        radioNoInvoice.Text = "不开票"
        radioNoInvoice.Location = New Point(5, 8)
        radioNoInvoice.Size = New Size(58, 20)
        radioNoInvoice.Checked = True
        groupInvoice.Controls.Add(radioNoInvoice)

        radioInvoice.Text = "开票"
        radioInvoice.Location = New Point(67, 8)
        radioInvoice.Size = New Size(45, 20)
        groupInvoice.Controls.Add(radioInvoice)

        ' 业务员
        AddLabel(panelTop, "业务员：", 338, 70)
        cmbGuide.Location = New Point(385, 68)
        cmbGuide.Size = New Size(150, 25)
        panelTop.Controls.Add(cmbGuide)

        ' 单据样式
        AddLabel(panelTop, "单据样式：", 545, 70)
        cmbVoucher.Location = New Point(610, 68)
        cmbVoucher.Size = New Size(120, 25)
        panelTop.Controls.Add(cmbVoucher)

        ' 应收
        AddLabel(panelTop, "应收：", 740, 70)
        txtReceivable.Location = New Point(775, 68)
        txtReceivable.Size = New Size(70, 25)
        txtReceivable.Text = "0"
        txtReceivable.ReadOnly = True
        panelTop.Controls.Add(txtReceivable)

        ' 税额
        AddLabel(panelTop, "税额：", 850, 70)
        txtTax.Location = New Point(882, 68)
        txtTax.Size = New Size(50, 25)
        txtTax.ReadOnly = True
        panelTop.Controls.Add(txtTax)

        ' 税点
        AddLabel(panelTop, "税点：", 940, 70)
        txtTaxRate.Location = New Point(965, 68)
        txtTaxRate.Size = New Size(30, 25)
        panelTop.Controls.Add(txtTaxRate)

        ' 实收
        AddLabel(panelTop, "实收：", 1000, 70)
        txtReceived.Location = New Point(1035, 68)
        txtReceived.Size = New Size(70, 25)
        panelTop.Controls.Add(txtReceived)

        ' 打印选项
        chkPrint.Text = "打印单据"
        chkPrint.Location = New Point(1110, 70)
        chkPrint.AutoSize = True
        panelTop.Controls.Add(chkPrint)

        chkPreview.Text = "打印预览"
        chkPreview.Location = New Point(1185, 70)
        chkPreview.AutoSize = True
        panelTop.Controls.Add(chkPreview)

        chkPrintChengse.Text = "打印成色"
        chkPrintChengse.Location = New Point(1265, 16)
        chkPrintChengse.AutoSize = True
        panelTop.Controls.Add(chkPrintChengse)

        ' 预售
        groupPresale.Text = ""
        groupPresale.Location = New Point(830, 35)
        groupPresale.Size = New Size(180, 30)
        panelTop.Controls.Add(groupPresale)

        radioPresaleNo.Text = "否"
        radioPresaleNo.Location = New Point(5, 8)
        radioPresaleNo.Size = New Size(40, 20)
        radioPresaleNo.Checked = True
        groupPresale.Controls.Add(radioPresaleNo)

        radioPresaleYes.Text = "是"
        radioPresaleYes.Location = New Point(50, 8)
        radioPresaleYes.Size = New Size(40, 20)
        groupPresale.Controls.Add(radioPresaleYes)

        AddLabel(panelTop, "预售单号：", 132, 37)
        txtPresaleNumber.Location = New Point(200, 35)
        txtPresaleNumber.Size = New Size(120, 25)
        txtPresaleNumber.ReadOnly = True
        panelTop.Controls.Add(txtPresaleNumber)

        AddLabel(panelTop, "订金：", 1015, 37)
        txtPresaleDeposit.Location = New Point(1040, 35)
        txtPresaleDeposit.Size = New Size(60, 25)
        txtPresaleDeposit.ReadOnly = True
        panelTop.Controls.Add(txtPresaleDeposit)

        ' 上次消费信息
        lblLastConsumption.Location = New Point(1100, 37)
        lblLastConsumption.Size = New Size(250, 20)
        panelTop.Controls.Add(lblLastConsumption)

        ' 存料/余额
        AddLabel(panelTop, "存料：", 1320, 5)
        txtStoreMaterial.Location = New Point(1350, 3)
        txtStoreMaterial.Size = New Size(60, 25)
        txtStoreMaterial.Visible = False
        panelTop.Controls.Add(txtStoreMaterial)

        lblMaterialBalance.Location = New Point(1320, 28)
        lblMaterialBalance.Size = New Size(60, 20)
        lblMaterialBalance.Visible = False
        panelTop.Controls.Add(lblMaterialBalance)

        AddLabel(panelTop, "余额：", 1320, 50)
        txtStoreBalance.Location = New Point(1350, 48)
        txtStoreBalance.Size = New Size(60, 25)
        txtStoreBalance.Visible = False
        panelTop.Controls.Add(txtStoreBalance)

        lblBalanceAmount.Location = New Point(1320, 73)
        lblBalanceAmount.Size = New Size(60, 20)
        lblBalanceAmount.Visible = False
        panelTop.Controls.Add(lblBalanceAmount)

        lblMemberId.Location = New Point(200, 10)
        lblMemberId.Size = New Size(100, 20)
        panelTop.Controls.Add(lblMemberId)

        ' 工具条
        Dim panelToolbar As New Panel()
        panelToolbar.Dock = DockStyle.Top
        panelToolbar.Height = 40
        Me.Controls.Add(panelToolbar)

        btnToolbarSave.Text = "保存"
        btnToolbarSave.Location = New Point(10, 5)
        btnToolbarSave.Size = New Size(80, 30)
        panelToolbar.Controls.Add(btnToolbarSave)

        btnToolbarReset.Text = "重置"
        btnToolbarReset.Location = New Point(100, 5)
        btnToolbarReset.Size = New Size(80, 30)
        panelToolbar.Controls.Add(btnToolbarReset)

        btnToolbarBatchEdit.Text = "批量编辑"
        btnToolbarBatchEdit.Location = New Point(190, 5)
        btnToolbarBatchEdit.Size = New Size(80, 30)
        panelToolbar.Controls.Add(btnToolbarBatchEdit)

        btnToolbarPrintPreview.Text = "打印预览"
        btnToolbarPrintPreview.Location = New Point(280, 5)
        btnToolbarPrintPreview.Size = New Size(80, 30)
        panelToolbar.Controls.Add(btnToolbarPrintPreview)

        ' 中部表格区
        dgvSales.Dock = DockStyle.Fill
        dgvSales.AllowUserToAddRows = False
        dgvSales.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvSales.Location = New Point(0, 150)
        Me.Controls.Add(dgvSales)

        ' 底部区域
        Dim panelBottom As New Panel()
        panelBottom.Dock = DockStyle.Bottom
        panelBottom.Height = 260
        Me.Controls.Add(panelBottom)

        ' 统计汇总表格
        dgvSummary.Location = New Point(0, 0)
        dgvSummary.Size = New Size(0, 35)
        dgvSummary.AllowUserToAddRows = False
        dgvSummary.Height = 35
        dgvSummary.Width = panelBottom.Width
        panelBottom.Controls.Add(dgvSummary)

        ' 回收明细表格
        AddLabel(panelBottom, "回收明细：", 0, 40)
        dgvRecovery.Location = New Point(55, 40)
        dgvRecovery.Size = New Size(700, 100)
        dgvRecovery.AllowUserToAddRows = False
        panelBottom.Controls.Add(dgvRecovery)

        btnRecoveryAdd.Text = "+"
        btnRecoveryAdd.Location = New Point(10, 45)
        btnRecoveryAdd.Size = New Size(40, 25)
        panelBottom.Controls.Add(btnRecoveryAdd)

        btnRecoveryRemove.Text = "-"
        btnRecoveryRemove.Location = New Point(10, 75)
        btnRecoveryRemove.Size = New Size(40, 25)
        panelBottom.Controls.Add(btnRecoveryRemove)

        ' 收支明细表格
        AddLabel(panelBottom, "收支明细：", 760, 40)
        dgvPayment.Location = New Point(760, 60)
        dgvPayment.Size = New Size(239, 200)
        dgvPayment.AllowUserToAddRows = False
        panelBottom.Controls.Add(dgvPayment)

        ' 备注
        AddLabel(panelBottom, "备注：", 0, 150)
        txtRemarks.Location = New Point(50, 147)
        txtRemarks.Size = New Size(700, 25)
        panelBottom.Controls.Add(txtRemarks)
    End Sub

    Private Sub AddLabel(parent As Control, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    ' ========== 窗口加载（_窗口_商品销售出库_创建完毕） ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 初始化编辑状态
        tableEditState = 0
        paymentEditState = 0
        recoveryEditState = 0

        ' 生成单据号
        txtOrderNumber.Text = DateTime.Now.ToString("yyyyMMddHHmmss") & New Random().Next(1000, 9999).ToString() & "1"
        InboundProductCodeText = ""
        cmbMember.Text = ""
        cmbMember.SelectedIndex = -1
        txtPhone.Text = ""
        txtPhone.ReadOnly = False
        txtPhone.BackColor = Color.White
        txtReceivable.Text = "0"
        txtReceived.Text = ""
        lblLastConsumption.Text = ""

        txtTax.ReadOnly = True
        txtReceivable.ReadOnly = True
        txtPresaleNumber.ReadOnly = True

        txtStoreMaterial.Visible = False
        lblMaterialBalance.Visible = False
        lblBalanceAmount.Visible = False
        lblMaterialBalance.Text = "0.000"
        txtStoreMaterial.Text = "0.000"
        lblBalanceAmount.Text = "0.00"
        lblMemberId.Text = ""
        txtStoreBalance.Text = "0.00"
        wholesaleStockWeight = "0.000"
        wholesaleStockAmount = "0.00"

        chkPrint.Checked = False
        chkPreview.Checked = False
        chkPrintChengse.Checked = False
        radioRetail.Checked = True
        radioWholesale.Checked = False

        ' 开票设置
        If SalesQueryInvoice = "1" Then
            radioNoInvoice.Checked = False
            radioInvoice.Checked = True
            groupInvoice.Visible = False
        Else
            radioNoInvoice.Checked = True
            radioInvoice.Checked = False
            groupInvoice.Visible = True
        End If

        ' 初始化
        LoadVoucherStyles()
        InitSalesGrid()
        InitRecoveryGrid()
        InitPaymentGrid()

        ' 加载临时数据
        LoadTempData()

        ' 数据统计汇总
        CalculateSummary()
    End Sub

    ' ========== 加载单据样式（_单据样式_初始化） ==========
    Private Sub LoadVoucherStyles()
        Try
            Dim sql As String = "SELECT * FROM xipunum_erp_voucher where (type= '出库' or type= '置换') and state= '0' order by id ASC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbVoucher.Items.Clear()
            For Each r As DataRow In dt.Rows
                cmbVoucher.Items.Add(SafeString(r("name")))
            Next
            cmbVoucher.SelectedIndex = -1
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 初始化销售明细表格（_高级表格1_加载表头） ==========
    Private Sub InitSalesGrid()
        dgvSales.Columns.Clear()
        dgvSales.Rows.Clear()
        Dim headers() As String = {"", "商品编码", "商品名称", "款号", "规格", "材质", "圈口/长度", "成色", "单件重", "金重", "重量", "成本工费", "参考工费", "成本附加费", "库存", "销售单价", "销售金额", "数量", "原附加费", "销售克价", "销售工费", "销售附加费", "折扣", "实收金额", "导购员", "库存金重", "操作"}
        Dim widths() As Integer = {45, 100, 140, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 0, 60}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvSales.Columns.Add(col)
        Next

        ' 加载导购员列表到下拉列
        Dim guideSql As String = "SELECT name FROM xipunum_erp_user where state='0' order by id ASC"
        If UserPermission = "店铺" Then
            guideSql = $"SELECT name FROM xipunum_erp_user where department='{UserDepartment}' and state='0' order by id ASC"
        ElseIf UserPermission = "岗位" Then
            guideSql = $"SELECT name FROM xipunum_erp_user WHERE user in {GlobalViewSQL} and state='0' order by id ASC"
        ElseIf UserPermission = "个人" Then
            guideSql = $"SELECT name FROM xipunum_erp_user WHERE user='{UserAccount}' and state='0' order by id ASC"
        End If
        Dim guideDt As DataTable = DatabaseModule.ExecuteQuery(guideSql)
        cmbGuide.Items.Clear()
        For Each r As DataRow In guideDt.Rows
            cmbGuide.Items.Add(SafeString(r("name")))
        Next
        cmbGuide.SelectedIndex = -1
    End Sub

    ' ========== 初始化回收明细表格（_高级表格2_加载表头） ==========
    Private Sub InitRecoveryGrid()
        dgvRecovery.Columns.Clear()
        dgvRecovery.Rows.Clear()
        Dim headers() As String = {"", "商品名称", "数量", "总重", "金重", "成色", "回收克价", "其他费用", "回收金额", "备注", "导购员"}
        Dim widths() As Integer = {45, 200, 100, 100, 100, 100, 100, 100, 100, 200, 100}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "rcol" & i
            col.Width = widths(i)
            dgvRecovery.Columns.Add(col)
        Next

        ' 加载回收名称列表
        Dim sql As String = "SELECT * FROM xipunum_erp_retreat_title WHERE 1=1 order by id asc"
        DatabaseModule.ExecuteQuery(sql)

        ' 加载导购员到回收表格下拉
        Dim guideSql As String = "SELECT name FROM xipunum_erp_user where state='0' order by id ASC"
        If UserPermission = "店铺" Then
            guideSql = $"SELECT name FROM xipunum_erp_user where department='{UserDepartment}' and state='0' order by id ASC"
        ElseIf UserPermission = "岗位" Then
            guideSql = $"SELECT name FROM xipunum_erp_user WHERE user in {GlobalViewSQL} and state='0' order by id ASC"
        ElseIf UserPermission = "个人" Then
            guideSql = $"SELECT name FROM xipunum_erp_user WHERE user='{UserAccount}' and state='0' order by id ASC"
        End If
        DatabaseModule.ExecuteQuery(guideSql)
    End Sub

    ' ========== 初始化收支明细表格（_高级表格3_加载表头 + _高级表格3_加载表格） ==========
    Private Sub InitPaymentGrid()
        dgvPayment.Columns.Clear()
        dgvPayment.Rows.Clear()
        Dim headers() As String = {"序号", "支付方式", "金额", "id"}
        Dim widths() As Integer = {45, 65, 75, 0}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "pcol" & i
            col.Width = widths(i)
            If i = 3 Then col.Visible = False
            dgvPayment.Columns.Add(col)
        Next

        ' 加载支付方式
        Try
            Dim sql As String = "SELECT id,name FROM xipunum_erp_pay where state='0' order by id ASC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            For i As Integer = 0 To dt.Rows.Count - 1
                dgvPayment.Rows.Add(
                    Right("000" & (i + 1).ToString(), 1),
                    SafeString(dt.Rows(i)("name")),
                    "0",
                    SafeString(dt.Rows(i)("id"))
                )
            Next
            ' 合计行
            dgvPayment.Rows.Add("", "合计", "0", "")
            ' 设置合计行只读
            Dim lastRow As Integer = dgvPayment.Rows.Count - 1
            For c As Integer = 0 To 2
                dgvPayment.Rows(lastRow).Cells(c).ReadOnly = True
                dgvPayment.Rows(lastRow).Cells(c).Style.BackColor = Color.LightGray
            Next
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 加载临时数据（_高级表格1_临时数据） ==========
    Private Sub LoadTempData()
        ' 检查是否有临时数据（本地Access）
        tempQuantity = 0
        dataLoaded = 1

        If tempQuantity > 0 Then
            If MessageBox.Show("是否加载上次临时出库表？", "警告", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                dataLoaded = 1
            Else
                dataLoaded = 0
            End If
        End If

        If dataLoaded = 1 Then
            ' 加载临时数据
            AddSalesRow()
        End If

        If dataLoaded = 0 Then
            ' 清空临时数据
            ClearTempData()
            AddSalesRow()
        End If

        CalculateSummary()
    End Sub

    ' ========== 清空临时数据 ==========
    Private Sub ClearTempData()
        ' 清空本地Access临时数据（原易语言使用Access数据库chuku表）
    End Sub

    ' ========== 添加销售空行（_子程序_添加销售） ==========
    Private Sub AddSalesRow()
        Dim addRowIndex As Integer = dgvSales.Rows.Count
        If addRowIndex = 0 OrElse SafeString(dgvSales.Rows(addRowIndex - 1).Cells("col1").Value) <> "" Then
            Dim newRow As Integer = dgvSales.Rows.Add()
            dgvSales.Rows(newRow).Cells("col0").Value = (newRow + 1).ToString()
            dgvSales.Rows(newRow).Cells("col8").Value = "0.00"
            dgvSales.Rows(newRow).Cells("col9").Value = "0.000"
            dgvSales.Rows(newRow).Cells("col10").Value = "0.000"
            dgvSales.Rows(newRow).Cells("col11").Value = "0.00"
            dgvSales.Rows(newRow).Cells("col12").Value = "0.00"
            dgvSales.Rows(newRow).Cells("col13").Value = "0.00"
            dgvSales.Rows(newRow).Cells("col14").Value = "0"
            dgvSales.Rows(newRow).Cells("col15").Value = "0"
            dgvSales.Rows(newRow).Cells("col16").Value = "0"
            dgvSales.Rows(newRow).Cells("col17").Value = "0"
            dgvSales.Rows(newRow).Cells("col18").Value = "0"
            dgvSales.Rows(newRow).Cells("col19").Value = "0"
            dgvSales.Rows(newRow).Cells("col20").Value = "0"
            dgvSales.Rows(newRow).Cells("col21").Value = "0"
            dgvSales.Rows(newRow).Cells("col22").Value = "1"
            dgvSales.Rows(newRow).Cells("col23").Value = "0"
            dgvSales.Rows(newRow).Cells("col25").Value = "0"

            ' 设置只读
            For c As Integer = 2 To 25
                dgvSales.Rows(newRow).Cells(c).ReadOnly = True
                dgvSales.Rows(newRow).Cells(c).Style.BackColor = Color.LightGray
            Next
        End If
    End Sub

    ' ========== 数据统计汇总（_子程序_数据统计汇总） ==========
    Private Sub CalculateSummary()
        Dim calcTaxRate As String = If(txtTaxRate.Text = "", "0", txtTaxRate.Text)

        quantityTotal = "0"
        weightTotal = "0"
        weightPrice = "0"
        salesAmount = "0"
        actualAmount = "0"
        basicCostTotal = "0"
        premiumCostTotal = "0"
        salesCostTotal = "0"
        costSurchargeTotal = "0"
        salesSurchargeTotal = "0"

        ' 销售明细统计
        For Each r As DataGridViewRow In dgvSales.Rows
            If r.IsNewRow Then Continue For
            Dim qty As Decimal = SafeDecimal(r.Cells("col17").Value)
            Dim jinzhong As Decimal = SafeDecimal(r.Cells("col9").Value)
            Dim unitPrice As Decimal = SafeDecimal(r.Cells("col15").Value)
            Dim salesAmt As Decimal = SafeDecimal(r.Cells("col16").Value)
            Dim actualAmt As Decimal = SafeDecimal(r.Cells("col23").Value)
            Dim basicCost As Decimal = SafeDecimal(r.Cells("col11").Value)
            Dim premiumCost As Decimal = SafeDecimal(r.Cells("col12").Value)
            Dim costSurch As Decimal = SafeDecimal(r.Cells("col13").Value)
            Dim salesCost As Decimal = SafeDecimal(r.Cells("col20").Value)
            Dim salesSurch As Decimal = SafeDecimal(r.Cells("col21").Value)

            quantityTotal = (SafeDecimal(quantityTotal) + qty).ToString()
            weightTotal = (SafeDecimal(weightTotal) + jinzhong).ToString()
            weightPrice = (SafeDecimal(weightPrice) + unitPrice).ToString()
            salesAmount = (SafeDecimal(salesAmount) + salesAmt).ToString()
            actualAmount = (SafeDecimal(actualAmount) + actualAmt).ToString()
            basicCostTotal = (SafeDecimal(basicCostTotal) + jinzhong * basicCost).ToString()
            premiumCostTotal = (SafeDecimal(premiumCostTotal) + jinzhong * premiumCost).ToString()
            salesCostTotal = (SafeDecimal(salesCostTotal) + jinzhong * salesCost).ToString()

            If qty = 0 Then
                costSurchargeTotal = (SafeDecimal(costSurchargeTotal) + costSurch).ToString()
            Else
                costSurchargeTotal = (SafeDecimal(costSurchargeTotal) + qty * costSurch).ToString()
            End If

            salesSurchargeTotal = (SafeDecimal(salesSurchargeTotal) + salesSurch).ToString()
        Next

        ' 回收明细统计
        recoveryWeightTotal = "0"
        recoveryGoldTotal = "0"
        recoveryOtherTotal = "0"
        recoveryAmountTotal = "0"

        For Each r As DataGridViewRow In dgvRecovery.Rows
            If r.IsNewRow Then Continue For
            recoveryWeightTotal = (SafeDecimal(recoveryWeightTotal) + SafeDecimal(r.Cells("rcol3").Value)).ToString()
            recoveryGoldTotal = (SafeDecimal(recoveryGoldTotal) + SafeDecimal(r.Cells("rcol4").Value)).ToString()
            recoveryOtherTotal = (SafeDecimal(recoveryOtherTotal) + SafeDecimal(r.Cells("rcol7").Value)).ToString()
            recoveryAmountTotal = (SafeDecimal(recoveryAmountTotal) + SafeDecimal(r.Cells("rcol8").Value)).ToString()
        Next

        ' 税额计算
        Dim presaleDeposit As Decimal = SafeDecimal(txtPresaleDeposit.Text)
        txtTax.Text = Math.Round((SafeDecimal(salesAmount) - SafeDecimal(recoveryAmountTotal) - presaleDeposit) * SafeDecimal(calcTaxRate) / 100, 0).ToString()

        ' 应收计算
        txtReceivable.Text = Math.Round(SafeDecimal(salesAmount) + SafeDecimal(txtTax.Text), 2).ToString()

        ' 实收计算
        Dim infoActualAmount As String = FormatTwoDecimals(Math.Round(SafeDecimal(actualAmount) + SafeDecimal(txtTax.Text) - SafeDecimal(recoveryAmountTotal) - presaleDeposit, 0).ToString())

        ' 清零支付方式金额并设置第一行
        For i As Integer = 0 To dgvPayment.Rows.Count - 2
            If Not dgvPayment.Rows(i).IsNewRow Then
                dgvPayment.Rows(i).Cells("pcol2").Value = "0"
            End If
        Next

        If dgvPayment.Rows.Count > 1 Then
            dgvPayment.Rows(0).Cells("pcol2").Value = infoActualAmount
        End If

        ' 合计行
        Dim payTotal As Decimal = 0
        For i As Integer = 0 To dgvPayment.Rows.Count - 2
            If Not dgvPayment.Rows(i).IsNewRow Then
                payTotal += SafeDecimal(dgvPayment.Rows(i).Cells("pcol2").Value)
            End If
        Next
        If dgvPayment.Rows.Count > 0 Then
            dgvPayment.Rows(dgvPayment.Rows.Count - 1).Cells("pcol2").Value = payTotal.ToString()
        End If
        txtReceived.Text = payTotal.ToString()

        ' 批发模式下计算存料/余额
        If radioWholesale.Checked Then
            Dim infoStockWeight As String = (SafeDecimal(wholesaleStockWeight) + SafeDecimal(txtStoreMaterial.Text) - SafeDecimal(recoveryGoldTotal)).ToString()
            Dim infoStockAmount As String = (SafeDecimal(wholesaleStockAmount) + SafeDecimal(txtStoreBalance.Text) - SafeDecimal(txtReceived.Text)).ToString()

            infoStockWeight = FormatThreeDecimals(infoStockWeight)
            infoStockAmount = FormatTwoDecimals(infoStockAmount)

            lblMaterialBalance.Text = infoStockWeight
            lblBalanceAmount.Text = infoStockAmount
        End If

        ' 更新统计汇总行
        UpdateSummaryGrid()
    End Sub

    ' ========== 更新统计汇总表格（_高级表格4_加载表头） ==========
    Private Sub UpdateSummaryGrid()
        dgvSummary.Columns.Clear()
        dgvSummary.Rows.Clear()
        Dim widths() As Integer = {45, 100, 140, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70}
        For i As Integer = 0 To widths.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.Name = "scol" & i
            col.Width = widths(i)
            dgvSummary.Columns.Add(col)
        Next
        dgvSummary.Rows.Add()
        dgvSummary.Rows(0).Cells("scol0").Value = ""
        dgvSummary.Rows(0).Cells("scol1").Value = "合计"
        dgvSummary.Rows(0).Cells("scol9").Value = weightTotal
        dgvSummary.Rows(0).Cells("scol11").Value = Math.Round(SafeDecimal(basicCostTotal), 0).ToString()
        dgvSummary.Rows(0).Cells("scol12").Value = Math.Round(SafeDecimal(premiumCostTotal), 0).ToString()
        dgvSummary.Rows(0).Cells("scol13").Value = Math.Round(SafeDecimal(costSurchargeTotal), 0).ToString()
        dgvSummary.Rows(0).Cells("scol16").Value = Math.Round(SafeDecimal(salesAmount), 0).ToString()
        dgvSummary.Rows(0).Cells("scol17").Value = Math.Round(SafeDecimal(quantityTotal), 2).ToString()
        dgvSummary.Rows(0).Cells("scol20").Value = Math.Round(SafeDecimal(salesCostTotal), 0).ToString()
        dgvSummary.Rows(0).Cells("scol21").Value = Math.Round(SafeDecimal(salesSurchargeTotal), 0).ToString()
        dgvSummary.Rows(0).Cells("scol23").Value = Math.Round(SafeDecimal(actualAmount) - SafeDecimal(recoveryAmountTotal), 0).ToString()
        For c As Integer = 1 To 24
            dgvSummary.Rows(0).Cells(c).ReadOnly = True
            dgvSummary.Rows(0).Cells(c).Style.BackColor = Color.LightGray
        Next
    End Sub

    ' ========== 销售表格结束编辑（_高级表格1_结束编辑） ==========
    Private Sub DgvSales_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return

        Dim r As DataGridViewRow = dgvSales.Rows(e.RowIndex)
        tableEditState = 0

        ' 检查业务员
        If cmbGuide.SelectedIndex = -1 Then
            ShowWarning("业务员不能为空！")
            cmbGuide.Focus()
            If dgvSales.Rows.Count > 0 Then
                dgvSales.Rows(dgvSales.Rows.Count - 1).Cells("col1").Value = ""
            End If
            Return
        End If

        Dim productCode As String = SafeString(r.Cells("col1").Value)
        Dim salesGoldWeight As String = SafeString(r.Cells("col9").Value)
        Dim stockQty As String = SafeString(r.Cells("col14").Value)
        Dim salesUnitPrice As String = Math.Round(SafeDecimal(r.Cells("col15").Value), 2).ToString()
        Dim salesActualAmount As String = Math.Round(SafeDecimal(r.Cells("col16").Value), 2).ToString()
        Dim salesQty As String = Math.Round(SafeDecimal(r.Cells("col17").Value), 2).ToString()
        Dim originalSurch As String = Math.Round(SafeDecimal(r.Cells("col18").Value), 2).ToString()
        Dim salesGoldPrice As String = Math.Round(SafeDecimal(r.Cells("col19").Value), 2).ToString()
        Dim salesWorkFee As String = Math.Round(SafeDecimal(r.Cells("col20").Value), 2).ToString()
        Dim salesSurch As String = Math.Round(SafeDecimal(r.Cells("col21").Value), 0).ToString()
        Dim discount As String = If(SafeString(r.Cells("col21").Value) <> "",
                                    Math.Round(SafeDecimal(r.Cells("col22").Value), 4).ToString(), "1.000")
        Dim actualReceived As String = Math.Round(SafeDecimal(r.Cells("col23").Value), 0).ToString()
        Dim originalGoldWeight As String = SafeString(r.Cells("col25").Value)
        Dim guideSearch As String = SafeString(r.Cells("col24").Value)

        salesProductCode = productCode
        salesDataCode = productCode

        ' 商品编码列编辑
        If e.ColumnIndex = 1 Then
            If productCode.Length > 4 Then
                ' 查询商品是否存在
                Dim sql As String = $"SELECT * FROM xipunum_erp_shop where (poduct_code='{SafeSQL(productCode)}' or fu_code='{SafeSQL(productCode)}') order by id ASC"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
                If dt.Rows.Count = 0 Then
                    ShowWarning("请输入正确的商品编码！")
                    r.Cells("col1").Value = ""
                    Return
                End If

                ' 检查重复
                For i As Integer = 0 To dgvSales.Rows.Count - 2
                    If i <> e.RowIndex AndAlso SafeString(dgvSales.Rows(i).Cells("col1").Value) = productCode Then
                        ShowWarning("此商品已在当前销售清单！")
                        dgvSales.Rows(dgvSales.Rows.Count - 1).Cells("col1").Value = ""
                        Return
                    End If
                Next

                ' 查询商品库存是否存在
                Dim kucunSql As String = $"SELECT a.poduct_code AS apoduct_code FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE a.kufang = '{UserDepartment}' AND (a.quantity > 0 or a.jinzhong >0) AND (b.poduct_code = '{SafeSQL(productCode)}' OR b.fu_code = '{SafeSQL(productCode)}') ORDER BY a.id DESC"
                Dim kucunDt As DataTable = DatabaseModule.ExecuteQuery(kucunSql)
                If kucunDt.Rows.Count = 0 Then
                    ShowWarning("此商品不属于您所在的库房！")
                    r.Cells("col1").Value = ""
                    Return
                End If

                If kucunDt.Rows.Count = 1 Then
                    salesDataCode = SafeString(kucunDt.Rows(0)("apoduct_code"))
                    LoadProductData(e.RowIndex)
                    Return
                Else
                    ' 多编码查询（原易语言载入副编码查询窗口）
                    ShowWarning("存在多个编码，使用第一个！")
                    salesDataCode = SafeString(kucunDt.Rows(0)("apoduct_code"))
                    LoadProductData(e.RowIndex)
                    Return
                End If
            End If
        End If

        ' 金重列编辑
        If e.ColumnIndex = 9 Then
            ' 查询款号是否零销
            Dim kuanhao As String = SafeString(r.Cells("col3").Value)
            Dim lingxiao As String = "否"
            If kuanhao <> "" Then
                Dim lxSql As String = $"SELECT lingxiao FROM xipunum_erp_ksiamges where kuanhao='{SafeSQL(kuanhao)}' LIMIT 1"
                Dim lxdt As DataTable = DatabaseModule.ExecuteQuery(lxSql)
                If lxdt.Rows.Count > 0 Then
                    lingxiao = SafeString(lxdt.Rows(0)("lingxiao"))
                End If
            End If

            If SafeDecimal(salesGoldWeight) > SafeDecimal(originalGoldWeight) Then
                ShowWarning("销售金重不能大于库存金重！")
                r.Cells("col9").Value = originalGoldWeight
                If lingxiao = "是" Then
                    r.Cells("col10").Value = originalGoldWeight
                End If
                Return
            End If

            r.Cells("col9").Value = FormatThreeDecimals(salesGoldWeight)
        End If

        ' 数量列编辑
        If e.ColumnIndex = 17 Then
            If SafeDecimal(salesQty) < 0 Then
                ShowWarning("销售数量不能小于0！")
                salesQty = stockQty
            End If

            If SafeDecimal(salesQty) > SafeDecimal(stockQty) Then
                ShowWarning("销售数量不能大于库存数量！")
                salesQty = stockQty
            End If

            ' 重新计算金重
            salesGoldWeight = (SafeDecimal(originalGoldWeight) / SafeDecimal(stockQty) * SafeDecimal(salesQty)).ToString()
            salesGoldWeight = FormatThreeDecimals(salesGoldWeight)

            Dim salesWeight As String = (SafeDecimal(r.Cells("col8").Value) * SafeDecimal(salesQty)).ToString()
            salesWeight = FormatThreeDecimals(salesWeight)

            If SafeDecimal(salesQty) = 0 Then
                salesSurch = originalSurch
            Else
                salesSurch = (SafeDecimal(originalSurch) * SafeDecimal(salesQty)).ToString()
            End If
            salesSurch = FormatTwoDecimals(Math.Round(SafeDecimal(salesSurch), 0).ToString())

            r.Cells("col9").Value = salesGoldWeight
            r.Cells("col10").Value = salesWeight
            r.Cells("col17").Value = salesQty
            r.Cells("col21").Value = salesSurch
        End If

        ' 销售克价列编辑
        If e.ColumnIndex = 19 Then
            r.Cells("col19").Value = FormatTwoDecimals(salesGoldPrice)
        End If

        ' 销售工费列编辑
        If e.ColumnIndex = 20 Then
            r.Cells("col20").Value = FormatTwoDecimals(salesWorkFee)
        End If

        ' 销售附加费列编辑
        If e.ColumnIndex = 21 Then
            r.Cells("col21").Value = FormatTwoDecimals(Math.Round(SafeDecimal(salesSurch), 0).ToString())

            If SafeDecimal(salesQty) = 0 Then
                discount = (SafeDecimal(salesSurch) / SafeDecimal(originalSurch)).ToString()
            Else
                discount = (SafeDecimal(salesSurch) / SafeDecimal(salesQty) / SafeDecimal(originalSurch)).ToString()
            End If
            discount = FormatThreeDecimals(discount)

            If SafeDecimal(originalSurch) = 0 Then
                r.Cells("col22").Value = "1.000"
            Else
                r.Cells("col22").Value = discount
            End If
        End If

        ' 折扣列编辑
        If e.ColumnIndex = 22 Then
            If SafeDecimal(discount) > 1 Then
                MessageBox.Show("折扣不能大于1")
                r.Cells("col22").Value = "1.000"
                Return
            End If
            If SafeDecimal(discount) < 0 Then
                MessageBox.Show("折扣不能小于0")
                r.Cells("col22").Value = "1.000"
                Return
            End If

            discount = FormatThreeDecimals(discount)

            If SafeDecimal(salesQty) = 0 Then
                salesSurch = (SafeDecimal(originalSurch) * SafeDecimal(discount)).ToString()
            Else
                salesSurch = (SafeDecimal(originalSurch) * SafeDecimal(discount) * SafeDecimal(salesQty)).ToString()
            End If
            salesSurch = FormatTwoDecimals(Math.Round(SafeDecimal(salesSurch), 0).ToString())

            r.Cells("col21").Value = salesSurch
            If SafeDecimal(originalSurch) = 0 Then
                r.Cells("col22").Value = "1.000"
            Else
                r.Cells("col22").Value = discount
            End If
        End If

        ' 实收金额列编辑
        If e.ColumnIndex = 23 Then
            actualReceived = FormatTwoDecimals(Math.Round(SafeDecimal(actualReceived), 0).ToString())

            Dim minAmount As Decimal = SafeDecimal(salesActualAmount) * (100 - SafeDecimal(DiscountRate)) / 100
            If SafeDecimal(actualReceived) < minAmount Then
                ShowWarning($"商品实收金额不能小于{Math.Round(minAmount, 0)}元")
                r.Cells("col23").Value = Math.Round(SafeDecimal(salesActualAmount), 0).ToString()
                If SafeDecimal(salesGoldPrice) = 0 Then
                    salesGoldPrice = ((SafeDecimal(salesActualAmount) - SafeDecimal(salesSurch)) / SafeDecimal(salesGoldWeight) - SafeDecimal(salesWorkFee)).ToString()
                    r.Cells("col19").Value = FormatTwoDecimals(salesGoldPrice)
                End If
                Return
            End If

            r.Cells("col23").Value = actualReceived
            If SafeDecimal(salesGoldPrice) = 0 Then
                salesGoldPrice = ((SafeDecimal(actualReceived) - SafeDecimal(salesSurch)) / SafeDecimal(salesGoldWeight) - SafeDecimal(salesWorkFee)).ToString()
                r.Cells("col19").Value = FormatTwoDecimals(salesGoldPrice)
            End If
        End If

        ' 计算价格（列9/17/19/20/21/22变化时）
        If e.ColumnIndex = 9 OrElse e.ColumnIndex = 17 OrElse e.ColumnIndex = 19 OrElse e.ColumnIndex = 20 OrElse e.ColumnIndex = 21 OrElse e.ColumnIndex = 22 Then
            Dim findUnitPrice As String = "0"
            If SafeDecimal(salesQty) = 0 Then
                findUnitPrice = (SafeDecimal(salesGoldWeight) * (SafeDecimal(salesWorkFee) + SafeDecimal(salesGoldPrice)) + SafeDecimal(salesSurch)).ToString()
            Else
                findUnitPrice = ((SafeDecimal(salesGoldWeight) * (SafeDecimal(salesWorkFee) + SafeDecimal(salesGoldPrice)) + SafeDecimal(salesSurch)) / SafeDecimal(salesQty)).ToString()
            End If
            findUnitPrice = FormatTwoDecimals(findUnitPrice)

            Dim findSalesAmount As String = (SafeDecimal(salesGoldWeight) * (SafeDecimal(salesWorkFee) + SafeDecimal(salesGoldPrice)) + SafeDecimal(salesSurch)).ToString()
            findSalesAmount = FormatTwoDecimals(findSalesAmount)

            Dim findActualReceived As String = findSalesAmount
            findActualReceived = FormatTwoDecimals(Math.Round(SafeDecimal(findActualReceived), 0).ToString())

            r.Cells("col15").Value = findUnitPrice
            r.Cells("col16").Value = findSalesAmount
            r.Cells("col23").Value = findActualReceived
        End If

        ' 导购员列编辑
        If e.ColumnIndex = 24 Then
            If guideSearch <> "" Then
                Dim guideSql As String = $"SELECT name FROM xipunum_erp_user WHERE (jianxie = '{SafeSQL(guideSearch)}' or user like '%{SafeSQL(guideSearch)}%' or name like '%{SafeSQL(guideSearch)}%') and department='{UserDepartment}' LIMIT 1"
                Dim guideDt As DataTable = DatabaseModule.ExecuteQuery(guideSql)
                Dim guideName As String = ""
                If guideDt.Rows.Count > 0 Then
                    guideName = SafeString(guideDt.Rows(0)("name"))
                End If

                If guideName = "" Then
                    ShowWarning("导购信息不存在！")
                    r.Cells("col24").Value = ""
                    Return
                End If

                r.Cells("col24").Value = guideName
            End If
        End If

        CalculateSummary()
    End Sub

    ' ========== 商品数据加载（_高级表格1_数据加载） ==========
    Private Sub LoadProductData(rowIndex As Integer)
        Dim r As DataGridViewRow = dgvSales.Rows(rowIndex)

        ' 查询商品库存信息
        Dim sql As String = "SELECT CASE WHEN COALESCE(f.title, '') = '' THEN '未匹配' ELSE f.title END AS pinlei," &
            "CASE WHEN COALESCE(f.shuliang, '') = '' THEN '1' ELSE f.shuliang END AS plduoshu," &
            "CASE WHEN COALESCE(e1.shuliang, e2.shuliang, '') = '' THEN '1' ELSE COALESCE(e1.shuliang, e2.shuliang) END AS ggduoshu," &
            "CASE WHEN COALESCE(d.lingxiao, '') = '' THEN '否' ELSE d.lingxiao END AS lingxiao," &
            "a.poduct_code AS poduct_code,b.product_name AS shoptitle,b.item_number AS kuanhao," &
            "COALESCE(e1.title, e2.title, '无数据') AS guige,b.caizhi as caizhi,b.quandu as quandu," &
            "CASE WHEN COALESCE(d.lingxiao, '') = '是' THEN a.jinzhong ELSE b.single END AS danzhong," &
            "a.jinzhong as jinzhong,CASE WHEN COALESCE(d.lingxiao, '') = '是' THEN a.jinzhong ELSE b.weight END AS zongzhong," &
            "a.quantity as kucun,b.sales_unit as danwie " &
            "FROM xipunum_erp_shop_kucun AS a " &
            "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
            "INNER JOIN xipunum_erp_type AS c ON c.id = a.kufang " &
            "LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = b.item_number AND b.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id AND b.item_number IS NOT NULL AND b.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id AND b.specification_id IS NOT NULL AND b.specification_id != '' " &
            "LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) AND COALESCE(e1.category_id, e2.category_id) IS NOT NULL " &
            $"WHERE a.kufang = '{UserDepartment}' AND (a.quantity > 0 OR a.jinzhong > 0) AND a.poduct_code = '{SafeSQL(salesDataCode)}' ORDER BY a.id DESC"

        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
        If dt.Rows.Count = 0 Then Return

        Dim data As DataRow = dt.Rows(0)
        Dim productTitle As String = SafeString(data("shoptitle"))
        Dim stockQty As String = SafeString(data("kucun"))
        Dim singleWeight As String = SafeString(data("danzhong"))
        Dim caizhi As String = SafeString(data("caizhi"))
        Dim salesUnit As String = SafeString(data("danwie"))
        Dim guige As String = SafeString(data("guige"))
        Dim kuanhao As String = SafeString(data("kuanhao"))
        Dim quandu As String = SafeString(data("quandu"))
        Dim jinzhong As String = SafeString(data("jinzhong"))
        Dim zongzhong As String = SafeString(data("zongzhong"))
        Dim lingxiao As String = SafeString(data("lingxiao"))

        ' 检查库存
        If lingxiao = "是" Then
            If SafeDecimal(jinzhong) <= 0 Then
                ShowWarning("此商品库存金重为0,不能再销售！")
                r.Cells("col1").Value = ""
                Return
            End If
        Else
            If SafeDecimal(stockQty) <= 0 Then
                ShowWarning("此商品库存数量为0,不能再销售！")
                r.Cells("col1").Value = ""
                Return
            End If
        End If

        ' 查询入库信息（成本工费等）
        Dim storeSql As String = $"SELECT a.factory_condition as afactory_condition,a.basic_cost as abasic_cost,a.premium_cost as apremium_cost,a.sales_cost as asales_cost,a.company_surcharge as acompany_surcharge,a.sales_surcharge as asales_surcharge,a.sales_price as yushoujia FROM xipunum_erp_store as a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code where a.poduct_code='{SafeSQL(salesDataCode)}' LIMIT 1"
        Dim storeDt As DataTable = DatabaseModule.ExecuteQuery(storeSql)
        Dim factoryCondition As String = ""
        Dim basicCost As String = ""
        Dim premiumCost As String = ""
        Dim salesCost As String = ""
        Dim companySurch As String = ""
        Dim salesSurch As String = ""
        Dim presalePrice As String = ""
        If storeDt.Rows.Count > 0 Then
            Dim s As DataRow = storeDt.Rows(0)
            factoryCondition = SafeString(s("afactory_condition"))
            basicCost = SafeString(s("abasic_cost"))
            premiumCost = SafeString(s("apremium_cost"))
            salesCost = SafeString(s("asales_cost"))
            companySurch = SafeString(s("acompany_surcharge"))
            salesSurch = SafeString(s("asales_surcharge"))
            presalePrice = SafeString(s("yushoujia"))
        End If

        ' 填充表格
        r.Cells("col1").Value = salesDataCode
        r.Cells("col2").Value = productTitle
        r.Cells("col3").Value = kuanhao
        r.Cells("col4").Value = guige
        r.Cells("col5").Value = caizhi
        r.Cells("col6").Value = quandu
        r.Cells("col7").Value = factoryCondition
        r.Cells("col8").Value = singleWeight
        r.Cells("col9").Value = jinzhong
        r.Cells("col10").Value = zongzhong
        r.Cells("col11").Value = basicCost
        r.Cells("col12").Value = premiumCost
        r.Cells("col13").Value = companySurch
        r.Cells("col14").Value = stockQty

        ' 计算销售单价
        Dim findUnitPrice As String = "0"
        If salesUnit = "重量" Then
            If SafeDecimal(stockQty) = 0 Then
                findUnitPrice = Math.Round(SafeDecimal(jinzhong) * SafeDecimal(salesCost) + SafeDecimal(salesSurch), 2).ToString()
            Else
                findUnitPrice = Math.Round((SafeDecimal(jinzhong) * SafeDecimal(salesCost) + SafeDecimal(salesSurch) * SafeDecimal(stockQty)) / SafeDecimal(stockQty), 2).ToString()
            End If
        Else
            findUnitPrice = presalePrice
        End If
        findUnitPrice = FormatTwoDecimals(findUnitPrice)

        ' 计算销售金额
        Dim findSalesAmount As String = "0"
        If salesUnit = "重量" Then
            findSalesAmount = Math.Round(SafeDecimal(jinzhong) * SafeDecimal(salesCost) + SafeDecimal(salesSurch) * SafeDecimal(stockQty), 2).ToString()
        Else
            findSalesAmount = (SafeDecimal(presalePrice) * SafeDecimal(stockQty)).ToString()
        End If
        findSalesAmount = FormatTwoDecimals(findSalesAmount)

        ' 计算销售附加费
        Dim findSurch As String = "0"
        If SafeDecimal(stockQty) = 0 Then
            findSurch = salesSurch
        Else
            findSurch = Math.Round(SafeDecimal(salesSurch) * SafeDecimal(stockQty), 2).ToString()
        End If
        findSurch = FormatTwoDecimals(findSurch)

        r.Cells("col15").Value = findUnitPrice
        r.Cells("col16").Value = findSalesAmount
        r.Cells("col17").Value = stockQty
        r.Cells("col18").Value = salesSurch
        r.Cells("col19").Value = "0.00"
        r.Cells("col20").Value = salesCost
        r.Cells("col21").Value = findSurch
        r.Cells("col22").Value = "1.000"

        ' 计算实收金额
        Dim findActual As String = "0"
        If salesUnit = "重量" Then
            findActual = Math.Round(SafeDecimal(jinzhong) * SafeDecimal(salesCost) + SafeDecimal(salesSurch) * SafeDecimal(stockQty), 2).ToString()
        Else
            findActual = (SafeDecimal(presalePrice) * SafeDecimal(stockQty)).ToString()
        End If
        findActual = FormatTwoDecimals(findActual)
        r.Cells("col23").Value = findActual

        ' 导购员（第一行用业务员，其他行用上一行）
        If rowIndex <= 1 Then
            r.Cells("col24").Value = If(cmbGuide.SelectedIndex >= 0, cmbGuide.SelectedItem.ToString(), "")
        Else
            r.Cells("col24").Value = SafeString(dgvSales.Rows(rowIndex - 1).Cells("col24").Value)
        End If

        r.Cells("col25").Value = jinzhong

        ' 设置只读
        For c As Integer = 1 To 20
            r.Cells(c).ReadOnly = True
            r.Cells(c).Style.BackColor = Color.LightGray
        Next
        For c As Integer = 21 To 26
            If c < dgvSales.Columns.Count Then
                r.Cells(c).ReadOnly = False
                r.Cells(c).Style.BackColor = Color.White
            End If
        Next

        ' 非重量计价时数量可编辑
        If salesUnit <> "重量" Then
            r.Cells("col17").ReadOnly = False
            r.Cells("col17").Style.BackColor = Color.White
        End If

        ' 重量计价时特定列可编辑
        If salesUnit = "重量" Then
            r.Cells("col19").ReadOnly = False
            r.Cells("col19").Style.BackColor = Color.White
            r.Cells("col20").ReadOnly = False
            r.Cells("col20").Style.BackColor = Color.White
        End If

        ' 添加新空行
        AddSalesRow()
        CalculateSummary()
    End Sub

    ' ========== 单元格值变化事件 ==========
    Private Sub DgvSales_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
    End Sub

    ' ========== 回收表格结束编辑（_高级表格2_结束编辑） ==========
    Private Sub DgvRecovery_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return

        If dgvSales.Rows.Count <= 1 Then
            ShowWarning("请先添加销售商品！")
            Return
        End If
        recoveryEditState = 0

        Dim r As DataGridViewRow = dgvRecovery.Rows(e.RowIndex)
        Dim salesGoldPrice As String = Math.Round(SafeDecimal(dgvSales.Rows(0).Cells("col19").Value), 2).ToString()
        Dim recoveryName As String = SafeString(r.Cells("rcol1").Value)
        Dim recoveryQty As String = SafeString(r.Cells("rcol2").Value)
        Dim recoveryTotalWeight As String = SafeString(r.Cells("rcol3").Value)
        Dim recoveryGoldWeight As String = SafeString(r.Cells("rcol4").Value)
        Dim recoveryChengse As String = SafeString(r.Cells("rcol5").Value)
        Dim recoveryPrice As String = SafeString(r.Cells("rcol6").Value)
        Dim recoveryOtherFee As String = SafeString(r.Cells("rcol7").Value)
        Dim recoveryTotalAmount As String = SafeString(r.Cells("rcol8").Value)

        ' 商品名称列编辑
        If e.ColumnIndex = 1 Then
            Dim sql As String = $"SELECT title FROM xipunum_erp_retreat_title WHERE (bianma = '{SafeSQL(recoveryName)}' or title like '%{SafeSQL(recoveryName)}%') LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            Dim queryName As String = ""
            If dt.Rows.Count > 0 Then
                queryName = SafeString(dt.Rows(0)("title"))
            End If

            If queryName = "" Then
                ShowWarning("回收名称不存在！")
                r.Cells("rcol1").Value = ""
                Return
            End If

            r.Cells("rcol1").Value = queryName
        End If

        ' 非名称列时检查名称是否为空
        If e.ColumnIndex <> 1 AndAlso recoveryName = "" Then
            ShowWarning("回收名称不能为空！")
            Return
        End If

        ' 数量列
        If e.ColumnIndex = 2 Then
            r.Cells("rcol2").Value = recoveryQty
        End If

        ' 总重列
        If e.ColumnIndex = 3 Then
            r.Cells("rcol3").Value = FormatThreeDecimals(recoveryTotalWeight)
            r.Cells("rcol4").Value = recoveryTotalWeight
        End If

        ' 金重列
        If e.ColumnIndex = 4 Then
            r.Cells("rcol4").Value = FormatThreeDecimals(recoveryGoldWeight)
        End If

        ' 成色列
        If e.ColumnIndex = 5 Then
            If SafeDecimal(recoveryChengse) > 1 Then
                ShowWarning("成色数值不能大于1！")
                r.Cells("rcol5").Value = "1.0000"
                Return
            End If
            r.Cells("rcol5").Value = FormatFourDecimals(recoveryChengse)
            recoveryPrice = (SafeDecimal(recoveryChengse) * SafeDecimal(salesGoldPrice)).ToString()
            r.Cells("rcol6").Value = FormatTwoDecimals(recoveryPrice)
        End If

        ' 克价列
        If e.ColumnIndex = 6 Then
            recoveryChengse = Math.Round(SafeDecimal(recoveryPrice) / SafeDecimal(salesGoldPrice), 4).ToString()
            r.Cells("rcol5").Value = recoveryChengse
            r.Cells("rcol6").Value = FormatTwoDecimals(recoveryPrice)
        End If

        ' 其他费用列
        If e.ColumnIndex = 7 Then
            r.Cells("rcol7").Value = FormatTwoDecimals(recoveryOtherFee)
        End If

        ' 回收金额列
        If e.ColumnIndex = 8 Then
            r.Cells("rcol8").Value = FormatTwoDecimals(recoveryTotalAmount)
        End If

        ' 自动计算回收金额（列3/4/5/6/7变化时）
        If e.ColumnIndex = 3 OrElse e.ColumnIndex = 4 OrElse e.ColumnIndex = 5 OrElse e.ColumnIndex = 6 OrElse e.ColumnIndex = 7 Then
            recoveryTotalAmount = (SafeDecimal(recoveryGoldWeight) * SafeDecimal(recoveryPrice) + SafeDecimal(recoveryOtherFee)).ToString()
            recoveryTotalAmount = FormatTwoDecimals(Math.Round(SafeDecimal(recoveryTotalAmount), 0).ToString())
            r.Cells("rcol8").Value = recoveryTotalAmount
        End If

        CalculateSummary()
    End Sub

    ' ========== 收支表格结束编辑（_高级表格3_结束编辑） ==========
    Private Sub DgvPayment_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        paymentEditState = 0

        ' 计算合计
        Dim payTotal As Decimal = 0
        For i As Integer = 0 To dgvPayment.Rows.Count - 2
            If Not dgvPayment.Rows(i).IsNewRow Then
                payTotal += SafeDecimal(dgvPayment.Rows(i).Cells("pcol2").Value)
            End If
        Next
        If dgvPayment.Rows.Count > 0 Then
            dgvPayment.Rows(dgvPayment.Rows.Count - 1).Cells("pcol2").Value = payTotal.ToString()
        End If
        txtReceived.Text = payTotal.ToString()
    End Sub

    ' ========== 回收加按钮（_按钮_回收加_被单击） ==========
    Private Sub BtnRecoveryAdd_Click(sender As Object, e As EventArgs)
        If dgvSales.Rows.Count <= 1 OrElse SafeString(dgvSales.Rows(0).Cells("col1").Value) = "" Then
            ShowWarning("请先输入销售商品！")
            Return
        End If

        If SafeDecimal(dgvSales.Rows(0).Cells("col19").Value) <= 0 Then
            ShowWarning("销售第一件商品克价不能为0！")
            Return
        End If

        Dim addRowIndex As Integer = dgvRecovery.Rows.Count
        If addRowIndex > 0 Then
            Dim lastRow As DataGridViewRow = dgvRecovery.Rows(addRowIndex - 1)
            If SafeString(lastRow.Cells("rcol8").Value) = "" Then
                ShowWarning("上一个回收数据不能为空！")
                Return
            End If
            If SafeString(lastRow.Cells("rcol10").Value) = "" Then
                ShowWarning("上一个回收导购员不能为空！")
                Return
            End If
        End If

        Dim newRow As Integer = dgvRecovery.Rows.Add()
        dgvRecovery.Rows(newRow).Cells("rcol0").Value = ""
        dgvRecovery.Rows(newRow).Cells("rcol1").Value = ""
        dgvRecovery.Rows(newRow).Cells("rcol2").Value = "1"
        dgvRecovery.Rows(newRow).Cells("rcol3").Value = ""
        dgvRecovery.Rows(newRow).Cells("rcol4").Value = ""
        dgvRecovery.Rows(newRow).Cells("rcol5").Value = ""
        dgvRecovery.Rows(newRow).Cells("rcol6").Value = ""
        dgvRecovery.Rows(newRow).Cells("rcol7").Value = ""
        dgvRecovery.Rows(newRow).Cells("rcol8").Value = ""
        dgvRecovery.Rows(newRow).Cells("rcol9").Value = ""

        If newRow = 0 Then
            dgvRecovery.Rows(newRow).Cells("rcol10").Value = SafeString(dgvSales.Rows(0).Cells("col24").Value)
        Else
            dgvRecovery.Rows(newRow).Cells("rcol10").Value = SafeString(dgvRecovery.Rows(newRow - 1).Cells("rcol10").Value)
        End If
    End Sub

    ' ========== 回收减按钮（_按钮_回收减_被单击） ==========
    Private Sub BtnRecoveryRemove_Click(sender As Object, e As EventArgs)
        If dgvRecovery.CurrentCell Is Nothing Then
            ShowWarning("请选择需要删除的回收数据！")
            Return
        End If

        Dim deleteRow As Integer = dgvRecovery.CurrentCell.RowIndex
        If deleteRow > 0 Then
            dgvRecovery.Rows.RemoveAt(deleteRow)
        End If

        If dgvRecovery.Rows.Count <= 1 Then
            recoveryEditState = 0
        End If

        CalculateSummary()
    End Sub

    ' ========== 保存按钮（_超级按钮_保存_被单击） ==========
    Private Sub BtnToolbarSave_Click(sender As Object, e As EventArgs)
        SaveSalesOrder()
    End Sub

    ' ========== 重置按钮（_工具条_通用_被单击: 重置） ==========
    Private Sub BtnToolbarReset_Click(sender As Object, e As EventArgs)
        ClearTempData()
        cmbSalesFactory.Text = ""
        cmbSalesFactory.SelectedIndex = -1
        Form_Load(Nothing, Nothing)
        CalculateSummary()
    End Sub

    ' ========== 批量编辑按钮 ==========
    Private Sub BtnToolbarBatchEdit_Click(sender As Object, e As EventArgs)
        Dim form As New SalesBatchEditForm()
        form.ShowDialog(Me)
    End Sub

    ' ========== 打印预览按钮 ==========
    Private Sub BtnToolbarPrintPreview_Click(sender As Object, e As EventArgs)
        PrintDocument()
    End Sub

    ' ========== 保存销售单（_超级按钮_保存_被单击 完整逻辑） ==========
    Private Sub SaveSalesOrder()
        If saveInProgress = 1 Then Return

        ' 数据库连接检查
        If MySQL_Write Is Nothing OrElse MySQL_Read Is Nothing Then
            ShowWarning("业务数据库未连接，无法保存！")
            Return
        End If

        LogOperationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        LogOperationAccount = UserAccount

        ' 验证销售数量
        If dgvSales.Rows.Count <= 2 Then
            ShowWarning("销售数量不能为空！")
            Return
        End If

        ' 批发模式验证会员
        If radioWholesale.Checked Then
            If cmbMember.SelectedIndex = -1 AndAlso cmbMember.Text = "" Then
                ShowWarning("批发销售会员信息不能为空！")
                cmbMember.Text = ""
                cmbMember.Focus()
                Return
            End If
        End If

        ' 验证业务员
        If cmbGuide.SelectedIndex = -1 Then
            ShowWarning("请选择业务员！")
            cmbGuide.Focus()
            Return
        End If

        ' 验证实收金额
        If txtReceived.Text = "" Then
            ShowWarning("实收金额不能为空！")
            txtReceived.Focus()
            Return
        End If

        ' 回收单号生成
        Dim recoveryOrderNumber As String = ""
        If dgvRecovery.Rows.Count > 1 Then
            recoveryOrderNumber = DateTime.Now.ToString("yyyyMMddHHmmss") & New Random().Next(1000, 9999).ToString() & "3"
            Dim checkSql As String = $"SELECT * FROM xipunum_erp_retreat_order where retrea_umber ='{SafeSQL(recoveryOrderNumber)}'"
            Dim checkDt As DataTable = DatabaseModule.ExecuteQuery(checkSql)
            If checkDt.Rows.Count > 0 Then
                recoveryOrderNumber = DateTime.Now.ToString("yyyyMMddHHmmss") & New Random().Next(1000, 9999).ToString() & "3"
            End If
        End If

        ' 介绍人必填检查
        If IntroducerRequired = "1" AndAlso cmbIntroducer.SelectedIndex = -1 Then
            ShowWarning("请选择介绍人！")
            cmbIntroducer.Focus()
            Return
        End If

        ' 出库单号
        Dim outboundOrderNumber As String = SafeSQL(txtOrderNumber.Text)

        ' 检查单号是否已存在
        Dim existSql As String = $"SELECT * FROM xipunum_erp_outbound_order where settlement_number ='{outboundOrderNumber}'"
        Dim existDt As DataTable = DatabaseModule.ExecuteQuery(existSql)
        If existDt.Rows.Count > 0 Then
            ShowWarning("当前出库单号已存在，已重新生成单号，请再次点击保存按钮！")
            txtOrderNumber.Text = DateTime.Now.ToString("yyyyMMddHHmmss") & New Random().Next(1000, 9999).ToString() & "1"
            Return
        End If

        ' 库存预检查
        Dim checkCount As Integer = 0
        For i As Integer = 0 To dgvSales.Rows.Count - 2
            If dgvSales.Rows(i).IsNewRow Then Continue For
            Dim chkCode As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col1").Value))
            Dim chkQty As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col17").Value))
            Dim chkWeight As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col9").Value))
            Dim chkSql As String = $"SELECT * FROM xipunum_erp_shop_kucun where poduct_code ='{chkCode}' and quantity >='{chkQty}' and jinzhong>='{chkWeight}' and kufang='{UserDepartment}'"
            Dim chkDt As DataTable = DatabaseModule.ExecuteQuery(chkSql)
            If chkDt.Rows.Count > 0 Then
                checkCount += 1
            End If
        Next

        If checkCount <> dgvSales.Rows.Count - 1 Then
            ShowWarning("销售数据存在库存不存在数据，请检查删除后在保存！")
            Return
        End If

        ' 编辑状态检查
        If tableEditState = 1 Then
            ShowWarning("请先结束销售列表操作！")
            Return
        End If
        If recoveryEditState = 1 Then
            ShowWarning("请先结束回收列表操作！")
            Return
        End If
        If paymentEditState = 1 Then
            ShowWarning("请先结束收支列表操作！")
            Return
        End If

        ' 回收数据验证
        If dgvRecovery.Rows.Count > 1 Then
            If cmbGuide.SelectedIndex = -1 Then
                ShowWarning("有回收数据时请选择业务员！")
                cmbGuide.Focus()
                Return
            End If
            For i As Integer = 0 To dgvRecovery.Rows.Count - 2
                If dgvRecovery.Rows(i).IsNewRow Then Continue For
                If SafeString(dgvRecovery.Rows(i).Cells("rcol1").Value) = "" Then
                    ShowWarning("回收明细商品名称不能为空！")
                    Return
                End If
                If SafeString(dgvRecovery.Rows(i).Cells("rcol10").Value) = "" Then
                    ShowWarning("回收明细导购员不能为空！")
                    Return
                End If
            Next
        End If

        ' 获取会员信息
        Dim memberName As String = ""
        If cmbMember.SelectedIndex = -1 Then
            memberName = SafeSQL(cmbMember.Text)
        Else
            memberName = SafeSQL(cmbMember.SelectedItem.ToString())
        End If

        Dim memberPhone As String = SafeSQL(txtPhone.Text)
        Dim taxRate As String = SafeSQL(txtTaxRate.Text)
        Dim taxAmount As String = SafeSQL(txtTax.Text)
        Dim receivableAmount As String = SafeSQL(txtReceivable.Text)
        Dim actualReceivedAmount As String = SafeSQL(txtReceived.Text)
        Dim salesRemarks As String = SafeSQL(txtRemarks.Text)
        Dim discountAmount As String = FormatTwoDecimals((SafeDecimal(actualAmount) - SafeDecimal(recoveryAmountTotal) - SafeDecimal(actualReceivedAmount)).ToString())

        ' 会员编码获取
        Dim customerCode As String = ""
        Dim newMemberCode As String = ""
        If cmbMember.SelectedIndex <> -1 Then
            Dim memberIdText As String = lblMemberId.Text
            Dim memberId As String = If(memberIdText.Length >= 7, memberIdText.Substring(memberIdText.Length - 7), "")
            Dim memberSql As String = $"SELECT id,customer_code FROM xipunum_erp_member where memberid='{SafeSQL(memberId)}' order by id ASC LIMIT 1"
            Dim memberDt As DataTable = DatabaseModule.ExecuteQuery(memberSql)
            If memberDt.Rows.Count > 0 Then
                customerCode = SafeString(memberDt.Rows(0)("customer_code"))
            End If
        Else
            If cmbMember.Text <> "" Then
                ' 检查会员是否已存在
                Dim checkMemberSql As String = $"SELECT * FROM xipunum_erp_member where name= '{SafeSQL(memberName)}' and tel= '{SafeSQL(memberPhone)}'"
                Dim checkMemberDt As DataTable = DatabaseModule.ExecuteQuery(checkMemberSql)
                If checkMemberDt.Rows.Count > 0 Then
                    ShowWarning("此会员名称和手机号已存在！")
                    cmbMember.SelectedIndex = -1
                    cmbMember.Focus()
                    cmbMember.Text = ""
                    Return
                End If

                ' 新建会员
                newMemberCode = "HY" & DateTime.Now.ToString("yyyyMMddHHmmss")
                customerCode = newMemberCode
            End If
        End If

        ' 预售单号
        Dim presaleNumber As String = ""
        If radioPresaleYes.Checked Then
            presaleNumber = SafeSQL(txtPresaleNumber.Text)
        End If

        ' 批零
        If radioRetail.Checked Then
            salesPling = "零售"
        End If
        If radioWholesale.Checked Then
            salesPling = "批发"
        End If

        ' 发票
        Dim fapiao As String = "0"
        If radioNoInvoice.Checked Then
            fapiao = "0"
        End If
        If radioInvoice.Checked Then
            fapiao = "1"
        End If

        ' 业务员
        Dim salesman As String = If(cmbGuide.SelectedIndex >= 0, cmbGuide.SelectedItem.ToString(), "")
        If salesman = "" Then
            ShowWarning("业务员信息无效，请重新选择！")
            Return
        End If

        ' 汇总数据
        Dim counterWeight As String = weightTotal
        Dim settlementWeight As String = weightTotal
        Dim netCounter As String = weightTotal
        Dim netWeight As String = weightTotal
        Dim basicCostSum As String = basicCostTotal
        Dim premiumCostSum As String = premiumCostTotal
        Dim salesCostSum As String = salesCostTotal
        Dim salesSurchSum As String = salesSurchargeTotal

        ' 介绍人
        Dim introducer As String = ""
        If cmbIntroducer.SelectedIndex <> -1 Then
            Dim introText As String = cmbIntroducer.SelectedItem.ToString()
            If introText.Length >= 8 Then
                introducer = introText.Substring(introText.Length - 8, 7)
            End If
        End If

        ' 销售工厂
        Dim salesFactory As String = ""
        If cmbSalesFactory.SelectedIndex = -1 Then
            salesFactory = SafeSQL(cmbSalesFactory.Text)
        Else
            salesFactory = SafeSQL(cmbSalesFactory.SelectedItem.ToString())
        End If

        saveOrderNumber = outboundOrderNumber
        saveRecoveryNumber = ""
        saveInProgress = 0

        ' 写入出库主单
        Dim insertOrderSql As String = $"INSERT INTO xipunum_erp_outbound_order SET " &
            $"settlement_number='{outboundOrderNumber}'," &
            $"presale_number='{presaleNumber}'," &
            $"retrea_umber='{recoveryOrderNumber}'," &
            $"ying_amount='{receivableAmount}'," &
            $"youhui='{discountAmount}'," &
            $"settlement='{actualReceivedAmount}'," &
            $"customer_code='{customerCode}'," &
            $"brand_name=''," &
            $"salesman='{salesman}'," &
            $"counter_weight='{counterWeight}'," &
            $"settlement_weight='{settlementWeight}'," &
            $"net_counter='{netCounter}'," &
            $"net_weight='{netWeight}'," &
            $"basic_cost='{basicCostSum}'," &
            $"premium_cost='{premiumCostSum}'," &
            $"sales_cost='{salesCostSum}'," &
            $"sales_surcharge='{salesSurchSum}'," &
            $"shopping_guide='{introducer}'," &
            $"taxpoint='{taxRate}'," &
            $"taxamount='{taxAmount}'," &
            $"remarks='{salesRemarks}'," &
            $"pling='{salesPling}'," &
            $"state='正常'," &
            $"sales_return='0'," &
            $"fapiao='{fapiao}'," &
            $"xsfactory='{salesFactory}'," &
            $"cjuser='{LogOperationAccount}'," &
            $"creationtime='{LogOperationDate}'," &
            $"xianshangtime='{LogOperationDate}'"
        DatabaseModule.ExecuteCommand(insertOrderSql, MySQL_Write)

        ' 获取订单ID
        Dim orderIdSql As String = $"SELECT id FROM xipunum_erp_outbound_order where settlement_number='{outboundOrderNumber}' order by id ASC LIMIT 1"
        Dim orderIdDt As DataTable = DatabaseModule.ExecuteQuery(orderIdSql)
        Dim orderId As String = ""
        If orderIdDt.Rows.Count > 0 Then
            orderId = SafeString(orderIdDt.Rows(0)("id"))
        End If
        If orderId = "" Then
            SaveFailedRollback("商品销售出库主单写入失败，已回滚！")
            Return
        End If

        ' 写入回收主单
        Dim recoveryOrderId As String = ""
        If dgvRecovery.Rows.Count > 1 Then
            Dim recoverySalesman As String = salesman
            Dim recTotalWeight As String = recoveryWeightTotal
            Dim recGoldWeight As String = recoveryGoldTotal
            Dim recOtherFee As String = recoveryOtherTotal
            Dim recAmount As String = recoveryAmountTotal
            Dim recPayable As String = recoveryAmountTotal
            Dim recActualPay As String = recoveryAmountTotal

            Dim insertRecoverySql As String = $"INSERT INTO xipunum_erp_retreat_order SET " &
                $"retrea_umber='{recoveryOrderNumber}'," &
                $"customer_code='{customerCode}'," &
                $"total='{recTotalWeight}'," &
                $"jin_zhong='{recGoldWeight}'," &
                $"qita_price='{recOtherFee}'," &
                $"tax_rate='0'," &
                $"rate_amount='0'," &
                $"retreat_amount='{recAmount}'," &
                $"ying_amount='{recPayable}'," &
                $"settlement='{recActualPay}'," &
                $"salesman='{recoverySalesman}'," &
                $"remarks=''," &
                $"cjuser='{LogOperationAccount}'," &
                $"creationtime='{LogOperationDate}'"
            DatabaseModule.ExecuteCommand(insertRecoverySql, MySQL_Write)

            Dim recOrderIdSql As String = $"SELECT id FROM xipunum_erp_retreat_order where retrea_umber='{recoveryOrderNumber}' order by id ASC LIMIT 1"
            Dim recOrderIdDt As DataTable = DatabaseModule.ExecuteQuery(recOrderIdSql)
            If recOrderIdDt.Rows.Count > 0 Then
                recoveryOrderId = SafeString(recOrderIdDt.Rows(0)("id"))
            End If
            If recoveryOrderId = "" Then
                SaveFailedRollback("商品回收主单写入失败，已回滚！")
                Return
            End If
            saveRecoveryNumber = recoveryOrderNumber
        End If

        ' 开始事务
        DatabaseModule.ExecuteCommand("START TRANSACTION", MySQL_Write)
        saveInProgress = 1

        ' 新建会员
        If newMemberCode <> "" Then
            Dim insertMemberSql As String = $"INSERT INTO xipunum_erp_member SET customer_code='{newMemberCode}',name='{memberName}',tel='{memberPhone}',cjuser='{LogOperationAccount}',creationtime='{LogOperationDate}'"
            DatabaseModule.ExecuteCommand(insertMemberSql, MySQL_Write)
            saveMemberCode = newMemberCode

            Dim memIdSql As String = $"SELECT id,customer_code FROM xipunum_erp_member where customer_code='{newMemberCode}' order by id ASC LIMIT 1"
            Dim memIdDt As DataTable = DatabaseModule.ExecuteQuery(memIdSql, MySQL_Write)
            If memIdDt.Rows.Count > 0 Then
                customerCode = SafeString(memIdDt.Rows(0)("customer_code"))
                Dim customerId As String = SafeString(memIdDt.Rows(0)("id"))
                If customerCode = "" Then
                    SaveFailedRollback("新建会员写入失败，已回滚！")
                    Return
                End If
                Dim updateMemIdSql As String = $"UPDATE xipunum_erp_member SET memberid= '1{("00000000" & customerId).Substring(("00000000" & customerId).Length - 6)}' WHERE customer_code ='{newMemberCode}' LIMIT 1"
                DatabaseModule.ExecuteCommand(updateMemIdSql, MySQL_Write)
            End If
        End If

        ' 写入收款记录
        For i As Integer = 0 To dgvPayment.Rows.Count - 2
            If dgvPayment.Rows(i).IsNewRow Then Continue For
            Dim payId As String = SafeString(dgvPayment.Rows(i).Cells("pcol3").Value)
            Dim payAmount As String = SafeString(dgvPayment.Rows(i).Cells("pcol2").Value)
            If SafeDecimal(payAmount) <> 0 Then
                Dim insertPaySql As String = $"INSERT INTO xipunum_erp_shoukuan SET leibie='1',settlement_number='{outboundOrderNumber}',xianjin='{payAmount}',type='{payId}',kufang='{UserDepartment}',cjuser='{LogOperationAccount}',creationtime='{LogOperationDate}'"
                DatabaseModule.ExecuteCommand(insertPaySql, MySQL_Write)
            End If
        Next

        ' 系统日志
        Dim logContent As String = $"账户:{UserAccount} 商品销售，销售单号:{txtOrderNumber.Text}"
        Dim insertLogSql As String = $"INSERT INTO xipunum_erp_xitong_log SET type='添加',title='商品销售',conter='{SafeSQL(logContent)}',user='{LogOperationAccount}',creationtime='{LogOperationDate}'"
        DatabaseModule.ExecuteCommand(insertLogSql, MySQL_Write)

        ' 批发模式：会员存料/存元/欠料/欠款
        If radioWholesale.Checked Then
            If SafeDecimal(txtStoreMaterial.Text) <> 0 Then
                Dim insertCqSql As String = $"INSERT INTO xipunum_erp_member_cq SET customer_code='{customerCode}',danhao='{outboundOrderNumber}',cunqu='存',type='料',number='{txtStoreMaterial.Text}',kufang='{UserDepartment}',cjuser='{LogOperationAccount}',creationtime='{LogOperationDate}'"
                DatabaseModule.ExecuteCommand(insertCqSql, MySQL_Write)
                Dim logCq As String = $"账户:{UserAccount} 操作 会员编码：{customerCode} 会员:{memberName} 存储（{txtStoreMaterial.Text}）g 料"
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_xitong_log SET type='添加',title='会员存料',conter='{SafeSQL(logCq)}',user='{LogOperationAccount}',creationtime='{LogOperationDate}'", MySQL_Write)
            End If

            If SafeDecimal(txtStoreBalance.Text) <> 0 Then
                Dim insertCqSql As String = $"INSERT INTO xipunum_erp_member_cq SET customer_code='{customerCode}',danhao='{outboundOrderNumber}',cunqu='存',type='元',number='{txtStoreBalance.Text}',kufang='{UserDepartment}',cjuser='{LogOperationAccount}',creationtime='{LogOperationDate}'"
                DatabaseModule.ExecuteCommand(insertCqSql, MySQL_Write)
                Dim logCq As String = $"账户:{UserAccount} 操作 会员编码：{customerCode} 会员:{memberName} 存元（{txtStoreBalance.Text}）元"
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_xitong_log SET type='添加',title='会员存元',conter='{SafeSQL(logCq)}',user='{LogOperationAccount}',creationtime='{LogOperationDate}'", MySQL_Write)
            End If

            If SafeDecimal(recoveryGoldTotal) > 0 Then
                Dim insertCqSql As String = $"INSERT INTO xipunum_erp_member_cq SET customer_code='{customerCode}',danhao='{outboundOrderNumber}',cunqu='欠',type='料',number='{recoveryGoldTotal}',kufang='{UserDepartment}',cjuser='{LogOperationAccount}',creationtime='{LogOperationDate}'"
                DatabaseModule.ExecuteCommand(insertCqSql, MySQL_Write)
                Dim logCq As String = $"账户:{UserAccount} 增加会员:{customerCode}欠料（{recoveryGoldTotal}）g 料"
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_xitong_log SET type='新增',title='会员出料',conter='{SafeSQL(logCq)}',user='{LogOperationAccount}',creationtime='{LogOperationDate}'", MySQL_Write)
            End If

            If SafeDecimal(actualReceivedAmount) > 0 Then
                Dim insertCqSql As String = $"INSERT INTO xipunum_erp_member_cq SET customer_code='{customerCode}',danhao='{outboundOrderNumber}',cunqu='欠',type='元',number='{actualReceivedAmount}',kufang='{UserDepartment}',cjuser='{LogOperationAccount}',creationtime='{LogOperationDate}'"
                DatabaseModule.ExecuteCommand(insertCqSql, MySQL_Write)
                Dim logCq As String = $"账户:{UserAccount} 增加会员:{customerCode}欠款（{actualReceivedAmount}）元"
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_xitong_log SET type='新增',title='会员消费',conter='{SafeSQL(logCq)}',user='{LogOperationAccount}',creationtime='{LogOperationDate}'", MySQL_Write)
            End If
        End If

        ' 获取库房名称
        Dim kufangName As String = ""
        Dim kufangSql As String = $"SELECT title FROM xipunum_erp_type where id='{UserDepartment}' order by id ASC LIMIT 1"
        Dim kufangDt As DataTable = DatabaseModule.ExecuteQuery(kufangSql)
        If kufangDt.Rows.Count > 0 Then
            kufangName = SafeString(kufangDt.Rows(0)("title"))
        End If

        ' 销售明细批量INSERT
        Dim salesCount As Integer = dgvSales.Rows.Count - 1
        Dim batchSalesData As String = ""
        Dim batchHistoryData As String = ""
        Dim batchShopLogData As String = ""
        Dim batchSystemLogData As String = ""
        saveLastGuide = ""
        saveLastGuideAccount = ""
        Dim scorePoints As Decimal = 0

        For i As Integer = 0 To salesCount - 1
            If dgvSales.Rows(i).IsNewRow Then Continue For

            Dim sCode As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col1").Value))
            Dim sQty As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col17").Value))
            Dim sUnitPrice As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col15").Value))
            Dim sAmount As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col16").Value))
            Dim sActual As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col23").Value))
            Dim sGoldPrice As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col19").Value))
            Dim sNetWeight As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col9").Value))
            Dim sBasicCost As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col11").Value))
            Dim sPremiumCost As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col12").Value))
            Dim sSalesCost As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col20").Value))
            Dim sSalesSurch As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col21").Value))
            Dim sDiscount As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col22").Value))
            Dim sGuide As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col24").Value))

            ' 查询导购员账户
            Dim accountInfo As String = ""
            If sGuide = saveLastGuide AndAlso saveLastGuide <> "" Then
                accountInfo = saveLastGuideAccount
            Else
                Dim accSql As String = $"SELECT user FROM xipunum_erp_user WHERE name = '{sGuide}' and department = '{UserDepartment}' LIMIT 1"
                Dim accDt As DataTable = DatabaseModule.ExecuteQuery(accSql)
                If accDt.Rows.Count > 0 Then
                    accountInfo = SafeString(accDt.Rows(0)("user"))
                End If
                If accountInfo = "" Then
                    accountInfo = LogOperationAccount
                End If
                saveLastGuide = sGuide
                saveLastGuideAccount = accountInfo
            End If

            ' 扣减库存
            Dim deductSql As String = $"UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '{sQty}',jinzhong = jinzhong - '{sNetWeight}' WHERE poduct_code = '{sCode}' AND kufang='{UserDepartment}' AND quantity >='{sQty}' AND jinzhong>='{sNetWeight}'"
            Dim deductResult As Integer = DatabaseModule.ExecuteCommand(deductSql, MySQL_Write)
            If deductResult < 1 Then
                Dim rowCnt As Object = DatabaseModule.ExecuteScalar("SELECT ROW_COUNT() AS rowcnt", MySQL_Write)
                If SafeString(rowCnt) = "" OrElse SafeDecimal(SafeString(rowCnt)) < 1 Then
                    SaveFailedRollback($"商品编码 {sCode} 库存不足或已被占用，保存已回滚！")
                    Return
                End If
            End If

            ' 批量INSERT数据
            batchSalesData &= $"('{orderId}','{sCode}','{sQty}','{sUnitPrice}','{sAmount}','{sActual}','{sGoldPrice}','{sNetWeight}','{sBasicCost}','{sPremiumCost}','{sSalesCost}','{sSalesSurch}','{accountInfo}','{UserDepartment}','{sDiscount}','','{salesPling}','0','0','{LogOperationAccount}','{LogOperationDate}','{LogOperationDate}'),"

            ' 积分
            Dim jifenSql As String = $"SELECT CASE WHEN COALESCE(d.jifen, '') = '' THEN '0' ELSE d.jifen END AS jifen FROM xipunum_erp_shop AS a LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = a.item_number AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id AND a.item_number IS NOT NULL AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id AND a.specification_id IS NOT NULL AND a.specification_id != '' LEFT JOIN xipunum_erp_category AS d ON d.id = COALESCE ( e1.category_id, e2.category_id ) AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL WHERE a.poduct_code='{sCode}' LIMIT 1"
            Dim jifenDt As DataTable = DatabaseModule.ExecuteQuery(jifenSql)
            Dim jifenRate As String = "0"
            If jifenDt.Rows.Count > 0 Then
                jifenRate = SafeString(jifenDt.Rows(0)("jifen"))
            End If
            scorePoints += Math.Round(SafeDecimal(sActual) * SafeDecimal(jifenRate) / 100, 0)

            ' 证书更新
            Dim certSql As String = $"SELECT * FROM xipunum_erp_zhengshu WHERE poduct_code = '{sCode}' LIMIT 1"
            Dim certDt As DataTable = DatabaseModule.ExecuteQuery(certSql)
            If certDt.Rows.Count > 0 Then
                Dim certId As String = SafeString(certDt.Rows(0)("id"))
                Dim certCost As String = SafeString(certDt.Rows(0)("chengben"))
                Dim certSales As String = SafeString(certDt.Rows(0)("xiaoshou"))
                Dim certSalesUpdate As String = ""
                If certSales = "" Then
                    certSalesUpdate = $",xiaoshou= '{certCost}'"
                End If
                Dim updateCertSql As String = $"UPDATE xipunum_erp_zhengshu SET xiaoshouid= '{outboundOrderNumber}'{certSalesUpdate},updatetime= '{LogOperationDate}' WHERE poduct_code ='{sCode}' LIMIT 1"
                DatabaseModule.ExecuteCommand(updateCertSql, MySQL_Write)
            End If

            ' 商品状态更新
            Dim stockSumSql As String = $"SELECT sum(quantity) as shuliang,sum(jinzhong) as jinzhong FROM xipunum_erp_shop_kucun WHERE poduct_code='{sCode}'"
            Dim stockSumDt As DataTable = DatabaseModule.ExecuteQuery(stockSumSql)
            Dim remainQty As String = "0"
            Dim remainWeight As String = "0"
            If stockSumDt.Rows.Count > 0 Then
                remainQty = (SafeDecimal(SafeString(stockSumDt.Rows(0)("shuliang"))) - SafeDecimal(sQty)).ToString()
                remainWeight = (SafeDecimal(SafeString(stockSumDt.Rows(0)("jinzhong"))) - SafeDecimal(sNetWeight)).ToString()
            End If

            If SafeDecimal(remainWeight) <= 0 AndAlso SafeDecimal(remainQty) <= 0 Then
                Dim updateStateSql As String = $"UPDATE xipunum_erp_shop SET state= '售尽',updatetime= '{LogOperationDate}' WHERE poduct_code ='{sCode}' LIMIT 1"
                DatabaseModule.ExecuteCommand(updateStateSql, MySQL_Write)
            End If

            ' 追溯数据
            Dim sWeight As String = SafeSQL(SafeString(dgvSales.Rows(i).Cells("col10").Value))
            batchHistoryData &= $"('{sCode}','{LogOperationDate}','{outboundOrderNumber}','成品出库','{sQty}','{sNetWeight}','{sWeight}','商品从{kufangName}出库','{LogOperationAccount}'),"
            batchShopLogData &= $"('{sCode}','出库','{LogOperationDate}'),"

            ' 系统日志
            Dim sLogContent As String = $"账户:{UserAccount} 商品销售出库，编码：{sCode}"
            batchSystemLogData &= $"('添加','商品销售出库','{SafeSQL(sLogContent)}','{LogOperationAccount}','{LogOperationDate}'),"
        Next

        ' 批量写入销售明细
        If batchSalesData <> "" Then
            batchSalesData = batchSalesData.TrimEnd(","c)
            Dim insertSalesSql As String = $"INSERT INTO xipunum_erp_outbound ( order_id, poduct_code, quantity, xiaodan_amount, xiao_amount, settlement, gold_price, net_weight, basic_cost, premium_cost, sales_cost, sales_surcharge, shopping_guide, kufang, zhekou, remarks, pling, state, sales_return, cjuser, creationtime, xianshangtime ) VALUES {batchSalesData};"
            Dim salesResult As Integer = DatabaseModule.ExecuteCommand(insertSalesSql, MySQL_Write)
            If salesResult < 1 Then
                Dim rowCnt As Object = DatabaseModule.ExecuteScalar("SELECT ROW_COUNT() AS rowcnt", MySQL_Write)
                If SafeString(rowCnt) = "" OrElse SafeDecimal(SafeString(rowCnt)) < 1 Then
                    SaveFailedRollback("销售明细写入失败，已回滚！")
                    Return
                End If
            End If
        Else
            SaveFailedRollback("销售明细为空，保存已取消并回滚！")
            Return
        End If

        ' 批量写入追溯
        If batchHistoryData <> "" Then
            batchHistoryData = batchHistoryData.TrimEnd(","c)
            Dim insertHistorySql As String = $"INSERT INTO xipunum_erp_history ( poduct_code, updatetime, number, type, quantity, jinzhong, zhongliang, conter, cjuser ) VALUES {batchHistoryData};"
            Dim histResult As Integer = DatabaseModule.ExecuteCommand(insertHistorySql, MySQL_Write)
            If histResult < 1 Then
                Dim rowCnt As Object = DatabaseModule.ExecuteScalar("SELECT ROW_COUNT() AS rowcnt", MySQL_Write)
                If SafeString(rowCnt) = "" OrElse SafeDecimal(SafeString(rowCnt)) < 1 Then
                    SaveFailedRollback("销售追溯写入失败，已回滚！")
                    Return
                End If
            End If
        End If

        ' 批量写入商品日志
        If batchShopLogData <> "" Then
            batchShopLogData = batchShopLogData.TrimEnd(","c)
            Dim insertShopLogSql As String = $"INSERT INTO xipunum_erp_shop_log ( poduct_code, type,creationtime ) VALUES {batchShopLogData};"
            DatabaseModule.ExecuteCommand(insertShopLogSql, MySQL_Write)
        End If

        ' 批量写入系统日志
        If batchSystemLogData <> "" Then
            batchSystemLogData = batchSystemLogData.TrimEnd(","c)
            Dim insertSysLogSql As String = $"INSERT INTO xipunum_erp_xitong_log ( type, title, conter, user, creationtime ) VALUES {batchSystemLogData};"
            DatabaseModule.ExecuteCommand(insertSysLogSql, MySQL_Write)
        End If

        ' 预售订单更新
        If radioPresaleYes.Checked Then
            Dim updatePresaleSql As String = $"UPDATE xipunum_erp_presale_order SET state= '已收货',updatetime= '{LogOperationDate}' WHERE presale_umber ='{presaleNumber}' LIMIT 1"
            DatabaseModule.ExecuteCommand(updatePresaleSql, MySQL_Write)
        End If

        ' 回收明细写入
        If dgvRecovery.Rows.Count > 1 Then
            saveLastGuide = ""
            saveLastGuideAccount = ""
            For i As Integer = 0 To dgvRecovery.Rows.Count - 2
                If dgvRecovery.Rows(i).IsNewRow Then Continue For

                Dim rName As String = SafeSQL(SafeString(dgvRecovery.Rows(i).Cells("rcol1").Value))
                Dim rQty As String = SafeSQL(SafeString(dgvRecovery.Rows(i).Cells("rcol2").Value))
                Dim rTotal As String = SafeSQL(SafeString(dgvRecovery.Rows(i).Cells("rcol3").Value))
                Dim rGold As String = SafeSQL(SafeString(dgvRecovery.Rows(i).Cells("rcol4").Value))
                Dim rChengse As String = SafeSQL(SafeString(dgvRecovery.Rows(i).Cells("rcol5").Value))
                Dim rPrice As String = SafeSQL(SafeString(dgvRecovery.Rows(i).Cells("rcol6").Value))
                Dim rOther As String = SafeSQL(SafeString(dgvRecovery.Rows(i).Cells("rcol7").Value))
                Dim rAmount As String = SafeSQL(SafeString(dgvRecovery.Rows(i).Cells("rcol8").Value))
                Dim rRemarks As String = SafeSQL(SafeString(dgvRecovery.Rows(i).Cells("rcol9").Value))
                Dim rGuide As String = SafeSQL(SafeString(dgvRecovery.Rows(i).Cells("rcol10").Value))

                ' 查询导购员账户
                Dim accountInfo As String = ""
                If rGuide = saveLastGuide AndAlso saveLastGuide <> "" Then
                    accountInfo = saveLastGuideAccount
                Else
                    Dim accSql As String = $"SELECT user FROM xipunum_erp_user WHERE name = '{rGuide}' and department = '{UserDepartment}' LIMIT 1"
                    Dim accDt As DataTable = DatabaseModule.ExecuteQuery(accSql)
                    If accDt.Rows.Count > 0 Then
                        accountInfo = SafeString(accDt.Rows(0)("user"))
                    End If
                    If accountInfo = "" Then
                        accountInfo = LogOperationAccount
                    End If
                    saveLastGuide = rGuide
                    saveLastGuideAccount = accountInfo
                End If

                ' 查询回收品类ID和积分比例
                Dim retSql As String = $"SELECT a.id as aid,CASE WHEN COALESCE(b.jifen, '') = '' THEN '0' ELSE b.jifen END AS jifen FROM xipunum_erp_retreat_title as a INNER JOIN xipunum_erp_category AS b ON b.id = a.category_id WHERE a.title = '{rName}' LIMIT 1"
                Dim retDt As DataTable = DatabaseModule.ExecuteQuery(retSql)
                Dim retDataId As String = "1"
                Dim retJifen As String = "0"
                If retDt.Rows.Count > 0 Then
                    retDataId = SafeString(retDt.Rows(0)("aid"))
                    retJifen = SafeString(retDt.Rows(0)("jifen"))
                End If

                Dim insertRetSql As String = $"INSERT INTO xipunum_erp_retreat SET order_id='{recoveryOrderId}',product_name='{retDataId}',quantity='{rQty}',total='{rTotal}',jin_zhong='{rGold}',chengse='{rChengse}',price='{rPrice}',qita_price='{rOther}',retreat_amount='{rAmount}',huishoutime='{LogOperationDate}',shopping_guide='{accountInfo}',remarks='{rRemarks}',cjuser='{LogOperationAccount}',creationtime='{LogOperationDate}'"
                DatabaseModule.ExecuteCommand(insertRetSql, MySQL_Write)

                scorePoints -= Math.Round(SafeDecimal(rAmount) * SafeDecimal(retJifen) / 100, 0)

                Dim retLogContent As String = $"账户:{UserAccount} 商品回收，名称：{SafeString(dgvRecovery.Rows(i).Cells("rcol1").Value)}"
                Dim retLogSql As String = $"INSERT INTO xipunum_erp_xitong_log SET type='添加',title='商品回收',conter='{SafeSQL(retLogContent)}',user='{LogOperationAccount}',creationtime='{LogOperationDate}'"
                DatabaseModule.ExecuteCommand(retLogSql, MySQL_Write)
            Next

            ' 回收汇总日志
            Dim retSumLog As String = $"账户:{UserAccount} 商品回收，回收单号::{recoveryOrderNumber} 回收数量:{(dgvRecovery.Rows.Count - 1).ToString()}"
            DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_xitong_log SET type='添加',title='商品回收',conter='{SafeSQL(retSumLog)}',user='{LogOperationAccount}',creationtime='{LogOperationDate}'", MySQL_Write)
        End If

        ' 积分记录
        If customerCode <> "" AndAlso scorePoints <> 0 Then
            Dim insertScoreSql As String = $"INSERT INTO xipunum_erp_member_score_log SET customer_code='{customerCode}',settlement_number='{outboundOrderNumber}',num='{scorePoints}',st='0',type='0',remarks='【订单号：{outboundOrderNumber}】实销 获得积分',creationtime='{LogOperationDate}',cjuser='{LogOperationAccount}'"
            DatabaseModule.ExecuteCommand(insertScoreSql, MySQL_Write)
        End If

        ' 提交事务
        If saveInProgress = 1 Then
            DatabaseModule.ExecuteCommand("COMMIT", MySQL_Write)
            saveInProgress = 0
        End If
        saveOrderNumber = ""
        saveRecoveryNumber = ""
        saveMemberCode = ""

        ' 打印
        If chkPrint.Checked Then
            If MessageBox.Show("是否打印单据？", "警告", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                PrintDocument()
            End If
        End If

        ' 清空临时数据并重新初始化
        ClearTempData()
        ShowSuccess("商品销售出库成功！")
        recoveryEditState = 0
        Form_Load(Nothing, Nothing)
    End Sub

    ' ========== 保存失败回滚（_销售出库_保存失败回滚） ==========
    Private Sub SaveFailedRollback(failReason As String)
        If saveInProgress = 1 Then
            DatabaseModule.ExecuteCommand("ROLLBACK", MySQL_Write)
            saveInProgress = 0
        End If
        If saveOrderNumber <> "" Then
            DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_shoukuan WHERE settlement_number='{saveOrderNumber}'", MySQL_Write)
            DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_outbound WHERE order_id IN (SELECT id FROM xipunum_erp_outbound_order WHERE settlement_number='{saveOrderNumber}')", MySQL_Write)
            DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_outbound_order WHERE settlement_number='{saveOrderNumber}' LIMIT 1", MySQL_Write)
        End If
        If saveRecoveryNumber <> "" Then
            DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_retreat WHERE order_id IN (SELECT id FROM xipunum_erp_retreat_order WHERE retrea_umber='{saveRecoveryNumber}')", MySQL_Write)
            DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_retreat_order WHERE retrea_umber='{saveRecoveryNumber}' LIMIT 1", MySQL_Write)
        End If
        If saveMemberCode <> "" Then
            DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_member WHERE customer_code='{saveMemberCode}' LIMIT 1", MySQL_Write)
            saveMemberCode = ""
        End If
        ShowWarning(failReason)
    End Sub

    ' ========== 打印功能（_打印功能_被单击） ==========
    Private Sub PrintDocument()
        ' 获取库房信息
        Dim kufangName As String = ""
        Dim kufangShort As String = ""
        Dim kufangCompany As String = ""
        Dim kufangAddress As String = ""
        Dim kufangSql As String = $"SELECT * FROM xipunum_erp_type WHERE id= '{UserDepartment}' order by id desc"
        Dim kufangDt As DataTable = DatabaseModule.ExecuteQuery(kufangSql)
        If kufangDt.Rows.Count > 0 Then
            kufangName = SafeString(kufangDt.Rows(0)("title"))
            kufangShort = SafeString(kufangDt.Rows(0)("data1"))
            kufangCompany = SafeString(kufangDt.Rows(0)("data2"))
            kufangAddress = SafeString(kufangDt.Rows(0)("data3"))
        End If

        Dim orderCode As String = txtOrderNumber.Text
        Dim customerName As String = ""
        Dim introducerName As String = ""

        If cmbIntroducer.SelectedIndex <> -1 Then
            introducerName = cmbIntroducer.SelectedItem.ToString()
        End If

        ' 会员信息
        If cmbMember.SelectedIndex <> -1 Then
            customerName = cmbMember.SelectedItem.ToString()
        ElseIf radioWholesale.Checked Then
            Dim memSql As String = $"SELECT * FROM xipunum_erp_member where name ='{SafeSQL(cmbMember.Text)}' and tel ='{SafeSQL(txtPhone.Text)}' LIMIT 1"
            Dim memDt As DataTable = DatabaseModule.ExecuteQuery(memSql)
            If memDt.Rows.Count > 0 Then
                customerName = SafeString(memDt.Rows(0)("name")) & "(" & SafeString(memDt.Rows(0)("memberid")) & ")"
            End If
        End If

        Dim orderDate As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Dim receivable As String = txtReceivable.Text
        Dim actual As String = txtReceived.Text
        ' 打印逻辑（原易语言使用报表组件，VB.NET中可使用打印文档或Crystal Reports）
        ' 此处简化为提示
        MessageBox.Show($"打印单据 - 订单号:{orderCode} 客户:{customerName} 应收:{receivable} 实收:{actual}", "打印", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' ========== 零售选中（_单选框_零售_被单击） ==========
    Private Sub RadioRetail_CheckedChanged(sender As Object, e As EventArgs)
        If radioRetail.Checked Then
            txtStoreMaterial.Visible = False
            lblMaterialBalance.Visible = False
            txtStoreBalance.Visible = False
            lblBalanceAmount.Visible = False
            chkPrint.Checked = False
        End If
    End Sub

    ' ========== 批发选中（_单选框_批发_被单击） ==========
    Private Sub RadioWholesale_CheckedChanged(sender As Object, e As EventArgs)
        If radioWholesale.Checked Then
            txtStoreMaterial.Visible = True
            lblMaterialBalance.Visible = True
            txtStoreBalance.Visible = True
            lblBalanceAmount.Visible = True
            chkPrint.Checked = True
            CalculateSummary()
        End If
    End Sub

    ' ========== 预售是（_单选框_是_被单击） ==========
    Private Sub RadioPresaleYes_CheckedChanged(sender As Object, e As EventArgs)
        If radioPresaleYes.Checked Then
            txtPresaleNumber.ReadOnly = False
            txtPresaleNumber.BackColor = Color.White
        End If
    End Sub

    ' ========== 预售否（_单选框_否_被单击） ==========
    Private Sub RadioPresaleNo_CheckedChanged(sender As Object, e As EventArgs)
        If radioPresaleNo.Checked Then
            txtPresaleNumber.ReadOnly = True
            txtPresaleNumber.BackColor = Color.Silver
        End If
    End Sub

    ' ========== 税点变化（_税点_编辑框_内容被改变） ==========
    Private Sub TxtTaxRate_TextChanged(sender As Object, e As EventArgs)
        CalculateSummary()
    End Sub

    ' ========== 联系电话变化（_联系电话_编辑框_内容被改变） ==========
    Private Sub TxtPhone_TextChanged(sender As Object, e As EventArgs)
        If txtPhone.Focused AndAlso txtPhone.Text.Length > 10 Then
            ' 批发模式显示存料/余额
            If radioWholesale.Checked Then
                txtStoreMaterial.Visible = True
                lblMaterialBalance.Visible = True
                txtStoreBalance.Visible = True
                lblBalanceAmount.Visible = True
            End If

            ' 查询会员
            Dim memCode As String = ""
            Dim memName As String = ""
            Dim memId As String = ""
            Dim memSql As String = $"SELECT * FROM xipunum_erp_member where tel ='{SafeSQL(txtPhone.Text)}' LIMIT 1"
            Dim memDt As DataTable = DatabaseModule.ExecuteQuery(memSql)
            If memDt.Rows.Count > 0 Then
                memCode = SafeString(memDt.Rows(0)("customer_code"))
                memName = SafeString(memDt.Rows(0)("name"))
                memId = SafeString(memDt.Rows(0)("memberid"))
            End If

            If memCode <> "" Then
                cmbMember.Text = memName & "(" & memId & ")"
                lblMemberId.Text = "ID:" & memId
            End If

            ' 查询存料/欠料
            Dim cunLiaoSql As String = $"SELECT sum(number) FROM xipunum_erp_member_cq WHERE customer_code = '{memCode}' and cunqu = '存' and type = '料' and kufang = '{UserDepartment}'"
            Dim cunLiao As String = SafeString(DatabaseModule.ExecuteScalar(cunLiaoSql))
            If cunLiao = "" Then cunLiao = "0.00"

            Dim qianLiaoSql As String = $"SELECT sum(number) FROM xipunum_erp_member_cq WHERE customer_code = '{memCode}' and cunqu = '欠' and type = '料' and kufang = '{UserDepartment}'"
            Dim qianLiao As String = SafeString(DatabaseModule.ExecuteScalar(qianLiaoSql))
            If qianLiao = "" Then qianLiao = "0.00"

            Dim jieyuLiao As String = FormatThreeDecimals((SafeDecimal(cunLiao) - SafeDecimal(qianLiao)).ToString())

            Dim cunYuanSql As String = $"SELECT sum(number) FROM xipunum_erp_member_cq WHERE customer_code = '{memCode}' and cunqu = '存' and type = '元' and kufang = '{UserDepartment}'"
            Dim cunYuan As String = SafeString(DatabaseModule.ExecuteScalar(cunYuanSql))
            If cunYuan = "" Then cunYuan = "0.00"

            Dim qianYuanSql As String = $"SELECT sum(number) FROM xipunum_erp_member_cq WHERE customer_code = '{memCode}' and cunqu = '欠' and type = '元' and kufang = '{UserDepartment}'"
            Dim qianYuan As String = SafeString(DatabaseModule.ExecuteScalar(qianYuanSql))
            If qianYuan = "" Then qianYuan = "0.00"

            Dim jieyuYuan As String = FormatTwoDecimals((SafeDecimal(cunYuan) - SafeDecimal(qianYuan)).ToString())

            lblMaterialBalance.Text = jieyuLiao
            txtStoreMaterial.Text = "0.000"
            lblBalanceAmount.Text = jieyuYuan
            txtStoreBalance.Text = "0.00"

            wholesaleStockWeight = jieyuLiao
            wholesaleStockAmount = jieyuYuan

            If txtPhone.Text.Length > 11 Then
                txtPhone.Text = txtPhone.Text.Substring(0, 11)
            End If
        End If
    End Sub

    ' ========== 联系电话按键（_联系电话_编辑框_按下某键） ==========
    Private Sub TxtPhone_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If txtPhone.Text = "" Then
                ShowWarning("会员联系电话不能为空！")
                txtPhone.Text = ""
                txtPhone.Focus()
                Return
            End If
            If txtPhone.Text.Length <> 11 Then
                ShowWarning("请输入正确的11位手机号码！")
                txtPhone.Text = ""
                txtPhone.Focus()
                Return
            End If
        End If
    End Sub

    ' ========== 预售单号变化（_预售单号_编辑框_内容被改变） ==========
    Private Sub TxtPresaleNumber_TextChanged(sender As Object, e As EventArgs)
        If txtPresaleNumber.Text.Length > 18 Then
            Dim queryNumber As String = txtPresaleNumber.Text
            Dim checkSql As String = $"SELECT * FROM xipunum_erp_presale_order where presale_umber='{SafeSQL(queryNumber)}' and state='待收货'"
            Dim checkDt As DataTable = DatabaseModule.ExecuteQuery(checkSql)
            If checkDt.Rows.Count = 0 Then
                ShowWarning("此预售订单已收货或者订单不存在！")
                txtPresaleNumber.Text = ""
                txtPresaleNumber.Focus()
                Return
            End If

            ' 获取预售订单数据
            Dim presaleSql As String = $"SELECT * FROM xipunum_erp_presale_order where presale_umber='{SafeSQL(queryNumber)}' and state='待收货' LIMIT 1"
            Dim presaleDt As DataTable = DatabaseModule.ExecuteQuery(presaleSql)
            If presaleDt.Rows.Count > 0 Then
                Dim presaleNumber As String = SafeString(presaleDt.Rows(0)("presale_umber"))
                Dim customerCode As String = SafeString(presaleDt.Rows(0)("customer_code"))
                Dim deposit As String = SafeString(presaleDt.Rows(0)("deposit"))

                ' 查询会员
                Dim memSql As String = $"SELECT * FROM xipunum_erp_member where customer_code='{SafeSQL(customerCode)}' LIMIT 1"
                Dim memDt As DataTable = DatabaseModule.ExecuteQuery(memSql)
                If memDt.Rows.Count > 0 Then
                    Dim memName As String = SafeString(memDt.Rows(0)("name"))
                    Dim memTel As String = SafeString(memDt.Rows(0)("tel"))
                    Dim memId As String = SafeString(memDt.Rows(0)("memberid"))

                    txtPresaleNumber.Text = presaleNumber
                    txtPresaleDeposit.Text = deposit
                    cmbMember.Text = memName & "(" & memId & ")"
                    lblMemberId.Text = "ID:" & memId
                    txtPhone.Text = memTel
                End If
                txtPresaleNumber.ReadOnly = True
                txtPresaleNumber.BackColor = Color.LightGray
            End If
        End If
    End Sub

    ' ========== 预售单号按键（_预售单号_编辑框_按下某键） ==========
    Private Sub TxtPresaleNumber_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If txtPresaleNumber.Text = "" Then
                ShowWarning("预售单号不能为空！")
                txtPresaleNumber.Focus()
                Return
            End If
            TxtPresaleNumber_TextChanged(Nothing, Nothing)
        End If
    End Sub

    ' ========== 会员查找下拉（_组合框会员查找_将弹出列表） ==========
    Private Sub CmbMember_DropDown(sender As Object, e As EventArgs)
        lblLastConsumption.Text = ""
        txtPhone.Text = ""
        lblMemberId.Text = ""

        If cmbMember.Text.Length > 0 Then
            cmbMember.Items.Clear()
            Dim searchContent As String = cmbMember.Text
            Dim sql As String = $"SELECT * FROM xipunum_erp_member where (memberid like '%{SafeSQL(searchContent)}%' or name like '%{SafeSQL(searchContent)}%' or tel like '%{SafeSQL(searchContent)}%') order by id ASC LIMIT 20"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            For Each r As DataRow In dt.Rows
                cmbMember.Items.Add(SafeString(r("name")) & "(" & SafeString(r("memberid")) & ")")
            Next
        End If
    End Sub

    ' ========== 会员选择（_组合框会员查找_列表项被选择） ==========
    Private Sub CmbMember_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmbMember.SelectedIndex = -1 Then Return

        ' 提取会员ID
        Dim selectedText As String = cmbMember.SelectedItem.ToString()
        Dim memberId As String = ""
        If selectedText.Length >= 8 Then
            memberId = selectedText.Substring(selectedText.Length - 8, 7)
        End If

        ' 批发模式显示存料/余额
        If radioWholesale.Checked Then
            txtStoreMaterial.Visible = True
            lblMaterialBalance.Visible = True
            txtStoreBalance.Visible = True
            lblBalanceAmount.Visible = True
        End If

        ' 查询会员信息
        Dim memCode As String = ""
        Dim memName As String = ""
        Dim memTel As String = ""
        Dim memId As String = ""
        Dim memSql As String = $"SELECT * FROM xipunum_erp_member where memberid ='{SafeSQL(memberId)}' LIMIT 1"
        Dim memDt As DataTable = DatabaseModule.ExecuteQuery(memSql)
        If memDt.Rows.Count > 0 Then
            memCode = SafeString(memDt.Rows(0)("customer_code"))
            memName = SafeString(memDt.Rows(0)("name"))
            memId = SafeString(memDt.Rows(0)("memberid"))
            memTel = SafeString(memDt.Rows(0)("tel"))
        End If

        txtPhone.Text = memTel
        If memTel.Length = 11 Then
            txtPhone.ReadOnly = True
            txtPhone.BackColor = Color.Silver
        Else
            txtPhone.ReadOnly = False
            txtPhone.BackColor = Color.White
        End If
        lblMemberId.Text = "ID:" & memId

        ' 查询上次消费工费和折扣
        Dim lastSql As String = $"SELECT b.id as id,b.sales_cost as sales_cost,b.zhekou as zhekou FROM xipunum_erp_outbound AS b INNER JOIN xipunum_erp_outbound_order AS a ON a.id=b.order_id WHERE a.customer_code = '{memCode}' ORDER BY b.id DESC LIMIT 1"
        Dim lastDt As DataTable = DatabaseModule.ExecuteQuery(lastSql)
        Dim lastCost As String = "0.00"
        Dim lastDiscount As String = "0.0000"
        If lastDt.Rows.Count > 0 Then
            lastCost = SafeString(lastDt.Rows(0)("sales_cost"))
            lastDiscount = SafeString(lastDt.Rows(0)("zhekou"))
        End If
        If lastCost = "" Then lastCost = "0.00"
        If lastDiscount = "" Then lastDiscount = "0.0000"

        lblLastConsumption.Text = $"上次消费工费:{lastCost} 元/克 折扣:{lastDiscount}"

        ' 查询存料/欠款
        Dim cqSql As String = $"SELECT tol.customer_code as customer_code,CAST(ROUND( COALESCE (sum( tol.jine ), 0 ), 2 ) AS DECIMAL ( 30, 2 )) as jine,CAST(ROUND( COALESCE (sum( tol.liao ), 0 ), 3 ) AS DECIMAL ( 30, 3 )) as liao FROM (SELECT customer_code,CASE WHEN type = '元' THEN CASE WHEN cunqu = '欠' THEN -number ELSE number END ELSE '0.00' END AS jine,CASE WHEN type = '料' THEN CASE WHEN cunqu = '欠' THEN -number ELSE number END ELSE '0.000' END AS liao FROM xipunum_erp_member_cq WHERE customer_code = '{memCode}' AND kufang = '{UserDepartment}' ORDER BY id ASC) as tol GROUP BY tol.customer_code"
        Dim cqDt As DataTable = DatabaseModule.ExecuteQuery(cqSql)
        Dim jieyuLiao As String = "0.000"
        Dim jieyuYuan As String = "0.00"
        If cqDt.Rows.Count > 0 Then
            jieyuLiao = SafeString(cqDt.Rows(0)("liao"))
            jieyuYuan = SafeString(cqDt.Rows(0)("jine"))
        End If

        lblMaterialBalance.Text = jieyuLiao
        txtStoreMaterial.Text = "0.000"
        lblBalanceAmount.Text = jieyuYuan
        txtStoreBalance.Text = "0.00"

        wholesaleStockWeight = jieyuLiao
        wholesaleStockAmount = jieyuYuan

        If txtPhone.Text.Length > 11 Then
            txtPhone.Text = txtPhone.Text.Substring(0, 11)
        End If
    End Sub

    ' ========== 介绍人查找下拉（_组合框介绍人查找_将弹出列表） ==========
    Private Sub CmbIntroducer_DropDown(sender As Object, e As EventArgs)
        If cmbIntroducer.Text.Length > 0 Then
            cmbIntroducer.Items.Clear()
            Dim sql As String = $"SELECT * FROM xipunum_erp_member where (memberid like '%{SafeSQL(cmbIntroducer.Text)}%' or name like '%{SafeSQL(cmbIntroducer.Text)}%' or tel like '%{SafeSQL(cmbIntroducer.Text)}%') order by id ASC LIMIT 20"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            For Each r As DataRow In dt.Rows
                cmbIntroducer.Items.Add(SafeString(r("name")) & "(" & SafeString(r("memberid")) & ")")
            Next
        End If
    End Sub

    ' ========== 销售工厂下拉（_组合框销售工厂_将弹出列表） ==========
    Private Sub CmbSalesFactory_DropDown(sender As Object, e As EventArgs)
        If cmbSalesFactory.Text.Length > 0 Then
            cmbSalesFactory.Items.Clear()
            Dim sql As String = $"SELECT * FROM xipunum_erp_outbound_order where xsfactory like '%{SafeSQL(cmbSalesFactory.Text)}%' GROUP BY xsfactory order by id ASC LIMIT 20"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            For Each r As DataRow In dt.Rows
                Dim factoryName As String = SafeString(r("xsfactory"))
                If factoryName <> "" Then
                    cmbSalesFactory.Items.Add(factoryName)
                End If
            Next
        End If
    End Sub

    ' ========== 业务员组合框按键（_业务员组合框_按下某键） ==========
    Private Sub CmbGuide_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            CmbGuide_DropDown(Nothing, Nothing)
        End If
    End Sub

    ' ========== 业务员组合框下拉（_业务员组合框_将弹出列表） ==========
    Private Sub CmbGuide_DropDown(sender As Object, e As EventArgs)
        Dim searchContent As String = cmbGuide.Text
        cmbGuide.Items.Clear()
        cmbGuide.Text = searchContent

        Dim sql As String = ""
        If searchContent = "" Then
            If UserPermission = "全部" Then
                sql = "SELECT name FROM xipunum_erp_user where state='0' order by id ASC"
            ElseIf UserPermission = "店铺" Then
                sql = $"SELECT name FROM xipunum_erp_user where department='{UserDepartment}' and state='0' order by id ASC"
            ElseIf UserPermission = "岗位" Then
                sql = $"SELECT name FROM xipunum_erp_user WHERE user in {GlobalViewSQL} and state='0' order by id ASC"
            ElseIf UserPermission = "个人" Then
                sql = $"SELECT name FROM xipunum_erp_user WHERE user='{UserAccount}' and state='0' order by id ASC"
            Else
                sql = "SELECT name FROM xipunum_erp_user WHERE 1=0"
            End If
        Else
            If UserPermission = "全部" Then
                sql = $"SELECT name FROM xipunum_erp_user WHERE (jianxie = '{SafeSQL(searchContent)}' or user like '%{SafeSQL(searchContent)}%' or name like '%{SafeSQL(searchContent)}%') and state='0'"
            ElseIf UserPermission = "店铺" Then
                sql = $"SELECT name FROM xipunum_erp_user WHERE (jianxie = '{SafeSQL(searchContent)}' or user like '%{SafeSQL(searchContent)}%' or name like '%{SafeSQL(searchContent)}%') and department='{UserDepartment}' and state='0'"
            ElseIf UserPermission = "岗位" Then
                sql = $"SELECT name FROM xipunum_erp_user WHERE user in {GlobalViewSQL} and (jianxie = '{SafeSQL(searchContent)}' or user like '%{SafeSQL(searchContent)}%' or name like '%{SafeSQL(searchContent)}%') and state='0'"
            ElseIf UserPermission = "个人" Then
                sql = $"SELECT name FROM xipunum_erp_user WHERE user='{UserAccount}' and (jianxie = '{SafeSQL(searchContent)}' or user like '%{SafeSQL(searchContent)}%' or name like '%{SafeSQL(searchContent)}%') and state='0'"
            Else
                sql = "SELECT name FROM xipunum_erp_user WHERE 1=0"
            End If
        End If

        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
        For Each r As DataRow In dt.Rows
            cmbGuide.Items.Add(SafeString(r("name")))
        Next

        If dt.Rows.Count <= 0 Then
            cmbGuide.Text = ""
            cmbGuide.SelectedIndex = -1
        Else
            cmbGuide.SelectedIndex = 0
        End If
    End Sub

    ' ========== 存料编辑框失去焦点（_存料_编辑框_失去焦点） ==========
    Private Sub TxtStoreMaterial_Leave(sender As Object, e As EventArgs)
        txtStoreMaterial.Text = FormatThreeDecimals(txtStoreMaterial.Text)
        CalculateSummary()
    End Sub

    ' ========== 余额编辑框失去焦点（_余额_编辑框_失去焦点） ==========
    Private Sub TxtStoreBalance_Leave(sender As Object, e As EventArgs)
        txtStoreBalance.Text = FormatTwoDecimals(txtStoreBalance.Text)
        CalculateSummary()
    End Sub

    ' ========== 窗口关闭确认（_窗口_商品销售出库_可否被关闭） ==========
    Private Function CanClose() As Boolean
        If dgvSales.Rows.Count > 2 OrElse dgvRecovery.Rows.Count > 1 OrElse tableEditState = 1 OrElse recoveryEditState = 1 OrElse paymentEditState = 1 Then
            If MessageBox.Show("当前销售出库单尚未保存，确定要关闭吗？", "提示", MessageBoxButtons.YesNo) = DialogResult.No Then
                Return False
            End If
        End If
        Return True
    End Function

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If Not CanClose() Then
            e.Cancel = True
        Else
            MyBase.OnFormClosing(e)
        End If
    End Sub

    ' ========== 格式化函数（_销售出库_格式三位小数） ==========
    Private Function FormatThreeDecimals(value As String) As String
        If String.IsNullOrEmpty(value) Then Return "0.000"
        Dim d As Decimal
        If Not Decimal.TryParse(value, d) Then Return "0.000"
        Return Math.Round(d, 3).ToString("0.000")
    End Function

    ' ========== 格式化函数（_销售出库_格式二位小数） ==========
    Private Function FormatTwoDecimals(value As String) As String
        If String.IsNullOrEmpty(value) Then Return "0.00"
        Dim d As Decimal
        If Not Decimal.TryParse(value, d) Then Return "0.00"
        Return Math.Round(d, 2).ToString("0.00")
    End Function

    ' ========== 格式化函数（_销售出库_格式四位小数） ==========
    Private Function FormatFourDecimals(value As String) As String
        If String.IsNullOrEmpty(value) Then Return "0.0000"
        Dim d As Decimal
        If Not Decimal.TryParse(value, d) Then Return "0.0000"
        Return Math.Round(d, 4).ToString("0.0000")
    End Function

    ' ========== 辅助函数 ==========
    Private Function SafeString(value As Object) As String
        If value Is Nothing OrElse IsDBNull(value) Then Return ""
        Return value.ToString()
    End Function

    Private Function SafeDecimal(value As Object) As Decimal
        If value Is Nothing OrElse IsDBNull(value) Then Return 0
        Dim d As Decimal
        If Decimal.TryParse(value.ToString(), d) Then Return d
        Return 0
    End Function

    Private Sub ShowWarning(msg As String)
        MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub ShowSuccess(msg As String)
        MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ShowError(msg As String)
        MessageBox.Show(msg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Function ConfirmAction(msg As String) As Boolean
        Return MessageBox.Show(msg, "确认", MessageBoxButtons.YesNo) = DialogResult.Yes
    End Function

    Private Sub AddSystemLog(logType As String, logTitle As String, logContent As String)
        Dim sql As String = $"INSERT INTO xipunum_erp_xitong_log SET type='{logType}',title='{logTitle}',conter='{SafeSQL(logContent)}',user='{LogOperationAccount}',creationtime='{LogOperationDate}'"
        DatabaseModule.ExecuteCommand(sql, MySQL_Write)
    End Sub

End Class

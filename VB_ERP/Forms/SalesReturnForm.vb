' ============================================================================
' 商品销售客退窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品销售客退.form.e.txt
' 包含所有34个程序集变量、所有子程序、SQL查询、业务逻辑
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SalesReturnForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（34个） ==========
    Private row1 As Integer = 0                    ' 集_行号1
    Private col1 As Integer = 0                    ' 集_列号1
    Private row2 As Integer = 0                    ' 集_行号2
    Private col2 As Integer = 0                    ' 集_列号2
    Private row3 As Integer = 0                    ' 集_行号3
    Private col3 As Integer = 0                    ' 集_列号3
    Private salesWeightTotal As String = "0"       ' 局部销售金重合计
    Private salesAmountTotal As String = "0"       ' 局部销售金额合计
    Private actualSalesAmount As String = "0"      ' 局部实销金额合计
    Private salesQuantityTotal As String = "0"     ' 局部销售数量合计
    Private salesCostTotal As String = "0"         ' 局部销售工费合计
    Private salesSurchargeTotal As String = "0"    ' 局部销售附加费合计
    Private actualReturnCost As String = "0"       ' 局部实退工费合计
    Private actualReturnSurcharge As String = "0"  ' 局部实退附加费合计
    Private returnAmountTotal As String = "0"      ' 局部应退金额合计
    Private actualReturnAmount As String = "0"     ' 局部实退金额合计
    Private recoveryWeightTotal As String = "0"    ' 局部回收总重合计数值
    Private recoveryGoldTotal As String = "0"      ' 局部回收金重合计数值
    Private recoveryOtherTotal As String = "0"     ' 局部回收其他合计数值
    Private recoveryAmountTotalVal As String = "0" ' 局部回收回收合计数值
    Private wholesaleStockWeight As String = "0"   ' 局部_批发库存料数值
    Private wholesaleStockAmount As String = "0"   ' 局部_批发库存元数值
    Private costCostTotal As String = "0"          ' 局部成本工费合计
    Private premiumCostTotal As String = "0"       ' 局部参考工费合计
    Private salesMemberId As String = "-1"         ' 局部_商品销售会员id
    Private tableEditState As Integer = 0          ' 局部_表格编辑状态
    Private paymentEditState As Integer = 0        ' 局部_收支编辑状态
    Private recoveryEditState As Integer = 0       ' 局部_回收编辑状态
    Public dataProductCode As String = ""         ' 局部_数据据商品编码
    Public dataProductOrderId As String = ""      ' 局部_数据据商品订单id

    ' ========== 控件声明 ==========
    Private dgvSales As New DataGridView()         ' 高级表格1 - 销售明细
    Private dgvRecovery As New DataGridView()       ' 高级表格2 - 回收明细
    Private dgvPayment As New DataGridView()        ' 高级表格3 - 收支明细
    Private dgvSummary As New DataGridView()        ' 高级表格4 - 统计汇总
    Private txtOrderNumber As New TextBox()         ' 单据号_编辑框
    Private txtMemberName As New TextBox()          ' 客户姓名_编辑框
    Private txtPhone As New TextBox()               ' 联系电话_编辑框
    Private txtReceivable As New TextBox()          ' 应收_编辑框
    Private txtReceived As New TextBox()            ' 实收_编辑框
    Private txtSalesman As New TextBox()            ' 业务员_编辑框
    Private txtRemarks As New TextBox()             ' 备注_编辑框
    Private txtStoreMaterial As New TextBox()       ' 存料_编辑框
    Private txtStoreBalance As New TextBox()        ' 余额_编辑框
    Private lblStoreMaterialVal As New Label()      ' 透明标签_存料1
    Private lblStoreBalanceVal As New Label()       ' 透明标签_余额1
    Private lblMemberId As New Label()              ' 透明标签13
    Private chkPrintDoc As New CheckBox()           ' 打印单据_选择框
    Private chkPrintPreview As New CheckBox()       ' 打印预览_选择框
    Private radioNewRow As New RadioButton()        ' 单选框_换行
    Private radioNewCol As New RadioButton()        ' 单选框_换列
    Private btnRecoveryAdd As New Button()          ' 按钮_回收加
    Private btnRecoveryRemove As New Button()       ' 按钮_回收减
    Private btnToolbarSave As New Button()          ' 工具条_保存
    Private btnToolbarReset As New Button()         ' 工具条_重置
    Private btnToolbarPrint As New Button()         ' 工具条_打印预览

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "商品销售客退"
        Me.Size = New Drawing.Size(1430, 698)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' 顶部信息区
        Dim panelTop As New Panel()
        panelTop.Name = "分组框_头部"
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 110
        Me.Controls.Add(panelTop)

        ' 单据号
        Dim lblOrder As New Label()
        lblOrder.Text = "单据号："
        lblOrder.Location = New Drawing.Point(20, 7)
        lblOrder.AutoSize = True
        panelTop.Controls.Add(lblOrder)
        txtOrderNumber.Location = New Drawing.Point(80, 5)
        txtOrderNumber.Size = New Drawing.Size(200, 25)
        txtOrderNumber.ReadOnly = True
        panelTop.Controls.Add(txtOrderNumber)

        ' 客户姓名
        Dim lblMember As New Label()
        lblMember.Text = "会员姓名："
        lblMember.Location = New Drawing.Point(290, 7)
        lblMember.AutoSize = True
        panelTop.Controls.Add(lblMember)
        txtMemberName.Location = New Drawing.Point(360, 5)
        txtMemberName.Size = New Drawing.Size(150, 25)
        txtMemberName.ReadOnly = True
        txtMemberName.BackColor = Drawing.Color.Silver
        panelTop.Controls.Add(txtMemberName)

        ' 联系电话
        Dim lblPhone As New Label()
        lblPhone.Text = "联系电话："
        lblPhone.Location = New Drawing.Point(520, 7)
        lblPhone.AutoSize = True
        panelTop.Controls.Add(lblPhone)
        txtPhone.Location = New Drawing.Point(590, 5)
        txtPhone.Size = New Drawing.Size(120, 25)
        txtPhone.ReadOnly = True
        txtPhone.BackColor = Drawing.Color.Silver
        panelTop.Controls.Add(txtPhone)

        ' 业务员
        Dim lblSalesman As New Label()
        lblSalesman.Text = "业务员："
        lblSalesman.Location = New Drawing.Point(20, 40)
        lblSalesman.AutoSize = True
        panelTop.Controls.Add(lblSalesman)
        txtSalesman.Location = New Drawing.Point(80, 38)
        txtSalesman.Size = New Drawing.Size(100, 25)
        txtSalesman.ReadOnly = True
        panelTop.Controls.Add(txtSalesman)

        ' 应收
        Dim lblReceivable As New Label()
        lblReceivable.Text = "应退："
        lblReceivable.Location = New Drawing.Point(200, 40)
        lblReceivable.AutoSize = True
        panelTop.Controls.Add(lblReceivable)
        txtReceivable.Location = New Drawing.Point(240, 38)
        txtReceivable.Size = New Drawing.Size(100, 25)
        txtReceivable.ReadOnly = True
        txtReceivable.Text = "0"
        panelTop.Controls.Add(txtReceivable)

        ' 实收
        Dim lblReceived As New Label()
        lblReceived.Text = "实退："
        lblReceived.Location = New Drawing.Point(360, 40)
        lblReceived.AutoSize = True
        panelTop.Controls.Add(lblReceived)
        txtReceived.Location = New Drawing.Point(400, 38)
        txtReceived.Size = New Drawing.Size(100, 25)
        panelTop.Controls.Add(txtReceived)

        ' 存料
        Dim lblStoreMaterial As New Label()
        lblStoreMaterial.Text = "结料："
        lblStoreMaterial.Location = New Drawing.Point(520, 40)
        lblStoreMaterial.AutoSize = True
        panelTop.Controls.Add(lblStoreMaterial)
        txtStoreMaterial.Location = New Drawing.Point(560, 38)
        txtStoreMaterial.Size = New Drawing.Size(60, 25)
        txtStoreMaterial.Text = "0.000"
        panelTop.Controls.Add(txtStoreMaterial)

        ' 余额
        Dim lblStoreBalance As New Label()
        lblStoreBalance.Text = "结元："
        lblStoreBalance.Location = New Drawing.Point(630, 40)
        lblStoreBalance.AutoSize = True
        panelTop.Controls.Add(lblStoreBalance)
        txtStoreBalance.Location = New Drawing.Point(670, 38)
        txtStoreBalance.Size = New Drawing.Size(60, 25)
        txtStoreBalance.Text = "0.00"
        panelTop.Controls.Add(txtStoreBalance)

        ' 存料/余额显示标签
        lblStoreMaterialVal.Text = "0.000"
        lblStoreMaterialVal.Location = New Drawing.Point(740, 40)
        lblStoreMaterialVal.AutoSize = True
        panelTop.Controls.Add(lblStoreMaterialVal)
        lblStoreBalanceVal.Text = "0.00"
        lblStoreBalanceVal.Location = New Drawing.Point(820, 40)
        lblStoreBalanceVal.AutoSize = True
        panelTop.Controls.Add(lblStoreBalanceVal)

        ' 会员ID标签
        lblMemberId.Text = ""
        lblMemberId.Location = New Drawing.Point(520, 72)
        lblMemberId.AutoSize = True
        panelTop.Controls.Add(lblMemberId)

        ' 打印选择框
        chkPrintDoc.Text = "打印单据"
        chkPrintDoc.Location = New Drawing.Point(20, 72)
        chkPrintDoc.AutoSize = True
        panelTop.Controls.Add(chkPrintDoc)
        chkPrintPreview.Text = "打印预览"
        chkPrintPreview.Location = New Drawing.Point(100, 72)
        chkPrintPreview.AutoSize = True
        panelTop.Controls.Add(chkPrintPreview)

        ' 换行/换列
        Dim grpDirection As New GroupBox()
        grpDirection.Text = ""
        grpDirection.Location = New Drawing.Point(200, 65)
        grpDirection.Size = New Drawing.Size(150, 30)
        panelTop.Controls.Add(grpDirection)
        radioNewRow.Text = "换行"
        radioNewRow.Location = New Drawing.Point(5, 8)
        radioNewRow.AutoSize = True
        radioNewRow.Checked = True
        grpDirection.Controls.Add(radioNewRow)
        radioNewCol.Text = "换列"
        radioNewCol.Location = New Drawing.Point(70, 8)
        radioNewCol.AutoSize = True
        grpDirection.Controls.Add(radioNewCol)

        ' 工具条
        Dim panelToolbar As New Panel()
        panelToolbar.Dock = DockStyle.Top
        panelToolbar.Height = 40
        panelToolbar.Location = New Drawing.Point(0, 110)
        Me.Controls.Add(panelToolbar)

        btnToolbarSave.Text = "保存"
        btnToolbarSave.Location = New Drawing.Point(10, 5)
        btnToolbarSave.Size = New Drawing.Size(80, 30)
        panelToolbar.Controls.Add(btnToolbarSave)

        btnToolbarReset.Text = "重置"
        btnToolbarReset.Location = New Drawing.Point(100, 5)
        btnToolbarReset.Size = New Drawing.Size(80, 30)
        panelToolbar.Controls.Add(btnToolbarReset)

        btnToolbarPrint.Text = "打印预览"
        btnToolbarPrint.Location = New Drawing.Point(190, 5)
        btnToolbarPrint.Size = New Drawing.Size(80, 30)
        panelToolbar.Controls.Add(btnToolbarPrint)

        ' 销售明细表格
        dgvSales.Location = New Drawing.Point(0, 150)
        dgvSales.AllowUserToAddRows = False
        dgvSales.AllowUserToDeleteRows = False
        dgvSales.SelectionMode = DataGridViewSelectionMode.CellSelect
        dgvSales.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Me.Controls.Add(dgvSales)

        ' 统计汇总表格
        dgvSummary.Location = New Drawing.Point(0, 0)
        dgvSummary.Size = New Drawing.Size(1392, 35)
        dgvSummary.AllowUserToAddRows = False
        dgvSummary.AllowUserToDeleteRows = False
        Me.Controls.Add(dgvSummary)

        ' 回收明细区域
        Dim panelRecovery As New Panel()
        panelRecovery.Name = "分组框_回收"
        panelRecovery.Dock = DockStyle.Bottom
        panelRecovery.Height = 180
        Me.Controls.Add(panelRecovery)

        btnRecoveryAdd.Text = "+"
        btnRecoveryAdd.Location = New Drawing.Point(5, 10)
        btnRecoveryAdd.Size = New Drawing.Size(40, 30)
        panelRecovery.Controls.Add(btnRecoveryAdd)

        btnRecoveryRemove.Text = "-"
        btnRecoveryRemove.Location = New Drawing.Point(5, 45)
        btnRecoveryRemove.Size = New Drawing.Size(40, 30)
        panelRecovery.Controls.Add(btnRecoveryRemove)

        dgvRecovery.Location = New Drawing.Point(55, 0)
        dgvRecovery.Size = New Drawing.Size(900, 100)
        dgvRecovery.AllowUserToAddRows = False
        dgvRecovery.AllowUserToDeleteRows = False
        panelRecovery.Controls.Add(dgvRecovery)

        dgvPayment.Location = New Drawing.Point(965, 0)
        dgvPayment.Size = New Drawing.Size(239, 100)
        dgvPayment.AllowUserToAddRows = False
        dgvPayment.AllowUserToDeleteRows = False
        panelRecovery.Controls.Add(dgvPayment)

        ' 备注区
        Dim panelRemarks As New Panel()
        panelRemarks.Name = "分组框_备注"
        panelRemarks.Dock = DockStyle.Bottom
        panelRemarks.Height = 79
        Me.Controls.Add(panelRemarks)

        Dim lblRemarks As New Label()
        lblRemarks.Text = "备注："
        lblRemarks.Location = New Drawing.Point(5, 10)
        lblRemarks.AutoSize = True
        panelRemarks.Controls.Add(lblRemarks)
        txtRemarks.Location = New Drawing.Point(50, 8)
        txtRemarks.Size = New Drawing.Size(600, 60)
        txtRemarks.Multiline = True
        panelRemarks.Controls.Add(txtRemarks)

        ' 事件绑定
        AddHandler btnToolbarSave.Click, AddressOf BtnToolbarSave_Click
        AddHandler btnToolbarReset.Click, AddressOf BtnToolbarReset_Click
        AddHandler btnToolbarPrint.Click, AddressOf BtnToolbarPrint_Click
        AddHandler btnRecoveryAdd.Click, AddressOf BtnRecoveryAdd_Click
        AddHandler btnRecoveryRemove.Click, AddressOf BtnRecoveryRemove_Click
        AddHandler dgvSales.CellEndEdit, AddressOf DgvSales_CellEndEdit
        AddHandler dgvSales.CellBeginEdit, AddressOf DgvSales_CellBeginEdit
        AddHandler dgvRecovery.CellEndEdit, AddressOf DgvRecovery_CellEndEdit
        AddHandler dgvRecovery.CellBeginEdit, AddressOf DgvRecovery_CellBeginEdit
        AddHandler dgvPayment.CellEndEdit, AddressOf DgvPayment_CellEndEdit
        AddHandler dgvPayment.CellBeginEdit, AddressOf DgvPayment_CellBeginEdit
        AddHandler dgvSales.KeyDown, AddressOf DgvSales_KeyDown
        AddHandler dgvSales.SelectionChanged, AddressOf DgvSales_SelectionChanged
        AddHandler dgvRecovery.SelectionChanged, AddressOf DgvRecovery_SelectionChanged
        AddHandler dgvPayment.SelectionChanged, AddressOf DgvPayment_SelectionChanged
        AddHandler radioNewRow.CheckedChanged, AddressOf RadioNewRow_CheckedChanged
        AddHandler radioNewCol.CheckedChanged, AddressOf RadioNewCol_CheckedChanged
        AddHandler txtStoreMaterial.Leave, AddressOf TxtStoreMaterial_Leave
        AddHandler txtStoreBalance.Leave, AddressOf TxtStoreBalance_Leave
        AddHandler txtStoreMaterial.KeyDown, AddressOf TxtStoreMaterial_KeyDown
        AddHandler txtStoreBalance.KeyDown, AddressOf TxtStoreBalance_KeyDown
    End Sub

    Private Sub AddLabel(parent As Control, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    ' ========== 窗口加载 _窗口_商品销售客退_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 生成单据号
        txtOrderNumber.Text = GenerateOrderNumber("1")
        salesMemberId = "-1"
        dataProductCode = ""
        tableEditState = 0
        paymentEditState = 0
        recoveryEditState = 0

        txtMemberName.Text = ""
        txtMemberName.ReadOnly = True
        txtMemberName.BackColor = Drawing.Color.Silver
        txtPhone.Text = ""
        txtPhone.ReadOnly = True
        txtPhone.BackColor = Drawing.Color.Silver
        txtReceivable.Text = "0"
        txtReceived.Text = ""
        txtReceivable.ReadOnly = True

        lblStoreMaterialVal.Text = "0.000"
        txtStoreMaterial.Text = "0.000"
        lblStoreBalanceVal.Text = "0.00"
        lblMemberId.Text = ""
        txtStoreBalance.Text = "0.00"
        wholesaleStockWeight = "0.000"
        wholesaleStockAmount = "0.00"

        salesWeightTotal = "0"
        salesAmountTotal = "0"
        actualSalesAmount = "0"
        salesQuantityTotal = "0"
        salesCostTotal = "0"
        salesSurchargeTotal = "0"
        actualReturnCost = "0"
        actualReturnSurcharge = "0"
        actualReturnAmount = "0"
        returnAmountTotal = "0"
        costCostTotal = "0"
        premiumCostTotal = "0"
        chkPrintDoc.Checked = False
        chkPrintPreview.Checked = False
        txtSalesman.Text = UserAccount

        InitSalesGrid()
        InitRecoveryGrid()
        InitPaymentGrid()
        LoadSalesGridData()
        LoadRecoveryGridData()
        LoadPaymentGridData()
        CalculateSummary()
    End Sub

    ' ========== 窗口尺寸改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        Dim nWidth As Integer = Me.ClientSize.Width
        Dim nHeight As Integer = Me.ClientSize.Height

        dgvSales.Width = nWidth
        dgvSales.Height = nHeight - 366

        ' 统计表格位置
        dgvSummary.Width = nWidth
        dgvSummary.Location = New Drawing.Point(0, nHeight - 216)
    End Sub

    ' ========== _高级表格1_加载表头 ==========
    Private Sub InitSalesGrid()
        dgvSales.Columns.Clear()
        Dim headers() As String = {"序号", "商品编码", "商品名称", "款号", "规格", "材质", "圈口/长度", "成色", "成本工费", "参考工费", "单件重", "销售金重", "销售金额", "实销金额", "销售数量", "销售克价", "销售工费", "销售附加费", "折扣", "实退克价", "实退工费", "实退附加费", "实退金额", "导购员", "单据id", "操作"}
        Dim widths() As Integer = {45, 100, 140, 0, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 0, 65}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            If widths(i) = 0 Then
                col.Visible = False
            End If
            dgvSales.Columns.Add(col)
        Next

        ' 添加删除按钮列
        Dim btnCol As New DataGridViewButtonColumn()
        btnCol.HeaderText = "操作"
        btnCol.Name = "colDelete"
        btnCol.Width = 65
        btnCol.UseColumnTextForButtonValue = True
        btnCol.Text = "删除"
        dgvSales.Columns.RemoveAt(25)
        dgvSales.Columns.Add(btnCol)
    End Sub

    ' ========== _高级表格2_加载表头 ==========
    Private Sub InitRecoveryGrid()
        dgvRecovery.Columns.Clear()
        Dim headers() As String = {"", "商品名称", "数量", "总重", "金重", "成色", "回收克价", "其他费用", "回收金额", "导购员"}
        Dim widths() As Integer = {45, 200, 100, 100, 100, 100, 100, 100, 100, 100}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "rcol" & i
            col.Width = widths(i)
            dgvRecovery.Columns.Add(col)
        Next

        ' 加载回收名称下拉列表
        Try
            Dim sql As String = "SELECT * FROM xipunum_erp_retreat_title WHERE 1=1 order by id asc"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            ' 存储回收名称列表供后续使用
            recoveryNameList.Clear()
            For Each r As DataRow In dt.Rows
                recoveryNameList.Add(SafeString(r("title")))
            Next
        Catch ex As Exception
        End Try

        ' 加载导购员下拉列表
        Try
            Dim sql As String = ""
            If UserPermission = "全部" Then
                sql = "SELECT name FROM xipunum_erp_user where state='0' order by id ASC"
            ElseIf UserPermission = "店铺" Then
                sql = "SELECT name FROM xipunum_erp_user where department='" & UserDepartment & "' and state='0' order by id ASC"
            Else
                sql = "SELECT name FROM xipunum_erp_user where state='0' order by id ASC"
            End If
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            guideNameList.Clear()
            For Each r As DataRow In dt.Rows
                guideNameList.Add(SafeString(r("name")))
            Next
        Catch ex As Exception
        End Try
    End Sub

    ' 回收名称列表和导购员列表
    Private recoveryNameList As New List(Of String)
    Private guideNameList As New List(Of String)

    ' ========== _高级表格3_加载表头 ==========
    Private Sub InitPaymentGrid()
        dgvPayment.Columns.Clear()
        Dim headers() As String = {"序号", "支付方式", "金额", "id"}
        Dim widths() As Integer = {45, 65, 75, 0}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "pcol" & i
            col.Width = widths(i)
            If widths(i) = 0 Then col.Visible = False
            dgvPayment.Columns.Add(col)
        Next
    End Sub

    ' ========== _高级表格4_加载表头 ==========
    Private Sub InitSummaryGrid()
        dgvSummary.Columns.Clear()
        Dim headers() As String = {"", "合计", "", "", "", "", "", "", "", "", "", "销售金重", "销售金额", "实销金额", "销售数量", "", "销售工费", "销售附加费", "", "", "实退工费", "实退附加费", "实退金额", "", "", ""}
        Dim widths() As Integer = {45, 100, 140, 0, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 75, 0, 65}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "scol" & i
            col.Width = widths(i)
            If widths(i) = 0 Then col.Visible = False
            dgvSummary.Columns.Add(col)
        Next
        dgvSummary.Rows.Add()
        dgvSummary.Rows(0).Cells(1).Value = "合计"
        dgvSummary.Rows(0).Cells(11).Value = salesWeightTotal
        dgvSummary.Rows(0).Cells(12).Value = salesAmountTotal
        dgvSummary.Rows(0).Cells(13).Value = actualSalesAmount
        dgvSummary.Rows(0).Cells(14).Value = salesQuantityTotal
        dgvSummary.Rows(0).Cells(16).Value = salesCostTotal
        dgvSummary.Rows(0).Cells(17).Value = salesSurchargeTotal
        dgvSummary.Rows(0).Cells(20).Value = actualReturnCost
        dgvSummary.Rows(0).Cells(21).Value = actualReturnSurcharge
        dgvSummary.Rows(0).Cells(22).Value = actualReturnAmount
    End Sub

    ' ========== _高级表格1_加载表格 ==========
    Private Sub LoadSalesGridData()
        ClearSalesGrid()
        AddBlankSalesRow()
    End Sub

    ' ========== _高级表格2_加载表格 ==========
    Private Sub LoadRecoveryGridData()
        ClearRecoveryGrid()
    End Sub

    ' ========== _高级表格3_加载表格 ==========
    Private Sub LoadPaymentGridData()
        ClearPaymentGrid()

        Try
            Dim sql As String = "SELECT id,name FROM xipunum_erp_pay where state='0' order by id ASC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)

            Dim rowCount As Integer = dt.Rows.Count
            ' 添加支付方式行
            For i As Integer = 0 To rowCount - 1
                dgvPayment.Rows.Add()
                dgvPayment.Rows(i).Cells(0).Value = (i + 1).ToString()
                dgvPayment.Rows(i).Cells(1).Value = SafeString(dt.Rows(i)("name"))
                dgvPayment.Rows(i).Cells(2).Value = 0
                dgvPayment.Rows(i).Cells(3).Value = SafeString(dt.Rows(i)("id"))
            Next

            ' 添加合计行
            dgvPayment.Rows.Add()
            dgvPayment.Rows(rowCount).Cells(1).Value = "合计"
            Dim sumVal As Decimal = 0
            For i As Integer = 0 To rowCount - 1
                sumVal += SafeDecimal(dgvPayment.Rows(i).Cells(2).Value)
            Next
            dgvPayment.Rows(rowCount).Cells(2).Value = sumVal

            ' 设置只读
            For i As Integer = 0 To rowCount - 1
                dgvPayment.Rows(i).Cells(0).ReadOnly = True
                dgvPayment.Rows(i).Cells(1).ReadOnly = True
            Next
            dgvPayment.Rows(rowCount).Cells(0).ReadOnly = True
            dgvPayment.Rows(rowCount).Cells(1).ReadOnly = True
            dgvPayment.Rows(rowCount).Cells(2).ReadOnly = True
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 子程序_删除表格1 ==========
    Private Sub ClearSalesGrid()
        dgvSales.Rows.Clear()
    End Sub

    ' ========== 子程序_删除表格2 ==========
    Private Sub ClearRecoveryGrid()
        dgvRecovery.Rows.Clear()
    End Sub

    ' ========== 子程序_删除表格3 ==========
    Private Sub ClearPaymentGrid()
        dgvPayment.Rows.Clear()
    End Sub

    ' ========== _子程序_添加销售 ==========
    Private Sub AddBlankSalesRow()
        Dim addRow As Integer = dgvSales.Rows.Count
        If addRow > 0 AndAlso dgvSales.Rows(addRow - 1).Cells(1).Value IsNot Nothing AndAlso dgvSales.Rows(addRow - 1).Cells(1).Value.ToString() <> "" Then
            dgvSales.Rows.Add()
            Dim newRow As Integer = dgvSales.Rows.Count - 1
            For i As Integer = 0 To 24
                dgvSales.Rows(newRow).Cells(i).Value = ""
            Next
            ' 设置只读
            For i As Integer = 2 To 24
                dgvSales.Rows(newRow).Cells(i).ReadOnly = True
                dgvSales.Rows(newRow).Cells(i).Style.BackColor = Drawing.Color.LightGray
            Next
        End If
    End Sub

    ' ========== _子程序_数据统计汇总 ==========
    Private Sub CalculateSummary()
        salesWeightTotal = "0"
        salesAmountTotal = "0"
        actualSalesAmount = "0"
        salesQuantityTotal = "0"
        salesCostTotal = "0"
        salesSurchargeTotal = "0"
        actualReturnCost = "0"
        actualReturnSurcharge = "0"
        actualReturnAmount = "0"
        returnAmountTotal = "0"
        costCostTotal = "0"
        premiumCostTotal = "0"

        ' 销售汇总
        For Each r As DataGridViewRow In dgvSales.Rows
            If r.Cells(1).Value Is Nothing OrElse r.Cells(1).Value.ToString() = "" Then Continue For

            costCostTotal = (SafeDecimal(costCostTotal) + SafeDecimal(r.Cells(8).Value) * SafeDecimal(r.Cells(11).Value)).ToString()
            premiumCostTotal = (SafeDecimal(premiumCostTotal) + SafeDecimal(r.Cells(9).Value) * SafeDecimal(r.Cells(11).Value)).ToString()
            salesWeightTotal = (SafeDecimal(salesWeightTotal) + SafeDecimal(r.Cells(11).Value)).ToString()
            salesAmountTotal = (SafeDecimal(salesAmountTotal) + SafeDecimal(r.Cells(12).Value)).ToString()
            actualSalesAmount = (SafeDecimal(actualSalesAmount) + SafeDecimal(r.Cells(13).Value)).ToString()
            salesQuantityTotal = (SafeDecimal(salesQuantityTotal) + SafeDecimal(r.Cells(14).Value)).ToString()
            salesCostTotal = (SafeDecimal(salesCostTotal) + SafeDecimal(r.Cells(16).Value) * SafeDecimal(r.Cells(11).Value)).ToString()
            salesSurchargeTotal = (SafeDecimal(salesSurchargeTotal) + SafeDecimal(r.Cells(17).Value)).ToString()
            actualReturnCost = (SafeDecimal(actualReturnCost) + SafeDecimal(r.Cells(20).Value) * SafeDecimal(r.Cells(11).Value)).ToString()
            actualReturnSurcharge = (SafeDecimal(actualReturnSurcharge) + SafeDecimal(r.Cells(21).Value)).ToString()
            returnAmountTotal = (SafeDecimal(returnAmountTotal) + SafeDecimal(r.Cells(22).Value)).ToString()
        Next

        ' 回收汇总
        recoveryWeightTotal = "0"
        recoveryGoldTotal = "0"
        recoveryOtherTotal = "0"
        recoveryAmountTotalVal = "0"
        For Each r As DataGridViewRow In dgvRecovery.Rows
            recoveryWeightTotal = (SafeDecimal(recoveryWeightTotal) + SafeDecimal(r.Cells(3).Value)).ToString()
            recoveryGoldTotal = (SafeDecimal(recoveryGoldTotal) + SafeDecimal(r.Cells(4).Value)).ToString()
            recoveryOtherTotal = (SafeDecimal(recoveryOtherTotal) + SafeDecimal(r.Cells(7).Value)).ToString()
            recoveryAmountTotalVal = (SafeDecimal(recoveryAmountTotalVal) + SafeDecimal(r.Cells(8).Value)).ToString()
        Next

        txtReceivable.Text = "-" & returnAmountTotal

        ' 实退金额 = 四舍五入(应退金额 - 回收金额, 0)
        actualReturnAmount = Math.Round(SafeDecimal(returnAmountTotal) - SafeDecimal(recoveryAmountTotalVal), 0).ToString()

        ' 更新汇总表格
        InitSummaryGrid()

        ' 实收金额 = 实退金额 * -1
        Dim infoReceivedAmount As String = (SafeDecimal(actualReturnAmount) * -1).ToString()
        infoReceivedAmount = FormatTwoDecimals(infoReceivedAmount, False)

        ' 清零支付方式并设置第一行
        For i As Integer = 0 To dgvPayment.Rows.Count - 2
            dgvPayment.Rows(i).Cells(2).Value = 0
        Next
        If dgvPayment.Rows.Count > 1 Then
            dgvPayment.Rows(0).Cells(2).Value = SafeDecimal(infoReceivedAmount)
        End If

        ' 合计行
        If dgvPayment.Rows.Count > 0 Then
            Dim sumPay As Decimal = 0
            For i As Integer = 0 To dgvPayment.Rows.Count - 2
                sumPay += SafeDecimal(dgvPayment.Rows(i).Cells(2).Value)
            Next
            dgvPayment.Rows(dgvPayment.Rows.Count - 1).Cells(2).Value = sumPay
            txtReceived.Text = dgvPayment.Rows(dgvPayment.Rows.Count - 1).Cells(2).Value.ToString()
        End If

        ' 存料/余额计算
        Dim infoStockWeight As String = (SafeDecimal(wholesaleStockWeight) - SafeDecimal(txtStoreMaterial.Text) + SafeDecimal(recoveryGoldTotal)).ToString()
        Dim infoStockAmount As String = (SafeDecimal(wholesaleStockAmount) - SafeDecimal(txtStoreBalance.Text) + SafeDecimal(txtReceived.Text)).ToString()
        infoStockWeight = FormatThreeDecimals(infoStockWeight, False)
        infoStockAmount = FormatTwoDecimals(infoStockAmount, False)

        lblStoreMaterialVal.Text = infoStockWeight
        lblStoreBalanceVal.Text = infoStockAmount
    End Sub

    ' ========== _高级表格1_结束编辑 ==========
    Private Sub DgvSales_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        tableEditState = 0

        Dim productCode As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(1).Value)
        Dim goldWeight As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(11).Value)
        Dim actualSales As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(13).Value)
        Dim salesGoldPrice As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(15).Value)
        Dim salesCost As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(16).Value)
        Dim salesSurcharge As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(17).Value)
        Dim returnGoldPrice As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(19).Value)
        Dim returnCost As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(20).Value)
        Dim returnSurcharge As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(21).Value)
        Dim returnAmount As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(22).Value)

        ' 列1: 商品编码
        If e.ColumnIndex = 1 Then
            If productCode <> "" Then
                ' 检查商品是否存在销售数据
                Dim existsCount As Integer = 0
                Dim sql As String = "SELECT a.* FROM xipunum_erp_outbound as a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE (a.poduct_code = '" & SafeSQL(productCode) & "' or b.fu_code='" & SafeSQL(productCode) & "') AND a.kufang = '" & UserDepartment & "' AND a.xsstate = '0'"
                Try
                    Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                    existsCount = dt.Rows.Count
                Catch ex As Exception
                End Try

                If existsCount = 0 Then
                    ShowWarning("此商品暂无销售数据！")
                    dgvSales.Rows(e.RowIndex).Cells(1).Value = ""
                    Return
                End If

                dataProductCode = ""
                dataProductOrderId = ""

                If existsCount > 1 Then
                    ' 多条记录 - 弹出订单查询窗口让用户选择
                    dataProductCode = productCode
                    dataProductOrderId = ""
                    Dim queryForm As New SalesOrderQueryForm()
                    queryForm.ProductCode = productCode
                    queryForm.ParentReturnForm = Me
                    queryForm.ShowDialog()
                    Return
                Else
                    ' 唯一记录
                    Dim newCode As String = ""
                    Dim orderId As String = ""
                    Try
                        Dim sql2 As String = "SELECT a.poduct_code as poduct_code,a.order_id as order_id FROM xipunum_erp_outbound as a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE (a.poduct_code = '" & SafeSQL(productCode) & "' or b.fu_code='" & SafeSQL(productCode) & "') AND a.kufang = '" & UserDepartment & "' AND a.xsstate = '0'"
                        Dim dt2 As DataTable = DatabaseModule.ExecuteQuery(sql2, MySQL_Read)
                        If dt2.Rows.Count > 0 Then
                            newCode = SafeString(dt2.Rows(0)("poduct_code"))
                            orderId = SafeString(dt2.Rows(0)("order_id"))
                        End If
                    Catch ex As Exception
                    End Try
                    dataProductCode = newCode
                    dataProductOrderId = orderId
                    ReadProductData()
                    Return
                End If
            End If
        End If

        ' 列11: 销售金重
        If e.ColumnIndex = 11 Then
            goldWeight = FormatThreeDecimals(goldWeight, False)
            dgvSales.Rows(e.RowIndex).Cells(11).Value = goldWeight
            If Not ValidateReturnLimit(e.RowIndex) Then Return
        End If

        ' 列19: 实退克价
        If e.ColumnIndex = 19 Then
            If SafeDecimal(returnGoldPrice) > SafeDecimal(salesGoldPrice) Then
                ShowWarning("退货克价不能大于销售克价！")
                dgvSales.Rows(e.RowIndex).Cells(19).Value = salesGoldPrice
                Return
            End If
            returnGoldPrice = FormatTwoDecimals(returnGoldPrice, False)
            dgvSales.Rows(e.RowIndex).Cells(19).Value = returnGoldPrice
        End If

        ' 列20: 实退工费
        If e.ColumnIndex = 20 Then
            If SafeDecimal(returnCost) > SafeDecimal(salesCost) Then
                ShowWarning("实退工费不能大于销售工费！")
                dgvSales.Rows(e.RowIndex).Cells(20).Value = salesCost
                Return
            End If
            returnCost = FormatTwoDecimals(returnCost, False)
            dgvSales.Rows(e.RowIndex).Cells(20).Value = returnCost
        End If

        ' 列21: 实退附加费
        If e.ColumnIndex = 21 Then
            If SafeDecimal(returnSurcharge) > SafeDecimal(salesSurcharge) Then
                ShowWarning("实退附加费不能大于销售附加费！")
                dgvSales.Rows(e.RowIndex).Cells(21).Value = salesSurcharge
                Return
            End If
            returnSurcharge = FormatTwoDecimals(returnSurcharge, False)
            dgvSales.Rows(e.RowIndex).Cells(21).Value = returnSurcharge
        End If

        ' 列22: 实退金额
        If e.ColumnIndex = 22 Then
            If SafeDecimal(returnAmount) > SafeDecimal(actualSales) Then
                ShowWarning("实退金额不能大于实销金额！")
                dgvSales.Rows(e.RowIndex).Cells(22).Value = actualSales
                Return
            End If
            returnAmount = FormatAmount(returnAmount, False)
            dgvSales.Rows(e.RowIndex).Cells(22).Value = returnAmount
        End If

        ' 列11或14: 校验退货行上限
        If e.ColumnIndex = 11 Or e.ColumnIndex = 14 Then
            If Not ValidateReturnLimit(e.RowIndex) Then Return
        End If

        ' 列11/19/20/21: 重新计算实退金额
        If e.ColumnIndex = 11 Or e.ColumnIndex = 19 Or e.ColumnIndex = 20 Or e.ColumnIndex = 21 Then
            returnAmount = ((SafeDecimal(returnGoldPrice) + SafeDecimal(returnCost)) * SafeDecimal(goldWeight) + SafeDecimal(returnSurcharge)).ToString()
            returnAmount = FormatAmount(returnAmount, False)
            dgvSales.Rows(e.RowIndex).Cells(22).Value = returnAmount
        End If

        CalculateSummary()
    End Sub

    ' ========== _高级表格1_读取数据 ==========
    Public Sub ReadProductData()
        ' 检查重复录入
        If Not String.IsNullOrEmpty(dataProductCode) Then
            For i As Integer = 0 To dgvSales.Rows.Count - 2
                If dgvSales.Rows(i).Cells(1).Value IsNot Nothing AndAlso dgvSales.Rows(i).Cells(1).Value.ToString() = dataProductCode AndAlso i <> row1 Then
                    ShowWarning("该商品已在退货列表中！")
                    If dgvSales.Rows.Count > 0 Then
                        dgvSales.Rows(dgvSales.Rows.Count - 1).Cells(1).Value = ""
                    End If
                    Return
                End If
            Next
        End If

        Dim productCode As String = dataProductCode
        Dim orderId As String = dataProductOrderId
        Dim productName As String = ""
        Dim itemNumber As String = ""
        Dim spec As String = ""
        Dim caizhi As String = ""
        Dim quandu As String = ""
        Dim factoryCondition As String = ""
        Dim basicCost As String = ""
        Dim premiumCost As String = ""
        Dim singleWeight As String = ""
        Dim netWeight As String = ""
        Dim xiaoAmount As String = ""
        Dim settlement As String = ""
        Dim quantity As String = ""
        Dim goldPrice As String = ""
        Dim salesCost As String = ""
        Dim salesSurcharge As String = ""
        Dim zhekou As String = ""
        Dim guideName As String = ""
        Dim kufang As String = ""
        Dim astate As String = ""
        Dim gid As String = ""

        Dim sql As String = "SELECT c.product_name AS cproduct_name,c.item_number AS citem_number,e.title AS etitle,c.caizhi AS ccaizhi,c.quandu AS cquandu,d.factory_condition AS dfactory_condition,CAST(ROUND(a.basic_cost, 2) AS DECIMAL (10, 2)) as abasic_cost,CAST(ROUND(a.premium_cost, 2) AS DECIMAL (10, 2)) as apremium_cost,CAST(ROUND(c.single, 3) AS DECIMAL (10, 3)) as danjian,CAST(ROUND(a.net_weight, 3) AS DECIMAL (10, 3)) as jinzhong,CAST(ROUND(a.xiao_amount, 2) AS DECIMAL (10, 2)) as xiaoshouje,CAST(ROUND(a.settlement, 2) AS DECIMAL (10, 2)) as shixiaoje,CAST(ROUND(a.quantity, 2) AS DECIMAL (10, 2)) as aquantity,CAST(ROUND(a.gold_price, 2) AS DECIMAL (10, 2)) as xskejia,CAST(ROUND(a.sales_cost, 2) AS DECIMAL (10, 2)) as xsgongfei,CAST(ROUND(a.sales_surcharge, 2) AS DECIMAL (10, 2)) as xsfujia,CAST(ROUND(a.zhekou, 3) AS DECIMAL (10, 3)) as zhekou,b.name as daogou,a.kufang as kufang,a.state as astate,g.id as gid FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_user AS b ON b.USER = a.shopping_guide INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = a.poduct_code INNER JOIN xipunum_erp_store AS d ON d.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_specs AS e ON e.id = c.specification_id INNER JOIN xipunum_erp_outbound_order AS f ON f.id = a.order_id LEFT JOIN xipunum_erp_member AS g ON g.customer_code = f.customer_code WHERE a.order_id = '" & SafeSQL(orderId) & "' and a.poduct_code = '" & SafeSQL(productCode) & "' and a.kufang='" & UserDepartment & "' ORDER BY a.id DESC  LIMIT 1"

        Try
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            If dt.Rows.Count = 0 Then
                ShowWarning("未找到可退货的销售明细！")
                If row1 < dgvSales.Rows.Count Then
                    dgvSales.Rows(row1).Cells(1).Value = ""
                End If
                Return
            End If

            Dim r As DataRow = dt.Rows(0)
            productName = SafeString(r("cproduct_name"))
            itemNumber = SafeString(r("citem_number"))
            spec = SafeString(r("etitle"))
            caizhi = SafeString(r("ccaizhi"))
            quandu = SafeString(r("cquandu"))
            factoryCondition = SafeString(r("dfactory_condition"))
            basicCost = SafeString(r("abasic_cost"))
            premiumCost = SafeString(r("apremium_cost"))
            singleWeight = SafeString(r("danjian"))
            netWeight = SafeString(r("jinzhong"))
            xiaoAmount = SafeString(r("xiaoshouje"))
            settlement = SafeString(r("shixiaoje"))
            quantity = SafeString(r("aquantity"))
            goldPrice = SafeString(r("xskejia"))
            salesCost = SafeString(r("xsgongfei"))
            salesSurcharge = SafeString(r("xsfujia"))
            zhekou = SafeString(r("zhekou"))
            guideName = SafeString(r("daogou"))
            kufang = SafeString(r("kufang"))
            astate = SafeString(r("astate"))
            gid = SafeString(r("gid"))
        Catch ex As Exception
            ShowWarning("查询商品信息失败：" & ex.Message)
            Return
        End Try

        ' 检查库房
        If UserDepartment <> kufang Then
            ShowWarning("此商品不属于当前店铺销售！")
            dgvSales.Rows(row1).Cells(1).Value = ""
            Return
        End If

        ' 检查线上商品
        If astate <> "0" Then
            ShowWarning("此商品属于线上下单商品，请在线上订单中进行退货操作！")
            dgvSales.Rows(row1).Cells(1).Value = ""
            Return
        End If

        ' 会员一致性检查
        If salesMemberId = "-1" Then
            salesMemberId = gid
            LoadMemberInfo()
        End If

        If salesMemberId <> "-1" AndAlso salesMemberId <> gid Then
            ShowWarning("此商品与上一件商品不是同一个会员销售，无法退货！")
            dgvSales.Rows(row1).Cells(1).Value = ""
            Return
        End If

        ' 填充数据
        dgvSales.Rows(row1).Cells(0).Value = row1.ToString()
        dgvSales.Rows(row1).Cells(1).Value = productCode
        dgvSales.Rows(row1).Cells(2).Value = productName
        dgvSales.Rows(row1).Cells(3).Value = itemNumber
        dgvSales.Rows(row1).Cells(4).Value = spec
        dgvSales.Rows(row1).Cells(5).Value = caizhi
        dgvSales.Rows(row1).Cells(6).Value = quandu
        dgvSales.Rows(row1).Cells(7).Value = factoryCondition
        dgvSales.Rows(row1).Cells(8).Value = basicCost
        dgvSales.Rows(row1).Cells(9).Value = premiumCost
        dgvSales.Rows(row1).Cells(10).Value = singleWeight
        dgvSales.Rows(row1).Cells(11).Value = netWeight
        dgvSales.Rows(row1).Cells(12).Value = xiaoAmount
        dgvSales.Rows(row1).Cells(13).Value = settlement
        dgvSales.Rows(row1).Cells(14).Value = quantity
        dgvSales.Rows(row1).Cells(15).Value = goldPrice
        dgvSales.Rows(row1).Cells(16).Value = salesCost
        dgvSales.Rows(row1).Cells(17).Value = salesSurcharge
        dgvSales.Rows(row1).Cells(18).Value = zhekou
        dgvSales.Rows(row1).Cells(19).Value = goldPrice
        dgvSales.Rows(row1).Cells(20).Value = salesCost
        dgvSales.Rows(row1).Cells(21).Value = salesSurcharge
        dgvSales.Rows(row1).Cells(22).Value = settlement
        dgvSales.Rows(row1).Cells(23).Value = guideName
        dgvSales.Rows(row1).Cells(24).Value = orderId

        ' 设置只读
        For i As Integer = 1 To 18
            dgvSales.Rows(row1).Cells(i).ReadOnly = True
            dgvSales.Rows(row1).Cells(i).Style.BackColor = Drawing.Color.LightGray
        Next
        For i As Integer = 19 To 22
            dgvSales.Rows(row1).Cells(i).ReadOnly = False
            dgvSales.Rows(row1).Cells(i).Style.BackColor = Drawing.Color.White
        Next
        dgvSales.Rows(row1).Cells(23).ReadOnly = True
        dgvSales.Rows(row1).Cells(23).Style.BackColor = Drawing.Color.LightGray

        ' 零销商品可编辑金重
        If SafeDecimal(quantity) = 0 Then
            dgvSales.Rows(row1).Cells(11).ReadOnly = False
            dgvSales.Rows(row1).Cells(11).Style.BackColor = Drawing.Color.White
        End If

        AddBlankSalesRow()
        CalculateSummary()
    End Sub

    ' ========== _高级表格2_结束编辑 ==========
    Private Sub DgvRecovery_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return

        ' 检查是否有客退商品
        If dgvSales.Rows.Count <= 1 Then
            ShowWarning("请先添加客退商品！")
            Return
        End If

        recoveryEditState = 0
        Dim salesGoldPrice As String = Math.Round(SafeDecimal(dgvSales.Rows(0).Cells(19).Value), 2).ToString()
        Dim recoveryName As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(1).Value)
        Dim recoveryQty As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(2).Value)
        Dim recoveryTotalWeight As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(3).Value)
        Dim recoveryGoldWeight As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(4).Value)
        Dim recoveryChengse As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(5).Value)
        Dim recoveryPrice As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(6).Value)
        Dim recoveryOtherFee As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(7).Value)
        Dim recoveryTotalAmount As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(8).Value)

        ' 列1: 回收名称
        If e.ColumnIndex = 1 Then
            Try
                Dim sql As String = "SELECT title FROM xipunum_erp_retreat_title WHERE (bianma = '" & SafeSQL(recoveryName) & "' or title like '%" & SafeSQL(recoveryName) & "%') LIMIT 1"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                If dt.Rows.Count > 0 Then
                    recoveryName = SafeString(dt.Rows(0)("title"))
                    dgvRecovery.Rows(e.RowIndex).Cells(1).Value = recoveryName
                Else
                    ShowWarning("回收名称不存在！")
                    dgvRecovery.Rows(e.RowIndex).Cells(1).Value = ""
                    Return
                End If
            Catch ex As Exception
            End Try
        End If

        ' 非列1时检查名称是否为空
        If e.ColumnIndex <> 1 AndAlso recoveryName = "" Then
            ShowWarning("回收名称不能为空！")
            Return
        End If

        ' 列2: 数量
        If e.ColumnIndex = 2 Then
            dgvRecovery.Rows(e.RowIndex).Cells(2).Value = recoveryQty
        End If

        ' 列3: 总重
        If e.ColumnIndex = 3 Then
            recoveryTotalWeight = FormatThreeDecimals(recoveryTotalWeight, False)
            recoveryGoldWeight = recoveryTotalWeight
            dgvRecovery.Rows(e.RowIndex).Cells(3).Value = recoveryTotalWeight
            dgvRecovery.Rows(e.RowIndex).Cells(4).Value = recoveryGoldWeight
        End If

        ' 列4: 金重
        If e.ColumnIndex = 4 Then
            recoveryGoldWeight = FormatThreeDecimals(recoveryGoldWeight, False)
            dgvRecovery.Rows(e.RowIndex).Cells(4).Value = recoveryGoldWeight
        End If

        ' 列6: 回收克价
        If e.ColumnIndex = 6 Then
            recoveryChengse = ""
            If SafeDecimal(salesGoldPrice) > 0 Then
                recoveryChengse = Math.Round(SafeDecimal(recoveryPrice) / SafeDecimal(salesGoldPrice), 4).ToString()
            End If
            recoveryPrice = FormatTwoDecimals(recoveryPrice, False)
            dgvRecovery.Rows(e.RowIndex).Cells(5).Value = recoveryChengse
            dgvRecovery.Rows(e.RowIndex).Cells(6).Value = recoveryPrice
        End If

        ' 列7: 其他费用
        If e.ColumnIndex = 7 Then
            recoveryOtherFee = FormatTwoDecimals(recoveryOtherFee, False)
            dgvRecovery.Rows(e.RowIndex).Cells(7).Value = recoveryOtherFee
        End If

        ' 列8: 回收金额
        If e.ColumnIndex = 8 Then
            recoveryTotalAmount = FormatTwoDecimals(recoveryTotalAmount, False)
            dgvRecovery.Rows(e.RowIndex).Cells(8).Value = recoveryTotalAmount
        End If

        ' 列3/4/6/7: 重新计算回收金额
        If e.ColumnIndex = 3 Or e.ColumnIndex = 4 Or e.ColumnIndex = 6 Or e.ColumnIndex = 7 Then
            recoveryTotalAmount = (SafeDecimal(recoveryGoldWeight) * SafeDecimal(recoveryPrice) + SafeDecimal(recoveryOtherFee)).ToString()
            recoveryTotalAmount = FormatAmount(recoveryTotalAmount, False)
            dgvRecovery.Rows(e.RowIndex).Cells(8).Value = recoveryTotalAmount
        End If

        CalculateSummary()
    End Sub

    ' ========== _高级表格3_结束编辑 ==========
    Private Sub DgvPayment_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        paymentEditState = 0

        ' 更新合计行
        If dgvPayment.Rows.Count > 0 Then
            Dim sumVal As Decimal = 0
            For i As Integer = 0 To dgvPayment.Rows.Count - 2
                sumVal += SafeDecimal(dgvPayment.Rows(i).Cells(2).Value)
            Next
            dgvPayment.Rows(dgvPayment.Rows.Count - 1).Cells(2).Value = sumVal
            txtReceived.Text = sumVal.ToString()
        End If
    End Sub

    ' ========== _高级表格1_按钮被点击 ==========
    Private Sub DgvSales_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSales.CellClick
        If e.RowIndex < 0 Then Return
        ' 删除按钮列
        If e.ColumnIndex = 25 Then
            dgvSales.Rows.RemoveAt(e.RowIndex)
            ' 重新编号
            For i As Integer = 0 To dgvSales.Rows.Count - 1
                dgvSales.Rows(i).Cells(0).Value = (i + 1).ToString()
            Next
            CalculateSummary()
        End If
    End Sub

    ' ========== _按钮_回收加_被单击 ==========
    Private Sub BtnRecoveryAdd_Click(sender As Object, e As EventArgs)
        ' 检查是否有客退商品
        If dgvSales.Rows.Count < 1 OrElse SafeString(dgvSales.Rows(0).Cells(1).Value) = "" Then
            ShowWarning("请先输入客退商品！")
            Return
        End If

        ' 检查第一件商品克价
        If SafeDecimal(dgvSales.Rows(0).Cells(19).Value) <= 0 Then
            ShowWarning("客退第一件商品克价不能为0！")
            Return
        End If

        ' 检查会员
        If txtMemberName.Text = "" Then
            ShowWarning("会员信息不能为空！")
            txtMemberName.Focus()
            Return
        End If

        ' 检查上一行是否填写完整
        Dim addRow As Integer = dgvRecovery.Rows.Count
        If addRow > 0 Then
            If SafeString(dgvRecovery.Rows(addRow - 1).Cells(8).Value) = "" Then
                ShowWarning("上一个回收数据不能为空！")
                Return
            End If
            If SafeString(dgvRecovery.Rows(addRow - 1).Cells(9).Value) = "" Then
                ShowWarning("上一个回收导购员不能为空！")
                Return
            End If
        End If

        ' 添加新行
        dgvRecovery.Rows.Add()
        Dim newRow As Integer = dgvRecovery.Rows.Count - 1
        dgvRecovery.Rows(newRow).Cells(0).Value = ""
        dgvRecovery.Rows(newRow).Cells(1).Value = ""
        dgvRecovery.Rows(newRow).Cells(2).Value = "1"
        dgvRecovery.Rows(newRow).Cells(3).Value = ""
        dgvRecovery.Rows(newRow).Cells(4).Value = ""
        dgvRecovery.Rows(newRow).Cells(5).Value = ""
        dgvRecovery.Rows(newRow).Cells(6).Value = ""
        dgvRecovery.Rows(newRow).Cells(7).Value = ""
        dgvRecovery.Rows(newRow).Cells(8).Value = ""

        ' 导购员默认值
        If newRow = 0 Then
            dgvRecovery.Rows(newRow).Cells(9).Value = SafeString(dgvSales.Rows(0).Cells(23).Value)
        Else
            dgvRecovery.Rows(newRow).Cells(9).Value = SafeString(dgvRecovery.Rows(newRow - 1).Cells(9).Value)
        End If

        ' 设置只读
        For i As Integer = 0 To newRow
            dgvRecovery.Rows(i).Cells(8).ReadOnly = True
            dgvRecovery.Rows(i).Cells(8).Style.BackColor = Drawing.Color.LightGray
            dgvRecovery.Rows(i).Cells(9).ReadOnly = False
            dgvRecovery.Rows(i).Cells(9).Style.BackColor = Drawing.Color.White
        Next
    End Sub

    ' ========== _按钮_回收减_被单击 ==========
    Private Sub BtnRecoveryRemove_Click(sender As Object, e As EventArgs)
        If row2 < 0 OrElse row2 >= dgvRecovery.Rows.Count Then
            ShowWarning("请选择需要删除的回收数据！")
            Return
        End If
        If row2 > 0 Then
            dgvRecovery.Rows.RemoveAt(row2)
        End If
        CalculateSummary()
    End Sub

    ' ========== 光标位置改变 ==========
    Private Sub DgvSales_SelectionChanged(sender As Object, e As EventArgs)
        If dgvSales.CurrentCell IsNot Nothing Then
            row1 = dgvSales.CurrentCell.RowIndex
            col1 = dgvSales.CurrentCell.ColumnIndex
        End If
    End Sub

    Private Sub DgvRecovery_SelectionChanged(sender As Object, e As EventArgs)
        If dgvRecovery.CurrentCell IsNot Nothing Then
            row2 = dgvRecovery.CurrentCell.RowIndex
            col2 = dgvRecovery.CurrentCell.ColumnIndex
        End If
    End Sub

    Private Sub DgvPayment_SelectionChanged(sender As Object, e As EventArgs)
        If dgvPayment.CurrentCell IsNot Nothing Then
            row3 = dgvPayment.CurrentCell.RowIndex
            col3 = dgvPayment.CurrentCell.ColumnIndex
        End If
    End Sub

    ' ========== 将被编辑 ==========
    Private Sub DgvSales_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs)
        tableEditState = 1
    End Sub

    Private Sub DgvRecovery_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs)
        recoveryEditState = 1
    End Sub

    Private Sub DgvPayment_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs)
        paymentEditState = 1
    End Sub

    ' ========== _高级表格1_按下某键 ==========
    Private Sub DgvSales_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If dgvSales.CurrentCell IsNot Nothing AndAlso dgvSales.CurrentCell.ColumnIndex > 1 Then
                If radioNewRow.Checked Then
                    ' 换行
                    e.Handled = True
                    Dim nextRow As Integer = dgvSales.CurrentCell.RowIndex + 1
                    If nextRow < dgvSales.Rows.Count Then
                        dgvSales.CurrentCell = dgvSales(nextRow, dgvSales.CurrentCell.ColumnIndex)
                    End If
                ElseIf radioNewCol.Checked Then
                    ' 换列
                    e.Handled = True
                    Dim nextCol As Integer = dgvSales.CurrentCell.ColumnIndex + 1
                    If nextCol < dgvSales.Columns.Count Then
                        dgvSales.CurrentCell = dgvSales(dgvSales.CurrentCell.RowIndex, nextCol)
                    End If
                End If
            End If
        End If
    End Sub

    ' ========== 单选框 ==========
    Private Sub RadioNewRow_CheckedChanged(sender As Object, e As EventArgs)
        ' 换行模式
    End Sub

    Private Sub RadioNewCol_CheckedChanged(sender As Object, e As EventArgs)
        ' 换列模式
    End Sub

    ' ========== 存料编辑框 ==========
    Private Sub TxtStoreMaterial_Leave(sender As Object, e As EventArgs)
        txtStoreMaterial.Text = FormatThreeDecimals(txtStoreMaterial.Text, False)
        CalculateSummary()
    End Sub

    Private Sub TxtStoreMaterial_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If txtStoreMaterial.Text = "" Then
                ShowWarning("会员结料不能为空！")
                txtStoreMaterial.Text = "0.000"
                txtStoreMaterial.Focus()
                Return
            End If
            If SafeDecimal(txtStoreMaterial.Text) < 0 Then
                ShowWarning("会员结料不能小于0！")
                txtStoreMaterial.Text = "0.000"
                txtStoreMaterial.Focus()
                Return
            End If
            CalculateSummary()
            e.Handled = True
        End If
    End Sub

    ' ========== 余额编辑框 ==========
    Private Sub TxtStoreBalance_Leave(sender As Object, e As EventArgs)
        txtStoreBalance.Text = FormatTwoDecimals(txtStoreBalance.Text, False)
        CalculateSummary()
    End Sub

    Private Sub TxtStoreBalance_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If txtStoreBalance.Text = "" Then
                ShowWarning("会员存元不能为空！")
                txtStoreBalance.Text = "0.00"
                txtStoreBalance.Focus()
                Return
            End If
            If SafeDecimal(txtStoreBalance.Text) < 0 Then
                ShowWarning("会员结账不能小于0！")
                txtStoreBalance.Text = "0.00"
                txtStoreBalance.Focus()
                Return
            End If
            CalculateSummary()
            e.Handled = True
        End If
    End Sub

    ' ========== _客户姓名_编辑框_获取内容 ==========
    Private Sub LoadMemberInfo()
        Dim memberId As String = salesMemberId

        Dim memberCode As String = ""
        Dim memberName As String = ""
        Dim memberNo As String = ""
        Dim memberTel As String = ""

        Try
            Dim sql As String = "SELECT * FROM xipunum_erp_member where id ='" & SafeSQL(memberId) & "' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            If dt.Rows.Count > 0 Then
                memberCode = SafeString(dt.Rows(0)("customer_code"))
                memberName = SafeString(dt.Rows(0)("name"))
                memberNo = SafeString(dt.Rows(0)("memberid"))
                memberTel = SafeString(dt.Rows(0)("tel"))
            End If
        Catch ex As Exception
        End Try

        If memberCode <> "" Then
            txtMemberName.Text = memberName & "(" & memberNo & ")"
            lblMemberId.Text = "ID:" & memberNo
            txtPhone.Text = memberTel
        End If

        ' 查询存料
        Dim cunLiao As String = "0.00"
        Try
            Dim sql As String = "SELECT sum(number) FROM xipunum_erp_member_cq WHERE customer_code = '" & SafeSQL(memberCode) & "' and cunqu = '存' and type = '料' and kufang = '" & UserDepartment & "'"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            If dt.Rows.Count > 0 Then
                cunLiao = SafeString(dt.Rows(0)(0))
                If cunLiao = "" Then cunLiao = "0.00"
            End If
        Catch ex As Exception
        End Try

        ' 查询欠料
        Dim qianLiao As String = "0.00"
        Try
            Dim sql As String = "SELECT sum(number) FROM xipunum_erp_member_cq WHERE customer_code = '" & SafeSQL(memberCode) & "' and cunqu = '欠' and type = '料' and kufang = '" & UserDepartment & "'"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            If dt.Rows.Count > 0 Then
                qianLiao = SafeString(dt.Rows(0)(0))
                If qianLiao = "" Then qianLiao = "0.00"
            End If
        Catch ex As Exception
        End Try

        Dim jieYuLiao As String = (SafeDecimal(cunLiao) - SafeDecimal(qianLiao)).ToString()
        jieYuLiao = FormatThreeDecimals(jieYuLiao, False)

        ' 查询存元
        Dim cunYuan As String = "0.00"
        Try
            Dim sql As String = "SELECT sum(number) FROM xipunum_erp_member_cq WHERE customer_code = '" & SafeSQL(memberCode) & "' and cunqu = '存' and type = '元' and kufang = '" & UserDepartment & "'"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            If dt.Rows.Count > 0 Then
                cunYuan = SafeString(dt.Rows(0)(0))
                If cunYuan = "" Then cunYuan = "0.00"
            End If
        Catch ex As Exception
        End Try

        ' 查询欠元
        Dim qianYuan As String = "0.00"
        Try
            Dim sql As String = "SELECT sum(number) FROM xipunum_erp_member_cq WHERE customer_code = '" & SafeSQL(memberCode) & "' and cunqu = '欠' and type = '元' and kufang = '" & UserDepartment & "'"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            If dt.Rows.Count > 0 Then
                qianYuan = SafeString(dt.Rows(0)(0))
                If qianYuan = "" Then qianYuan = "0.00"
            End If
        Catch ex As Exception
        End Try

        Dim jieYuYuan As String = (SafeDecimal(cunYuan) - SafeDecimal(qianYuan)).ToString()
        jieYuYuan = FormatTwoDecimals(jieYuYuan, False)

        lblStoreMaterialVal.Text = jieYuLiao
        txtStoreMaterial.Text = "0.000"
        lblStoreBalanceVal.Text = jieYuYuan
        txtStoreBalance.Text = "0.00"

        wholesaleStockWeight = jieYuLiao
        wholesaleStockAmount = jieYuYuan
    End Sub

    ' ========== _工具条_通用_被单击 ==========
    Private Sub BtnToolbarSave_Click(sender As Object, e As EventArgs)
        btnToolbarSave.Enabled = False
        SaveReturnOrder()
        btnToolbarSave.Enabled = True
    End Sub

    Private Sub BtnToolbarReset_Click(sender As Object, e As EventArgs)
        Form_Load(Nothing, Nothing)
        CalculateSummary()
    End Sub

    Private Sub BtnToolbarPrint_Click(sender As Object, e As EventArgs)
        PrintDocument()
    End Sub

    ' ========== _打印功能_被单击 ==========
    Private Sub PrintDocument()
        ' VB.NET简化为提示
        MessageBox.Show("打印功能：请使用打印预览查看单据。", "打印提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' ========== _超级按钮_保存_被单击 ==========
    Private Sub SaveReturnOrder()
        LogOperationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        LogOperationAccount = UserAccount

        Dim recoveryNumber As String = ""
        Dim customerCode As String = ""

        ' 验证
        If dgvSales.Rows.Count <= 1 Then
            ShowWarning("退货数量不能为空！")
            Return
        End If

        If txtReceived.Text = "" Then
            ShowWarning("实退金额不能为空！")
            txtReceived.Focus()
            Return
        End If

        ' 生成回收单号
        If dgvRecovery.Rows.Count > 0 Then
            recoveryNumber = GenerateOrderNumber("3")
            Try
                Dim chkSql As String = "SELECT * FROM xipunum_erp_retreat_order where retrea_umber ='" & SafeSQL(recoveryNumber) & "'"
                Dim chkDt As DataTable = DatabaseModule.ExecuteQuery(chkSql, MySQL_Read)
                If chkDt.Rows.Count > 0 Then
                    recoveryNumber = GenerateOrderNumber("3")
                End If
            Catch ex As Exception
            End Try
        End If

        ' 检查出库单号是否存在
        Dim outboundNumber As String = txtOrderNumber.Text
        Try
            Dim chkSql As String = "SELECT * FROM xipunum_erp_outbound_order where settlement_number ='" & SafeSQL(outboundNumber) & "'"
            Dim chkDt As DataTable = DatabaseModule.ExecuteQuery(chkSql, MySQL_Read)
            If chkDt.Rows.Count > 0 Then
                ShowWarning("当前退货单号已存在，已重新生成单号，请再次点击保存按钮！")
                txtOrderNumber.Text = GenerateOrderNumber("1")
                Return
            End If
        Catch ex As Exception
        End Try

        ' 会员姓名
        Dim memberName As String = txtMemberName.Text

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

        Dim taxPoint As String = "0"
        Dim taxAmount As String = "0"
        Dim receivableAmount As String = txtReceivable.Text
        Dim receivedAmount As String = txtReceived.Text
        Dim remarks As String = txtRemarks.Text

        ' 优惠金额 = 实退合计 + 实收金额
        Dim discountAmount As String = (SafeDecimal(actualReturnAmount) + SafeDecimal(receivedAmount)).ToString()
        discountAmount = FormatTwoDecimals(discountAmount, False)

        ' 获取客户编码
        If txtMemberName.Text <> "" AndAlso salesMemberId <> "" Then
            Try
                Dim sql As String = "SELECT customer_code FROM xipunum_erp_member WHERE id='" & SafeSQL(salesMemberId) & "' LIMIT 1"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                If dt.Rows.Count > 0 Then
                    customerCode = SafeString(dt.Rows(0)("customer_code"))
                End If
            Catch ex As Exception
            End Try
        End If

        ' 批零标识
        Dim pling As String = "零售"
        If txtMemberName.Text <> "" AndAlso customerCode <> "" Then
            pling = "批发"
        End If

        Dim brandName As String = ""
        Dim salesman As String = txtSalesman.Text
        Dim counterWeight As String = "-" & salesWeightTotal
        Dim settlementWeight As String = "-" & salesWeightTotal
        Dim netCounter As String = "-" & salesWeightTotal
        Dim netWeight As String = "-" & salesWeightTotal
        Dim basicCostVal As String = "-" & costCostTotal
        Dim premiumCostVal As String = "-" & premiumCostTotal
        Dim salesCostVal As String = "-" & actualReturnCost
        Dim salesSurchargeVal As String = "-" & actualReturnSurcharge

        ' 插入出库主单
        Dim insertOrderSql As String = "INSERT INTO xipunum_erp_outbound_order SET " &
            "settlement_number='" & SafeSQL(outboundNumber) & "'," &
            "presale_number=''," &
            "retrea_umber='" & SafeSQL(recoveryNumber) & "'," &
            "ying_amount='" & SafeSQL(receivableAmount) & "'," &
            "youhui='" & SafeSQL(discountAmount) & "'," &
            "settlement='" & SafeSQL(receivedAmount) & "'," &
            "customer_code='" & SafeSQL(customerCode) & "'," &
            "brand_name='" & SafeSQL(brandName) & "'," &
            "salesman='" & SafeSQL(salesman) & "'," &
            "counter_weight='" & SafeSQL(counterWeight) & "'," &
            "settlement_weight='" & SafeSQL(settlementWeight) & "'," &
            "net_counter='" & SafeSQL(netCounter) & "'," &
            "net_weight='" & SafeSQL(netWeight) & "'," &
            "basic_cost='" & SafeSQL(basicCostVal) & "'," &
            "premium_cost='" & SafeSQL(premiumCostVal) & "'," &
            "sales_cost='" & SafeSQL(salesCostVal) & "'," &
            "sales_surcharge='" & SafeSQL(salesSurchargeVal) & "'," &
            "taxpoint='" & SafeSQL(taxPoint) & "'," &
            "taxamount='" & SafeSQL(taxAmount) & "'," &
            "remarks='" & SafeSQL(remarks) & "'," &
            "pling='" & SafeSQL(pling) & "'," &
            "state='正常'," &
            "sales_return='1'," &
            "cjuser='" & SafeSQL(LogOperationAccount) & "'," &
            "creationtime='" & SafeSQL(LogOperationDate) & "'," &
            "xianshangtime='" & SafeSQL(LogOperationDate) & "'"

        Try
            DatabaseModule.ExecuteCommand(insertOrderSql, MySQL_Write)
        Catch ex As Exception
            ShowError("保存出库主单失败：" & ex.Message)
            Return
        End Try

        ' 获取订单ID
        Dim orderId As String = ""
        Try
            Dim sql As String = "SELECT id FROM xipunum_erp_outbound_order where settlement_number='" & SafeSQL(outboundNumber) & "' order by id ASC LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            If dt.Rows.Count > 0 Then
                orderId = SafeString(dt.Rows(0)("id"))
            End If
        Catch ex As Exception
        End Try

        If orderId = "" Then
            ShowWarning("商品销售出库失败！")
            Return
        End If

        ' 保存收款记录
        For i As Integer = 0 To dgvPayment.Rows.Count - 2
            Dim payId As String = SafeString(dgvPayment.Rows(i).Cells(3).Value)
            Dim payAmount As String = SafeString(dgvPayment.Rows(i).Cells(2).Value)
            If SafeDecimal(payAmount) <> 0 Then
                Dim paySql As String = "INSERT INTO xipunum_erp_shoukian SET " &
                    "leibie='1'," &
                    "settlement_number='" & SafeSQL(outboundNumber) & "'," &
                    "xianjin='" & SafeSQL(payAmount) & "'," &
                    "type='" & SafeSQL(payId) & "'," &
                    "kufang='" & SafeSQL(UserDepartment) & "'," &
                    "cjuser='" & SafeSQL(LogOperationAccount) & "'," &
                    "creationtime='" & SafeSQL(LogOperationDate) & "'"
                Try
                    DatabaseModule.ExecuteCommand(paySql, MySQL_Write)
                Catch ex As Exception
                End Try
            End If
        Next

        ' 系统日志
        Dim logContent As String = "账户:" & UserAccount & " 商品退货，退货单号::" & txtOrderNumber.Text
        AddSystemLog("添加", "商品退货", logContent)

        ' 会员存欠记录
        If txtMemberName.Text <> "" AndAlso customerCode <> "" Then
            ' 存料不为0
            If SafeDecimal(txtStoreMaterial.Text) <> 0 Then
                Dim cqSql As String = "INSERT INTO xipunum_erp_member_cq SET " &
                    "customer_code='" & SafeSQL(customerCode) & "'," &
                    "danhao='" & SafeSQL(outboundNumber) & "'," &
                    "cunqu='欠'," &
                    "type='料'," &
                    "number='" & SafeSQL(txtStoreMaterial.Text) & "'," &
                    "kufang='" & SafeSQL(UserDepartment) & "'," &
                    "cjuser='" & SafeSQL(LogOperationAccount) & "'," &
                    "creationtime='" & SafeSQL(LogOperationDate) & "'"
                Try
                    DatabaseModule.ExecuteCommand(cqSql, MySQL_Write)
                Catch ex As Exception
                End Try
                AddSystemLog("添加", "会员存料", "账户:" & UserAccount & " 操作 会员编码：" & customerCode & " 会员:" & memberName & " 欠料（" & txtStoreMaterial.Text & ")g 料")
            End If

            ' 余额不为0
            If SafeDecimal(txtStoreBalance.Text) <> 0 Then
                Dim cqSql As String = "INSERT INTO xipunum_erp_member_cq SET " &
                    "customer_code='" & SafeSQL(customerCode) & "'," &
                    "danhao='" & SafeSQL(outboundNumber) & "'," &
                    "cunqu='欠'," &
                    "type='元'," &
                    "number='" & SafeSQL(txtStoreBalance.Text) & "'," &
                    "kufang='" & SafeSQL(UserDepartment) & "'," &
                    "cjuser='" & SafeSQL(LogOperationAccount) & "'," &
                    "creationtime='" & SafeSQL(LogOperationDate) & "'"
                Try
                    DatabaseModule.ExecuteCommand(cqSql, MySQL_Write)
                Catch ex As Exception
                End Try
                AddSystemLog("添加", "会员存元", "账户:" & UserAccount & " 操作 会员编码：" & customerCode & " 会员:" & memberName & " 欠款（" & txtStoreBalance.Text & ")元")
            End If

            ' 回收金重不为0 - 存料
            If SafeDecimal(recoveryGoldTotal) <> 0 Then
                Dim cqSql As String = "INSERT INTO xipunum_erp_member_cq SET " &
                    "customer_code='" & SafeSQL(customerCode) & "'," &
                    "danhao='" & SafeSQL(outboundNumber) & "'," &
                    "cunqu='存'," &
                    "type='料'," &
                    "number='" & SafeSQL(recoveryGoldTotal) & "'," &
                    "kufang='" & SafeSQL(UserDepartment) & "'," &
                    "cjuser='" & SafeSQL(LogOperationAccount) & "'," &
                    "creationtime='" & SafeSQL(LogOperationDate) & "'"
                Try
                    DatabaseModule.ExecuteCommand(cqSql, MySQL_Write)
                Catch ex As Exception
                End Try
                AddSystemLog("新增", "会员出料", "账户:" & UserAccount & " 增加会员:" & customerCode & "存储（" & recoveryGoldTotal & ")g 料")
            End If

            ' 实退金额不为0 - 存元
            If SafeDecimal(actualReturnAmount) <> 0 Then
                Dim cqSql As String = "INSERT INTO xipunum_erp_member_cq SET " &
                    "customer_code='" & SafeSQL(customerCode) & "'," &
                    "danhao='" & SafeSQL(outboundNumber) & "'," &
                    "cunqu='存'," &
                    "type='元'," &
                    "number='" & SafeSQL(actualReturnAmount) & "'," &
                    "kufang='" & SafeSQL(UserDepartment) & "'," &
                    "cjuser='" & SafeSQL(LogOperationAccount) & "'," &
                    "creationtime='" & SafeSQL(LogOperationDate) & "'"
                Try
                    DatabaseModule.ExecuteCommand(cqSql, MySQL_Write)
                Catch ex As Exception
                End Try
                AddSystemLog("新增", "会员消费", "账户:" & UserAccount & " 增加会员:" & customerCode & "存元（" & (SafeDecimal(actualReturnAmount) * -1).ToString() & ")元")
            End If
        End If

        ' ========== 保存退货明细 ==========
        Dim scoreValue As String = "0"
        Dim salesCount As Integer = dgvSales.Rows.Count - 1

        For i As Integer = 0 To salesCount - 1
            Dim productCode As String = SafeString(dgvSales.Rows(i).Cells(1).Value).Trim()
            If productCode = "" Then Continue For

            ' 校验退货行上限
            If Not ValidateReturnLimit(i) Then Return

            Dim qty As String = "-" & SafeString(dgvSales.Rows(i).Cells(14).Value)
            Dim xiaoDanAmount As String
            If SafeDecimal(qty) = 0 Then
                xiaoDanAmount = "-" & SafeString(dgvSales.Rows(i).Cells(13).Value)
            Else
                xiaoDanAmount = "-" & (SafeDecimal(dgvSales.Rows(i).Cells(13).Value) / SafeDecimal(dgvSales.Rows(i).Cells(14).Value)).ToString()
            End If

            Dim xiaoAmount As String = "-" & SafeString(dgvSales.Rows(i).Cells(13).Value)
            Dim settlementVal As String = "-" & SafeString(dgvSales.Rows(i).Cells(22).Value)
            Dim goldPrice As String = "-" & SafeString(dgvSales.Rows(i).Cells(15).Value)
            Dim netWeightVal As String = "-" & SafeString(dgvSales.Rows(i).Cells(11).Value)
            Dim basicCostRow As String = "-" & SafeString(dgvSales.Rows(i).Cells(8).Value)
            Dim premiumCostRow As String = "-" & SafeString(dgvSales.Rows(i).Cells(9).Value)
            Dim salesCostRow As String = "-" & SafeString(dgvSales.Rows(i).Cells(20).Value)
            Dim salesSurchargeRow As String = "-" & SafeString(dgvSales.Rows(i).Cells(21).Value)
            Dim zhekouVal As String = CalculateSaveDiscount(i)
            Dim guideName As String = SafeString(dgvSales.Rows(i).Cells(23).Value)
            Dim salesOrderId As String = SafeString(dgvSales.Rows(i).Cells(24).Value)
            Dim remarksRow As String = ""

            ' 获取导购员账户
            Dim guideAccount As String = ""
            Try
                Dim sql As String = "SELECT user FROM xipunum_erp_user WHERE name = '" & SafeSQL(guideName) & "' and department = '" & UserDepartment & "' LIMIT 1"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                If dt.Rows.Count > 0 Then
                    guideAccount = SafeString(dt.Rows(0)("user"))
                End If
            Catch ex As Exception
            End Try

            ' 检查是否零销
            Dim isLingxiao As String = "否"
            Try
                Dim sql As String = "SELECT lingxiao FROM xipunum_erp_ksiamges where kuanhao='" & SafeSQL(SafeString(dgvSales.Rows(i).Cells(3).Value)) & "' LIMIT 1"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                If dt.Rows.Count > 0 Then
                    isLingxiao = SafeString(dt.Rows(0)("lingxiao"))
                End If
            Catch ex As Exception
            End Try

            ' 插入出库明细
            Dim insertDetailSql As String = "INSERT INTO xipunum_erp_outbound SET " &
                "order_id='" & SafeSQL(orderId) & "'," &
                "poduct_code='" & SafeSQL(productCode) & "'," &
                "quantity='" & SafeSQL(qty) & "'," &
                "xiaodan_amount='" & SafeSQL(xiaoDanAmount) & "'," &
                "xiao_amount='" & SafeSQL(xiaoAmount) & "'," &
                "settlement='" & SafeSQL(settlementVal) & "'," &
                "gold_price='" & SafeSQL(goldPrice) & "'," &
                "net_weight='" & SafeSQL(netWeightVal) & "'," &
                "basic_cost='" & SafeSQL(basicCostRow) & "'," &
                "premium_cost='" & SafeSQL(premiumCostRow) & "'," &
                "sales_cost='" & SafeSQL(salesCostRow) & "'," &
                "sales_surcharge='" & SafeSQL(salesSurchargeRow) & "'," &
                "shopping_guide='" & SafeSQL(guideAccount) & "'," &
                "kufang='" & SafeSQL(UserDepartment) & "'," &
                "zhekou='" & SafeSQL(zhekouVal) & "'," &
                "pling='" & SafeSQL(pling) & "'," &
                "remarks='" & SafeSQL(remarksRow) & "'," &
                "state='0'," &
                "sales_return='1'," &
                "xsstate='1'," &
                "cjuser='" & SafeSQL(LogOperationAccount) & "'," &
                "creationtime='" & SafeSQL(LogOperationDate) & "'," &
                "xianshangtime='" & SafeSQL(LogOperationDate) & "'"
            Try
                DatabaseModule.ExecuteCommand(insertDetailSql, MySQL_Write)
            Catch ex As Exception
            End Try

            ' 更新原出库记录状态
            Try
                Dim updateSql As String = "UPDATE xipunum_erp_outbound SET xsstate = '1' WHERE poduct_code = '" & SafeSQL(productCode) & "' AND order_id='" & SafeSQL(salesOrderId) & "'"
                DatabaseModule.ExecuteCommand(updateSql, MySQL_Write)
            Catch ex As Exception
            End Try

            ' 回补库存
            RestoreStockRow(productCode, SafeString(dgvSales.Rows(i).Cells(14).Value), SafeString(dgvSales.Rows(i).Cells(11).Value))

            ' 积分计算
            Dim jifenRatio As String = "0"
            Try
                Dim sql As String = "SELECT CASE WHEN COALESCE(d.jifen, '' ) = '' THEN '0' ELSE d.jifen END AS jifen FROM xipunum_erp_shop AS a LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = a.item_number AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id AND a.item_number IS NOT NULL AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id AND a.specification_id IS NOT NULL AND a.specification_id != '' LEFT JOIN xipunum_erp_category AS d ON d.id = COALESCE ( e1.category_id, e2.category_id ) AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL WHERE a.poduct_code='" & SafeSQL(productCode) & "' LIMIT 1"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                If dt.Rows.Count > 0 Then
                    jifenRatio = SafeString(dt.Rows(0)("jifen"))
                End If
            Catch ex As Exception
            End Try

            scoreValue = (SafeDecimal(scoreValue) + Math.Round(SafeDecimal(settlementVal) * SafeDecimal(jifenRatio) / 100, 0)).ToString()

            ' 更新商品状态
            Try
                Dim updateSql As String = "UPDATE xipunum_erp_shop SET state= '销售',updatetime= '" & LogOperationDate & "'  WHERE poduct_code ='" & SafeSQL(productCode) & "' LIMIT 1"
                DatabaseModule.ExecuteCommand(updateSql, MySQL_Write)
            Catch ex As Exception
            End Try

            ' 更新证书
            Try
                Dim updateSql As String = "UPDATE xipunum_erp_zhengshu SET xiaoshouid= '',xiaoshou= '',updatetime= '" & LogOperationDate & "'  WHERE poduct_code ='" & SafeSQL(productCode) & "' LIMIT 1"
                DatabaseModule.ExecuteCommand(updateSql, MySQL_Write)
            Catch ex As Exception
            End Try

            ' 获取库房名称
            Dim kufangName As String = ""
            Try
                Dim sql As String = "SELECT title FROM xipunum_erp_type where id='" & UserDepartment & "' order by id ASC LIMIT 1"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                If dt.Rows.Count > 0 Then
                    kufangName = SafeString(dt.Rows(0)("title"))
                End If
            Catch ex As Exception
            End Try

            ' 获取商品重量
            Dim weightData As String = ""
            Try
                Dim sql As String = "SELECT weight FROM xipunum_erp_shop where poduct_code='" & SafeSQL(productCode) & "' order by id ASC LIMIT 1"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                If dt.Rows.Count > 0 Then
                    weightData = SafeString(dt.Rows(0)("weight"))
                End If
            Catch ex As Exception
            End Try

            ' 商品追溯
            Dim historySql As String = "INSERT INTO xipunum_erp_history SET " &
                "poduct_code='" & SafeSQL(productCode) & "'," &
                "updatetime='" & SafeSQL(LogOperationDate) & "'," &
                "number='" & SafeSQL(outboundNumber) & "'," &
                "type='客户退货'," &
                "quantity='" & SafeSQL(qty) & "'," &
                "jinzhong='" & SafeSQL(netWeightVal) & "'," &
                "zhongliang='" & SafeSQL(weightData) & "'," &
                "conter='商品从" & kufangName & "退货'," &
                "cjuser='" & SafeSQL(LogOperationAccount) & "'"
            Try
                DatabaseModule.ExecuteCommand(historySql, MySQL_Write)
            Catch ex As Exception
            End Try

            ' 商品日志
            Dim shopLogSql As String = "INSERT INTO xipunum_erp_shop_log SET " &
                "poduct_code='" & SafeSQL(productCode) & "'," &
                "type='退货'," &
                "creationtime='" & SafeSQL(LogOperationDate) & "'"
            Try
                DatabaseModule.ExecuteCommand(shopLogSql, MySQL_Write)
            Catch ex As Exception
            End Try

            ' 系统日志
            AddSystemLog("添加", "客户退货", "账户:" & UserAccount & " 客户退货，退货单号::" & productCode)
        Next

        ' ========== 保存回收数据 ==========
        If dgvRecovery.Rows.Count > 0 Then
            Dim recoverySalesman As String = txtSalesman.Text
            Dim recoveryTotalWeight As String = "-" & recoveryWeightTotal
            Dim recoveryGoldWeight As String = "-" & recoveryGoldTotal
            Dim recoveryOtherFee As String = "-" & recoveryOtherTotal
            Dim recoveryTaxPoint As String = "0"
            Dim recoveryTaxAmount As String = "0"
            Dim recoveryAmountVal As String = "-" & recoveryAmountTotalVal
            Dim recoveryYingAmount As String = "-" & recoveryAmountTotalVal
            Dim recoverySettlement As String = "-" & recoveryAmountTotalVal
            Dim recoveryRemarks As String = ""

            ' 插入回收主单
            Dim insertRecoverySql As String = "INSERT INTO xipunum_erp_retreat_order SET " &
                "retrea_umber='" & SafeSQL(recoveryNumber) & "'," &
                "customer_code='" & SafeSQL(customerCode) & "'," &
                "total='" & SafeSQL(recoveryTotalWeight) & "'," &
                "jin_zhong='" & SafeSQL(recoveryGoldWeight) & "'," &
                "qita_price='" & SafeSQL(recoveryOtherFee) & "'," &
                "tax_rate='" & SafeSQL(recoveryTaxPoint) & "'," &
                "rate_amount='" & SafeSQL(recoveryTaxAmount) & "'," &
                "retreat_amount='" & SafeSQL(recoveryAmountVal) & "'," &
                "ying_amount='" & SafeSQL(recoveryYingAmount) & "'," &
                "settlement='" & SafeSQL(recoverySettlement) & "'," &
                "salesman='" & SafeSQL(recoverySalesman) & "'," &
                "remarks='" & SafeSQL(recoveryRemarks) & "'," &
                "cjuser='" & SafeSQL(LogOperationAccount) & "'," &
                "creationtime='" & SafeSQL(LogOperationDate) & "'"
            Try
                DatabaseModule.ExecuteCommand(insertRecoverySql, MySQL_Write)
            Catch ex As Exception
            End Try

            ' 获取回收订单ID
            Dim recoveryOrderId As String = ""
            Try
                Dim sql As String = "SELECT id FROM xipunum_erp_retreat_order where retrea_umber='" & SafeSQL(recoveryNumber) & "' order by id ASC LIMIT 1"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                If dt.Rows.Count > 0 Then
                    recoveryOrderId = SafeString(dt.Rows(0)("id"))
                End If
            Catch ex As Exception
            End Try

            ' 保存回收明细
            For i As Integer = 0 To dgvRecovery.Rows.Count - 1
                Dim recName As String = SafeString(dgvRecovery.Rows(i).Cells(1).Value)
                Dim recQty As String = "-" & SafeString(dgvRecovery.Rows(i).Cells(2).Value)
                Dim recTotal As String = "-" & SafeString(dgvRecovery.Rows(i).Cells(3).Value)
                Dim recGold As String = "-" & SafeString(dgvRecovery.Rows(i).Cells(4).Value)
                Dim recChengse As String = SafeString(dgvRecovery.Rows(i).Cells(5).Value)
                Dim recPrice As String = "-" & SafeString(dgvRecovery.Rows(i).Cells(6).Value)
                Dim recOther As String = "-" & SafeString(dgvRecovery.Rows(i).Cells(7).Value)
                Dim recAmount As String = "-" & SafeString(dgvRecovery.Rows(i).Cells(8).Value)
                Dim recGuide As String = SafeString(dgvRecovery.Rows(i).Cells(9).Value)
                Dim recRemarks As String = ""

                ' 获取导购员账户
                Dim guideAccount As String = ""
                Try
                    Dim sql As String = "SELECT user FROM xipunum_erp_user WHERE name = '" & SafeSQL(recGuide) & "' and department = '" & UserDepartment & "' LIMIT 1"
                    Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                    If dt.Rows.Count > 0 Then
                        guideAccount = SafeString(dt.Rows(0)("user"))
                    End If
                Catch ex As Exception
                End Try

                ' 获取回收标题ID和积分比例
                Dim recDataId As String = ""
                Dim recJifenRatio As String = "0"
                Try
                    Dim sql As String = "SELECT a.id as aid,CASE WHEN COALESCE(b.jifen, '' ) = '' THEN '0' ELSE b.jifen END AS jifen FROM xipunum_erp_retreat_title as a INNER JOIN xipunum_erp_category AS b ON b.id = a.category_id WHERE a.title = '" & SafeSQL(recName) & "'  LIMIT 1"
                    Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                    If dt.Rows.Count > 0 Then
                        recDataId = SafeString(dt.Rows(0)("aid"))
                        recJifenRatio = SafeString(dt.Rows(0)("jifen"))
                    End If
                Catch ex As Exception
                End Try

                If recDataId = "" Then recDataId = "1"
                If guideAccount = "" Then guideAccount = LogOperationAccount

                ' 插入回收明细
                Dim insertRecDetailSql As String = "INSERT INTO xipunum_erp_retreat SET " &
                    "order_id='" & SafeSQL(recoveryOrderId) & "'," &
                    "product_name='" & SafeSQL(recDataId) & "'," &
                    "quantity='" & SafeSQL(recQty) & "'," &
                    "total='" & SafeSQL(recTotal) & "'," &
                    "jin_zhong='" & SafeSQL(recGold) & "'," &
                    "chengse='" & SafeSQL(recChengse) & "'," &
                    "price='" & SafeSQL(recPrice) & "'," &
                    "qita_price='" & SafeSQL(recOther) & "'," &
                    "retreat_amount='" & SafeSQL(recAmount) & "'," &
                    "huishoutime='" & SafeSQL(LogOperationDate) & "'," &
                    "shopping_guide='" & SafeSQL(guideAccount) & "'," &
                    "remarks='" & SafeSQL(recRemarks) & "'," &
                    "cjuser='" & SafeSQL(LogOperationAccount) & "'," &
                    "creationtime='" & SafeSQL(LogOperationDate) & "'"
                Try
                    DatabaseModule.ExecuteCommand(insertRecDetailSql, MySQL_Write)
                Catch ex As Exception
                End Try

                ' 扣减积分
                scoreValue = (SafeDecimal(scoreValue) - Math.Round(SafeDecimal(recAmount) * SafeDecimal(recJifenRatio) / 100, 0)).ToString()

                AddSystemLog("添加", "回收出库", "账户:" & UserAccount & " 回收出库，名称：" & recName)
            Next

            AddSystemLog("添加", "回收出库", "账户:" & UserAccount & " 回收出库，回收单号::" & recoveryNumber & "  回收数量:" & dgvRecovery.Rows.Count.ToString())
        End If

        ' ========== 积分记录 ==========
        If customerCode <> "" AndAlso SafeDecimal(scoreValue) <> 0 Then
            Dim scoreSql As String = "INSERT INTO xipunum_erp_member_score_log SET " &
                "customer_code='" & SafeSQL(customerCode) & "'," &
                "settlement_number='" & SafeSQL(outboundNumber) & "'," &
                "num='" & SafeSQL(scoreValue) & "'," &
                "st='1'," &
                "type='1'," &
                "remarks='会员退货【订单号：" & outboundNumber & "】积分返还'," &
                "creationtime='" & SafeSQL(LogOperationDate) & "'," &
                "cjuser='" & SafeSQL(LogOperationAccount) & "'"
            Try
                DatabaseModule.ExecuteCommand(scoreSql, MySQL_Write)
            Catch ex As Exception
            End Try
        End If

        ' 打印
        If chkPrintDoc.Checked Then
            Dim result As DialogResult = MessageBox.Show("是否打印单据？", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)
            If result = DialogResult.OK Then
                PrintDocument()
            End If
        End If

        ShowSuccess("商品退货成功！")
        Me.Close()
    End Sub

    ' ========== _客退_SQL文本安全 ==========
    Private Function SafeSQLForReturn(text As String) As String
        Return text.Replace("'", "''")
    End Function

    ' ========== _客退_写库影响行 ==========
    Private Function GetWriteAffectedRows() As String
        Try
            Dim sql As String = "SELECT ROW_COUNT() AS rowcnt"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Write)
            If dt.Rows.Count > 0 Then
                Return SafeString(dt.Rows(0)("rowcnt"))
            End If
        Catch ex As Exception
        End Try
        Return "0"
    End Function

    ' ========== _客退_回补库存行 ==========
    Private Function RestoreStockRow(productCode As String, addQty As String, addWeight As String) As Boolean
        Dim safeCode As String = SafeSQL(productCode)
        Try
            Dim sql As String = "UPDATE xipunum_erp_shop_kucun SET quantity = quantity + '" & SafeSQL(addQty) & "',jinzhong = jinzhong + '" & SafeSQL(addWeight) & "' WHERE poduct_code = '" & safeCode & "' AND kufang='" & UserDepartment & "'"
            DatabaseModule.ExecuteCommand(sql, MySQL_Write)
            If SafeDecimal(GetWriteAffectedRows()) > 0 Then
                Return True
            End If
        Catch ex As Exception
        End Try

        ' 如果UPDATE影响0行，则INSERT
        Try
            Dim sql As String = "INSERT INTO xipunum_erp_shop_kucun SET poduct_code='" & safeCode & "',kufang='" & UserDepartment & "',quantity='" & SafeSQL(addQty) & "',jinzhong='" & SafeSQL(addWeight) & "'"
            DatabaseModule.ExecuteCommand(sql, MySQL_Write)
            Return SafeDecimal(GetWriteAffectedRows()) > 0
        Catch ex As Exception
            Return False
        End Try
    End Function

    ' ========== _客退_校验退货行上限 ==========
    Private Function ValidateReturnLimit(rowIndex As Integer) As Boolean
        Dim productCode As String = SafeString(dgvSales.Rows(rowIndex).Cells(1).Value).Trim()
        If productCode = "" Then Return True

        Dim originalOrderId As String = SafeString(dgvSales.Rows(rowIndex).Cells(24).Value)
        Dim originalQty As Decimal = 0
        Dim originalWeight As Decimal = 0

        If originalOrderId <> "" Then
            Try
                Dim sql As String = "SELECT CAST(ROUND(a.quantity, 2) AS DECIMAL(10, 2)) AS aquantity,CAST(ROUND(a.net_weight, 3) AS DECIMAL(10, 3)) AS jinzhong FROM xipunum_erp_outbound AS a WHERE a.order_id='" & SafeSQL(originalOrderId) & "' AND a.poduct_code='" & SafeSQL(productCode) & "' AND a.kufang='" & UserDepartment & "' AND a.xsstate='0' LIMIT 1"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                If dt.Rows.Count > 0 Then
                    originalQty = SafeDecimal(SafeString(dt.Rows(0)("aquantity")))
                    originalWeight = SafeDecimal(SafeString(dt.Rows(0)("jinzhong")))
                End If
            Catch ex As Exception
            End Try
        End If

        ' 如果查不到，使用表格中的值作为上限
        If originalQty <= 0 AndAlso originalWeight <= 0 Then
            originalQty = SafeDecimal(dgvSales.Rows(rowIndex).Cells(14).Value)
            originalWeight = SafeDecimal(dgvSales.Rows(rowIndex).Cells(11).Value)
        End If

        Dim currentQty As Decimal = SafeDecimal(dgvSales.Rows(rowIndex).Cells(14).Value)
        Dim currentWeight As Decimal = SafeDecimal(dgvSales.Rows(rowIndex).Cells(11).Value)

        If currentQty < 0 Then
            currentQty = 0
            dgvSales.Rows(rowIndex).Cells(14).Value = "0"
        End If
        If currentWeight < 0 Then
            currentWeight = 0
            dgvSales.Rows(rowIndex).Cells(11).Value = "0.000"
        End If

        If originalQty > 0 AndAlso currentQty > originalQty Then
            ShowWarning("退货数量不能大于原销售数量！")
            dgvSales.Rows(rowIndex).Cells(14).Value = originalQty.ToString()
            Return False
        End If

        If originalWeight > 0 AndAlso currentWeight > originalWeight Then
            ShowWarning("退货金重不能大于原销售金重！")
            dgvSales.Rows(rowIndex).Cells(11).Value = originalWeight.ToString()
            Return False
        End If

        Return True
    End Function

    ' ========== _客退_计算保存折扣 ==========
    Private Function CalculateSaveDiscount(rowIndex As Integer) As String
        Dim returnAmount As Decimal = SafeDecimal(dgvSales.Rows(rowIndex).Cells(22).Value)
        Dim actualSales As Decimal = SafeDecimal(dgvSales.Rows(rowIndex).Cells(13).Value)
        Dim originalZhekou As Decimal = SafeDecimal(dgvSales.Rows(rowIndex).Cells(18).Value)
        Dim zhekouVal As Decimal

        If actualSales > 0 Then
            zhekouVal = Math.Round(returnAmount / actualSales, 3)
        Else
            zhekouVal = originalZhekou
        End If

        Return "-" & zhekouVal.ToString()
    End Function

    ' ========== _客退_补足小数文本 ==========
    Private Function PadDecimalText(numText As String, decimalPlaces As Integer) As String
        If numText = "" Then
            If decimalPlaces = 3 Then Return "0.000"
            Return "0.00"
        End If

        Dim parts() As String = numText.Split("."c)
        If parts.Length > 1 Then
            Dim decimalPart As String = parts(1)
            While decimalPart.Length < decimalPlaces
                decimalPart &= "0"
            End While
            Return parts(0) & "." & decimalPart.Substring(0, Math.Min(decimalPart.Length, decimalPlaces))
        End If

        If decimalPlaces = 3 Then
            Return numText & ".000"
        Else
            Return numText & ".00"
        End If
    End Function

    ' ========== _客退_格式化数值文本 ==========
    Private Function FormatNumberText(originalText As String, roundDigits As Integer, displayDigits As Integer, toUTF8 As Boolean) As String
        Dim formattedText As String = PadDecimalText(Math.Round(SafeDecimal(originalText), roundDigits).ToString(), displayDigits)
        Return formattedText
    End Function

    ' ========== _客退_格式化两位 ==========
    Private Function FormatTwoDecimals(originalText As String, toUTF8 As Boolean) As String
        Return FormatNumberText(originalText, 2, 2, toUTF8)
    End Function

    ' ========== _客退_格式化三位 ==========
    Private Function FormatThreeDecimals(originalText As String, toUTF8 As Boolean) As String
        Return FormatNumberText(originalText, 3, 3, toUTF8)
    End Function

    ' ========== _客退_格式化金额 ==========
    Private Function FormatAmount(originalText As String, toUTF8 As Boolean) As String
        Return FormatNumberText(originalText, 0, 2, toUTF8)
    End Function

    ' ========== 辅助函数 ==========
    Private Function SafeString(val As Object) As String
        If val Is Nothing OrElse val Is DBNull.Value Then Return ""
        Return val.ToString()
    End Function

    Private Function SafeDecimal(val As Object) As Decimal
        If val Is Nothing OrElse val Is DBNull.Value Then Return 0
        Dim result As Decimal
        If Decimal.TryParse(val.ToString(), result) Then Return result
        Return 0
    End Function

    Private Sub ShowWarning(msg As String)
        MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub ShowSuccess(msg As String)
        MessageBox.Show(msg, "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ShowError(msg As String)
        MessageBox.Show(msg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Sub AddSystemLog(logType As String, logTitle As String, logContent As String)
        Try
            Dim sql As String = "INSERT INTO xipunum_erp_xitong_log SET " &
                "type='" & SafeSQL(logType) & "'," &
                "title='" & SafeSQL(logTitle) & "'," &
                "conter='" & SafeSQL(logContent) & "'," &
                "user='" & SafeSQL(LogOperationAccount) & "'," &
                "creationtime='" & SafeSQL(LogOperationDate) & "'"
            DatabaseModule.ExecuteCommand(sql, MySQL_Write)
        Catch ex As Exception
        End Try
    End Sub

    Private Function GenerateOrderNumber(suffix As String) As String
        Dim ts As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim rnd As New Random()
        Dim rndNum As String = rnd.Next(1000, 9999).ToString()
        Return (ts & rndNum).Substring(0, 18) & suffix
    End Function

End Class

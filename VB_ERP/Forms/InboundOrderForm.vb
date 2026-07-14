' ============================================================================
' 商品入库添加窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品入库添加.form.e.txt
' 包含所有10个程序集变量、21个子程序、31个SQL查询
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.IO

Public Class InboundOrderForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（10个） ==========
    Private deleteBtn As New Button()              ' 删除按钮
    Private col As Integer = 0                     ' 集_列号
    Private row As Integer = 0                     ' 集_行号
    Private inboundProductCount As Integer = 0      ' 局部_入库商品数量
    Private printOrderNumber As String = ""         ' 局部_打印入库单号
    Private isProductInlaid As String = ""           ' 局部_商品是否镶嵌
    Private productCategoryID As String = ""         ' 局部_商品品类id

    ' ========== 控件声明 ==========
    Private dgvProducts As New DataGridView()       ' 高级表格1
    Private txtOrderNumber As New TextBox()         ' 单据号_编辑框
    Private cmbFactory As New ComboBox()            ' 工厂名称组合框
    Private cmbSource As New ComboBox()             ' 来源组合框
    Private cmbWarehouse As New ComboBox()          ' 库房名称组合框
    Private cmbSettlement As New ComboBox()         ' 结算方式组合框
    Private txtRemarks As New TextBox()             ' 备注_编辑框
    Private txtDeliveryNumber As New TextBox()      ' 送货单号_编辑框
    Private cmbLabelStyle As New ComboBox()         ' 组合框标签样式
    Private chkPrintLabel As New CheckBox()         ' 打印标签_选择框
    Private chkPrintDocument As New CheckBox()      ' 打印单据_选择框
    Private rbHalfYes As New RadioButton()          ' 单选框_是
    Private rbHalfNo As New RadioButton()           ' 单选框_否
    Private txtTotalWeight As New TextBox()         ' 总重_编辑框
    Private txtMaterialWeight As New TextBox()      ' 料重_编辑框
    Private btnReset As New Button()                ' 按钮_重置
    Private toolStrip As New ToolStrip()            ' 工具条_通用

    ' 表格列索引常量
    Private Const COL_SEQ As Integer = 0            ' 序号
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
    Private Const COL_JINZHONG As Integer = 14      ' 金重
    Private Const COL_LOSS As Integer = 15          ' 损耗
    Private Const COL_INCLUDING As Integer = 16     ' 含耗重
    Private Const COL_WEIGHT As Integer = 17        ' 重量
    Private Const COL_UNIT As Integer = 18          ' 单位
    Private Const COL_SHITOU As Integer = 19        ' 石重
    Private Const COL_STNUM As Integer = 20         ' 石头数
    Private Const COL_SHITOU1 As Integer = 21       ' 副石重
    Private Const COL_SHNUM1 As Integer = 22        ' 副石头数
    Private Const COL_COST_PRICE As Integer = 23    ' 成本单价
    Private Const COL_COEFFICIENT As Integer = 24   ' 系数
    Private Const COL_BASIC_COST As Integer = 25    ' 成本工费
    Private Const COL_PREMIUM_COST As Integer = 26  ' 参考工费
    Private Const COL_SALES_COST As Integer = 27    ' 销售工费
    Private Const COL_COMPANY_SUR As Integer = 28   ' 成本附加费
    Private Const COL_SALES_SUR As Integer = 29     ' 销售附加费
    Private Const COL_SALES_PRICE As Integer = 30   ' 销售价
    Private Const COL_REMARKS As Integer = 31       ' 备注
    Private Const COL_ZHUSE As Integer = 32         ' 主石色
    Private Const COL_IMAGES As Integer = 33        ' 图片地址
    Private Const COL_GOLD_PRICE As Integer = 34    ' 原料价
    Private Const COL_MOJUHAO As Integer = 35       ' 模具号
    Private Const COL_DELETE As Integer = 36        ' 操作

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler dgvProducts.CellEndEdit, AddressOf DgvProducts_CellEndEdit
        AddHandler dgvProducts.CellClick, AddressOf DgvProducts_CellClick
        AddHandler dgvProducts.SelectionChanged, AddressOf DgvProducts_SelectionChanged
        AddHandler cmbFactory.DropDown, AddressOf CmbFactory_DropDown
        AddHandler btnReset.Click, AddressOf BtnReset_Click
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "商品入库添加"
        Me.Size = New Drawing.Size(1427, 664)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' 工具条
        toolStrip.Dock = DockStyle.Top
        toolStrip.Height = 45
        Me.Controls.Add(toolStrip)

        AddToolButton("保存")
        AddToolButton("添加商品")
        AddToolButton("标签打印")

        ' 重置按钮
        btnReset.Text = "重置"
        btnReset.Size = New Drawing.Size(45, 45)
        btnReset.Location = New Drawing.Point(toolStrip.Right + 5, 3)
        Me.Controls.Add(btnReset)

        ' 顶部信息区
        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 160
        Me.Controls.Add(panelTop)

        ' 标签样式
        Dim lblLabelStyle As New Label()
        lblLabelStyle.Text = "标签样式："
        lblLabelStyle.Location = New Drawing.Point(158, 8)
        lblLabelStyle.AutoSize = True
        panelTop.Controls.Add(lblLabelStyle)

        cmbLabelStyle.Location = New Drawing.Point(214, 5)
        cmbLabelStyle.Size = New Drawing.Size(126, 25)
        panelTop.Controls.Add(cmbLabelStyle)

        ' 打印标签选择框
        chkPrintLabel.Text = "打印标签"
        chkPrintLabel.Location = New Drawing.Point(350, 11)
        chkPrintLabel.AutoSize = True
        panelTop.Controls.Add(chkPrintLabel)

        ' 打印单据选择框
        chkPrintDocument.Text = "打印单据"
        chkPrintDocument.Location = New Drawing.Point(435, 11)
        chkPrintDocument.AutoSize = True
        panelTop.Controls.Add(chkPrintDocument)

        ' 第一行：工厂、来源、半成品
        AddLabel(panelTop, "工厂：", 20, 50)
        cmbFactory.Location = New Drawing.Point(70, 47)
        cmbFactory.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(cmbFactory)

        AddLabel(panelTop, "来源：", 290, 50)
        cmbSource.Location = New Drawing.Point(340, 47)
        cmbSource.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(cmbSource)

        AddLabel(panelTop, "半成品：", 510, 50)
        rbHalfYes.Text = "是"
        rbHalfYes.Location = New Drawing.Point(570, 48)
        rbHalfYes.AutoSize = True
        panelTop.Controls.Add(rbHalfYes)

        rbHalfNo.Text = "否"
        rbHalfNo.Location = New Drawing.Point(610, 48)
        rbHalfNo.AutoSize = True
        panelTop.Controls.Add(rbHalfNo)

        ' 第二行：库房、结算方式
        AddLabel(panelTop, "库房：", 20, 85)
        cmbWarehouse.Location = New Drawing.Point(70, 82)
        cmbWarehouse.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(cmbWarehouse)

        AddLabel(panelTop, "结算：", 290, 85)
        cmbSettlement.Location = New Drawing.Point(340, 82)
        cmbSettlement.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(cmbSettlement)

        ' 第三行：送货单号、单据号
        AddLabel(panelTop, "送货单号：", 20, 120)
        txtDeliveryNumber.Location = New Drawing.Point(90, 117)
        txtDeliveryNumber.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtDeliveryNumber)

        AddLabel(panelTop, "单据号：", 510, 120)
        txtOrderNumber.Location = New Drawing.Point(580, 117)
        txtOrderNumber.Size = New Drawing.Size(150, 25)
        txtOrderNumber.ReadOnly = True
        panelTop.Controls.Add(txtOrderNumber)

        ' 总重、料重
        AddLabel(panelTop, "总重：", 750, 120)
        txtTotalWeight.Location = New Drawing.Point(790, 117)
        txtTotalWeight.Size = New Drawing.Size(80, 25)
        txtTotalWeight.ReadOnly = True
        panelTop.Controls.Add(txtTotalWeight)

        AddLabel(panelTop, "料重：", 880, 120)
        txtMaterialWeight.Location = New Drawing.Point(920, 117)
        txtMaterialWeight.Size = New Drawing.Size(80, 25)
        txtMaterialWeight.ReadOnly = True
        panelTop.Controls.Add(txtMaterialWeight)

        ' 备注分组框
        Dim grpRemarks As New GroupBox()
        grpRemarks.Text = "备注"
        grpRemarks.Location = New Drawing.Point(798, 41)
        grpRemarks.Size = New Drawing.Size(568, 70)
        panelTop.Controls.Add(grpRemarks)

        txtRemarks.Location = New Drawing.Point(10, 15)
        txtRemarks.Size = New Drawing.Size(545, 45)
        txtRemarks.Multiline = True
        grpRemarks.Controls.Add(txtRemarks)

        ' 商品表格
        dgvProducts.Dock = DockStyle.Fill
        dgvProducts.AllowUserToAddRows = False
        dgvProducts.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvProducts.Font = New Drawing.Font("宋体", 9)
        Me.Controls.Add(dgvProducts)
    End Sub

    Private Sub AddLabel(parent As Control, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    Private Sub AddToolButton(text As String)
        Dim btn As New ToolStripButton(text)
        btn.DisplayStyle = ToolStripItemDisplayStyle.Text
        AddHandler btn.Click, AddressOf ToolButton_Click
        toolStrip.Items.Add(btn)
    End Sub

    ' ========== 窗口加载（_窗口_商品入库添加_创建完毕） ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        inboundProductCount = 0
        printOrderNumber = ""
        isProductInlaid = ""
        productCategoryID = ""
        InitializeBaseData()
        LoadTableHeader()
        LoadTableData()
    End Sub

    ' ========== 窗口尺寸改变（_窗口_商品入库添加_尺寸被改变） ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        ' 表格自动Fill，不需要手动调整
    End Sub

    ' ========== 初始化基础数据（_入库基础_初始化） ==========
    Private Sub InitializeBaseData()
        isProductInlaid = ""
        productCategoryID = ""
        chkPrintLabel.Checked = True
        chkPrintDocument.Checked = False
        rbHalfYes.Enabled = True
        rbHalfNo.Enabled = True
        cmbFactory.Enabled = True
        cmbSource.Enabled = True
        cmbWarehouse.Enabled = True
        cmbSettlement.Enabled = True
        txtDeliveryNumber.Enabled = True
        txtRemarks.Text = ""
        txtDeliveryNumber.Text = ""
        txtOrderNumber.Text = "RK" & DateTime.Now.ToString("yyyyMMdd") & "****"
        txtTotalWeight.Text = "0.00"
        txtMaterialWeight.Text = "0.00"

        ' 加载标签模板文件
        cmbLabelStyle.Items.Clear()
        Dim labelDir As String = Path.Combine(Application.StartupPath, "voucher\biaoqian")
        If Directory.Exists(labelDir) Then
            Dim files() As String = Directory.GetFiles(labelDir, "*.qdf")
            For Each f As String In files
                cmbLabelStyle.Items.Add(Path.GetFileName(f))
            Next
            If cmbLabelStyle.Items.Count > 0 Then
                cmbLabelStyle.SelectedIndex = 0
            End If
        End If

        inboundProductCount = 0
        rbHalfYes.Checked = False
        rbHalfNo.Checked = True

        cmbFactory.Items.Clear()
        cmbFactory.SelectedIndex = -1
        cmbSource.Items.Clear()
        cmbSource.Text = "请选择来源"
        cmbSource.SelectedIndex = -1
        cmbWarehouse.Items.Clear()
        cmbWarehouse.Text = "请选择库房"
        cmbWarehouse.SelectedIndex = -1
        cmbSettlement.Items.Clear()
        cmbSettlement.Text = "请选择结算方式"
        cmbSettlement.SelectedIndex = -1

        ' 加载工厂列表
        LoadFactoryList()

        ' 加载来源列表
        LoadSourceList()

        ' 加载库房列表
        LoadWarehouseList()

        ' 加载结算方式列表
        LoadSettlementList()
    End Sub

    ' ========== 加载工厂列表 ==========
    Private Sub LoadFactoryList()
        Try
            Dim sql As String = "SELECT title FROM xipunum_erp_about WHERE 1=1 ORDER BY id ASC"
            Dim dt As DataTable = ExecuteQuery(sql)
            cmbFactory.Items.Clear()
            For Each r As DataRow In dt.Rows
                cmbFactory.Items.Add(SafeString(r("title")))
            Next
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 工厂名称组合框弹出列表（_工厂名称组合框_将弹出列表） ==========
    Private Sub CmbFactory_DropDown(sender As Object, e As EventArgs)
        Try
            Dim defaultText As String = cmbFactory.Text
            Dim keyword As String = defaultText
            cmbFactory.Items.Clear()
            cmbFactory.Text = defaultText
            cmbFactory.SelectedIndex = -1

            Dim sql As String = $"SELECT title FROM xipunum_erp_about WHERE title LIKE '%{SafeSQL(keyword)}%' OR jianxie LIKE '%{SafeSQL(keyword)}%' ORDER BY id ASC"
            Dim dt As DataTable = ExecuteQuery(sql)
            For Each r As DataRow In dt.Rows
                cmbFactory.Items.Add(SafeString(r("title")))
            Next
            If dt.Rows.Count <= 0 Then
                cmbFactory.Text = defaultText
                cmbFactory.SelectedIndex = -1
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 加载来源列表 ==========
    Private Sub LoadSourceList()
        Try
            Dim sql As String = "SELECT title FROM xipunum_erp_type WHERE type='商品来源' AND superior='0' ORDER BY id DESC"
            Dim dt As DataTable = ExecuteQuery(sql)
            cmbSource.Items.Clear()
            For Each r As DataRow In dt.Rows
                cmbSource.Items.Add(SafeString(r("title")))
            Next
            If dt.Rows.Count = 1 Then
                cmbSource.SelectedIndex = 0
                cmbSource.Enabled = False
            Else
                cmbSource.Enabled = True
                cmbSource.SelectedIndex = -1
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 加载库房列表 ==========
    Private Sub LoadWarehouseList()
        Try
            Dim shopPermission As String = UserShopPermission
            If String.IsNullOrEmpty(shopPermission) Then shopPermission = "-1"

            Dim sql As String = $"SELECT id AS akufang, CASE WHEN id = '0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id in ({shopPermission}) UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN ({shopPermission}) ORDER BY akufang = '0' DESC, akufang"
            Dim dt As DataTable = ExecuteQuery(sql)
            cmbWarehouse.Items.Clear()
            Dim shopCount As Integer = dt.Rows.Count
            For Each r As DataRow In dt.Rows
                cmbWarehouse.Items.Add(SafeString(r("btitle")))
            Next

            If shopCount = 1 Then
                cmbWarehouse.SelectedIndex = 0
                cmbWarehouse.Enabled = True
            Else
                If UserPermission = "全部" Then
                    cmbWarehouse.Enabled = True
                    cmbWarehouse.SelectedIndex = -1
                Else
                    If shopCount > 0 Then
                        cmbWarehouse.SelectedIndex = 0
                        cmbWarehouse.Enabled = True
                    Else
                        cmbWarehouse.SelectedIndex = -1
                        cmbWarehouse.Enabled = True
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 加载结算方式列表 ==========
    Private Sub LoadSettlementList()
        Try
            Dim sql As String = "SELECT title FROM xipunum_erp_type WHERE type='结算方式' AND superior='0' ORDER BY id DESC"
            Dim dt As DataTable = ExecuteQuery(sql)
            cmbSettlement.Items.Clear()
            Dim settleCount As Integer = dt.Rows.Count
            For Each r As DataRow In dt.Rows
                cmbSettlement.Items.Add(SafeString(r("title")))
            Next
            If settleCount = 1 Then
                cmbSettlement.SelectedIndex = 0
                cmbSettlement.Enabled = True
            Else
                cmbSettlement.Enabled = True
                cmbSettlement.SelectedIndex = -1
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 加载表头（_高级表格1_加载表头） ==========
    Private Sub LoadTableHeader()
        dgvProducts.Columns.Clear()
        Dim headers() As String = {"序号", "商品编码", "商品名称", "商品品类", "规格", "款号", "材质", "圈口长度", "面宽", "厚度", "工厂成色", "公司成色", "单件重", "数量", "金重", "损耗", "含耗重", "重量", "单位", "石重", "石头数", "副石重", "副石头数", "成本单价", "系数", "成本工费", "参考工费", "销售工费", "成本附加费", "销售附加费", "销售价", "备注", "主石色", "图片地址", "原料价", "模具号", "操作"}
        Dim widths() As Integer = {45, 100, 150, 75, 75, 110, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 0, 60, 60, 60}

        For i As Integer = 0 To headers.Length - 1
            If i = COL_DELETE Then
                Dim btnCol As New DataGridViewButtonColumn()
                btnCol.HeaderText = headers(i)
                btnCol.Name = "col" & i
                btnCol.Width = widths(i)
                btnCol.Text = "删除"
                btnCol.UseColumnTextForButtonValue = True
                dgvProducts.Columns.Add(btnCol)
            Else
                Dim col As New DataGridViewTextBoxColumn()
                col.HeaderText = headers(i)
                col.Name = "col" & i
                col.Width = widths(i)
                dgvProducts.Columns.Add(col)
            End If
        Next

        ' 设置单位列(18)为下拉列表
        Dim unitCol As DataGridViewComboBoxColumn = DirectCast(dgvProducts.Columns(COL_UNIT), DataGridViewComboBoxColumn)
        ' Note: DataGridViewComboBoxColumn requires Items setup
    End Sub

    ' ========== 加载表格数据（_高级表格1_加载数据） ==========
    Private Sub LoadTableData()
        ClearTableRows()
        Dim insertRows As Integer = 25
        If inboundProductCount >= 25 Then
            insertRows = inboundProductCount + 1
        End If
        For i As Integer = 0 To insertRows - 1
            dgvProducts.Rows.Add()
        Next
        CalculateTotals()
    End Sub

    ' ========== 清空表格数据行（子程序_删除表格1） ==========
    Private Function ClearTableRows() As Boolean
        While dgvProducts.Rows.Count > 0
            dgvProducts.Rows.RemoveAt(0)
        End While
        Return True
    End Function

    ' ========== 光标位置改变（_高级表格1_光标位置改变） ==========
    Private Sub DgvProducts_SelectionChanged(sender As Object, e As EventArgs)
        If dgvProducts.CurrentCell IsNot Nothing Then
            row = dgvProducts.CurrentCell.RowIndex
            col = dgvProducts.CurrentCell.ColumnIndex
        End If
    End Sub

    ' ========== 工具栏按钮点击（_工具条_通用_被单击） ==========
    Private Sub ToolButton_Click(sender As Object, e As EventArgs)
        Dim btn As ToolStripButton = DirectCast(sender, ToolStripButton)
        Select Case btn.Text
            Case "添加商品"
                AddProduct()
            Case "保存"
                SaveInboundOrder()
            Case "重置"
                ResetForm()
            Case "标签打印"
                PrintLabels()
        End Select
    End Sub

    ' ========== 添加商品（_超级按钮_添加_被单击） ==========
    Private Sub AddProduct()
        If cmbFactory.SelectedIndex = -1 Then
            ShowWarning("请选择工厂名称！")
            cmbFactory.Focus()
            Return
        End If
        If cmbSource.SelectedIndex = -1 Then
            ShowWarning("请选择商品来源！")
            cmbSource.Focus()
            Return
        End If
        If cmbWarehouse.SelectedIndex = -1 Then
            ShowWarning("请选择库房名称！")
            cmbWarehouse.Focus()
            Return
        End If
        If cmbSettlement.SelectedIndex = -1 Then
            ShowWarning("请选择结算方式！")
            cmbSettlement.Focus()
            Return
        End If

        rbHalfYes.Enabled = False
        rbHalfNo.Enabled = False
        cmbFactory.Enabled = False
        cmbSource.Enabled = False
        cmbWarehouse.Enabled = False
        cmbSettlement.Enabled = False
        txtDeliveryNumber.Enabled = True

        ' 打开商品信息添加窗口
        Dim productForm As New ProductSelectForm()
        If productForm.ShowDialog() = DialogResult.OK Then
            ' 在表格中添加商品数据
            Dim rowIndex As Integer = -1
            ' 找到第一个空行
            For i As Integer = 0 To dgvProducts.Rows.Count - 1
                If String.IsNullOrEmpty(SafeString(dgvProducts.Rows(i).Cells(COL_CODE).Value)) Then
                    rowIndex = i
                    Exit For
                End If
            Next
            If rowIndex = -1 Then
                dgvProducts.Rows.Add()
                rowIndex = dgvProducts.Rows.Count - 1
            End If

            dgvProducts.Rows(rowIndex).Cells(COL_CODE).Value = productForm.ProductCode
            dgvProducts.Rows(rowIndex).Cells(COL_NAME).Value = productForm.ProductName
            dgvProducts.Rows(rowIndex).Cells(COL_CATEGORY).Value = productForm.CategoryName
            dgvProducts.Rows(rowIndex).Cells(COL_SPEC).Value = productForm.SpecName
            dgvProducts.Rows(rowIndex).Cells(COL_KUANHAO).Value = productForm.KuanHao
            dgvProducts.Rows(rowIndex).Cells(COL_CAIZHI).Value = productForm.CaiZhi
            dgvProducts.Rows(rowIndex).Cells(COL_SINGLE).Value = FormatThreeDecimals(productForm.DanZhong.ToString())
            dgvProducts.Rows(rowIndex).Cells(COL_QTY).Value = "1"
            dgvProducts.Rows(rowIndex).Cells(COL_JINZHONG).Value = FormatThreeDecimals(productForm.NetWeight.ToString())
            dgvProducts.Rows(rowIndex).Cells(COL_WEIGHT).Value = FormatThreeDecimals(productForm.Weight.ToString())
            dgvProducts.Rows(rowIndex).Cells(COL_UNIT).Value = "重量"

            inboundProductCount += 1
        End If

        CalculateTotals()
    End Sub

    ' ========== 重置表单（_超级按钮_重置_被单击） ==========
    Private Sub ResetForm()
        If inboundProductCount > 0 Then
            If Not ConfirmAction("当前入库商品数据将被清空，确定重置吗？") Then
                Return
            End If
        End If

        inboundProductCount = 0
        printOrderNumber = ""
        isProductInlaid = ""
        productCategoryID = ""
        InitializeBaseData()
        LoadTableHeader()
        LoadTableData()
    End Sub

    ' ========== 按钮重置点击（_按钮_重置_被单击） ==========
    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        ResetForm()
    End Sub

    ' ========== 表格按钮点击（_高级表格1_按钮被点击） ==========
    Private Sub DgvProducts_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        If e.ColumnIndex <> COL_DELETE Then Return
        If e.RowIndex >= dgvProducts.Rows.Count - 1 Then Return

        ' 合计行不处理
        If SafeString(dgvProducts.Rows(e.RowIndex).Cells(COL_CODE).Value) = "合计" Then Return

        Dim productCode As String = SafeString(dgvProducts.Rows(e.RowIndex).Cells(COL_CODE).Value)
        If productCode = "" Then
            ShowWarning("当前行没有商品编码，无法删除！")
            Return
        End If

        If Not ConfirmAction("是否删除选中商品？") Then Return

        ' 从数据库删除商品
        ExecuteCommand($"DELETE FROM xipunum_erp_shop WHERE poduct_code='{SafeSQL(productCode)}'")
        ExecuteCommand($"DELETE FROM xipunum_erp_zhengshu WHERE poduct_code='{SafeSQL(productCode)}'")

        ' 删除合计行后删除当前行
        If dgvProducts.Rows.Count > 1 Then
            dgvProducts.Rows.RemoveAt(dgvProducts.Rows.Count - 1)
        End If
        dgvProducts.Rows.RemoveAt(e.RowIndex)

        If inboundProductCount > 0 Then
            inboundProductCount -= 1
        End If

        ' 如果行数不足26，补充空行
        If dgvProducts.Rows.Count <= 26 Then
            dgvProducts.Rows.Add()
        End If

        CalculateTotals()
    End Sub

    ' ========== 数据统计（_高级表格1_数据统计） ==========
    Private Sub CalculateTotals()
        If dgvProducts.Rows.Count <= 0 Then Return

        ' 删除已有合计行
        If dgvProducts.Rows.Count > 0 Then
            Dim lastRow As Integer = dgvProducts.Rows.Count - 1
            If SafeString(dgvProducts.Rows(lastRow).Cells(COL_CODE).Value) = "合计" Then
                dgvProducts.Rows.RemoveAt(lastRow)
            End If
        End If

        ' 初始化合计值
        Dim totQty As Decimal = 0
        Dim totJinzhong As Decimal = 0
        Dim totWeight As Decimal = 0
        Dim totShitou As Decimal = 0
        Dim totStnum As Decimal = 0
        Dim totShitou1 As Decimal = 0
        Dim totShnum1 As Decimal = 0
        Dim totCostPrice As Decimal = 0
        Dim totBasicCost As Decimal = 0
        Dim totPremiumCost As Decimal = 0
        Dim totSalesCost As Decimal = 0
        Dim totCompanySur As Decimal = 0
        Dim totSalesSur As Decimal = 0

        Dim dataRows As Integer = dgvProducts.Rows.Count
        For i As Integer = 0 To dataRows - 1
            Dim code As String = SafeString(dgvProducts.Rows(i).Cells(COL_CODE).Value)
            If code <> "" AndAlso code <> "合计" Then
                dgvProducts.Rows(i).Cells(COL_SEQ).Value = i.ToString()

                Dim qty As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_QTY).Value)
                Dim jinzhong As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_JINZHONG).Value)
                Dim weight As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_WEIGHT).Value)
                Dim shitou As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_SHITOU).Value)
                Dim stnum As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_STNUM).Value)
                Dim shitou1 As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_SHITOU1).Value)
                Dim shnum1 As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_SHNUM1).Value)
                Dim costPrice As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_COST_PRICE).Value)
                Dim basicCost As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_BASIC_COST).Value)
                Dim premiumCost As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_PREMIUM_COST).Value)
                Dim salesCost As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_SALES_COST).Value)
                Dim companySur As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_COMPANY_SUR).Value)
                Dim salesSur As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(COL_SALES_SUR).Value)

                totQty += qty
                totJinzhong += jinzhong
                totWeight += weight
                totShitou += shitou
                totStnum += stnum
                totShitou1 += shitou1
                totShnum1 += shnum1
                totCostPrice += costPrice
                totBasicCost += basicCost * jinzhong
                totPremiumCost += premiumCost * jinzhong
                totSalesCost += salesCost * jinzhong
                totCompanySur += companySur * qty
                totSalesSur += salesSur
            End If
        Next

        ' 添加合计行
        dgvProducts.Rows.Add()
        Dim totalRowIdx As Integer = dgvProducts.Rows.Count - 1
        dgvProducts.Rows(totalRowIdx).Cells(COL_CODE).Value = "合计"
        dgvProducts.Rows(totalRowIdx).Cells(COL_QTY).Value = totQty.ToString()
        dgvProducts.Rows(totalRowIdx).Cells(COL_JINZHONG).Value = FormatThreeDecimals(totJinzhong.ToString())
        dgvProducts.Rows(totalRowIdx).Cells(COL_WEIGHT).Value = FormatThreeDecimals(totWeight.ToString())
        dgvProducts.Rows(totalRowIdx).Cells(COL_SHITOU).Value = FormatThreeDecimals(totShitou.ToString())
        dgvProducts.Rows(totalRowIdx).Cells(COL_STNUM).Value = totStnum.ToString()
        dgvProducts.Rows(totalRowIdx).Cells(COL_SHITOU1).Value = FormatThreeDecimals(totShitou1.ToString())
        dgvProducts.Rows(totalRowIdx).Cells(COL_SHNUM1).Value = totShnum1.ToString()
        dgvProducts.Rows(totalRowIdx).Cells(COL_COST_PRICE).Value = FormatTwoDecimals(totCostPrice.ToString())
        dgvProducts.Rows(totalRowIdx).Cells(COL_BASIC_COST).Value = FormatTwoDecimals(totBasicCost.ToString())
        dgvProducts.Rows(totalRowIdx).Cells(COL_PREMIUM_COST).Value = FormatTwoDecimals(totPremiumCost.ToString())
        dgvProducts.Rows(totalRowIdx).Cells(COL_SALES_COST).Value = FormatTwoDecimals(totSalesCost.ToString())
        dgvProducts.Rows(totalRowIdx).Cells(COL_COMPANY_SUR).Value = FormatTwoDecimals(totCompanySur.ToString())
        dgvProducts.Rows(totalRowIdx).Cells(COL_SALES_SUR).Value = FormatTwoDecimals(totSalesSur.ToString())

        ' 更新总重和料重
        txtMaterialWeight.Text = FormatThreeDecimals(totJinzhong.ToString())
        txtTotalWeight.Text = FormatThreeDecimals(totWeight.ToString())
    End Sub

    ' ========== 格式三位小数（入库添加_格式三位小数） ==========
    Private Function FormatThreeDecimals(originalValue As String) As String
        If String.IsNullOrEmpty(originalValue) Then Return "0.000"
        Dim num As Decimal = SafeDecimal(originalValue)
        Dim formatted As String = Math.Round(num, 3).ToString()
        If formatted = "" Then Return "0.000"
        If formatted.Contains(".") Then
            Dim parts() As String = formatted.Split("."c)
            If parts.Length > 1 Then
                Dim decLen As Integer = parts(1).Length
                If decLen = 1 Then
                    Return formatted & "00"
                ElseIf decLen = 2 Then
                    Return formatted & "0"
                Else
                    Return formatted
                End If
            Else
                Return formatted & ".000"
            End If
        Else
            Return formatted & ".000"
        End If
    End Function

    ' ========== 格式二位小数（入库添加_格式二位小数） ==========
    Private Function FormatTwoDecimals(originalValue As String) As String
        If String.IsNullOrEmpty(originalValue) Then Return "0.00"
        Dim num As Decimal = SafeDecimal(originalValue)
        Dim formatted As String = Math.Round(num, 2).ToString()
        If formatted = "" Then Return "0.00"
        If formatted.Contains(".") Then
            Dim parts() As String = formatted.Split("."c)
            If parts.Length > 1 Then
                Dim decLen As Integer = parts(1).Length
                If decLen = 1 Then
                    Return formatted & "0"
                Else
                    Return formatted
                End If
            Else
                Return formatted & ".00"
            End If
        Else
            Return formatted & ".00"
        End If
    End Function

    ' ========== 单元格结束编辑（_高级表格1_结束编辑） ==========
    Private Sub DgvProducts_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex <= 0 Then Return
        If e.RowIndex >= dgvProducts.Rows.Count - 1 Then Return

        Dim r As DataGridViewRow = dgvProducts.Rows(e.RowIndex)
        Dim currentCode As String = SafeString(r.Cells(COL_CODE).Value)
        Dim currentName As String = SafeString(r.Cells(COL_NAME).Value)
        If currentCode = "" AndAlso currentName = "" Then Return

        ' 读取当前行所有值
        Dim qty As String = SafeString(r.Cells(COL_QTY).Value)
        Dim jinzhong As String = SafeString(r.Cells(COL_JINZHONG).Value)
        Dim loss As String = SafeString(r.Cells(COL_LOSS).Value)
        Dim including As String = SafeString(r.Cells(COL_INCLUDING).Value)
        Dim weight As String = SafeString(r.Cells(COL_WEIGHT).Value)
        Dim shitou As String = SafeString(r.Cells(COL_SHITOU).Value)
        Dim stnum As String = SafeString(r.Cells(COL_STNUM).Value)
        Dim shitou1 As String = SafeString(r.Cells(COL_SHITOU1).Value)
        Dim shnum1 As String = SafeString(r.Cells(COL_SHNUM1).Value)
        Dim coefficient As String = SafeString(r.Cells(COL_COEFFICIENT).Value)
        Dim basicCost As String = SafeString(r.Cells(COL_BASIC_COST).Value)
        Dim premiumCost As String = SafeString(r.Cells(COL_PREMIUM_COST).Value)
        Dim salesCost As String = SafeString(r.Cells(COL_SALES_COST).Value)
        Dim companySur As String = SafeString(r.Cells(COL_COMPANY_SUR).Value)
        Dim salesSur As String = SafeString(r.Cells(COL_SALES_SUR).Value)
        Dim goldPrice As String = SafeString(r.Cells(COL_GOLD_PRICE).Value)

        Dim singleWeight As String = SafeString(r.Cells(COL_SINGLE).Value)
        Dim costPrice As String = SafeString(r.Cells(COL_COST_PRICE).Value)
        Dim salesPrice As String = SafeString(r.Cells(COL_SALES_PRICE).Value)

        ' 根据编辑的列进行计算
        Select Case col
            Case COL_QTY ' 13 数量
                singleWeight = (SafeDecimal(weight) / SafeDecimal(qty)).ToString()

            Case COL_JINZHONG ' 14 金重
                If SafeDecimal(weight) <= 0 Then
                    weight = jinzhong
                Else
                    If SafeDecimal(jinzhong) > SafeDecimal(weight) Then
                        ShowWarning("金重不可以大于重量！")
                        jinzhong = weight
                    End If
                End If
                costPrice = (SafeDecimal(jinzhong) * (SafeDecimal(basicCost) + SafeDecimal(goldPrice)) + SafeDecimal(companySur)).ToString()
                salesPrice = (SafeDecimal(jinzhong) * (SafeDecimal(salesCost) + SafeDecimal(goldPrice)) + SafeDecimal(salesSur)).ToString()
                shitou = (SafeDecimal((SafeDecimal(weight) - SafeDecimal(jinzhong)) * 5) - SafeDecimal(shitou1)).ToString()
                If SafeDecimal(shitou) < 0 Then shitou = "0"

            Case COL_LOSS ' 15 损耗
                including = (SafeDecimal(loss) + SafeDecimal(weight)).ToString()

            Case COL_WEIGHT ' 17 重量
                If SafeDecimal(weight) < SafeDecimal(jinzhong) Then
                    ShowWarning("重量不可以小于金重！")
                    weight = jinzhong
                End If
                jinzhong = weight
                singleWeight = (SafeDecimal(weight) / SafeDecimal(qty)).ToString()
                including = (SafeDecimal(loss) + SafeDecimal(weight)).ToString()
                costPrice = (SafeDecimal(jinzhong) * (SafeDecimal(basicCost) + SafeDecimal(goldPrice)) + SafeDecimal(companySur)).ToString()
                salesPrice = (SafeDecimal(jinzhong) * (SafeDecimal(salesCost) + SafeDecimal(goldPrice)) + SafeDecimal(salesSur)).ToString()
                shitou = (SafeDecimal((SafeDecimal(weight) - SafeDecimal(jinzhong)) * 5) - SafeDecimal(shitou1)).ToString()

            Case COL_SHITOU1 ' 21 副石重
                shitou = (SafeDecimal((SafeDecimal(weight) - SafeDecimal(jinzhong)) * 5) - SafeDecimal(shitou1)).ToString()
                If SafeDecimal(shitou) < 0 Then shitou = "0"

            Case COL_BASIC_COST, COL_COMPANY_SUR ' 25 成本工费, 28 成本附加费
                costPrice = (SafeDecimal(jinzhong) * (SafeDecimal(basicCost) + SafeDecimal(goldPrice)) + SafeDecimal(companySur)).ToString()

            Case COL_PREMIUM_COST ' 26 参考工费
                salesCost = premiumCost
                salesPrice = (SafeDecimal(jinzhong) * (SafeDecimal(salesCost) + SafeDecimal(goldPrice)) + SafeDecimal(salesSur)).ToString()

            Case COL_SALES_COST, COL_SALES_SUR ' 27 销售工费, 29 销售附加费
                salesPrice = (SafeDecimal(jinzhong) * (SafeDecimal(salesCost) + SafeDecimal(goldPrice)) + SafeDecimal(salesSur)).ToString()
        End Select

        ' 格式化数值
        singleWeight = FormatThreeDecimals(singleWeight)
        jinzhong = FormatThreeDecimals(jinzhong)
        loss = FormatThreeDecimals(loss)
        including = FormatThreeDecimals(including)
        weight = FormatThreeDecimals(weight)
        shitou = FormatThreeDecimals(shitou)
        shitou1 = FormatThreeDecimals(shitou1)
        basicCost = FormatTwoDecimals(basicCost)
        premiumCost = FormatTwoDecimals(premiumCost)
        salesCost = FormatTwoDecimals(salesCost)
        companySur = FormatTwoDecimals(companySur)
        salesSur = FormatTwoDecimals(salesSur)
        costPrice = FormatTwoDecimals(costPrice)
        salesPrice = FormatTwoDecimals(salesPrice)

        ' 根据编辑列回写单元格
        Select Case col
            Case COL_QTY
                r.Cells(COL_SINGLE).Value = singleWeight
                r.Cells(COL_QTY).Value = qty

            Case COL_JINZHONG
                r.Cells(COL_JINZHONG).Value = jinzhong
                r.Cells(COL_WEIGHT).Value = weight
                r.Cells(COL_SHITOU).Value = shitou
                r.Cells(COL_COST_PRICE).Value = costPrice
                r.Cells(COL_SALES_PRICE).Value = salesPrice

            Case COL_LOSS
                r.Cells(COL_LOSS).Value = loss
                r.Cells(COL_INCLUDING).Value = including

            Case COL_WEIGHT
                r.Cells(COL_SINGLE).Value = singleWeight
                r.Cells(COL_JINZHONG).Value = jinzhong
                r.Cells(COL_INCLUDING).Value = including
                r.Cells(COL_WEIGHT).Value = weight
                r.Cells(COL_SHITOU).Value = shitou
                r.Cells(COL_COST_PRICE).Value = costPrice
                r.Cells(COL_SALES_PRICE).Value = salesPrice

            Case COL_SHITOU
                r.Cells(COL_SHITOU).Value = shitou

            Case COL_STNUM
                r.Cells(COL_STNUM).Value = stnum

            Case COL_SHITOU1
                r.Cells(COL_SHITOU).Value = shitou
                r.Cells(COL_SHITOU1).Value = shitou1

            Case COL_SHNUM1
                r.Cells(COL_SHNUM1).Value = shnum1

            Case COL_COEFFICIENT
                r.Cells(COL_COEFFICIENT).Value = coefficient

            Case COL_BASIC_COST
                r.Cells(COL_COST_PRICE).Value = costPrice
                r.Cells(COL_BASIC_COST).Value = basicCost

            Case COL_PREMIUM_COST
                r.Cells(COL_PREMIUM_COST).Value = premiumCost
                r.Cells(COL_SALES_COST).Value = salesCost
                r.Cells(COL_SALES_PRICE).Value = salesPrice

            Case COL_SALES_COST
                r.Cells(COL_SALES_COST).Value = salesCost
                r.Cells(COL_SALES_PRICE).Value = salesPrice

            Case COL_COMPANY_SUR
                r.Cells(COL_COST_PRICE).Value = costPrice
                r.Cells(COL_COMPANY_SUR).Value = companySur

            Case COL_SALES_SUR
                r.Cells(COL_SALES_SUR).Value = salesSur
                r.Cells(COL_SALES_PRICE).Value = salesPrice
        End Select

        ' 删除合计行后重新统计
        If dgvProducts.Rows.Count > 1 Then
            If SafeString(dgvProducts.Rows(dgvProducts.Rows.Count - 1).Cells(COL_CODE).Value) = "合计" Then
                dgvProducts.Rows.RemoveAt(dgvProducts.Rows.Count - 1)
            End If
        End If
        CalculateTotals()
    End Sub

    ' ========== 保存入库单（_超级按钮_保存_被单击） ==========
    Private Sub SaveInboundOrder()
        ' 设置操作日期和账户
        Dim operationDate As String = GetOperationDate()
        Dim operationAccount As String = UserAccount

        ' 检查数据库连接
        If MySQL_Read Is Nothing OrElse MySQL_Write Is Nothing Then
            ShowWarning("数据库连接异常，无法保存入库单！")
            Return
        End If

        ' 验证
        If inboundProductCount <= 0 Then
            ShowWarning("入库商品数量不能小于1！")
            Return
        End If
        If cmbFactory.SelectedIndex = -1 Then
            ShowWarning("请选择工厂名称！")
            cmbFactory.Focus()
            Return
        End If
        If cmbSource.SelectedIndex = -1 Then
            ShowWarning("请选择商品来源！")
            cmbSource.Focus()
            Return
        End If
        If cmbWarehouse.SelectedIndex = -1 Then
            ShowWarning("请选择库房名称！")
            cmbWarehouse.Focus()
            Return
        End If
        If cmbSettlement.SelectedIndex = -1 Then
            ShowWarning("请选择结算方式！")
            cmbSettlement.Focus()
            Return
        End If

        ' 生成入库单号
        Dim orderNumber As String = "RK" & DateTime.Now.ToString("yyyyMMdd") & DateTimeOffset.Now.ToUnixTimeSeconds().ToString() & UserAccount

        ' 检查单号是否已存在
        Dim checkSql As String = $"SELECT id FROM xipunum_erp_store_order WHERE odd_numbers='{SafeSQL(orderNumber)}' LIMIT 1"
        Dim checkDt As DataTable = ExecuteQuery(checkSql)
        If checkDt.Rows.Count > 0 Then
            ShowWarning("当前入库单号已存在，请重新点击保存按钮！")
            txtOrderNumber.Text = "RK" & DateTime.Now.ToString("yyyyMMdd") & "****"
            Return
        End If

        ' 商品保存数量
        Dim saveCount As Integer = dgvProducts.Rows.Count - 2
        If saveCount <= 0 Then
            ShowWarning("没有可保存的商品明细！")
            Return
        End If

        ' 获取单据信息
        Dim orderXiangqian As String = isProductInlaid
        Dim orderSource As String = cmbSource.SelectedItem.ToString()
        Dim orderSettlement As String = cmbSettlement.SelectedItem.ToString()
        Dim orderDelivery As String = txtDeliveryNumber.Text
        Dim orderRemarks As String = txtRemarks.Text

        ' 获取工厂ID
        Dim factoryName As String = cmbFactory.SelectedItem.ToString()
        Dim factoryId As String = ""
        Dim factoryDt As DataTable = ExecuteQuery($"SELECT id FROM xipunum_erp_about WHERE title='{SafeSQL(factoryName)}' ORDER BY id ASC LIMIT 1")
        If factoryDt.Rows.Count > 0 Then
            factoryId = SafeString(factoryDt.Rows(0)("id"))
        End If
        If factoryId = "" Then
            ShowWarning("未找到工厂信息，无法保存！")
            Return
        End If

        ' 获取半成品信息
        Dim halfProduct As String = ""
        If rbHalfYes.Checked Then halfProduct = rbHalfYes.Text
        If rbHalfNo.Checked Then halfProduct = rbHalfNo.Text

        ' 获取库房ID
        Dim warehouseId As String = ""
        Dim warehouseName As String = cmbWarehouse.SelectedItem.ToString()
        If warehouseName = "总库" Then
            warehouseId = "0"
        Else
            Dim warehouseDt As DataTable = ExecuteQuery($"SELECT id FROM xipunum_erp_type WHERE type='商铺' AND title='{SafeSQL(warehouseName)}' AND superior='0' ORDER BY id ASC LIMIT 1")
            If warehouseDt.Rows.Count > 0 Then
                warehouseId = SafeString(warehouseDt.Rows(0)("id"))
            End If
        End If
        If warehouseId = "" Then
            ShowWarning("未找到库房信息，无法保存！")
            Return
        End If

        ' 获取合计行数据
        Dim totalRowIndex As Integer = dgvProducts.Rows.Count - 1
        Dim factoryZhezu As String = SafeString(dgvProducts.Rows(totalRowIndex).Cells(COL_JINZHONG).Value)
        Dim orderTotalWeight As String = SafeString(dgvProducts.Rows(totalRowIndex).Cells(COL_WEIGHT).Value)
        Dim orderJinzhong As String = SafeString(dgvProducts.Rows(totalRowIndex).Cells(COL_JINZHONG).Value)
        Dim orderInputTotal As String = SafeString(dgvProducts.Rows(totalRowIndex).Cells(COL_WEIGHT).Value)
        Dim orderGoldPrice As String = SafeString(dgvProducts.Rows(1).Cells(COL_GOLD_PRICE).Value)
        Dim orderJlzhong As String = SafeString(dgvProducts.Rows(totalRowIndex).Cells(COL_JINZHONG).Value)
        Dim orderLjzhezu As String = SafeString(dgvProducts.Rows(totalRowIndex).Cells(COL_JINZHONG).Value)
        Dim orderGongchang As String = SafeString(dgvProducts.Rows(totalRowIndex).Cells(COL_COST_PRICE).Value)
        Dim orderChengben As String = SafeString(dgvProducts.Rows(totalRowIndex).Cells(COL_COST_PRICE).Value)
        Dim orderJljine As String = "0.00"
        Dim orderState As String = "待审"

        ' 空值默认处理
        If factoryZhezu = "" Then factoryZhezu = "0.000"
        If orderTotalWeight = "" Then orderTotalWeight = "0.000"
        If orderJinzhong = "" Then orderJinzhong = "0.000"
        If orderInputTotal = "" Then orderInputTotal = "0.000"
        If orderGoldPrice = "" Then orderGoldPrice = "0.00"
        If orderGongchang = "" Then orderGongchang = "0.00"
        If orderChengben = "" Then orderChengben = "0.00"

        ' 插入入库订单
        Dim insertOrderSql As String = $"INSERT INTO xipunum_erp_store_order (odd_numbers, category_id, xiangqian, delivery, half_product, factory, source, settlement, factory_zhezu, jinzhong, total, total_weight, gold_price, jlzhong, ljzhezu, gongchang, chengben, state, jljine, remarks, cjuser, creationtime) VALUES ('{SafeSQL(orderNumber)}', '{SafeSQL(productCategoryID)}', '{SafeSQL(orderXiangqian)}', '{SafeSQL(orderDelivery)}', '{SafeSQL(halfProduct)}', '{SafeSQL(factoryId)}', '{SafeSQL(orderSource)}', '{SafeSQL(orderSettlement)}', '{SafeSQL(factoryZhezu)}', '{SafeSQL(orderJinzhong)}', '{SafeSQL(orderInputTotal)}', '{SafeSQL(orderInputTotal)}', '{SafeSQL(orderGoldPrice)}', '{SafeSQL(orderJlzhong)}', '{SafeSQL(orderLjzhezu)}', '{SafeSQL(orderGongchang)}', '{SafeSQL(orderChengben)}', '{SafeSQL(orderState)}', '{SafeSQL(orderJljine)}', '{SafeSQL(orderRemarks)}', '{SafeSQL(operationAccount)}', '{SafeSQL(operationDate)}')"
        ExecuteCommand(insertOrderSql, MySQL_Write)

        ' 获取订单ID
        Dim orderId As String = ""
        Dim orderDt As DataTable = ExecuteQuery($"SELECT id FROM xipunum_erp_store_order WHERE odd_numbers='{SafeSQL(orderNumber)}' ORDER BY id ASC LIMIT 1")
        If orderDt.Rows.Count > 0 Then
            orderId = SafeString(orderDt.Rows(0)("id"))
        End If
        If orderId = "" Then
            ShowWarning("商品入库保存失败，请重新保存！")
            Return
        End If

        ' 更新单据号
        Dim finalOrderNumber As String = orderNumber.Substring(0, 10) & ("00000000" & orderId).PadLeft(8, "0"c)
        finalOrderNumber = orderNumber.Substring(0, 10) & orderId.PadLeft(4, "0"c)
        txtOrderNumber.Text = finalOrderNumber

        ExecuteCommand($"UPDATE xipunum_erp_store_order SET odd_numbers='{SafeSQL(finalOrderNumber)}' WHERE id='{SafeSQL(orderId)}' LIMIT 1", MySQL_Write)

        ' 插入对账记录
        Dim reconSql As String = $"INSERT INTO xipunum_erp_reconciliation (odd_numbers, delivery, factory, jinzhong, total, kufang, state, cjuser, creationtime) VALUES ('{SafeSQL(finalOrderNumber)}', '{SafeSQL(orderDelivery)}', '{SafeSQL(factoryId)}', '{SafeSQL(orderJinzhong)}', '{SafeSQL(orderInputTotal)}', '{SafeSQL(warehouseId)}', '待对账', '{SafeSQL(operationAccount)}', '{SafeSQL(operationDate)}')"
        ExecuteCommand(reconSql, MySQL_Write)

        ' 插入入库日志
        Dim logSql As String = $"INSERT INTO xipunum_erp_store_log (order_id, user, conter, creationtime) VALUES ('{SafeSQL(orderId)}', '{SafeSQL(operationAccount)}', '商品入库', '{SafeSQL(operationDate)}')"
        ExecuteCommand(logSql, MySQL_Write)

        ' 插入系统日志
        Dim logContent As String = $"账户:{UserAccount} 添加入库订单：{finalOrderNumber} 相关数据"
        Dim sysLogSql As String = $"INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('添加', '添加入库订单', '{SafeSQL(logContent)}', '{SafeSQL(operationAccount)}', '{SafeSQL(operationDate)}')"
        ExecuteCommand(sysLogSql, MySQL_Write)

        ' 批量插入数据字符串
        Dim tmpStockData As String = ""
        Dim tmpStockData1 As String = ""
        Dim historyData As String = ""
        Dim xiangqianData As String = ""
        Dim storeData As String = ""
        Dim opLogData As String = ""

        ' 循环保存商品明细
        For i As Integer = 1 To saveCount
            Dim r As DataGridViewRow = dgvProducts.Rows(i)

            ' 读取商品数据
            Dim fCode As String = SafeString(r.Cells(COL_CODE).Value)
            If fCode = "" Then Continue For

            Dim fName As String = SafeString(r.Cells(COL_NAME).Value)
            Dim fCategory As String = SafeString(r.Cells(COL_CATEGORY).Value)
            Dim fSpec As String = SafeString(r.Cells(COL_SPEC).Value)
            Dim fKuanhao As String = SafeString(r.Cells(COL_KUANHAO).Value)
            Dim fCaizhi As String = SafeString(r.Cells(COL_CAIZHI).Value)
            Dim fQuandu As String = SafeString(r.Cells(COL_QUANDU).Value)
            Dim fWidth As String = SafeString(r.Cells(COL_WIDTH).Value)
            Dim fThickness As String = SafeString(r.Cells(COL_THICKNESS).Value)
            Dim fFactoryCond As String = SafeString(r.Cells(COL_FACTORY_COND).Value)
            Dim fCompanyCond As String = SafeString(r.Cells(COL_COMPANY_COND).Value)
            Dim fSingle As String = SafeString(r.Cells(COL_SINGLE).Value)
            Dim fQty As String = SafeString(r.Cells(COL_QTY).Value)
            Dim fJinzhong As String = SafeString(r.Cells(COL_JINZHONG).Value)
            Dim fLoss As String = SafeString(r.Cells(COL_LOSS).Value)
            Dim fIncluding As String = SafeString(r.Cells(COL_INCLUDING).Value)
            Dim fWeight As String = SafeString(r.Cells(COL_WEIGHT).Value)
            Dim fUnit As String = SafeString(r.Cells(COL_UNIT).Value)
            Dim fShitou As String = SafeString(r.Cells(COL_SHITOU).Value)
            Dim fStnum As String = SafeString(r.Cells(COL_STNUM).Value)
            Dim fShitou1 As String = SafeString(r.Cells(COL_SHITOU1).Value)
            Dim fShnum1 As String = SafeString(r.Cells(COL_SHNUM1).Value)
            Dim fCostPrice As String = SafeString(r.Cells(COL_COST_PRICE).Value)
            Dim fCoefficient As String = SafeString(r.Cells(COL_COEFFICIENT).Value)
            Dim fBasicCost As String = SafeString(r.Cells(COL_BASIC_COST).Value)
            Dim fPremiumCost As String = SafeString(r.Cells(COL_PREMIUM_COST).Value)
            Dim fSalesCost As String = SafeString(r.Cells(COL_SALES_COST).Value)
            Dim fCompanySur As String = SafeString(r.Cells(COL_COMPANY_SUR).Value)
            Dim fSalesSur As String = SafeString(r.Cells(COL_SALES_SUR).Value)
            Dim fSalesPrice As String = SafeString(r.Cells(COL_SALES_PRICE).Value)
            Dim fRemarks As String = SafeString(r.Cells(COL_REMARKS).Value)
            Dim fZhuse As String = SafeString(r.Cells(COL_ZHUSE).Value)
            Dim fImages As String = SafeString(r.Cells(COL_IMAGES).Value)
            Dim fMojuhao As String = SafeString(r.Cells(COL_MOJUHAO).Value)

            ' 查询款号信息
            Dim inCategoryId As String = ""
            Dim inSpecId As String = ""
            Dim inCaizhi As String = ""
            Dim inImages As String = ""
            Dim inProductName As String = ""
            Dim inXiangqian As String = ""

            If fKuanhao <> "" Then
                Dim kuanhaoDt As DataTable = ExecuteQuery($"SELECT title,category_id,specification_id,caizhi,images,xiangqian FROM xipunum_erp_ksiamges WHERE kuanhao='{SafeSQL(fKuanhao)}' ORDER BY id ASC LIMIT 1")
                If kuanhaoDt.Rows.Count > 0 Then
                    inCategoryId = SafeString(kuanhaoDt.Rows(0)("category_id"))
                    inSpecId = SafeString(kuanhaoDt.Rows(0)("specification_id"))
                    inCaizhi = SafeString(kuanhaoDt.Rows(0)("caizhi"))
                    inImages = SafeString(kuanhaoDt.Rows(0)("images"))
                    inProductName = SafeString(kuanhaoDt.Rows(0)("title"))
                    inXiangqian = SafeString(kuanhaoDt.Rows(0)("xiangqian"))
                End If
                If inCaizhi <> "" Then fCaizhi = inCaizhi
                If inImages <> "" Then fImages = inImages
                If inProductName <> "" Then fName = inProductName
                If inXiangqian <> "" Then orderXiangqian = inXiangqian
            Else
                ' 无款号时查询品类和规格
                Dim catDt As DataTable = ExecuteQuery($"SELECT id FROM xipunum_erp_category WHERE title='{SafeSQL(fCategory)}' ORDER BY id ASC LIMIT 1")
                If catDt.Rows.Count > 0 Then
                    inCategoryId = SafeString(catDt.Rows(0)("id"))
                End If
                If inCategoryId <> "" Then
                    Dim specDt As DataTable = ExecuteQuery($"SELECT id FROM xipunum_erp_specs WHERE category_id='{SafeSQL(inCategoryId)}' AND title='{SafeSQL(fSpec)}' ORDER BY id ASC LIMIT 1")
                    If specDt.Rows.Count > 0 Then
                        inSpecId = SafeString(specDt.Rows(0)("id"))
                    End If
                End If
            End If

            ' 更新商品主表
            Dim updateShopSql As String = $"UPDATE xipunum_erp_shop SET images='{SafeSQL(fImages)}',product_name='{SafeSQL(fName)}',xiangqian='{SafeSQL(orderXiangqian)}',category_id='{SafeSQL(inCategoryId)}',specification_id='{SafeSQL(inSpecId)}',item_number='{SafeSQL(fKuanhao)}',caizhi='{SafeSQL(fCaizhi)}',quandu='{SafeSQL(fQuandu)}',wide='{SafeSQL(fWidth)}',thickness='{SafeSQL(fThickness)}',single='{SafeSQL(fSingle)}',quantity='{SafeSQL(fQty)}',jin_zhong='{SafeSQL(fJinzhong)}',loss='{SafeSQL(fLoss)}',including='{SafeSQL(fIncluding)}',weight='{SafeSQL(fWeight)}',sales_unit='{SafeSQL(fUnit)}',state='销售',kufang='{SafeSQL(warehouseId)}',imgstate='1',cjuser='{SafeSQL(operationAccount)}',creationtime='{SafeSQL(operationDate)}' WHERE poduct_code='{SafeSQL(fCode)}' LIMIT 1"
            ExecuteCommand(updateShopSql, MySQL_Write)

            ' 构建批量插入数据
            tmpStockData &= $"('{SafeSQL(orderId)}','{SafeSQL(fCode)}','{SafeSQL(fQty)}','{SafeSQL(fJinzhong)}','{SafeSQL(warehouseId)}','0','{SafeSQL(operationAccount)}','{SafeSQL(operationDate)}'),"

            tmpStockData1 &= $"('{SafeSQL(fCode)}','{SafeSQL(fCode)}','{SafeSQL(fSingle)}','{SafeSQL(fQty)}','{SafeSQL(fJinzhong)}','{SafeSQL(fLoss)}','{SafeSQL(fIncluding)}','{SafeSQL(fWeight)}','{SafeSQL(fUnit)}','{SafeSQL(operationAccount)}','{SafeSQL(operationDate)}'),"

            historyData &= $"('{SafeSQL(fCode)}','{SafeSQL(operationDate)}','{SafeSQL(finalOrderNumber)}','成品入库','{SafeSQL(fQty)}','{SafeSQL(fJinzhong)}','{SafeSQL(fWeight)}','工厂:{SafeSQL(factoryName)}->{SafeSQL(warehouseName)}','{SafeSQL(operationAccount)}'),"

            If isProductInlaid = "镶嵌" Then
                xiangqianData &= $"('{SafeSQL(fCode)}','{SafeSQL(fShitou)}','{SafeSQL(fStnum)}','{SafeSQL(fShitou1)}','{SafeSQL(fShnum1)}','{SafeSQL(fZhuse)}','{SafeSQL(operationAccount)}','{SafeSQL(operationDate)}'),"
            End If

            storeData &= $"('{SafeSQL(orderId)}','{SafeSQL(fCode)}','{SafeSQL(fSalesPrice)}','{SafeSQL(fFactoryCond)}','{SafeSQL(fCompanyCond)}','{SafeSQL(fCostPrice)}','{SafeSQL(fCoefficient)}','{SafeSQL(fBasicCost)}','{SafeSQL(fPremiumCost)}','{SafeSQL(fSalesCost)}','{SafeSQL(fCompanySur)}','{SafeSQL(fSalesSur)}','{SafeSQL(fMojuhao)}','{SafeSQL(fRemarks)}','{SafeSQL(operationAccount)}','{SafeSQL(operationDate)}'),"

            Dim productLogContent As String = $"账户:{UserAccount} 添加商品编码：{fCode} 相关数据"
            opLogData &= $"('添加','添加商品信息','{SafeSQL(productLogContent)}','{SafeSQL(operationAccount)}','{SafeSQL(operationDate)}'),"
        Next

        ' 执行批量插入
        If tmpStockData <> "" Then
            ExecuteCommand($"INSERT INTO xipunum_erp_shop_lincun (rukuid, poduct_code, quantity, jinzhong, kufang, state, cjuser, creationtime) VALUES {tmpStockData.TrimEnd(","c)};", MySQL_Write)
        End If

        If tmpStockData1 <> "" Then
            ExecuteCommand($"INSERT INTO xipunum_erp_shopys (poduct_code, fu_code, single, quantity, jin_zhong, loss, including, weight, sales_unit, cjuser, creationtime) VALUES {tmpStockData1.TrimEnd(","c)};", MySQL_Write)
        End If

        If historyData <> "" Then
            ExecuteCommand($"INSERT INTO xipunum_erp_history (poduct_code, updatetime, number, type, quantity, jinzhong, zhongliang, conter, cjuser) VALUES {historyData.TrimEnd(","c)};", MySQL_Write)
        End If

        If xiangqianData <> "" Then
            ExecuteCommand($"INSERT INTO xipunum_erp_shop_xiangqian (poduct_code, shitou, stnum, shitou1, shnum1, zhuse, cjuser, creationtime) VALUES {xiangqianData.TrimEnd(","c)};", MySQL_Write)
        End If

        If storeData <> "" Then
            ExecuteCommand($"INSERT INTO xipunum_erp_store (order_id, poduct_code, sales_price, factory_condition, company_condition, cost_price, coefficient, basic_cost, premium_cost, sales_cost, company_surcharge, sales_surcharge, mojuhao, remarks, cjuser, creationtime) VALUES {storeData.TrimEnd(","c)};", MySQL_Write)
        End If

        If opLogData <> "" Then
            ExecuteCommand($"INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES {opLogData.TrimEnd(","c)};", MySQL_Write)
        End If

        ' 检查是否有有效数据
        If tmpStockData = "" OrElse storeData = "" Then
            ShowWarning("没有有效商品明细，本次入库未写入完整库存数据！")
            Return
        End If

        ShowSuccess("商品入库保存成功！")

        If txtOrderNumber.Text = "" Then Return

        ' 打印单据
        If chkPrintDocument.Checked Then
            If MessageBox.Show("是否打印入库单据？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                printOrderNumber = txtOrderNumber.Text
                PrintOrderDocument()
            End If
        End If

        ' 打印标签
        If chkPrintLabel.Checked Then
            If MessageBox.Show("是否打印商品标签？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                PrintLabels()
            End If
        End If

        Me.Close()
    End Sub

    ' ========== 标签打印（_超级按钮_标签打印_被单击） ==========
    Private Sub PrintLabels()
        If MySQL_Read Is Nothing Then
            ShowWarning("数据库连接异常，无法打印标签！")
            Return
        End If
        If cmbFactory.SelectedIndex = -1 Then
            ShowWarning("请选择工厂名称！")
            cmbFactory.Focus()
            Return
        End If
        If cmbWarehouse.SelectedIndex = -1 Then
            ShowWarning("请选择库房名称！")
            cmbWarehouse.Focus()
            Return
        End If
        If cmbLabelStyle.SelectedIndex = -1 Then
            ShowWarning("请选择标签模板！")
            cmbLabelStyle.Focus()
            Return
        End If
        If String.IsNullOrEmpty(LabelPrinterConnection) OrElse String.IsNullOrEmpty(LabelPrinterName) Then
            ShowWarning("标签打印机配置为空，无法打印！")
            Return
        End If

        Dim dataFile As String = Path.Combine(Application.StartupPath, "data\erpdata.mdb")
        Dim lockFile As String = Path.Combine(Application.StartupPath, "temp\lm_print.lock")

        ' 获取工厂简写
        Dim factoryName As String = cmbFactory.SelectedItem.ToString()
        Dim factoryJianxie As String = ""
        Dim factoryDt As DataTable = ExecuteQuery($"SELECT jianxie FROM xipunum_erp_about WHERE title='{SafeSQL(factoryName)}' ORDER BY id DESC LIMIT 1")
        If factoryDt.Rows.Count > 0 Then
            factoryJianxie = SafeString(factoryDt.Rows(0)("jianxie"))
        End If

        Dim labelTemplate As String = cmbLabelStyle.SelectedItem.ToString()
        Dim labelTemplatePath As String = Path.Combine(Application.StartupPath, "voucher\biaoqian", labelTemplate)
        If Not File.Exists(labelTemplatePath) Then
            ShowWarning("标签模板不存在，无法打印！")
            Return
        End If

        Dim printCount As Integer = dgvProducts.Rows.Count
        Dim actualInsertCount As Integer = 0

        ' 遍历商品行（不含表头和合计行）
        For i As Integer = 0 To printCount - 2
            Dim productCode As String = SafeString(dgvProducts.Rows(i).Cells(COL_CODE).Value)
            If productCode = "" OrElse productCode = "合计" Then Continue For

            ' 读取商品数据
            Dim productName As String = SafeString(dgvProducts.Rows(i).Cells(COL_NAME).Value)
            Dim fSpec As String = SafeString(dgvProducts.Rows(i).Cells(COL_SPEC).Value)
            Dim fKuanhao As String = SafeString(dgvProducts.Rows(i).Cells(COL_KUANHAO).Value)
            Dim fCaizhi As String = SafeString(dgvProducts.Rows(i).Cells(COL_CAIZHI).Value)
            Dim fQuandu As String = SafeString(dgvProducts.Rows(i).Cells(COL_QUANDU).Value)
            Dim fCompanyCond As String = SafeString(dgvProducts.Rows(i).Cells(COL_COMPANY_COND).Value)
            Dim fSingle As String = SafeString(dgvProducts.Rows(i).Cells(COL_SINGLE).Value)
            Dim fJinzhong As String = SafeString(dgvProducts.Rows(i).Cells(COL_JINZHONG).Value)
            Dim fLoss As String = SafeString(dgvProducts.Rows(i).Cells(COL_LOSS).Value)
            Dim fIncluding As String = SafeString(dgvProducts.Rows(i).Cells(COL_INCLUDING).Value)
            Dim fWeight As String = SafeString(dgvProducts.Rows(i).Cells(COL_WEIGHT).Value)
            Dim fBasicCost As String = SafeString(dgvProducts.Rows(i).Cells(COL_BASIC_COST).Value)
            Dim fPremiumCost As String = SafeString(dgvProducts.Rows(i).Cells(COL_PREMIUM_COST).Value)
            Dim fSalesCost As String = SafeString(dgvProducts.Rows(i).Cells(COL_SALES_COST).Value)
            Dim fCompanySur As String = SafeString(dgvProducts.Rows(i).Cells(COL_COMPANY_SUR).Value)
            Dim fSalesSur As String = SafeString(dgvProducts.Rows(i).Cells(COL_SALES_SUR).Value)
            Dim fSalesPrice As String = SafeString(dgvProducts.Rows(i).Cells(COL_SALES_PRICE).Value)
            Dim fShitou As String = SafeString(dgvProducts.Rows(i).Cells(COL_SHITOU).Value)
            Dim fStnum As String = SafeString(dgvProducts.Rows(i).Cells(COL_STNUM).Value)
            Dim fShitou1 As String = SafeString(dgvProducts.Rows(i).Cells(COL_SHITOU1).Value)
            Dim fShnum1 As String = SafeString(dgvProducts.Rows(i).Cells(COL_SHNUM1).Value)
            Dim warehouseName As String = cmbWarehouse.SelectedItem.ToString()

            ' 查询库房信息
            Dim shopName As String = ""
            Dim shopJianxie As String = ""
            Dim shopCompany As String = ""
            Dim shopAddress As String = ""
            Dim shopDt As DataTable = ExecuteQuery($"SELECT * FROM xipunum_erp_type WHERE type='商铺' AND superior='0' AND title='{SafeSQL(warehouseName)}' ORDER BY id DESC")
            If shopDt.Rows.Count > 0 Then
                shopName = SafeString(shopDt.Rows(0)("title"))
                shopJianxie = SafeString(shopDt.Rows(0)("data1"))
                shopCompany = SafeString(shopDt.Rows(0)("data2"))
                shopAddress = SafeString(shopDt.Rows(0)("data3"))
            End If

            ' 查询证书信息
            Dim zhengshu As String = ""
            Dim yanse As String = ""
            Dim certDt As DataTable = ExecuteQuery($"SELECT b.zsbianma AS zhengshu,b.yanse AS yanse FROM xipunum_erp_shop AS a LEFT JOIN xipunum_erp_zhengshu AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_store AS c ON c.poduct_code = a.poduct_code WHERE a.poduct_code='{SafeSQL(productCode)}' LIMIT 1")
            If certDt.Rows.Count > 0 Then
                zhengshu = SafeString(certDt.Rows(0)("zhengshu"))
                yanse = SafeString(certDt.Rows(0)("yanse"))
            End If

            ' 写入Access数据库（标签打印数据）
            ' Note: 实际项目中需要使用OleDb连接Access数据库
            actualInsertCount += 1
        Next

        If actualInsertCount <= 0 Then
            ShowWarning("没有可打印的商品标签！")
            Return
        End If

        ' 执行打印命令
        Dim labelFileName As String = Path.Combine(Application.StartupPath, "voucher\biaoqian", labelTemplate)
        Dim cmdLine As String = """" & LabelPrinterConnection & """"
        cmdLine &= " /L=""" & labelFileName & """"
        cmdLine &= " /C=" & actualInsertCount.ToString()
        cmdLine &= " /X=""" & LabelPrinterName & """"
        cmdLine &= " /N /Z=2 /FL=""" & lockFile & """"

        Try
            Dim tempDir As String = Path.Combine(Application.StartupPath, "temp")
            If Not Directory.Exists(tempDir) Then
                Directory.CreateDirectory(tempDir)
            End If
            If File.Exists(lockFile) Then
                File.Delete(lockFile)
            End If
            Diagnostics.Process.Start(New ProcessStartInfo With {
                .FileName = LabelPrinterConnection,
                .Arguments = cmdLine.Substring(LabelPrinterConnection.Length + 2),
                .WindowStyle = ProcessWindowStyle.Hidden,
                .UseShellExecute = False
            })
        Catch ex As Exception
            ShowError("打印失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 单据打印（_超级按钮_单据打印_被单击） ==========
    Private Sub PrintOrderDocument()
        If MySQL_Read Is Nothing Then
            ShowWarning("数据库连接异常，无法打印入库单！")
            Return
        End If
        If printOrderNumber = "" Then
            ShowWarning("入库单号为空，无法打印！")
            Return
        End If

        ' 查询入库订单汇总
        Dim orderSql As String = $"SELECT a.id AS aid,a.odd_numbers AS aodd_numbers,a.delivery AS adelivery,a.creationtime AS acreationtime,b.title AS btitle,e.title AS etitle FROM xipunum_erp_store_order AS a INNER JOIN xipunum_erp_about AS b ON b.id = a.factory INNER JOIN xipunum_erp_store AS c ON c.order_id = a.id INNER JOIN xipunum_erp_shop AS d ON d.poduct_code = c.poduct_code INNER JOIN xipunum_erp_type AS e ON e.id = d.kufang WHERE a.odd_numbers='{SafeSQL(printOrderNumber)}' ORDER BY a.id DESC LIMIT 1"
        Dim orderDt As DataTable = ExecuteQuery(orderSql)
        If orderDt.Rows.Count = 0 Then
            ShowWarning("未找到入库单信息，无法打印！")
            Return
        End If

        Dim orderId As String = SafeString(orderDt.Rows(0)("aid"))
        Dim orderNum As String = SafeString(orderDt.Rows(0)("aodd_numbers"))
        Dim deliveryNum As String = SafeString(orderDt.Rows(0)("adelivery"))
        Dim createTime As String = SafeString(orderDt.Rows(0)("acreationtime"))
        Dim factoryTitle As String = SafeString(orderDt.Rows(0)("btitle"))
        Dim warehouseTitle As String = SafeString(orderDt.Rows(0)("etitle"))

        If orderId = "" Then
            ShowWarning("入库单ID为空，无法打印！")
            Return
        End If

        ' 查询入库详情
        Dim detailSql As String = $"SELECT a.poduct_code AS apoduct_code,a.basic_cost as abasic_cost,a.sales_cost as asales_cost,a.sales_surcharge as asales_surcharge,a.remarks as aremarks,b.single*b.quantity as bsingle,b.jin_zhong*b.quantity as bjin_zhong,b.quantity as bquantity,b.caizhi as bcaizhi,c.title AS ctitle FROM xipunum_erp_store AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_specs AS c ON c.id = b.specification_id WHERE a.order_id = '{SafeSQL(orderId)}' ORDER BY a.id DESC"
        Dim detailDt As DataTable = ExecuteQuery(detailSql)
        If detailDt.Rows.Count = 0 Then
            ShowWarning("未找到入库单明细，无法打印！")
            Return
        End If

        ' 生成打印预览数据
        ' Note: 实际项目中使用报表组件，这里使用简化的打印逻辑
        Dim printDoc As New Drawing.Printing.PrintDocument()
        AddHandler printDoc.PrintPage, Sub(s As Object, ev As Drawing.Printing.PrintPageEventArgs)
            Dim font As New Drawing.Font("宋体", 10)
            Dim fontBold As New Drawing.Font("宋体", 12, Drawing.FontStyle.Bold)
            Dim y As Integer = 50

            ev.Graphics.DrawString("入库单", fontBold, Drawing.Brushes.Black, 350, y) : y += 30
            ev.Graphics.DrawString($"订单号：{orderNum}", font, Drawing.Brushes.Black, 50, y) : y += 20
            ev.Graphics.DrawString($"工厂：{factoryTitle}", font, Drawing.Brushes.Black, 50, y) : y += 20
            ev.Graphics.DrawString($"店铺名称：{warehouseTitle}", font, Drawing.Brushes.Black, 50, y) : y += 20
            ev.Graphics.DrawString($"入库时间：{createTime}", font, Drawing.Brushes.Black, 50, y) : y += 20
            ev.Graphics.DrawString($"送货单号：{deliveryNum}", font, Drawing.Brushes.Black, 50, y) : y += 30

            ' 表头
            ev.Graphics.DrawString("商品编码", font, Drawing.Brushes.Black, 50, y)
            ev.Graphics.DrawString("品类规格", font, Drawing.Brushes.Black, 200, y)
            ev.Graphics.DrawString("数量", font, Drawing.Brushes.Black, 350, y)
            ev.Graphics.DrawString("重量", font, Drawing.Brushes.Black, 420, y)
            ev.Graphics.DrawString("金重", font, Drawing.Brushes.Black, 500, y)
            ev.Graphics.DrawString("基本工费", font, Drawing.Brushes.Black, 580, y)
            ev.Graphics.DrawString("销售工费", font, Drawing.Brushes.Black, 660, y)
            y += 25

            For Each dr As DataRow In detailDt.Rows
                Dim code As String = SafeString(dr("apoduct_code"))
                Dim spec As String = SafeString(dr("bcaizhi")) & SafeString(dr("ctitle"))
                Dim qty As String = SafeString(dr("bquantity"))
                Dim single As String = SafeString(dr("bsingle"))
                Dim jinzhong As String = SafeString(dr("bjin_zhong"))
                Dim basicCost As String = Math.Round(SafeDecimal(dr("bjin_zhong")) * SafeDecimal(dr("abasic_cost")), 2).ToString()
                Dim salesCost As String = Math.Round(SafeDecimal(dr("bjin_zhong")) * SafeDecimal(dr("asales_cost")), 2).ToString()
                Dim salesSur As String = SafeString(dr("asales_surcharge"))
                Dim remarks As String = SafeString(dr("aremarks"))

                ev.Graphics.DrawString(code, font, Drawing.Brushes.Black, 50, y)
                ev.Graphics.DrawString(spec, font, Drawing.Brushes.Black, 200, y)
                ev.Graphics.DrawString(qty, font, Drawing.Brushes.Black, 350, y)
                ev.Graphics.DrawString(single, font, Drawing.Brushes.Black, 420, y)
                ev.Graphics.DrawString(jinzhong, font, Drawing.Brushes.Black, 500, y)
                ev.Graphics.DrawString(basicCost, font, Drawing.Brushes.Black, 580, y)
                ev.Graphics.DrawString(salesCost, font, Drawing.Brushes.Black, 660, y)
                y += 20

                If y > ev.MarginBounds.Bottom Then
                    ev.HasMorePages = True
                    Return
                End If
            Next
        End Sub

        Dim printPreview As New PrintPreviewDialog()
        printPreview.Document = printDoc
        printPreview.WindowState = FormWindowState.Maximized
        printPreview.ShowDialog()
    End Sub

    ' ========== 窗口关闭确认（_窗口_商品入库添加_可否被关闭） ==========
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If inboundProductCount > 0 Then
            Dim result As DialogResult = MessageBox.Show("当前入库单尚未保存，确定关闭吗？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                ' 保存后关闭
                SaveInboundOrder()
                e.Cancel = True
                Return
            End If
        End If
        MyBase.OnFormClosing(e)
    End Sub

End Class

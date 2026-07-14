' ============================================================================
' 商品信息退货窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品信息退货.form.e.txt
' 包含所有6个程序集变量、14个子程序、所有SQL查询
' 两种模式：添加（新建退货单）、详情（查看已有退货单）
' 退货目标：从退库仓退回给工厂
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class RefundDetailForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private selectedRow As Integer = 0                    ' 集_行号
    Private selectedCol As Integer = 0                    ' 集_列号
    Private deleteBtn As New Button()                    ' 删除按钮
    Private orderSelected As Integer = -1                ' 局部_订单是否选中
    Private refundProductCode As String = ""              ' 局部_退货商品编码
    Private refundDataCode As String = ""                 ' 局部_退货商品数据编码

    ' ========== 控件声明 ==========
    Private dgvProducts As New DataGridView()             ' 高级表格1
    Private txtOrderNumber As New TextBox()               ' 单据号_编辑框
    Private txtRemarks As New TextBox()                   ' 备注_编辑框
    Private rbScan As New RadioButton()                   ' 单选框_扫码
    Private rbManual As New RadioButton()                 ' 单选框_手动
    Private txtProductCode As New TextBox()               ' 商品编码_编辑框
    Private btnAdd As New Button()                       ' 超级按钮_添加
    Private btnSave As New Button()                       ' 超级按钮_保存
    Private btnReset As New Button()                     ' 超级按钮_重置
    Private panelToolbar As New Panel()                   ' 外形框_工具条
    Private panelHeader As New Panel()                     ' 外形框_头部
    Private grpRemarks As New GroupBox()                  ' 分组框_备注
    Private lblTitle1 As New Label()                      ' 透明标签1
    Private lblTitle2 As New Label()                      ' 透明标签2

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler dgvProducts.CellClick, AddressOf DgvProducts_CellClick
        AddHandler dgvProducts.SelectionChanged, AddressOf DgvProducts_SelectionChanged
        AddHandler rbScan.CheckedChanged, AddressOf RbScan_CheckedChanged
        AddHandler rbManual.CheckedChanged, AddressOf RbManual_CheckedChanged
        AddHandler txtProductCode.KeyDown, AddressOf TxtProductCode_KeyDown
        AddHandler btnAdd.Click, AddressOf BtnAdd_Click
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnReset.Click, AddressOf BtnReset_Click
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "商品信息退货"
        Me.Size = New Drawing.Size(1427, 800)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' 工具条面板
        panelToolbar.Dock = DockStyle.Top
        panelToolbar.Height = 30
        panelToolbar.BackColor = Drawing.Color.FromArgb(240, 240, 240)
        Me.Controls.Add(panelToolbar)

        ' 添加按钮
        btnAdd.Text = "添加"
        btnAdd.Location = New Drawing.Point(8, 3)
        btnAdd.Size = New Drawing.Size(60, 24)
        panelToolbar.Controls.Add(btnAdd)

        ' 保存按钮
        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(73, 3)
        btnSave.Size = New Drawing.Size(60, 24)
        panelToolbar.Controls.Add(btnSave)

        ' 重置按钮
        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(138, 3)
        btnReset.Size = New Drawing.Size(60, 24)
        panelToolbar.Controls.Add(btnReset)

        ' 扫码单选框
        rbScan.Text = "扫码"
        rbScan.Location = New Drawing.Point(210, 5)
        rbScan.Size = New Drawing.Size(50, 24)
        rbScan.Checked = True
        panelToolbar.Controls.Add(rbScan)

        ' 手动单选框
        rbManual.Text = "手动"
        rbManual.Location = New Drawing.Point(265, 5)
        rbManual.Size = New Drawing.Size(50, 24)
        panelToolbar.Controls.Add(rbManual)

        ' 商品编码输入框
        txtProductCode.Location = New Drawing.Point(320, 5)
        txtProductCode.Size = New Drawing.Size(200, 24)
        panelToolbar.Controls.Add(txtProductCode)

        ' 头部面板
        panelHeader.Dock = DockStyle.Top
        panelHeader.Height = 65
        panelHeader.BackColor = Drawing.Color.FromArgb(248, 248, 248)
        Me.Controls.Add(panelHeader)

        ' 透明标签1（标题）
        lblTitle1.Text = "商品信息退货"
        lblTitle1.Font = New Drawing.Font("微软雅黑", 14, Drawing.FontStyle.Bold)
        lblTitle1.AutoSize = True
        lblTitle1.Location = New Drawing.Point(5, 5)
        panelHeader.Controls.Add(lblTitle1)

        ' 透明标签2（单据号标签）
        lblTitle2.Text = "单据号："
        lblTitle2.AutoSize = True
        lblTitle2.Location = New Drawing.Point(0, 12)
        panelHeader.Controls.Add(lblTitle2)

        ' 单据号编辑框
        txtOrderNumber.Location = New Drawing.Point(0, 12)
        txtOrderNumber.Size = New Drawing.Size(130, 25)
        txtOrderNumber.ReadOnly = True
        panelHeader.Controls.Add(txtOrderNumber)

        ' 商品表格
        dgvProducts.Dock = DockStyle.Fill
        dgvProducts.AllowUserToAddRows = False
        dgvProducts.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        dgvProducts.RowHeadersVisible = False
        dgvProducts.BackgroundColor = Drawing.Color.White
        dgvProducts.RowTemplate.Height = 28
        Me.Controls.Add(dgvProducts)

        ' 备注分组框
        grpRemarks.Dock = DockStyle.Bottom
        grpRemarks.Height = 85
        grpRemarks.Text = "备注"
        Me.Controls.Add(grpRemarks)

        ' 备注编辑框
        txtRemarks.Dock = DockStyle.Fill
        txtRemarks.Multiline = True
        grpRemarks.Controls.Add(txtRemarks)

        ' 删除按钮样式
        deleteBtn.Text = "删除"
        deleteBtn.BackColor = Drawing.Color.FromArgb(255, 200, 200)

        ' 调整顺序
        Me.Controls.SetChildIndex(dgvProducts, 0)
        Me.Controls.SetChildIndex(grpRemarks, 1)
        Me.Controls.SetChildIndex(panelHeader, 2)
        Me.Controls.SetChildIndex(panelToolbar, 3)
    End Sub

    ' ========== 窗口创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 获取主窗口选中的行号
        orderSelected = -1
        Try
            Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
            If mainForm IsNot Nothing AndAlso mainForm.dgvMain IsNot Nothing AndAlso mainForm.dgvMain.CurrentCell IsNot Nothing Then
                orderSelected = mainForm.dgvMain.CurrentCell.RowIndex
            End If
        Catch
        End Try

        ' 加载表头
        LoadTableHeader()
        ' 清空表格
        ClearGrid()

        If RefundOrderNumber <> "" Then
            ' 查看已有退货单（详情模式）
            txtOrderNumber.Text = RefundOrderNumber
            txtRemarks.Text = ""
            txtRemarks.ReadOnly = True
            LoadRefundDetails()
            panelToolbar.Visible = False
            panelHeader.Enabled = False
            dgvProducts.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Else
            If orderSelected = -1 Then
                ' 新建退货单
                txtOrderNumber.Text = "TH" & Now.ToString("yyyyMMdd") & "****"
                txtRemarks.Text = ""
                panelToolbar.Visible = True
                panelHeader.Enabled = True
                RbScan_CheckedChanged(Nothing, Nothing)
                dgvProducts.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
                CalculateStatistics()
            Else
                ' 从主窗口选中行查看
                txtOrderNumber.Text = RefundOrderNumber
                Try
                    Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
                    If mainForm IsNot Nothing AndAlso mainForm.dgvMain IsNot Nothing AndAlso orderSelected >= 0 Then
                        txtRemarks.Text = SafeString(mainForm.dgvMain.Rows(orderSelected).Cells(4).Value)
                    End If
                Catch
                End Try
                txtRemarks.ReadOnly = True
                LoadRefundDetails()
                panelToolbar.Visible = False
                panelHeader.Enabled = False
                dgvProducts.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            End If
        End If
    End Sub

    ' ========== 窗口尺寸改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        Dim nWidth As Integer = Me.ClientRectangle.Width
        Dim nHeight As Integer = Me.ClientRectangle.Height

        panelToolbar.Width = nWidth - 10
        panelToolbar.Left = 5
        panelToolbar.Top = 0

        panelHeader.Width = nWidth - 10
        panelHeader.Left = 5
        panelHeader.Top = 30

        dgvProducts.Width = nWidth - 10
        dgvProducts.Left = 5
        dgvProducts.Top = 95
        dgvProducts.RowTemplate.Height = 28
        dgvProducts.Height = nHeight - panelHeader.Height - 116

        grpRemarks.Left = 5
        grpRemarks.Top = nHeight - 85
        grpRemarks.Width = nWidth - 10
        txtRemarks.Width = grpRemarks.Width - 40

        lblTitle1.Width = panelHeader.Width
        lblTitle1.Top = 5
        lblTitle2.Left = panelHeader.Width - 220
        lblTitle2.Top = 12
        txtOrderNumber.Left = panelHeader.Width - 130
        txtOrderNumber.Top = 12
    End Sub

    ' ========== 加载表头 ==========
    Private Sub LoadTableHeader()
        dgvProducts.Columns.Clear()

        Dim tableHead() As String = {}
        Dim dataWidth() As Integer = {}

        If orderSelected = -1 Then
            ' 新建模式
            tableHead = {"序号", "商品编码", "商品名称", "送货单号", "半成品", "单件重", "退货数量", "退货金重", "退货重量", "工厂", "来源", "工厂成色", "备注", "操作"}
            dataWidth = {50, 120, 150, 120, 70, 100, 80, 80, 80, 80, 80, 80, 150, 60}
        Else
            ' 详情模式
            tableHead = {"序号", "商品编码", "商品名称", "送货单号", "半成品", "单件重", "退货数量", "退货金重", "退货重量", "工厂", "来源", "工厂成色", "退货时间", "备注"}
            dataWidth = {50, 120, 150, 120, 70, 100, 80, 80, 80, 80, 80, 80, 140, 150}
        End If

        For i As Integer = 0 To tableHead.Length - 1
            If i = tableHead.Length - 1 AndAlso orderSelected = -1 Then
                ' 操作列使用按钮列
                Dim btnCol As New DataGridViewButtonColumn()
                btnCol.HeaderText = tableHead(i)
                btnCol.Name = "col" & i
                btnCol.Width = dataWidth(i)
                dgvProducts.Columns.Add(btnCol)
            Else
                Dim col As New DataGridViewTextBoxColumn()
                col.HeaderText = tableHead(i)
                col.Name = "col" & i
                col.Width = dataWidth(i)
                dgvProducts.Columns.Add(col)
            End If
        Next
    End Sub

    ' ========== 清空表格 ==========
    Private Function ClearGrid() As Boolean
        dgvProducts.Rows.Clear()
        Return True
    End Function

    ' ========== 扫码模式 ==========
    Private Sub RbScan_CheckedChanged(sender As Object, e As EventArgs)
        If rbScan.Checked Then
            dgvProducts.CurrentCell = Nothing
            btnAdd.Enabled = True
        End If
    End Sub

    ' ========== 手动模式 ==========
    Private Sub RbManual_CheckedChanged(sender As Object, e As EventArgs)
        If rbManual.Checked Then
            dgvProducts.CurrentCell = Nothing
            btnAdd.Enabled = True
        End If
    End Sub

    ' ========== 重置按钮 ==========
    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        Form_Load(Nothing, Nothing)
    End Sub

    ' ========== 商品编码按下某键 ==========
    Private Sub TxtProductCode_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtProductCode.Text) Then
                ShowWarning("商品编码不能为空！")
                txtProductCode.Focus()
                Return
            End If
            BtnAdd_Click(Nothing, Nothing)
        End If
    End Sub

    ' ========== 光标位置改变 ==========
    Private Sub DgvProducts_SelectionChanged(sender As Object, e As EventArgs)
        If dgvProducts.CurrentCell IsNot Nothing Then
            selectedRow = dgvProducts.CurrentCell.RowIndex
            selectedCol = dgvProducts.CurrentCell.ColumnIndex
        End If
    End Sub

    ' ========== 退货详情被点击（详情模式） ==========
    Private Sub LoadRefundDetails()
        Dim sql As String = "SELECT a.poduct_code AS apoduct_code,c.product_name AS cproduct_name," &
            "a.delivery AS adelivery,a.half_product AS ahalf_product," &
            "CAST(ROUND(c.single, 3 ) AS DECIMAL ( 10, 3 )) AS csingle," &
            "a.quantity as aquantity,CAST(ROUND(a.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) AS ajinzhong," &
            "CAST(ROUND(a.zhongliang, 3 ) AS DECIMAL ( 10, 3 )) AS azhongliang," &
            "f.title AS ftitle,a.source AS asource,a.factory_condition AS afactory_condition," &
            "a.creationtime as acreationtime,a.remarks AS aremarks " &
            "FROM xipunum_erp_return AS a " &
            "INNER JOIN xipunum_erp_return_order AS b ON b.id = a.order_id AND b.return_umber = '" & txtOrderNumber.Text & "' " &
            "INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = a.poduct_code " &
            "INNER JOIN xipunum_erp_about AS f ON f.id = a.factory " &
            "WHERE 1 = 1 ORDER BY a.id ASC"

        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
        Dim detailCount As Integer = dt.Rows.Count

        ' 进度条
        Dim progressForm As New ProgressBarForm()
        progressForm.Show(Me)
        progressForm.LabelText = "数据正在加载中..."
        progressForm.MaxValue = detailCount

        For i As Integer = 0 To detailCount - 1
            progressForm.LabelText = "正在获取所需数据...(" & CInt((i / detailCount) * 100) & "%)"
            progressForm.Value = i + 1

            Dim row As DataRow = dt.Rows(i)
            Dim rowIdx As Integer = dgvProducts.Rows.Add()

            dgvProducts.Rows(rowIdx).Cells(0).Value = (i + 1).ToString().PadLeft(detailCount.ToString().Length, "0"c)
            dgvProducts.Rows(rowIdx).Cells(1).Value = SafeString(row("apoduct_code"))
            dgvProducts.Rows(rowIdx).Cells(2).Value = SafeString(row("cproduct_name"))
            dgvProducts.Rows(rowIdx).Cells(3).Value = SafeString(row("adelivery"))
            dgvProducts.Rows(rowIdx).Cells(4).Value = SafeString(row("ahalf_product"))
            dgvProducts.Rows(rowIdx).Cells(5).Value = SafeString(row("csingle"))
            dgvProducts.Rows(rowIdx).Cells(6).Value = SafeString(row("aquantity"))
            dgvProducts.Rows(rowIdx).Cells(7).Value = SafeString(row("ajinzhong"))
            dgvProducts.Rows(rowIdx).Cells(8).Value = SafeString(row("azhongliang"))
            dgvProducts.Rows(rowIdx).Cells(9).Value = SafeString(row("ftitle"))
            dgvProducts.Rows(rowIdx).Cells(10).Value = SafeString(row("asource"))
            dgvProducts.Rows(rowIdx).Cells(11).Value = SafeString(row("afactory_condition"))
            dgvProducts.Rows(rowIdx).Cells(12).Value = SafeString(row("acreationtime"))
            dgvProducts.Rows(rowIdx).Cells(13).Value = SafeString(row("aremarks"))
        Next

        progressForm.Close()

        ' 查询合计
        Dim sumSql As String = "SELECT sum(a.quantity) as aquantity," &
            "CAST(ROUND(sum(a.jinzhong), 3 ) AS DECIMAL ( 10, 3 )) AS ajinzhong," &
            "CAST(ROUND(sum(a.zhongliang), 3 ) AS DECIMAL ( 10, 3 )) AS azhongliang " &
            "FROM xipunum_erp_return AS a " &
            "INNER JOIN xipunum_erp_return_order AS b ON b.id = a.order_id AND b.return_umber = '" & txtOrderNumber.Text & "' " &
            "INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = a.poduct_code " &
            "INNER JOIN xipunum_erp_about AS f ON f.id = a.factory " &
            "WHERE 1 = 1 ORDER BY a.id ASC"

        Dim sumDt As DataTable = DatabaseModule.ExecuteQuery(sumSql, MySQL_Read)
        If sumDt.Rows.Count > 0 Then
            Dim sumRow As DataRow = sumDt.Rows(0)
            Dim sumQuantity As String = SafeString(sumRow("aquantity"))
            Dim sumJinzhong As String = SafeString(sumRow("ajinzhong"))
            Dim sumWeight As String = SafeString(sumRow("azhongliang"))

            ' 添加合计行
            Dim sumRowIndex As Integer = dgvProducts.Rows.Add()
            dgvProducts.Rows(sumRowIndex).Cells(1).Value = "合计"
            dgvProducts.Rows(sumRowIndex).Cells(6).Value = sumQuantity
            dgvProducts.Rows(sumRowIndex).Cells(7).Value = sumJinzhong
            dgvProducts.Rows(sumRowIndex).Cells(8).Value = sumWeight

            ' 设置只读
            For i As Integer = 0 To dgvProducts.Rows.Count - 1
                For j As Integer = 0 To dgvProducts.Columns.Count - 1
                    dgvProducts.Rows(i).Cells(j).ReadOnly = True
                Next
            Next
        End If
    End Sub

    ' ========== 添加按钮被点击 ==========
    Private Sub BtnAdd_Click(sender As Object, e As EventArgs)
        refundProductCode = ""
        refundProductCode = GBKToUTF8(txtProductCode.Text)

        ' 检查商品是否存在
        Dim sql As String = "SELECT * FROM xipunum_erp_shop where (poduct_code='" & refundProductCode & "' or fu_code='" & refundProductCode & "') order by id ASC"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
        If dt.Rows.Count = 0 Then
            ShowWarning("此商品不存在，请输入正确的商品编码！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        ' 检查商品是否在退库仓有库存
        sql = "SELECT a.poduct_code AS apoduct_code FROM xipunum_erp_shop_kucun AS a " &
            "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
            "WHERE a.kufang = '" & ReturnWarehouse & "' AND (a.quantity > 0 or a.jinzhong >0) " &
            "AND (b.poduct_code = '" & refundProductCode & "' OR b.fu_code = '" & refundProductCode & "') ORDER BY a.id DESC"
        Dim dtStock As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
        If dtStock.Rows.Count = 0 Then
            ShowWarning("此商品不属于" & ReturnWarehouseName & "！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        If dtStock.Rows.Count = 1 Then
            ' 只有一条记录，直接加载
            refundDataCode = SafeString(dtStock.Rows(0)("apoduct_code"))
            AddProductDataLoad()
        Else
            ' 多条记录，打开副编码查询窗口
            Dim fuCodeForm As New ReturnFuCodeQueryForm()
            fuCodeForm.SearchCode = refundProductCode
            If fuCodeForm.ShowDialog() = DialogResult.OK Then
                refundDataCode = fuCodeForm.SelectedProductCode
                AddProductDataLoad()
            End If
        End If
    End Sub

    ' ========== 添加数据加载 ==========
    Private Sub AddProductDataLoad()
        ' 检查商品是否已在退货清单
        For i As Integer = 0 To dgvProducts.Rows.Count - 2
            If SafeString(dgvProducts.Rows(i).Cells(1).Value) = refundDataCode Then
                ShowWarning("此商品已在当前退货清单！")
                txtProductCode.Text = ""
                txtProductCode.Focus()
                Return
            End If
        Next

        ' 查询商品信息（从退库仓库存）
        Dim sql As String = "SELECT a.poduct_code AS apoduct_code,b.product_name AS aproduct_name," &
            "g.delivery AS cdelivery,g.half_product AS chalf_product," &
            "CAST(ROUND( b.single, 3 ) AS DECIMAL ( 10, 3 )) AS asingle," &
            "a.quantity AS kucun,CAST(ROUND( a.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) AS jinzhong," &
            "b.item_number AS aitem_number," &
            "CASE WHEN COALESCE (d.lingxiao, '') = '是' THEN CAST(ROUND( a.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) " &
            "ELSE CAST(ROUND( b.single * a.quantity, 3 ) AS DECIMAL ( 10, 3 )) END AS zongzhong," &
            "h.title AS ftitle,g.source AS csource,f.factory_condition AS bfactory_condition," &
            "b.state AS astate,a.kufang AS akufang," &
            "CASE WHEN COALESCE (i.shuliang, '' ) = '' THEN '1' ELSE i.shuliang END AS plduoshu," &
            "CASE WHEN COALESCE (d.lingxiao, '' ) = '' THEN '否' ELSE d.lingxiao END AS lingxiao " &
            "FROM xipunum_erp_shop_kucun AS a " &
            "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang " &
            "LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = b.item_number AND b.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id AND b.item_number IS NOT NULL AND b.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id AND b.specification_id IS NOT NULL AND b.specification_id != '' " &
            "LEFT JOIN xipunum_erp_store AS f ON f.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_store_order AS g ON g.id = f.order_id " &
            "INNER JOIN xipunum_erp_about AS h ON h.id = g.factory " &
            "LEFT JOIN xipunum_erp_category AS i ON i.id = COALESCE ( e1.category_id, e2.category_id ) " &
            "AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL " &
            "WHERE a.kufang = '" & ReturnWarehouse & "'  AND ( a.quantity > 0 OR a.jinzhong > 0 )  " &
            "AND b.poduct_code = '" & refundDataCode & "' ORDER BY a.id DESC"

        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
        If dt.Rows.Count = 0 Then
            ShowWarning("未找到商品库存数据！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        Dim row As DataRow = dt.Rows(0)
        Dim productCode As String = SafeString(row("apoduct_code"))
        Dim productName As String = SafeString(row("aproduct_name"))
        Dim delivery As String = SafeString(row("cdelivery"))
        Dim halfProduct As String = SafeString(row("chalf_product"))
        Dim single As String = SafeString(row("asingle"))
        Dim kucun As String = SafeString(row("kucun"))
        Dim jinzhong As String = SafeString(row("jinzhong"))
        Dim zongzhong As String = SafeString(row("zongzhong"))
        Dim factory As String = SafeString(row("ftitle"))
        Dim source As String = SafeString(row("csource"))
        Dim factoryCondition As String = SafeString(row("bfactory_condition"))
        Dim state As String = SafeString(row("astate"))
        Dim akufang As String = SafeString(row("akufang"))
        Dim plduoshu As String = SafeString(row("plduoshu"))
        Dim lingxiao As String = SafeString(row("lingxiao"))

        ' 检查是否在退库仓
        If akufang <> ReturnWarehouse Then
            ShowWarning("此商品不在" & ReturnWarehouseName & ",不能退货！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        ' 零销商品检查金重
        If lingxiao = "是" Then
            If SafeDecimal(jinzhong) <= 0 Then
                ShowWarning("此商品库存金重为0,不能退货！")
                txtProductCode.Text = ""
                txtProductCode.Focus()
                Return
            End If
        End If

        ' 非零销商品检查数量
        If lingxiao = "否" Then
            If SafeDecimal(kucun) <= 0 Then
                ShowWarning("此商品库存数量为0,不能退货！")
                txtProductCode.Text = ""
                txtProductCode.Focus()
                Return
            End If
        End If

        ' 删除合计行，插入数据行
        Dim insertRowIdx As Integer = dgvProducts.Rows.Count
        If dgvProducts.Rows.Count > 0 Then
            dgvProducts.Rows.RemoveAt(dgvProducts.Rows.Count - 1)
        End If
        insertRowIdx = dgvProducts.Rows.Add()

        dgvProducts.Rows(insertRowIdx).Cells(0).Value = (insertRowIdx + 1).ToString()
        dgvProducts.Rows(insertRowIdx).Cells(1).Value = productCode
        dgvProducts.Rows(insertRowIdx).Cells(2).Value = productName
        dgvProducts.Rows(insertRowIdx).Cells(3).Value = delivery
        dgvProducts.Rows(insertRowIdx).Cells(4).Value = halfProduct
        dgvProducts.Rows(insertRowIdx).Cells(5).Value = single
        dgvProducts.Rows(insertRowIdx).Cells(6).Value = kucun
        dgvProducts.Rows(insertRowIdx).Cells(7).Value = jinzhong
        dgvProducts.Rows(insertRowIdx).Cells(8).Value = zongzhong
        dgvProducts.Rows(insertRowIdx).Cells(9).Value = factory
        dgvProducts.Rows(insertRowIdx).Cells(10).Value = source
        dgvProducts.Rows(insertRowIdx).Cells(11).Value = factoryCondition
        dgvProducts.Rows(insertRowIdx).Cells(12).Value = ""
        dgvProducts.Rows(insertRowIdx).Cells(13).Value = "删除"

        ' 设置只读和背景色
        For j As Integer = 0 To 11
            dgvProducts.Rows(insertRowIdx).Cells(j).ReadOnly = True
            dgvProducts.Rows(insertRowIdx).Cells(j).Style.BackColor = Drawing.Color.LightGray
        Next

        ' 品类多数为0时可编辑数量
        If plduoshu = "0" Then
            dgvProducts.Rows(insertRowIdx).Cells(6).ReadOnly = False
            dgvProducts.Rows(insertRowIdx).Cells(6).Style.BackColor = Drawing.Color.White
        End If

        ' 品类多数为1且零销时可编辑金重
        If plduoshu = "1" AndAlso lingxiao = "是" Then
            dgvProducts.Rows(insertRowIdx).Cells(7).ReadOnly = False
            dgvProducts.Rows(insertRowIdx).Cells(7).Style.BackColor = Drawing.Color.White
        End If

        ' 统计合计
        CalculateStatistics()

        ' 清空编码输入框
        txtProductCode.Text = ""
        txtProductCode.Focus()
    End Sub

    ' ========== 统计数据 ==========
    Private Sub CalculateStatistics()
        Dim insertRowIdx As Integer = dgvProducts.Rows.Count
        Dim sumQuantity As String = ""
        Dim sumJinzhong As String = ""
        Dim sumWeight As String = ""

        ' 累加（排除合计行）
        Dim loopCount As Integer = insertRowIdx
        If insertRowIdx > 0 AndAlso SafeString(dgvProducts.Rows(insertRowIdx - 1).Cells(1).Value) = "合计" Then
            loopCount = insertRowIdx - 1
        End If

        For i As Integer = 0 To loopCount - 1
            sumQuantity = (SafeDecimal(sumQuantity) + SafeDecimal(dgvProducts.Rows(i).Cells(6).Value)).ToString()
            sumJinzhong = (SafeDecimal(sumJinzhong) + SafeDecimal(dgvProducts.Rows(i).Cells(7).Value)).ToString()
            sumWeight = (SafeDecimal(sumWeight) + SafeDecimal(dgvProducts.Rows(i).Cells(8).Value)).ToString()
        Next

        ' 三位数处理
        sumJinzhong = FormatThreeDecimal(sumJinzhong)
        sumWeight = FormatThreeDecimal(sumWeight)

        ' 添加合计行
        Dim sumRowIndex As Integer = dgvProducts.Rows.Add()
        dgvProducts.Rows(sumRowIndex).Cells(1).Value = "合计"
        dgvProducts.Rows(sumRowIndex).Cells(6).Value = sumQuantity
        dgvProducts.Rows(sumRowIndex).Cells(7).Value = sumJinzhong
        dgvProducts.Rows(sumRowIndex).Cells(8).Value = sumWeight

        ' 设置只读和背景色
        For j As Integer = 0 To dgvProducts.Columns.Count - 1
            dgvProducts.Rows(sumRowIndex).Cells(j).ReadOnly = True
            dgvProducts.Rows(sumRowIndex).Cells(j).Style.BackColor = Drawing.Color.LightGray
        Next
    End Sub

    ' ========== 三位数处理 ==========
    Private Function FormatThreeDecimal(value As String) As String
        If String.IsNullOrEmpty(value) Then Return "0.000"
        Dim num As Decimal = SafeDecimal(value)
        Dim rounded As Decimal = Math.Round(num, 3)
        Return rounded.ToString("0.000")
    End Function

    ' ========== 表格按钮被点击（删除行） ==========
    Private Sub DgvProducts_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return

        ' 检查是否点击了操作列（最后一列）
        Dim opColIndex As Integer = dgvProducts.Columns.Count - 1
        If e.ColumnIndex <> opColIndex Then Return

        ' 不能删除合计行
        If SafeString(dgvProducts.Rows(e.RowIndex).Cells(1).Value) = "合计" Then Return

        Dim deleteRowIdx As Integer = e.RowIndex

        ' 删除当前行
        dgvProducts.Rows.RemoveAt(deleteRowIdx)
        ' 删除合计行
        If dgvProducts.Rows.Count > 0 Then
            dgvProducts.Rows.RemoveAt(dgvProducts.Rows.Count - 1)
        End If
        ' 重新编号
        For i As Integer = 0 To dgvProducts.Rows.Count - 1
            dgvProducts.Rows(i).Cells(0).Value = (i + 1).ToString()
        Next
        ' 重新统计
        CalculateStatistics()
    End Sub

    ' ========== 保存按钮 ==========
    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        InformationOperationDate = Now.ToString("yyyy-MM-dd HH:mm:ss")
        InformationOperationAccount = GBKToUTF8(UserAccount)

        Dim productCount As Integer = dgvProducts.Rows.Count - 2  ' 减去表头和合计行
        If dgvProducts.Rows.Count < 2 Then productCount = 0

        If productCount <= 0 Then
            ShowWarning("退货商品数量不能为0！")
            Return
        End If

        Dim infoRefundNumber As String = ""
        infoRefundNumber = GBKToUTF8(txtOrderNumber.Text) & CLng(DateTime.UtcNow.ToString("yyyyMMddHHmmss")) & UserAccount

        ' 检查单号是否已存在
        Dim checkSql As String = "SELECT * FROM xipunum_erp_return_order where return_umber ='" & infoRefundNumber & "'"
        Dim checkDt As DataTable = DatabaseModule.ExecuteQuery(checkSql, MySQL_Read)
        If checkDt.Rows.Count > 0 Then
            ShowWarning("当前退货单号已存在，已重新生成单号，请再次点击保存按钮！")
            txtOrderNumber.Text = "TH" & Now.ToString("yyyyMMdd") & "****"
            Return
        End If

        Dim productRefundJinzhong As String = SafeString(dgvProducts.Rows(dgvProducts.Rows.Count - 1).Cells(7).Value)
        Dim productRefundWeight As String = SafeString(dgvProducts.Rows(dgvProducts.Rows.Count - 1).Cells(8).Value)
        Dim infoRemarks As String = GBKToUTF8(txtRemarks.Text)

        ' 插入退货订单
        Dim insertOrderSql As String = "INSERT INTO xipunum_erp_return_order (return_umber, return_jinzhong, return_weight, remarks, cjuser, creationtime) VALUES (" &
            "'" & infoRefundNumber & "'," &
            "'" & productRefundJinzhong & "'," &
            "'" & productRefundWeight & "'," &
            "'" & infoRemarks & "'," &
            "'" & InformationOperationAccount & "'," &
            "'" & InformationOperationDate & "')"
        DatabaseModule.ExecuteCommand(insertOrderSql, MySQL_Write)

        ' 获取订单ID
        Dim orderDataId As String = ""
        Dim idSql As String = "SELECT id FROM xipunum_erp_return_order where return_umber='" & infoRefundNumber & "' order by id ASC LIMIT 1"
        Dim idDt As DataTable = DatabaseModule.ExecuteQuery(idSql, MySQL_Read)
        If idDt.Rows.Count > 0 Then
            orderDataId = SafeString(idDt.Rows(0)("id"))
        End If

        ' 更新单据号
        txtOrderNumber.Text = infoRefundNumber.Substring(0, 10) & Right("00000000" & orderDataId, 4)
        Dim updateNumSql As String = "UPDATE xipunum_erp_return_order SET return_umber= '" & txtOrderNumber.Text & "' WHERE id ='" & orderDataId & "' LIMIT 1"
        DatabaseModule.ExecuteCommand(updateNumSql, MySQL_Write)

        ' 系统日志
        LogSaveContent = ""
        LogSaveContent = "账户:" & UserAccount & " 增加退货订单，单号：" & txtOrderNumber.Text
        Dim logSql As String = "INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES (" &
            "'" & GBKToUTF8("添加") & "','" & GBKToUTF8("增加退货订单") & "','" & LogSaveContent & "','" & InformationOperationAccount & "','" & InformationOperationDate & "')"
        DatabaseModule.ExecuteCommand(logSql, MySQL_Write)

        ' 进度条
        Dim progressForm As New ProgressBarForm()
        progressForm.Show(Me)
        progressForm.LabelText = "数据正在保存中..."
        progressForm.MaxValue = productCount

        ' 遍历商品明细保存
        For i As Integer = 0 To productCount - 1
            progressForm.LabelText = "正在保存数据...(" & CInt((i / productCount) * 100) & "%)"
            progressForm.Value = i + 1

            Dim refundInfoCode As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(1).Value))
            Dim refundInfoQuantity As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(6).Value))
            Dim refundInfoJinzhong As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(7).Value))
            Dim refundInfoWeight As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(8).Value))
            Dim refundInfoDelivery As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(3).Value))
            Dim refundInfoHalfProduct As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(4).Value))
            Dim refundInfoFactory As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(9).Value))
            Dim refundInfoSource As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(10).Value))
            Dim refundInfoFactoryCondition As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(11).Value))
            Dim refundInfoRemarks As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(12).Value))

            ' 获取工厂id
            Dim factoryId As String = ""
            Dim factorySql As String = "SELECT id FROM xipunum_erp_about where title='" & refundInfoFactory & "' order by id ASC"
            Dim factoryDt As DataTable = DatabaseModule.ExecuteQuery(factorySql, MySQL_Read)
            If factoryDt.Rows.Count > 0 Then
                factoryId = SafeString(factoryDt.Rows(0)("id"))
            End If

            ' 获取商品总库存
            Dim stockQuantity As String = ""
            Dim stockSql As String = "SELECT sum(quantity) FROM xipunum_erp_shop_kucun where poduct_code='" & refundInfoCode & "' order by id ASC"
            Dim stockDt As DataTable = DatabaseModule.ExecuteQuery(stockSql, MySQL_Read)
            If stockDt.Rows.Count > 0 Then
                stockQuantity = SafeString(stockDt.Rows(0)(0))
            End If

            ' 插入退货明细
            Dim insertDetailSql As String = "INSERT INTO xipunum_erp_return (order_id, poduct_code, quantity, jinzhong, zhongliang, delivery, half_product, factory, source, factory_condition, remarks, cjuser, creationtime) VALUES (" &
                "'" & orderDataId & "'," &
                "'" & refundInfoCode & "'," &
                "'" & refundInfoQuantity & "'," &
                "'" & refundInfoJinzhong & "'," &
                "'" & refundInfoWeight & "'," &
                "'" & refundInfoDelivery & "'," &
                "'" & refundInfoHalfProduct & "'," &
                "'" & factoryId & "'," &
                "'" & refundInfoSource & "'," &
                "'" & refundInfoFactoryCondition & "'," &
                "'" & refundInfoRemarks & "'," &
                "'" & InformationOperationAccount & "'," &
                "'" & InformationOperationDate & "')"
            DatabaseModule.ExecuteCommand(insertDetailSql, MySQL_Write)

            ' 更新退库仓库存（扣减）
            Dim updateStockSql As String = "UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '" & refundInfoQuantity & "',jinzhong = jinzhong - '" & refundInfoJinzhong & "' WHERE poduct_code = '" & refundInfoCode & "' AND kufang='" & ReturnWarehouse & "'"
            DatabaseModule.ExecuteCommand(updateStockSql, MySQL_Write)

            ' 如果总库存为0，更新商品状态为退货
            If SafeDecimal(stockQuantity) - SafeDecimal(stockQuantity) <= 0 Then
                Dim updateStateSql As String = "UPDATE xipunum_erp_shop SET state= '" & GBKToUTF8("退货") & "',updatetime= '" & InformationOperationDate & "' WHERE poduct_code ='" & refundInfoCode & "' and kufang='0' LIMIT 1"
                DatabaseModule.ExecuteCommand(updateStateSql, MySQL_Write)
            End If

            ' 插入历史记录
            Dim historySql As String = "INSERT INTO xipunum_erp_history (poduct_code, updatetime, number, type, quantity, conter, cjuser) VALUES (" &
                "'" & refundInfoCode & "'," &
                "'" & InformationOperationDate & "'," &
                "'" & txtOrderNumber.Text & "'," &
                "'" & GBKToUTF8("成品退货") & "'," &
                "'" & refundInfoQuantity & "'," &
                "'" & GBKToUTF8("商品从原：" & ReturnWarehouseName & "-> 新：" & SafeString(dgvProducts.Rows(i).Cells(9).Value)) & "'," &
                "'" & InformationOperationAccount & "')"
            DatabaseModule.ExecuteCommand(historySql, MySQL_Write)

            ' 插入商品日志
            Dim shopLogSql As String = "INSERT INTO xipunum_erp_shop_log (poduct_code, type, creationtime) VALUES (" &
                "'" & refundInfoCode & "','" & GBKToUTF8("退货") & "','" & InformationOperationDate & "')"
            DatabaseModule.ExecuteCommand(shopLogSql, MySQL_Write)

            ' 插入系统日志
            LogSaveContent = ""
            LogSaveContent = "账户:" & UserAccount & " 退货商品，编码：" & refundInfoCode & " 从原：" & ReturnWarehouseName & "-> 新：" & SafeString(dgvProducts.Rows(i).Cells(9).Value)
            Dim sysLogSql As String = "INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES (" &
                "'" & GBKToUTF8("修改") & "','" & GBKToUTF8("修改商品参数") & "','" & LogSaveContent & "','" & InformationOperationAccount & "','" & InformationOperationDate & "')"
            DatabaseModule.ExecuteCommand(sysLogSql, MySQL_Write)

            System.Threading.Thread.Sleep(50)
        Next

        progressForm.Close()

        ShowSuccess("商品退货操作完成！")
        Me.Close()

        ' 刷新主窗口
        Try
            Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
            If mainForm IsNot Nothing Then
                mainForm.RefreshSubFolderTable()
            End If
        Catch
        End Try
    End Sub

End Class

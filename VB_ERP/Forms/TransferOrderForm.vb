' ============================================================================
' 商品信息调拨窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品信息调拨.form.e.txt
' 包含所有程序集变量、26个子程序、所有SQL查询
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Diagnostics

Public Class TransferOrderForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private transferMethod As String = "扫码"           ' 局部_商品调拨方法
    Private selectedRow As Integer = 0                   ' 集_行号
    Private selectedCol As Integer = 0                   ' 集_列号
    Private deleteBtn As New Button()                   ' 删除按钮
    Private orderSelected As Integer = -1               ' 局部_订单是否选中
    Private transferProductCode As String = ""          ' 局部_调拨商品编码
    Private transferDataCode As String = ""              ' 局部_调拨商品数据编码
    Private warehouseDataID As String = ""              ' 局部_库房数据id

    ' ========== 控件声明 ==========
    Private dgvProducts As New DataGridView()            ' 高级表格1
    Private txtOrderNumber As New TextBox()              ' 单据号_编辑框
    Private cmbToWarehouse As New ComboBox()             ' 库房名称组合框（调拨库房）
    Private cmbFromWarehouse As New ComboBox()           ' 调出库房名称组合框
    Private txtToWarehouseID As New TextBox()            ' 调拨库房_编辑框id
    Private txtFromWarehouseID As New TextBox()          ' 调出库房_编辑框id
    Private rbTransferIn As New RadioButton()            ' 单选框_调入
    Private rbTransferOut As New RadioButton()            ' 单选框_调出
    Private rbManual As New RadioButton()                ' 单选框_手动
    Private rbScan As New RadioButton()                  ' 单选框_扫码
    Private txtProductCode As New TextBox()              ' 商品编码_编辑框
    Private cmbLabelStyle As New ComboBox()              ' 组合框标签样式
    Private btnLabelPrint As New Button()                ' 超级按钮_标签打印
    Private btnSelectAll As New Button()                 ' 超级按钮_全选
    Private btnInvertSelect As New Button()              ' 超级按钮_反选
    Private btnExtractCode As New Button()                ' 超级按钮_提取编码
    Private panelToolbar As New Panel()                   ' 外形框_工具条
    Private panelHeader As New Panel()                   ' 外形框_头部
    Private lblLabelStyle As New Label()                  ' 透明标签a9
    Private lblTitle1 As New Label()                     ' 透明标签1
    Private lblTitle2 As New Label()                     ' 透明标签2
    Private lblTitle3 As New Label()                     ' 透明标签3
    Private lblTitle8 As New Label()                     ' 透明标签8
    Private btnAdd As New Button()                       ' 超级按钮_添加
    Private btnSave As New Button()                       ' 超级按钮_保存
    Private btnReset As New Button()                     ' 超级按钮_重置

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler dgvProducts.CellEndEdit, AddressOf DgvProducts_CellEndEdit
        AddHandler dgvProducts.CellClick, AddressOf DgvProducts_CellClick
        AddHandler dgvProducts.SelectionChanged, AddressOf DgvProducts_SelectionChanged
        AddHandler rbManual.CheckedChanged, AddressOf RbManual_CheckedChanged
        AddHandler rbScan.CheckedChanged, AddressOf RbScan_CheckedChanged
        AddHandler rbTransferIn.CheckedChanged, AddressOf RbTransferIn_CheckedChanged
        AddHandler rbTransferOut.CheckedChanged, AddressOf RbTransferOut_CheckedChanged
        AddHandler cmbToWarehouse.SelectedIndexChanged, AddressOf CmbToWarehouse_SelectedIndexChanged
        AddHandler cmbFromWarehouse.SelectedIndexChanged, AddressOf CmbFromWarehouse_SelectedIndexChanged
        AddHandler txtProductCode.KeyDown, AddressOf TxtProductCode_KeyDown
        AddHandler txtProductCode.TextChanged, AddressOf TxtProductCode_TextChanged
        AddHandler btnAdd.Click, AddressOf BtnAdd_Click
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnReset.Click, AddressOf BtnReset_Click
        AddHandler btnLabelPrint.Click, AddressOf BtnLabelPrint_Click
        AddHandler btnSelectAll.Click, AddressOf BtnSelectAll_Click
        AddHandler btnInvertSelect.Click, AddressOf BtnInvertSelect_Click
        AddHandler btnExtractCode.Click, AddressOf BtnExtractCode_Click
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "商品信息调拨"
        Me.Size = New Drawing.Size(1427, 664)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' 工具条面板
        panelToolbar.Dock = DockStyle.Top
        panelToolbar.Height = 35
        panelToolbar.BackColor = Drawing.Color.FromArgb(240, 240, 240)
        Me.Controls.Add(panelToolbar)

        ' 工具条按钮
        btnAdd.Text = "添加"
        btnAdd.Location = New Drawing.Point(8, 5)
        btnAdd.Size = New Drawing.Size(60, 24)
        panelToolbar.Controls.Add(btnAdd)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(73, 5)
        btnSave.Size = New Drawing.Size(60, 24)
        panelToolbar.Controls.Add(btnSave)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(138, 5)
        btnReset.Size = New Drawing.Size(60, 24)
        panelToolbar.Controls.Add(btnReset)

        ' 标签样式
        lblLabelStyle.Text = "标签样式："
        lblLabelStyle.Location = New Drawing.Point(206, 9)
        lblLabelStyle.AutoSize = True
        panelToolbar.Controls.Add(lblLabelStyle)

        cmbLabelStyle.Location = New Drawing.Point(261, 6)
        cmbLabelStyle.Size = New Drawing.Size(120, 21)
        panelToolbar.Controls.Add(cmbLabelStyle)

        btnLabelPrint.Text = "标签打印"
        btnLabelPrint.Location = New Drawing.Point(387, 5)
        btnLabelPrint.Size = New Drawing.Size(60, 24)
        panelToolbar.Controls.Add(btnLabelPrint)

        btnSelectAll.Text = "全选"
        btnSelectAll.Location = New Drawing.Point(453, 5)
        btnSelectAll.Size = New Drawing.Size(60, 24)
        panelToolbar.Controls.Add(btnSelectAll)

        btnInvertSelect.Text = "反选"
        btnInvertSelect.Location = New Drawing.Point(518, 5)
        btnInvertSelect.Size = New Drawing.Size(60, 24)
        panelToolbar.Controls.Add(btnInvertSelect)

        btnExtractCode.Text = "提取编码"
        btnExtractCode.Location = New Drawing.Point(583, 5)
        btnExtractCode.Size = New Drawing.Size(75, 24)
        panelToolbar.Controls.Add(btnExtractCode)

        ' 头部面板
        panelHeader.Dock = DockStyle.Top
        panelHeader.Height = 70
        panelHeader.BackColor = Drawing.Color.White
        Me.Controls.Add(panelHeader)

        ' 调拨类型
        lblTitle8.Text = "调入库房"
        lblTitle8.Location = New Drawing.Point(8, 8)
        lblTitle8.AutoSize = True
        panelHeader.Controls.Add(lblTitle8)

        rbTransferIn.Text = "调入"
        rbTransferIn.Location = New Drawing.Point(8, 28)
        rbTransferIn.AutoSize = True
        panelHeader.Controls.Add(rbTransferIn)

        rbTransferOut.Text = "调出"
        rbTransferOut.Location = New Drawing.Point(60, 28)
        rbTransferOut.AutoSize = True
        panelHeader.Controls.Add(rbTransferOut)

        cmbFromWarehouse.Location = New Drawing.Point(8, 48)
        cmbFromWarehouse.Size = New Drawing.Size(120, 21)
        panelHeader.Controls.Add(cmbFromWarehouse)

        txtFromWarehouseID.Location = New Drawing.Point(130, 48)
        txtFromWarehouseID.Size = New Drawing.Size(30, 21)
        txtFromWarehouseID.Visible = False
        panelHeader.Controls.Add(txtFromWarehouseID)

        ' 调拨库房
        lblTitle3.Text = "调出库房"
        lblTitle3.Location = New Drawing.Point(170, 8)
        lblTitle3.AutoSize = True
        panelHeader.Controls.Add(lblTitle3)

        cmbToWarehouse.Location = New Drawing.Point(170, 48)
        cmbToWarehouse.Size = New Drawing.Size(120, 21)
        panelHeader.Controls.Add(cmbToWarehouse)

        txtToWarehouseID.Location = New Drawing.Point(292, 48)
        txtToWarehouseID.Size = New Drawing.Size(30, 21)
        txtToWarehouseID.Visible = False
        panelHeader.Controls.Add(txtToWarehouseID)

        ' 输入模式
        Dim lblInput As New Label()
        lblInput.Text = "输入："
        lblInput.Location = New Drawing.Point(330, 8)
        lblInput.AutoSize = True
        panelHeader.Controls.Add(lblInput)

        rbScan.Text = "扫码"
        rbScan.Location = New Drawing.Point(330, 28)
        rbScan.AutoSize = True
        panelHeader.Controls.Add(rbScan)

        rbManual.Text = "手动"
        rbManual.Location = New Drawing.Point(382, 28)
        rbManual.AutoSize = True
        panelHeader.Controls.Add(rbManual)

        ' 商品编码
        Dim lblCode As New Label()
        lblCode.Text = "商品编码："
        lblCode.Location = New Drawing.Point(440, 8)
        lblCode.AutoSize = True
        panelHeader.Controls.Add(lblCode)

        txtProductCode.Location = New Drawing.Point(440, 28)
        txtProductCode.Size = New Drawing.Size(200, 21)
        panelHeader.Controls.Add(txtProductCode)

        ' 单据号
        lblTitle2.Text = "单据号："
        lblTitle2.Location = New Drawing.Point(660, 8)
        lblTitle2.AutoSize = True
        panelHeader.Controls.Add(lblTitle2)

        txtOrderNumber.Location = New Drawing.Point(660, 28)
        txtOrderNumber.Size = New Drawing.Size(150, 21)
        txtOrderNumber.ReadOnly = True
        panelHeader.Controls.Add(txtOrderNumber)

        ' 商品表格
        dgvProducts.Dock = DockStyle.Fill
        dgvProducts.AllowUserToAddRows = False
        dgvProducts.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        dgvProducts.BackgroundColor = Drawing.Color.White
        dgvProducts.RowHeadersVisible = False
        dgvProducts.AllowUserToResizeRows = False
        Me.Controls.Add(dgvProducts)
        dgvProducts.BringToFront()
    End Sub

    ' ========== 窗口加载（_窗口_商品信息调拨_创建完毕） ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 获取主窗口选中行号
        Try
            Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
            If mainForm IsNot Nothing AndAlso mainForm.dgvMain.CurrentCell IsNot Nothing Then
                orderSelected = mainForm.dgvMain.CurrentCell.RowIndex
            Else
                orderSelected = -1
            End If
        Catch
            orderSelected = -1
        End Try

        ' 加载表头
        LoadTableHeader()

        ' 清空表格
        ClearGrid()

        ' 设置只读
        dgvProducts.ReadOnly = True

        If TransferOrderNumber <> "" Then
            ' 查看已有调拨单
            cmbLabelStyle.Visible = True
            btnLabelPrint.Visible = True
            lblLabelStyle.Visible = True
            btnSelectAll.Visible = True
            btnInvertSelect.Visible = True
            btnExtractCode.Visible = True

            txtOrderNumber.Text = TransferOrderNumber
            rbTransferIn.Checked = True
            rbTransferOut.Checked = False
            dgvProducts.Columns(0).HeaderText = ""

            ' 加载调拨详情
            LoadTransferDetails()

            panelToolbar.Visible = False
            panelHeader.Enabled = False
            cmbToWarehouse.Visible = False
            cmbFromWarehouse.Visible = False
        Else
            If orderSelected = -1 Then
                ' 新建调拨单
                cmbLabelStyle.Visible = False
                btnLabelPrint.Visible = False
                lblLabelStyle.Visible = False
                btnSelectAll.Visible = False
                btnInvertSelect.Visible = False
                btnExtractCode.Visible = False

                cmbLabelStyle.Items.Clear()

                ' 生成单据号 DByyyyMMdd****
                txtOrderNumber.Text = "DB" & DateTime.Now.ToString("yyyyMMdd") & "****"
                panelToolbar.Visible = True
                panelHeader.Enabled = True
                cmbToWarehouse.Visible = True
                cmbFromWarehouse.Visible = True
                cmbToWarehouse.Enabled = True
                dgvProducts.ReadOnly = False

                ' 加载基础参数
                LoadBaseParameters()

                rbManual.Checked = False
                rbScan.Checked = True
                RbScan_CheckedChanged(Nothing, Nothing)
                RbTransferOut_CheckedChanged(Nothing, Nothing)
                CalculateStatistics()
            Else
                ' 从主窗口选中行查看
                cmbLabelStyle.Visible = True
                btnLabelPrint.Visible = True
                lblLabelStyle.Visible = True
                btnSelectAll.Visible = True
                btnInvertSelect.Visible = True
                btnExtractCode.Visible = True

                cmbLabelStyle.Items.Clear()

                txtOrderNumber.Text = TransferOrderNumber
                dgvProducts.ReadOnly = False
                LoadTransferDetails()
                panelToolbar.Visible = False
                panelHeader.Enabled = True

                ' 根据主窗口数据设置调入/调出
                Try
                    Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
                    If mainForm IsNot Nothing Then
                        Dim cellValue As String = SafeString(mainForm.dgvMain.Rows(orderSelected).Cells(2).Value)
                        If cellValue = rbTransferIn.Text Then
                            rbTransferIn.Checked = True
                            rbTransferOut.Checked = False
                        ElseIf cellValue = rbTransferOut.Text Then
                            rbTransferIn.Checked = False
                            rbTransferOut.Checked = True
                        End If
                    End If
                Catch
                End Try

                cmbToWarehouse.Visible = False
                cmbFromWarehouse.Visible = False
            End If
        End If

        ' 加载标签样式文件
        cmbLabelStyle.Items.Clear()
        Try
            Dim labelDir As String = Path.Combine(Application.StartupPath, "voucher\biaoqian\")
            If Directory.Exists(labelDir) Then
                Dim files() As String = Directory.GetFiles(labelDir, "*.qdf")
                For Each f As String In files
                    cmbLabelStyle.Items.Add(Path.GetFileName(f))
                Next
                If cmbLabelStyle.Items.Count > 0 Then
                    cmbLabelStyle.SelectedIndex = 0
                End If
            End If
        Catch
        End Try
    End Sub

    ' ========== 窗口尺寸改变（_窗口_商品信息调拨_尺寸被改变） ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        Dim nWidth As Integer = Me.ClientSize.Width
        Dim nHeight As Integer = Me.ClientSize.Height

        panelToolbar.Width = nWidth - 10
        panelToolbar.Left = 5
        panelToolbar.Top = 0

        panelHeader.Width = nWidth - 10
        panelHeader.Left = 5
        panelHeader.Top = 30

        dgvProducts.Width = nWidth - 10
        dgvProducts.Left = 5
        dgvProducts.Top = 105
        dgvProducts.RowTemplate.Height = 28
        dgvProducts.Height = nHeight - panelHeader.Height - 41
    End Sub

    ' ========== 手动模式（_单选框_手动_被单击） ==========
    Private Sub RbManual_CheckedChanged(sender As Object, e As EventArgs)
        If rbManual.Checked Then
            transferMethod = "添加"
            dgvProducts.CurrentCell = Nothing
            btnAdd.Enabled = True
        End If
    End Sub

    ' ========== 扫码模式（_单选框_扫码_被单击） ==========
    Private Sub RbScan_CheckedChanged(sender As Object, e As EventArgs)
        If rbScan.Checked Then
            transferMethod = "扫码"
            dgvProducts.CurrentCell = Nothing
            btnAdd.Enabled = False
        End If
    End Sub

    ' ========== 加载基础参数（_调拨基础参数_被单击） ==========
    Private Sub LoadBaseParameters()
        cmbToWarehouse.Items.Clear()
        cmbToWarehouse.Text = "请选择库房"
        cmbToWarehouse.Items.Add("总库")

        If AccountType = "后台" Then
            cmbFromWarehouse.Items.Clear()
            cmbFromWarehouse.Text = "请选择库房"
            cmbFromWarehouse.Items.Add("总库")
            txtFromWarehouseID.Text = "0"
            cmbFromWarehouse.Enabled = True
            rbTransferIn.Enabled = True
            rbTransferOut.Enabled = True
        Else
            cmbFromWarehouse.Items.Clear()
            cmbFromWarehouse.Text = UserDepartmentName
            txtFromWarehouseID.Text = UserDepartment
            cmbFromWarehouse.Enabled = False
            rbTransferIn.Enabled = False
            rbTransferOut.Enabled = False
        End If

        rbTransferIn.Checked = False
        rbTransferOut.Checked = True

        ' 查询商铺列表
        Dim sql As String = "SELECT title FROM xipunum_erp_type WHERE type= '商铺' and superior ='0' order by id desc"
        Dim dt As DataTable = ExecuteQuery(sql)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim shopName As String = SafeString(dt.Rows(i)("title"))
            cmbToWarehouse.Items.Add(shopName)
            If AccountType = "后台" Then
                cmbFromWarehouse.Items.Add(shopName)
                cmbFromWarehouse.SelectedIndex = 0
            End If
        Next

        txtToWarehouseID.Text = ""
    End Sub

    ' ========== 重置（_超级按钮_重置_被单击） ==========
    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        Form_Load(Nothing, Nothing)
    End Sub

    ' ========== 商品编码按键（_商品编码_编辑框_按下某键） ==========
    Private Sub TxtProductCode_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If txtProductCode.Text = "" Then
                ShowWarning("商品编码不能为空！")
                txtProductCode.Focus()
                Return
            End If
            BtnAdd_Click(Nothing, Nothing)
            Return
        End If
    End Sub

    ' ========== 光标位置改变（_高级表格1_光标位置改变） ==========
    Private Sub DgvProducts_SelectionChanged(sender As Object, e As EventArgs)
        If dgvProducts.CurrentCell IsNot Nothing Then
            selectedRow = dgvProducts.CurrentCell.RowIndex
            selectedCol = dgvProducts.CurrentCell.ColumnIndex
        End If
    End Sub

    ' ========== 清空表格（子程序_删除表格） ==========
    Private Sub ClearGrid()
        dgvProducts.Rows.Clear()
        dgvProducts.Columns.Clear()
    End Sub

    ' ========== 加载表头（_高级表格1_加载表头） ==========
    Private Sub LoadTableHeader()
        dgvProducts.Columns.Clear()

        Dim headers() As String
        Dim widths() As Integer

        If orderSelected = -1 Then
            headers = {"序号", "商品编码", "品类", "规格", "商品名称", "款号", "单件重", "调拨数量", "调拨金重", "原库房", "新库房", "库存数量", "库存重量", "操作"}
            widths = {50, 120, 100, 100, 200, 120, 80, 80, 80, 80, 80, 78, 78, 60}
        Else
            headers = {"序号", "商品编码", "品类", "规格", "商品名称", "款号", "单件重", "调拨数量", "调拨金重", "原库房", "新库房", "调拨时间", "备注"}
            widths = {50, 120, 100, 100, 200, 120, 80, 80, 80, 80, 80, 140, 150}
        End If

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            col.SortMode = DataGridViewColumnSortMode.NotSortable
            dgvProducts.Columns.Add(col)
        Next

        If orderSelected = -1 Then
            ' 添加操作列按钮
            Dim btnCol As New DataGridViewButtonColumn()
            btnCol.HeaderText = "操作"
            btnCol.Name = "colAction"
            btnCol.Width = 60
            btnCol.Text = "删除"
            btnCol.UseColumnTextForButtonValue = True
            dgvProducts.Columns.Add(btnCol)
            ' 删除文本列的操作列
            dgvProducts.Columns.RemoveAt(13)
            dgvProducts.Columns.Add(btnCol)
        End If
    End Sub

    ' ========== 加载调拨详情（_调拨详情_被点击） ==========
    Private Sub LoadTransferDetails()
        Dim sql As String = "SELECT a.poduct_code AS apoduct_code,d.title AS pinlei,e.title AS guige,c.product_name AS cproduct_name,c.item_number AS citem_number," &
            "CAST(ROUND( c.single,3) AS DECIMAL (10,3)) AS csingle,a.quantity AS aquantity," &
            "CAST(ROUND( a.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) AS zhongliang," &
            "CASE WHEN a.ykufang = '0' THEN '总库' ELSE f.title END AS yuanku," &
            "CASE WHEN a.xkufang = '0' THEN '总库' ELSE g.title END AS xinku," &
            "a.creationtime AS acreationtime,a.remarks AS aremarks " &
            "FROM xipunum_erp_transfer AS a " &
            "INNER JOIN xipunum_erp_transfer_order AS b ON b.id = a.order_id AND b.transfer_umber = '" & txtOrderNumber.Text & "' " &
            "INNER JOIN (SELECT * FROM xipunum_erp_shop GROUP BY poduct_code) AS c ON c.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_category AS d ON d.id = c.category_id " &
            "LEFT JOIN xipunum_erp_specs AS e ON e.id = c.specification_id " &
            "left JOIN xipunum_erp_type AS f ON f.id = a.ykufang AND a.ykufang != '0' " &
            "left JOIN xipunum_erp_type AS g ON g.id = a.xkufang AND a.xkufang != '0' " &
            "WHERE 1 = 1 ORDER BY a.id asc"

        Dim dt As DataTable = ExecuteQuery(sql)
        Dim detailCount As Integer = dt.Rows.Count

        ' 显示进度条
        Dim progressForm As ProgressBarForm = Nothing
        If detailCount > 0 Then
            progressForm = New ProgressBarForm("数据正在加载中...")
            progressForm.Show(Me)
            progressForm.SetProgress(0, "数据正在加载中...")
        End If

        For i As Integer = 0 To detailCount - 1
            Dim pct As Integer = CInt(Math.Round(i / detailCount * 100))
            progressForm?.SetProgress(pct, $"正在获取所需数据...({pct}%)")

            dgvProducts.Rows.Add()
            Dim rowIdx As Integer = dgvProducts.Rows.Count - 1

            If TransferOrderNumber <> "" Then
                dgvProducts.Rows(rowIdx).Cells(0).Value = False
            Else
                If orderSelected = -1 Then
                    dgvProducts.Rows(rowIdx).Cells(0).Value = (i + 1).ToString().PadLeft(dgvProducts.Rows.Count.ToString().Length, "0"c)
                Else
                    dgvProducts.Rows(rowIdx).Cells(0).Value = False
                End If
            End If

            dgvProducts.Rows(rowIdx).Cells(1).Value = SafeString(dt.Rows(i)("apoduct_code"))
            dgvProducts.Rows(rowIdx).Cells(2).Value = SafeString(dt.Rows(i)("pinlei"))
            dgvProducts.Rows(rowIdx).Cells(3).Value = SafeString(dt.Rows(i)("guige"))
            dgvProducts.Rows(rowIdx).Cells(4).Value = SafeString(dt.Rows(i)("cproduct_name"))
            dgvProducts.Rows(rowIdx).Cells(5).Value = SafeString(dt.Rows(i)("citem_number"))
            dgvProducts.Rows(rowIdx).Cells(6).Value = SafeString(dt.Rows(i)("csingle"))
            dgvProducts.Rows(rowIdx).Cells(7).Value = SafeString(dt.Rows(i)("aquantity"))
            dgvProducts.Rows(rowIdx).Cells(8).Value = SafeString(dt.Rows(i)("zhongliang"))
            dgvProducts.Rows(rowIdx).Cells(9).Value = SafeString(dt.Rows(i)("yuanku"))
            dgvProducts.Rows(rowIdx).Cells(10).Value = SafeString(dt.Rows(i)("xinku"))
            dgvProducts.Rows(rowIdx).Cells(11).Value = SafeString(dt.Rows(i)("acreationtime"))
            dgvProducts.Rows(rowIdx).Cells(12).Value = SafeString(dt.Rows(i)("aremarks"))
        Next

        progressForm?.Close()

        ' 查询合计
        Dim sumSql As String = "SELECT sum(a.quantity) AS aquantity,CAST(ROUND( sum(a.jinzhong), 3 ) AS DECIMAL ( 10, 3 )) AS zhongliang " &
            "FROM xipunum_erp_transfer AS a " &
            "INNER JOIN xipunum_erp_transfer_order AS b ON b.id = a.order_id AND b.transfer_umber = '" & txtOrderNumber.Text & "' " &
            "INNER JOIN (SELECT * FROM xipunum_erp_shop GROUP BY poduct_code) AS c ON c.poduct_code = a.poduct_code " &
            "WHERE 1 = 1 ORDER BY a.id DESC"
        Dim sumDt As DataTable = ExecuteQuery(sumSql)
        Dim sumQty As String = ""
        Dim sumWeight As String = ""
        If sumDt.Rows.Count > 0 Then
            sumQty = SafeString(sumDt.Rows(0)("aquantity"))
            sumWeight = SafeString(sumDt.Rows(0)("zhongliang"))
        End If

        ' 添加合计行
        dgvProducts.Rows.Add()
        Dim totalRow As Integer = dgvProducts.Rows.Count - 1
        dgvProducts.Rows(totalRow).Cells(1).Value = "合计"
        dgvProducts.Rows(totalRow).Cells(7).Value = sumQty
        dgvProducts.Rows(totalRow).Cells(8).Value = sumWeight

        ' 设置只读
        For i As Integer = 0 To totalRow
            For j As Integer = 1 To 12
                dgvProducts.Rows(i).Cells(j).ReadOnly = True
            Next
        Next
    End Sub

    ' ========== 商品编码内容改变（_商品编码_编辑框_内容被改变） ==========
    Private Sub TxtProductCode_TextChanged(sender As Object, e As EventArgs)
        If transferMethod = "扫码" Then
            Dim code As String = txtProductCode.Text
            If code.Length > 0 AndAlso code.Substring(0, code.Length - 1) <> "T" Then
                ' Actually check if left part (all except last char) != "T"
                If code.Length > 0 Then
                    Dim leftPart As String = code.Substring(0, Math.Max(0, code.Length - 1))
                    If leftPart <> "T" AndAlso code.Length >= 12 Then
                        BtnAdd_Click(Nothing, Nothing)
                        Return
                    End If
                End If
            End If
        End If
    End Sub

    ' ========== 调入单选（_单选框_调入_被单击） ==========
    Private Sub RbTransferIn_CheckedChanged(sender As Object, e As EventArgs)
        If rbTransferIn.Checked Then
            lblTitle8.Text = "调出库房"
            lblTitle3.Text = "调入库房"
        End If
    End Sub

    ' ========== 调出单选（_单选框_调出_被单击） ==========
    Private Sub RbTransferOut_CheckedChanged(sender As Object, e As EventArgs)
        If rbTransferOut.Checked Then
            lblTitle8.Text = "调入库房"
            lblTitle3.Text = "调出库房"
        End If
    End Sub

    ' ========== 调拨库房选择（_库房名称组合框_列表项被选择） ==========
    Private Sub CmbToWarehouse_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmbToWarehouse.SelectedIndex < 0 Then Return
        Dim selectedText As String = cmbToWarehouse.SelectedItem.ToString()
        txtToWarehouseID.Text = ""

        If selectedText = "总库" Then
            txtToWarehouseID.Text = "0"
        Else
            Dim sql As String = "SELECT id FROM xipunum_erp_type where type='商铺' and title='" & selectedText & "' and superior='0' order by id ASC"
            Dim dt As DataTable = ExecuteQuery(sql)
            If dt.Rows.Count > 0 Then
                txtToWarehouseID.Text = SafeString(dt.Rows(0)("id"))
            End If
        End If
    End Sub

    ' ========== 调出库房选择（_调出库房名称组合框_列表项被选择） ==========
    Private Sub CmbFromWarehouse_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmbFromWarehouse.SelectedIndex < 0 Then Return
        Dim selectedText As String = cmbFromWarehouse.SelectedItem.ToString()
        txtFromWarehouseID.Text = ""

        If selectedText = "总库" Then
            txtFromWarehouseID.Text = "0"
        Else
            Dim sql As String = "SELECT id FROM xipunum_erp_type where type='商铺' and title='" & selectedText & "' and superior='0' order by id ASC"
            Dim dt As DataTable = ExecuteQuery(sql)
            If dt.Rows.Count > 0 Then
                txtFromWarehouseID.Text = SafeString(dt.Rows(0)("id"))
            End If
        End If
    End Sub

    ' ========== 添加商品（_超级按钮_添加_被单击） ==========
    Private Sub BtnAdd_Click(sender As Object, e As EventArgs)
        ' 验证库房选择
        If cmbToWarehouse.SelectedIndex = -1 Then
            ShowWarning("请选择库房名称！")
            txtProductCode.Text = ""
            cmbToWarehouse.Focus()
            Return
        End If

        If txtFromWarehouseID.Text = "" Then
            ShowWarning("请选择库房名称！")
            txtProductCode.Text = ""
            cmbToWarehouse.Focus()
            Return
        End If

        transferProductCode = txtProductCode.Text

        ' 检查商品是否存在
        Dim checkSql As String = "SELECT * FROM xipunum_erp_shop where (poduct_code='" & transferProductCode & "' or fu_code='" & transferProductCode & "') order by id ASC"
        Dim checkDt As DataTable = ExecuteQuery(checkSql)
        If checkDt.Rows.Count = 0 Then
            ShowWarning("此商品不存在，请输入正确的商品编码！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        ' 确定库房数据id
        If rbTransferIn.Checked Then
            warehouseDataID = txtToWarehouseID.Text
        End If
        If rbTransferOut.Checked Then
            warehouseDataID = txtFromWarehouseID.Text
        End If

        ' 检查商品是否属于所选调出店铺
        Dim belongSql As String = "SELECT a.poduct_code AS apoduct_code FROM xipunum_erp_shop_kucun AS a " &
            "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
            "WHERE a.kufang = '" & warehouseDataID & "' AND (a.quantity > 0 or a.jinzhong >0) " &
            "AND (b.poduct_code = '" & transferProductCode & "' OR b.fu_code = '" & transferProductCode & "') ORDER BY a.id DESC"
        Dim belongDt As DataTable = ExecuteQuery(belongSql)

        If belongDt.Rows.Count = 0 Then
            ShowWarning("此商品不属于所选调出店铺！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        If belongDt.Rows.Count = 1 Then
            transferDataCode = SafeString(belongDt.Rows(0)("apoduct_code"))
            AddProductDataLoad()
            Return
        Else
            ' 多个匹配时打开副编码查询窗口
            Dim fuCodeForm As New TransferFuCodeQueryForm(warehouseDataID)
            If fuCodeForm.ShowDialog() = DialogResult.OK Then
                transferDataCode = fuCodeForm.SelectedProductCode
                AddProductDataLoad()
            End If
            Return
        End If
    End Sub

    ' ========== 添加商品数据加载（_超级按钮_添加_数据加载） ==========
    Private Sub AddProductDataLoad()
        ' 检查是否已在调拨清单中
        For i As Integer = 0 To dgvProducts.Rows.Count - 2
            If SafeString(dgvProducts.Rows(i).Cells(1).Value) = transferDataCode Then
                ShowWarning("此商品已在当前调拨清单！")
                txtProductCode.Text = ""
                cmbToWarehouse.Focus()
                Return
            End If
        Next

        ' 获取调出库房名称
        Dim writeWarehouseState As String = ""
        If txtFromWarehouseID.Text = "0" Then
            writeWarehouseState = "总库"
        Else
            Dim whSql As String = "SELECT * FROM xipunum_erp_type where id='" & txtFromWarehouseID.Text & "' order by id ASC LIMIT 1"
            Dim whDt As DataTable = ExecuteQuery(whSql)
            If whDt.Rows.Count > 0 Then
                writeWarehouseState = SafeString(whDt.Rows(0)("title"))
            End If
        End If

        ' 获取库房数据名称
        Dim warehouseName As String = ""
        If AccountType = "后台" Then
            If rbTransferIn.Checked Then
                warehouseName = cmbFromWarehouse.SelectedItem?.ToString()
            End If
            If rbTransferOut.Checked Then
                warehouseName = cmbToWarehouse.SelectedItem?.ToString()
            End If
        Else
            If rbTransferIn.Checked Then
                warehouseName = cmbFromWarehouse.Text
            End If
            If rbTransferOut.Checked Then
                warehouseName = cmbToWarehouse.SelectedItem?.ToString()
            End If
        End If

        Dim transferWhId As String = txtToWarehouseID.Text
        Dim fromWhId As String = txtFromWarehouseID.Text
        Dim kufangDataId As String = warehouseDataID

        rbTransferIn.Enabled = True
        rbTransferOut.Enabled = True
        cmbFromWarehouse.Enabled = True

        ' 查询商品库存信息
        Dim sql As String = "SELECT CASE WHEN COALESCE(f.shuliang, '') = '' THEN '1' ELSE f.shuliang END AS plduoshu," &
            "CASE WHEN COALESCE(e1.shuliang, e2.shuliang, '') = '' THEN '1' ELSE COALESCE(e1.shuliang, e2.shuliang) END AS ggduoshu," &
            "CASE WHEN COALESCE(d.lingxiao, '') = '' THEN '否' ELSE d.lingxiao END AS lingxiao," &
            "a.poduct_code AS apoduct_code," &
            "CASE WHEN COALESCE(f.title, '') = '' THEN '未匹配' ELSE f.title END AS pinlei," &
            "COALESCE(e1.title, e2.title, '无数据') AS guige," &
            "b.product_name AS aproduct_name,b.item_number AS aitem_number," &
            "CAST(ROUND( b.single, 3 ) AS DECIMAL ( 10, 3 )) AS asingle," &
            "a.quantity AS kucun,CAST(ROUND( a.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) AS zongzhong," &
            "CASE WHEN a.kufang = '0' THEN '总库' ELSE c.title END AS yuanku," &
            "b.state AS astate,a.kufang AS akufang " &
            "FROM xipunum_erp_shop_kucun AS a " &
            "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang " &
            "LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = b.item_number AND b.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id AND b.item_number IS NOT NULL AND b.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id AND b.specification_id IS NOT NULL AND b.specification_id != '' " &
            "LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE ( e1.category_id, e2.category_id ) " &
            "AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL " &
            "WHERE a.kufang = '" & kufangDataId & "' AND (a.quantity > 0 OR a.jinzhong > 0) " &
            "AND b.poduct_code = '" & transferDataCode & "' ORDER BY a.id DESC"

        Dim dt As DataTable = ExecuteQuery(sql)
        If dt.Rows.Count = 0 Then
            ShowWarning("未找到商品库存数据！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        Dim data As DataRow = dt.Rows(0)
        Dim productCode As String = SafeString(data("apoduct_code"))
        Dim pinlei As String = SafeString(data("pinlei"))
        Dim guige As String = SafeString(data("guige"))
        Dim productName As String = SafeString(data("aproduct_name"))
        Dim itemNumber As String = SafeString(data("aitem_number"))
        Dim single As String = SafeString(data("asingle"))
        Dim kucun As String = SafeString(data("kucun"))
        Dim zongzhong As String = SafeString(data("zongzhong"))
        Dim yuanku As String = SafeString(data("yuanku"))
        Dim astate As String = SafeString(data("astate"))
        Dim akufang As String = SafeString(data("akufang"))
        Dim plduoshu As String = SafeString(data("plduoshu"))
        Dim lingxiao As String = SafeString(data("lingxiao"))

        ' 零销商品检查
        If lingxiao = "是" Then
            If SafeDecimal(zongzhong) <= 0 Then
                ShowWarning("此商品库存金重为0,不能调拨！")
                txtProductCode.Text = ""
                txtProductCode.Focus()
                Return
            End If
        End If

        If lingxiao = "否" Then
            If SafeDecimal(kucun) <= 0 Then
                ShowWarning("此商品库存数量为0,不能调拨！")
                txtProductCode.Text = ""
                txtProductCode.Focus()
                Return
            End If
        End If

        ' 商品状态检查
        If astate <> "销售" Then
            ShowWarning("此商品状态为:""" & astate & """,无法调拨！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        ' 调入检查
        If rbTransferIn.Checked Then
            If akufang = fromWhId Then
                If AccountType <> "后台" Then
                    ShowWarning("此商品已在:""" & cmbToWarehouse.Text & """库房,无法调拨！")
                Else
                    ShowWarning("此商品已在:""" & cmbFromWarehouse.SelectedItem?.ToString() & """库房,无法调拨！")
                End If
                txtProductCode.Text = ""
                txtProductCode.Focus()
                Return
            End If
        End If

        ' 调出检查
        If rbTransferOut.Checked Then
            If akufang = transferWhId Then
                MessageBox.Show("此商品已在:""" & cmbToWarehouse.SelectedItem?.ToString() & """库房,无法调拨！")
                txtProductCode.Text = ""
                txtProductCode.Focus()
                Return
            End If
        End If

        ' 添加到表格
        Dim insertRow As Integer = dgvProducts.Rows.Count
        ' 删除合计行
        If dgvProducts.Rows.Count > 0 Then
            dgvProducts.Rows.RemoveAt(dgvProducts.Rows.Count - 1)
        End If
        dgvProducts.Rows.Add()
        Dim newIdx As Integer = dgvProducts.Rows.Count - 1

        dgvProducts.Rows(newIdx).Cells(0).Value = (newIdx).ToString()
        dgvProducts.Rows(newIdx).Cells(1).Value = productCode
        dgvProducts.Rows(newIdx).Cells(2).Value = pinlei
        dgvProducts.Rows(newIdx).Cells(3).Value = guige
        dgvProducts.Rows(newIdx).Cells(4).Value = productName
        dgvProducts.Rows(newIdx).Cells(5).Value = itemNumber
        dgvProducts.Rows(newIdx).Cells(6).Value = single
        dgvProducts.Rows(newIdx).Cells(7).Value = kucun
        dgvProducts.Rows(newIdx).Cells(8).Value = zongzhong
        dgvProducts.Rows(newIdx).Cells(9).Value = yuanku
        dgvProducts.Rows(newIdx).Cells(10).Value = warehouseName
        dgvProducts.Rows(newIdx).Cells(11).Value = kucun
        dgvProducts.Rows(newIdx).Cells(12).Value = zongzhong

        ' 设置只读和背景色
        For j As Integer = 0 To 12
            dgvProducts.Rows(newIdx).Cells(j).ReadOnly = True
            dgvProducts.Rows(newIdx).Cells(j).Style.BackColor = Drawing.Color.LightGray
        Next

        ' 品类多数为0时允许编辑数量
        If plduoshu = "0" Then
            dgvProducts.Rows(newIdx).Cells(7).ReadOnly = False
            dgvProducts.Rows(newIdx).Cells(7).Style.BackColor = Drawing.Color.White
        End If

        ' 品类多数为1且零销时允许编辑金重
        If plduoshu = "1" AndAlso lingxiao = "是" Then
            dgvProducts.Rows(newIdx).Cells(8).ReadOnly = False
            dgvProducts.Rows(newIdx).Cells(8).Style.BackColor = Drawing.Color.White
        End If

        CalculateStatistics()
        txtProductCode.Text = ""
        txtProductCode.Focus()
    End Sub

    ' ========== 保存调拨单（_超级按钮_保存_被单击） ==========
    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        InformationOperationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        InformationOperationAccount = UserAccount

        Dim productCount As Integer = dgvProducts.Rows.Count - 2
        If productCount <= 0 Then
            ShowWarning("调拨商品数量不能为0！")
            Return
        End If

        ' 生成调拨单号
        Dim orderNumber As String = txtOrderNumber.Text & DateTimeOffset.Now.ToUnixTimeSeconds().ToString() & UserAccount

        ' 检查单号是否存在
        Dim checkSql As String = "SELECT * FROM xipunum_erp_transfer_order where transfer_umber ='" & orderNumber & "'"
        Dim checkDt As DataTable = ExecuteQuery(checkSql)
        If checkDt.Rows.Count > 0 Then
            ShowWarning("当前调拨单号已存在，已重新生成单号，请再次点击保存按钮！")
            txtOrderNumber.Text = "DB" & DateTime.Now.ToString("yyyyMMdd") & "****"
            Return
        End If

        ' 验证库存
        Dim validCount As Integer = 0
        For i As Integer = 0 To productCount - 1
            Dim productCode As String = SafeString(dgvProducts.Rows(i).Cells(1).Value)
            Dim quantity As String = SafeString(dgvProducts.Rows(i).Cells(7).Value)
            Dim jinzhong As String = SafeString(dgvProducts.Rows(i).Cells(8).Value)
            Dim originalWarehouse As String = SafeString(dgvProducts.Rows(i).Cells(9).Value)

            Dim ykufangId As String = ""
            If originalWarehouse = "总库" Then
                ykufangId = "0"
            Else
                Dim ykSql As String = "SELECT id FROM xipunum_erp_type where type='商铺' and title='" & originalWarehouse & "' and superior='0' order by id ASC"
                Dim ykDt As DataTable = ExecuteQuery(ykSql)
                If ykDt.Rows.Count > 0 Then
                    ykufangId = SafeString(ykDt.Rows(0)("id"))
                End If
            End If

            Dim stockSql As String = "SELECT * FROM xipunum_erp_shop_kucun where poduct_code ='" & productCode & "' and quantity >='" & quantity & "' and jinzhong>='" & jinzhong & "' and kufang='" & ykufangId & "'"
            Dim stockDt As DataTable = ExecuteQuery(stockSql)
            If stockDt.Rows.Count > 0 Then
                validCount += 1
            End If
        Next

        If validCount <> productCount Then
            ShowWarning("调拨数据存在库存不存在数据，请检查删除后在保存！")
            Return
        End If

        ' 获取总重和调拨类型
        Dim totalWeight As String = SafeString(dgvProducts.Rows(dgvProducts.Rows.Count - 1).Cells(8).Value)
        Dim transferType As String = ""
        If rbTransferIn.Checked Then
            transferType = rbTransferIn.Text
        End If
        If rbTransferOut.Checked Then
            transferType = rbTransferOut.Text
        End If

        ' 获取原库房和新库房id
        Dim lastYkufangId As String = ""
        Dim lastXkufangId As String = ""
        If dgvProducts.Rows.Count > 1 Then
            Dim lastOriginal As String = SafeString(dgvProducts.Rows(0).Cells(9).Value)
            Dim lastNew As String = SafeString(dgvProducts.Rows(0).Cells(10).Value)

            If lastOriginal = "总库" Then
                lastYkufangId = "0"
            Else
                Dim ykSql As String = "SELECT id FROM xipunum_erp_type where type='商铺' and title='" & lastOriginal & "' and superior='0' order by id ASC"
                Dim ykDt As DataTable = ExecuteQuery(ykSql)
                If ykDt.Rows.Count > 0 Then lastYkufangId = SafeString(ykDt.Rows(0)("id"))
            End If

            If lastNew = "总库" Then
                lastXkufangId = "0"
            Else
                Dim xkSql As String = "SELECT id FROM xipunum_erp_type where type='商铺' and title='" & lastNew & "' and superior='0' order by id ASC"
                Dim xkDt As DataTable = ExecuteQuery(xkSql)
                If xkDt.Rows.Count > 0 Then lastXkufangId = SafeString(xkDt.Rows(0)("id"))
            End If
        End If

        ' 插入调拨订单
        Dim insertOrderSql As String = "INSERT INTO xipunum_erp_transfer_order SET " &
            "transfer_umber='" & orderNumber & "'," &
            "type='" & transferType & "'," &
            "total='" & totalWeight & "'," &
            "ykufang='" & lastYkufangId & "'," &
            "xkufang='" & lastXkufangId & "'," &
            "state='0'," &
            "cjuser='" & InformationOperationAccount & "'," &
            "creationtime='" & InformationOperationDate & "'"
        ExecuteCommand(insertOrderSql)

        ' 获取订单ID
        Dim idSql As String = "SELECT id FROM xipunum_erp_transfer_order where transfer_umber='" & orderNumber & "' order by id ASC LIMIT 1"
        Dim idDt As DataTable = ExecuteQuery(idSql)
        Dim orderId As String = ""
        If idDt.Rows.Count > 0 Then
            orderId = SafeString(idDt.Rows(0)("id"))
        End If

        ' 更新单据号
        txtOrderNumber.Text = orderNumber.Substring(0, 10) & orderId.PadLeft(4, "0"c).Substring(Math.Max(0, orderId.Length - 4))
        Dim updateSql As String = "UPDATE xipunum_erp_transfer_order SET transfer_umber= '" & txtOrderNumber.Text & "' WHERE id ='" & orderId & "' LIMIT 1"
        ExecuteCommand(updateSql)

        ' 系统日志
        LogSaveContent = "账户:" & UserAccount & " 增加调拨订单，单号：" & txtOrderNumber.Text
        Dim logSql As String = "INSERT INTO xipunum_erp_xitong_log SET " &
            "type='添加'," &
            "title='增加调拨订单'," &
            "conter='" & LogSaveContent & "'," &
            "user='" & InformationOperationAccount & "'," &
            "creationtime='" & InformationOperationDate & "'"
        ExecuteCommand(logSql)

        ' 保存商品明细
        Dim progressForm As New ProgressBarForm("数据正在保存中...")
        progressForm.Show(Me)

        For i As Integer = 0 To productCount - 1
            Dim pct As Integer = CInt(Math.Round(i / productCount * 100))
            progressForm.SetProgress(pct, $"正在保存数据...({pct}%)")

            Dim productCode As String = SafeString(dgvProducts.Rows(i).Cells(1).Value)
            Dim quantity As String = SafeString(dgvProducts.Rows(i).Cells(7).Value)
            Dim originalWarehouse As String = SafeString(dgvProducts.Rows(i).Cells(9).Value)
            Dim newWarehouse As String = SafeString(dgvProducts.Rows(i).Cells(10).Value)
            Dim jinzhong As String = SafeString(dgvProducts.Rows(i).Cells(8).Value)
            Dim remarks As String = ""
            Dim weight As String = "0.000"
            Dim single As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(6).Value)
            Dim qty As Decimal = SafeDecimal(dgvProducts.Rows(i).Cells(7).Value)
            weight = Math.Round(single * qty, 3).ToString("F3")

            ' 原库房id
            Dim ykufangId As String = ""
            If originalWarehouse = "总库" Then
                ykufangId = "0"
            Else
                Dim ykSql As String = "SELECT id FROM xipunum_erp_type where type='商铺' and title='" & originalWarehouse & "' and superior='0' order by id ASC"
                Dim ykDt As DataTable = ExecuteQuery(ykSql)
                If ykDt.Rows.Count > 0 Then ykufangId = SafeString(ykDt.Rows(0)("id"))
            End If

            ' 新库房id
            Dim xkufangId As String = ""
            If newWarehouse = "总库" Then
                xkufangId = "0"
            Else
                Dim xkSql As String = "SELECT id FROM xipunum_erp_type where type='商铺' and title='" & newWarehouse & "' and superior='0' order by id ASC"
                Dim xkDt As DataTable = ExecuteQuery(xkSql)
                If xkDt.Rows.Count > 0 Then xkufangId = SafeString(xkDt.Rows(0)("id"))
            End If

            ' 插入调拨明细
            Dim insertDetailSql As String = "INSERT INTO xipunum_erp_transfer SET " &
                "order_id='" & orderId & "'," &
                "poduct_code='" & productCode & "'," &
                "type='" & transferType & "'," &
                "quantity='" & quantity & "'," &
                "jinzhong='" & jinzhong & "'," &
                "ykufang='" & ykufangId & "'," &
                "xkufang='" & xkufangId & "'," &
                "remarks='" & remarks & "'," &
                "cjuser='" & InformationOperationAccount & "'," &
                "creationtime='" & InformationOperationDate & "'"
            ExecuteCommand(insertDetailSql)

            ' 原库房扣减库存
            Dim subtractSql As String = "UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '" & quantity & "',jinzhong = jinzhong - '" & jinzhong & "' WHERE poduct_code = '" & productCode & "' AND kufang='" & ykufangId & "'"
            ExecuteCommand(subtractSql)

            ' 目标库房增加库存
            Dim existSql As String = "SELECT * FROM xipunum_erp_shop_kucun where poduct_code ='" & productCode & "' and kufang='" & xkufangId & "'"
            Dim existDt As DataTable = ExecuteQuery(existSql)
            If existDt.Rows.Count = 0 Then
                Dim insertStockSql As String = "INSERT INTO xipunum_erp_shop_kucun SET " &
                    "poduct_code='" & productCode & "'," &
                    "quantity='" & quantity & "'," &
                    "jinzhong='" & jinzhong & "'," &
                    "kufang='" & xkufangId & "'"
                ExecuteCommand(insertStockSql)
            Else
                Dim addSql As String = "UPDATE xipunum_erp_shop_kucun SET quantity = quantity + '" & quantity & "',jinzhong = jinzhong +'" & jinzhong & "' WHERE poduct_code = '" & productCode & "' AND kufang='" & xkufangId & "'"
                ExecuteCommand(addSql)
            End If

            ' 插入历史记录
            Dim historySql As String = "INSERT INTO xipunum_erp_history SET " &
                "poduct_code='" & productCode & "'," &
                "updatetime='" & InformationOperationDate & "'," &
                "number='" & txtOrderNumber.Text & "'," &
                "type='成品调拨'," &
                "quantity='" & quantity & "'," &
                "jinzhong='" & jinzhong & "'," &
                "zhongliang='" & weight & "'," &
                "conter='商品从原：" & originalWarehouse & "-> 新：" & newWarehouse & "'," &
                "cjuser='" & InformationOperationAccount & "'"
            ExecuteCommand(historySql)

            ' 插入商品日志
            Dim shopLogSql As String = "INSERT INTO xipunum_erp_shop_log SET " &
                "poduct_code='" & productCode & "'," &
                "type='调拨'," &
                "creationtime='" & InformationOperationDate & "'"
            ExecuteCommand(shopLogSql)

            ' 系统日志
            LogSaveContent = "账户:" & UserAccount & " 调拨商品，编码：" & productCode & " 从原：" & originalWarehouse & "-> 新：" & newWarehouse
            Dim sysLogSql As String = "INSERT INTO xipunum_erp_xitong_log SET " &
                "type='修改'," &
                "title='修改商品参数'," &
                "conter='" & LogSaveContent & "'," &
                "user='" & InformationOperationAccount & "'," &
                "creationtime='" & InformationOperationDate & "'"
            ExecuteCommand(sysLogSql)

            System.Threading.Thread.Sleep(50)
        Next

        progressForm.Close()

        ' 更新订单库房
        Dim updateOrderSql As String = "UPDATE xipunum_erp_transfer_order SET ykufang= '" & lastYkufangId & "',xkufang= '" & lastXkufangId & "' WHERE id ='" & orderId & "' LIMIT 1"
        ExecuteCommand(updateOrderSql)

        ShowSuccess("商品调拨操作完成！")
        Me.Close()

        ' 刷新主窗口
        If HomePageQueryText = "商品调拨" Then
            Try
                Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
                mainForm?.RefreshSubFolderTable()
            Catch
            End Try
        End If
    End Sub

    ' ========== 表格按钮点击（_高级表格1_按钮被点击） ==========
    Private Sub DgvProducts_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        ' 检查是否点击了操作列
        If dgvProducts.Columns(e.ColumnIndex).Name = "colAction" Then
            Dim deleteRow As Integer = e.RowIndex
            If deleteRow < dgvProducts.Rows.Count - 1 Then
                dgvProducts.Rows.RemoveAt(deleteRow)
                ' 删除合计行后重新添加
                If dgvProducts.Rows.Count > 0 Then
                    dgvProducts.Rows.RemoveAt(dgvProducts.Rows.Count - 1)
                End If
                ' 重新编号
                For i As Integer = 0 To dgvProducts.Rows.Count - 1
                    dgvProducts.Rows(i).Cells(0).Value = (i + 1).ToString()
                Next
                CalculateStatistics()
            End If
        End If
    End Sub

    ' ========== 统计数据（_高级表格1_统计数据） ==========
    Private Sub CalculateStatistics()
        Dim sumQty As Decimal = 0
        Dim sumWeight As Decimal = 0
        Dim dataCount As Integer = dgvProducts.Rows.Count

        For i As Integer = 0 To dataCount - 1
            sumQty += SafeDecimal(dgvProducts.Rows(i).Cells(7).Value)
            sumWeight += SafeDecimal(dgvProducts.Rows(i).Cells(8).Value)
        Next

        ' 格式化重量为3位小数
        Dim weightStr As String = Math.Round(sumWeight, 3).ToString("F3")

        ' 添加合计行
        dgvProducts.Rows.Add()
        Dim totalRow As Integer = dgvProducts.Rows.Count - 1
        dgvProducts.Rows(totalRow).Cells(1).Value = "合计"
        dgvProducts.Rows(totalRow).Cells(7).Value = sumQty.ToString()
        dgvProducts.Rows(totalRow).Cells(8).Value = weightStr

        For j As Integer = 0 To 12
            dgvProducts.Rows(totalRow).Cells(j).ReadOnly = True
            dgvProducts.Rows(totalRow).Cells(j).Style.BackColor = Drawing.Color.LightGray
        Next
    End Sub

    ' ========== 结束编辑（_高级表格1_结束编辑） ==========
    Private Sub DgvProducts_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return

        ' 获取原始值
        Dim originalQty As String = SafeString(dgvProducts.Rows(e.RowIndex).Cells(11).Value)
        Dim currentQty As String = SafeString(dgvProducts.Rows(e.RowIndex).Cells(7).Value)
        Dim originalWeight As String = SafeString(dgvProducts.Rows(e.RowIndex).Cells(12).Value)
        Dim currentWeight As String = SafeString(dgvProducts.Rows(e.RowIndex).Cells(8).Value)

        If selectedCol = 7 Then
            ' 编辑数量列
            If SafeDecimal(currentQty) > SafeDecimal(originalQty) Then
                ShowWarning("调拨数量不能大于库存数量！")
                dgvProducts.Rows(e.RowIndex).Cells(7).Value = originalQty
                Return
            End If

            ' 计算重量 = 原重量 / 原数量 * 当前数量
            If SafeDecimal(originalQty) > 0 Then
                currentWeight = Math.Round(SafeDecimal(currentWeight) / SafeDecimal(originalQty) * SafeDecimal(currentQty), 3).ToString("F3")
            End If
            If SafeDecimal(currentWeight) > SafeDecimal(originalWeight) Then
                currentWeight = originalWeight
            End If
            dgvProducts.Rows(e.RowIndex).Cells(8).Value = currentWeight
        End If

        If selectedCol = 8 Then
            ' 编辑重量列
            If SafeDecimal(currentWeight) > SafeDecimal(originalWeight) Then
                ShowWarning("调拨重量不能大于库存重量！")
                dgvProducts.Rows(e.RowIndex).Cells(8).Value = originalWeight
                Return
            End If
            currentWeight = Math.Round(SafeDecimal(currentWeight), 3).ToString("F3")
            dgvProducts.Rows(e.RowIndex).Cells(8).Value = currentWeight
        End If

        ' 删除合计行后重新计算
        If dgvProducts.Rows.Count > 0 Then
            dgvProducts.Rows.RemoveAt(dgvProducts.Rows.Count - 1)
        End If
        CalculateStatistics()
    End Sub

    ' ========== 标签打印（_超级按钮_标签打印_被单击） ==========
    Private Sub BtnLabelPrint_Click(sender As Object, e As EventArgs)
        ' 统计选中数量
        Dim selectedCount As Integer = 0
        For i As Integer = 0 To dgvProducts.Rows.Count - 2
            Dim code As String = SafeString(dgvProducts.Rows(i).Cells(1).Value)
            If code <> "" Then
                If dgvProducts.Rows(i).Cells(0).Value IsNot Nothing AndAlso
                   (dgvProducts.Rows(i).Cells(0).Value.ToString() = "True" OrElse
                    dgvProducts.Rows(i).Cells(0).Value.ToString() = "1") Then
                    selectedCount += 1
                End If
            End If
        Next

        If selectedCount = 0 Then
            ShowWarning("打印标签数量不能为0！")
            Return
        End If

        ' 连接Access数据库
        Dim accessPath As String = Path.Combine(Application.StartupPath, "data\erpdata.mdb")
        If Not File.Exists(accessPath) Then
            ShowWarning("标签数据库文件不存在！")
            Return
        End If

        Try
            Dim connStr As String = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={accessPath};"
            Using conn As New OleDb.OleDbConnection(connStr)
                conn.Open()
                ' 清空标签表
                Using cmd As New OleDb.OleDbCommand("delete from biaoqian where 1=1", conn)
                    cmd.ExecuteNonQuery()
                End Using
                ' 重置自增ID
                Using cmd As New OleDb.OleDbCommand("ALTER TABLE biaoqian ALTER COLUMN [ID] COUNTER(1, 1)", conn)
                    cmd.ExecuteNonQuery()
                End Using

                ' 遍历商品
                For i As Integer = 0 To dgvProducts.Rows.Count - 2
                    Dim productCode As String = SafeString(dgvProducts.Rows(i).Cells(1).Value)
                    Dim warehouseName As String = SafeString(dgvProducts.Rows(i).Cells(10).Value)

                    If productCode = "" Then Continue For
                    If dgvProducts.Rows(i).Cells(0).Value Is Nothing OrElse
                       (dgvProducts.Rows(i).Cells(0).Value.ToString() <> "True" AndAlso
                        dgvProducts.Rows(i).Cells(0).Value.ToString() <> "1") Then
                        Continue For
                    End If

                    ' 查询商品信息
                    Dim sql As String = "SELECT c.jianxie AS gcjianxie,a.premium_cost AS gongfei,a.basic_cost AS cbgongfei," &
                        "d.caizhi AS caizhi," &
                        "CASE WHEN d.specification_id = '' THEN '未匹配' ELSE e.title END AS guige," &
                        "d.quandu as quanhao,a.factory_condition as chengse," &
                        "d.jin_zhong as jinzhong,d.weight as zhongliang," &
                        "a.sales_cost as xsgongfei,a.sales_surcharge AS xsfujia," &
                        "f.shitou + f.shitou1 AS shizhong,f.stnum + f.shnum1 AS shishu " &
                        "FROM xipunum_erp_store AS a " &
                        "INNER JOIN xipunum_erp_store_order AS b ON b.id = a.order_id " &
                        "LEFT JOIN xipunum_erp_about AS c ON c.id = b.factory " &
                        "INNER JOIN xipunum_erp_shop AS d ON d.poduct_code = a.poduct_code " &
                        "LEFT JOIN xipunum_erp_specs AS e ON e.id = d.specification_id " &
                        "LEFT JOIN xipunum_erp_shop_xiangqian AS f ON f.poduct_code = a.poduct_code " &
                        "WHERE a.poduct_code = '" & productCode & "'"

                    Dim dt As DataTable = ExecuteQuery(sql)
                    If dt.Rows.Count = 0 Then Continue For

                    Dim data As DataRow = dt.Rows(0)
                    Dim gcjianxie As String = SafeString(data("gcjianxie"))
                    Dim gongfei As String = SafeString(data("gongfei"))
                    Dim cbgongfei As String = SafeString(data("cbgongfei"))
                    Dim caizhi As String = SafeString(data("caizhi"))
                    Dim guige As String = SafeString(data("guige"))
                    Dim quanhao As String = SafeString(data("quanhao"))
                    Dim chengse As String = SafeString(data("chengse"))
                    Dim jinzhong As String = SafeString(data("jinzhong"))
                    Dim zhongliang As String = SafeString(data("zhongliang"))
                    Dim xsgongfei As String = SafeString(data("xsgongfei"))
                    Dim xsfujia As String = SafeString(data("xsfujia"))
                    Dim shizhong As String = SafeString(data("shizhong"))
                    Dim shishu As String = SafeString(data("shishu"))

                    ' 查询商铺信息
                    Dim shopSql As String = "SELECT * FROM xipunum_erp_type WHERE type= '商铺' and superior ='0' and title ='" & warehouseName & "' order by id desc"
                    Dim shopDt As DataTable = ExecuteQuery(shopSql)
                    Dim shopName As String = ""
                    Dim shopJianxie As String = ""
                    Dim shopCompany As String = ""
                    Dim shopAddress As String = ""
                    If shopDt.Rows.Count > 0 Then
                        shopName = SafeString(shopDt.Rows(0)("title"))
                        shopJianxie = SafeString(shopDt.Rows(0)("data1"))
                        shopCompany = SafeString(shopDt.Rows(0)("data2"))
                        shopAddress = SafeString(shopDt.Rows(0)("data3"))
                    End If

                    ' 插入标签数据
                    Dim insertSql As String = "insert into biaoqian(工厂简写,工厂工费,成本工费,商铺简写,商品编码,商铺地址,商铺名称,材质,规格,圈号,净含量,净重,总重,销售工费,附加费,石重,石头数) values('" &
                        gcjianxie & "','" & gongfei & "','" & cbgongfei & "','" & shopJianxie & "','" &
                        productCode & "','" & shopAddress & "','" & shopCompany & "','" &
                        caizhi & "','" & guige & "','" & quanhao & "','" & chengse & "','" &
                        jinzhong & "','" & zhongliang & "','" & xsgongfei & "','" & xsfujia & "','" &
                        shizhong & "','" & shishu & "')"
                    Using cmd As New OleDb.OleDbCommand(insertSql, conn)
                        cmd.ExecuteNonQuery()
                    End Using
                Next
            End Using

            ' 运行打印程序
            Dim labelFile As String = ""
            If cmbLabelStyle.SelectedIndex >= 0 Then
                labelFile = cmbLabelStyle.SelectedItem.ToString()
            End If
            Dim cmdLine As String = """" & LabelPrinterConnection & """ /L=""" & Path.Combine(Application.StartupPath, "voucher\biaoqian\") & labelFile & """ /C=0 /X=""" & LabelPrinterName & """ /N /M /Y"
            Process.Start(cmdLine)

            ' 等待打印完成
            System.Threading.Thread.Sleep(7000)

            ' 终止打印进程
            For Each p As Process In Process.GetProcessesByName("lmwprint.exe")
                p.Kill()
            Next
        Catch ex As Exception
            ShowError("标签打印失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 全选（_超级按钮_全选_被单击） ==========
    Private Sub BtnSelectAll_Click(sender As Object, e As EventArgs)
        Dim count As Integer = dgvProducts.Rows.Count - 2
        For i As Integer = 0 To count - 1
            Dim code As String = SafeString(dgvProducts.Rows(i).Cells(1).Value)
            If code <> "" Then
                ' 检查库存是否存在
                Dim sql As String = "SELECT * FROM xipunum_erp_shop_kucun WHERE poduct_code = '" & code & "' AND kufang = '" & UserDepartment & "' AND (quantity > 0 or jinzhong >0) GROUP BY poduct_code,kufang ORDER BY id desc"
                Dim dt As DataTable = ExecuteQuery(sql)
                If dt.Rows.Count > 0 Then
                    dgvProducts.Rows(i).Cells(0).Value = True
                End If
            End If
        Next
    End Sub

    ' ========== 反选（_超级按钮_反选_被单击） ==========
    Private Sub BtnInvertSelect_Click(sender As Object, e As EventArgs)
        Dim count As Integer = dgvProducts.Rows.Count - 2
        For i As Integer = 0 To count - 1
            Dim code As String = SafeString(dgvProducts.Rows(i).Cells(1).Value)
            If code <> "" Then
                Dim sql As String = "SELECT * FROM xipunum_erp_shop_kucun WHERE poduct_code = '" & code & "' AND kufang = '" & UserDepartment & "' AND (quantity > 0 or jinzhong >0) GROUP BY poduct_code,kufang ORDER BY id desc"
                Dim dt As DataTable = ExecuteQuery(sql)
                If dt.Rows.Count > 0 Then
                    Dim currentVal As Boolean = False
                    If dgvProducts.Rows(i).Cells(0).Value IsNot Nothing Then
                        Boolean.TryParse(dgvProducts.Rows(i).Cells(0).Value.ToString(), currentVal)
                    End If
                    dgvProducts.Rows(i).Cells(0).Value = Not currentVal
                End If
            End If
        Next
    End Sub

    ' ========== 提取编码（_超级按钮_提取编码_被单击） ==========
    Private Sub BtnExtractCode_Click(sender As Object, e As EventArgs)
        InboundProductCodeText = ""
        Dim count As Integer = dgvProducts.Rows.Count - 2
        Dim extractCount As Integer = 0
        Dim codeText As String = ""

        For i As Integer = 0 To count - 1
            Dim code As String = SafeString(dgvProducts.Rows(i).Cells(1).Value)
            If code <> "" Then
                Dim isSelected As Boolean = False
                If dgvProducts.Rows(i).Cells(0).Value IsNot Nothing Then
                    Boolean.TryParse(dgvProducts.Rows(i).Cells(0).Value.ToString(), isSelected)
                End If

                If isSelected Then
                    Dim sql As String = "SELECT * FROM xipunum_erp_shop_kucun WHERE poduct_code = '" & code & "' AND kufang = '" & UserDepartment & "' AND (quantity > 0 or jinzhong >0) GROUP BY poduct_code,kufang ORDER BY id desc"
                    Dim dt As DataTable = ExecuteQuery(sql)
                    If dt.Rows.Count > 0 Then
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
        Me.Close()
    End Sub

End Class

' ============================================================================
' 商品信息退库窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品信息退库.form.e.txt
' 包含所有7个程序集变量、18个子程序、所有SQL查询
' 三种模式：添加、详情、再次提交
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class ReturnOrderForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private selectedRow As Integer = 0                    ' 集_行号
    Private selectedCol As Integer = 0                    ' 集_列号
    Private deleteBtn As New Button()                    ' 删除按钮
    Private orderSelected As Integer = -1                ' 局部_订单是否选中
    Private returnProductCode As String = ""             ' 局部_退库商品编码
    Private returnDataCode As String = ""                ' 局部_退库商品数据编码
    Private returnButtonName As String = ""              ' 局部_退库按钮名称（添加/详情/再次提交）

    ' ========== 控件声明 ==========
    Private dgvProducts As New DataGridView()             ' 高级表格1
    Private dgvDeletedData As New DataGridView()          ' 高级表格5（隐藏，记录删除的再次提交数据）
    Private txtOrderNumber As New TextBox()               ' 单据号_编辑框
    Private txtRemarks As New TextBox()                   ' 备注_编辑框
    Private txtWarehouseID As New TextBox()               ' 退库库房_编辑框id
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
        AddHandler dgvProducts.CellEndEdit, AddressOf DgvProducts_CellEndEdit
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
        Me.Text = "商品信息退库"
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
        lblTitle1.Text = "商品信息退库"
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

        ' 退库库房编辑框id（隐藏）
        txtWarehouseID.Text = "0"
        txtWarehouseID.Visible = False
        panelHeader.Controls.Add(txtWarehouseID)

        ' 商品表格
        dgvProducts.Dock = DockStyle.Fill
        dgvProducts.AllowUserToAddRows = False
        dgvProducts.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        dgvProducts.RowHeadersVisible = False
        dgvProducts.BackgroundColor = Drawing.Color.White
        dgvProducts.RowTemplate.Height = 28
        Me.Controls.Add(dgvProducts)

        ' 隐藏的删除数据表格（高级表格5）
        dgvDeletedData.Visible = False
        dgvDeletedData.AllowUserToAddRows = False
        dgvDeletedData.ColumnCount = 2
        dgvDeletedData.Columns(0).Name = "id"
        dgvDeletedData.Columns(1).Name = "code"
        Me.Controls.Add(dgvDeletedData)

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
            If mainForm IsNot Nothing AndAlso mainForm.dgvMain IsNot Nothing AndAlso mainForm.dgvMain.CurrentRow IsNot Nothing Then
                orderSelected = mainForm.dgvMain.CurrentRow.Index
            End If
        Catch
        End Try

        ' 获取退库按钮名称
        returnButtonName = MainForm.ReturnOrderButtonName
        If String.IsNullOrEmpty(returnButtonName) Then returnButtonName = "添加"

        ' 加载表头
        LoadTableHeader()
        ' 清空表格
        ClearGrid()
        ' 设置退库库房id
        txtWarehouseID.Text = "0"

        If returnButtonName = "添加" Then
            ' 生成单据号
            txtOrderNumber.Text = "TK" & Now.ToString("yyyyMMdd") & "****"
            txtRemarks.Text = ""
            panelToolbar.Visible = True
            panelHeader.Enabled = True
            RbScan_CheckedChanged(Nothing, Nothing)
            dgvProducts.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            CalculateStatistics()

        ElseIf returnButtonName = "详情" Then
            ' 详情模式
            txtOrderNumber.Text = ReturnOrderNumber
            Try
                Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
                If mainForm IsNot Nothing AndAlso mainForm.dgvMain IsNot Nothing AndAlso orderSelected >= 0 Then
                    txtRemarks.Text = SafeString(mainForm.dgvMain.Rows(orderSelected).Cells(8).Value)
                End If
            Catch
            End Try
            txtRemarks.ReadOnly = True
            LoadReturnDetails()
            panelToolbar.Visible = False
            panelHeader.Enabled = False
            dgvProducts.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        ElseIf returnButtonName = "再次提交" Then
            ' 再次提交模式
            txtOrderNumber.Text = ReturnOrderNumber
            Try
                Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
                If mainForm IsNot Nothing AndAlso mainForm.dgvMain IsNot Nothing AndAlso orderSelected >= 0 Then
                    txtRemarks.Text = SafeString(mainForm.dgvMain.Rows(orderSelected).Cells(8).Value)
                End If
            Catch
            End Try
            panelToolbar.Visible = True
            panelHeader.Enabled = True
            RbScan_CheckedChanged(Nothing, Nothing)
            dgvProducts.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            LoadResubmitDetails()
            ClearGrid5()
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

        ' 调整头部标签位置
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

        If ReturnOrderNumber <> "" Then
            If returnButtonName = "详情" Then
                tableHead = {"序号", "商品编码", "品类", "规格", "商品名称", "款号", "单件重", "退库数量", "退库金重", "退库重量", "原库房", "新库房", "退库时间", "备注"}
                dataWidth = {50, 120, 100, 100, 200, 120, 80, 80, 80, 80, 80, 80, 140, 150}
            ElseIf returnButtonName = "再次提交" Then
                tableHead = {"序号", "商品编码", "品类", "规格", "商品名称", "款号", "单件重", "退库数量", "退库金重", "退库重量", "原库房", "新库房", "库存数量", "库存重量", "id", "操作"}
                dataWidth = {50, 120, 100, 100, 200, 120, 80, 80, 80, 80, 80, 80, 80, 80, 0, 60}
            End If
        Else
            If returnButtonName = "添加" Then
                tableHead = {"序号", "商品编码", "品类", "规格", "商品名称", "款号", "单件重", "退库数量", "退库金重", "退库重量", "原库房", "新库房", "库存数量", "库存重量", "操作"}
                dataWidth = {50, 120, 100, 100, 200, 120, 80, 80, 80, 80, 80, 80, 80, 80, 60}
            End If
        End If

        For i As Integer = 0 To tableHead.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = tableHead(i)
            col.Name = "col" & i
            col.Width = dataWidth(i)
            If i = tableHead.Length - 1 AndAlso (returnButtonName = "添加" OrElse returnButtonName = "再次提交") Then
                ' 操作列使用按钮列
                Dim btnCol As New DataGridViewButtonColumn()
                btnCol.HeaderText = tableHead(i)
                btnCol.Name = "col" & i
                btnCol.Width = dataWidth(i)
                dgvProducts.Columns.Add(btnCol)
            Else
                dgvProducts.Columns.Add(col)
            End If
        Next

        ' id列隐藏（再次提交模式）
        If returnButtonName = "再次提交" Then
            dgvProducts.Columns("col14").Visible = False
        End If
    End Sub

    ' ========== 清空表格 ==========
    Private Function ClearGrid() As Boolean
        dgvProducts.Rows.Clear()
        Return True
    End Function

    ' ========== 清空隐藏表格5 ==========
    Private Function ClearGrid5() As Boolean
        dgvDeletedData.Rows.Clear()
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

    ' ========== 退库详情被点击（详情模式） ==========
    Private Sub LoadReturnDetails()
        Dim sql As String = "SELECT a.poduct_code AS apoduct_code,CASE WHEN COALESCE(f.title, '') = '' THEN '未匹配' ELSE f.title END AS pinlei," &
            "COALESCE (e1.title, e2.title, '无数据') AS guige,c.product_name AS cproduct_name,c.item_number AS citem_number," &
            "CAST(ROUND( c.single, 3 ) AS DECIMAL ( 10, 3 )) AS csingle,a.quantity AS aquantity," &
            "CAST(ROUND( a.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) AS jinzhong," &
            "CASE WHEN COALESCE (d.lingxiao, '') = '是' THEN CAST(ROUND( a.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) " &
            "ELSE CAST(ROUND(c.single*a.quantity, 3 ) AS DECIMAL ( 10, 3 )) END AS zongzhong," &
            "CASE WHEN a.ykufang = '0' THEN '总库' ELSE g.title END AS yuanku," &
            "CASE WHEN a.xkufang = '0' THEN '总库' ELSE h.title  END AS xinku," &
            "a.creationtime AS acreationtime,a.remarks AS aremarks " &
            "FROM xipunum_erp_tuiku AS a " &
            "INNER JOIN xipunum_erp_tuiku_order AS b ON b.id = a.order_id AND b.tuiku_umber = '" & txtOrderNumber.Text & "' " &
            "INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = c.item_number AND c.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id AND c.item_number IS NOT NULL AND c.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = c.specification_id AND c.specification_id IS NOT NULL AND c.specification_id != '' " &
            "LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE ( e1.category_id, e2.category_id ) " &
            "AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL " &
            "LEFT JOIN xipunum_erp_type AS g ON g.id = a.ykufang AND a.ykufang != '0' " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = a.xkufang AND a.xkufang != '0' " &
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
            dgvProducts.Rows(rowIdx).Cells(2).Value = SafeString(row("pinlei"))
            dgvProducts.Rows(rowIdx).Cells(3).Value = SafeString(row("guige"))
            dgvProducts.Rows(rowIdx).Cells(4).Value = SafeString(row("cproduct_name"))
            dgvProducts.Rows(rowIdx).Cells(5).Value = SafeString(row("citem_number"))
            dgvProducts.Rows(rowIdx).Cells(6).Value = SafeString(row("csingle"))
            dgvProducts.Rows(rowIdx).Cells(7).Value = SafeString(row("aquantity"))
            dgvProducts.Rows(rowIdx).Cells(8).Value = SafeString(row("jinzhong"))
            dgvProducts.Rows(rowIdx).Cells(9).Value = SafeString(row("zongzhong"))
            dgvProducts.Rows(rowIdx).Cells(10).Value = SafeString(row("yuanku"))
            dgvProducts.Rows(rowIdx).Cells(11).Value = SafeString(row("xinku"))
            dgvProducts.Rows(rowIdx).Cells(12).Value = SafeString(row("acreationtime"))
            dgvProducts.Rows(rowIdx).Cells(13).Value = SafeString(row("aremarks"))
        Next

        progressForm.Close()

        ' 查询合计
        Dim sumSql As String = "SELECT sum(tol.aquantity) as aquantity,sum(tol.jinzhong) as jinzhong,sum(tol.zongzhong) as zhongliang FROM (" &
            "SELECT a.quantity AS aquantity,CAST(ROUND( a.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) AS jinzhong," &
            "CASE WHEN COALESCE (d.lingxiao, '') = '是' THEN CAST(ROUND( a.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) " &
            "ELSE CAST(ROUND(c.single*a.quantity, 3 ) AS DECIMAL ( 10, 3 )) END AS zongzhong " &
            "FROM xipunum_erp_tuiku AS a " &
            "INNER JOIN xipunum_erp_tuiku_order AS b ON b.id = a.order_id AND b.tuiku_umber = '" & txtOrderNumber.Text & "' " &
            "INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = c.item_number AND c.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id AND c.item_number IS NOT NULL AND c.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = c.specification_id AND c.specification_id IS NOT NULL AND c.specification_id != '' " &
            "LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE ( e1.category_id, e2.category_id ) " &
            "AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL " &
            "LEFT JOIN xipunum_erp_type AS g ON g.id = a.ykufang AND a.ykufang != '0' " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = a.xkufang AND a.xkufang != '0' " &
            "WHERE 1 = 1 ORDER BY a.id ASC) as tol"

        Dim sumDt As DataTable = DatabaseModule.ExecuteQuery(sumSql, MySQL_Read)
        If sumDt.Rows.Count > 0 Then
            Dim sumRow As DataRow = sumDt.Rows(0)
            Dim sumQuantity As String = SafeString(sumRow("aquantity"))
            Dim sumJinzhong As String = SafeString(sumRow("jinzhong"))
            Dim sumWeight As String = SafeString(sumRow("zhongliang"))

            ' 添加合计行
            Dim sumRowIndex As Integer = dgvProducts.Rows.Add()
            dgvProducts.Rows(sumRowIndex).Cells(1).Value = "合计"
            dgvProducts.Rows(sumRowIndex).Cells(7).Value = sumQuantity
            dgvProducts.Rows(sumRowIndex).Cells(8).Value = sumJinzhong
            dgvProducts.Rows(sumRowIndex).Cells(9).Value = sumWeight

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
        returnProductCode = ""
        returnProductCode = GBKToUTF8(txtProductCode.Text)

        ' 检查商品是否存在
        Dim sql As String = "SELECT * FROM xipunum_erp_shop where (poduct_code='" & returnProductCode & "' or fu_code='" & returnProductCode & "') order by id ASC"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
        If dt.Rows.Count = 0 Then
            ShowWarning("此商品不存在，请输入正确的商品编码！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        ' 检查商品是否在当前库房有库存
        sql = "SELECT a.poduct_code AS apoduct_code FROM xipunum_erp_shop_kucun AS a " &
            "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
            "WHERE a.kufang = '" & UserDepartment & "' AND (a.quantity > 0 or a.jinzhong >0) " &
            "AND (b.poduct_code = '" & returnProductCode & "' OR b.fu_code = '" & returnProductCode & "') ORDER BY a.id DESC"
        Dim dtStock As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
        If dtStock.Rows.Count = 0 Then
            ShowWarning("此商品再当前库房无库存数据！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        If dtStock.Rows.Count = 1 Then
            ' 只有一条记录，直接加载
            returnDataCode = SafeString(dtStock.Rows(0)("apoduct_code"))
            AddProductDataLoad()
        Else
            ' 多条记录，打开副编码查询窗口
            Dim fuCodeForm As New ReturnFuCodeQueryForm()
            fuCodeForm.SearchCode = returnProductCode
            If fuCodeForm.ShowDialog() = DialogResult.OK Then
                returnDataCode = fuCodeForm.SelectedProductCode
                AddProductDataLoad()
            End If
        End If
    End Sub

    ' ========== 添加数据加载 ==========
    Private Sub AddProductDataLoad()
        Dim warehouseName As String = ReturnWarehouseName

        ' 检查商品是否已在退库清单
        For i As Integer = 0 To dgvProducts.Rows.Count - 2
            If SafeString(dgvProducts.Rows(i).Cells(1).Value) = returnDataCode Then
                ShowWarning("此商品已在当前退库清单！")
                txtProductCode.Text = ""
                txtProductCode.Focus()
                Return
            End If
        Next

        ' 查询商品信息
        Dim sql As String = "SELECT a.poduct_code AS apoduct_code," &
            "CASE WHEN COALESCE ( f.title, '' ) = '' THEN'未匹配' ELSE f.title END AS pinlei," &
            "COALESCE ( e1.title, e2.title, '无数据' ) AS guige," &
            "b.product_name AS aproduct_name,b.item_number AS aitem_number," &
            "CAST(ROUND(b.single, 3) AS DECIMAL ( 10, 3 )) AS asingle," &
            "a.quantity AS kucun,CAST(ROUND(a.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) AS jinzhong," &
            "CASE WHEN COALESCE ( d.lingxiao, '' ) = '是' THEN CAST(ROUND(a.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) " &
            "ELSE CAST(ROUND(b.single*a.quantity, 3 ) AS DECIMAL ( 10, 3 )) END AS zongzhong," &
            "CASE WHEN a.kufang = '0' THEN '总库' ELSE c.title END AS yuanku," &
            "b.state AS astate,a.kufang AS akufang," &
            "CASE WHEN COALESCE ( f.shuliang, '' ) = '' THEN '1' ELSE f.shuliang END AS plduoshu," &
            "CASE  WHEN COALESCE ( d.lingxiao, '' ) = '' THEN '否' ELSE d.lingxiao END AS lingxiao " &
            "FROM xipunum_erp_shop_kucun AS a " &
            "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang " &
            "LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = b.item_number AND b.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id AND b.item_number IS NOT NULL AND b.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id AND b.specification_id IS NOT NULL AND b.specification_id != '' " &
            "LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE ( e1.category_id, e2.category_id ) " &
            "AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL " &
            "WHERE a.kufang = '" & UserDepartment & "'  AND (a.quantity > 0 OR a.jinzhong > 0)  " &
            "AND b.poduct_code = '" & returnDataCode & "' ORDER BY a.id DESC"

        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
        If dt.Rows.Count = 0 Then
            ShowWarning("未找到商品库存数据！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        Dim row As DataRow = dt.Rows(0)
        Dim productCode As String = SafeString(row("apoduct_code"))
        Dim pinlei As String = SafeString(row("pinlei"))
        Dim guige As String = SafeString(row("guige"))
        Dim productName As String = SafeString(row("aproduct_name"))
        Dim itemNumber As String = SafeString(row("aitem_number"))
        Dim single As String = SafeString(row("asingle"))
        Dim kucun As String = SafeString(row("kucun"))
        Dim jinzhong As String = SafeString(row("jinzhong"))
        Dim zongzhong As String = SafeString(row("zongzhong"))
        Dim yuanku As String = SafeString(row("yuanku"))
        Dim state As String = SafeString(row("astate"))
        Dim akufang As String = SafeString(row("akufang"))
        Dim plduoshu As String = SafeString(row("plduoshu"))
        Dim lingxiao As String = SafeString(row("lingxiao"))

        ' 检查是否已在退库仓
        If akufang = ReturnWarehouse Then
            ShowWarning("此商品已在退库仓,无需退库！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        ' 零销商品检查金重
        If lingxiao = "是" Then
            If SafeDecimal(jinzhong) <= 0 Then
                ShowWarning("此商品库存金重为0,不能退库！")
                txtProductCode.Text = ""
                txtProductCode.Focus()
                Return
            End If
        End If

        ' 非零销商品检查数量
        If lingxiao = "否" Then
            If SafeDecimal(kucun) <= 0 Then
                ShowWarning("此商品库存数量为0,不能退库！")
                txtProductCode.Text = ""
                txtProductCode.Focus()
                Return
            End If
        End If

        ' 检查商品状态
        If state <> "销售" Then
            ShowWarning("此商品状态为:""" & state & """,无法退库！")
            txtProductCode.Text = ""
            txtProductCode.Focus()
            Return
        End If

        ' 删除合计行，插入数据行
        Dim insertRowIdx As Integer = dgvProducts.Rows.Count
        If dgvProducts.Rows.Count > 0 Then
            dgvProducts.Rows.RemoveAt(dgvProducts.Rows.Count - 1)
        End If
        insertRowIdx = dgvProducts.Rows.Add()

        dgvProducts.Rows(insertRowIdx).Cells(0).Value = (insertRowIdx + 1).ToString()
        dgvProducts.Rows(insertRowIdx).Cells(1).Value = productCode
        dgvProducts.Rows(insertRowIdx).Cells(2).Value = pinlei
        dgvProducts.Rows(insertRowIdx).Cells(3).Value = guige
        dgvProducts.Rows(insertRowIdx).Cells(4).Value = productName
        dgvProducts.Rows(insertRowIdx).Cells(5).Value = itemNumber
        dgvProducts.Rows(insertRowIdx).Cells(6).Value = single
        dgvProducts.Rows(insertRowIdx).Cells(7).Value = kucun
        dgvProducts.Rows(insertRowIdx).Cells(8).Value = jinzhong
        dgvProducts.Rows(insertRowIdx).Cells(9).Value = zongzhong
        dgvProducts.Rows(insertRowIdx).Cells(10).Value = yuanku
        dgvProducts.Rows(insertRowIdx).Cells(11).Value = warehouseName
        dgvProducts.Rows(insertRowIdx).Cells(12).Value = kucun
        dgvProducts.Rows(insertRowIdx).Cells(13).Value = jinzhong

        ' 操作列按钮
        If returnButtonName = "添加" Then
            dgvProducts.Rows(insertRowIdx).Cells(14).Value = "删除"
        ElseIf returnButtonName = "再次提交" Then
            dgvProducts.Rows(insertRowIdx).Cells(14).Value = ""
            dgvProducts.Rows(insertRowIdx).Cells(15).Value = "删除"
        End If

        ' 设置只读和背景色
        For j As Integer = 0 To 13
            dgvProducts.Rows(insertRowIdx).Cells(j).ReadOnly = True
            dgvProducts.Rows(insertRowIdx).Cells(j).Style.BackColor = Drawing.Color.LightGray
        Next

        ' 品类多数为0时可编辑数量
        If plduoshu = "0" Then
            dgvProducts.Rows(insertRowIdx).Cells(7).ReadOnly = False
            dgvProducts.Rows(insertRowIdx).Cells(7).Style.BackColor = Drawing.Color.White
        End If

        ' 品类多数为1且零销时可编辑金重
        If plduoshu = "1" AndAlso lingxiao = "是" Then
            dgvProducts.Rows(insertRowIdx).Cells(8).ReadOnly = False
            dgvProducts.Rows(insertRowIdx).Cells(8).Style.BackColor = Drawing.Color.White
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
            sumQuantity = (SafeDecimal(sumQuantity) + SafeDecimal(dgvProducts.Rows(i).Cells(7).Value)).ToString()
            sumJinzhong = (SafeDecimal(sumJinzhong) + SafeDecimal(dgvProducts.Rows(i).Cells(8).Value)).ToString()
            sumWeight = (SafeDecimal(sumWeight) + SafeDecimal(dgvProducts.Rows(i).Cells(9).Value)).ToString()
        Next

        ' 三位数处理金重
        sumJinzhong = FormatThreeDecimal(sumJinzhong)
        sumWeight = FormatThreeDecimal(sumWeight)

        ' 添加合计行
        Dim sumRowIndex As Integer = dgvProducts.Rows.Add()
        dgvProducts.Rows(sumRowIndex).Cells(1).Value = "合计"
        dgvProducts.Rows(sumRowIndex).Cells(7).Value = sumQuantity
        dgvProducts.Rows(sumRowIndex).Cells(8).Value = sumJinzhong
        dgvProducts.Rows(sumRowIndex).Cells(9).Value = sumWeight

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
        Dim deleteProductCode As String = SafeString(dgvProducts.Rows(deleteRowIdx).Cells(1).Value)
        Dim deleteProductId As String = ""

        ' 再次提交模式：获取id列数据
        If returnButtonName = "再次提交" Then
            deleteProductId = SafeString(dgvProducts.Rows(deleteRowIdx).Cells(14).Value)
            If deleteProductId <> "" Then
                ' 记录到隐藏表格5
                Dim delIdx As Integer = dgvDeletedData.Rows.Add()
                dgvDeletedData.Rows(delIdx).Cells(0).Value = deleteProductId
                dgvDeletedData.Rows(delIdx).Cells(1).Value = deleteProductCode
            End If
        End If

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
            ShowWarning("退库商品数量不能为0！")
            Return
        End If

        If String.IsNullOrEmpty(txtRemarks.Text) Then
            ShowWarning("退库备注不能为空！")
            txtRemarks.Focus()
            Return
        End If

        Dim infoReturnNumber As String = ""
        Dim productTotalWeight As String = ""
        Dim infoRemarks As String = ""

        infoReturnNumber = GBKToUTF8(txtOrderNumber.Text) & CLng(DateTime.UtcNow.ToString("yyyyMMddHHmmss")) & UserAccount
        productTotalWeight = GBKToUTF8(SafeString(dgvProducts.Rows(dgvProducts.Rows.Count - 1).Cells(8).Value))
        infoRemarks = GBKToUTF8(txtRemarks.Text)

        Dim orderDataId As String = ""

        If returnButtonName = "添加" Then
            ' 检查单号是否已存在
            Dim checkSql As String = "SELECT * FROM xipunum_erp_tuiku_order where tuiku_umber ='" & infoReturnNumber & "'"
            Dim checkDt As DataTable = DatabaseModule.ExecuteQuery(checkSql, MySQL_Read)
            If checkDt.Rows.Count > 0 Then
                ShowWarning("当前退库单号已存在，已重新生成单号，请再次点击保存按钮！")
                txtOrderNumber.Text = "TK" & Now.ToString("yyyyMMdd") & "****"
                Return
            End If

            ' 插入退库订单
            Dim insertOrderSql As String = "INSERT INTO xipunum_erp_tuiku_order (tuiku_umber, total, state, remarks, cjuser, creationtime) VALUES (" &
                "'" & infoReturnNumber & "'," &
                "'" & productTotalWeight & "'," &
                "'" & GBKToUTF8("正常") & "'," &
                "'" & infoRemarks & "'," &
                "'" & InformationOperationAccount & "'," &
                "'" & InformationOperationDate & "')"
            DatabaseModule.ExecuteCommand(insertOrderSql, MySQL_Write)

            ' 获取订单ID
            Dim idSql As String = "SELECT id FROM xipunum_erp_tuiku_order where tuiku_umber='" & infoReturnNumber & "' order by id ASC LIMIT 1"
            Dim idDt As DataTable = DatabaseModule.ExecuteQuery(idSql, MySQL_Read)
            If idDt.Rows.Count > 0 Then
                orderDataId = SafeString(idDt.Rows(0)("id"))
            End If

            ' 更新单据号
            txtOrderNumber.Text = infoReturnNumber.Substring(0, 10) & Right("00000000" & orderDataId, 4)
            Dim updateNumSql As String = "UPDATE xipunum_erp_tuiku_order SET tuiku_umber= '" & txtOrderNumber.Text & "' WHERE id ='" & orderDataId & "' LIMIT 1"
            DatabaseModule.ExecuteCommand(updateNumSql, MySQL_Write)

            ' 系统日志
            LogSaveContent = ""
            LogSaveContent = "账户:" & UserAccount & " 增加退库订单，单号：" & txtOrderNumber.Text
            Dim logSql As String = "INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES (" &
                "'添加','增加退库订单','" & LogSaveContent & "','" & InformationOperationAccount & "','" & InformationOperationDate & "')"
            DatabaseModule.ExecuteCommand(logSql, MySQL_Write)

        ElseIf returnButtonName = "再次提交" Then
            ' 更新退库订单
            Dim updateSql As String = "UPDATE xipunum_erp_tuiku_order SET " &
                "total='" & productTotalWeight & "'," &
                "state='" & GBKToUTF8("正常") & "'," &
                "remarks='" & infoRemarks & "'," &
                "cjuser='" & InformationOperationAccount & "'," &
                "updatetime='" & InformationOperationDate & "' " &
                "WHERE tuiku_umber ='" & txtOrderNumber.Text & "' LIMIT 1"
            DatabaseModule.ExecuteCommand(updateSql, MySQL_Write)

            ' 获取订单ID
            Dim idSql As String = "SELECT id FROM xipunum_erp_tuiku_order where tuiku_umber='" & txtOrderNumber.Text & "' order by id ASC LIMIT 1"
            Dim idDt As DataTable = DatabaseModule.ExecuteQuery(idSql, MySQL_Read)
            If idDt.Rows.Count > 0 Then
                orderDataId = SafeString(idDt.Rows(0)("id"))
            End If
        End If

        ' 删除隐藏表格5中记录的已删除数据
        For i As Integer = 0 To dgvDeletedData.Rows.Count - 1
            Dim delId As String = SafeString(dgvDeletedData.Rows(i).Cells(0).Value)
            If delId <> "" Then
                Dim delSql As String = "DELETE FROM xipunum_erp_tuiku where id= '" & delId & "'"
                DatabaseModule.ExecuteCommand(delSql, MySQL_Write)
            End If
        Next

        ' 进度条
        Dim progressForm As New ProgressBarForm()
        progressForm.Show(Me)
        progressForm.LabelText = "数据正在保存中..."
        progressForm.MaxValue = productCount

        ' 遍历商品明细保存
        For i As Integer = 0 To productCount - 1
            progressForm.LabelText = "正在保存数据...(" & CInt((i / productCount) * 100) & "%)"
            progressForm.Value = i + 1

            Dim returnInfoCode As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(1).Value))
            Dim returnInfoQuantity As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(7).Value))
            Dim returnInfoJinzhong As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(8).Value))
            Dim returnInfoOriginWarehouse As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(10).Value))
            Dim returnInfoNewWarehouse As String = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(11).Value))
            Dim returnInfoRemarks As String = ""

            ' 获取原库房id
            Dim originWarehouseId As String = ""
            If SafeString(dgvProducts.Rows(i).Cells(10).Value) = "总库" Then
                originWarehouseId = GBKToUTF8("0")
            Else
                Dim originSql As String = "SELECT id FROM xipunum_erp_type where type='" & GBKToUTF8("商铺") & "' and title='" & returnInfoOriginWarehouse & "' and superior='0' order by id ASC"
                Dim originDt As DataTable = DatabaseModule.ExecuteQuery(originSql, MySQL_Read)
                If originDt.Rows.Count > 0 Then
                    originWarehouseId = SafeString(originDt.Rows(0)("id"))
                End If
            End If

            ' 获取新库房id
            Dim newWarehouseId As String = ""
            If SafeString(dgvProducts.Rows(i).Cells(11).Value) = "总库" Then
                newWarehouseId = GBKToUTF8("0")
            Else
                Dim newSql As String = "SELECT id FROM xipunum_erp_type where type='" & GBKToUTF8("商铺") & "' and title='" & returnInfoNewWarehouse & "' and superior='0' order by id ASC"
                Dim newDt As DataTable = DatabaseModule.ExecuteQuery(newSql, MySQL_Read)
                If newDt.Rows.Count > 0 Then
                    newWarehouseId = SafeString(newDt.Rows(0)("id"))
                End If
            End If

            ' 获取商品id（再次提交模式）
            Dim returnInfoProductId As String = ""
            If returnButtonName = "再次提交" Then
                returnInfoProductId = GBKToUTF8(SafeString(dgvProducts.Rows(i).Cells(14).Value))
            End If

            ' 只有数量或金重大于0才处理
            If SafeDecimal(returnInfoQuantity) > 0 OrElse SafeDecimal(returnInfoJinzhong) > 0 Then

                ' 插入或更新退库明细
                If returnInfoProductId = "" Then
                    ' 新增
                    Dim insertDetailSql As String = "INSERT INTO xipunum_erp_tuiku (order_id, poduct_code, quantity, jinzhong, ykufang, xkufang, remarks, cjuser, creationtime) VALUES (" &
                        "'" & orderDataId & "'," &
                        "'" & returnInfoCode & "'," &
                        "'" & returnInfoQuantity & "'," &
                        "'" & returnInfoJinzhong & "'," &
                        "'" & originWarehouseId & "'," &
                        "'" & newWarehouseId & "'," &
                        "'" & returnInfoRemarks & "'," &
                        "'" & InformationOperationAccount & "'," &
                        "'" & InformationOperationDate & "')"
                    DatabaseModule.ExecuteCommand(insertDetailSql, MySQL_Write)
                Else
                    ' 更新
                    Dim updateDetailSql As String = "UPDATE xipunum_erp_tuiku SET quantity = '" & returnInfoQuantity & "',jinzhong =  '" & returnInfoJinzhong & "' WHERE id='" & returnInfoProductId & "'"
                    DatabaseModule.ExecuteCommand(updateDetailSql, MySQL_Write)
                End If

                ' 更新源库房库存（扣减）
                Dim updateOriginSql As String = "UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '" & returnInfoQuantity & "',jinzhong = jinzhong - '" & returnInfoJinzhong & "' WHERE poduct_code = '" & returnInfoCode & "' AND kufang='" & originWarehouseId & "'"
                DatabaseModule.ExecuteCommand(updateOriginSql, MySQL_Write)

                ' 检查目标库房是否有库存记录
                Dim checkStockSql As String = "SELECT * FROM xipunum_erp_shop_kucun where poduct_code ='" & returnInfoCode & "' and kufang='" & newWarehouseId & "'"
                Dim checkStockDt As DataTable = DatabaseModule.ExecuteQuery(checkStockSql, MySQL_Read)
                If checkStockDt.Rows.Count = 0 Then
                    ' 新增库存记录
                    Dim insertStockSql As String = "INSERT INTO xipunum_erp_shop_kucun (poduct_code, quantity, jinzhong, kufang) VALUES (" &
                        "'" & returnInfoCode & "','" & returnInfoQuantity & "','" & returnInfoJinzhong & "','" & newWarehouseId & "')"
                    DatabaseModule.ExecuteCommand(insertStockSql, MySQL_Write)
                Else
                    ' 更新库存（增加）
                    Dim updateStockSql As String = "UPDATE xipunum_erp_shop_kucun SET quantity = quantity + '" & returnInfoQuantity & "',jinzhong = jinzhong + '" & returnInfoJinzhong & "' WHERE poduct_code = '" & returnInfoCode & "' AND kufang='" & newWarehouseId & "'"
                    DatabaseModule.ExecuteCommand(updateStockSql, MySQL_Write)
                End If

                ' 插入历史记录
                Dim historySql As String = "INSERT INTO xipunum_erp_history (poduct_code, updatetime, number, type, quantity, conter, cjuser) VALUES (" &
                    "'" & returnInfoCode & "'," &
                    "'" & InformationOperationDate & "'," &
                    "'" & txtOrderNumber.Text & "'," &
                    "'" & GBKToUTF8("成品退库") & "'," &
                    "'" & returnInfoQuantity & "'," &
                    "'" & GBKToUTF8("商品从原：" & SafeString(dgvProducts.Rows(i).Cells(10).Value) & "-> 新：" & SafeString(dgvProducts.Rows(i).Cells(11).Value)) & "'," &
                    "'" & InformationOperationAccount & "')"
                DatabaseModule.ExecuteCommand(historySql, MySQL_Write)

                ' 插入商品日志
                Dim shopLogSql As String = "INSERT INTO xipunum_erp_shop_log (poduct_code, type, creationtime) VALUES (" &
                    "'" & returnInfoCode & "','" & GBKToUTF8("退库") & "','" & InformationOperationDate & "')"
                DatabaseModule.ExecuteCommand(shopLogSql, MySQL_Write)

                ' 插入系统日志
                LogSaveContent = ""
                LogSaveContent = "账户:" & UserAccount & " 退库商品，编码：" & returnInfoCode & " 从原：" & SafeString(dgvProducts.Rows(i).Cells(10).Value) & "-> 新：" & SafeString(dgvProducts.Rows(i).Cells(11).Value)
                Dim sysLogSql As String = "INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES (" &
                    "'修改','修改商品参数','" & LogSaveContent & "','" & InformationOperationAccount & "','" & InformationOperationDate & "')"
                DatabaseModule.ExecuteCommand(sysLogSql, MySQL_Write)
            End If

            System.Threading.Thread.Sleep(50)
        Next

        progressForm.Close()

        ShowSuccess("商品退库操作完成！")
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

    ' ========== 结束编辑 ==========
    Private Sub DgvProducts_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        If SafeString(dgvProducts.Rows(e.RowIndex).Cells(1).Value) = "合计" Then Return

        Dim originQuantity As String = SafeString(dgvProducts.Rows(e.RowIndex).Cells(12).Value)
        Dim originJinzhong As String = SafeString(dgvProducts.Rows(e.RowIndex).Cells(13).Value)
        Dim singleWeight As String = SafeString(dgvProducts.Rows(e.RowIndex).Cells(6).Value)
        Dim editQuantity As String = SafeString(dgvProducts.Rows(e.RowIndex).Cells(7).Value)
        Dim editJinzhong As String = SafeString(dgvProducts.Rows(e.RowIndex).Cells(8).Value)
        Dim editWeight As String = SafeString(dgvProducts.Rows(e.RowIndex).Cells(9).Value)

        ' 编辑数量列（列7）
        If selectedCol = 7 Then
            If SafeDecimal(editQuantity) > SafeDecimal(originQuantity) Then
                ShowWarning("退库数量不能大于库存数量！")
                dgvProducts.Rows(e.RowIndex).Cells(7).Value = originQuantity
                Return
            End If

            ' 按比例计算金重
            editJinzhong = (SafeDecimal(originJinzhong) / SafeDecimal(originQuantity) * SafeDecimal(editQuantity)).ToString()
            If SafeDecimal(editJinzhong) > SafeDecimal(originJinzhong) Then
                editJinzhong = originJinzhong
            End If

            ' 计算重量
            editWeight = (SafeDecimal(singleWeight) * SafeDecimal(editQuantity)).ToString()

            ' 三位数处理
            editJinzhong = FormatThreeDecimal(editJinzhong)
            editWeight = FormatThreeDecimal(editWeight)

            dgvProducts.Rows(e.RowIndex).Cells(8).Value = editJinzhong
            dgvProducts.Rows(e.RowIndex).Cells(9).Value = editWeight
        End If

        ' 编辑金重列（列8）
        If selectedCol = 8 Then
            If SafeDecimal(editJinzhong) > SafeDecimal(originJinzhong) Then
                ShowWarning("退库金重不能大于库存金重！")
                dgvProducts.Rows(e.RowIndex).Cells(8).Value = originJinzhong
                Return
            End If

            ' 三位数处理
            editJinzhong = FormatThreeDecimal(editJinzhong)
            dgvProducts.Rows(e.RowIndex).Cells(8).Value = editJinzhong
            dgvProducts.Rows(e.RowIndex).Cells(9).Value = editJinzhong
        End If

        ' 删除合计行并重新统计
        If dgvProducts.Rows.Count > 0 Then
            Dim lastRowIdx As Integer = dgvProducts.Rows.Count - 1
            If SafeString(dgvProducts.Rows(lastRowIdx).Cells(1).Value) = "合计" Then
                dgvProducts.Rows.RemoveAt(lastRowIdx)
            End If
        End If
        CalculateStatistics()
    End Sub

    ' ========== 再次提交（加载已有退库详情） ==========
    Private Sub LoadResubmitDetails()
        Dim sql As String = "SELECT a.id as aid,a.poduct_code AS apoduct_code," &
            "CASE WHEN COALESCE(f.title, '') = '' THEN '未匹配' ELSE f.title END AS pinlei," &
            "COALESCE (e1.title, e2.title, '无数据') AS guige," &
            "c.product_name AS cproduct_name,c.item_number AS citem_number," &
            "CAST(ROUND( c.single, 3 ) AS DECIMAL ( 10, 3 )) AS csingle," &
            "a.quantity AS aquantity,CAST(ROUND( i.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) AS jinzhong," &
            "CASE WHEN COALESCE (d.lingxiao, '') = '是' THEN CAST(ROUND( i.jinzhong, 3 ) AS DECIMAL ( 10, 3 )) " &
            "ELSE CAST(ROUND(c.single*a.quantity, 3 ) AS DECIMAL ( 10, 3 )) END AS zongzhong," &
            "CASE WHEN a.ykufang = '0' THEN '总库' ELSE g.title END AS yuanku," &
            "CASE WHEN a.xkufang = '0' THEN '总库' ELSE h.title  END AS xinku," &
            "CASE WHEN COALESCE ( f.shuliang, '' ) = '' THEN '1' ELSE f.shuliang END AS plduoshu," &
            "CASE  WHEN COALESCE ( d.lingxiao, '' ) = '' THEN '否' ELSE d.lingxiao END AS lingxiao " &
            "FROM xipunum_erp_tuiku AS a " &
            "INNER JOIN xipunum_erp_tuiku_order AS b ON b.id = a.order_id AND b.tuiku_umber = '" & txtOrderNumber.Text & "' " &
            "INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = c.item_number AND c.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id AND c.item_number IS NOT NULL AND c.item_number != '' " &
            "LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = c.specification_id AND c.specification_id IS NOT NULL AND c.specification_id != '' " &
            "LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE ( e1.category_id, e2.category_id ) " &
            "AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL " &
            "LEFT JOIN xipunum_erp_type AS g ON g.id = a.ykufang AND a.ykufang != '0' " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = a.xkufang AND a.xkufang != '0' " &
            "INNER JOIN xipunum_erp_shop_kucun AS i ON i.poduct_code = a.poduct_code AND i.kufang = a.ykufang " &
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
            Dim infoProductId As String = SafeString(row("aid"))
            Dim infoProductCode As String = SafeString(row("apoduct_code"))
            Dim infoPinlei As String = SafeString(row("pinlei"))
            Dim infoGuige As String = SafeString(row("guige"))
            Dim infoProductName As String = SafeString(row("cproduct_name"))
            Dim infoItemNumber As String = SafeString(row("citem_number"))
            Dim infoSingle As String = SafeString(row("csingle"))
            Dim infoQuantity As String = SafeString(row("aquantity"))
            Dim infoJinzhong As String = SafeString(row("jinzhong"))
            Dim infoWeight As String = SafeString(row("zongzhong"))
            Dim infoOriginKu As String = SafeString(row("yuanku"))
            Dim infoNewKu As String = SafeString(row("xinku"))
            Dim infoPlduoshu As String = SafeString(row("plduoshu"))
            Dim infoLingxiao As String = SafeString(row("lingxiao"))

            Dim rowIdx As Integer = dgvProducts.Rows.Add()
            dgvProducts.Rows(rowIdx).Cells(0).Value = (i + 1).ToString().PadLeft(detailCount.ToString().Length, "0"c)
            dgvProducts.Rows(rowIdx).Cells(1).Value = infoProductCode
            dgvProducts.Rows(rowIdx).Cells(2).Value = infoPinlei
            dgvProducts.Rows(rowIdx).Cells(3).Value = infoGuige
            dgvProducts.Rows(rowIdx).Cells(4).Value = infoProductName
            dgvProducts.Rows(rowIdx).Cells(5).Value = infoItemNumber
            dgvProducts.Rows(rowIdx).Cells(6).Value = infoSingle
            dgvProducts.Rows(rowIdx).Cells(7).Value = infoQuantity
            dgvProducts.Rows(rowIdx).Cells(8).Value = infoJinzhong
            dgvProducts.Rows(rowIdx).Cells(9).Value = infoWeight
            dgvProducts.Rows(rowIdx).Cells(10).Value = infoOriginKu
            dgvProducts.Rows(rowIdx).Cells(11).Value = infoNewKu
            dgvProducts.Rows(rowIdx).Cells(12).Value = infoQuantity
            dgvProducts.Rows(rowIdx).Cells(13).Value = infoJinzhong
            dgvProducts.Rows(rowIdx).Cells(14).Value = infoProductId
            dgvProducts.Rows(rowIdx).Cells(15).Value = "删除"

            ' 设置只读和背景色
            For j As Integer = 0 To 15
                dgvProducts.Rows(rowIdx).Cells(j).ReadOnly = True
                dgvProducts.Rows(rowIdx).Cells(j).Style.BackColor = Drawing.Color.LightGray
            Next

            ' 品类多数为0时可编辑数量
            If infoPlduoshu = "0" Then
                dgvProducts.Rows(rowIdx).Cells(7).ReadOnly = False
                dgvProducts.Rows(rowIdx).Cells(7).Style.BackColor = Drawing.Color.White
            End If

            ' 品类多数为1且零销时可编辑金重
            If infoPlduoshu = "1" AndAlso infoLingxiao = "是" Then
                dgvProducts.Rows(rowIdx).Cells(8).ReadOnly = False
                dgvProducts.Rows(rowIdx).Cells(8).Style.BackColor = Drawing.Color.White
            End If
        Next

        progressForm.Close()

        ' 统计合计
        CalculateStatistics()
    End Sub

End Class

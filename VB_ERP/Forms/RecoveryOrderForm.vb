' ============================================================================
' 商品信息回收窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品信息回收.form.e.txt
' 包含所有5个程序集变量、17个子程序、所有SQL查询
' 两种模式：新建回收单（orderSelected=-1）、查看已有回收单
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class RecoveryOrderForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（5个） ==========
    Private row As Integer = 0                           ' 集_行号
    Private col As Integer = 0                           ' 集_列号
    Private deleteBtn As New Button()                   ' 删除按钮
    Private orderSelected As Integer = -1                ' 局部_订单是否选中
    Private editState As Integer = 0                     ' 局部_回收编辑状态

    ' ========== 控件声明 ==========
    Private dgvProducts As New DataGridView()           ' 高级表格1
    Private txtOrderNumber As New TextBox()              ' 单据号_编辑框
    Private txtRemarks As New TextBox()                  ' 备注_编辑框
    Private txtMemberName As New TextBox()              ' 会员姓名_编辑框
    Private txtPhone As New TextBox()                   ' 联系电话_编辑框
    Private txtSalesman As New TextBox()                 ' 业务员_编辑框
    Private txtTaxRate As New TextBox()                  ' 税点_编辑框
    Private txtTaxAmount As New TextBox()                ' 税收_编辑框
    Private txtPayable As New TextBox()                  ' 应付_编辑框
    Private txtActualPay As New TextBox()                ' 实付_编辑框
    Private btnSave As New Button()                     ' 按钮_保存
    Private btnReset As New Button()                     ' 按钮_重置
    Private panelHeader As New Panel()                   ' 外形框_头部
    Private grpRemarks As New GroupBox()                 ' 分组框_备注
    Private lblTitle1 As New Label()                     ' 透明标签1
    Private lblTitle2 As New Label()                     ' 透明标签2

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler dgvProducts.SelectionChanged, AddressOf DgvProducts_SelectionChanged
        AddHandler dgvProducts.CellClick, AddressOf DgvProducts_CellClick
        AddHandler dgvProducts.CellBeginEdit, AddressOf DgvProducts_CellBeginEdit
        AddHandler dgvProducts.CellEndEdit, AddressOf DgvProducts_CellEndEdit
        AddHandler txtPhone.TextChanged, AddressOf TxtPhone_TextChanged
        AddHandler txtPhone.KeyDown, AddressOf TxtPhone_KeyDown
        AddHandler txtTaxRate.TextChanged, AddressOf TxtTaxRate_TextChanged
        AddHandler txtPayable.TextChanged, AddressOf TxtPayable_TextChanged
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnReset.Click, AddressOf BtnReset_Click
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "商品信息回收"
        Me.Size = New Drawing.Size(1427, 800)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' 头部面板
        panelHeader.Dock = DockStyle.Top
        panelHeader.Height = 65
        panelHeader.BackColor = Drawing.Color.FromArgb(248, 248, 248)
        Me.Controls.Add(panelHeader)

        ' 透明标签1（标题）
        lblTitle1.Text = "商品信息回收"
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

        ' 会员姓名标签和编辑框
        Dim lblMember As New Label()
        lblMember.Text = "会员姓名："
        lblMember.AutoSize = True
        lblMember.Location = New Drawing.Point(140, 15)
        panelHeader.Controls.Add(lblMember)

        txtMemberName.Location = New Drawing.Point(210, 12)
        txtMemberName.Size = New Drawing.Size(100, 25)
        panelHeader.Controls.Add(txtMemberName)

        ' 联系电话标签和编辑框
        Dim lblPhone As New Label()
        lblPhone.Text = "联系电话："
        lblPhone.AutoSize = True
        lblPhone.Location = New Drawing.Point(320, 15)
        panelHeader.Controls.Add(lblPhone)

        txtPhone.Location = New Drawing.Point(390, 12)
        txtPhone.Size = New Drawing.Size(100, 25)
        panelHeader.Controls.Add(txtPhone)

        ' 业务员标签和编辑框
        Dim lblSales As New Label()
        lblSales.Text = "业务员："
        lblSales.AutoSize = True
        lblSales.Location = New Drawing.Point(500, 15)
        panelHeader.Controls.Add(lblSales)

        txtSalesman.Location = New Drawing.Point(555, 12)
        txtSalesman.Size = New Drawing.Size(80, 25)
        panelHeader.Controls.Add(txtSalesman)

        ' 税点标签和编辑框
        Dim lblTaxRate As New Label()
        lblTaxRate.Text = "税点："
        lblTaxRate.AutoSize = True
        lblTaxRate.Location = New Drawing.Point(645, 15)
        panelHeader.Controls.Add(lblTaxRate)

        txtTaxRate.Location = New Drawing.Point(680, 12)
        txtTaxRate.Size = New Drawing.Size(40, 25)
        panelHeader.Controls.Add(txtTaxRate)

        ' 税收标签和编辑框
        Dim lblTaxAmount As New Label()
        lblTaxAmount.Text = "税收："
        lblTaxAmount.AutoSize = True
        lblTaxAmount.Location = New Drawing.Point(725, 15)
        panelHeader.Controls.Add(lblTaxAmount)

        txtTaxAmount.Location = New Drawing.Point(760, 12)
        txtTaxAmount.Size = New Drawing.Size(58, 25)
        txtTaxAmount.ReadOnly = True
        panelHeader.Controls.Add(txtTaxAmount)

        ' 应付标签和编辑框
        Dim lblPayable As New Label()
        lblPayable.Text = "应付："
        lblPayable.AutoSize = True
        lblPayable.Location = New Drawing.Point(825, 15)
        panelHeader.Controls.Add(lblPayable)

        txtPayable.Location = New Drawing.Point(860, 12)
        txtPayable.Size = New Drawing.Size(90, 25)
        txtPayable.ReadOnly = True
        panelHeader.Controls.Add(txtPayable)

        ' 实付标签和编辑框
        Dim lblActualPay As New Label()
        lblActualPay.Text = "实付："
        lblActualPay.AutoSize = True
        lblActualPay.Location = New Drawing.Point(955, 15)
        panelHeader.Controls.Add(lblActualPay)

        txtActualPay.Location = New Drawing.Point(990, 12)
        txtActualPay.Size = New Drawing.Size(78, 25)
        panelHeader.Controls.Add(txtActualPay)

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

        ' 保存按钮
        btnSave.Text = "保存"
        btnSave.Size = New Drawing.Size(75, 30)
        panelHeader.Controls.Add(btnSave)

        ' 重置按钮
        btnReset.Text = "重置"
        btnReset.Size = New Drawing.Size(75, 30)
        panelHeader.Controls.Add(btnReset)

        ' 删除按钮样式
        deleteBtn.Text = "删除"
        deleteBtn.BackColor = Drawing.Color.FromArgb(255, 200, 200)

        ' 调整ZOrder
        Me.Controls.SetChildIndex(dgvProducts, 0)
        Me.Controls.SetChildIndex(grpRemarks, 1)
        Me.Controls.SetChildIndex(panelHeader, 2)
    End Sub

    ' ========== 窗口创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        editState = 0

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

        If orderSelected = -1 Then
            ' 新建回收单
            txtOrderNumber.Text = (Now.ToString("yyyyMMddHHmmss") & New Random().Next(1000, 9999).ToString()).Substring(0, 18) & "3"
            txtRemarks.Text = ""
            txtRemarks.ReadOnly = False
            txtMemberName.Text = ""
            txtMemberName.ReadOnly = False
            txtPhone.Text = ""
            txtPhone.ReadOnly = False
            txtSalesman.Text = UserName
            txtSalesman.ReadOnly = False
            txtTaxRate.Text = "0"
            txtTaxRate.ReadOnly = False
            txtTaxAmount.Text = ""
            txtTaxAmount.ReadOnly = True
            txtPayable.Text = ""
            txtPayable.ReadOnly = True
            txtActualPay.Text = ""
            txtActualPay.ReadOnly = False

            btnSave.Visible = True
            btnReset.Visible = True
            AddProductRow()
            dgvProducts.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        Else
            ' 查看已有回收单
            Dim orderNumber As String = ""
            Try
                Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
                If mainForm IsNot Nothing AndAlso mainForm.dgvMain IsNot Nothing Then
                    ' 尝试从主表格获取单据号
                    orderNumber = SafeString(mainForm.dgvMain.Rows(orderSelected).Cells(1).Value)
                End If
            Catch
            End Try

            If String.IsNullOrEmpty(orderNumber) AndAlso RecoveryOrderNumber <> "" Then
                orderNumber = RecoveryOrderNumber
            End If

            txtOrderNumber.Text = orderNumber

            ' 从数据库查询订单头部信息
            Try
                Dim headerSql As String = "SELECT a.*, b.name AS member_name, b.tel AS member_tel " &
                    "FROM xipunum_erp_retreat_order AS a " &
                    "LEFT JOIN xipunum_erp_member AS b ON b.customer_code = a.customer_code " &
                    "WHERE a.retrea_umber = '" & SafeSQL(orderNumber) & "' LIMIT 1"
                Dim headerDt As DataTable = DatabaseModule.ExecuteQuery(headerSql, MySQL_Read)
                If headerDt.Rows.Count > 0 Then
                    Dim headerRow As DataRow = headerDt.Rows(0)
                    txtMemberName.Text = SafeString(headerRow("member_name"))
                    txtPhone.Text = SafeString(headerRow("member_tel"))
                    txtSalesman.Text = SafeString(headerRow("salesman"))
                    txtTaxRate.Text = SafeString(headerRow("tax_rate"))
                    txtTaxAmount.Text = SafeString(headerRow("rate_amount"))
                    txtPayable.Text = SafeString(headerRow("ying_amount"))
                    txtActualPay.Text = SafeString(headerRow("settlement"))
                    txtRemarks.Text = SafeString(headerRow("remarks"))
                End If
            Catch
            End Try

            txtRemarks.ReadOnly = True
            txtMemberName.ReadOnly = True
            txtPhone.ReadOnly = True
            txtSalesman.ReadOnly = True
            txtTaxRate.ReadOnly = True
            txtTaxAmount.ReadOnly = True
            txtPayable.ReadOnly = True
            txtActualPay.ReadOnly = True

            btnSave.Visible = False
            btnReset.Visible = False
            LoadRecoveryDetails()
            dgvProducts.ReadOnly = True
            dgvProducts.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End If
    End Sub

    ' ========== 窗口尺寸改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        Dim nWidth As Integer = Me.ClientRectangle.Width
        Dim nHeight As Integer = Me.ClientRectangle.Height

        panelHeader.Width = nWidth - 10
        panelHeader.Left = 5
        panelHeader.Top = 0

        dgvProducts.Width = nWidth - 10
        dgvProducts.Left = 5
        dgvProducts.Top = 70
        dgvProducts.RowTemplate.Height = 28
        dgvProducts.Height = nHeight - panelHeader.Height - 100

        grpRemarks.Left = 5
        grpRemarks.Top = nHeight - 85
        If orderSelected = -1 Then
            grpRemarks.Width = nWidth - 415
        Else
            grpRemarks.Width = nWidth - 10
        End If
        txtRemarks.Width = grpRemarks.Width - 40

        btnReset.Top = 30
        btnReset.Left = panelHeader.Width - 170
        btnSave.Top = 30
        btnSave.Left = panelHeader.Width - 85

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
            tableHead = {"序号", "商品名称", "数量", "总重", "金重", "成色", "回收克价", "其他费用", "回收金额", "备注", "导购员", "操作"}
            dataWidth = {50, 200, 100, 100, 100, 100, 100, 100, 100, 260, 100, 65}
        Else
            ' 查看模式
            tableHead = {"序号", "商品名称", "数量", "总重", "金重", "成色", "回收克价", "其他费用", "回收金额", "导购员", "回收时间", "备注"}
            dataWidth = {50, 200, 100, 100, 100, 100, 100, 100, 100, 100, 140, 260}
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
                col.SortMode = DataGridViewColumnSortMode.NotSortable
                dgvProducts.Columns.Add(col)
            End If
        Next

        ' 新建模式：加载回收名称下拉列表和导购员下拉列表
        If orderSelected = -1 Then
            ' 回收名称列表（列1）
            Try
                Dim titleSql As String = "SELECT * FROM xipunum_erp_retreat_title WHERE 1=1 order by id asc"
                Dim titleDt As DataTable = DatabaseModule.ExecuteQuery(titleSql, MySQL_Write)
                ' 存储回收名称列表供后续验证使用
                recoveryTitleList = New List(Of String)()
                For Each r As DataRow In titleDt.Rows
                    recoveryTitleList.Add(SafeString(r("title")))
                Next
            Catch
            End Try

            ' 导购员列表（列10，即倒数第2列）
            Try
                Dim guideSql As String = ""
                If UserPermission = "全部" Then
                    guideSql = "SELECT name FROM xipunum_erp_user where state='0' order by id ASC"
                ElseIf UserPermission = "店铺" Then
                    guideSql = "SELECT name FROM xipunum_erp_user where department='" & SafeSQL(UserDepartment) & "' and state='0' order by id ASC"
                Else
                    guideSql = "SELECT name FROM xipunum_erp_user where state='0' order by id ASC"
                End If
                Dim guideDt As DataTable = DatabaseModule.ExecuteQuery(guideSql, MySQL_Read)
                guideList = New List(Of String)()
                For Each r As DataRow In guideDt.Rows
                    guideList.Add(SafeString(r("name")))
                Next
            Catch
            End Try
        End If
    End Sub

    ' 回收名称列表和导购员列表（供验证使用）
    Private recoveryTitleList As New List(Of String)()
    Private guideList As New List(Of String)()

    ' ========== 清空表格 ==========
    Private Function ClearGrid() As Boolean
        dgvProducts.Rows.Clear()
        Return True
    End Function

    ' ========== 光标位置改变 ==========
    Private Sub DgvProducts_SelectionChanged(sender As Object, e As EventArgs)
        If dgvProducts.CurrentCell IsNot Nothing Then
            row = dgvProducts.CurrentCell.RowIndex
            col = dgvProducts.CurrentCell.ColumnIndex
        End If
    End Sub

    ' ========== 回收详情（查看模式） ==========
    Private Sub LoadRecoveryDetails()
        Dim sql As String = "SELECT c.title AS aproduct_name,a.quantity AS aquantity," &
            "CAST(ROUND( a.total, 3 ) AS DECIMAL ( 10, 3 )) AS atotal," &
            "CAST(ROUND( a.jin_zhong, 3 ) AS DECIMAL ( 10, 3 )) AS ajin_zhong," &
            "a.chengse AS achengse," &
            "CAST(ROUND( a.price, 3 ) AS DECIMAL ( 10, 2 )) AS aprice," &
            "CAST(ROUND( a.qita_price, 3 ) AS DECIMAL ( 10, 2 )) AS aqita_price," &
            "CAST(ROUND( a.retreat_amount, 3 ) AS DECIMAL ( 10, 2 )) AS aretreat_amount," &
            "a.shopping_guide as daogou,a.remarks AS aremarks," &
            "a.huishoutime AS ahuishoutime,g.name as gname " &
            "FROM xipunum_erp_retreat AS a " &
            "INNER JOIN xipunum_erp_retreat_order AS b ON b.id = a.order_id " &
            "AND b.retrea_umber = '" & SafeSQL(txtOrderNumber.Text) & "' " &
            "LEFT JOIN xipunum_erp_user AS g ON g.user = a.shopping_guide " &
            "INNER JOIN xipunum_erp_retreat_title AS c ON c.id = a.product_name " &
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

            Dim r As DataRow = dt.Rows(i)
            Dim rowIdx As Integer = dgvProducts.Rows.Add()

            dgvProducts.Rows(rowIdx).Cells(0).Value = (i + 1).ToString().PadLeft(detailCount.ToString().Length, "0"c)
            dgvProducts.Rows(rowIdx).Cells(1).Value = SafeString(r("aproduct_name"))
            dgvProducts.Rows(rowIdx).Cells(2).Value = SafeString(r("aquantity"))
            dgvProducts.Rows(rowIdx).Cells(3).Value = SafeString(r("atotal"))
            dgvProducts.Rows(rowIdx).Cells(4).Value = SafeString(r("ajin_zhong"))
            dgvProducts.Rows(rowIdx).Cells(5).Value = SafeString(r("achengse"))
            dgvProducts.Rows(rowIdx).Cells(6).Value = SafeString(r("aprice"))
            dgvProducts.Rows(rowIdx).Cells(7).Value = SafeString(r("aqita_price"))
            dgvProducts.Rows(rowIdx).Cells(8).Value = SafeString(r("aretreat_amount"))
            dgvProducts.Rows(rowIdx).Cells(9).Value = SafeString(r("gname"))
            dgvProducts.Rows(rowIdx).Cells(10).Value = SafeString(r("ahuishoutime"))
            dgvProducts.Rows(rowIdx).Cells(11).Value = SafeString(r("aremarks"))
        Next

        progressForm.Close()

        CalculateStatistics()

        ' 居中对齐
        For i As Integer = 0 To dgvProducts.Rows.Count - 1
            For j As Integer = 0 To dgvProducts.Columns.Count - 1
                dgvProducts.Rows(i).Cells(j).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Next
        Next
    End Sub

    ' ========== 联系电话内容改变 ==========
    Private Sub TxtPhone_TextChanged(sender As Object, e As EventArgs)
        If txtPhone.Text.Length > 10 Then
            Try
                Dim sql As String = "SELECT * FROM xipunum_erp_member where tel ='" & SafeSQL(txtPhone.Text) & "' LIMIT 1"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                If dt.Rows.Count > 0 Then
                    txtMemberName.Text = SafeString(dt.Rows(0)("name"))
                End If
            Catch
            End Try
        End If
    End Sub

    ' ========== 联系电话按键 ==========
    Private Sub TxtPhone_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If String.IsNullOrEmpty(txtPhone.Text) Then
                ShowWarning("回收客户联系电话不能为空！")
                txtPhone.Text = ""
                txtPhone.Focus()
                Return
            End If
            TxtPhone_TextChanged(Nothing, Nothing)
        End If
    End Sub

    ' ========== 回收数据合计 ==========
    Private Sub CalculateStatistics()
        Dim dataRowCount As Integer = dgvProducts.Rows.Count - 1  ' 减去合计行

        ' 如果没有合计行，dataRowCount就是总行数
        If dgvProducts.Rows.Count = 0 Then Return

        ' 检查最后一行是否是合计行，如果是则删除
        If dgvProducts.Rows.Count > 0 AndAlso SafeString(dgvProducts.Rows(dgvProducts.Rows.Count - 1).Cells(1).Value) = "合计" Then
            dgvProducts.Rows.RemoveAt(dgvProducts.Rows.Count - 1)
            dataRowCount = dgvProducts.Rows.Count
        Else
            dataRowCount = dgvProducts.Rows.Count
        End If

        ' 如果是查看模式，不需要额外插入行（易语言中的insert+delete已简化）

        Dim sumQuantity As Decimal = 0
        Dim sumTotal As Decimal = 0
        Dim sumJinZhong As Decimal = 0
        Dim sumOther As Decimal = 0
        Dim sumAmount As Decimal = 0

        ' 统计数据行数：排除空行（最后一行空数据行）
        Dim calcCount As Integer = dataRowCount
        ' 在新建模式下，最后一行是空行（未填写产品名称的行），不参与统计
        If orderSelected = -1 AndAlso dataRowCount > 0 Then
            ' 检查最后一行是否为空行（产品名称为空）
            If String.IsNullOrEmpty(SafeString(dgvProducts.Rows(dataRowCount - 1).Cells(1).Value)) Then
                calcCount = dataRowCount - 1
            End If
        End If

        For i As Integer = 0 To calcCount - 1
            sumQuantity += SafeDecimal(dgvProducts.Rows(i).Cells(2).Value)
            sumTotal += SafeDecimal(dgvProducts.Rows(i).Cells(3).Value)
            sumJinZhong += SafeDecimal(dgvProducts.Rows(i).Cells(4).Value)
            sumOther += SafeDecimal(dgvProducts.Rows(i).Cells(7).Value)
            sumAmount += SafeDecimal(dgvProducts.Rows(i).Cells(8).Value)
        Next

        ' 添加合计行
        Dim sumRowIndex As Integer = dgvProducts.Rows.Add()
        dgvProducts.Rows(sumRowIndex).Cells(0).Value = ""
        dgvProducts.Rows(sumRowIndex).Cells(1).Value = "合计"
        dgvProducts.Rows(sumRowIndex).Cells(2).Value = sumQuantity.ToString()
        dgvProducts.Rows(sumRowIndex).Cells(3).Value = sumTotal.ToString()
        dgvProducts.Rows(sumRowIndex).Cells(4).Value = sumJinZhong.ToString()
        dgvProducts.Rows(sumRowIndex).Cells(7).Value = sumOther.ToString()
        dgvProducts.Rows(sumRowIndex).Cells(8).Value = sumAmount.ToString()

        ' 设置合计行只读和背景色
        For j As Integer = 0 To dgvProducts.Columns.Count - 1
            dgvProducts.Rows(sumRowIndex).Cells(j).ReadOnly = True
            dgvProducts.Rows(sumRowIndex).Cells(j).Style.BackColor = Drawing.Color.LightGray
            dgvProducts.Rows(sumRowIndex).Cells(j).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next

        ' 新建模式：设置空行（统计行）的只读和背景色
        If orderSelected = -1 AndAlso calcCount < dataRowCount Then
            Dim emptyRowIndex As Integer = calcCount
            For j As Integer = 2 To dgvProducts.Columns.Count - 2
                dgvProducts.Rows(emptyRowIndex).Cells(j).ReadOnly = True
                dgvProducts.Rows(emptyRowIndex).Cells(j).Style.BackColor = Drawing.Color.LightGray
            Next

            ' 计算税点、应付、实付
            Dim taxRate As Decimal = SafeDecimal(txtTaxRate.Text)
            Dim taxAmount As Decimal = Math.Round(sumAmount * taxRate / 100, 2)
            Dim payable As Decimal = Math.Round(sumAmount + sumAmount * taxRate / 100, 2)

            txtPayable.Text = payable.ToString()
            txtTaxAmount.Text = taxAmount.ToString()
            txtActualPay.Text = txtPayable.Text
        End If
    End Sub

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

        ' 如果行数不足，添加新行
        If dgvProducts.Rows.Count <= 3 Then
            AddProductRow()
        End If
    End Sub

    ' ========== 回收数据添加 ==========
    Private Sub AddProductRow()
        Dim currentRowCount As Integer = dgvProducts.Rows.Count

        If currentRowCount = 0 Then
            ' 第一次添加：数据行 + 合计行
            Dim dataIdx As Integer = dgvProducts.Rows.Add()
            dgvProducts.Rows(dataIdx).Cells(0).Value = "1"
            dgvProducts.Rows(dataIdx).Cells(1).Value = ""
            dgvProducts.Rows(dataIdx).Cells(2).Value = "0"
            dgvProducts.Rows(dataIdx).Cells(3).Value = ""
            dgvProducts.Rows(dataIdx).Cells(4).Value = ""
            dgvProducts.Rows(dataIdx).Cells(5).Value = ""
            dgvProducts.Rows(dataIdx).Cells(6).Value = ""
            dgvProducts.Rows(dataIdx).Cells(7).Value = ""
            dgvProducts.Rows(dataIdx).Cells(8).Value = ""
            dgvProducts.Rows(dataIdx).Cells(9).Value = ""

            ' 设置空行cols 2-9为灰色只读
            For j As Integer = 2 To dgvProducts.Columns.Count - 2
                dgvProducts.Rows(dataIdx).Cells(j).ReadOnly = True
                dgvProducts.Rows(dataIdx).Cells(j).Style.BackColor = Drawing.Color.LightGray
            Next

            ' 添加合计行
            Dim sumIdx As Integer = dgvProducts.Rows.Add()
            dgvProducts.Rows(sumIdx).Cells(0).Value = ""
            dgvProducts.Rows(sumIdx).Cells(1).Value = "合计"
            dgvProducts.Rows(sumIdx).Cells(2).Value = "0"

            For j As Integer = 0 To dgvProducts.Columns.Count - 1
                dgvProducts.Rows(sumIdx).Cells(j).ReadOnly = True
                dgvProducts.Rows(sumIdx).Cells(j).Style.BackColor = Drawing.Color.LightGray
            Next
        Else
            ' 后续添加：删除合计行，添加数据行，重新添加合计行
            ' 删除合计行
            dgvProducts.Rows.RemoveAt(dgvProducts.Rows.Count - 1)

            ' 添加新数据行
            Dim newIdx As Integer = dgvProducts.Rows.Add()
            dgvProducts.Rows(newIdx).Cells(0).Value = (newIdx + 1).ToString()
            dgvProducts.Rows(newIdx).Cells(1).Value = ""
            dgvProducts.Rows(newIdx).Cells(2).Value = "0"
            dgvProducts.Rows(newIdx).Cells(3).Value = ""
            dgvProducts.Rows(newIdx).Cells(4).Value = ""
            dgvProducts.Rows(newIdx).Cells(5).Value = ""
            dgvProducts.Rows(newIdx).Cells(6).Value = ""
            dgvProducts.Rows(newIdx).Cells(7).Value = ""
            dgvProducts.Rows(newIdx).Cells(8).Value = ""
            dgvProducts.Rows(newIdx).Cells(9).Value = ""

            ' 重新统计
            CalculateStatistics()
        End If

        ' 重新编号
        For i As Integer = 0 To dgvProducts.Rows.Count - 2
            If SafeString(dgvProducts.Rows(i).Cells(1).Value) <> "合计" Then
                dgvProducts.Rows(i).Cells(0).Value = (i + 1).ToString()
            End If
        Next
    End Sub

    ' ========== 表格将被编辑 ==========
    Private Sub DgvProducts_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs)
        editState = 1
    End Sub

    ' ========== 表格结束编辑 ==========
    Private Sub DgvProducts_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        editState = 0
        If e.RowIndex < 0 Then Return

        Dim editRow As Integer = e.RowIndex
        Dim editCol As Integer = e.ColumnIndex

        ' 列1：商品名称验证
        If editCol = 1 Then
            Dim productName As String = SafeString(dgvProducts.Rows(editRow).Cells(1).Value)
            If String.IsNullOrEmpty(productName) Then
                ShowWarning("回收名称不能为空！")
                Return
            End If

            ' 查询回收名称是否存在
            Dim queryName As String = ""
            Try
                Dim sql As String = "SELECT title FROM xipunum_erp_retreat_title WHERE (bianma = '" & SafeSQL(productName) & "' or title like '%" & SafeSQL(productName) & "%') LIMIT 1"
                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                If dt.Rows.Count > 0 Then
                    queryName = SafeString(dt.Rows(0)("title"))
                End If
            Catch
            End Try

            If String.IsNullOrEmpty(queryName) Then
                ShowWarning("回收名称不存在！")
                dgvProducts.Rows(editRow).Cells(1).Value = ""
                Return
            End If

            ' 设置回收名称
            dgvProducts.Rows(editRow).Cells(1).Value = queryName
            ' 设置默认值
            dgvProducts.Rows(editRow).Cells(2).Value = "1"
            dgvProducts.Rows(editRow).Cells(5).Value = "1"

            ' 设置导购员
            If editRow <= 0 Then
                dgvProducts.Rows(editRow).Cells(10).Value = txtSalesman.Text
            Else
                dgvProducts.Rows(editRow).Cells(10).Value = SafeString(dgvProducts.Rows(editRow - 1).Cells(10).Value)
            End If

            ' 解锁cols 2-10，设置白色背景
            For j As Integer = 2 To 10
                dgvProducts.Rows(editRow).Cells(j).ReadOnly = False
                dgvProducts.Rows(editRow).Cells(j).Style.BackColor = Drawing.Color.White
            Next

            ' 如果是最后一行数据行，添加新行
            If editRow = dgvProducts.Rows.Count - 2 Then
                AddProductRow()
            End If
        End If

        ' 列5：成色验证
        If editCol = 5 Then
            Dim chengse As String = SafeString(dgvProducts.Rows(editRow).Cells(5).Value)
            Dim chengseVal As Decimal = SafeDecimal(chengse)
            If chengseVal > 1 Then
                MessageBox.Show("成色不能大于1")
                dgvProducts.Rows(editRow).Cells(5).Value = "1"
                Return
            End If
            If chengseVal < 0 Then
                MessageBox.Show("成色不能小于0")
                dgvProducts.Rows(editRow).Cells(5).Value = "1"
                Return
            End If
        End If

        ' 列3,4,6,7,8：数值格式化和金额计算
        If editCol = 3 OrElse editCol = 4 OrElse editCol = 6 OrElse editCol = 7 OrElse editCol = 8 Then
            Dim totalWeight As String = SafeString(dgvProducts.Rows(editRow).Cells(3).Value)
            Dim jinZhong As String = SafeString(dgvProducts.Rows(editRow).Cells(4).Value)
            Dim price As String = SafeString(dgvProducts.Rows(editRow).Cells(6).Value)
            Dim otherFee As String = SafeString(dgvProducts.Rows(editRow).Cells(7).Value)
            Dim amount As String = SafeString(dgvProducts.Rows(editRow).Cells(8).Value)

            ' 如果编辑的是总重（列3），金重等于总重
            If editCol = 3 Then
                jinZhong = totalWeight
            End If

            ' 如果编辑的不是回收金额（列8），自动计算回收金额
            If editCol <> 8 Then
                amount = (SafeDecimal(jinZhong) * SafeDecimal(price) + SafeDecimal(otherFee)).ToString()
            End If

            ' 三位数处理：总重
            totalWeight = FormatThreeDecimal(totalWeight)
            ' 三位数处理：金重
            jinZhong = FormatThreeDecimal(jinZhong)
            ' 二位数处理：回收克价
            price = FormatTwoDecimal(price)
            ' 二位数处理：其他费用
            otherFee = FormatTwoDecimal(otherFee)
            ' 二位数处理：回收金额
            amount = FormatTwoDecimal(amount)

            ' 设置回单元格
            dgvProducts.Rows(editRow).Cells(3).Value = totalWeight
            dgvProducts.Rows(editRow).Cells(4).Value = jinZhong
            dgvProducts.Rows(editRow).Cells(6).Value = price
            dgvProducts.Rows(editRow).Cells(7).Value = otherFee
            dgvProducts.Rows(editRow).Cells(8).Value = amount

            ' 重新统计
            CalculateStatistics()
        End If

        ' 列10：导购员验证
        If editCol = 10 Then
            Dim guideName As String = SafeString(dgvProducts.Rows(editRow).Cells(10).Value)
            Dim foundName As String = ""

            If Not String.IsNullOrEmpty(guideName) Then
                Try
                    Dim sql As String = ""
                    If UserPermission = "全部" Then
                        sql = "SELECT name FROM xipunum_erp_user WHERE (jianxie = '" & SafeSQL(guideName) & "' or user like '%" & SafeSQL(guideName) & "%' or name like '%" & SafeSQL(guideName) & "%') and state='0' order by id ASC LIMIT 1"
                    Else
                        sql = "SELECT name FROM xipunum_erp_user WHERE (jianxie = '" & SafeSQL(guideName) & "' or user like '%" & SafeSQL(guideName) & "%' or name like '%" & SafeSQL(guideName) & "%') and department='" & SafeSQL(UserDepartment) & "' and state='0' order by id ASC LIMIT 1"
                    End If
                    Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
                    If dt.Rows.Count > 0 Then
                        foundName = SafeString(dt.Rows(0)("name"))
                    End If
                Catch
                End Try
            End If

            If String.IsNullOrEmpty(foundName) Then
                ' 未找到导购员，使用默认值
                If editRow <= 0 Then
                    dgvProducts.Rows(editRow).Cells(10).Value = txtSalesman.Text
                Else
                    dgvProducts.Rows(editRow).Cells(10).Value = SafeString(dgvProducts.Rows(editRow - 1).Cells(10).Value)
                End If
            Else
                dgvProducts.Rows(editRow).Cells(10).Value = foundName
            End If
        End If
    End Sub

    ' ========== 三位数处理 ==========
    Private Function FormatThreeDecimal(value As String) As String
        If String.IsNullOrEmpty(value) Then Return "0.000"
        Dim num As Decimal = SafeDecimal(value)
        Dim rounded As Decimal = Math.Round(num, 3)
        Return rounded.ToString("0.000")
    End Function

    ' ========== 二位数处理 ==========
    Private Function FormatTwoDecimal(value As String) As String
        If String.IsNullOrEmpty(value) Then Return "0.00"
        Dim num As Decimal = SafeDecimal(value)
        Dim rounded As Decimal = Math.Round(num, 2)
        Return rounded.ToString("0.00")
    End Function

    ' ========== 重置按钮 ==========
    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        Form_Load(Nothing, Nothing)
    End Sub

    ' ========== 税点改变 ==========
    Private Sub TxtTaxRate_TextChanged(sender As Object, e As EventArgs)
        Dim amountData As String = ""
        If dgvProducts.Rows.Count > 0 Then
            ' 获取合计行的回收金额
            Dim sumRowIdx As Integer = dgvProducts.Rows.Count - 1
            amountData = SafeString(dgvProducts.Rows(sumRowIdx).Cells(8).Value)
        End If

        Dim taxRate As Decimal = SafeDecimal(txtTaxRate.Text)
        Dim amount As Decimal = SafeDecimal(amountData)
        Dim taxAmount As Decimal = Math.Round(amount * taxRate / 100, 2)
        Dim payable As Decimal = Math.Round(amount + amount * taxRate / 100, 2)

        txtTaxAmount.Text = taxAmount.ToString()
        txtPayable.Text = payable.ToString()
        txtActualPay.Text = txtPayable.Text
    End Sub

    ' ========== 应付改变 ==========
    Private Sub TxtPayable_TextChanged(sender As Object, e As EventArgs)
        txtActualPay.Text = txtPayable.Text
    End Sub

    ' ========== 保存按钮 ==========
    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        InformationOperationDate = Now.ToString("yyyy-MM-dd HH:mm:ss")
        InformationOperationAccount = GBKToUTF8(UserAccount)

        ' 商品回收数量 = 总行数 - 2（空行 + 合计行）
        Dim productCount As Integer = dgvProducts.Rows.Count - 2

        If productCount <= 0 Then
            ShowWarning("回收数量不能为空！")
            Return
        End If

        If String.IsNullOrEmpty(txtActualPay.Text) Then
            ShowWarning("实付金额不能为空！")
            txtActualPay.Focus()
            Return
        End If

        If SafeDecimal(txtActualPay.Text) <= 0 Then
            ShowWarning("实付不能小于等于0元！")
            txtActualPay.Focus()
            Return
        End If

        ' 验证最后一个数据行（空行前的最后一行）
        Dim lastDataRowIdx As Integer = productCount - 1
        If lastDataRowIdx >= 0 Then
            If String.IsNullOrEmpty(SafeString(dgvProducts.Rows(lastDataRowIdx).Cells(2).Value)) Then
                ShowWarning("上一个回收数量不能为空！")
                Return
            End If
            If String.IsNullOrEmpty(SafeString(dgvProducts.Rows(lastDataRowIdx).Cells(3).Value)) Then
                ShowWarning("上一个回收总重不能为空！")
                Return
            End If
            If String.IsNullOrEmpty(SafeString(dgvProducts.Rows(lastDataRowIdx).Cells(4).Value)) Then
                ShowWarning("上一个回收金重不能为空！")
                Return
            End If
            If String.IsNullOrEmpty(SafeString(dgvProducts.Rows(lastDataRowIdx).Cells(6).Value)) Then
                ShowWarning("上一个回收回收克价不能为空！")
                Return
            End If
            If String.IsNullOrEmpty(SafeString(dgvProducts.Rows(lastDataRowIdx).Cells(7).Value)) Then
                ShowWarning("上一个回收其他费用不能为空！")
                Return
            End If
            If String.IsNullOrEmpty(SafeString(dgvProducts.Rows(lastDataRowIdx).Cells(8).Value)) Then
                ShowWarning("上一个回收金额不能为空！")
                Return
            End If
        End If

        If editState = 1 Then
            ShowWarning("请先结束回收列表操作！")
            Return
        End If

        ' 获取表单数据
        Dim orderNumber As String = txtOrderNumber.Text
        Dim salesman As String = txtSalesman.Text
        Dim memberPhone As String = txtPhone.Text
        Dim memberName As String = txtMemberName.Text

        ' 合计行数据
        Dim sumRowIdx As Integer = dgvProducts.Rows.Count - 1
        Dim totalWeight As String = SafeString(dgvProducts.Rows(sumRowIdx).Cells(3).Value)
        Dim jinZhong As String = SafeString(dgvProducts.Rows(sumRowIdx).Cells(4).Value)
        Dim otherFee As String = SafeString(dgvProducts.Rows(sumRowIdx).Cells(7).Value)
        Dim retreatAmount As String = SafeString(dgvProducts.Rows(sumRowIdx).Cells(8).Value)

        Dim taxRate As String = txtTaxRate.Text
        Dim taxAmount As String = txtTaxAmount.Text
        Dim payableAmount As String = txtPayable.Text
        Dim actualPay As String = txtActualPay.Text
        Dim remarks As String = txtRemarks.Text

        Dim savedCustomerCode As String = ""

        ' 检查会员是否存在
        If Not String.IsNullOrEmpty(memberPhone) Then
            Try
                Dim checkSql As String = "SELECT * FROM xipunum_erp_member where tel= '" & SafeSQL(memberPhone) & "'"
                Dim checkDt As DataTable = DatabaseModule.ExecuteQuery(checkSql, MySQL_Read)
                Dim memberExists As Integer = checkDt.Rows.Count

                If memberExists = 0 Then
                    ' 自动创建会员
                    Dim memberCode As String = "HY" & Now.ToString("yyyyMMddHHmmss")
                    Dim insertMemberSql As String = "INSERT INTO xipunum_erp_member (customer_code, name, tel, cjuser, creationtime) VALUES ('" &
                        SafeSQL(memberCode) & "','" & SafeSQL(memberName) & "','" & SafeSQL(memberPhone) & "','" &
                        SafeSQL(InformationOperationAccount) & "','" & SafeSQL(InformationOperationDate) & "')"
                    DatabaseModule.ExecuteCommand(insertMemberSql, MySQL_Write)
                End If

                ' 获取客户编码
                Dim codeSql As String = "SELECT customer_code FROM xipunum_erp_member where tel='" & SafeSQL(memberPhone) & "' order by id ASC LIMIT 1"
                Dim codeDt As DataTable = DatabaseModule.ExecuteQuery(codeSql, MySQL_Read)
                If codeDt.Rows.Count > 0 Then
                    Dim customerCode As String = SafeString(codeDt.Rows(0)("customer_code"))
                    savedCustomerCode = ",customer_code='" & SafeSQL(customerCode) & "'"
                End If
            Catch ex As Exception
                ShowError("会员处理失败：" & ex.Message)
                Return
            End Try
        End If

        ' 插入回收订单
        Try
            Dim insertOrderSql As String = "INSERT INTO xipunum_erp_retreat_order (retrea_umber" & savedCustomerCode &
                ",total,jin_zhong,qita_price,tax_rate,rate_amount,retreat_amount,ying_amount,settlement,salesman,remarks,cjuser,creationtime) VALUES ('" &
                SafeSQL(orderNumber) & "','" & SafeSQL(totalWeight) & "','" & SafeSQL(jinZhong) & "','" &
                SafeSQL(otherFee) & "','" & SafeSQL(taxRate) & "','" & SafeSQL(taxAmount) & "','" &
                SafeSQL(retreatAmount) & "','" & SafeSQL(payableAmount) & "','" & SafeSQL(actualPay) & "','" &
                SafeSQL(salesman) & "','" & SafeSQL(remarks) & "','" & SafeSQL(InformationOperationAccount) & "','" &
                SafeSQL(InformationOperationDate) & "')"
            DatabaseModule.ExecuteCommand(insertOrderSql, MySQL_Write)
        Catch ex As Exception
            ShowError("保存回收订单失败：" & ex.Message)
            Return
        End Try

        ' 插入收款记录
        Try
            Dim insertShoukuanSql As String = "INSERT INTO xipunum_erp_shoukuan (leibie,settlement_number,xianjin,type,kufang,cjuser,creationtime) VALUES ('2','" &
                SafeSQL(orderNumber) & "','" & SafeSQL(actualPay) & "','1','" & SafeSQL(UserDepartment) & "','" &
                SafeSQL(InformationOperationAccount) & "','" & SafeSQL(InformationOperationDate) & "')"
            DatabaseModule.ExecuteCommand(insertShoukuanSql, MySQL_Write)
        Catch ex As Exception
        End Try

        ' 插入系统日志
        LogSaveContent = "账户:" & UserAccount & " 商品回收，回收单号:" & orderNumber
        Try
            Dim insertLogSql As String = "INSERT INTO xipunum_erp_xitong_log (type,title,conter,user,creationtime) VALUES ('添加','商品回收','" &
                SafeSQL(LogSaveContent) & "','" & SafeSQL(InformationOperationAccount) & "','" & SafeSQL(InformationOperationDate) & "')"
            DatabaseModule.ExecuteCommand(insertLogSql, MySQL_Write)
        Catch
        End Try

        ' 获取回收订单ID
        Dim orderId As String = ""
        Try
            Dim orderSql As String = "SELECT id FROM xipunum_erp_retreat_order where retrea_umber='" & SafeSQL(orderNumber) & "' order by id ASC LIMIT 1"
            Dim orderDt As DataTable = DatabaseModule.ExecuteQuery(orderSql, MySQL_Read)
            If orderDt.Rows.Count > 0 Then
                orderId = SafeString(orderDt.Rows(0)("id"))
            End If
        Catch
        End Try

        ' 循环保存商品明细
        For i As Integer = 0 To productCount - 1
            Try
                Dim productName As String = SafeString(dgvProducts.Rows(i).Cells(1).Value)
                Dim quantity As String = SafeString(dgvProducts.Rows(i).Cells(2).Value)
                Dim productTotal As String = SafeString(dgvProducts.Rows(i).Cells(3).Value)
                Dim productJinZhong As String = SafeString(dgvProducts.Rows(i).Cells(4).Value)
                Dim chengse As String = SafeString(dgvProducts.Rows(i).Cells(5).Value)
                Dim productPrice As String = SafeString(dgvProducts.Rows(i).Cells(6).Value)
                Dim productOtherFee As String = SafeString(dgvProducts.Rows(i).Cells(7).Value)
                Dim productAmount As String = SafeString(dgvProducts.Rows(i).Cells(8).Value)
                Dim productRemarks As String = SafeString(dgvProducts.Rows(i).Cells(9).Value)
                Dim guideName As String = SafeString(dgvProducts.Rows(i).Cells(10).Value)

                ' 成色处理
                If chengse = "1" Then
                    chengse = "1.0000"
                Else
                    chengse = Math.Round(SafeDecimal(chengse), 4).ToString()
                End If

                ' 获取导购员账户
                Dim guideAccount As String = ""
                Try
                    Dim guideSql As String = "SELECT user FROM xipunum_erp_user WHERE name = '" & SafeSQL(guideName) & "' and department = '" & SafeSQL(UserDepartment) & "' LIMIT 1"
                    Dim guideDt As DataTable = DatabaseModule.ExecuteQuery(guideSql, MySQL_Read)
                    If guideDt.Rows.Count > 0 Then
                        guideAccount = SafeString(guideDt.Rows(0)("user"))
                    End If
                Catch
                End Try

                ' 获取回收名称ID
                Dim productId As String = ""
                Try
                    Dim titleSql As String = "SELECT id FROM xipunum_erp_retreat_title WHERE title='" & SafeSQL(productName) & "' LIMIT 1"
                    Dim titleDt As DataTable = DatabaseModule.ExecuteQuery(titleSql, MySQL_Read)
                    If titleDt.Rows.Count > 0 Then
                        productId = SafeString(titleDt.Rows(0)("id"))
                    End If
                Catch
                End Try

                If String.IsNullOrEmpty(productId) Then
                    productId = "1"
                End If

                ' 插入回收明细
                Dim insertDetailSql As String = "INSERT INTO xipunum_erp_retreat (order_id,product_name,quantity,total,jin_zhong,chengse,price,qita_price,retreat_amount,huishoutime,shopping_guide,remarks,cjuser,creationtime) VALUES ('" &
                    SafeSQL(orderId) & "','" & SafeSQL(productId) & "','" & SafeSQL(quantity) & "','" & SafeSQL(productTotal) & "','" &
                    SafeSQL(productJinZhong) & "','" & SafeSQL(chengse) & "','" & SafeSQL(productPrice) & "','" &
                    SafeSQL(productOtherFee) & "','" & SafeSQL(productAmount) & "','" & SafeSQL(InformationOperationDate) & "','" &
                    SafeSQL(guideAccount) & "','" & SafeSQL(productRemarks) & "','" & SafeSQL(InformationOperationAccount) & "','" &
                    SafeSQL(InformationOperationDate) & "')"
                DatabaseModule.ExecuteCommand(insertDetailSql, MySQL_Write)

                ' 插入系统日志
                LogSaveContent = "账户:" & UserAccount & " 商品回收，名称：" & productName
                Try
                    Dim insertDetailLogSql As String = "INSERT INTO xipunum_erp_xitong_log (type,title,conter,user,creationtime) VALUES ('添加','商品回收','" &
                        SafeSQL(LogSaveContent) & "','" & SafeSQL(InformationOperationAccount) & "','" & SafeSQL(InformationOperationDate) & "')"
                    DatabaseModule.ExecuteCommand(insertDetailLogSql, MySQL_Write)
                Catch
                End Try

                Threading.Thread.Sleep(50)
            Catch
            End Try
        Next

        ShowSuccess("回收商品操作完成！")
        Me.Close()

        ' 刷新主表格
        If HomePageQueryText = "商品回收" Then
            Try
                Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
                If mainForm IsNot Nothing Then
                    mainForm.RefreshSubFolderTable()
                End If
            Catch
            End Try
        End If
    End Sub

End Class

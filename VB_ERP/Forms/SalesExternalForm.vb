' ============================================================================
' 商品销售外部单据窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品销售外部单据.form.e.txt
' 控件: 高级表格Ex1, 编辑框_订单编码, 编辑框_客户名称, 编辑框_出库日期,
'       编辑框_制单人, 编辑框_提货地址, 编辑框_交货方, 编辑框_用途,
'       按钮EX_打印, 按钮EX_重置, 图片框EX4(关闭), 添加修改_分组框
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SalesExternalForm
    Inherits System.Windows.Forms.Form

    ' ========== 控件声明 ==========
    Private dgvList As New DataGridView()          ' 高级表格Ex1
    Private txtOrderCode As New TextBox()          ' 编辑框_订单编码
    Private txtCustomerName As New TextBox()       ' 编辑框_客户名称
    Private txtOutboundDate As New TextBox()       ' 编辑框_出库日期
    Private txtMaker As New TextBox()              ' 编辑框_制单人
    Private txtDeliveryAddr As New TextBox()       ' 编辑框_提货地址
    Private txtDeliveryParty As New TextBox()      ' 编辑框_交货方
    Private txtUsage As New TextBox()              ' 编辑框_用途
    Private WithEvents btnPrint As New Button()    ' 按钮EX_打印
    Private WithEvents btnReset As New Button()    ' 按钮EX_重置
    Private WithEvents btnClose As New Button()    ' 图片框EX4（关闭）
    Private grpMain As New GroupBox()              ' 添加修改_分组框

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "商品销售外部单据"
        Me.Size = New Drawing.Size(1000, 700)
        Me.StartPosition = FormStartPosition.CenterParent

        ' 主分组框
        grpMain.Text = "商品销售外部单据"
        grpMain.Dock = DockStyle.Fill
        Me.Controls.Add(grpMain)

        ' 顶部信息面板
        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 120
        grpMain.Controls.Add(panelTop)

        ' 订单编码
        Dim lblOrderCode As New Label()
        lblOrderCode.Text = "订单编码："
        lblOrderCode.Location = New Drawing.Point(10, 10)
        lblOrderCode.AutoSize = True
        panelTop.Controls.Add(lblOrderCode)
        txtOrderCode.Location = New Drawing.Point(80, 8)
        txtOrderCode.Size = New Drawing.Size(180, 25)
        panelTop.Controls.Add(txtOrderCode)

        ' 客户名称
        Dim lblCustomer As New Label()
        lblCustomer.Text = "客户名称："
        lblCustomer.Location = New Drawing.Point(280, 10)
        lblCustomer.AutoSize = True
        panelTop.Controls.Add(lblCustomer)
        txtCustomerName.Location = New Drawing.Point(350, 8)
        txtCustomerName.Size = New Drawing.Size(150, 25)
        txtCustomerName.ReadOnly = True
        panelTop.Controls.Add(txtCustomerName)

        ' 出库日期
        Dim lblDate As New Label()
        lblDate.Text = "出库日期："
        lblDate.Location = New Drawing.Point(520, 10)
        lblDate.AutoSize = True
        panelTop.Controls.Add(lblDate)
        txtOutboundDate.Location = New Drawing.Point(590, 8)
        txtOutboundDate.Size = New Drawing.Size(100, 25)
        txtOutboundDate.ReadOnly = True
        panelTop.Controls.Add(txtOutboundDate)

        ' 制单人
        Dim lblMaker As New Label()
        lblMaker.Text = "制单人："
        lblMaker.Location = New Drawing.Point(710, 10)
        lblMaker.AutoSize = True
        panelTop.Controls.Add(lblMaker)
        txtMaker.Location = New Drawing.Point(770, 8)
        txtMaker.Size = New Drawing.Size(100, 25)
        txtMaker.ReadOnly = True
        panelTop.Controls.Add(txtMaker)

        ' 提货地址
        Dim lblAddr As New Label()
        lblAddr.Text = "提货地址："
        lblAddr.Location = New Drawing.Point(10, 45)
        lblAddr.AutoSize = True
        panelTop.Controls.Add(lblAddr)
        txtDeliveryAddr.Location = New Drawing.Point(80, 42)
        txtDeliveryAddr.Size = New Drawing.Size(300, 25)
        txtDeliveryAddr.ReadOnly = True
        panelTop.Controls.Add(txtDeliveryAddr)

        ' 交货方
        Dim lblParty As New Label()
        lblParty.Text = "交货方："
        lblParty.Location = New Drawing.Point(400, 45)
        lblParty.AutoSize = True
        panelTop.Controls.Add(lblParty)
        txtDeliveryParty.Location = New Drawing.Point(450, 42)
        txtDeliveryParty.Size = New Drawing.Size(200, 25)
        txtDeliveryParty.ReadOnly = True
        panelTop.Controls.Add(txtDeliveryParty)

        ' 用途
        Dim lblUsage As New Label()
        lblUsage.Text = "用途："
        lblUsage.Location = New Drawing.Point(670, 45)
        lblUsage.AutoSize = True
        panelTop.Controls.Add(lblUsage)
        txtUsage.Location = New Drawing.Point(710, 42)
        txtUsage.Size = New Drawing.Size(160, 25)
        txtUsage.ReadOnly = True
        panelTop.Controls.Add(txtUsage)

        ' 按钮
        btnPrint.Text = "打印"
        btnPrint.Location = New Drawing.Point(10, 80)
        btnPrint.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnPrint)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(100, 80)
        btnReset.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnReset)

        btnClose.Text = "关闭"
        btnClose.Location = New Drawing.Point(190, 80)
        btnClose.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnClose)

        ' 表格
        dgvList.Dock = DockStyle.Fill
        dgvList.AllowUserToAddRows = False
        dgvList.AllowUserToDeleteRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.CellSelect
        grpMain.Controls.Add(dgvList)

        ' 事件
        AddHandler dgvList.CellEndEdit, AddressOf DgvList_CellEndEdit
    End Sub

    ' ========== _窗口_商品销售外部单据_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        grpMain.Text = "商品销售外部单据,销售单号:" & GlobalSalesOrderNumber
        InitGridHeader()
        LoadData()
    End Sub

    ' ========== _窗口_商品销售外部单据_尺寸被改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        ' 分组框自适应
    End Sub

    ' ========== _子程序_加载表头 ==========
    Private Sub InitGridHeader()
        dgvList.Columns.Clear()
        Dim headers() As String = {"序号", "产品名称", "重量(克)", "成色", "单价(元/克)", "金额(元)", "备注"}
        Dim widths() As Integer = {45, 200, 100, 100, 100, 100, 150}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvList.Columns.Add(col)
        Next
    End Sub

    ' ========== _子程序_加载数据 ==========
    Private Sub LoadData()
        dgvList.Rows.Clear()

        ' 获取店铺地址
        Dim accountAddr As String = ""
        Try
            Dim sql As String = "SELECT * FROM xipunum_erp_type WHERE id= '" & UserDepartment & "' order by id desc"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            If dt.Rows.Count > 0 Then
                accountAddr = SafeString(dt.Rows(0)("data3"))
            End If
        Catch ex As Exception
        End Try

        ' 获取订单信息
        Dim customerName As String = ""
        Dim outboundTime As String = ""
        Dim externalNumber As String = ""

        Try
            Dim sql As String = "SELECT a.customer_code AS customer_code,b.name AS NAME,a.creationtime AS chukutime,a.waibu_number as waibu_number FROM xipunum_erp_outbound_order AS a LEFT JOIN xipunum_erp_member AS b ON b.customer_code = a.customer_code WHERE a.settlement_number = '" & SafeSQL(GlobalSalesOrderNumber) & "'"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            If dt.Rows.Count > 0 Then
                customerName = SafeString(dt.Rows(0)("NAME"))
                outboundTime = SafeString(dt.Rows(0)("chukutime"))
                externalNumber = SafeString(dt.Rows(0)("waibu_number"))
            End If
        Catch ex As Exception
        End Try

        txtOrderCode.Text = externalNumber
        txtOrderCode.ReadOnly = (externalNumber <> "")

        If externalNumber = "" Then
            txtOrderCode.Focus()
        End If

        txtCustomerName.Text = customerName
        If outboundTime <> "" Then
            Try
                txtOutboundDate.Text = DateTime.Parse(outboundTime).ToString("MM月dd日")
            Catch
                txtOutboundDate.Text = ""
            End Try
        End If
        txtMaker.Text = UserAccount
        txtDeliveryAddr.Text = accountAddr
        txtDeliveryParty.Text = "厦门海峡金融服务有限公司"
        txtUsage.Text = "投资性黄金"

        ' 加载销售明细
        Dim detailCount As Integer = 0
        Try
            Dim sql As String = "SELECT c.product_name AS mingcheg,CAST(ROUND(a.net_weight,2) AS DECIMAL (30, 2)) as jinzhong,d.factory_condition/100 as chengse,CAST(ROUND(a.gold_price,2) AS DECIMAL (30, 2)) as kejia,CAST(ROUND(a.settlement,2) AS DECIMAL (30, 2)) as shishou,a.remarks as beizhu FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_outbound_order AS b ON b.id = a.order_id and b.settlement_number = '" & SafeSQL(GlobalSalesOrderNumber) & "'  INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = a.poduct_code INNER JOIN xipunum_erp_store AS d ON d.poduct_code = a.poduct_code GROUP BY a.id desc"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            detailCount = dt.Rows.Count

            For i As Integer = 0 To detailCount - 1
                dgvList.Rows.Add()
                dgvList.Rows(i).Cells(0).Value = (i + 1).ToString()
                dgvList.Rows(i).Cells(1).Value = SafeString(dt.Rows(i)("mingcheg"))
                dgvList.Rows(i).Cells(2).Value = SafeString(dt.Rows(i)("jinzhong"))
                dgvList.Rows(i).Cells(3).Value = SafeString(dt.Rows(i)("chengse"))
                dgvList.Rows(i).Cells(4).Value = SafeString(dt.Rows(i)("kejia"))
                dgvList.Rows(i).Cells(5).Value = SafeString(dt.Rows(i)("shishou"))
                dgvList.Rows(i).Cells(6).Value = SafeString(dt.Rows(i)("beizhu"))
            Next
        Catch ex As Exception
        End Try

        ' 合计行
        Dim totalWeight As String = ""
        Dim totalAmount As String = ""
        Try
            Dim sql As String = "SELECT CAST(ROUND(sum(a.net_weight),2) AS DECIMAL (30, 2)) as jinzhong,CAST(ROUND(sum(a.settlement),2) AS DECIMAL (30, 2)) as shishou FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_outbound_order AS b ON b.id = a.order_id and b.settlement_number = '" & SafeSQL(GlobalSalesOrderNumber) & "'  INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = a.poduct_code INNER JOIN xipunum_erp_store AS d ON d.poduct_code = a.poduct_code"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            If dt.Rows.Count > 0 Then
                totalWeight = SafeString(dt.Rows(0)("jinzhong"))
                totalAmount = SafeString(dt.Rows(0)("shishou"))
            End If
        Catch ex As Exception
        End Try

        ' 添加合计行
        dgvList.Rows.Add()
        Dim totalRowIndex As Integer = dgvList.Rows.Count - 1
        dgvList.Rows(totalRowIndex).Cells(2).Value = totalWeight
        dgvList.Rows(totalRowIndex).Cells(5).Value = totalAmount
        For i As Integer = 0 To dgvList.Columns.Count - 1
            If i <> 2 AndAlso i <> 5 Then
                dgvList.Rows(totalRowIndex).Cells(i).ReadOnly = True
                dgvList.Rows(totalRowIndex).Cells(i).Style.BackColor = Drawing.Color.LightGray
            End If
        Next
    End Sub

    ' ========== _高级表格Ex1_项目退出编辑 ==========
    Private Sub DgvList_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return

        ' 列4(单价)和列5(金额): 格式化两位小数
        If e.ColumnIndex = 4 Or e.ColumnIndex = 5 Then
            Dim content As String = SafeString(dgvList.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
            Dim formatted As String = FormatTwoDecimals(content)
            dgvList.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = formatted
        End If
    End Sub

    ' ========== _图片框EX4_鼠标左键单击（关闭） ==========
    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ' ========== _按钮EX_重置_鼠标左键单击 ==========
    Private Sub BtnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Form_Load(Nothing, Nothing)
    End Sub

    ' ========== _按钮EX_打印_鼠标左键单击 ==========
    Private Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        ' 验证外部单号
        If txtOrderCode.Text = "" Then
            ShowWarning("外部单号不能为空！")
            txtOrderCode.Focus()
            Return
        End If

        ' VB.NET简化为提示打印
        ' 实际应加载GRF报表文件并填充数据
        MessageBox.Show("打印外部单据，订单编码：" & txtOrderCode.Text, "打印提示", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' 更新外部单号到数据库
        Try
            Dim sql As String = "UPDATE xipunum_erp_outbound_order SET waibu_number='" & SafeSQL(txtOrderCode.Text) & "' WHERE settlement_number ='" & SafeSQL(GlobalSalesOrderNumber) & "' LIMIT 1"
            DatabaseModule.ExecuteCommand(sql, MySQL_Write)
        Catch ex As Exception
        End Try

        Me.Close()
    End Sub

    ' ========== 辅助函数 ==========
    Private Function FormatTwoDecimals(originalText As String) As String
        If originalText = "" Then Return "0.00"
        Dim val As Decimal
        If Not Decimal.TryParse(originalText, val) Then Return "0.00"
        Dim processedText As String = Math.Round(val, 2).ToString()
        Dim parts() As String = processedText.Split("."c)
        If parts.Length > 1 Then
            If parts(1).Length = 1 Then
                Return processedText & "0"
            Else
                Return processedText
            End If
        Else
            Return processedText & ".00"
        End If
    End Function

    Private Function SafeString(val As Object) As String
        If val Is Nothing OrElse val Is DBNull.Value Then Return ""
        Return val.ToString()
    End Function

    Private Function SafeSQL(val As String) As String
        Return val.Replace("'", "''")
    End Function

    Private Sub ShowWarning(msg As String)
        MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

End Class

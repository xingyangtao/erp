' ============================================================================
' 入库工厂修改窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_入库工厂修改.form.e.txt
' 包含3个程序集变量、8个子程序
' 功能：修改入库订单的工厂/成本工费/成本附加费/原料克价，批量重算商品成本
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class InboundFactoryEditForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（3个） ==========
    Private mainWindowRowIndex As Integer = -1      ' 局部_主窗口行号
    Private orderId As String = ""                  ' 局部_订单id
    Private isSaving As Boolean = False             ' 局部_正在保存

    ' ========== 控件声明 ==========
    Private grpMain As New GroupBox()               ' 添加修改_分组框
    Private txtOrderCode As New TextBox()           ' 编辑框_订单编码
    Private txtOriginalFactory As New TextBox()     ' 编辑框_原入库工厂
    Private cmbNewFactory As New ComboBox()         ' 组合框新入库工厂
    Private txtBasicCost As New TextBox()           ' 编辑框_成本工费
    Private txtCompanySurcharge As New TextBox()    ' 编辑框_成本附加费
    Private txtGoldPrice As New TextBox()           ' 编辑框_原料克价
    Private btnSave As New Button()                 ' 按钮EX1（保存）
    Private btnReset As New Button()                ' 按钮EX2（重置）
    Private btnCancel As New Button()               ' 图片框EX4（关闭）

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "入库工厂修改"
        Me.Size = New Drawing.Size(500, 400)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 分组框
        grpMain.Text = "入库工厂修改"
        grpMain.Size = New Drawing.Size(420, 320)
        grpMain.Location = New Drawing.Point(30, 20)
        Me.Controls.Add(grpMain)

        ' 订单编码
        Dim lblOrderCode As New Label()
        lblOrderCode.Text = "订单编码："
        lblOrderCode.Location = New Drawing.Point(20, 25)
        lblOrderCode.AutoSize = True
        grpMain.Controls.Add(lblOrderCode)
        txtOrderCode.Location = New Drawing.Point(100, 22)
        txtOrderCode.Size = New Drawing.Size(200, 25)
        txtOrderCode.ReadOnly = True
        grpMain.Controls.Add(txtOrderCode)

        ' 原入库工厂
        Dim lblOrigFactory As New Label()
        lblOrigFactory.Text = "原入库工厂："
        lblOrigFactory.Location = New Drawing.Point(20, 60)
        lblOrigFactory.AutoSize = True
        grpMain.Controls.Add(lblOrigFactory)
        txtOriginalFactory.Location = New Drawing.Point(100, 57)
        txtOriginalFactory.Size = New Drawing.Size(200, 25)
        txtOriginalFactory.ReadOnly = True
        grpMain.Controls.Add(txtOriginalFactory)

        ' 新入库工厂
        Dim lblNewFactory As New Label()
        lblNewFactory.Text = "新入库工厂："
        lblNewFactory.Location = New Drawing.Point(20, 95)
        lblNewFactory.AutoSize = True
        grpMain.Controls.Add(lblNewFactory)
        cmbNewFactory.Location = New Drawing.Point(100, 92)
        cmbNewFactory.Size = New Drawing.Size(200, 25)
        cmbNewFactory.DropDownStyle = ComboBoxStyle.DropDown
        grpMain.Controls.Add(cmbNewFactory)

        ' 成本工费
        Dim lblBasicCost As New Label()
        lblBasicCost.Text = "成本工费："
        lblBasicCost.Location = New Drawing.Point(20, 130)
        lblBasicCost.AutoSize = True
        grpMain.Controls.Add(lblBasicCost)
        txtBasicCost.Location = New Drawing.Point(100, 127)
        txtBasicCost.Size = New Drawing.Size(200, 25)
        grpMain.Controls.Add(txtBasicCost)

        ' 成本附加费
        Dim lblSurcharge As New Label()
        lblSurcharge.Text = "成本附加费："
        lblSurcharge.Location = New Drawing.Point(20, 165)
        lblSurcharge.AutoSize = True
        grpMain.Controls.Add(lblSurcharge)
        txtCompanySurcharge.Location = New Drawing.Point(100, 162)
        txtCompanySurcharge.Size = New Drawing.Size(200, 25)
        grpMain.Controls.Add(txtCompanySurcharge)

        ' 原料克价
        Dim lblGoldPrice As New Label()
        lblGoldPrice.Text = "原料克价："
        lblGoldPrice.Location = New Drawing.Point(20, 200)
        lblGoldPrice.AutoSize = True
        grpMain.Controls.Add(lblGoldPrice)
        txtGoldPrice.Location = New Drawing.Point(100, 197)
        txtGoldPrice.Size = New Drawing.Size(200, 25)
        grpMain.Controls.Add(txtGoldPrice)

        ' 按钮
        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(60, 250)
        btnSave.Size = New Drawing.Size(90, 35)
        grpMain.Controls.Add(btnSave)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(165, 250)
        btnReset.Size = New Drawing.Size(90, 35)
        grpMain.Controls.Add(btnReset)

        btnCancel.Text = "关闭"
        btnCancel.Location = New Drawing.Point(270, 250)
        btnCancel.Size = New Drawing.Size(90, 35)
        grpMain.Controls.Add(btnCancel)

        ' 事件绑定
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnReset.Click, AddressOf BtnReset_Click
        AddHandler btnCancel.Click, AddressOf BtnCancel_Click
        AddHandler cmbNewFactory.DropDown, AddressOf CmbNewFactory_DropDown
        AddHandler cmbNewFactory.SelectedIndexChanged, AddressOf CmbNewFactory_SelectedIndexChanged
    End Sub

    ' ========== _窗口_入库工厂修改_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        isSaving = False

        ' 从全局变量获取订单信息（由调用方设置）
        If orderId = "" Then
            ShowWarning("请先在主窗口选择一条入库订单！")
            Me.Close()
            Return
        End If

        cmbNewFactory.SelectedIndex = -1
        cmbNewFactory.Text = "请选择工厂"
        txtBasicCost.Text = ""
        txtCompanySurcharge.Text = ""
        txtGoldPrice.Text = ""
    End Sub

    ' ========== 设置订单信息 ==========
    Public Sub SetOrderInfo(id As String, code As String, originalFactory As String)
        orderId = id
        txtOrderCode.Text = code
        txtOriginalFactory.Text = originalFactory
    End Sub

    ' ========== _窗口_入库工厂修改_尺寸被改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        grpMain.Left = (Me.Width - grpMain.Width) \ 2
        grpMain.Top = (Me.Height - grpMain.Height) \ 2
    End Sub

    ' ========== _组合框新入库工厂_项目内容改变（搜索工厂） ==========
    Private Sub CmbNewFactory_DropDown(sender As Object, e As EventArgs)
        Dim searchText As String = cmbNewFactory.Text
        If searchText = "请选择工厂" Then searchText = ""

        cmbNewFactory.Items.Clear()

        Dim sql As String = "SELECT id,title FROM xipunum_erp_about " &
            "WHERE (title LIKE '%" & SafeSQL(searchText) & "%' OR jianxie LIKE '%" & SafeSQL(searchText) & "%') " &
            "ORDER BY id ASC"
        Dim dt As DataTable = ExecuteQuery(sql, MySQL_Read)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Dim factoryId As String = SafeString(row("id"))
                Dim factoryName As String = SafeString(row("title"))
                cmbNewFactory.Items.Add(factoryName & "|" & factoryId)
            Next
        End If
    End Sub

    ' ========== _组合框新入库工厂_内容被改变 ==========
    Private Sub CmbNewFactory_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' 选中后只显示工厂名称（不含ID）
        If cmbNewFactory.SelectedIndex >= 0 Then
            Dim selectedItem As String = cmbNewFactory.SelectedItem.ToString()
            If selectedItem.Contains("|") Then
                cmbNewFactory.Text = selectedItem.Split("|"c)(0)
            End If
        End If
    End Sub

    ' ========== 获取选中工厂的ID ==========
    Private Function GetSelectedFactoryId() As String
        If cmbNewFactory.SelectedIndex < 0 Then Return ""
        Dim selectedItem As String = cmbNewFactory.Items(cmbNewFactory.SelectedIndex).ToString()
        If selectedItem.Contains("|") Then
            Return selectedItem.Split("|"c)(1)
        End If
        Return ""
    End Function

    ' ========== _按钮EX1_鼠标左键单击（保存） ==========
    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        If isSaving Then Return
        If orderId = "" Then
            ShowWarning("请先选择入库订单！")
            Return
        End If

        isSaving = True
        Dim operationDate As String = GetOperationDate()
        Dim operationAccount As String = GetOperationAccount()

        Dim orderNumber As String = txtOrderCode.Text
        Dim inputBasicCost As String = txtBasicCost.Text
        Dim inputCompanySurcharge As String = txtCompanySurcharge.Text
        Dim inputGoldPrice As String = txtGoldPrice.Text

        If orderNumber = "" Then
            ShowWarning("入库单号不能为空！")
            isSaving = False
            Return
        End If

        ' 判断哪些有变更
        Dim factoryChanged As Boolean = False
        Dim laborChanged As Boolean = (SafeDecimal(inputBasicCost) <> 0)
        Dim surchargeChanged As Boolean = (SafeDecimal(inputCompanySurcharge) <> 0)
        Dim goldPriceChanged As Boolean = (SafeDecimal(inputGoldPrice) <> 0)
        Dim newFactoryName As String = ""

        If cmbNewFactory.SelectedIndex >= 0 Then
            newFactoryName = cmbNewFactory.Text
            If newFactoryName <> txtOriginalFactory.Text Then
                factoryChanged = True
            End If
        End If

        If Not factoryChanged AndAlso Not laborChanged AndAlso Not surchargeChanged AndAlso Not goldPriceChanged Then
            ShowWarning("请填写需要修改的入库信息！")
            isSaving = False
            Return
        End If

        Try
            ' 1. 工厂变更
            If factoryChanged Then
                Dim factoryId As String = GetSelectedFactoryId()
                ExecuteCommand("UPDATE xipunum_erp_store_order SET factory='" & SafeSQL(factoryId) & "',updatetime='" & operationDate & "' WHERE odd_numbers='" & SafeSQL(orderNumber) & "' LIMIT 1", MySQL_Write)
                ExecuteCommand("INSERT INTO xipunum_erp_store_log(order_id,user,conter,creationtime) VALUES('" & SafeSQL(orderId) & "','" & SafeSQL(operationAccount) & "','修改工厂','" & operationDate & "')", MySQL_Write)
                LogContent = "账户:" & UserAccount & " 修改入库订单：" & orderNumber & " 工厂名称由：" & txtOriginalFactory.Text & "->" & newFactoryName
                ExecuteCommand("INSERT INTO xipunum_erp_xitong_log(type,title,conter,user,creationtime) VALUES('修改','修改入库订单','" & SafeSQL(LogContent) & "','" & SafeSQL(operationAccount) & "','" & operationDate & "')", MySQL_Write)
                txtOriginalFactory.Text = newFactoryName
            End If

            ' 2. 原料价变更
            If goldPriceChanged Then
                inputGoldPrice = FormatTwoDecimals(inputGoldPrice)
                ExecuteCommand("UPDATE xipunum_erp_store_order SET gold_price='" & SafeSQL(inputGoldPrice) & "',updatetime='" & operationDate & "' WHERE odd_numbers='" & SafeSQL(orderNumber) & "' LIMIT 1", MySQL_Write)
                ExecuteCommand("INSERT INTO xipunum_erp_store_log(order_id,user,conter,creationtime) VALUES('" & SafeSQL(orderId) & "','" & SafeSQL(operationAccount) & "','修改原料价','" & operationDate & "')", MySQL_Write)
                LogContent = "账户:" & UserAccount & " 修改入库订单：" & orderNumber & " 原料价修改为：" & inputGoldPrice
                ExecuteCommand("INSERT INTO xipunum_erp_xitong_log(type,title,conter,user,creationtime) VALUES('修改','修改入库订单','" & SafeSQL(LogContent) & "','" & SafeSQL(operationAccount) & "','" & operationDate & "')", MySQL_Write)
            End If

            ' 3. 工费/附加费/原料价变更：批量重算商品成本
            If laborChanged OrElse surchargeChanged OrElse goldPriceChanged Then
                If laborChanged Then
                    inputBasicCost = FormatTwoDecimals(inputBasicCost)
                End If
                If surchargeChanged Then
                    inputCompanySurcharge = FormatTwoDecimals(inputCompanySurcharge)
                End If
                If goldPriceChanged Then
                    inputGoldPrice = FormatTwoDecimals(inputGoldPrice)
                End If

                ' 查询入库明细
                Dim detailSql As String = "SELECT b.gold_price AS gold_price,a.id AS rukuid,a.poduct_code AS bianam," &
                    "c.jin_zhong AS jinzhong,c.quantity AS rukushu,a.basic_cost AS chengbengf," &
                    "a.company_surcharge AS chengebnfj,a.cost_price AS chengben " &
                    "FROM xipunum_erp_store AS a INNER JOIN xipunum_erp_store_order AS b ON b.id=a.order_id " &
                    "INNER JOIN xipunum_erp_shop AS c ON c.poduct_code=a.poduct_code " &
                    "WHERE b.odd_numbers='" & SafeSQL(orderNumber) & "' ORDER BY a.id ASC"
                Dim detailDt As DataTable = ExecuteQuery(detailSql, MySQL_Read)

                If detailDt Is Nothing Then
                    ShowError("查询入库明细失败，请检查数据库连接！")
                    isSaving = False
                    Return
                End If

                If detailDt.Rows.Count = 0 Then
                    ShowWarning("未找到该入库单明细，无法修改成本！")
                    isSaving = False
                    Return
                End If

                Dim totalCost As Decimal = 0

                For i As Integer = 0 To detailDt.Rows.Count - 1
                    Dim row As DataRow = detailDt.Rows(i)
                    Dim storeId As String = SafeString(row("rukuid"))
                    Dim productCode As String = SafeString(row("bianam"))
                    Dim jinZhong As Decimal = SafeDecimal(row("jinzhong"))
                    Dim quantity As Decimal = SafeDecimal(row("rukushu"))
                    Dim origBasicCost As String = SafeString(row("chengbengf"))
                    Dim origCompanySurcharge As String = SafeString(row("chengebnfj"))
                    Dim origCostPrice As String = SafeString(row("chengben"))
                    Dim origGoldPrice As String = SafeString(row("gold_price"))

                    ' 确定当前使用的值
                    Dim currentBasicCost As String = inputBasicCost
                    Dim currentCompanySurcharge As String = inputCompanySurcharge
                    Dim currentGoldPrice As String = inputGoldPrice

                    If Not laborChanged OrElse currentBasicCost = "" OrElse SafeDecimal(currentBasicCost) = 0 Then
                        currentBasicCost = origBasicCost
                    End If
                    If Not surchargeChanged OrElse currentCompanySurcharge = "" OrElse SafeDecimal(currentCompanySurcharge) = 0 Then
                        currentCompanySurcharge = origCompanySurcharge
                    End If
                    If Not goldPriceChanged OrElse currentGoldPrice = "" OrElse SafeDecimal(currentGoldPrice) = 0 Then
                        currentGoldPrice = origGoldPrice
                    End If

                    ' 计算成本单价
                    Dim costUnitPrice As Decimal
                    If quantity = 0 Then
                        costUnitPrice = jinZhong * (SafeDecimal(currentBasicCost) + SafeDecimal(currentGoldPrice)) + SafeDecimal(currentCompanySurcharge)
                        totalCost += costUnitPrice
                    Else
                        costUnitPrice = (jinZhong / quantity) * (SafeDecimal(currentBasicCost) + SafeDecimal(currentGoldPrice)) + SafeDecimal(currentCompanySurcharge)
                        totalCost += costUnitPrice * quantity
                    End If

                    Dim costUnitPriceStr As String = FormatTwoDecimals(costUnitPrice.ToString())
                    Dim totalCostStr As String = FormatTwoDecimals(totalCost.ToString())

                    ' 更新商品成本
                    ExecuteCommand("UPDATE xipunum_erp_store SET cost_price='" & SafeSQL(costUnitPriceStr) & "',basic_cost='" & SafeSQL(currentBasicCost) & "',company_surcharge='" & SafeSQL(currentCompanySurcharge) & "',updatetime='" & operationDate & "' WHERE id='" & SafeSQL(storeId) & "' LIMIT 1", MySQL_Write)

                    ' 记录日志
                    LogContent = "账户:" & UserAccount & " 修改商品编码：" & productCode & " 成本工费：" & origBasicCost & "->" & currentBasicCost & " 成本附加费：" & origCompanySurcharge & "->" & currentCompanySurcharge & " 成本单价：" & origCostPrice & "->" & costUnitPriceStr
                    ExecuteCommand("INSERT INTO xipunum_erp_xitong_log(type,title,conter,user,creationtime) VALUES('修改','修改商品参数','" & SafeSQL(LogContent) & "','" & SafeSQL(operationAccount) & "','" & operationDate & "')", MySQL_Write)
                Next

                Dim totalCostFinal As String = FormatTwoDecimals(totalCost.ToString())
                ' 更新订单总成本
                ExecuteCommand("UPDATE xipunum_erp_store_order SET chengben='" & SafeSQL(totalCostFinal) & "',updatetime='" & operationDate & "' WHERE odd_numbers='" & SafeSQL(orderNumber) & "' LIMIT 1", MySQL_Write)
            End If

            ShowSuccess("入库订单信息修改完成！")
            Me.Close()
        Catch ex As Exception
            ShowError("修改失败：" & ex.Message)
            isSaving = False
        End Try
    End Sub

    ' ========== _按钮EX2_鼠标左键单击（重置） ==========
    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        Form_Load(Nothing, Nothing)
    End Sub

    ' ========== _图片框EX4_鼠标左键单击（关闭） ==========
    Private Sub BtnCancel_Click(sender As Object, e As EventArgs)
        If isSaving Then Return
        Me.Close()
    End Sub

    ' ========== 工厂修改_格式二位小数 ==========
    Private Function FormatTwoDecimals(value As String) As String
        Dim numValue As Decimal = SafeDecimal(value)
        Dim rounded As Decimal = Math.Round(numValue, 2)
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

End Class

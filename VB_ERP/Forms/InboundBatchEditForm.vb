' ============================================================================
' 商品入库批量修改窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品入库批量修改.form.e.txt
' 包含2个程序集变量、7个子程序
' 功能：从InboundDetailForm表格中选择列名，批量修改所有行的该列值
' ============================================================================

Imports System.Data

Public Class InboundBatchEditForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（2个） ==========
    Private selectedColumnName As String = ""          ' 组合框列名称数据
    Private parentDetailForm As InboundDetailForm = Nothing  ' 父窗口引用

    ' ========== 控件声明 ==========
    Private grpEdit As New GroupBox()                  ' 添加修改_分组框
    Private cmbColumnName As New ComboBox()            ' 组合框列名称
    Private txtColumnValue As New TextBox()            ' 编辑框_列参数
    Private cmbColumnValue As New ComboBox()           ' 组合框列参数（单位选择）
    Private btnApply As New Button()                   ' 按钮EX1（确定修改）
    Private btnReset As New Button()                   ' 按钮EX2（重置）
    Private btnCancel As New Button()                  ' 图片框EX4（关闭）

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
    End Sub

    ' ========== 设置父窗口引用 ==========
    Public Sub SetParentForm(parent As InboundDetailForm)
        parentDetailForm = parent
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "商品入库批量修改"
        Me.Size = New Drawing.Size(500, 350)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 分组框
        grpEdit.Text = "批量修改"
        grpEdit.Size = New Drawing.Size(380, 220)
        grpEdit.Location = New Drawing.Point(50, 20)
        Me.Controls.Add(grpEdit)

        ' 列名称标签
        Dim lblColumn As New Label()
        lblColumn.Text = "列名称："
        lblColumn.Location = New Drawing.Point(20, 30)
        lblColumn.AutoSize = True
        grpEdit.Controls.Add(lblColumn)

        ' 列名称组合框
        cmbColumnName.Location = New Drawing.Point(90, 27)
        cmbColumnName.Size = New Drawing.Size(180, 25)
        cmbColumnName.DropDownStyle = ComboBoxStyle.DropDownList
        grpEdit.Controls.Add(cmbColumnName)

        ' 列参数标签
        Dim lblValue As New Label()
        lblValue.Text = "列参数："
        lblValue.Location = New Drawing.Point(20, 70)
        lblValue.AutoSize = True
        grpEdit.Controls.Add(lblValue)

        ' 编辑框_列参数
        txtColumnValue.Location = New Drawing.Point(90, 67)
        txtColumnValue.Size = New Drawing.Size(180, 25)
        txtColumnValue.Visible = True
        txtColumnValue.Enabled = False
        grpEdit.Controls.Add(txtColumnValue)

        ' 组合框列参数（单位选择）
        cmbColumnValue.Location = New Drawing.Point(90, 67)
        cmbColumnValue.Size = New Drawing.Size(180, 25)
        cmbColumnValue.DropDownStyle = ComboBoxStyle.DropDownList
        cmbColumnValue.Visible = False
        grpEdit.Controls.Add(cmbColumnValue)

        ' 确定修改按钮
        btnApply.Text = "确定修改"
        btnApply.Location = New Drawing.Point(60, 120)
        btnApply.Size = New Drawing.Size(90, 35)
        grpEdit.Controls.Add(btnApply)

        ' 重置按钮
        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(165, 120)
        btnReset.Size = New Drawing.Size(90, 35)
        grpEdit.Controls.Add(btnReset)

        ' 关闭按钮
        btnCancel.Text = "关闭"
        btnCancel.Location = New Drawing.Point(270, 120)
        btnCancel.Size = New Drawing.Size(90, 35)
        grpEdit.Controls.Add(btnCancel)

        ' 事件绑定
        AddHandler cmbColumnName.SelectedIndexChanged, AddressOf CmbColumnName_Changed
        AddHandler btnApply.Click, AddressOf BtnApply_Click
        AddHandler btnReset.Click, AddressOf BtnReset_Click
        AddHandler btnCancel.Click, AddressOf BtnCancel_Click
    End Sub

    ' ========== _窗口_商品入库批量修改_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        selectedColumnName = ""
        cmbColumnName.Items.Clear()
        cmbColumnName.Text = "请选择列名称"

        If parentDetailForm Is Nothing Then
            ShowWarning("入库详情窗口不存在，无法批量修改！")
            Return
        End If

        Dim dgv As DataGridView = parentDetailForm.GetDetailGrid()
        If dgv.Columns.Count <= 6 Then
            ShowWarning("入库详情表格列数异常！")
            Return
        End If

        ' 从第6列开始添加列名（跳过序号、编码、名称、品类、规格、款号）
        For i As Integer = 6 To dgv.Columns.Count - 1
            ' 跳过操作列（最后一列按钮列）
            If TypeOf dgv.Columns(i) Is DataGridViewButtonColumn Then Continue For
            cmbColumnName.Items.Add(dgv.Columns(i).HeaderText)
        Next

        txtColumnValue.Visible = True
        txtColumnValue.Enabled = False
        txtColumnValue.Text = ""
        cmbColumnValue.Items.Clear()
        cmbColumnValue.Text = "请选择列参数"
        cmbColumnValue.Visible = False
    End Sub

    ' ========== _窗口_商品入库批量修改_尺寸被改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        grpEdit.Left = (Me.Width - grpEdit.Width) \ 2
        grpEdit.Top = (Me.Height - grpEdit.Height) \ 2
        If grpEdit.Left < 0 Then grpEdit.Left = 0
        If grpEdit.Top < 0 Then grpEdit.Top = 0
    End Sub

    ' ========== _组合框列名称_内容被改变 ==========
    Private Sub CmbColumnName_Changed(sender As Object, e As EventArgs)
        If cmbColumnName.SelectedIndex = -1 Then
            selectedColumnName = ""
            txtColumnValue.Enabled = False
            txtColumnValue.Text = ""
            cmbColumnValue.Visible = False
            txtColumnValue.Visible = True
            Return
        End If

        selectedColumnName = cmbColumnName.SelectedItem.ToString()

        ' 不可修改的列
        If selectedColumnName = "库房" OrElse selectedColumnName = "公司成色" OrElse
           selectedColumnName = "单件重" OrElse selectedColumnName = "库存" OrElse
           selectedColumnName = "含耗重" OrElse selectedColumnName = "成本价" OrElse
           selectedColumnName = "销售价" Then
            ShowWarning("所选数据不可修改！")
            cmbColumnName.SelectedIndex = -1
            cmbColumnName.Focus()
            Return
        End If

        ' 非镶嵌模式下不可修改镶嵌相关列
        If parentDetailForm IsNot Nothing AndAlso parentDetailForm.IsNotInlaid() Then
            If selectedColumnName = "石重" OrElse selectedColumnName = "石头数" OrElse
               selectedColumnName = "副石重" OrElse selectedColumnName = "副石头数" OrElse
               selectedColumnName = "主石色" Then
                ShowWarning("所选数据不可修改！")
                cmbColumnName.SelectedIndex = -1
                cmbColumnName.Focus()
                Return
            End If
        End If

        txtColumnValue.Enabled = False
        txtColumnValue.Text = ""

        If selectedColumnName = "单位" Then
            txtColumnValue.Visible = False
            cmbColumnValue.Visible = True
            cmbColumnValue.Items.Clear()
            cmbColumnValue.Items.Add("重量")
            cmbColumnValue.Items.Add("单件")
            cmbColumnValue.SelectedIndex = -1
        Else
            txtColumnValue.Visible = True
            cmbColumnValue.Visible = False
            txtColumnValue.Enabled = True
        End If
    End Sub

    ' ========== _按钮EX1_鼠标左键单击（确定修改） ==========
    Private Sub BtnApply_Click(sender As Object, e As EventArgs)
        If cmbColumnName.SelectedIndex = -1 Then
            ShowWarning("请选择列名称！")
            cmbColumnName.Focus()
            Return
        End If

        selectedColumnName = cmbColumnName.SelectedItem.ToString()
        Dim selectedColIndex As Integer = cmbColumnName.SelectedIndex + 6

        Dim baseValue As String = ""
        If selectedColumnName = "单位" Then
            If cmbColumnValue.SelectedIndex = -1 Then
                ShowWarning("请选择单位！")
                cmbColumnValue.Focus()
                Return
            End If
            baseValue = cmbColumnValue.SelectedItem.ToString()
        Else
            If txtColumnValue.Text = "" Then
                ShowWarning("列参数不能为空！")
                txtColumnValue.Focus()
                Return
            End If
            baseValue = txtColumnValue.Text
        End If

        ' 验证：数量(13)和重量(18)必须>0
        If selectedColIndex = 13 OrElse selectedColIndex = 18 Then
            If SafeDecimal(baseValue) <= 0 Then
                ShowWarning("该列参数必须大于0！")
                txtColumnValue.Focus()
                Return
            End If
        End If

        ' 验证：金重/损耗/含耗重/石重/副石重/成本工费/参考工费/销售工费/成本附加费/销售附加费不能<0
        If selectedColIndex = 15 OrElse selectedColIndex = 16 OrElse selectedColIndex = 17 OrElse
           selectedColIndex = 20 OrElse selectedColIndex = 22 OrElse selectedColIndex = 26 OrElse
           selectedColIndex = 27 OrElse selectedColIndex = 28 OrElse selectedColIndex = 29 OrElse
           selectedColIndex = 30 Then
            If SafeDecimal(baseValue) < 0 Then
                ShowWarning("该列参数不能小于0！")
                txtColumnValue.Focus()
                Return
            End If
        End If

        Dim dgv As DataGridView = parentDetailForm.GetDetailGrid()
        Dim modifiableRows As Integer = dgv.Rows.Count - 2
        If modifiableRows <= 0 Then
            ShowWarning("没有可修改的数据！")
            Return
        End If

        ' 遍历所有数据行进行批量修改
        For i As Integer = 0 To modifiableRows - 1
            If SafeString(dgv.Rows(i).Cells(1).Value) <> "" Then
                Dim currentValue As String = baseValue

                ' 直接赋值的列：6,7,8,9,19,21,23,25,33,34
                If selectedColIndex = 6 OrElse selectedColIndex = 7 OrElse selectedColIndex = 8 OrElse
                   selectedColIndex = 9 OrElse selectedColIndex = 19 OrElse selectedColIndex = 21 OrElse
                   selectedColIndex = 23 OrElse selectedColIndex = 25 OrElse selectedColIndex = 33 OrElse
                   selectedColIndex = 34 Then
                    dgv.Rows(i).Cells(selectedColIndex).Value = currentValue
                End If

                ' 工厂成色(10)：同时设置公司成色(11)
                If selectedColIndex = 10 Then
                    dgv.Rows(i).Cells(10).Value = currentValue
                    dgv.Rows(i).Cells(11).Value = currentValue
                End If

                ' 数量(13)：重新计算单件重 = 重量/数量
                If selectedColIndex = 13 Then
                    Dim weightVal As Decimal = SafeDecimal(dgv.Rows(i).Cells(18).Value)
                    Dim qtyVal As Decimal = SafeDecimal(currentValue)
                    Dim singleWeight As String = FormatThreeDecimals((weightVal / qtyVal).ToString())
                    dgv.Rows(i).Cells(12).Value = singleWeight
                    dgv.Rows(i).Cells(13).Value = currentValue
                End If

                ' 重量(18)：重新计算单件重 = 重量/数量
                If selectedColIndex = 18 Then
                    Dim qtyVal As Decimal = SafeDecimal(dgv.Rows(i).Cells(13).Value)
                    If qtyVal <= 0 Then
                        dgv.Rows(i).Cells(13).Value = "1"
                        qtyVal = 1
                    End If
                    Dim singleWeight As String = FormatThreeDecimals((SafeDecimal(currentValue) / qtyVal).ToString())
                    dgv.Rows(i).Cells(12).Value = singleWeight
                    dgv.Rows(i).Cells(18).Value = FormatThreeDecimals(currentValue)
                End If

                ' 损耗(16)/含耗重(17)/石重(20)/副石重(22)：格式三位小数
                If selectedColIndex = 16 OrElse selectedColIndex = 17 OrElse
                   selectedColIndex = 20 OrElse selectedColIndex = 22 Then
                    dgv.Rows(i).Cells(selectedColIndex).Value = FormatThreeDecimals(currentValue)
                End If

                ' 金重(15)：重算成本价 = 金重*成本工费 + 成本附加费
                If selectedColIndex = 15 Then
                    Dim jinZhong As Decimal = SafeDecimal(currentValue)
                    Dim basicCost As Decimal = SafeDecimal(dgv.Rows(i).Cells(26).Value)
                    Dim companySur As Decimal = SafeDecimal(dgv.Rows(i).Cells(29).Value)
                    Dim costPrice As String = FormatTwoDecimals((jinZhong * basicCost + companySur).ToString())
                    dgv.Rows(i).Cells(15).Value = FormatThreeDecimals(currentValue)
                    dgv.Rows(i).Cells(24).Value = costPrice
                End If

                ' 成本工费(26)：重算成本价 = 金重*成本工费 + 成本附加费
                If selectedColIndex = 26 Then
                    Dim basicCost As Decimal = SafeDecimal(currentValue)
                    Dim jinZhong As Decimal = SafeDecimal(dgv.Rows(i).Cells(15).Value)
                    Dim companySur As Decimal = SafeDecimal(dgv.Rows(i).Cells(29).Value)
                    Dim costPrice As String = FormatTwoDecimals((basicCost * jinZhong + companySur).ToString())
                    dgv.Rows(i).Cells(24).Value = costPrice
                    dgv.Rows(i).Cells(26).Value = FormatTwoDecimals(currentValue)
                End If

                ' 参考工费(27)：同步到销售工费(28)，重算销售价 = 参考工费*金重 + 销售附加费
                If selectedColIndex = 27 Then
                    Dim premiumCost As Decimal = SafeDecimal(currentValue)
                    Dim jinZhong As Decimal = SafeDecimal(dgv.Rows(i).Cells(15).Value)
                    Dim salesSur As Decimal = SafeDecimal(dgv.Rows(i).Cells(30).Value)
                    Dim salesPrice As String = FormatTwoDecimals((premiumCost * jinZhong + salesSur).ToString())
                    dgv.Rows(i).Cells(27).Value = FormatTwoDecimals(currentValue)
                    dgv.Rows(i).Cells(28).Value = FormatTwoDecimals(currentValue)
                    dgv.Rows(i).Cells(31).Value = salesPrice
                End If

                ' 销售工费(28)：重算销售价 = 销售工费*金重 + 销售附加费
                If selectedColIndex = 28 Then
                    Dim salesCost As Decimal = SafeDecimal(currentValue)
                    Dim jinZhong As Decimal = SafeDecimal(dgv.Rows(i).Cells(15).Value)
                    Dim salesSur As Decimal = SafeDecimal(dgv.Rows(i).Cells(30).Value)
                    Dim salesPrice As String = FormatTwoDecimals((salesCost * jinZhong + salesSur).ToString())
                    dgv.Rows(i).Cells(28).Value = FormatTwoDecimals(currentValue)
                    dgv.Rows(i).Cells(31).Value = salesPrice
                End If

                ' 成本附加费(29)：重算成本价 = 金重*成本工费 + 成本附加费
                If selectedColIndex = 29 Then
                    Dim jinZhong As Decimal = SafeDecimal(dgv.Rows(i).Cells(15).Value)
                    Dim basicCost As Decimal = SafeDecimal(dgv.Rows(i).Cells(26).Value)
                    Dim companySur As Decimal = SafeDecimal(currentValue)
                    Dim costPrice As String = FormatTwoDecimals((jinZhong * basicCost + companySur).ToString())
                    dgv.Rows(i).Cells(24).Value = costPrice
                    dgv.Rows(i).Cells(29).Value = FormatTwoDecimals(currentValue)
                End If

                ' 销售附加费(30)：重算销售价 = 销售工费*金重 + 销售附加费
                If selectedColIndex = 30 Then
                    Dim salesCost As Decimal = SafeDecimal(dgv.Rows(i).Cells(28).Value)
                    Dim jinZhong As Decimal = SafeDecimal(dgv.Rows(i).Cells(15).Value)
                    Dim salesSur As Decimal = SafeDecimal(currentValue)
                    Dim salesPrice As String = FormatTwoDecimals((salesCost * jinZhong + salesSur).ToString())
                    dgv.Rows(i).Cells(30).Value = FormatTwoDecimals(currentValue)
                    dgv.Rows(i).Cells(31).Value = salesPrice
                End If
            End If
        Next

        ' 触发父窗口重新统计
        parentDetailForm.RecalculateStatistics()
        Me.Close()
    End Sub

    ' ========== _按钮EX2_鼠标左键单击（重置） ==========
    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        Form_Load(Nothing, Nothing)
    End Sub

    ' ========== _图片框EX4_鼠标左键单击（关闭） ==========
    Private Sub BtnCancel_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ' ========== 格式三位小数 ==========
    Private Function FormatThreeDecimals(value As String) As String
        Dim numValue As Decimal = SafeDecimal(value)
        Dim rounded As Decimal = Math.Round(numValue, 3)
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

' ============================================================================
' 商品销售批量修改窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品销售批量修改.form.e.txt
' 控件：编辑框_实时克价, 编辑框_销售工费, 编辑框_销售附加费, 编辑框_附加费折扣,
'       组合框_导购员, 按钮EX_修改, 按钮EX_重置, 图片框EX4(关闭), 添加修改_分组框
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SalesBatchEditForm
    Inherits System.Windows.Forms.Form

    ' ========== 控件声明 ==========
    Private txtGoldPrice As New TextBox()          ' 编辑框_实时克价
    Private txtSalesCost As New TextBox()           ' 编辑框_销售工费
    Private txtSalesSurcharge As New TextBox()      ' 编辑框_销售附加费
    Private txtSurchargeDiscount As New TextBox()   ' 编辑框_附加费折扣
    Private cmbGuide As New ComboBox()              ' 组合框_导购员
    Private WithEvents btnModify As New Button()    ' 按钮EX_修改
    Private WithEvents btnReset As New Button()     ' 按钮EX_重置
    Private WithEvents btnClose As New Button()     ' 图片框EX4（关闭）
    Private grpMain As New GroupBox()               ' 添加修改_分组框

    ' 引用销售出库窗口的表格（由调用方设置）
    Public Property SalesGrid As DataGridView

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "商品销售批量修改"
        Me.Size = New Drawing.Size(541, 247)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 主分组框
        grpMain.Text = "批量修改"
        grpMain.Location = New Drawing.Point(10, 10)
        grpMain.Size = New Drawing.Size(515, 195)
        Me.Controls.Add(grpMain)

        ' 实时克价
        Dim lblGoldPrice As New Label()
        lblGoldPrice.Text = "实时克价"
        lblGoldPrice.Location = New Drawing.Point(16, 30)
        lblGoldPrice.AutoSize = True
        grpMain.Controls.Add(lblGoldPrice)
        txtGoldPrice.Location = New Drawing.Point(88, 27)
        txtGoldPrice.Size = New Drawing.Size(150, 25)
        grpMain.Controls.Add(txtGoldPrice)

        ' 销售工费
        Dim lblSalesCost As New Label()
        lblSalesCost.Text = "销售工费"
        lblSalesCost.Location = New Drawing.Point(16, 65)
        lblSalesCost.AutoSize = True
        grpMain.Controls.Add(lblSalesCost)
        txtSalesCost.Location = New Drawing.Point(88, 62)
        txtSalesCost.Size = New Drawing.Size(150, 25)
        grpMain.Controls.Add(txtSalesCost)

        ' 销售附加费
        Dim lblSurcharge As New Label()
        lblSurcharge.Text = "销售附加费"
        lblSurcharge.Location = New Drawing.Point(16, 100)
        lblSurcharge.AutoSize = True
        grpMain.Controls.Add(lblSurcharge)
        txtSalesSurcharge.Location = New Drawing.Point(88, 97)
        txtSalesSurcharge.Size = New Drawing.Size(150, 25)
        grpMain.Controls.Add(txtSalesSurcharge)

        ' 附加费折扣
        Dim lblDiscount As New Label()
        lblDiscount.Text = "附加费折扣"
        lblDiscount.Location = New Drawing.Point(16, 135)
        lblDiscount.AutoSize = True
        grpMain.Controls.Add(lblDiscount)
        txtSurchargeDiscount.Location = New Drawing.Point(88, 132)
        txtSurchargeDiscount.Size = New Drawing.Size(150, 25)
        grpMain.Controls.Add(txtSurchargeDiscount)

        ' 导购员
        Dim lblGuide As New Label()
        lblGuide.Text = "导购员"
        lblGuide.Location = New Drawing.Point(270, 135)
        lblGuide.AutoSize = True
        grpMain.Controls.Add(lblGuide)
        cmbGuide.Location = New Drawing.Point(344, 132)
        cmbGuide.Size = New Drawing.Size(150, 25)
        cmbGuide.DropDownStyle = ComboBoxStyle.DropDownList
        grpMain.Controls.Add(cmbGuide)

        ' 按钮
        btnModify.Text = "修改"
        btnModify.Location = New Drawing.Point(270, 27)
        btnModify.Size = New Drawing.Size(80, 30)
        grpMain.Controls.Add(btnModify)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(360, 27)
        btnReset.Size = New Drawing.Size(80, 30)
        grpMain.Controls.Add(btnReset)

        btnClose.Text = "关闭"
        btnClose.Location = New Drawing.Point(440, 27)
        btnClose.Size = New Drawing.Size(60, 30)
        grpMain.Controls.Add(btnClose)

        ' 事件绑定
        AddHandler txtSurchargeDiscount.TextChanged, AddressOf TxtSurchargeDiscount_TextChanged
        AddHandler txtSalesSurcharge.TextChanged, AddressOf TxtSalesSurcharge_TextChanged
    End Sub

    ' ========== _窗口_商品销售批量修改_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        txtGoldPrice.Text = ""
        txtSalesCost.Text = ""
        txtSalesSurcharge.Text = ""
        txtSurchargeDiscount.Text = ""

        ' 加载导购员列表
        cmbGuide.Items.Clear()
        cmbGuide.Items.Add("请选择导购员")
        cmbGuide.SelectedIndex = 0

        Try
            Dim sql As String = ""
            If UserPermission = "全部" Then
                sql = "SELECT name FROM xipunum_erp_user where state='0' order by id ASC"
            ElseIf UserPermission = "店铺" Then
                sql = "SELECT name FROM xipunum_erp_user where department='" & UserDepartment & "' and state='0' order by id ASC"
            ElseIf UserPermission = "岗位" Then
                sql = "SELECT name FROM xipunum_erp_user WHERE user in " & GlobalViewSQL & " and state='0' order by id ASC"
            ElseIf UserPermission = "个人" Then
                sql = "SELECT name FROM xipunum_erp_user WHERE user='" & UserAccount & "' and state='0' order by id ASC"
            Else
                sql = "SELECT name FROM xipunum_erp_user WHERE 1=0"
            End If

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            For Each row As DataRow In dt.Rows
                cmbGuide.Items.Add(SafeString(row("name")))
            Next
        Catch ex As Exception
        End Try
    End Sub

    ' ========== _窗口_商品销售批量修改_尺寸被改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        grpMain.Left = (Me.Width - grpMain.Width) \ 2
        grpMain.Top = (Me.Height - grpMain.Height) \ 2
    End Sub

    ' ========== _图片框EX4_鼠标左键单击（关闭） ==========
    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ' ========== _按钮EX_重置_鼠标左键单击 ==========
    Private Sub BtnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Form_Load(Nothing, Nothing)
    End Sub

    ' ========== _编辑框_附加费折扣_内容被改变 ==========
    Private Sub TxtSurchargeDiscount_TextChanged(sender As Object, e As EventArgs)
        If txtSurchargeDiscount.Text.Length > 0 Then
            If txtSalesSurcharge.Text <> "" Then
                ShowWarning("销售附加费用必须为空！")
                txtSurchargeDiscount.Text = ""
                txtSurchargeDiscount.Focus()
                txtSalesSurcharge.Text = ""
                Return
            End If

            If SafeDecimal(txtSurchargeDiscount.Text) > 1 Then
                ShowWarning("折扣不能大于1！")
                txtSurchargeDiscount.Text = ""
                txtSurchargeDiscount.Focus()
                Return
            End If

            If SafeDecimal(txtSurchargeDiscount.Text) < 0 Then
                ShowWarning("折扣不能小于0！")
                txtSurchargeDiscount.Text = ""
                txtSurchargeDiscount.Focus()
                Return
            End If
        End If
    End Sub

    ' ========== _编辑框_销售附加费_内容被改变 ==========
    Private Sub TxtSalesSurcharge_TextChanged(sender As Object, e As EventArgs)
        If txtSalesSurcharge.Text.Length > 0 Then
            If txtSurchargeDiscount.Text <> "" Then
                ShowWarning("折扣必须为空！")
                txtSalesSurcharge.Text = ""
                txtSalesSurcharge.Focus()
                txtSurchargeDiscount.Text = ""
                Return
            End If
        End If
    End Sub

    ' ========== _按钮EX_修改_鼠标左键单击 ==========
    Private Sub BtnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click
        Dim batchGoldPrice As String = txtGoldPrice.Text
        Dim batchSalesCost As String = txtSalesCost.Text
        Dim batchSurcharge As String = txtSalesSurcharge.Text
        Dim batchDiscount As String = txtSurchargeDiscount.Text

        Dim modifyGoldPrice As Boolean = (batchGoldPrice <> "")
        Dim modifySalesCost As Boolean = (batchSalesCost <> "")
        Dim modifySurcharge As Boolean = (batchSurcharge <> "")
        Dim modifyDiscount As Boolean = (batchDiscount <> "")
        Dim modifyGuide As Boolean = False
        Dim batchGuideName As String = ""

        If cmbGuide.SelectedIndex >= 0 Then
            batchGuideName = cmbGuide.SelectedItem.ToString()
            If batchGuideName <> "" AndAlso batchGuideName <> "请选择导购员" Then
                modifyGuide = True
            End If
        End If

        ' 验证
        If SalesGrid Is Nothing OrElse SalesGrid.Rows.Count <= 1 Then
            ShowWarning("当前没有可批量修改的销售明细！")
            Return
        End If

        If Not modifyGoldPrice AndAlso Not modifySalesCost AndAlso Not modifySurcharge AndAlso Not modifyDiscount AndAlso Not modifyGuide Then
            ShowWarning("请至少填写一项要修改的内容！")
            Return
        End If

        Dim batchCount As Integer = SalesGrid.Rows.Count - 1
        If batchCount <= 0 Then
            ShowWarning("当前没有可批量修改的销售明细！")
            Return
        End If

        ' 遍历修改
        For i As Integer = 0 To batchCount - 1
            Dim productCode As String = SafeString(SalesGrid.Rows(i).Cells(1).Value)
            If productCode = "" Then Continue For

            Dim salesWeight As String = SafeString(SalesGrid.Rows(i).Cells(9).Value)
            Dim salesQty As String = SafeString(SalesGrid.Rows(i).Cells(17).Value)
            Dim origSurchargeUnit As String = SafeString(SalesGrid.Rows(i).Cells(18).Value)
            Dim origSurchargeTotal As String

            If SafeDecimal(salesQty) = 0 Then
                origSurchargeTotal = origSurchargeUnit
            Else
                origSurchargeTotal = (SafeDecimal(origSurchargeUnit) * SafeDecimal(salesQty)).ToString()
            End If

            Dim goldPriceVal As String = SafeString(SalesGrid.Rows(i).Cells(19).Value)
            Dim costVal As String = SafeString(SalesGrid.Rows(i).Cells(20).Value)
            Dim surchargeVal As String = SafeString(SalesGrid.Rows(i).Cells(21).Value)
            Dim discountVal As String = SafeString(SalesGrid.Rows(i).Cells(22).Value)

            If SafeDecimal(discountVal) <= 0 Then
                discountVal = "1"
            End If

            If modifyGoldPrice Then
                goldPriceVal = batchGoldPrice
            End If
            If modifySalesCost Then
                costVal = batchSalesCost
            End If

            If modifySurcharge Then
                surchargeVal = Math.Round(SafeDecimal(batchSurcharge), 0).ToString()
                If SafeDecimal(origSurchargeTotal) = 0 Then
                    discountVal = "1"
                Else
                    discountVal = (SafeDecimal(surchargeVal) / SafeDecimal(origSurchargeTotal)).ToString()
                End If
            End If

            If modifyDiscount Then
                discountVal = batchDiscount
                If SafeDecimal(salesQty) = 0 Then
                    surchargeVal = Math.Round(SafeDecimal(origSurchargeUnit) * SafeDecimal(discountVal), 0).ToString()
                Else
                    surchargeVal = Math.Round(SafeDecimal(origSurchargeUnit) * SafeDecimal(discountVal) * SafeDecimal(salesQty), 0).ToString()
                End If
            End If

            Dim formattedGoldPrice As String = FormatTwoDecimals(goldPriceVal)
            Dim formattedCost As String = FormatTwoDecimals(costVal)
            Dim formattedSurcharge As String = FormatTwoDecimals(surchargeVal)
            Dim formattedDiscount As String
            If SafeDecimal(origSurchargeUnit) = 0 Then
                formattedDiscount = FormatThreeDecimals("1")
            Else
                formattedDiscount = FormatThreeDecimals(discountVal)
            End If

            ' 更新表格显示
            If modifyGoldPrice Then
                SalesGrid.Rows(i).Cells(19).Value = formattedGoldPrice
            End If
            If modifySalesCost Then
                SalesGrid.Rows(i).Cells(20).Value = formattedCost
            End If
            If modifySurcharge Or modifyDiscount Then
                SalesGrid.Rows(i).Cells(21).Value = formattedSurcharge
                SalesGrid.Rows(i).Cells(22).Value = formattedDiscount
            End If

            ' 重新计算销售单价/销售金额/实收金额
            Dim salesUnitPrice As String
            If SafeDecimal(salesQty) = 0 Then
                salesUnitPrice = (SafeDecimal(salesWeight) * (SafeDecimal(goldPriceVal) + SafeDecimal(costVal)) + SafeDecimal(surchargeVal)).ToString()
            Else
                salesUnitPrice = ((SafeDecimal(salesWeight) * (SafeDecimal(goldPriceVal) + SafeDecimal(costVal)) + SafeDecimal(surchargeVal)) / SafeDecimal(salesQty)).ToString()
            End If
            salesUnitPrice = FormatTwoDecimals(salesUnitPrice)

            Dim salesAmount As String = (SafeDecimal(salesWeight) * (SafeDecimal(goldPriceVal) + SafeDecimal(costVal)) + SafeDecimal(surchargeVal)).ToString()
            salesAmount = FormatTwoDecimals(salesAmount)

            Dim receivedAmount As String = FormatTwoDecimals(Math.Round(SafeDecimal(salesAmount), 0).ToString())

            SalesGrid.Rows(i).Cells(15).Value = salesUnitPrice
            SalesGrid.Rows(i).Cells(16).Value = salesAmount
            SalesGrid.Rows(i).Cells(23).Value = receivedAmount

            ' 更新数据库（本地Access库 - 在VB.NET中直接更新MySQL）
            Dim safeCode As String = SafeSQL(productCode)
            If modifyGoldPrice Or modifySalesCost Or modifySurcharge Or modifyDiscount Then
                Try
                    Dim sql As String = "UPDATE xipunum_erp_outbound SET " &
                        "gold_price='" & formattedGoldPrice & "'," &
                        "sales_cost='" & formattedCost & "'," &
                        "sales_surcharge='" & formattedSurcharge & "'," &
                        "settlement='" & receivedAmount & "' " &
                        "WHERE poduct_code='" & safeCode & "' AND order_id IN (SELECT id FROM xipunum_erp_outbound_order WHERE state='正常')"
                    DatabaseModule.ExecuteCommand(sql, MySQL_Write)
                Catch ex As Exception
                End Try
            End If

            If modifyGuide Then
                SalesGrid.Rows(i).Cells(24).Value = batchGuideName
                Try
                    Dim sql As String = "UPDATE xipunum_erp_outbound SET shopping_guide='" & SafeSQL(batchGuideName) & "' WHERE poduct_code='" & safeCode & "'"
                    DatabaseModule.ExecuteCommand(sql, MySQL_Write)
                Catch ex As Exception
                End Try
            End If
        Next

        Me.Close()

        ' 触发父窗口重新统计
        If SalesGrid IsNot Nothing AndAlso SalesGrid.Parent IsNot Nothing Then
            Dim parentForm As Form = TryCast(SalesGrid.Parent, Form)
            If parentForm IsNot Nothing Then
                Dim calcMethod As Reflection.MethodInfo = parentForm.GetType().GetMethod("CalculateSummary")
                If calcMethod IsNot Nothing Then
                    calcMethod.Invoke(parentForm, Nothing)
                End If
            End If
        End If
    End Sub

    ' ========== _批量修改_SQL文本处理 ==========
    Private Function SafeSQLForBatch(text As String) As String
        Return text.Replace("'", "''")
    End Function

    ' ========== _批量修改_格式二位小数 ==========
    Private Function FormatTwoDecimals(originalText As String) As String
        Dim processedText As String = Math.Round(SafeDecimal(originalText), 2).ToString()
        If processedText = "" Then Return "0.00"

        Dim parts() As String = processedText.Split("."c)
        If parts.Length > 1 Then
            Dim decimalLen As Integer = parts(1).Length
            If decimalLen = 1 Then
                Return processedText & "0"
            Else
                Return processedText
            End If
        Else
            Return processedText & ".00"
        End If
    End Function

    ' ========== _批量修改_格式三位小数 ==========
    Private Function FormatThreeDecimals(originalText As String) As String
        Dim processedText As String = Math.Round(SafeDecimal(originalText), 3).ToString()
        If processedText = "" Then Return "0.000"

        Dim parts() As String = processedText.Split("."c)
        If parts.Length > 1 Then
            Dim decimalLen As Integer = parts(1).Length
            If decimalLen = 1 Then
                Return processedText & "00"
            ElseIf decimalLen = 2 Then
                Return processedText & "0"
            Else
                Return processedText
            End If
        Else
            Return processedText & ".000"
        End If
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

    Private Sub ShowError(msg As String)
        MessageBox.Show(msg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

End Class

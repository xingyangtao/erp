' ============================================================================
' 商品销售编辑窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品销售编辑.form.e.txt
' ============================================================================
Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SalesEditForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（38个） ==========
    Private row1 As Integer = 0, col1 As Integer = 0
    Private row2 As Integer = 0, col2 As Integer = 0
    Private row3 As Integer = 0, col3 As Integer = 0
    Private salesOrderID As String = ""
    Private salesRecoveryID As String = ""
    Private salesRecoveryNumber As String = ""
    Private salesMemberCode As String = ""
    Private weightTotal As String = "0"
    Private weightPrice As String = "0"
    Private salesAmount As String = "0"
    Private basicCostTotal As String = "0"
    Private premiumCostTotal As String = "0"
    Private salesCostTotal As String = "0"
    Private costSurchargeTotal As String = "0"
    Private salesSurchargeTotal As String = "0"
    Private recoveryWeightTotal As String = "0"
    Private recoveryGoldTotal As String = "0"
    Private recoveryOtherTotal As String = "0"
    Private recoveryAmountTotal As String = "0"
    Private quantityTotal As String = "0"
    Private actualAmount As String = "0"
    Private salesOrderStatus As String = ""
    Private salesProductCode As String = ""
    Private salesDataCode As String = ""
    Private salesPling As String = "零售"
    Private tableEditState As Integer = 0
    Private paymentEditState As Integer = 0
    Private recoveryEditState As Integer = 0
    Private saveInProgress As Integer = 0

    ' ========== 控件声明 ==========
    Private dgvSales As New DataGridView()
    Private dgvRecovery As New DataGridView()
    Private dgvPayment As New DataGridView()
    Private dgvSummary As New DataGridView()
    Private dgvDeleted As New DataGridView()
    Private txtOrderNumber As New TextBox()
    Private txtMemberName As New TextBox()
    Private txtPhone As New TextBox()
    Private txtReceivable As New TextBox()
    Private txtReceived As New TextBox()
    Private txtTax As New TextBox()
    Private txtTaxRate As New TextBox()
    Private cmbGuide As New ComboBox()
    Private cmbVoucherStyle As New ComboBox()
    Private cmbIntroducer As New ComboBox()
    Private cmbSalesFactory As New ComboBox()
    Private radioRetail As New RadioButton()
    Private radioWholesale As New RadioButton()
    Private radioInvoice As New RadioButton()
    Private radioNoInvoice As New RadioButton()
    Private radioPresaleYes As New RadioButton()
    Private radioPresaleNo As New RadioButton()
    Private radioNewRow As New RadioButton()
    Private radioNewCol As New RadioButton()
    Private txtRemarks As New TextBox()
    Private txtPresaleNumber As New TextBox()
    Private txtPresaleDeposit As New TextBox()
    Private txtStoreMaterial As New TextBox()
    Private txtStoreBalance As New TextBox()
    Private lblStoreMaterial As New Label()
    Private lblStoreBalance As New Label()
    Private lblStoreMaterialVal As New Label()
    Private lblStoreBalanceVal As New Label()
    Private chkPrintDoc As New CheckBox()
    Private chkPrintPreview As New CheckBox()
    Private chkPrintChengse As New CheckBox()
    Private groupInvoice As New GroupBox()
    Private groupPresale As New GroupBox()
    Private groupBatch As New GroupBox()
    Private btnRecoveryAdd As New Button()
    Private btnRecoveryRemove As New Button()
    Private btnToolbarSave As New Button()
    Private btnToolbarReset As New Button()
    Private btnToolbarPrintPreview As New Button()
    Private btnToolbarBatchEdit As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "商品销售编辑"
        Me.Size = New Drawing.Size(1430, 698)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 110
        Me.Controls.Add(panelTop)

        AddLabel(panelTop, "单据号：", 10, 7)
        txtOrderNumber.Location = New Drawing.Point(60, 5) : txtOrderNumber.Size = New Drawing.Size(130, 25) : txtOrderNumber.ReadOnly = True
        panelTop.Controls.Add(txtOrderNumber)

        AddLabel(panelTop, "会员姓名：", 200, 7)
        txtMemberName.Location = New Drawing.Point(260, 5) : txtMemberName.Size = New Drawing.Size(120, 25) : txtMemberName.ReadOnly = True
        panelTop.Controls.Add(txtMemberName)

        AddLabel(panelTop, "联系电话：", 390, 7)
        txtPhone.Location = New Drawing.Point(450, 5) : txtPhone.Size = New Drawing.Size(100, 25) : txtPhone.ReadOnly = True
        panelTop.Controls.Add(txtPhone)

        AddLabel(panelTop, "业务员：", 560, 7)
        cmbGuide.Location = New Drawing.Point(610, 5) : cmbGuide.Size = New Drawing.Size(100, 25)
        panelTop.Controls.Add(cmbGuide)

        AddLabel(panelTop, "预售单号：", 720, 7)
        txtPresaleNumber.Location = New Drawing.Point(780, 5) : txtPresaleNumber.Size = New Drawing.Size(130, 25)
        panelTop.Controls.Add(txtPresaleNumber)

        AddLabel(panelTop, "预售订金：", 920, 7)
        txtPresaleDeposit.Location = New Drawing.Point(980, 5) : txtPresaleDeposit.Size = New Drawing.Size(80, 25) : txtPresaleDeposit.ReadOnly = True
        panelTop.Controls.Add(txtPresaleDeposit)

        groupPresale.Location = New Drawing.Point(1070, 2) : groupPresale.Size = New Drawing.Size(80, 28)
        radioPresaleYes.Text = "是" : radioPresaleYes.Location = New Drawing.Point(5, 10) : radioPresaleYes.AutoSize = True
        groupPresale.Controls.Add(radioPresaleYes)
        radioPresaleNo.Text = "否" : radioPresaleNo.Location = New Drawing.Point(40, 10) : radioPresaleNo.AutoSize = True : radioPresaleNo.Checked = True
        groupPresale.Controls.Add(radioPresaleNo)
        panelTop.Controls.Add(groupPresale)

        AddLabel(panelTop, "批零：", 10, 40)
        radioRetail.Text = "零售" : radioRetail.Location = New Drawing.Point(50, 40) : radioRetail.AutoSize = True : radioRetail.Checked = True
        panelTop.Controls.Add(radioRetail)
        radioWholesale.Text = "批发" : radioWholesale.Location = New Drawing.Point(110, 40) : radioWholesale.AutoSize = True
        panelTop.Controls.Add(radioWholesale)

        groupInvoice.Location = New Drawing.Point(200, 35) : groupInvoice.Size = New Drawing.Size(120, 30)
        radioNoInvoice.Text = "不开票" : radioNoInvoice.Location = New Drawing.Point(5, 8) : radioNoInvoice.AutoSize = True : radioNoInvoice.Checked = True
        groupInvoice.Controls.Add(radioNoInvoice)
        radioInvoice.Text = "开票" : radioInvoice.Location = New Drawing.Point(67, 8) : radioInvoice.AutoSize = True
        groupInvoice.Controls.Add(radioInvoice)
        panelTop.Controls.Add(groupInvoice)

        AddLabel(panelTop, "介绍人：", 330, 40)
        cmbIntroducer.Location = New Drawing.Point(380, 38) : cmbIntroducer.Size = New Drawing.Size(118, 25)
        panelTop.Controls.Add(cmbIntroducer)

        AddLabel(panelTop, "销售工厂：", 510, 40)
        cmbSalesFactory.Location = New Drawing.Point(570, 38) : cmbSalesFactory.Size = New Drawing.Size(118, 25)
        panelTop.Controls.Add(cmbSalesFactory)

        groupBatch.Location = New Drawing.Point(700, 35) : groupBatch.Size = New Drawing.Size(110, 30)
        radioNewRow.Text = "换行" : radioNewRow.Location = New Drawing.Point(5, 8) : radioNewRow.AutoSize = True : radioNewRow.Checked = True
        groupBatch.Controls.Add(radioNewRow)
        radioNewCol.Text = "换列" : radioNewCol.Location = New Drawing.Point(55, 8) : radioNewCol.AutoSize = True
        groupBatch.Controls.Add(radioNewCol)
        panelTop.Controls.Add(groupBatch)

        AddLabel(panelTop, "单据样式：", 820, 40)
        cmbVoucherStyle.Location = New Drawing.Point(880, 38) : cmbVoucherStyle.Size = New Drawing.Size(100, 25) : cmbVoucherStyle.DropDownStyle = ComboBoxStyle.DropDownList
        panelTop.Controls.Add(cmbVoucherStyle)

        chkPrintDoc.Text = "打印单据" : chkPrintDoc.Location = New Drawing.Point(990, 40) : chkPrintDoc.AutoSize = True
        panelTop.Controls.Add(chkPrintDoc)
        chkPrintPreview.Text = "打印预览" : chkPrintPreview.Location = New Drawing.Point(1070, 40) : chkPrintPreview.AutoSize = True
        panelTop.Controls.Add(chkPrintPreview)
        chkPrintChengse.Text = "打印成色" : chkPrintChengse.Location = New Drawing.Point(1160, 40) : chkPrintChengse.AutoSize = True
        panelTop.Controls.Add(chkPrintChengse)

        AddLabel(panelTop, "税点：", 10, 75)
        txtTaxRate.Location = New Drawing.Point(50, 72) : txtTaxRate.Size = New Drawing.Size(50, 25)
        panelTop.Controls.Add(txtTaxRate)
        AddLabel(panelTop, "税收：", 110, 75)
        txtTax.Location = New Drawing.Point(150, 72) : txtTax.Size = New Drawing.Size(80, 25) : txtTax.ReadOnly = True
        panelTop.Controls.Add(txtTax)
        AddLabel(panelTop, "应收：", 240, 75)
        txtReceivable.Location = New Drawing.Point(280, 72) : txtReceivable.Size = New Drawing.Size(80, 25) : txtReceivable.ReadOnly = True
        panelTop.Controls.Add(txtReceivable)
        AddLabel(panelTop, "实收：", 370, 75)
        txtReceived.Location = New Drawing.Point(410, 72) : txtReceived.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(txtReceived)

        lblStoreMaterial.Text = "存料：" : lblStoreMaterial.Location = New Drawing.Point(500, 75) : lblStoreMaterial.AutoSize = True
        panelTop.Controls.Add(lblStoreMaterial)
        lblStoreMaterialVal.Text = "0" : lblStoreMaterialVal.Location = New Drawing.Point(530, 75) : lblStoreMaterialVal.AutoSize = True
        panelTop.Controls.Add(lblStoreMaterialVal)
        txtStoreMaterial.Location = New Drawing.Point(570, 72) : txtStoreMaterial.Size = New Drawing.Size(80, 25) : txtStoreMaterial.Text = "0.000" : txtStoreMaterial.Visible = False
        panelTop.Controls.Add(txtStoreMaterial)
        lblStoreBalance.Text = "余额：" : lblStoreBalance.Location = New Drawing.Point(660, 75) : lblStoreBalance.AutoSize = True
        panelTop.Controls.Add(lblStoreBalance)
        lblStoreBalanceVal.Text = "0" : lblStoreBalanceVal.Location = New Drawing.Point(690, 75) : lblStoreBalanceVal.AutoSize = True
        panelTop.Controls.Add(lblStoreBalanceVal)
        txtStoreBalance.Location = New Drawing.Point(730, 72) : txtStoreBalance.Size = New Drawing.Size(80, 25) : txtStoreBalance.Text = "0.00" : txtStoreBalance.Visible = False
        panelTop.Controls.Add(txtStoreBalance)

        Dim panelToolbar As New Panel()
        panelToolbar.Dock = DockStyle.Top : panelToolbar.Height = 40
        Me.Controls.Add(panelToolbar)
        btnToolbarSave.Text = "保存" : btnToolbarSave.Location = New Drawing.Point(10, 8) : btnToolbarSave.Size = New Drawing.Size(60, 25)
        panelToolbar.Controls.Add(btnToolbarSave)
        btnToolbarReset.Text = "重置" : btnToolbarReset.Location = New Drawing.Point(80, 8) : btnToolbarReset.Size = New Drawing.Size(60, 25)
        panelToolbar.Controls.Add(btnToolbarReset)
        btnToolbarPrintPreview.Text = "打印预览" : btnToolbarPrintPreview.Location = New Drawing.Point(150, 8) : btnToolbarPrintPreview.Size = New Drawing.Size(80, 25)
        panelToolbar.Controls.Add(btnToolbarPrintPreview)
        btnToolbarBatchEdit.Text = "批量编辑" : btnToolbarBatchEdit.Location = New Drawing.Point(240, 8) : btnToolbarBatchEdit.Size = New Drawing.Size(80, 25)
        panelToolbar.Controls.Add(btnToolbarBatchEdit)

        dgvSales.Dock = DockStyle.Fill : dgvSales.AllowUserToAddRows = False : dgvSales.BackgroundColor = Color.White
        Me.Controls.Add(dgvSales)

        Dim panelBottom As New Panel()
        panelBottom.Dock = DockStyle.Bottom : panelBottom.Height = 260
        Me.Controls.Add(panelBottom)

        dgvSummary.Location = New Drawing.Point(0, 0) : dgvSummary.Size = New Drawing.Size(1010, 35) : dgvSummary.AllowUserToAddRows = False : dgvSummary.BackgroundColor = Color.White
        panelBottom.Controls.Add(dgvSummary)

        btnRecoveryAdd.Text = "+" : btnRecoveryAdd.Location = New Drawing.Point(10, 20) : btnRecoveryAdd.Size = New Drawing.Size(35, 25)
        btnRecoveryRemove.Text = "-" : btnRecoveryRemove.Location = New Drawing.Point(10, 55) : btnRecoveryRemove.Size = New Drawing.Size(35, 25)
        Dim groupRecovery As New GroupBox()
        groupRecovery.Text = "回收" : groupRecovery.Location = New Drawing.Point(0, 40) : groupRecovery.Size = New Drawing.Size(55, 100)
        groupRecovery.Controls.Add(btnRecoveryAdd) : groupRecovery.Controls.Add(btnRecoveryRemove)
        panelBottom.Controls.Add(groupRecovery)

        dgvRecovery.Location = New Drawing.Point(55, 40) : dgvRecovery.Size = New Drawing.Size(955, 100) : dgvRecovery.AllowUserToAddRows = False : dgvRecovery.BackgroundColor = Color.White
        panelBottom.Controls.Add(dgvRecovery)

        dgvPayment.Location = New Drawing.Point(1010, 40) : dgvPayment.Size = New Drawing.Size(239, 180) : dgvPayment.AllowUserToAddRows = False : dgvPayment.BackgroundColor = Color.White
        panelBottom.Controls.Add(dgvPayment)

        Dim groupRemarks As New GroupBox()
        groupRemarks.Text = "备注" : groupRemarks.Location = New Drawing.Point(0, 150) : groupRemarks.Size = New Drawing.Size(1010, 79)
        txtRemarks.Location = New Drawing.Point(5, 20) : txtRemarks.Size = New Drawing.Size(970, 50) : txtRemarks.Multiline = True
        groupRemarks.Controls.Add(txtRemarks)
        panelBottom.Controls.Add(groupRemarks)

        dgvDeleted.Visible = False
        dgvDeleted.Columns.Add("c0", "id") : dgvDeleted.Columns.Add("c1", "code") : dgvDeleted.Columns.Add("c2", "orderid") : dgvDeleted.Columns.Add("c3", "qty") : dgvDeleted.Columns.Add("c4", "weight")
        Me.Controls.Add(dgvDeleted)

        AddHandler btnToolbarSave.Click, AddressOf SaveEdit
        AddHandler btnToolbarReset.Click, Sub() Form_Load(Nothing, Nothing) : CalculateSummary()
        AddHandler btnToolbarPrintPreview.Click, AddressOf PrintDocument
        AddHandler btnToolbarBatchEdit.Click, Sub() MsgBox("批量编辑功能需打开批量修改窗口")
        AddHandler btnRecoveryAdd.Click, AddressOf BtnRecoveryAdd_Click
        AddHandler btnRecoveryRemove.Click, AddressOf BtnRecoveryRemove_Click
        AddHandler txtTaxRate.TextChanged, Sub() CalculateSummary()
        AddHandler radioPresaleYes.CheckedChanged, Sub() If radioPresaleYes.Checked Then txtPresaleNumber.Enabled = True : txtPresaleNumber.BackColor = Color.White
        AddHandler radioPresaleNo.CheckedChanged, Sub() If radioPresaleNo.Checked Then txtPresaleNumber.Enabled = False : txtPresaleNumber.BackColor = Color.Silver
        AddHandler txtPresaleNumber.TextChanged, AddressOf TxtPresaleNumber_TextChanged
        AddHandler txtPresaleNumber.KeyDown, AddressOf TxtPresaleNumber_KeyDown
        AddHandler dgvSales.CellEndEdit, AddressOf DgvSales_CellEndEdit
        AddHandler dgvSales.UserDeletingRow, AddressOf DgvSales_UserDeletingRow
        AddHandler dgvRecovery.CellEndEdit, AddressOf DgvRecovery_CellEndEdit
        AddHandler dgvPayment.CellEndEdit, Sub(s, ev) If dgvPayment.Rows.Count >= 2 Then txtReceived.Text = SafeString(dgvPayment.Rows(dgvPayment.Rows.Count - 1).Cells(2).Value)
    End Sub

    Private Sub AddLabel(parent As Control, text As String, x As Integer, y As Integer)
        Dim lbl As New Label() With {.Text = text, .Location = New Drawing.Point(x, y), .AutoSize = True}
        parent.Controls.Add(lbl)
    End Sub

    ' ========== SQL文本处理 ==========
    Private Function SafeSQLForEdit(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Return text.Replace("'", "''")
    End Function

    Private Function GetWriteAffectedRows() As String
        Try
            Dim dt As DataTable = DatabaseModule.ExecuteQuery("SELECT ROW_COUNT() AS rowcnt", MySQL_Write)
            If dt.Rows.Count > 0 Then Return SafeString(dt.Rows(0)("rowcnt"))
        Catch
        End Try
        Return "0"
    End Function

    ' ========== 回补库存行 ==========
    Private Function RestoreStockRow(productCode As String, quantity As String, netWeight As String, warehouseID As String) As Boolean
        If Val(quantity) <= 0 AndAlso Val(netWeight) <= 0 Then Return True
        If String.IsNullOrEmpty(warehouseID) Then warehouseID = UserDepartment
        productCode = SafeSQLForEdit(productCode)
        Dim sql As String = $"SELECT id FROM xipunum_erp_shop_kucun WHERE poduct_code='{productCode}' AND kufang='{warehouseID}' LIMIT 1"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Write)
        If dt.Rows.Count = 0 Then
            sql = $"INSERT INTO xipunum_erp_shop_kucun (poduct_code, quantity, jinzhong, kufang, cjuser, creationtime) VALUES ('{productCode}', '{quantity}', '{netWeight}', '{warehouseID}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
            DatabaseModule.ExecuteCommand(sql, MySQL_Write)
        Else
            sql = $"UPDATE xipunum_erp_shop_kucun SET quantity = quantity + '{quantity}',jinzhong = jinzhong + '{netWeight}' WHERE poduct_code = '{productCode}' AND kufang='{warehouseID}'"
            DatabaseModule.ExecuteCommand(sql, MySQL_Write)
        End If
        Return Val(GetWriteAffectedRows()) >= 1
    End Function

    ' ========== 是否仅金重商品 ==========
    Private Function IsWeightOnlyProduct(productCode As String) As Boolean
        productCode = SafeSQLForEdit(productCode)
        Dim sql As String = $"SELECT CASE WHEN COALESCE(d.lingxiao, '') = '' THEN '否' ELSE d.lingxiao END AS lingxiao,b.sales_unit AS sales_unit FROM xipunum_erp_shop AS b LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = b.item_number AND b.item_number != '' WHERE b.poduct_code='{productCode}' LIMIT 1"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
        If dt.Rows.Count > 0 Then
            If SafeString(dt.Rows(0)("lingxiao")) = "是" Then Return True
            If SafeString(dt.Rows(0)("sales_unit")) = "重量" Then Return True
        End If
        Return False
    End Function

    Private Function BuildStockCondition(diffQty As String, diffWeight As String, isWeightOnly As Boolean) As String
        Dim c As String = ""
        If Not isWeightOnly AndAlso Val(diffQty) > 0 Then c &= $" AND quantity >='{diffQty}'"
        If Val(diffWeight) > 0 Then c &= $" AND jinzhong>='{diffWeight}'"
        Return c
    End Function

    Private Function ResolveDeductWarehouse(productCode As String, diffQty As String, diffWeight As String, preferredWH As String, isWeightOnly As Boolean) As String
        productCode = SafeSQLForEdit(productCode)
        Dim cond As String = BuildStockCondition(diffQty, diffWeight, isWeightOnly)
        If cond = "" Then Return If(preferredWH <> "", preferredWH, UserDepartment)
        If preferredWH <> "" Then
            Dim dt As DataTable = DatabaseModule.ExecuteQuery($"SELECT id FROM xipunum_erp_shop_kucun WHERE poduct_code='{productCode}' AND kufang='{preferredWH}'{cond} LIMIT 1")
            If dt.Rows.Count >= 1 Then Return preferredWH
        End If
        If UserDepartment <> "" Then
            Dim dt As DataTable = DatabaseModule.ExecuteQuery($"SELECT id FROM xipunum_erp_shop_kucun WHERE poduct_code='{productCode}' AND kufang='{UserDepartment}'{cond} LIMIT 1")
            If dt.Rows.Count >= 1 Then Return UserDepartment
        End If
        Return ""
    End Function

    Private Sub GetRowSalesLimit(gridRow As Integer, ByRef maxWeight As String, ByRef maxQty As String)
        Dim stockWeight As String = SafeString(dgvSales.Rows(gridRow).Cells(25).Value)
        Dim stockQty As String = SafeString(dgvSales.Rows(gridRow).Cells(14).Value)
        Dim origWeight As String = "0", origQty As String = "0"
        Dim detailID As String = SafeString(dgvSales.Rows(gridRow).Cells(26).Value)
        If detailID <> "" Then
            Dim dt As DataTable = DatabaseModule.ExecuteQuery($"SELECT quantity,net_weight FROM xipunum_erp_outbound WHERE id='{detailID}' LIMIT 1")
            If dt.Rows.Count > 0 Then
                origQty = SafeString(dt.Rows(0)("quantity")) : origWeight = SafeString(dt.Rows(0)("net_weight"))
            End If
            If salesOrderID <> "" Then
                Dim rowCode As String = SafeString(dgvSales.Rows(gridRow).Cells(1).Value)
                dt = DatabaseModule.ExecuteQuery($"SELECT quantity,net_weight FROM xipunum_erp_outbound WHERE order_id='{salesOrderID}' AND poduct_code='{rowCode}' ORDER BY id DESC LIMIT 1")
                If dt.Rows.Count > 0 Then
                    origQty = SafeString(dt.Rows(0)("quantity")) : origWeight = SafeString(dt.Rows(0)("net_weight"))
                End If
            End If
        End If
        maxWeight = (Val(origWeight) + Val(stockWeight)).ToString()
        maxQty = (Val(origQty) + Val(stockQty)).ToString()
    End Sub

    Private Function AutoLimitSalesRow(gridRow As Integer, ByRef salesWeight As String, ByRef salesQty As String) As Boolean
        Dim adj As Boolean = False
        If Val(salesWeight) < 0 Then salesWeight = "0" : adj = True
        If Val(salesQty) < 0 Then salesQty = "0" : adj = True
        Dim maxW As String = "", maxQ As String = ""
        GetRowSalesLimit(gridRow, maxW, maxQ)
        If Val(salesWeight) > Val(maxW) Then salesWeight = maxW : adj = True
        If Val(salesQty) > Val(maxQ) Then salesQty = Math.Round(Val(maxQ), 2).ToString() : adj = True
        If adj Then
            dgvSales.Rows(gridRow).Cells(9).Value = salesWeight
            dgvSales.Rows(gridRow).Cells(17).Value = salesQty
        End If
        Return adj
    End Function

    Private Function DeductStock(productCode As String, deductQty As String, deductWeight As String, warehouseID As String) As Boolean
        If String.IsNullOrEmpty(warehouseID) Then warehouseID = UserDepartment
        productCode = SafeSQLForEdit(productCode)
        Dim isWO As Boolean = IsWeightOnlyProduct(productCode)
        Dim wh As String = ResolveDeductWarehouse(productCode, deductQty, deductWeight, warehouseID, isWO)
        If wh = "" Then Return False
        Dim cond As String = BuildStockCondition(deductQty, deductWeight, isWO)
        DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '{deductQty}',jinzhong = jinzhong - '{deductWeight}' WHERE poduct_code = '{productCode}' AND kufang='{wh}'{cond}", MySQL_Write)
        Return Val(GetWriteAffectedRows()) >= 1
    End Function

    Private Function AdjustStock(productCode As String, origQty As String, origWeight As String, newQty As String, newWeight As String, warehouseID As String) As Boolean
        If String.IsNullOrEmpty(warehouseID) Then warehouseID = UserDepartment
        productCode = SafeSQLForEdit(productCode)
        Dim diffQty As String = (Val(newQty) - Val(origQty)).ToString()
        Dim diffWeight As String = (Val(newWeight) - Val(origWeight)).ToString()
        If Val(diffQty) = 0 AndAlso Val(diffWeight) = 0 Then Return True
        Dim actualWH As String = warehouseID
        If Val(diffQty) > 0 OrElse Val(diffWeight) > 0 Then
            Dim isWO As Boolean = IsWeightOnlyProduct(productCode)
            actualWH = ResolveDeductWarehouse(productCode, diffQty, diffWeight, warehouseID, isWO)
            If actualWH = "" Then Return False
            Dim cond As String = BuildStockCondition(diffQty, diffWeight, isWO)
            DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '{diffQty}',jinzhong = jinzhong - '{diffWeight}' WHERE poduct_code = '{productCode}' AND kufang='{actualWH}'{cond}", MySQL_Write)
            If Val(GetWriteAffectedRows()) < 1 Then Return False
        Else
            DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '{diffQty}',jinzhong = jinzhong - '{diffWeight}' WHERE poduct_code = '{productCode}' AND kufang='{actualWH}'", MySQL_Write)
            If Val(GetWriteAffectedRows()) < 1 Then
                If actualWH <> UserDepartment Then
                    actualWH = UserDepartment
                    DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_shop_kucun SET quantity = quantity - '{diffQty}',jinzhong = jinzhong - '{diffWeight}' WHERE poduct_code = '{productCode}' AND kufang='{actualWH}'", MySQL_Write)
                End If
                If Val(GetWriteAffectedRows()) < 1 Then
                    Dim rQty As String = (0 - Val(diffQty)).ToString(), rWt As String = (0 - Val(diffWeight)).ToString()
                    If Not RestoreStockRow(productCode, rQty, rWt, actualWH) Then Return False
                End If
            End If
        End If
        Return True
    End Function

    Private Sub SaveFailedRollback(reason As String)
        If saveInProgress = 1 Then
            DatabaseModule.ExecuteCommand("ROLLBACK", MySQL_Write)
            saveInProgress = 0
        End If
        ShowError(reason)
    End Sub

    ' ========== 窗口加载 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        If GlobalSalesOrderNumber <> "" Then txtOrderNumber.Text = GlobalSalesOrderNumber
        tableEditState = 0 : paymentEditState = 0 : recoveryEditState = 0

        Dim orderID As String = "", customerName As String = "", customerTel As String = ""
        Dim recoveryID As String = "", presaleNumber As String = "", presaleDeposit As String = ""
        Dim salesman As String = "", taxAmount As String = "", taxPoint As String = "", remarks As String = ""
        Dim orderState As String = "", recoveryNumber As String = "", memberCode As String = "", memberID As String = ""
        Dim thisStoreMaterial As String = "0.000", thisStoreBalance As String = "0.00"
        Dim warehouseID As String = "", salesTime As String = ""
        Dim totalStoreMaterial As String = "0.000", totalStoreBalance As String = "0.00"
        Dim introducer As String = "", introducerName As String = "", introducerInfo As String = ""
        Dim salesFactory As String = "", invoiceFlag As String = ""

        Dim sql As String = $"SELECT a.xsfactory as xsfactory,a.waibu_number as waibu_number,a.shopping_guide as shopping_guide,h.name AS hname,a.customer_code as acustomer_code,a.id AS aid,a.state AS astate,b.NAME AS bname,b.tel AS btel,d.id AS did,a.retrea_umber as aretrea_umber,a.presale_number AS apresale_number,ROUND(c.deposit,2) AS cdeposit,a.salesman AS asalesman,a.creationtime AS acreationtime,ROUND(a.taxamount,2) AS ataxamount,a.taxpoint AS ataxpoint,a.remarks AS aremarks,a.pling as piling,b.memberid AS memberid,e.department AS kufangdi,CAST(ROUND(COALESCE(f.number, 0), 3) AS DECIMAL (20, 3)) AS cunliao,CAST(ROUND(COALESCE(g.number, 0), 2) AS DECIMAL (20, 2)) AS cunyuan,a.fapiao as afapiao FROM xipunum_erp_outbound_order AS a LEFT JOIN xipunum_erp_member AS b ON b.customer_code = a.customer_code LEFT JOIN xipunum_erp_presale_order AS c ON c.presale_umber = a.presale_number LEFT JOIN xipunum_erp_retreat_order AS d ON d.retrea_umber = a.retrea_umber INNER JOIN xipunum_erp_user e ON e.USER = a.cjuser LEFT JOIN xipunum_erp_member_cq f ON f.customer_code = a.customer_code and a.creationtime=f.creationtime and e.department=f.kufang and f.cunqu='存' and f.type='料' LEFT JOIN xipunum_erp_member_cq g ON g.customer_code = a.customer_code and a.creationtime=g.creationtime and e.department=g.kufang and g.cunqu='存' and g.type='元' LEFT JOIN xipunum_erp_member AS h ON h.memberid = a.shopping_guide WHERE a.settlement_number = '{txtOrderNumber.Text}'  LIMIT 1"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
        If dt.Rows.Count > 0 Then
            Dim r As DataRow = dt.Rows(0)
            orderID = SafeString(r("aid")) : customerName = SafeString(r("bname")) : customerTel = SafeString(r("btel"))
            recoveryID = SafeString(r("did")) : recoveryNumber = SafeString(r("aretrea_umber"))
            presaleNumber = SafeString(r("apresale_number")) : presaleDeposit = SafeString(r("cdeposit"))
            salesman = SafeString(r("asalesman")) : taxAmount = SafeString(r("ataxamount"))
            taxPoint = SafeString(r("ataxpoint")) : remarks = SafeString(r("aremarks"))
            orderState = SafeString(r("astate")) : memberCode = SafeString(r("acustomer_code"))
            salesPling = SafeString(r("piling")) : memberID = SafeString(r("memberid"))
            thisStoreMaterial = SafeString(r("cunliao")) : thisStoreBalance = SafeString(r("cunyuan"))
            warehouseID = SafeString(r("kufangdi")) : salesTime = SafeString(r("acreationtime"))
            introducer = SafeString(r("shopping_guide")) : introducerName = SafeString(r("hname"))
            salesFactory = SafeString(r("xsfactory")) : invoiceFlag = SafeString(r("afapiao"))
        End If
        salesOrderStatus = orderState
        If customerName <> "" Then customerName = $"{customerName}({memberID})"
        If introducer <> "" Then introducerInfo = $"{introducerName}({introducer})"

        If SalesQueryInvoice = "1" Then groupInvoice.Visible = False Else groupInvoice.Visible = True
        If invoiceFlag = "1" Then radioNoInvoice.Checked = False : radioInvoice.Checked = True Else radioNoInvoice.Checked = True : radioInvoice.Checked = False

        Dim isNormal As Boolean = (salesOrderStatus = "正常")
        dgvPayment.Enabled = Not isNormal : dgvSummary.Enabled = Not isNormal
        txtTaxRate.Enabled = Not isNormal : radioNewRow.Enabled = Not isNormal : radioNewCol.Enabled = Not isNormal
        radioPresaleYes.Enabled = Not isNormal : radioPresaleNo.Enabled = Not isNormal : txtRemarks.Enabled = Not isNormal
        btnRecoveryAdd.Enabled = Not isNormal : btnRecoveryRemove.Enabled = Not isNormal
        cmbIntroducer.Enabled = Not isNormal : cmbSalesFactory.Enabled = Not isNormal
        radioNoInvoice.Enabled = Not isNormal : radioInvoice.Enabled = Not isNormal
        btnToolbarSave.Enabled = Not isNormal : btnToolbarReset.Enabled = Not isNormal : btnToolbarBatchEdit.Enabled = Not isNormal

        cmbGuide.Text = salesman : txtMemberName.Text = customerName : txtPhone.Text = customerTel
        txtTaxRate.Text = taxPoint : txtTax.Text = taxAmount : txtPresaleNumber.Text = presaleNumber
        txtPresaleDeposit.Text = presaleDeposit : txtRemarks.Text = remarks
        cmbIntroducer.Text = introducerInfo : cmbSalesFactory.Text = salesFactory

        If presaleNumber = "" Then radioPresaleYes.Checked = False : radioPresaleNo.Checked = True Else radioPresaleYes.Checked = True : radioPresaleNo.Checked = False

        salesOrderID = orderID : salesRecoveryID = recoveryID : salesRecoveryNumber = recoveryNumber : salesMemberCode = memberCode
        txtTax.Enabled = False : txtReceivable.Enabled = False
        chkPrintDoc.Checked = False : chkPrintPreview.Checked = False : chkPrintChengse.Checked = False

        LoadVoucherStyles() : InitSalesGrid() : InitRecoveryGrid() : InitPaymentGrid()

        If isNormal Then LoadSalesDetail() Else LoadSalesGridData()
        If salesRecoveryID <> "" Then LoadRecoveryGridData()
        LoadPaymentGridData() : CalculateSummary()

        sql = $"SELECT ROUND(IFNULL(SUM(CASE WHEN cq.cunqu = '存' AND cq.type = '料' THEN cq.number ELSE 0 END), 0) - IFNULL(SUM(CASE WHEN cq.cunqu = '欠' AND cq.type = '料' THEN cq.number ELSE 0 END), 0), 3) AS jieyuliao,ROUND(IFNULL(SUM(CASE WHEN cq.cunqu = '存' AND cq.type = '元' THEN cq.number ELSE 0 END), 0) - IFNULL(SUM(CASE WHEN cq.cunqu = '欠' AND cq.type = '元' THEN cq.number ELSE 0 END), 0), 2) AS jieyuyuan FROM xipunum_erp_member AS a INNER JOIN xipunum_erp_user AS b ON b.USER = a.cjuser LEFT JOIN xipunum_erp_member_cq AS cq ON cq.customer_code = a.customer_code AND cq.kufang = '{warehouseID}' and cq.creationtime <='{salesTime}' WHERE a.memberid = '{memberID}' GROUP BY a.customer_code, a.memberid"
        dt = DatabaseModule.ExecuteQuery(sql)
        If dt.Rows.Count > 0 Then
            totalStoreMaterial = SafeString(dt.Rows(0)("jieyuliao")) : totalStoreBalance = SafeString(dt.Rows(0)("jieyuyuan"))
        End If

        If salesPling = "批发" Then
            lblStoreMaterial.Visible = True : lblStoreMaterialVal.Visible = True : txtStoreMaterial.Visible = True
            lblStoreBalance.Visible = True : lblStoreBalanceVal.Visible = True : txtStoreBalance.Visible = True
            lblStoreMaterialVal.Text = totalStoreMaterial : txtStoreMaterial.Text = thisStoreMaterial
            lblStoreBalanceVal.Text = totalStoreBalance : txtStoreBalance.Text = thisStoreBalance
        Else
            lblStoreMaterial.Visible = False : lblStoreMaterialVal.Visible = False : txtStoreMaterial.Visible = False
            lblStoreBalance.Visible = False : lblStoreBalanceVal.Visible = False : txtStoreBalance.Visible = False
        End If
    End Sub

    Private Sub LoadVoucherStyles()
        cmbVoucherStyle.Items.Clear()
        Dim dt As DataTable = DatabaseModule.ExecuteQuery("SELECT * FROM xipunum_erp_voucher where (type= '出库' or type= '置换') and state= '0' order by id ASC")
        For Each r As DataRow In dt.Rows : cmbVoucherStyle.Items.Add(SafeString(r("name"))) : Next
        cmbVoucherStyle.SelectedIndex = -1
    End Sub

    Private Sub InitSalesGrid()
        dgvSales.Columns.Clear()
        Dim headers() As String, widths() As Integer
        If salesOrderStatus = "正常" Then
            headers = {"", "商品编码", "商品名称", "款号", "规格", "材质", "圈口/长度", "成色", "单件重", "金重", "重量", "成本工费", "参考工费", "成本附加费", "库存", "销售单价", "销售金额", "数量", "原附加费", "销售克价", "销售工费", "销售附加费", "折扣", "实收金额", "导购员"}
            widths = {45, 100, 140, 70, 70, 70, 70, 70, 70, 70, 70, 0, 0, 0, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70}
        Else
            headers = {"", "商品编码", "商品名称", "款号", "规格", "材质", "圈口/长度", "成色", "单件重", "金重", "重量", "成本工费", "参考工费", "成本附加费", "原库存", "销售单价", "销售金额", "数量", "原附加费", "销售克价", "销售工费", "销售附加费", "折扣", "实收金额", "导购员", "库存金重", "销售id", "操作"}
            widths = {45, 100, 140, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 0, 0, 60}
        End If
        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn() With {.HeaderText = headers(i), .Name = "col" & i, .Width = widths(i)}
            If widths(i) = 0 Then col.Visible = False
            dgvSales.Columns.Add(col)
        Next
    End Sub

    Private Sub InitRecoveryGrid()
        dgvRecovery.Columns.Clear()
        Dim headers() As String = {"", "商品名称", "数量", "总重", "金重", "成色", "回收克价", "其他费用", "回收金额", "备注", "导购员"}
        Dim widths() As Integer = {45, 200, 100, 100, 100, 100, 100, 100, 100, 200, 100}
        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn() With {.HeaderText = headers(i), .Name = "rcol" & i, .Width = widths(i)}
            dgvRecovery.Columns.Add(col)
        Next
        Dim retreatDt As DataTable = DatabaseModule.ExecuteQuery("SELECT * FROM xipunum_erp_retreat_title WHERE 1=1 order by id asc")
        Dim tCol As New DataGridViewComboBoxColumn() With {.Name = "rcol1_combo", .HeaderText = "商品名称", .Width = 200}
        For Each r As DataRow In retreatDt.Rows : tCol.Items.Add(SafeString(r("title"))) : Next
        dgvRecovery.Columns.RemoveAt(1) : dgvRecovery.Columns.Insert(1, tCol)
        Dim guideSQL As String = If(UserPermission = "全部", "SELECT name FROM xipunum_erp_user where state='0' order by id ASC", $"SELECT name FROM xipunum_erp_user where department='{UserDepartment}' and state='0' order by id ASC")
        Dim gDt As DataTable = DatabaseModule.ExecuteQuery(guideSQL)
        Dim gCol As New DataGridViewComboBoxColumn() With {.Name = "rcol10_combo", .HeaderText = "导购员", .Width = 100}
        For Each r As DataRow In gDt.Rows : gCol.Items.Add(SafeString(r("name"))) : Next
        dgvRecovery.Columns.RemoveAt(10) : dgvRecovery.Columns.Insert(10, gCol)
    End Sub

    Private Sub InitPaymentGrid()
        dgvPayment.Columns.Clear()
        Dim headers() As String = {"序号", "支付方式", "金额", "id"}
        Dim widths() As Integer = {55, 75, 65, 0}
        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn() With {.HeaderText = headers(i), .Name = "pcol" & i, .Width = widths(i)}
            If widths(i) = 0 Then col.Visible = False
            dgvPayment.Columns.Add(col)
        Next
    End Sub

    Private Sub LoadSalesDetail()
        dgvSales.Rows.Clear()
        Dim sql As String = $"SELECT a.id AS xiaoshouid,a.poduct_code AS outbound_product_code,b.product_name AS shop_product_name,b.item_number AS shop_item_number,c.title AS specification_title,b.caizhi AS material,b.quandu AS purity,MAX(d.company_condition) AS store_company_condition,b.single AS unit_weight,a.net_weight AS shop_gold_weight,CAST(ROUND(COALESCE(CASE WHEN b.quantity = 0 THEN a.net_weight ELSE b.weight / b.quantity * a.quantity END,a.net_weight),3) AS DECIMAL (20,3)) AS outbound_net_weight,a.basic_cost AS basic_cost,a.premium_cost AS apremium_cost,MAX(d.company_surcharge) AS store_company_surcharge,f.quantity AS shop_quantity,CAST(ROUND(COALESCE(f.jinzhong,0),3) AS DECIMAL (20,3)) AS kucunjin,a.xiaodan_amount AS xiaodan_amount,a.xiao_amount AS xiao_amount,a.quantity AS outbound_quantity,MAX(d.sales_surcharge) AS store_sales_surcharge,a.gold_price AS agold_price,a.sales_cost AS asales_cost,a.sales_surcharge AS asales_surcharge,a.zhekou AS discount,a.settlement AS settlement,a.shopping_guide AS ashopping_guide,g.name as gname FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_specs AS c ON c.id = b.specification_id INNER JOIN xipunum_erp_store AS d ON d.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_shop_kucun AS f ON f.poduct_code = a.poduct_code AND f.kufang = a.kufang LEFT JOIN xipunum_erp_user AS g ON g.user = a.shopping_guide WHERE a.order_id = '{salesOrderID}' GROUP BY a.id, a.poduct_code, b.product_name, b.item_number, c.title, b.caizhi, b.quandu, b.single, a.net_weight, a.basic_cost, a.premium_cost, a.xiaodan_amount, a.xiao_amount, a.quantity, a.gold_price, a.sales_cost, a.sales_surcharge, a.zhekou, a.settlement, c.shuliang, b.weight, b.quantity, f.quantity, f.jinzhong ORDER BY a.id ASC"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
        For Each d As DataRow In dt.Rows
            Dim idx As Integer = dgvSales.Rows.Add()
            Dim r As DataGridViewRow = dgvSales.Rows(idx)
            r.Cells(0).Value = (idx + 1).ToString() : r.Cells(1).Value = SafeString(d("outbound_product_code"))
            r.Cells(2).Value = SafeString(d("shop_product_name")) : r.Cells(3).Value = SafeString(d("shop_item_number"))
            r.Cells(4).Value = SafeString(d("specification_title")) : r.Cells(5).Value = SafeString(d("material"))
            r.Cells(6).Value = SafeString(d("purity")) : r.Cells(7).Value = SafeString(d("store_company_condition"))
            r.Cells(8).Value = SafeString(d("unit_weight")) : r.Cells(9).Value = SafeString(d("shop_gold_weight"))
            r.Cells(10).Value = SafeString(d("outbound_net_weight")) : r.Cells(11).Value = SafeString(d("basic_cost"))
            r.Cells(12).Value = SafeString(d("apremium_cost")) : r.Cells(13).Value = SafeString(d("store_company_surcharge"))
            r.Cells(14).Value = SafeString(d("shop_quantity")) : r.Cells(15).Value = SafeString(d("xiaodan_amount"))
            r.Cells(16).Value = SafeString(d("xiao_amount")) : r.Cells(17).Value = SafeString(d("outbound_quantity"))
            r.Cells(18).Value = SafeString(d("store_sales_surcharge")) : r.Cells(19).Value = SafeString(d("agold_price"))
            r.Cells(20).Value = SafeString(d("asales_cost")) : r.Cells(21).Value = SafeString(d("asales_surcharge"))
            r.Cells(22).Value = SafeString(d("discount")) : r.Cells(23).Value = SafeString(d("settlement"))
            r.Cells(24).Value = SafeString(d("gname")) : r.Cells(25).Value = SafeString(d("kucunjin"))
            r.Cells(26).Value = SafeString(d("xiaoshouid"))
        Next
        dgvSales.ReadOnly = True
    End Sub

    Private Sub LoadSalesGridData()
        dgvSales.Rows.Clear()
        Dim sql As String = $"SELECT CASE WHEN COALESCE(f.title, '') = '' THEN '未匹配' ELSE f.title END AS pinlei,CASE WHEN COALESCE(d.lingxiao, '') = '' THEN '否' ELSE d.lingxiao END AS lingxiao,a.poduct_code AS outbound_product_code,b.product_name AS shop_product_name,b.item_number AS shop_item_number,COALESCE(e1.title, e2.title, '无数据') AS guige,b.caizhi as caizhi,b.quandu as quandu,g.company_condition as chengse,CASE WHEN COALESCE(d.lingxiao, '') = '是' THEN a.net_weight ELSE b.single END AS danzhong,a.net_weight AS xsjinzhong,CASE WHEN COALESCE(d.lingxiao, '') = '是' THEN a.net_weight ELSE CAST(ROUND(b.single*a.quantity, 3) AS DECIMAL (10, 3)) END AS zongzhong,a.basic_cost AS basic_cost,a.premium_cost AS apremium_cost,g.company_surcharge as company_surcharge,h.quantity AS shop_quantity,a.xiaodan_amount AS xiaodan_amount,a.xiao_amount AS xiao_amount,a.quantity AS outbound_quantity,g.sales_surcharge as sales_surcharge,a.gold_price AS agold_price,a.sales_cost AS asales_cost,a.sales_surcharge AS asales_surcharge,a.zhekou AS discount,a.settlement AS settlement,i.name as daosgou,CAST(ROUND(COALESCE(h.jinzhong,0),3) AS DECIMAL (20,3)) AS kucunjin,a.id AS xiaoshouid FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_type AS c ON c.id = a.kufang LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = b.item_number AND b.item_number != '' LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE(e1.category_id, e2.category_id) INNER JOIN xipunum_erp_store AS g ON g.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_shop_kucun AS h ON h.poduct_code = a.poduct_code AND h.kufang = a.kufang LEFT JOIN xipunum_erp_user AS i ON i.user = a.shopping_guide WHERE a.order_id = '{salesOrderID}' ORDER BY a.id ASC"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
        For Each d As DataRow In dt.Rows
            Dim idx As Integer = dgvSales.Rows.Add()
            Dim r As DataGridViewRow = dgvSales.Rows(idx)
            r.Cells(0).Value = (idx + 1).ToString() : r.Cells(1).Value = SafeString(d("outbound_product_code"))
            r.Cells(2).Value = SafeString(d("shop_product_name")) : r.Cells(3).Value = SafeString(d("shop_item_number"))
            r.Cells(4).Value = SafeString(d("guige")) : r.Cells(5).Value = SafeString(d("caizhi"))
            r.Cells(6).Value = SafeString(d("quandu")) : r.Cells(7).Value = SafeString(d("chengse"))
            r.Cells(8).Value = SafeString(d("danzhong")) : r.Cells(9).Value = SafeString(d("xsjinzhong"))
            r.Cells(10).Value = SafeString(d("zongzhong")) : r.Cells(11).Value = SafeString(d("basic_cost"))
            r.Cells(12).Value = SafeString(d("apremium_cost")) : r.Cells(13).Value = SafeString(d("company_surcharge"))
            r.Cells(14).Value = SafeString(d("shop_quantity")) : r.Cells(15).Value = SafeString(d("xiaodan_amount"))
            r.Cells(16).Value = SafeString(d("xiao_amount")) : r.Cells(17).Value = SafeString(d("outbound_quantity"))
            r.Cells(18).Value = SafeString(d("sales_surcharge")) : r.Cells(19).Value = SafeString(d("agold_price"))
            r.Cells(20).Value = SafeString(d("asales_cost")) : r.Cells(21).Value = SafeString(d("asales_surcharge"))
            r.Cells(22).Value = SafeString(d("discount")) : r.Cells(23).Value = SafeString(d("settlement"))
            r.Cells(24).Value = SafeString(d("daosgou")) : r.Cells(25).Value = SafeString(d("kucunjin"))
            r.Cells(26).Value = SafeString(d("xiaoshouid"))
        Next
        AddBlankSalesRow()
    End Sub

    Private Sub AddBlankSalesRow()
        Dim cnt As Integer = dgvSales.Rows.Count
        If cnt > 0 AndAlso SafeString(dgvSales.Rows(cnt - 1).Cells(1).Value) <> "" Then
            Dim idx As Integer = dgvSales.Rows.Add()
            Dim r As DataGridViewRow = dgvSales.Rows(idx)
            r.Cells(0).Value = (idx + 1).ToString()
            For i As Integer = 8 To 23 : r.Cells(i).Value = If(i = 9 OrElse i = 10, "0.000", If(i = 22, "1", "0")) : Next
            r.Cells(25).Value = "0"
        End If
    End Sub

    Private Sub LoadRecoveryGridData()
        dgvRecovery.Rows.Clear()
        Dim sql As String = $"SELECT b.title AS product_name,a.quantity AS quantity,a.total AS total,a.jin_zhong AS jin_zhong,a.chengse AS chengse,a.price AS price,a.qita_price AS qita_price,a.retreat_amount AS retreat_amount,c.name AS shopping_guide,a.remarks AS remarks FROM xipunum_erp_retreat AS a INNER JOIN xipunum_erp_retreat_title AS b ON b.id = a.product_name LEFT JOIN xipunum_erp_user AS c ON c.user = a.shopping_guide WHERE a.order_id = '{salesRecoveryID}' ORDER BY a.id DESC"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
        For Each d As DataRow In dt.Rows
            Dim idx As Integer = dgvRecovery.Rows.Add()
            Dim r As DataGridViewRow = dgvRecovery.Rows(idx)
            r.Cells(0).Value = "" : r.Cells(1).Value = SafeString(d("product_name"))
            r.Cells(2).Value = SafeString(d("quantity")) : r.Cells(3).Value = SafeString(d("total"))
            r.Cells(4).Value = SafeString(d("jin_zhong")) : r.Cells(5).Value = SafeString(d("chengse"))
            r.Cells(6).Value = SafeString(d("price")) : r.Cells(7).Value = SafeString(d("qita_price"))
            r.Cells(8).Value = SafeString(d("retreat_amount")) : r.Cells(9).Value = SafeString(d("remarks"))
            r.Cells(10).Value = SafeString(d("shopping_guide"))
        Next
        If salesOrderStatus = "正常" Then dgvRecovery.ReadOnly = True Else dgvRecovery.ReadOnly = False
    End Sub

    Private Sub LoadPaymentGridData()
        dgvPayment.Rows.Clear()
        Dim dt As DataTable = DatabaseModule.ExecuteQuery("SELECT id,name FROM xipunum_erp_pay where state='0' order by id ASC")
        For Each d As DataRow In dt.Rows
            Dim payID As String = SafeString(d("id")), payName As String = SafeString(d("name"))
            Dim payAmount As String = "0.00"
            Dim amtDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT CAST(ROUND(xianjin, 2) AS DECIMAL (20, 2)) AS xianjin FROM xipunum_erp_shoukuan WHERE settlement_number = '{txtOrderNumber.Text}' and type= '{payID}'  LIMIT 1")
            If amtDt.Rows.Count > 0 Then payAmount = SafeString(amtDt.Rows(0)("xianjin"))
            If payAmount = "" Then payAmount = "0.00"
            Dim idx As Integer = dgvPayment.Rows.Add()
            dgvPayment.Rows(idx).Cells(0).Value = (idx + 1).ToString()
            dgvPayment.Rows(idx).Cells(1).Value = payName
            dgvPayment.Rows(idx).Cells(2).Value = payAmount
            dgvPayment.Rows(idx).Cells(3).Value = payID
        Next
        Dim tIdx As Integer = dgvPayment.Rows.Add()
        dgvPayment.Rows(tIdx).Cells(1).Value = "合计"
    End Sub

    ' ========== 数据统计汇总 ==========
    Private Sub CalculateSummary()
        Dim calcTaxRate As String = If(txtTaxRate.Text = "", "0", txtTaxRate.Text)
        quantityTotal = "0" : weightTotal = "0" : weightPrice = "0" : salesAmount = "0" : actualAmount = "0"
        basicCostTotal = "0" : premiumCostTotal = "0" : salesCostTotal = "0" : costSurchargeTotal = "0" : salesSurchargeTotal = "0"
        Dim loopCnt As Integer = If(salesOrderStatus <> "正常", Math.Max(0, dgvSales.Rows.Count - 1), dgvSales.Rows.Count)
        For i As Integer = 0 To loopCnt - 1
            quantityTotal = (Val(quantityTotal) + Val(SafeString(dgvSales.Rows(i).Cells(17).Value))).ToString()
            weightTotal = (Val(weightTotal) + Val(SafeString(dgvSales.Rows(i).Cells(9).Value))).ToString()
            weightPrice = (Val(weightPrice) + Val(SafeString(dgvSales.Rows(i).Cells(15).Value))).ToString()
            salesAmount = (Val(salesAmount) + Val(SafeString(dgvSales.Rows(i).Cells(16).Value))).ToString()
            actualAmount = (Val(actualAmount) + Val(SafeString(dgvSales.Rows(i).Cells(23).Value))).ToString()
            basicCostTotal = (Val(basicCostTotal) + Val(SafeString(dgvSales.Rows(i).Cells(9).Value)) * Val(SafeString(dgvSales.Rows(i).Cells(11).Value))).ToString()
            premiumCostTotal = (Val(premiumCostTotal) + Val(SafeString(dgvSales.Rows(i).Cells(9).Value)) * Val(SafeString(dgvSales.Rows(i).Cells(12).Value))).ToString()
            salesCostTotal = (Val(salesCostTotal) + Val(SafeString(dgvSales.Rows(i).Cells(9).Value)) * Val(SafeString(dgvSales.Rows(i).Cells(20).Value))).ToString()
            If Val(SafeString(dgvSales.Rows(i).Cells(17).Value)) = 0 Then
                costSurchargeTotal = (Val(costSurchargeTotal) + Val(SafeString(dgvSales.Rows(i).Cells(13).Value))).ToString()
            Else
                costSurchargeTotal = (Val(costSurchargeTotal) + Val(SafeString(dgvSales.Rows(i).Cells(17).Value)) * Val(SafeString(dgvSales.Rows(i).Cells(13).Value))).ToString()
            End If
            salesSurchargeTotal = (Val(salesSurchargeTotal) + Val(SafeString(dgvSales.Rows(i).Cells(21).Value))).ToString()
        Next
        recoveryWeightTotal = "0" : recoveryGoldTotal = "0" : recoveryOtherTotal = "0" : recoveryAmountTotal = "0"
        For i As Integer = 0 To dgvRecovery.Rows.Count - 1
            recoveryWeightTotal = (Val(recoveryWeightTotal) + Val(SafeString(dgvRecovery.Rows(i).Cells(3).Value))).ToString()
            recoveryGoldTotal = (Val(recoveryGoldTotal) + Val(SafeString(dgvRecovery.Rows(i).Cells(4).Value))).ToString()
            recoveryOtherTotal = (Val(recoveryOtherTotal) + Val(SafeString(dgvRecovery.Rows(i).Cells(7).Value))).ToString()
            recoveryAmountTotal = (Val(recoveryAmountTotal) + Val(SafeString(dgvRecovery.Rows(i).Cells(8).Value))).ToString()
        Next
        txtTax.Text = Math.Round((Val(salesAmount) - Val(recoveryAmountTotal) - Val(txtPresaleDeposit.Text)) * Val(calcTaxRate) / 100, 0).ToString()
        txtReceivable.Text = Math.Round(Val(salesAmount) + Val(txtTax.Text), 2).ToString()
        Dim infoActual As String = FormatTwoDecimals(Math.Round(Val(actualAmount) + Val(txtTax.Text) - Val(recoveryAmountTotal) - Val(txtPresaleDeposit.Text), 0).ToString())
        If salesOrderStatus <> "正常" Then
            For i As Integer = 0 To dgvPayment.Rows.Count - 1 : dgvPayment.Rows(i).Cells(2).Value = 0 : Next
            If dgvPayment.Rows.Count >= 2 Then dgvPayment.Rows(0).Cells(2).Value = Val(infoActual)
        End If
        txtReceived.Text = infoActual
        UpdateSummaryGrid()
    End Sub

    Private Sub UpdateSummaryGrid()
        dgvSummary.Columns.Clear() : dgvSummary.Rows.Clear()
        Dim widths() As Integer = {45, 100, 140, 70, 70, 70, 70, 70, 70, 70, 70, 0, 0, 0, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70}
        For i As Integer = 0 To 24
            Dim col As New DataGridViewTextBoxColumn() With {.HeaderText = "", .Name = "scol" & i, .Width = widths(i)}
            If widths(i) = 0 Then col.Visible = False
            dgvSummary.Columns.Add(col)
        Next
        Dim vals() As String = {"", "合计", "", "", "", "", "", "", "", "", weightTotal, Math.Round(Val(basicCostTotal), 0).ToString(), Math.Round(Val(premiumCostTotal), 0).ToString(), "", "", "", Math.Round(Val(salesAmount), 0).ToString(), Math.Round(Val(quantityTotal), 2).ToString(), "", "", Math.Round(Val(salesCostTotal), 0).ToString(), Math.Round(Val(salesSurchargeTotal), 0).ToString(), "", Math.Round(Val(actualAmount) - Val(recoveryAmountTotal), 0).ToString(), ""}
        dgvSummary.Rows.Add(vals)
        dgvSummary.ReadOnly = True
    End Sub

    ' ========== 销售表格结束编辑 ==========
    Private Sub DgvSales_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        tableEditState = 0
        Dim productCode As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(1).Value)
        salesProductCode = productCode
        Dim salesWeight As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(9).Value)
        Dim stockQty As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(14).Value)
        Dim salesGoldPrice As String = Math.Round(Val(SafeString(dgvSales.Rows(e.RowIndex).Cells(19).Value)), 2).ToString()
        Dim salesCost As String = Math.Round(Val(SafeString(dgvSales.Rows(e.RowIndex).Cells(20).Value)), 2).ToString()
        Dim salesSurcharge As String = Math.Round(Val(SafeString(dgvSales.Rows(e.RowIndex).Cells(21).Value)), 0).ToString()
        Dim salesQty As String = Math.Round(Val(SafeString(dgvSales.Rows(e.RowIndex).Cells(17).Value)), 2).ToString()
        Dim origSurcharge As String = Math.Round(Val(SafeString(dgvSales.Rows(e.RowIndex).Cells(18).Value)), 2).ToString()
        Dim discount As String = If(SafeString(dgvSales.Rows(e.RowIndex).Cells(21).Value) <> "", Math.Round(Val(SafeString(dgvSales.Rows(e.RowIndex).Cells(22).Value)), 4).ToString(), "1.000")
        Dim origWeight As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(25).Value)

        If e.ColumnIndex = 1 AndAlso productCode.Length > 4 Then
            Dim chkDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT * FROM xipunum_erp_shop where (poduct_code='{SafeSQLForEdit(productCode)}' or fu_code='{SafeSQLForEdit(productCode)}') order by id ASC")
            If chkDt.Rows.Count = 0 Then ShowWarning("请输入正确的商品编码！") : dgvSales.Rows(e.RowIndex).Cells(1).Value = "" : Return
            For i As Integer = 0 To dgvSales.Rows.Count - 2
                If i <> e.RowIndex AndAlso SafeString(dgvSales.Rows(i).Cells(1).Value) = productCode Then
                    ShowWarning("此商品已在当前销售清单！") : dgvSales.Rows(dgvSales.Rows.Count - 1).Cells(1).Value = "" : Return
                End If
            Next
            Dim codeDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT a.poduct_code AS apoduct_code FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE a.kufang = '{UserDepartment}' AND (a.quantity > 0 or a.jinzhong >0) AND (b.poduct_code = '{SafeSQLForEdit(productCode)}' OR b.fu_code = '{SafeSQLForEdit(productCode)}') ORDER BY a.id DESC")
            If codeDt.Rows.Count = 0 Then ShowWarning("此商品不属于您所在的库房！") : dgvSales.Rows(e.RowIndex).Cells(1).Value = "" : Return
            salesDataCode = SafeString(codeDt.Rows(0)("apoduct_code"))
            LoadProductData(e.RowIndex) : Return
        End If

        If e.ColumnIndex = 9 Then
            If Val(salesWeight) < 0 Then salesWeight = "0"
            If AutoLimitSalesRow(e.RowIndex, salesWeight, salesQty) Then ShowWarning("销售金重/数量不能小于0，且不能超过原销售加库存，已自动修正！")
            If IsWeightOnlyProduct(salesProductCode) Then salesQty = salesWeight
            salesWeight = FormatThreeDecimals(salesWeight)
            dgvSales.Rows(e.RowIndex).Cells(9).Value = salesWeight : dgvSales.Rows(e.RowIndex).Cells(17).Value = salesQty
        End If

        If e.ColumnIndex = 17 Then
            If AutoLimitSalesRow(e.RowIndex, salesWeight, salesQty) Then ShowWarning($"销售金重/数量已自动调整为原销售值加库存（数量 {salesQty}）！")
            If Val(salesQty) < 0 Then ShowWarning("销售数量不能小于0！") : salesQty = stockQty
            If Val(stockQty) <> 0 Then salesWeight = (Val(origWeight) / Val(stockQty) * Val(salesQty)).ToString()
            salesWeight = FormatThreeDecimals(salesWeight)
            Dim totalWt As String = FormatThreeDecimals((Val(SafeString(dgvSales.Rows(e.RowIndex).Cells(8).Value)) * Val(salesQty)).ToString())
            If Val(salesQty) = 0 Then salesSurcharge = origSurcharge Else salesSurcharge = (Val(origSurcharge) * Val(salesQty)).ToString()
            salesSurcharge = FormatTwoDecimals(Math.Round(Val(salesSurcharge), 0).ToString())
            dgvSales.Rows(e.RowIndex).Cells(9).Value = salesWeight : dgvSales.Rows(e.RowIndex).Cells(10).Value = totalWt
            dgvSales.Rows(e.RowIndex).Cells(17).Value = salesQty : dgvSales.Rows(e.RowIndex).Cells(21).Value = salesSurcharge
        End If

        If e.ColumnIndex = 19 Then salesGoldPrice = FormatTwoDecimals(Math.Round(Val(salesGoldPrice), 2).ToString()) : dgvSales.Rows(e.RowIndex).Cells(19).Value = salesGoldPrice
        If e.ColumnIndex = 20 Then salesCost = FormatTwoDecimals(Math.Round(Val(salesCost), 2).ToString()) : dgvSales.Rows(e.RowIndex).Cells(20).Value = salesCost

        If e.ColumnIndex = 21 Then
            salesSurcharge = FormatTwoDecimals(Math.Round(Val(salesSurcharge), 0).ToString())
            dgvSales.Rows(e.RowIndex).Cells(21).Value = salesSurcharge
            If Val(salesQty) = 0 Then discount = (Val(salesSurcharge) / Val(origSurcharge)).ToString() Else discount = (Val(salesSurcharge) / Val(salesQty) / Val(origSurcharge)).ToString()
            discount = FormatThreeDecimals(Math.Round(Val(discount), 3).ToString())
            dgvSales.Rows(e.RowIndex).Cells(22).Value = If(Val(origSurcharge) = 0, "1.000", discount)
        End If

        If e.ColumnIndex = 22 Then
            If Val(discount) > 1 Then MsgBox("折扣不能大于1") : dgvSales.Rows(e.RowIndex).Cells(22).Value = "1.000" : Return
            If Val(discount) < 0 Then MsgBox("折扣不能小于0") : dgvSales.Rows(e.RowIndex).Cells(22).Value = "1.000" : Return
            discount = FormatThreeDecimals(Math.Round(Val(discount), 3).ToString())
            If Val(salesQty) = 0 Then salesSurcharge = (Val(origSurcharge) * Val(discount)).ToString() Else salesSurcharge = (Val(origSurcharge) * Val(discount) * Val(salesQty)).ToString()
            salesSurcharge = FormatTwoDecimals(Math.Round(Val(salesSurcharge), 0).ToString())
            dgvSales.Rows(e.RowIndex).Cells(21).Value = Val(salesSurcharge).ToString()
            dgvSales.Rows(e.RowIndex).Cells(22).Value = If(Val(origSurcharge) = 0, "1.000", Val(discount).ToString())
        End If

        If e.ColumnIndex = 23 Then
            Dim actRecv As String = FormatTwoDecimals(Math.Round(Val(SafeString(dgvSales.Rows(e.RowIndex).Cells(23).Value)), 0).ToString())
            Dim minRecv As Decimal = Val(SafeString(dgvSales.Rows(e.RowIndex).Cells(16).Value)) * (100 - Val(DiscountRate)) / 100
            If Val(actRecv) < minRecv Then
                ShowWarning($"商品实收金额不能小于{Math.Round(minRecv, 0)}元")
                dgvSales.Rows(e.RowIndex).Cells(23).Value = Math.Round(Val(SafeString(dgvSales.Rows(e.RowIndex).Cells(16).Value)), 0).ToString() : Return
            End If
            dgvSales.Rows(e.RowIndex).Cells(23).Value = actRecv
        End If

        If {9, 17, 19, 20, 21, 22}.Contains(e.ColumnIndex) Then
            Dim calcUnit As String = If(Val(salesQty) = 0, (Val(salesWeight) * (Val(salesCost) + Val(salesGoldPrice)) + Val(salesSurcharge)).ToString(), ((Val(salesWeight) * (Val(salesCost) + Val(salesGoldPrice)) + Val(salesSurcharge)) / Val(salesQty)).ToString())
            calcUnit = FormatTwoDecimals(Math.Round(Val(calcUnit), 2).ToString())
            Dim calcTotal As String = (Val(salesWeight) * (Val(salesCost) + Val(salesGoldPrice)) + Val(salesSurcharge)).ToString()
            calcTotal = FormatTwoDecimals(Math.Round(Val(calcTotal), 2).ToString())
            Dim calcRecv As String = FormatTwoDecimals(Math.Round(Val(calcTotal), 0).ToString())
            dgvSales.Rows(e.RowIndex).Cells(15).Value = calcUnit
            dgvSales.Rows(e.RowIndex).Cells(16).Value = calcTotal
            dgvSales.Rows(e.RowIndex).Cells(23).Value = calcRecv
        End If

        If e.ColumnIndex = 24 Then
            Dim guideSearch As String = SafeString(dgvSales.Rows(e.RowIndex).Cells(24).Value)
            Dim gDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT name FROM xipunum_erp_user WHERE (jianxie = '{SafeSQLForEdit(guideSearch)}' or user like '%{SafeSQLForEdit(guideSearch)}%' or name like '%{SafeSQLForEdit(guideSearch)}%') and department='{UserDepartment}' LIMIT 1")
            If gDt.Rows.Count = 0 Then ShowWarning("导购信息不存在！") : dgvSales.Rows(e.RowIndex).Cells(24).Value = "" : Return
            dgvSales.Rows(e.RowIndex).Cells(24).Value = SafeString(gDt.Rows(0)("name"))
        End If
        CalculateSummary()
    End Sub

    Private Sub LoadProductData(gridRow As Integer)
        Dim sql As String = $"SELECT CASE WHEN COALESCE(d.lingxiao, '') = '' THEN '否' ELSE d.lingxiao END AS lingxiao,a.poduct_code AS poduct_code,b.product_name AS shoptitle,b.item_number AS kuanhao,COALESCE(e1.title, e2.title, '无数据') AS guige,b.caizhi as caizhi,b.quandu as quandu,CASE WHEN COALESCE(d.lingxiao, '') = '是' THEN a.jinzhong ELSE b.single END AS danzhong,a.jinzhong as jinzhong,CASE WHEN COALESCE(d.lingxiao, '') = '是' THEN a.jinzhong ELSE b.weight END AS zongzhong,a.quantity as kucun,b.sales_unit as danwie FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_type AS c ON c.id = a.kufang LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = b.item_number AND b.item_number != '' LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = b.specification_id WHERE a.kufang = '{UserDepartment}' AND (a.quantity > 0 OR a.jinzhong > 0) AND a.poduct_code = '{SafeSQLForEdit(salesDataCode)}' ORDER BY a.id DESC"
        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
        If dt.Rows.Count = 0 Then Return
        Dim d As DataRow = dt.Rows(0)
        Dim productName As String = SafeString(d("shoptitle")), stockQty As String = SafeString(d("kucun"))
        Dim unitWeight As String = SafeString(d("danzhong")), material As String = SafeString(d("caizhi"))
        Dim salesUnit As String = SafeString(d("danwie")), specName As String = SafeString(d("guige"))
        Dim itemNumber As String = SafeString(d("kuanhao")), quandu As String = SafeString(d("quandu"))
        Dim jinzhong As String = SafeString(d("jinzhong")), zongzhong As String = SafeString(d("zongzhong"))
        Dim isLingxiao As String = SafeString(d("lingxiao"))

        If isLingxiao = "是" AndAlso Val(jinzhong) <= 0 Then ShowWarning("此商品库存金重为0,不能再销售！") : dgvSales.Rows(gridRow).Cells(1).Value = "" : Return
        If isLingxiao = "否" AndAlso Val(stockQty) <= 0 Then ShowWarning("此商品库存数量为0,不能再销售！") : dgvSales.Rows(gridRow).Cells(1).Value = "" : Return

        Dim factoryCondition As String = "", basicCost As String = "", premiumCost As String = ""
        Dim salesCostVal As String = "", companySurcharge As String = "", salesSurchargeVal As String = "", salesPrice As String = ""
        Dim sDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT a.factory_condition as afactory_condition,a.basic_cost as abasic_cost,a.premium_cost as apremium_cost,a.sales_cost as asales_cost,a.company_surcharge as acompany_surcharge,a.sales_surcharge as asales_surcharge,a.sales_price as yushoujia FROM xipunum_erp_store as a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code where a.poduct_code='{SafeSQLForEdit(salesDataCode)}' LIMIT 1")
        If sDt.Rows.Count > 0 Then
            factoryCondition = SafeString(sDt.Rows(0)("afactory_condition")) : basicCost = SafeString(sDt.Rows(0)("abasic_cost"))
            premiumCost = SafeString(sDt.Rows(0)("apremium_cost")) : salesCostVal = SafeString(sDt.Rows(0)("asales_cost"))
            companySurcharge = SafeString(sDt.Rows(0)("acompany_surcharge")) : salesSurchargeVal = SafeString(sDt.Rows(0)("asales_surcharge"))
            salesPrice = SafeString(sDt.Rows(0)("yushoujia"))
        End If

        dgvSales.Rows(gridRow).Cells(1).Value = salesDataCode : dgvSales.Rows(gridRow).Cells(2).Value = productName
        dgvSales.Rows(gridRow).Cells(3).Value = itemNumber : dgvSales.Rows(gridRow).Cells(4).Value = specName
        dgvSales.Rows(gridRow).Cells(5).Value = material : dgvSales.Rows(gridRow).Cells(6).Value = quandu
        dgvSales.Rows(gridRow).Cells(7).Value = factoryCondition : dgvSales.Rows(gridRow).Cells(8).Value = unitWeight
        dgvSales.Rows(gridRow).Cells(9).Value = jinzhong : dgvSales.Rows(gridRow).Cells(10).Value = zongzhong
        dgvSales.Rows(gridRow).Cells(11).Value = basicCost : dgvSales.Rows(gridRow).Cells(12).Value = premiumCost
        dgvSales.Rows(gridRow).Cells(13).Value = companySurcharge : dgvSales.Rows(gridRow).Cells(14).Value = stockQty

        Dim calcUnit As String = "0"
        If salesUnit = "重量" Then
            If Val(stockQty) = 0 Then calcUnit = Math.Round(Val(jinzhong) * Val(salesCostVal) + Val(salesSurchargeVal), 2).ToString() Else calcUnit = Math.Round((Val(jinzhong) * Val(salesCostVal) + Val(salesSurchargeVal) * Val(stockQty)) / Val(stockQty), 2).ToString()
        Else : calcUnit = salesPrice
        End If
        calcUnit = FormatTwoDecimals(calcUnit)
        Dim calcTotal As String = If(salesUnit = "重量", Math.Round(Val(jinzhong) * Val(salesCostVal) + Val(salesSurchargeVal) * Val(stockQty), 2).ToString(), (Val(salesPrice) * Val(stockQty)).ToString())
        calcTotal = FormatTwoDecimals(calcTotal)
        Dim calcSurcharge As String = If(Val(stockQty) = 0, salesSurchargeVal, Math.Round(Val(salesSurchargeVal) * Val(stockQty), 2).ToString())
        calcSurcharge = FormatTwoDecimals(calcSurcharge)
        Dim calcRecv As String = FormatTwoDecimals(calcTotal)

        dgvSales.Rows(gridRow).Cells(15).Value = calcUnit : dgvSales.Rows(gridRow).Cells(16).Value = calcTotal
        dgvSales.Rows(gridRow).Cells(17).Value = stockQty : dgvSales.Rows(gridRow).Cells(18).Value = salesSurchargeVal
        dgvSales.Rows(gridRow).Cells(19).Value = "0.00" : dgvSales.Rows(gridRow).Cells(20).Value = salesCostVal
        dgvSales.Rows(gridRow).Cells(21).Value = calcSurcharge : dgvSales.Rows(gridRow).Cells(22).Value = "1.000"
        dgvSales.Rows(gridRow).Cells(23).Value = calcRecv
        If gridRow <= 0 Then dgvSales.Rows(gridRow).Cells(24).Value = cmbGuide.Text Else dgvSales.Rows(gridRow).Cells(24).Value = SafeString(dgvSales.Rows(gridRow - 1).Cells(24).Value)
        dgvSales.Rows(gridRow).Cells(25).Value = jinzhong
        AddBlankSalesRow() : CalculateSummary()
    End Sub

    Private Sub DgvSales_UserDeletingRow(sender As Object, e As DataGridViewRowCancelEventArgs)
        Dim delID As String = SafeString(e.Row.Cells(26).Value)
        If delID <> "" Then dgvDeleted.Rows.Add(delID, SafeString(e.Row.Cells(1).Value), salesOrderID, SafeString(e.Row.Cells(17).Value), SafeString(e.Row.Cells(9).Value))
        CalculateSummary()
    End Sub

    ' ========== 回收表格结束编辑 ==========
    Private Sub DgvRecovery_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        If dgvSales.Rows.Count < 2 Then ShowWarning("请先添加销售商品！") : Return
        recoveryEditState = 0
        Dim salesGoldPrice As String = Math.Round(Val(SafeString(dgvSales.Rows(0).Cells(19).Value)), 2).ToString()
        Dim rcName As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(1).Value)
        Dim rcTotal As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(3).Value)
        Dim rcGold As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(4).Value)
        Dim rcChengse As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(5).Value)
        Dim rcPrice As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(6).Value)
        Dim rcOther As String = SafeString(dgvRecovery.Rows(e.RowIndex).Cells(7).Value)

        If e.ColumnIndex = 1 Then
            Dim chkDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT title FROM xipunum_erp_retreat_title WHERE (bianma = '{SafeSQLForEdit(rcName)}' or title like '%{SafeSQLForEdit(rcName)}%') LIMIT 1")
            If chkDt.Rows.Count = 0 Then ShowWarning("回收名称不存在！") : dgvRecovery.Rows(e.RowIndex).Cells(1).Value = "" : Return
            dgvRecovery.Rows(e.RowIndex).Cells(1).Value = SafeString(chkDt.Rows(0)("title"))
        End If
        If e.ColumnIndex <> 1 AndAlso rcName = "" Then ShowWarning("回收名称不能为空！") : Return

        If e.ColumnIndex = 3 Then rcTotal = FormatThreeDecimals(rcTotal) : rcGold = rcTotal : dgvRecovery.Rows(e.RowIndex).Cells(3).Value = rcTotal : dgvRecovery.Rows(e.RowIndex).Cells(4).Value = rcGold
        If e.ColumnIndex = 4 Then rcGold = FormatThreeDecimals(rcGold) : dgvRecovery.Rows(e.RowIndex).Cells(4).Value = rcGold
        If e.ColumnIndex = 5 Then
            If Val(rcChengse) > 1 Then ShowWarning("成色数值不能大于1！") : dgvRecovery.Rows(e.RowIndex).Cells(5).Value = "1.0000" : Return
            rcChengse = FormatFourDecimals(rcChengse) : rcPrice = FormatTwoDecimals(Math.Round(Val(rcChengse) * Val(salesGoldPrice), 2).ToString())
            dgvRecovery.Rows(e.RowIndex).Cells(5).Value = rcChengse : dgvRecovery.Rows(e.RowIndex).Cells(6).Value = rcPrice
        End If
        If e.ColumnIndex = 6 Then
            rcChengse = Math.Round(Val(rcPrice) / Val(salesGoldPrice), 4).ToString()
            rcPrice = FormatTwoDecimals(Math.Round(Val(rcPrice), 2).ToString())
            dgvRecovery.Rows(e.RowIndex).Cells(5).Value = rcChengse : dgvRecovery.Rows(e.RowIndex).Cells(6).Value = rcPrice
        End If
        If e.ColumnIndex = 7 Then rcOther = FormatTwoDecimals(Math.Round(Val(rcOther), 2).ToString()) : dgvRecovery.Rows(e.RowIndex).Cells(7).Value = rcOther
        If e.ColumnIndex = 8 Then dgvRecovery.Rows(e.RowIndex).Cells(8).Value = FormatTwoDecimals(Math.Round(Val(SafeString(dgvRecovery.Rows(e.RowIndex).Cells(8).Value)), 2).ToString())

        If {3, 4, 5, 6, 7}.Contains(e.ColumnIndex) Then
            Dim rcAmount As String = FormatTwoDecimals(Math.Round(Val(rcGold) * Val(rcPrice) + Val(rcOther), 0).ToString())
            dgvRecovery.Rows(e.RowIndex).Cells(8).Value = rcAmount
        End If
        CalculateSummary()
    End Sub

    ' ========== 回收加减按钮 ==========
    Private Sub BtnRecoveryAdd_Click(sender As Object, e As EventArgs)
        If dgvSales.Rows.Count < 2 OrElse SafeString(dgvSales.Rows(0).Cells(1).Value) = "" Then ShowWarning("请先输入销售商品！") : Return
        If Val(SafeString(dgvSales.Rows(0).Cells(19).Value)) <= 0 Then ShowWarning("销售第一件商品克价不能为0！") : Return
        Dim addIdx As Integer = dgvRecovery.Rows.Count
        If addIdx > 0 Then
            If SafeString(dgvRecovery.Rows(addIdx - 1).Cells(8).Value) = "" Then ShowWarning("上一个回收数据不能为空！") : Return
            If SafeString(dgvRecovery.Rows(addIdx - 1).Cells(10).Value) = "" Then ShowWarning("上一个回收导购员不能为空！") : Return
        End If
        Dim idx As Integer = dgvRecovery.Rows.Add()
        dgvRecovery.Rows(idx).Cells(2).Value = "1"
        If idx = 0 Then dgvRecovery.Rows(idx).Cells(10).Value = SafeString(dgvSales.Rows(0).Cells(24).Value) Else dgvRecovery.Rows(idx).Cells(10).Value = SafeString(dgvRecovery.Rows(idx - 1).Cells(10).Value)
    End Sub

    Private Sub BtnRecoveryRemove_Click(sender As Object, e As EventArgs)
        If dgvRecovery.CurrentCell Is Nothing OrElse dgvRecovery.CurrentCell.RowIndex < 0 Then ShowWarning("请选择需要删除的回收数据！") : Return
        If dgvRecovery.CurrentCell.RowIndex >= 0 Then dgvRecovery.Rows.RemoveAt(dgvRecovery.CurrentCell.RowIndex)
        If dgvRecovery.Rows.Count = 0 Then recoveryEditState = 0
        CalculateSummary()
    End Sub

    ' ========== 预售单号 ==========
    Private Sub TxtPresaleNumber_TextChanged(sender As Object, e As EventArgs)
        If txtPresaleNumber.Text.Length > 18 Then
            Dim queryNum As String = txtPresaleNumber.Text
            Dim chkDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT * FROM xipunum_erp_presale_order where presale_umber='{queryNum}' and state='待收货'")
            If chkDt.Rows.Count = 0 Then ShowWarning("此预售订单已收货或者订单不存在！") : txtPresaleNumber.Text = "" : txtPresaleNumber.Focus() : Return
            Dim dataDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT * FROM xipunum_erp_presale_order where presale_umber='{queryNum}' and state='待收货' LIMIT 1")
            If dataDt.Rows.Count > 0 Then
                Dim presaleNum As String = SafeString(dataDt.Rows(0)("presale_umber"))
                Dim customerCode As String = SafeString(dataDt.Rows(0)("customer_code"))
                Dim deposit As String = SafeString(dataDt.Rows(0)("deposit"))
                Dim memberName As String = "", memberTel As String = ""
                Dim mDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT * FROM xipunum_erp_member where customer_code='{customerCode}' LIMIT 1")
                If mDt.Rows.Count > 0 Then memberName = SafeString(mDt.Rows(0)("name")) : memberTel = SafeString(mDt.Rows(0)("tel"))
                txtPresaleNumber.Text = "" : txtPresaleDeposit.Text = "" : txtMemberName.Text = "" : txtPhone.Text = ""
                txtPresaleNumber.Text = presaleNum : txtPresaleDeposit.Text = deposit
                txtMemberName.Text = memberName : txtPhone.Text = memberTel
                txtPresaleNumber.Enabled = True : txtPresaleNumber.BackColor = Color.Silver
            End If
        End If
    End Sub

    Private Sub TxtPresaleNumber_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If txtPresaleNumber.Text = "" Then ShowWarning("预售单号不能为空！") : txtPresaleNumber.Focus() : Return
            TxtPresaleNumber_TextChanged(Nothing, Nothing)
        End If
    End Sub

    ' ========== 打印 ==========
    Private Sub PrintDocument()
        MsgBox("打印功能：请选择打印收据后点击打印预览")
    End Sub

    ' ========== 保存编辑 ==========
    Private Sub SaveEdit()
        If saveInProgress = 1 Then Return
        LogOperationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        LogOperationAccount = UserAccount

        If dgvSales.Rows.Count <= 1 Then ShowWarning("销售数量不能为空！") : Return
        If txtReceived.Text = "" Then ShowWarning("实收金额不能为空！") : txtReceived.Focus() : Return
        If tableEditState = 1 Then ShowWarning("请先结束销售列表操作！") : Return
        If recoveryEditState = 1 Then ShowWarning("请先结束回收列表操作！") : Return
        If paymentEditState = 1 Then ShowWarning("请先结束收支列表操作！") : Return

        Dim salesCount As Integer = dgvSales.Rows.Count - 1
        For i As Integer = 0 To salesCount - 1
            If SafeString(dgvSales.Rows(i).Cells(26).Value) = "" Then
                Dim chkCode As String = SafeString(dgvSales.Rows(i).Cells(1).Value)
                Dim chkQty As String = SafeString(dgvSales.Rows(i).Cells(17).Value)
                Dim chkWeight As String = SafeString(dgvSales.Rows(i).Cells(9).Value)
                If chkCode = "" Then Continue For
                Dim chkDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT * FROM xipunum_erp_shop_kucun where poduct_code ='{SafeSQLForEdit(chkCode)}' and quantity >='{chkQty}' and jinzhong>='{chkWeight}' and kufang='{UserDepartment}'")
                If chkDt.Rows.Count = 0 Then ShowWarning($"商品编码 {chkCode} 库存不足，无法保存！") : Return
            End If
        Next

        saveInProgress = 0
        DatabaseModule.ExecuteCommand("START TRANSACTION", MySQL_Write)
        saveInProgress = 1

        Try
            SaveEditInner(salesCount)
            DatabaseModule.ExecuteCommand("COMMIT", MySQL_Write)
            saveInProgress = 0
            ShowSuccess("保存成功！")
        Catch ex As Exception
            If saveInProgress = 1 Then
                DatabaseModule.ExecuteCommand("ROLLBACK", MySQL_Write)
                saveInProgress = 0
            End If
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub

    Private Sub SaveEditInner(salesCount As Integer)
        Dim orderNumber As String = txtOrderNumber.Text
        Dim memberName As String = txtMemberName.Text
        Dim memberTel As String = txtPhone.Text
        Dim taxPoint As String = txtTaxRate.Text
        Dim taxAmount As String = txtTax.Text
        Dim receivableAmount As String = txtReceivable.Text
        Dim receivedAmount As String = txtReceived.Text
        Dim salesRemarks As String = txtRemarks.Text
        Dim discountAmount As String = FormatTwoDecimals(Math.Round(Val(actualAmount) - Val(recoveryAmountTotal) - Val(receivedAmount), 2).ToString())

        ' 会员处理
        Dim customerCode As String = ""
        If memberName <> "" Then
            Dim memberID As String = ""
            If memberName.Contains("(") AndAlso memberName.Contains(")") Then
                Dim s As Integer = memberName.IndexOf("(") + 1, e2 As Integer = memberName.IndexOf(")")
                If s < e2 Then memberID = memberName.Substring(s, e2 - s)
            End If
            Dim mExistDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT * FROM xipunum_erp_member where memberid= '{memberID}'")
            If mExistDt.Rows.Count = 0 Then
                customerCode = "HY" & DateTime.Now.ToString("yyyyMMddHHmmss")
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_member (customer_code, name, tel, cjuser, creationtime) VALUES ('{customerCode}', '{SafeSQLForEdit(memberName)}', '{SafeSQLForEdit(memberTel)}', '{SafeSQL(LogOperationAccount)}', '{LogOperationDate}')", MySQL_Write)
            End If
            Dim cDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT customer_code FROM xipunum_erp_member where tel='{SafeSQLForEdit(memberTel)}' order by id ASC LIMIT 1")
            If cDt.Rows.Count > 0 Then customerCode = SafeString(cDt.Rows(0)("customer_code"))
        End If

        Dim presaleNum As String = If(radioPresaleYes.Checked, txtPresaleNumber.Text, "")
        Dim counterWeight As String = weightTotal, netWeight As String = weightTotal
        Dim basicCostVal As String = basicCostTotal, premiumCostVal As String = premiumCostTotal
        Dim salesCostVal As String = salesCostTotal, salesSurchargeVal As String = salesSurchargeTotal

        ' 处理删除商品
        For i As Integer = 0 To dgvDeleted.Rows.Count - 1
            Dim delID As String = SafeString(dgvDeleted.Rows(i).Cells(0).Value)
            Dim delCode As String = SafeString(dgvDeleted.Rows(i).Cells(1).Value)
            Dim delQty As String = SafeString(dgvDeleted.Rows(i).Cells(3).Value)
            Dim delWeight As String = SafeString(dgvDeleted.Rows(i).Cells(4).Value)
            DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_shop_kucun SET quantity = quantity + '{delQty}',jinzhong = jinzhong + '{delWeight}' WHERE poduct_code ='{SafeSQLForEdit(delCode)}' and kufang ='{UserDepartment}'", MySQL_Write)
            DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_shop SET state= '销售',cjuser= '{LogOperationAccount}',updatetime= '{LogOperationDate}'  WHERE poduct_code ='{SafeSQLForEdit(delCode)}' LIMIT 1", MySQL_Write)
            DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_shop_log (poduct_code, type, creationtime) VALUES ('{SafeSQLForEdit(delCode)}', '撤单', '{LogOperationDate}')", MySQL_Write)
            DatabaseModule.ExecuteCommand($"delete from xipunum_erp_outbound where id= '{delID}'", MySQL_Write)
            DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_zhengshu SET xiaoshouid= '',xiaoshou= '',updatetime= '{LogOperationDate}'  WHERE poduct_code ='{SafeSQLForEdit(delCode)}' LIMIT 1", MySQL_Write)
            AddSystemLog("修改", "修改商品信息", $"账户:{UserAccount} 修改商品：{delCode} 状态为销售")
        Next

        ' 回收单号
        Dim recoveryNum As String = ""
        If dgvRecovery.Rows.Count > 0 Then
            recoveryNum = If(salesRecoveryNumber = "", (DateTime.Now.ToString("yyyyMMddHHmmss") & New Random().Next(1000, 9999).ToString()).Substring(0, 18) & "3", salesRecoveryNumber)
        End If

        ' 介绍人
        Dim introducerSQL As String = ""
        If cmbIntroducer.SelectedIndex >= 0 Then
            Dim introText As String = cmbIntroducer.Text
            If introText.Length >= 8 Then introducerSQL = $"shopping_guide='{introText.Substring(introText.Length - 8, 7)}',"
        End If

        Dim salesFactory As String = cmbSalesFactory.Text
        Dim invoiceFlag As String = If(radioInvoice.Checked, "1", "0")

        ' 更新出库主单
        DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_outbound_order SET settlement_number='{SafeSQLForEdit(orderNumber)}',retrea_umber='{SafeSQLForEdit(recoveryNum)}',ying_amount='{SafeSQLForEdit(receivableAmount)}',youhui='{SafeSQLForEdit(discountAmount)}',settlement='{SafeSQLForEdit(receivedAmount)}',counter_weight='{SafeSQLForEdit(counterWeight)}',settlement_weight='{SafeSQLForEdit(counterWeight)}',net_counter='{SafeSQLForEdit(netWeight)}',net_weight='{SafeSQLForEdit(netWeight)}',basic_cost='{SafeSQLForEdit(basicCostVal)}',premium_cost='{SafeSQLForEdit(premiumCostVal)}',sales_cost='{SafeSQLForEdit(salesCostVal)}',sales_surcharge='{SafeSQLForEdit(salesSurchargeVal)}',{introducerSQL}taxpoint='{SafeSQLForEdit(taxPoint)}',taxamount='{SafeSQLForEdit(taxAmount)}',remarks='{SafeSQLForEdit(salesRemarks)}',state='正常',fapiao='{invoiceFlag}',xsfactory='{SafeSQLForEdit(salesFactory)}',cjuser='{SafeSQL(LogOperationAccount)}',updatetime= '{LogOperationDate}'  WHERE id ='{salesOrderID}' LIMIT 1", MySQL_Write)

        ' 收款记录
        Dim payCount As Integer = dgvPayment.Rows.Count - 1
        For i As Integer = 0 To payCount - 1
            Dim payID As String = SafeString(dgvPayment.Rows(i).Cells(3).Value)
            Dim payAmount As String = SafeString(dgvPayment.Rows(i).Cells(2).Value)
            Dim payExistDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT * FROM xipunum_erp_shoukian where settlement_number= '{orderNumber}' and type= '{payID}'")
            If payExistDt.Rows.Count = 0 Then
                If Val(payAmount) <> 0 Then DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_shoukuan (leibie, settlement_number, xianjin, type, kufang, cjuser, creationtime) VALUES ('1', '{SafeSQLForEdit(orderNumber)}', '{payAmount}', '{payID}', '{UserDepartment}', '{SafeSQL(LogOperationAccount)}', '{LogOperationDate}')", MySQL_Write)
            Else
                DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_shoukuan SET xianjin='{payAmount}',cjuser='{SafeSQL(LogOperationAccount)}',updatetime= '{LogOperationDate}'  WHERE settlement_number ='{orderNumber}' and type ='{payID}' LIMIT 1", MySQL_Write)
            End If
        Next

        AddSystemLog("修改", "商品销售修改", $"账户:{UserAccount} 商品销售修改，销售单号:{orderNumber}")

        ' 查询销售单批零和时间
        Dim orderPling As String = "", orderSalesTime As String = ""
        Dim plingDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT pling,creationtime FROM xipunum_erp_outbound_order WHERE id='{salesOrderID}'")
        If plingDt.Rows.Count > 0 Then orderPling = SafeString(plingDt.Rows(0)("pling")) : orderSalesTime = SafeString(plingDt.Rows(0)("creationtime"))

        Dim warehouseName As String = ""
        Dim whDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT title FROM xipunum_erp_type where id='{UserDepartment}' order by id ASC LIMIT 1")
        If whDt.Rows.Count > 0 Then warehouseName = SafeString(whDt.Rows(0)("title"))

        ' 逐行处理销售明细
        Dim totalJifen As String = "0"
        For i As Integer = 0 To salesCount - 1
            Dim infoCode As String = SafeString(dgvSales.Rows(i).Cells(1).Value)
            If infoCode = "" Then Continue For
            Dim infoQty As String = SafeString(dgvSales.Rows(i).Cells(17).Value)
            Dim infoUnitPrice As String = SafeString(dgvSales.Rows(i).Cells(15).Value)
            Dim infoTotalPrice As String = SafeString(dgvSales.Rows(i).Cells(16).Value)
            Dim infoActualReceived As String = SafeString(dgvSales.Rows(i).Cells(23).Value)
            Dim infoGoldPrice As String = SafeString(dgvSales.Rows(i).Cells(19).Value)
            Dim infoNetWeight As String = SafeString(dgvSales.Rows(i).Cells(9).Value)
            Dim infoBasicCost As String = SafeString(dgvSales.Rows(i).Cells(11).Value)
            Dim infoPremiumCost As String = SafeString(dgvSales.Rows(i).Cells(12).Value)
            Dim infoSalesCost As String = SafeString(dgvSales.Rows(i).Cells(20).Value)
            Dim infoSalesSurcharge As String = SafeString(dgvSales.Rows(i).Cells(21).Value)
            Dim infoDiscount As String = SafeString(dgvSales.Rows(i).Cells(22).Value)
            Dim infoGuide As String = SafeString(dgvSales.Rows(i).Cells(24).Value)
            Dim infoDetailID As String = SafeString(dgvSales.Rows(i).Cells(26).Value)

            Dim origQty As String = "0", origWeight As String = "0", origWarehouse As String = UserDepartment
            If infoDetailID <> "" Then
                Dim origDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT quantity,net_weight,kufang FROM xipunum_erp_outbound WHERE id='{infoDetailID}'")
                If origDt.Rows.Count > 0 Then
                    origQty = SafeString(origDt.Rows(0)("quantity")) : origWeight = SafeString(origDt.Rows(0)("net_weight"))
                    origWarehouse = SafeString(origDt.Rows(0)("kufang")) : If origWarehouse = "" Then origWarehouse = UserDepartment
                End If
                If Not AdjustStock(infoCode, origQty, origWeight, infoQty, infoNetWeight, origWarehouse) Then
                    If Val(infoQty) >= Val(origQty) OrElse Val(infoNetWeight) >= Val(origWeight) Then
                        If AutoLimitSalesRow(i, infoNetWeight, infoQty) Then
                            If Not AdjustStock(infoCode, origQty, origWeight, infoQty, infoNetWeight, origWarehouse) Then
                                SaveFailedRollback($"商品编码 {infoCode} 库存不足，无法加大出库！") : Throw New Exception("rollback")
                            End If
                        Else
                            SaveFailedRollback($"商品编码 {infoCode} 库存不足，无法加大出库！") : Throw New Exception("rollback")
                        End If
                    Else
                        SaveFailedRollback($"商品编码 {infoCode} 库房库存记录不存在，无法回写库存！") : Throw New Exception("rollback")
                    End If
                End If
            End If

            Dim guideAccount As String = ""
            Dim gDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT user FROM xipunum_erp_user WHERE name = '{SafeSQLForEdit(infoGuide)}' and department = '{UserDepartment}' LIMIT 1")
            If gDt.Rows.Count > 0 Then guideAccount = SafeString(gDt.Rows(0)("user"))

            If infoDetailID <> "" Then
                DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_outbound SET quantity ='{infoQty}',xiaodan_amount ='{infoUnitPrice}',xiao_amount ='{infoTotalPrice}',settlement ='{infoActualReceived}',gold_price ='{infoGoldPrice}',net_weight ='{infoNetWeight}',basic_cost ='{infoBasicCost}',premium_cost ='{infoPremiumCost}',sales_cost ='{infoSalesCost}',sales_surcharge ='{infoSalesSurcharge}',shopping_guide ='{guideAccount}',zhekou ='{infoDiscount}',updatetime ='{LogOperationDate}' WHERE poduct_code = '{SafeSQLForEdit(infoCode)}' and order_id='{salesOrderID}'", MySQL_Write)
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_history (poduct_code, updatetime, number, type, quantity, jinzhong, zhongliang, conter, cjuser) VALUES ('{SafeSQLForEdit(infoCode)}', '{LogOperationDate}', '{SafeSQLForEdit(orderNumber)}', '成品出库修改', '{infoQty}', '{infoNetWeight}', '{SafeString(dgvSales.Rows(i).Cells(10).Value)}', '商品从{warehouseName}出库修改', '{SafeSQL(LogOperationAccount)}')", MySQL_Write)
                AddSystemLog("修改", "商品销售修改", $"账户:{UserAccount} 商品销售修改，编码：{infoCode}")
            Else
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_outbound (order_id, poduct_code, quantity, xiaodan_amount, xiao_amount, settlement, gold_price, net_weight, basic_cost, premium_cost, sales_cost, sales_surcharge, shopping_guide, kufang, zhekou, remarks, pling, state, sales_return, cjuser, creationtime, xianshangtime, updatetime) VALUES ('{salesOrderID}', '{SafeSQLForEdit(infoCode)}', '{infoQty}', '{infoUnitPrice}', '{infoTotalPrice}', '{infoActualReceived}', '{infoGoldPrice}', '{infoNetWeight}', '{infoBasicCost}', '{infoPremiumCost}', '{infoSalesCost}', '{infoSalesSurcharge}', '{guideAccount}', '{UserDepartment}', '{infoDiscount}', '', '{orderPling}', '0', '0', '{SafeSQL(LogOperationAccount)}', '{orderSalesTime}', '{orderSalesTime}', '{LogOperationDate}')", MySQL_Write)
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_history (poduct_code, updatetime, number, type, quantity, jinzhong, zhongliang, conter, cjuser) VALUES ('{SafeSQLForEdit(infoCode)}', '{LogOperationDate}', '{SafeSQLForEdit(orderNumber)}', '成品出库', '{infoQty}', '{infoNetWeight}', '{SafeString(dgvSales.Rows(i).Cells(10).Value)}', '商品从{warehouseName}出库', '{SafeSQL(LogOperationAccount)}')", MySQL_Write)

                ' 证书处理
                Dim certDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT * FROM xipunum_erp_zhengshu WHERE poduct_code = '{SafeSQLForEdit(infoCode)}' LIMIT 1")
                If certDt.Rows.Count > 0 Then
                    Dim certID As String = SafeString(certDt.Rows(0)("id"))
                    If certID <> "" Then
                        Dim certCost As String = SafeString(certDt.Rows(0)("chengben"))
                        Dim certSale As String = SafeString(certDt.Rows(0)("xiaoshou"))
                        Dim certSaleVal As String = If(certSale = "", $",xiaoshou= '{certCost}'", "")
                        DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_zhengshu SET xiaoshouid= '{SafeSQLForEdit(orderNumber)}'{certSaleVal},updatetime= '{LogOperationDate}'  WHERE poduct_code ='{SafeSQLForEdit(infoCode)}' LIMIT 1", MySQL_Write)
                    End If
                End If
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_shop_log (poduct_code, type, creationtime) VALUES ('{SafeSQLForEdit(infoCode)}', '出库', '{LogOperationDate}')", MySQL_Write)
                AddSystemLog("添加", "商品销售出库", $"账户:{UserAccount} 商品销售出库，编码：{infoCode}")
                If Not DeductStock(infoCode, infoQty, infoNetWeight, UserDepartment) Then SaveFailedRollback($"商品编码 {infoCode} 库存不足或已被占用，保存已回滚！") : Throw New Exception("rollback")
            End If

            ' 积分
            Dim jifenDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT CASE WHEN COALESCE(d.jifen, '' ) = '' THEN '0' ELSE d.jifen END AS jifen FROM xipunum_erp_shop AS a LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = a.item_number AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = i.specification_id LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id LEFT JOIN xipunum_erp_category AS d ON d.id = COALESCE ( e1.category_id, e2.category_id ) WHERE a.poduct_code='{SafeSQLForEdit(infoCode)}' LIMIT 1")
            Dim jifenRate As String = If(jifenDt.Rows.Count > 0, SafeString(jifenDt.Rows(0)("jifen")), "0")
            totalJifen = (Val(totalJifen) + Math.Round(Val(infoActualReceived) * Val(jifenRate) / 100, 0)).ToString()

            ' 检查售尽
            Dim sDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT sum(quantity) as shuliang,sum(jinzhong) as jinzhong FROM xipunum_erp_shop_kucun WHERE poduct_code='{SafeSQLForEdit(infoCode)}'")
            If sDt.Rows.Count > 0 Then
                If Val(SafeString(sDt.Rows(0)("jinzhong"))) <= 0 AndAlso Val(SafeString(sDt.Rows(0)("shuliang"))) <= 0 Then
                    DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_shop SET state= '售尽',updatetime= '{LogOperationDate}'  WHERE poduct_code ='{SafeSQLForEdit(infoCode)}' LIMIT 1", MySQL_Write)
                End If
            End If
        Next

        ' 预售订单状态
        If radioPresaleYes.Checked Then DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_presale_order SET state= '已收货',updatetime= '{LogOperationDate}'  WHERE presale_umber ='{SafeSQLForEdit(presaleNum)}' LIMIT 1", MySQL_Write)

        ' 无回收时删除回收记录
        If dgvRecovery.Rows.Count = 0 AndAlso salesRecoveryID <> "" Then
            DatabaseModule.ExecuteCommand($"delete from xipunum_erp_retreat where order_id= '{salesRecoveryID}'", MySQL_Write)
            DatabaseModule.ExecuteCommand($"delete from xipunum_erp_retreat_order where id= '{salesRecoveryID}'", MySQL_Write)
        End If

        ' 查询销售单信息
        Dim orderAccount As String = "", orderWarehouse As String = "", orderTime As String = "", orderMemberCode As String = ""
        Dim oiDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT a.kufang as kufang,b.cjuser AS cjuser,b.creationtime AS creationtime,b.customer_code AS customer_code,b.pling as pling FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_outbound_order AS b ON b.id = a.order_id WHERE b.id = '{salesOrderID}' LIMIT 1")
        If oiDt.Rows.Count > 0 Then
            orderAccount = SafeString(oiDt.Rows(0)("cjuser")) : orderWarehouse = SafeString(oiDt.Rows(0)("kufang"))
            orderTime = SafeString(oiDt.Rows(0)("creationtime")) : orderMemberCode = SafeString(oiDt.Rows(0)("customer_code"))
        End If

        ' 批发模式会员存欠
        If orderPling = "批发" Then
            If Val(recoveryGoldTotal) > 0 Then
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_member_cq (customer_code, danhao, cunqu, type, number, kufang, cjuser, creationtime) VALUES ('{orderMemberCode}', '{SafeSQLForEdit(orderNumber)}', '欠', '料', '{recoveryGoldTotal}', '{orderWarehouse}', '{orderAccount}', '{orderTime}')", MySQL_Write)
                AddSystemLog("新增", "会员出料", $"账户:{UserAccount} 增加会员:{orderMemberCode}欠料（{recoveryGoldTotal}）g 料")
            End If
            If Val(receivedAmount) > 0 Then
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_member_cq (customer_code, danhao, cunqu, type, number, kufang, cjuser, creationtime) VALUES ('{orderMemberCode}', '{SafeSQLForEdit(orderNumber)}', '欠', '元', '{receivedAmount}', '{orderWarehouse}', '{orderAccount}', '{orderTime}')", MySQL_Write)
                AddSystemLog("新增", "会员消费", $"账户:{UserAccount} 增加会员:{orderMemberCode}欠款（{receivedAmount}）元")
            End If
        End If

        ' 回收数据处理
        If dgvRecovery.Rows.Count > 0 Then
            If salesRecoveryNumber = "" Then
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_retreat_order (retrea_umber, customer_code, salesman, cjuser, creationtime) VALUES ('{SafeSQLForEdit(recoveryNum)}', '{SafeSQLForEdit(salesMemberCode)}', '{SafeSQLForEdit(cmbGuide.Text)}', '{SafeSQL(LogOperationAccount)}', '{orderTime}')", MySQL_Write)
                Dim tDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT id FROM xipunum_erp_retreat_order where retrea_umber='{SafeSQLForEdit(recoveryNum)}' order by id ASC LIMIT 1")
                If tDt.Rows.Count > 0 Then salesRecoveryID = SafeString(tDt.Rows(0)("id")) : salesRecoveryNumber = recoveryNum
            End If
            DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_retreat_order SET total='{recoveryWeightTotal}',jin_zhong='{recoveryGoldTotal}',qita_price='{recoveryOtherTotal}',retreat_amount='{recoveryAmountTotal}',ying_amount='{recoveryAmountTotal}',settlement='{recoveryAmountTotal}',remarks='',cjuser='{SafeSQL(LogOperationAccount)}',updatetime= '{LogOperationDate}'  WHERE id ='{salesRecoveryID}' LIMIT 1", MySQL_Write)
            DatabaseModule.ExecuteCommand($"delete from xipunum_erp_retreat where order_id= '{salesRecoveryID}'", MySQL_Write)
            Dim retreatTime As String = ""
            Dim rtDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT creationtime FROM xipunum_erp_retreat_order where id='{salesRecoveryID}' order by id ASC LIMIT 1")
            If rtDt.Rows.Count > 0 Then retreatTime = SafeString(rtDt.Rows(0)("creationtime"))

            For i As Integer = 0 To dgvRecovery.Rows.Count - 1
                Dim rcName As String = SafeString(dgvRecovery.Rows(i).Cells(1).Value)
                Dim rcQty As String = SafeString(dgvRecovery.Rows(i).Cells(2).Value)
                Dim rcTotal As String = SafeString(dgvRecovery.Rows(i).Cells(3).Value)
                Dim rcGold As String = SafeString(dgvRecovery.Rows(i).Cells(4).Value)
                Dim rcChengse As String = SafeString(dgvRecovery.Rows(i).Cells(5).Value)
                Dim rcPrice As String = SafeString(dgvRecovery.Rows(i).Cells(6).Value)
                Dim rcOther As String = SafeString(dgvRecovery.Rows(i).Cells(7).Value)
                Dim rcAmount As String = SafeString(dgvRecovery.Rows(i).Cells(8).Value)
                Dim rcRemarks As String = SafeString(dgvRecovery.Rows(i).Cells(9).Value)
                Dim rcGuide As String = SafeString(dgvRecovery.Rows(i).Cells(10).Value)
                Dim rcAccount As String = ""
                Dim gDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT user FROM xipunum_erp_user WHERE name = '{SafeSQLForEdit(rcGuide)}' and department = '{UserDepartment}' LIMIT 1")
                If gDt.Rows.Count > 0 Then rcAccount = SafeString(gDt.Rows(0)("user"))
                Dim rcDataID As String = "1"
                Dim rcJifen As String = "0"
                Dim rcDt As DataTable = DatabaseModule.ExecuteQuery($"SELECT a.id as aid,CASE WHEN COALESCE(b.jifen, '' ) = '' THEN '0' ELSE b.jifen END AS jifen FROM xipunum_erp_retreat_title as a INNER JOIN xipunum_erp_category AS b ON b.id = a.category_id WHERE a.title = '{SafeSQLForEdit(rcName)}'  LIMIT 1")
                If rcDt.Rows.Count > 0 Then rcDataID = SafeString(rcDt.Rows(0)("aid")) : rcJifen = SafeString(rcDt.Rows(0)("jifen"))
                DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_retreat (order_id, product_name, quantity, total, jin_zhong, chengse, price, qita_price, retreat_amount, huishoutime, shopping_guide, remarks, cjuser, creationtime) VALUES ('{salesRecoveryID}', '{rcDataID}', '{rcQty}', '{rcTotal}', '{rcGold}', '{rcChengse}', '{rcPrice}', '{rcOther}', '{rcAmount}', '{retreatTime}', '{rcAccount}', '{rcRemarks}', '{SafeSQL(LogOperationAccount)}', '{retreatTime}')", MySQL_Write)
                totalJifen = (Val(totalJifen) - Math.Round(Val(rcAmount) * Val(rcJifen) / 100, 0)).ToString()
                AddSystemLog("添加", "商品回收", $"账户:{UserAccount} 商品回收，名称：{rcName}")
            Next
            AddSystemLog("修改", "商品回收", $"账户:{UserAccount} 修改回收单号：{salesRecoveryNumber}原始相关数据")
        End If

        ' 积分记录
        If orderMemberCode <> "" AndAlso Val(totalJifen) <> 0 Then
            DatabaseModule.ExecuteCommand($"INSERT INTO xipunum_erp_member_score_log (customer_code, settlement_number, num, st, type, remarks, creationtime, cjuser) VALUES ('{orderMemberCode}', '{SafeSQLForEdit(orderNumber)}', '{totalJifen}', '0', '0', '【订单号：{SafeSQLForEdit(orderNumber)}】实销 获得积分', '{orderTime}', '{SafeSQL(LogOperationAccount)}')", MySQL_Write)
        End If
    End Sub

    ' ========== 辅助函数 ==========
    Private Function FormatThreeDecimals(s As String) As String
        If s = "" Then Return "0.000"
        Return Math.Round(Val(s), 3).ToString("0.000")
    End Function

    Private Function FormatTwoDecimals(s As String) As String
        If s = "" Then Return "0.00"
        Return Math.Round(Val(s), 2).ToString("0.00")
    End Function

    Private Function FormatFourDecimals(s As String) As String
        If s = "" Then Return "1.0000"
        Return Math.Round(Val(s), 4).ToString("0.0000")
    End Function

    Private Function SafeString(o As Object) As String
        If o Is Nothing OrElse IsDBNull(o) Then Return ""
        Return o.ToString()
    End Function

    Private Function SafeDecimal(o As Object) As Decimal
        If o Is Nothing OrElse IsDBNull(o) Then Return 0
        Dim v As Decimal
        If Decimal.TryParse(o.ToString(), v) Then Return v
        Return 0
    End Function

    Private Sub ShowWarning(msg As String)
        MsgBox(msg, MsgBoxStyle.Exclamation, "提示")
    End Sub

    Private Sub ShowSuccess(msg As String)
        MsgBox(msg, MsgBoxStyle.Information, "成功")
    End Sub

    Private Sub ShowError(msg As String)
        MsgBox(msg, MsgBoxStyle.Critical, "错误")
    End Sub

    Private Sub AddSystemLog(logType As String, logTitle As String, logContent As String)
        Dim sql As String = $"INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('{SafeSQLForEdit(logType)}', '{SafeSQLForEdit(logTitle)}', '{SafeSQLForEdit(logContent)}', '{SafeSQL(LogOperationAccount)}', '{LogOperationDate}')"
        DatabaseModule.ExecuteCommand(sql, MySQL_Write)
    End Sub

End Class

' ============================================================================
' 结账结料添加/编辑窗口
' 功能: 工厂结账结料的添加和编辑
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class FactorySettlementAddForm
    Inherits System.Windows.Forms.Form

    Private editMode As Boolean = False
    Private editId As String = ""

    Private cmbFactory As New ComboBox()
    Private cmbType As New ComboBox()
    Private txtOrderNumber As New TextBox()
    Private txtTotalQuantity As New TextBox()
    Private txtTotalWeight As New TextBox()
    Private txtTotalAmount As New TextBox()
    Private txtRemarks As New TextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnCancel As New Button()

    Public Sub New(Optional id As String = "")
        editMode = Not String.IsNullOrEmpty(id)
        editId = id
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = If(editMode, "编辑结账单", "添加结账单")
        Me.Size = New Drawing.Size(500, 400)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        Dim y As Integer = 20
        Dim labelWidth As Integer = 80
        Dim controlX As Integer = 100

        ' 工厂
        AddLabel("工厂：", 20, y)
        cmbFactory.Location = New Drawing.Point(controlX, y)
        cmbFactory.Size = New Drawing.Size(200, 25)
        Me.Controls.Add(cmbFactory)
        y += 35

        ' 类型
        AddLabel("类型：", 20, y)
        cmbType.Location = New Drawing.Point(controlX, y)
        cmbType.Size = New Drawing.Size(200, 25)
        cmbType.Items.AddRange({"结账", "结料", "结账结料"})
        cmbType.SelectedIndex = 0
        Me.Controls.Add(cmbType)
        y += 35

        ' 单据号
        AddLabel("单据号：", 20, y)
        txtOrderNumber.Location = New Drawing.Point(controlX, y)
        txtOrderNumber.Size = New Drawing.Size(200, 25)
        txtOrderNumber.ReadOnly = True
        Me.Controls.Add(txtOrderNumber)
        y += 35

        ' 总数量
        AddLabel("总数量：", 20, y)
        txtTotalQuantity.Location = New Drawing.Point(controlX, y)
        txtTotalQuantity.Size = New Drawing.Size(200, 25)
        Me.Controls.Add(txtTotalQuantity)
        y += 35

        ' 总金重
        AddLabel("总金重：", 20, y)
        txtTotalWeight.Location = New Drawing.Point(controlX, y)
        txtTotalWeight.Size = New Drawing.Size(200, 25)
        Me.Controls.Add(txtTotalWeight)
        y += 35

        ' 总金额
        AddLabel("总金额：", 20, y)
        txtTotalAmount.Location = New Drawing.Point(controlX, y)
        txtTotalAmount.Size = New Drawing.Size(200, 25)
        Me.Controls.Add(txtTotalAmount)
        y += 35

        ' 备注
        AddLabel("备注：", 20, y)
        txtRemarks.Location = New Drawing.Point(controlX, y)
        txtRemarks.Size = New Drawing.Size(300, 25)
        Me.Controls.Add(txtRemarks)
        y += 45

        ' 按钮
        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(150, y)
        btnSave.Size = New Drawing.Size(80, 30)
        Me.Controls.Add(btnSave)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(250, y)
        btnCancel.Size = New Drawing.Size(80, 30)
        btnCancel.DialogResult = DialogResult.Cancel
        Me.Controls.Add(btnCancel)

        Me.CancelButton = btnCancel
    End Sub

    Private Sub AddLabel(text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        Me.Controls.Add(lbl)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadFactoryList()
        If editMode Then
            LoadData()
        Else
            txtOrderNumber.Text = GenerateOrderNumber("JZ")
        End If
    End Sub

    Private Sub LoadFactoryList()
        Try
            Dim sql As String = "SELECT id, title FROM xipunum_erp_about WHERE 1=1 ORDER BY id DESC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbFactory.Items.Clear()
            For Each row As DataRow In dt.Rows
                cmbFactory.Items.Add(GBKToUTF8(SafeString(row("title"))))
            Next
            If cmbFactory.Items.Count > 0 Then cmbFactory.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim sql As String = $"SELECT * FROM xipunum_erp_delivery_order WHERE id='{editId}' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                txtOrderNumber.Text = SafeString(row("delivery_umber"))
                txtTotalQuantity.Text = SafeDecimal(row("total_quantity")).ToString()
                txtTotalWeight.Text = SafeDecimal(row("total_weight")).ToString()
                txtTotalAmount.Text = SafeDecimal(row("total_amount")).ToString()
                txtRemarks.Text = GBKToUTF8(SafeString(row("remarks")))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If cmbFactory.SelectedIndex < 0 Then
            ShowWarning("请选择工厂！")
            Return
        End If

        Try
            Dim factoryName As String = cmbFactory.SelectedItem.ToString()
            Dim factoryId As String = ""
            Dim sql As String = $"SELECT id FROM xipunum_erp_about WHERE title='{SafeSQL(factoryName)}' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            If dt.Rows.Count > 0 Then
                factoryId = SafeString(dt.Rows(0)("id"))
            End If

            Dim orderType As String = cmbType.SelectedItem.ToString()
            Dim totalQty As Decimal = SafeDecimal(txtTotalQuantity.Text)
            Dim totalWeight As Decimal = SafeDecimal(txtTotalWeight.Text)
            Dim totalAmount As Decimal = SafeDecimal(txtTotalAmount.Text)

            If editMode Then
                sql = $"UPDATE xipunum_erp_delivery_order SET gongchang='{factoryId}', type='{SafeSQL(orderType)}', total_quantity={totalQty}, total_weight={totalWeight}, total_amount={totalAmount}, remarks='{SafeSQL(txtRemarks.Text)}' WHERE id='{editId}'"
                DatabaseModule.ExecuteCommand(sql)
                AddSystemLog("编辑", "编辑结账单", $"单号:{txtOrderNumber.Text}")
            Else
                sql = $"INSERT INTO xipunum_erp_delivery_order (delivery_umber, gongchang, type, total_quantity, total_weight, total_amount, state, remarks, cjuser, creationtime) VALUES ('{SafeSQL(txtOrderNumber.Text)}', '{factoryId}', '{SafeSQL(orderType)}', {totalQty}, {totalWeight}, {totalAmount}, '未结算', '{SafeSQL(txtRemarks.Text)}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
                DatabaseModule.ExecuteCommand(sql)
                AddSystemLog("添加", "添加结账单", $"单号:{txtOrderNumber.Text}")
            End If

            ShowSuccess("保存成功！")
            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub
End Class

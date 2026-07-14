' ============================================================================
' 收支管理添加窗口
' 功能: 添加/编辑收支记录
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class FinanceAddForm
    Inherits System.Windows.Forms.Form

    Private cmbType As New ComboBox()
    Private cmbTitle As New ComboBox()
    Private cmbPay As New ComboBox()
    Private cmbWarehouse As New ComboBox()
    Private txtAmount As New TextBox()
    Private txtRemarks As New TextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnCancel As New Button()
    Private editMode As Boolean = False
    Private editId As String = ""

    Public Sub New(Optional id As String = "")
        editId = id
        editMode = Not String.IsNullOrEmpty(id)
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = If(editMode, "编辑收支", "添加收支")
        Me.Size = New Drawing.Size(450, 350)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblType As New Label()
        lblType.Text = "类型："
        lblType.Location = New Drawing.Point(30, 30)
        lblType.AutoSize = True
        Me.Controls.Add(lblType)

        cmbType.Location = New Drawing.Point(100, 27)
        cmbType.Size = New Drawing.Size(150, 25)
        cmbType.Items.AddRange(New String() {"收入", "支出"})
        cmbType.SelectedIndex = 0
        Me.Controls.Add(cmbType)

        Dim lblTitle As New Label()
        lblTitle.Text = "名称："
        lblTitle.Location = New Drawing.Point(30, 70)
        lblTitle.AutoSize = True
        Me.Controls.Add(lblTitle)

        cmbTitle.Location = New Drawing.Point(100, 67)
        cmbTitle.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbTitle)

        Dim lblPay As New Label()
        lblPay.Text = "支付方式："
        lblPay.Location = New Drawing.Point(30, 110)
        lblPay.AutoSize = True
        Me.Controls.Add(lblPay)

        cmbPay.Location = New Drawing.Point(100, 107)
        cmbPay.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbPay)

        Dim lblWarehouse As New Label()
        lblWarehouse.Text = "店铺："
        lblWarehouse.Location = New Drawing.Point(30, 150)
        lblWarehouse.AutoSize = True
        Me.Controls.Add(lblWarehouse)

        cmbWarehouse.Location = New Drawing.Point(100, 147)
        cmbWarehouse.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbWarehouse)

        Dim lblAmount As New Label()
        lblAmount.Text = "金额："
        lblAmount.Location = New Drawing.Point(30, 190)
        lblAmount.AutoSize = True
        Me.Controls.Add(lblAmount)

        txtAmount.Location = New Drawing.Point(100, 187)
        txtAmount.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(txtAmount)

        Dim lblRemarks As New Label()
        lblRemarks.Text = "备注："
        lblRemarks.Location = New Drawing.Point(30, 230)
        lblRemarks.AutoSize = True
        Me.Controls.Add(lblRemarks)

        txtRemarks.Location = New Drawing.Point(100, 227)
        txtRemarks.Size = New Drawing.Size(300, 25)
        Me.Controls.Add(txtRemarks)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(100, 270)
        btnSave.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnSave)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(220, 270)
        btnCancel.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadTitleList()
        LoadPayList()
        LoadWarehouseList()

        If editMode Then
            LoadData()
        End If
    End Sub

    Private Sub LoadTitleList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, title FROM xipunum_erp_finance_title ORDER BY id")
            cmbTitle.Items.Clear()
            For Each row As DataRow In dt.Rows
                cmbTitle.Items.Add(SafeString(row("title")))
            Next
            If cmbTitle.Items.Count > 0 Then cmbTitle.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadPayList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, name FROM xipunum_erp_pay ORDER BY id")
            cmbPay.Items.Clear()
            For Each row As DataRow In dt.Rows
                cmbPay.Items.Add(SafeString(row("name")))
            Next
            If cmbPay.Items.Count > 0 Then cmbPay.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadWarehouseList()
        Try
            Dim sql As String = $"SELECT id, CASE WHEN id = '0' THEN '总库' ELSE title END AS title FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id IN ({UserShopPermission}) ORDER BY id"
            Dim dt As DataTable = ExecuteQuery(sql)
            cmbWarehouse.Items.Clear()
            For Each row As DataRow In dt.Rows
                cmbWarehouse.Items.Add(SafeString(row("title")))
            Next
            If cmbWarehouse.Items.Count > 0 Then cmbWarehouse.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim sql As String = $"SELECT * FROM xipunum_erp_finance WHERE id='{editId}' LIMIT 1"
            Dim dt As DataTable = ExecuteQuery(sql)
            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                cmbType.Text = SafeString(row("type"))
                cmbTitle.Text = SafeString(row("title"))
                txtAmount.Text = SafeString(row("amount"))
                txtRemarks.Text = SafeString(row("remarks"))
            End If
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If String.IsNullOrEmpty(txtAmount.Text) Then
            ShowWarning("请输入金额！")
            Return
        End If

        Try
            Dim sql As String
            If editMode Then
                sql = $"UPDATE xipunum_erp_finance SET type='{cmbType.SelectedItem}', title='{SafeSQL(cmbTitle.Text)}', amount='{txtAmount.Text}', remarks='{SafeSQL(txtRemarks.Text)}', updatetime='{GetOperationDate()}' WHERE id='{editId}'"
            Else
                sql = $"INSERT INTO xipunum_erp_finance (type, title, amount, remarks, kufang, cjuser, creationtime) VALUES ('{cmbType.SelectedItem}', '{SafeSQL(cmbTitle.Text)}', '{txtAmount.Text}', '{SafeSQL(txtRemarks.Text)}', '{cmbWarehouse.SelectedIndex}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
            End If

            ExecuteCommand(sql)
            ShowSuccess("保存成功！")
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

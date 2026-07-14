' ============================================================================
' 收支卡号管理窗口
' 功能: 管理收支使用的银行卡号
' 权限: 83卡号新增/83卡号编辑
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class FinanceAccountForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private cmbPayType As New ComboBox()
    Private cmbWarehouse As New ComboBox()
    Private cmbBank As New ComboBox()
    Private txtZhName As New TextBox()
    Private txtAccount As New TextBox()
    Private txtAddress As New TextBox()
    Private txtRemarks As New TextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnDelete As New Button()
    Private WithEvents btnReset As New Button()
    Private WithEvents btnSearch As New Button()
    Private txtSearch As New TextBox()
    Private currentId As String = ""

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "收支卡号管理"
        Me.Size = New Drawing.Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 180
        Me.Controls.Add(panelTop)

        ' 搜索
        txtSearch.Location = New Drawing.Point(20, 10)
        txtSearch.Size = New Drawing.Size(150, 25)
        txtSearch.Text = "输入户名/账号搜索"
        AddHandler txtSearch.GotFocus, Sub() If txtSearch.Text = "输入户名/账号搜索" Then txtSearch.Text = ""
        AddHandler txtSearch.LostFocus, Sub() If String.IsNullOrEmpty(txtSearch.Text) Then txtSearch.Text = "输入户名/账号搜索"
        panelTop.Controls.Add(txtSearch)

        btnSearch.Text = "搜索"
        btnSearch.Location = New Drawing.Point(180, 10)
        btnSearch.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(btnSearch)

        ' 第一行
        AddLabel(panelTop, "支付方式：", 20, 45)
        cmbPayType.Location = New Drawing.Point(90, 42)
        cmbPayType.Size = New Drawing.Size(120, 25)
        panelTop.Controls.Add(cmbPayType)

        AddLabel(panelTop, "店铺：", 230, 45)
        cmbWarehouse.Location = New Drawing.Point(280, 42)
        cmbWarehouse.Size = New Drawing.Size(120, 25)
        panelTop.Controls.Add(cmbWarehouse)

        AddLabel(panelTop, "开户行：", 420, 45)
        cmbBank.Location = New Drawing.Point(480, 42)
        cmbBank.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(cmbBank)

        ' 第二行
        AddLabel(panelTop, "户名：", 20, 80)
        txtZhName.Location = New Drawing.Point(90, 77)
        txtZhName.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtZhName)

        AddLabel(panelTop, "账号：", 310, 80)
        txtAccount.Location = New Drawing.Point(360, 77)
        txtAccount.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtAccount)

        ' 第三行
        AddLabel(panelTop, "地址：", 20, 115)
        txtAddress.Location = New Drawing.Point(90, 112)
        txtAddress.Size = New Drawing.Size(400, 25)
        panelTop.Controls.Add(txtAddress)

        ' 第四行
        AddLabel(panelTop, "备注：", 20, 150)
        txtRemarks.Location = New Drawing.Point(90, 147)
        txtRemarks.Size = New Drawing.Size(400, 25)
        panelTop.Controls.Add(txtRemarks)

        ' 按钮
        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(520, 42)
        btnSave.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnSave)

        btnDelete.Text = "删除"
        btnDelete.Location = New Drawing.Point(520, 77)
        btnDelete.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnDelete)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(520, 112)
        btnReset.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnReset)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        AddHandler dgvList.CellClick, AddressOf dgvList_CellClick
        Me.Controls.Add(dgvList)

        InitGrid()
    End Sub

    Private Sub AddLabel(parent As Panel, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    Private Sub InitGrid()
        dgvList.Columns.Clear()
        Dim headers() As String = {"ID", "支付方式", "店铺", "开户行", "户名", "账号", "地址", "备注"}
        Dim widths() As Integer = {50, 100, 100, 120, 120, 150, 200, 150}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvList.Columns.Add(col)
        Next
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadPayTypeList()
        LoadWarehouseList()
        LoadBankList()
        LoadData()
    End Sub

    Private Sub LoadPayTypeList()
        Try
            Dim sql As String = "SELECT id, name FROM xipunum_erp_pay WHERE state='0' ORDER BY id ASC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbPayType.Items.Clear()
            cmbPayType.Items.Add(New ComboBoxItem("", "请选择"))
            For Each row As DataRow In dt.Rows
                cmbPayType.Items.Add(New ComboBoxItem(SafeString(row("id")), GBKToUTF8(SafeString(row("name")))))
            Next
            If cmbPayType.Items.Count > 0 Then cmbPayType.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadWarehouseList()
        Try
            Dim shopPermission As String = UserShopPermission
            If String.IsNullOrEmpty(shopPermission) Then shopPermission = "-1"

            Dim sql As String = $"SELECT id, CASE WHEN id = '0' THEN '总库' ELSE title END AS title FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id IN ({shopPermission}) ORDER BY id"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbWarehouse.Items.Clear()
            cmbWarehouse.Items.Add(New ComboBoxItem("", "请选择"))
            For Each row As DataRow In dt.Rows
                cmbWarehouse.Items.Add(New ComboBoxItem(SafeString(row("id")), GBKToUTF8(SafeString(row("title")))))
            Next
            If cmbWarehouse.Items.Count > 0 Then cmbWarehouse.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadBankList()
        Try
            Dim sql As String = "SELECT id, title FROM xipunum_erp_finance_yinhang WHERE state='0' ORDER BY id ASC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbBank.Items.Clear()
            cmbBank.Items.Add(New ComboBoxItem("", "请选择"))
            For Each row As DataRow In dt.Rows
                cmbBank.Items.Add(New ComboBoxItem(SafeString(row("id")), GBKToUTF8(SafeString(row("title")))))
            Next
            If cmbBank.Items.Count > 0 Then cmbBank.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim searchFilter As String = ""
            If Not String.IsNullOrEmpty(txtSearch.Text) AndAlso txtSearch.Text <> "输入户名/账号搜索" Then
                searchFilter = $" AND (a.zhname LIKE '%{SafeSQL(txtSearch.Text)}%' OR a.account LIKE '%{SafeSQL(txtSearch.Text)}%')"
            End If

            Dim sql As String = $"SELECT a.id, b.name AS pay_name, d.title AS dianpu, c.title AS kaihu, a.zhname, a.account, a.address, a.remark " &
                               $"FROM xipunum_erp_finance_account AS a " &
                               $"INNER JOIN xipunum_erp_pay AS b ON b.id = a.type " &
                               $"INNER JOIN xipunum_erp_finance_yinhang AS c ON c.id = a.kaihuhang " &
                               $"INNER JOIN xipunum_erp_type AS d ON d.id = a.kufang " &
                               $"WHERE 1=1 {searchFilter} ORDER BY a.id DESC"

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            dgvList.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    SafeString(row("id")),
                    GBKToUTF8(SafeString(row("pay_name"))),
                    GBKToUTF8(SafeString(row("dianpu"))),
                    GBKToUTF8(SafeString(row("kaihu"))),
                    GBKToUTF8(SafeString(row("zhname"))),
                    SafeString(row("account")),
                    GBKToUTF8(SafeString(row("address"))),
                    GBKToUTF8(SafeString(row("remark")))
                )
            Next
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub dgvList_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        currentId = SafeString(dgvList.Rows(e.RowIndex).Cells("col0").Value)
        txtZhName.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col4").Value)
        txtAccount.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col5").Value)
        txtAddress.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col6").Value)
        txtRemarks.Text = SafeString(dgvList.Rows(e.RowIndex).Cells("col7").Value)
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        LoadData()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If cmbPayType.SelectedIndex <= 0 Then
            ShowWarning("请选择支付方式！")
            Return
        End If
        If cmbWarehouse.SelectedIndex <= 0 Then
            ShowWarning("请选择店铺！")
            Return
        End If
        If String.IsNullOrEmpty(txtZhName.Text.Trim()) Then
            ShowWarning("请输入户名！")
            Return
        End If

        ' 权限检查
        If Not HasOperationPermission("83") Then
            ShowWarning("没有卡号管理权限！")
            Return
        End If

        Try
            Dim payTypeId As String = DirectCast(cmbPayType.SelectedItem, ComboBoxItem).ID
            Dim warehouseId As String = DirectCast(cmbWarehouse.SelectedItem, ComboBoxItem).ID
            Dim bankId As String = If(cmbBank.SelectedIndex > 0, DirectCast(cmbBank.SelectedItem, ComboBoxItem).ID, "")

            Dim sql As String
            If String.IsNullOrEmpty(currentId) Then
                sql = $"INSERT INTO xipunum_erp_finance_account (type, kufang, kaihuhang, zhname, account, address, remark, cjuser, creationtime) VALUES ('{payTypeId}', '{warehouseId}', '{bankId}', '{SafeSQL(txtZhName.Text)}', '{SafeSQL(txtAccount.Text)}', '{SafeSQL(txtAddress.Text)}', '{SafeSQL(txtRemarks.Text)}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
            Else
                sql = $"UPDATE xipunum_erp_finance_account SET kaihuhang='{bankId}', zhname='{SafeSQL(txtZhName.Text)}', account='{SafeSQL(txtAccount.Text)}', address='{SafeSQL(txtAddress.Text)}', remark='{SafeSQL(txtRemarks.Text)}', updatetime='{GetOperationDate()}' WHERE id='{currentId}'"
            End If
            DatabaseModule.ExecuteCommand(sql)
            AddSystemLog(If(String.IsNullOrEmpty(currentId), "添加", "编辑"), If(String.IsNullOrEmpty(currentId), "添加卡号", "编辑卡号"), "卡号：" & txtAccount.Text)
            ShowSuccess("保存成功！")
            LoadData()
            btnReset_Click(Nothing, Nothing)
        Catch ex As Exception
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If String.IsNullOrEmpty(currentId) Then
            ShowWarning("请先选择要删除的记录！")
            Return
        End If

        ' 权限检查
        If Not HasOperationPermission("83") Then
            ShowWarning("没有卡号管理权限！")
            Return
        End If

        If Not ConfirmAction("确定要删除吗？") Then Return
        Try
            DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_finance_account WHERE id='{currentId}'")
            AddSystemLog("删除", "删除卡号", "卡号：" & txtAccount.Text)
            ShowSuccess("删除成功！")
            LoadData()
            btnReset_Click(Nothing, Nothing)
        Catch ex As Exception
            ShowError("删除失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        currentId = ""
        txtZhName.Text = ""
        txtAccount.Text = ""
        txtAddress.Text = ""
        txtRemarks.Text = ""
        If cmbPayType.Items.Count > 0 Then cmbPayType.SelectedIndex = 0
        If cmbWarehouse.Items.Count > 0 Then cmbWarehouse.SelectedIndex = 0
        If cmbBank.Items.Count > 0 Then cmbBank.SelectedIndex = 0
    End Sub

    Private Class ComboBoxItem
        Public Property ID As String
        Public Property Text As String
        Public Sub New(id As String, text As String)
            Me.ID = id
            Me.Text = text
        End Sub
        Public Overrides Function ToString() As String
            Return Text
        End Function
    End Class
End Class

' ============================================================================
' 旧料管理添加窗口
' 功能: 添加旧料出入库单据
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class MaterialAddForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private cmbFactory As New ComboBox()
    Private txtRemarks As New TextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnCancel As New Button()
    Private WithEvents btnAddItem As New Button()
    Private WithEvents btnDeleteItem As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "添加旧料单据"
        Me.Size = New Drawing.Size(800, 500)
        Me.StartPosition = FormStartPosition.CenterParent

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 80
        Me.Controls.Add(panelTop)

        Dim lblFactory As New Label()
        lblFactory.Text = "工厂："
        lblFactory.Location = New Drawing.Point(20, 15)
        lblFactory.AutoSize = True
        panelTop.Controls.Add(lblFactory)

        cmbFactory.Location = New Drawing.Point(70, 12)
        cmbFactory.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(cmbFactory)

        Dim lblRemarks As New Label()
        lblRemarks.Text = "备注："
        lblRemarks.Location = New Drawing.Point(290, 15)
        lblRemarks.AutoSize = True
        panelTop.Controls.Add(lblRemarks)

        txtRemarks.Location = New Drawing.Point(340, 12)
        txtRemarks.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtRemarks)

        btnAddItem.Text = "添加项目"
        btnAddItem.Location = New Drawing.Point(20, 45)
        btnAddItem.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnAddItem)

        btnDeleteItem.Text = "删除项目"
        btnDeleteItem.Location = New Drawing.Point(120, 45)
        btnDeleteItem.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnDeleteItem)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(550, 45)
        btnSave.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnSave)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(650, 45)
        btnCancel.Size = New Drawing.Size(80, 25)
        panelTop.Controls.Add(btnCancel)

        dgvList.Dock = DockStyle.Fill
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadFactoryList()
        InitGrid()
    End Sub

    Private Sub LoadFactoryList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, title FROM xipunum_erp_about ORDER BY id")
            cmbFactory.Items.Clear()
            For Each row As DataRow In dt.Rows
                cmbFactory.Items.Add(SafeString(row("title")))
            Next
            If cmbFactory.Items.Count > 0 Then cmbFactory.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub InitGrid()
        dgvList.Columns.Clear()
        dgvList.Columns.Add("mingcheng", "名称")
        dgvList.Columns.Add("zhongliang", "重量")
        dgvList.Columns.Add("jieyu", "结余")
        dgvList.Columns.Add("beizhu", "备注")
        dgvList.Columns("mingcheng").Width = 200
        dgvList.Columns("zhongliang").Width = 100
        dgvList.Columns("jieyu").Width = 100
        dgvList.Columns("beizhu").Width = 200
    End Sub

    Private Sub btnAddItem_Click(sender As Object, e As EventArgs) Handles btnAddItem.Click
        dgvList.Rows.Add("", 0, 0, "")
    End Sub

    Private Sub btnDeleteItem_Click(sender As Object, e As EventArgs) Handles btnDeleteItem.Click
        If dgvList.SelectedRows.Count > 0 Then
            dgvList.Rows.Remove(dgvList.SelectedRows(0))
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If dgvList.Rows.Count = 0 Then
            ShowWarning("请添加项目！")
            Return
        End If

        Dim trans As MySqlTransaction = BeginTransaction()
        Try
            Dim orderNumber As String = GenerateOrderNumber("JL")
            Dim factoryId As String = cmbFactory.SelectedIndex.ToString()

            Dim sql As String = $"INSERT INTO xipunum_erp_material_order (number, gongchang, remarks, cjuser, creationtime) VALUES ('{orderNumber}', '{factoryId}', '{SafeSQL(txtRemarks.Text)}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
            ExecuteCommand(sql, MySQL_Write)

            sql = "SELECT LAST_INSERT_ID() AS id"
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_Write)
            Dim orderId As String = SafeString(dt.Rows(0)("id"))

            For Each row As DataGridViewRow In dgvList.Rows
                Dim mingcheng As String = SafeString(row.Cells("mingcheng").Value)
                Dim zhongliang As Decimal = SafeDecimal(row.Cells("zhongliang").Value)
                Dim jieyu As Decimal = SafeDecimal(row.Cells("jieyu").Value)
                Dim beizhu As String = SafeString(row.Cells("beizhu").Value)

                sql = $"INSERT INTO xipunum_erp_material (order_id, product_name, zhongliang, jieyu, kufang, remarks, cjuser, creationtime) VALUES ('{orderId}', '{SafeSQL(mingcheng)}', '{zhongliang}', '{jieyu}', '0', '{SafeSQL(beizhu)}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
                ExecuteCommand(sql, MySQL_Write)
            Next

            CommitTransaction(trans)
            ShowSuccess("保存成功！")
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            RollbackTransaction(trans)
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

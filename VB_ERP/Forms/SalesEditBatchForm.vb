' ============================================================================
' 销售编辑批量修改窗口
' 功能: 批量修改销售编辑信息
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SalesEditBatchForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private cmbGuide As New ComboBox()
    Private txtGoldPrice As New TextBox()
    Private txtSalesCost As New TextBox()
    Private txtSalesSurcharge As New TextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnCancel As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "销售编辑批量修改"
        Me.Size = New Drawing.Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterParent

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 100
        Me.Controls.Add(panelTop)

        Dim lblGuide As New Label()
        lblGuide.Text = "导购员："
        lblGuide.Location = New Drawing.Point(20, 15)
        lblGuide.AutoSize = True
        panelTop.Controls.Add(lblGuide)

        cmbGuide.Location = New Drawing.Point(80, 12)
        cmbGuide.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(cmbGuide)

        Dim lblGoldPrice As New Label()
        lblGoldPrice.Text = "克价："
        lblGoldPrice.Location = New Drawing.Point(250, 15)
        lblGoldPrice.AutoSize = True
        panelTop.Controls.Add(lblGoldPrice)

        txtGoldPrice.Location = New Drawing.Point(290, 12)
        txtGoldPrice.Size = New Drawing.Size(100, 25)
        panelTop.Controls.Add(txtGoldPrice)

        Dim lblSalesCost As New Label()
        lblSalesCost.Text = "工费："
        lblSalesCost.Location = New Drawing.Point(410, 15)
        lblSalesCost.AutoSize = True
        panelTop.Controls.Add(lblSalesCost)

        txtSalesCost.Location = New Drawing.Point(450, 12)
        txtSalesCost.Size = New Drawing.Size(100, 25)
        panelTop.Controls.Add(txtSalesCost)

        Dim lblSalesSurcharge As New Label()
        lblSalesSurcharge.Text = "附加费："
        lblSalesSurcharge.Location = New Drawing.Point(570, 15)
        lblSalesSurcharge.AutoSize = True
        panelTop.Controls.Add(lblSalesSurcharge)

        txtSalesSurcharge.Location = New Drawing.Point(630, 12)
        txtSalesSurcharge.Size = New Drawing.Size(100, 25)
        panelTop.Controls.Add(txtSalesSurcharge)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(20, 60)
        btnSave.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnSave)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(120, 60)
        btnCancel.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnCancel)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadGuideList()
        LoadData()
    End Sub

    Private Sub LoadGuideList()
        Try
            Dim sql As String = "SELECT name FROM xipunum_erp_user WHERE state='0' ORDER BY id"
            Dim dt As DataTable = ExecuteQuery(sql)
            cmbGuide.Items.Clear()
            cmbGuide.Items.Add("请选择导购员")
            For Each row As DataRow In dt.Rows
                cmbGuide.Items.Add(SafeString(row("name")))
            Next
            If cmbGuide.Items.Count > 0 Then cmbGuide.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT b.poduct_code, a.product_name, b.quantity, b.net_weight, b.gold_price, b.sales_cost, b.sales_surcharge, b.shopping_guide FROM xipunum_erp_outbound AS b INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code ORDER BY b.id DESC LIMIT 100")
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要修改的记录！")
            Return
        End If

        If Not ConfirmAction("确定要批量修改吗？") Then Return

        Try
            For Each row As DataGridViewRow In dgvList.SelectedRows
                Dim productCode As String = SafeString(row.Cells("poduct_code").Value)

                Dim sql As String = "UPDATE xipunum_erp_outbound SET "
                Dim updates As New List(Of String)()

                If cmbGuide.SelectedIndex > 0 Then
                    updates.Add($"shopping_guide='{SafeSQL(cmbGuide.SelectedItem.ToString())}'")
                End If
                If Not String.IsNullOrEmpty(txtGoldPrice.Text) Then
                    updates.Add($"gold_price='{txtGoldPrice.Text}'")
                End If
                If Not String.IsNullOrEmpty(txtSalesCost.Text) Then
                    updates.Add($"sales_cost='{txtSalesCost.Text}'")
                End If
                If Not String.IsNullOrEmpty(txtSalesSurcharge.Text) Then
                    updates.Add($"sales_surcharge='{txtSalesSurcharge.Text}'")
                End If

                If updates.Count > 0 Then
                    sql &= String.Join(", ", updates) & $", updatetime='{GetOperationDate()}' WHERE poduct_code='{SafeSQL(productCode)}'"
                    ExecuteCommand(sql)
                End If
            Next

            ShowSuccess("修改成功！")
            LoadData()
        Catch ex As Exception
            ShowError("修改失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class

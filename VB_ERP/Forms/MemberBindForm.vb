' ============================================================================
' 成品销售会员绑定窗口
' 功能: 为销售订单绑定/修改会员
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class MemberBindForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtSearch As New TextBox()
    Private WithEvents btnSearch As New Button()
    Private WithEvents btnBind As New Button()
    Private WithEvents btnUnbind As New Button()
    Private orderId As String = ""

    Public Sub New(orderId As String)
        Me.orderId = orderId
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "会员绑定"
        Me.Size = New Drawing.Size(800, 500)
        Me.StartPosition = FormStartPosition.CenterParent

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 60
        Me.Controls.Add(panelTop)

        Dim lblSearch As New Label()
        lblSearch.Text = "搜索："
        lblSearch.Location = New Drawing.Point(20, 18)
        lblSearch.AutoSize = True
        panelTop.Controls.Add(lblSearch)

        txtSearch.Location = New Drawing.Point(60, 15)
        txtSearch.Size = New Drawing.Size(250, 25)
        panelTop.Controls.Add(txtSearch)

        btnSearch.Text = "搜索"
        btnSearch.Location = New Drawing.Point(330, 13)
        btnSearch.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnSearch)

        btnBind.Text = "绑定"
        btnBind.Location = New Drawing.Point(430, 13)
        btnBind.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnBind)

        btnUnbind.Text = "解绑"
        btnUnbind.Location = New Drawing.Point(530, 13)
        btnUnbind.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnUnbind)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim searchFilter As String = ""
            If Not String.IsNullOrEmpty(txtSearch.Text.Trim()) Then
                searchFilter = $" AND (a.customer_code LIKE '%{SafeSQL(txtSearch.Text)}%' OR a.memberid LIKE '%{SafeSQL(txtSearch.Text)}%' OR a.name LIKE '%{SafeSQL(txtSearch.Text)}%' OR a.tel LIKE '%{SafeSQL(txtSearch.Text)}%')"
            End If

            Dim sql As String = $"SELECT a.customer_code, a.memberid, a.name, a.tel, a.shengri, a.dizhi FROM xipunum_erp_member AS a WHERE 1=1 {searchFilter} ORDER BY a.creationtime DESC LIMIT 1000"
            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        LoadData()
    End Sub

    Private Sub btnBind_Click(sender As Object, e As EventArgs) Handles btnBind.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要绑定的会员！")
            Return
        End If

        Try
            Dim customerCode As String = SafeString(dgvList.SelectedRows(0).Cells("customer_code").Value)
            ExecuteCommand($"UPDATE xipunum_erp_outbound_order SET customer_code='{SafeSQL(customerCode)}' WHERE id='{orderId}'")
            AddSystemLog("修改", "绑定会员", $"订单ID：{orderId}，会员：{customerCode}")
            ShowSuccess("绑定成功！")
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            ShowError("绑定失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnUnbind_Click(sender As Object, e As EventArgs) Handles btnUnbind.Click
        If Not ConfirmAction("确定要解绑吗？") Then Return

        Try
            ExecuteCommand($"UPDATE xipunum_erp_outbound_order SET customer_code='' WHERE id='{orderId}'")
            AddSystemLog("修改", "解绑会员", $"订单ID：{orderId}")
            ShowSuccess("解绑成功！")
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            ShowError("解绑失败：" & ex.Message)
        End Try
    End Sub
End Class

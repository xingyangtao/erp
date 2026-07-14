' ============================================================================
' 款号管理匹配窗口
' 功能: 款号匹配商品
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class KuanHaoMatchForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtSearch As New TextBox()
    Private WithEvents btnSearch As New Button()
    Private WithEvents btnMatch As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "款号管理匹配"
        Me.Size = New Drawing.Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 50
        Me.Controls.Add(panelTop)

        Dim lblSearch As New Label()
        lblSearch.Text = "搜索："
        lblSearch.Location = New Drawing.Point(20, 15)
        lblSearch.AutoSize = True
        panelTop.Controls.Add(lblSearch)

        txtSearch.Location = New Drawing.Point(60, 12)
        txtSearch.Size = New Drawing.Size(250, 25)
        panelTop.Controls.Add(txtSearch)

        btnSearch.Text = "搜索"
        btnSearch.Location = New Drawing.Point(330, 10)
        btnSearch.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnSearch)

        btnMatch.Text = "匹配"
        btnMatch.Location = New Drawing.Point(420, 10)
        btnMatch.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnMatch)

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
            Dim sql As String = "SELECT a.poduct_code, a.product_name, a.item_number, b.title AS pinlei, c.title AS guige FROM xipunum_erp_shop AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id WHERE a.item_number='' OR a.item_number IS NULL ORDER BY a.id DESC LIMIT 1000"
            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try
            Dim searchFilter As String = ""
            If Not String.IsNullOrEmpty(txtSearch.Text.Trim()) Then
                searchFilter = $" AND (a.poduct_code LIKE '%{SafeSQL(txtSearch.Text)}%' OR a.product_name LIKE '%{SafeSQL(txtSearch.Text)}%')"
            End If

            Dim sql As String = $"SELECT a.poduct_code, a.product_name, a.item_number, b.title AS pinlei, c.title AS guige FROM xipunum_erp_shop AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id WHERE (a.item_number='' OR a.item_number IS NULL) {searchFilter} ORDER BY a.id DESC LIMIT 1000"
            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnMatch_Click(sender As Object, e As EventArgs) Handles btnMatch.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要匹配的商品！")
            Return
        End If

        Dim form As New KuanHaoSelectForm()
        If form.ShowDialog() = DialogResult.OK Then
            Try
                For Each row As DataGridViewRow In dgvList.SelectedRows
                    Dim code As String = SafeString(row.Cells("poduct_code").Value)
                    ExecuteCommand($"UPDATE xipunum_erp_shop SET item_number='{SafeSQL(form.SelectedKuanHao)}' WHERE poduct_code='{code}'")
                Next
                ShowSuccess("匹配成功！")
                LoadData()
            Catch ex As Exception
                ShowError("匹配失败：" & ex.Message)
            End Try
        End If
    End Sub
End Class

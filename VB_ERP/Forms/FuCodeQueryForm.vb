' ============================================================================
' 成品修改副编码查询窗口
' 功能: 成品修改副编码查询
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class FuCodeQueryForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtCode As New TextBox()
    Private WithEvents btnSearch As New Button()
    Private warehouseId As String = ""

    Public Sub New(warehouseId As String)
        Me.warehouseId = warehouseId
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        Me.Text = "副编码查询"
        Me.Size = New Drawing.Size(800, 500)
        Me.StartPosition = FormStartPosition.CenterParent

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 50
        Me.Controls.Add(panelTop)

        Dim lblCode As New Label()
        lblCode.Text = "编码："
        lblCode.Location = New Drawing.Point(20, 15)
        lblCode.AutoSize = True
        panelTop.Controls.Add(lblCode)

        txtCode.Location = New Drawing.Point(60, 12)
        txtCode.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtCode)

        btnSearch.Text = "搜索"
        btnSearch.Location = New Drawing.Point(280, 10)
        btnSearch.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnSearch)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)

        InitGrid()
    End Sub

    Private Sub InitGrid()
        dgvList.Columns.Clear()
        dgvList.Columns.Add("poduct_code", "商品编码")
        dgvList.Columns.Add("fu_code", "副编码")
        dgvList.Columns.Add("kufang", "库房")
        dgvList.Columns.Add("quantity", "数量")
        dgvList.Columns.Add("jinzhong", "金重")
        dgvList.Columns.Add("creationtime", "入库时间")

        dgvList.Columns("poduct_code").Width = 120
        dgvList.Columns("fu_code").Width = 120
        dgvList.Columns("kufang").Width = 80
        dgvList.Columns("quantity").Width = 60
        dgvList.Columns("jinzhong").Width = 80
        dgvList.Columns("creationtime").Width = 140
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        If String.IsNullOrEmpty(txtCode.Text.Trim()) Then
            ShowWarning("请输入编码！")
            Return
        End If

        Try
            Dim sql As String = $"SELECT a.poduct_code, b.fu_code, CASE WHEN a.kufang='0' THEN '总库' ELSE c.title END AS kufang, a.quantity, a.jinzhong, b.creationtime FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code LEFT JOIN xipunum_erp_type AS c ON c.id = a.kufang WHERE a.kufang = '{warehouseId}' AND (a.quantity > 0 OR a.jinzhong > 0) AND (b.poduct_code = '{SafeSQL(txtCode.Text.Trim())}' OR b.fu_code = '{SafeSQL(txtCode.Text.Trim())}') ORDER BY a.id DESC"

            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    SafeString(row("poduct_code")),
                    SafeString(row("fu_code")),
                    SafeString(row("kufang")),
                    SafeDecimal(row("quantity")),
                    SafeDecimal(row("jinzhong")),
                    SafeString(row("creationtime"))
                )
            Next
        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        End Try
    End Sub
End Class

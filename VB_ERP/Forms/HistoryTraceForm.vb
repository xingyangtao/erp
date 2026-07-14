' ============================================================================
' 历史追溯窗口
' 功能: 商品历史追溯查询
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class HistoryTraceForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtCode As New TextBox()
    Private WithEvents btnSearch As New Button()

    Public Sub New()
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        Me.Text = "历史追溯"
        Me.Size = New Drawing.Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 50
        Me.Controls.Add(panelTop)

        Dim lblCode As New Label()
        lblCode.Text = "商品编码："
        lblCode.Location = New Drawing.Point(20, 15)
        lblCode.AutoSize = True
        panelTop.Controls.Add(lblCode)

        txtCode.Location = New Drawing.Point(90, 12)
        txtCode.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtCode)

        btnSearch.Text = "搜索"
        btnSearch.Location = New Drawing.Point(310, 10)
        btnSearch.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnSearch)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        If String.IsNullOrEmpty(txtCode.Text.Trim()) Then
            ShowWarning("请输入商品编码！")
            Return
        End If

        Try
            ' 查询入库历史
            Dim sql As String = $"SELECT '入库' AS type, b.number AS order_number, b.creationtime, a.quantity, a.net_weight FROM xipunum_erp_store AS a INNER JOIN xipunum_erp_store_order AS b ON b.id = a.order_id WHERE a.poduct_code='{SafeSQL(txtCode.Text.Trim())}'"
            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        End Try
    End Sub
End Class

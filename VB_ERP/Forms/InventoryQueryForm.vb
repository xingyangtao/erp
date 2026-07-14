' ============================================================================
' 实时库存查询窗口
' 功能: 查询当前实时库存数据
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class InventoryQueryForm
    Inherits System.Windows.Forms.Form

    Private dgvReport As New DataGridView()
    Private txtSearch As New TextBox()
    Private cmbCategory As New ComboBox()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnReset As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "实时库存查询"
        Me.Size = New Drawing.Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 60
        Me.Controls.Add(panelTop)

        AddLabel(panelTop, "搜索：", 20, 18)
        txtSearch.Location = New Drawing.Point(60, 15)
        txtSearch.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtSearch)

        AddLabel(panelTop, "品类：", 280, 18)
        cmbCategory.Location = New Drawing.Point(320, 15)
        cmbCategory.Size = New Drawing.Size(120, 25)
        panelTop.Controls.Add(cmbCategory)

        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(460, 13)
        btnQuery.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnQuery)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(550, 13)
        btnExport.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnExport)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(640, 13)
        btnReset.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnReset)

        dgvReport.Dock = DockStyle.Fill
        dgvReport.ReadOnly = True
        dgvReport.AllowUserToAddRows = False
        dgvReport.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvReport)

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
        dgvReport.Columns.Clear()
        Dim headers() As String = {"商品编码", "商品名称", "款号", "品类", "规格", "材质", "单件重", "数量", "金重", "库房"}
        Dim widths() As Integer = {120, 150, 100, 80, 80, 80, 70, 60, 80, 80}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvReport.Columns.Add(col)
        Next
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadCategoryList()
        LoadData()
    End Sub

    Private Sub LoadCategoryList()
        Try
            Dim dt As DataTable = DatabaseModule.ExecuteQuery("SELECT id, title FROM xipunum_erp_category ORDER BY id")
            cmbCategory.Items.Clear()
            cmbCategory.Items.Add("全部品类")
            For Each row As DataRow In dt.Rows
                cmbCategory.Items.Add(GBKToUTF8(SafeString(row("title"))))
            Next
            If cmbCategory.Items.Count > 0 Then cmbCategory.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim shopPermission As String = BuildShopPermission("b.kufang")
            Dim sql As String = GetRealtimeInventorySQL(shopPermission)
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)

            dgvReport.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvReport.Rows.Add(
                    SafeString(row("poduct_code")),
                    GBKToUTF8(SafeString(row("product_name"))),
                    SafeString(row("item_number")),
                    GBKToUTF8(SafeString(row("pinlei"))),
                    GBKToUTF8(SafeString(row("guige"))),
                    GBKToUTF8(SafeString(row("caizhi"))),
                    SafeDecimal(row("danzhong")).ToString("F3"),
                    SafeDecimal(row("shuliang")).ToString("F2"),
                    SafeDecimal(row("jinzhong")).ToString("F3"),
                    GBKToUTF8(SafeString(row("kufang")))
                )
            Next
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        LoadData()
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If dgvReport.Rows.Count = 0 Then
            ShowWarning("没有数据可导出！")
            Return
        End If
        Try
            Dim dt As New DataTable()
            For Each col As DataGridViewColumn In dgvReport.Columns
                dt.Columns.Add(col.HeaderText)
            Next
            For Each row As DataGridViewRow In dgvReport.Rows
                Dim dr As DataRow = dt.NewRow()
                For i As Integer = 0 To dgvReport.Columns.Count - 1
                    dr(i) = SafeString(row.Cells(i).Value)
                Next
                dt.Rows.Add(dr)
            Next
            ExportToExcel(dt, "实时库存查询.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        txtSearch.Text = ""
        cmbCategory.SelectedIndex = 0
        LoadData()
    End Sub
End Class

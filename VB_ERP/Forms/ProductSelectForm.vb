' ============================================================================
' 商品选择窗口
' 用于选择商品，返回商品信息
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class ProductSelectForm
    Inherits System.Windows.Forms.Form

    ' ========== 返回的商品信息 ==========
    Public Property ProductCode As String = ""
    Public Overloads Property ProductName As String = ""
    Public Property KuanHao As String = ""
    Public Property CategoryName As String = ""
    Public Property SpecName As String = ""
    Public Property CaiZhi As String = ""
    Public Property DanZhong As Decimal = 0
    Public Property NetWeight As Decimal = 0
    Public Property Weight As Decimal = 0

    ' ========== 控件声明 ==========
    Private WithEvents txtSearch As New TextBox()
    Private WithEvents btnSearch As New Button()
    Private dgvProducts As New DataGridView()
    Private WithEvents btnSelect As New Button()
    Private WithEvents btnCancel As New Button()

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf ProductSelectForm_Load
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "选择商品"
        Me.Size = New Drawing.Size(800, 500)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        ' 搜索区
        Dim panelSearch As New Panel()
        panelSearch.Dock = DockStyle.Top
        panelSearch.Height = 50
        Me.Controls.Add(panelSearch)

        Dim lblSearch As New Label()
        lblSearch.Text = "搜索："
        lblSearch.Location = New Drawing.Point(20, 15)
        lblSearch.AutoSize = True
        panelSearch.Controls.Add(lblSearch)

        txtSearch.Location = New Drawing.Point(70, 12)
        txtSearch.Size = New Drawing.Size(300, 25)
        panelSearch.Controls.Add(txtSearch)

        btnSearch.Text = "搜索"
        btnSearch.Location = New Drawing.Point(390, 10)
        btnSearch.Size = New Drawing.Size(80, 30)
        panelSearch.Controls.Add(btnSearch)

        ' 商品列表
        dgvProducts.Dock = DockStyle.Fill
        dgvProducts.Font = New Drawing.Font("微软雅黑", 10)
        dgvProducts.ReadOnly = True
        dgvProducts.AllowUserToAddRows = False
        dgvProducts.AllowUserToDeleteRows = False
        dgvProducts.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvProducts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Me.Controls.Add(dgvProducts)

        ' 按钮区
        Dim panelBottom As New Panel()
        panelBottom.Dock = DockStyle.Bottom
        panelBottom.Height = 50
        Me.Controls.Add(panelBottom)

        btnSelect.Text = "选择"
        btnSelect.Location = New Drawing.Point(600, 10)
        btnSelect.Size = New Drawing.Size(80, 30)
        panelBottom.Controls.Add(btnSelect)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(700, 10)
        btnCancel.Size = New Drawing.Size(80, 30)
        panelBottom.Controls.Add(btnCancel)
    End Sub

    ' ========== 窗口加载 ==========
    Private Sub ProductSelectForm_Load(sender As Object, e As EventArgs)
        ' 初始化表格列
        dgvProducts.Columns.Clear()
        dgvProducts.Columns.Add("bianma", "商品编码")
        dgvProducts.Columns.Add("mingcheng", "商品名称")
        dgvProducts.Columns.Add("kuanhao", "款号")
        dgvProducts.Columns.Add("pinlei", "品类")
        dgvProducts.Columns.Add("guige", "规格")
        dgvProducts.Columns.Add("caizhi", "材质")
        dgvProducts.Columns.Add("danzhong", "单件重")
        dgvProducts.Columns.Add("kucunshuliang", "库存数量")
        dgvProducts.Columns.Add("kucunjinzhong", "库存金重")

        ' 加载商品列表
        LoadProducts()
    End Sub

    ' ========== 加载商品列表 ==========
    Private Sub LoadProducts(Optional searchText As String = "")
        Try
            Dim searchFilter As String = ""
            If Not String.IsNullOrEmpty(searchText) Then
                searchFilter = $" AND (a.poduct_code LIKE '%{SafeSQL(searchText)}%' OR a.product_name LIKE '%{SafeSQL(searchText)}%' OR a.item_number LIKE '%{SafeSQL(searchText)}%')"
            End If

            Dim sql As String = $"SELECT a.poduct_code AS bianma, a.product_name AS mingcheng, 
                                        a.item_number AS kuanhao, d.title AS pinlei, e.title AS guige, 
                                        a.caizhi AS caizhi, a.single AS danzhong,
                                        CAST(ROUND(COALESCE(SUM(b.quantity),0), 2) AS DECIMAL(30,2)) AS kucunshuliang,
                                        CAST(ROUND(COALESCE(SUM(b.jinzhong),0), 3) AS DECIMAL(30,3)) AS kucunjinzhong
                                 FROM xipunum_erp_shop AS a
                                 LEFT JOIN xipunum_erp_shop_kucun AS b ON b.poduct_code = a.poduct_code
                                 LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id
                                 LEFT JOIN xipunum_erp_specs AS e ON e.id = a.specification_id
                                 WHERE 1=1 {searchFilter}
                                 GROUP BY a.poduct_code
                                 HAVING kucunshuliang > 0 OR kucunjinzhong > 0
                                 ORDER BY a.id DESC
                                 LIMIT 1000"
            Dim dt As DataTable = ExecuteQuery(sql)

            dgvProducts.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvProducts.Rows.Add(
                    SafeString(row("bianma")),
                    SafeString(row("mingcheng")),
                    SafeString(row("kuanhao")),
                    SafeString(row("pinlei")),
                    SafeString(row("guige")),
                    SafeString(row("caizhi")),
                    SafeDecimal(row("danzhong")),
                    SafeDecimal(row("kucunshuliang")),
                    SafeDecimal(row("kucunjinzhong"))
                )
            Next
        Catch ex As Exception
            ShowError("加载商品列表失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 搜索按钮点击 ==========
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        LoadProducts(txtSearch.Text.Trim())
    End Sub

    ' ========== 搜索框回车 ==========
    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            LoadProducts(txtSearch.Text.Trim())
        End If
    End Sub

    ' ========== 选择按钮点击 ==========
    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        If dgvProducts.SelectedRows.Count = 0 Then
            ShowWarning("请选择商品！")
            Return
        End If

        Dim row As DataGridViewRow = dgvProducts.SelectedRows(0)
        ProductCode = SafeString(row.Cells("bianma").Value)
        ProductName = SafeString(row.Cells("mingcheng").Value)
        KuanHao = SafeString(row.Cells("kuanhao").Value)
        CategoryName = SafeString(row.Cells("pinlei").Value)
        SpecName = SafeString(row.Cells("guige").Value)
        CaiZhi = SafeString(row.Cells("caizhi").Value)
        DanZhong = SafeDecimal(row.Cells("danzhong").Value)
        NetWeight = SafeDecimal(row.Cells("kucunjinzhong").Value)
        Weight = SafeDecimal(row.Cells("danzhong").Value)

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    ' ========== 取消按钮点击 ==========
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

End Class

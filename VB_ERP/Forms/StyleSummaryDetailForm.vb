' ============================================================================
' 款式数据汇总明细窗口
' 功能: 款号数据明细查看
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class StyleSummaryDetailForm
    Inherits System.Windows.Forms.Form

    Private dgvInbound As New DataGridView()
    Private dgvSales As New DataGridView()
    Private dgvInventory As New DataGridView()
    Private kuanHaoId As String = ""
    Private WithEvents tabControl As New TabControl()

    Public Sub New(id As String)
        kuanHaoId = id
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "款式数据汇总明细"
        Me.Size = New Drawing.Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        tabControl.Dock = DockStyle.Fill
        Me.Controls.Add(tabControl)

        ' 入库明细
        Dim tabInbound As New TabPage("入库明细")
        tabControl.TabPages.Add(tabInbound)
        dgvInbound.Dock = DockStyle.Fill
        dgvInbound.ReadOnly = True
        dgvInbound.AllowUserToAddRows = False
        tabInbound.Controls.Add(dgvInbound)

        ' 销售明细
        Dim tabSales As New TabPage("销售明细")
        tabControl.TabPages.Add(tabSales)
        dgvSales.Dock = DockStyle.Fill
        dgvSales.ReadOnly = True
        dgvSales.AllowUserToAddRows = False
        tabSales.Controls.Add(dgvSales)

        ' 库存明细
        Dim tabInventory As New TabPage("库存明细")
        tabControl.TabPages.Add(tabInventory)
        dgvInventory.Dock = DockStyle.Fill
        dgvInventory.ReadOnly = True
        dgvInventory.AllowUserToAddRows = False
        tabInventory.Controls.Add(dgvInventory)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            ' 查询款号信息
            Dim dt As DataTable = ExecuteQuery($"SELECT kuanhao FROM xipunum_erp_ksiamges WHERE id='{kuanHaoId}' LIMIT 1")
            If dt.Rows.Count = 0 Then
                ShowError("款号不存在！")
                Me.Close()
                Return
            End If
            Dim kuanhao As String = SafeString(dt.Rows(0)("kuanhao"))

            ' 入库明细
            Dim sql As String = $"SELECT b.number, b.creationtime, a.product_name, a.quantity, a.net_weight, a.weight, a.cost_price FROM xipunum_erp_store AS a INNER JOIN xipunum_erp_store_order AS b ON b.id = a.order_id WHERE a.item_number='{kuanhao}' ORDER BY b.creationtime DESC"
            dgvInbound.DataSource = ExecuteQuery(sql)
            dgvInbound.AutoResizeColumns()

            ' 销售明细
            sql = $"SELECT a.settlement_number, a.creationtime, b.product_name, b.quantity, b.net_weight, b.xiao_amount, b.settlement FROM xipunum_erp_outbound_order AS a INNER JOIN xipunum_erp_outbound AS b ON b.order_id = a.id INNER JOIN xipunum_erp_shop AS c ON c.poduct_code = b.poduct_code WHERE c.item_number='{kuanhao}' ORDER BY a.creationtime DESC"
            dgvSales.DataSource = ExecuteQuery(sql)
            dgvSales.AutoResizeColumns()

            ' 库存明细
            sql = $"SELECT a.poduct_code, a.product_name, b.quantity, b.jinzhong, b.kufang FROM xipunum_erp_shop AS a INNER JOIN xipunum_erp_shop_kucun AS b ON b.poduct_code = a.poduct_code WHERE a.item_number='{kuanhao}' AND (b.quantity > 0 OR b.jinzhong > 0) ORDER BY b.kufang"
            dgvInventory.DataSource = ExecuteQuery(sql)
            dgvInventory.AutoResizeColumns()
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub
End Class

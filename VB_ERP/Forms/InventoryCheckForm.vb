' ============================================================================
' 物资盘点添加窗口
' 功能: 库存盘点操作
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class InventoryCheckForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private cmbWarehouse As New ComboBox()
    Private WithEvents btnLoad As New Button()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnReset As New Button()
    Private WithEvents btnExport As New Button()
    Private lblSummary As New Label()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler dgvList.CellValueChanged, AddressOf DgvList_CellValueChanged
        AddHandler dgvList.CellEndEdit, AddressOf DgvList_CellEndEdit
    End Sub

    Private Sub InitializeUI()
        Me.Text = "物资盘点添加"
        Me.Size = New Drawing.Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 60
        Me.Controls.Add(panelTop)

        Dim lblWarehouse As New Label()
        lblWarehouse.Text = "库房："
        lblWarehouse.Location = New Drawing.Point(20, 18)
        lblWarehouse.AutoSize = True
        panelTop.Controls.Add(lblWarehouse)

        cmbWarehouse.Location = New Drawing.Point(60, 15)
        cmbWarehouse.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(cmbWarehouse)

        btnLoad.Text = "加载库存"
        btnLoad.Location = New Drawing.Point(230, 13)
        btnLoad.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnLoad)

        btnSave.Text = "保存盘点"
        btnSave.Location = New Drawing.Point(330, 13)
        btnSave.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnSave)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(430, 13)
        btnReset.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnReset)

        btnExport.Text = "导出Excel"
        btnExport.Location = New Drawing.Point(530, 13)
        btnExport.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnExport)

        ' 汇总信息
        lblSummary.Text = "共 0 条记录"
        lblSummary.Location = New Drawing.Point(650, 18)
        lblSummary.AutoSize = True
        lblSummary.ForeColor = Drawing.Color.Blue
        panelTop.Controls.Add(lblSummary)

        dgvList.Dock = DockStyle.Fill
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadWarehouseList()
        InitGrid()
    End Sub

    Private Sub LoadWarehouseList()
        Try
            Dim shopPermission As String = UserShopPermission
            If String.IsNullOrEmpty(shopPermission) Then shopPermission = "-1"

            Dim sql As String = $"SELECT id, CASE WHEN id = '0' THEN '总库' ELSE title END AS title FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id IN ({shopPermission}) ORDER BY id"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbWarehouse.Items.Clear()
            For Each row As DataRow In dt.Rows
                cmbWarehouse.Items.Add(GBKToUTF8(SafeString(row("title"))))
            Next
            If cmbWarehouse.Items.Count > 0 Then cmbWarehouse.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub InitGrid()
        dgvList.Columns.Clear()
        Dim headers() As String = {"序号", "商品编码", "商品名称", "款号", "品类", "规格", "库存数量", "库存金重", "盘点数量", "盘点金重", "差异(数量)", "差异(金重)", "备注"}
        Dim widths() As Integer = {50, 120, 150, 100, 80, 80, 80, 80, 80, 80, 80, 80, 200}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvList.Columns.Add(col)
        Next

        ' 设置盘点列可编辑，库存列只读
        dgvList.Columns("col0").ReadOnly = True
        dgvList.Columns("col1").ReadOnly = True
        dgvList.Columns("col2").ReadOnly = True
        dgvList.Columns("col3").ReadOnly = True
        dgvList.Columns("col4").ReadOnly = True
        dgvList.Columns("col5").ReadOnly = True
        dgvList.Columns("col6").ReadOnly = True
        dgvList.Columns("col7").ReadOnly = True
        dgvList.Columns("col10").ReadOnly = True
        dgvList.Columns("col11").ReadOnly = True
    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        If cmbWarehouse.SelectedIndex < 0 Then
            ShowWarning("请选择库房！")
            Return
        End If

        Try
            Dim warehouseId As String = cmbWarehouse.SelectedIndex.ToString()
            Dim sql As String = $"SELECT a.poduct_code, a.product_name, a.item_number AS kuanhao, " &
                               $"b.title AS pinlei, c.title AS guige, d.quantity, d.jinzhong " &
                               $"FROM xipunum_erp_shop AS a " &
                               $"INNER JOIN xipunum_erp_shop_kucun AS d ON d.poduct_code = a.poduct_code " &
                               $"LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id " &
                               $"LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id " &
                               $"WHERE d.kufang='{warehouseId}' AND (d.quantity > 0 OR d.jinzhong > 0) ORDER BY a.id"

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            dgvList.Rows.Clear()
            Dim seq As Integer = 1
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    seq,
                    SafeString(row("poduct_code")),
                    GBKToUTF8(SafeString(row("product_name"))),
                    GBKToUTF8(SafeString(row("kuanhao"))),
                    GBKToUTF8(SafeString(row("pinlei"))),
                    GBKToUTF8(SafeString(row("guige"))),
                    SafeDecimal(row("quantity")),
                    SafeDecimal(row("jinzhong")),
                    0, 0, 0, 0, ""
                )
                seq += 1
            Next
            lblSummary.Text = $"共 {dt.Rows.Count} 条记录"
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    ' ========== 单元格值变化事件（自动计算差异） ==========
    Private Sub DgvList_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        If e.ColumnIndex = 8 OrElse e.ColumnIndex = 9 Then ' 盘点数量或盘点金重
            CalculateDiff(e.RowIndex)
        End If
    End Sub

    ' ========== 单元格编辑结束事件 ==========
    Private Sub DgvList_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        If e.ColumnIndex = 8 OrElse e.ColumnIndex = 9 Then
            CalculateDiff(e.RowIndex)
        End If
    End Sub

    ' ========== 计算差异 ==========
    Private Sub CalculateDiff(rowIndex As Integer)
        Dim row As DataGridViewRow = dgvList.Rows(rowIndex)
        Dim stockQty As Decimal = SafeDecimal(row.Cells("col6").Value)
        Dim stockWeight As Decimal = SafeDecimal(row.Cells("col7").Value)
        Dim checkQty As Decimal = SafeDecimal(row.Cells("col8").Value)
        Dim checkWeight As Decimal = SafeDecimal(row.Cells("col9").Value)

        row.Cells("col10").Value = (checkQty - stockQty).ToString("F0")
        row.Cells("col11").Value = (checkWeight - stockWeight).ToString("F3")
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If dgvList.Rows.Count = 0 Then
            ShowWarning("请先加载库存数据！")
            Return
        End If

        If cmbWarehouse.SelectedIndex < 0 Then
            ShowWarning("请选择库房！")
            Return
        End If

        ' 权限检查
        If Not HasOperationPermission("盘点管理") Then
            ShowWarning("没有盘点管理权限！")
            Return
        End If

        Dim trans As MySqlTransaction = Nothing
        Try
            trans = BeginTransaction()
            Dim warehouseId As String = cmbWarehouse.SelectedIndex.ToString()

            For Each row As DataGridViewRow In dgvList.Rows
                Dim productCode As String = SafeString(row.Cells("col1").Value)
                If String.IsNullOrEmpty(productCode) Then Continue For

                Dim checkQty As Decimal = SafeDecimal(row.Cells("col8").Value)
                Dim checkWeight As Decimal = SafeDecimal(row.Cells("col9").Value)
                Dim diffQty As Decimal = SafeDecimal(row.Cells("col10").Value)
                Dim diffWeight As Decimal = SafeDecimal(row.Cells("col11").Value)

                ' 只处理有差异的记录
                If diffQty <> 0 OrElse diffWeight <> 0 Then
                    ' 更新库存为盘点数量
                    Dim sql As String = $"UPDATE xipunum_erp_shop_kucun SET quantity={checkQty}, jinzhong={checkWeight} WHERE poduct_code='{SafeSQL(productCode)}' AND kufang='{warehouseId}'"
                    DatabaseModule.ExecuteCommand(sql, MySQL_Write)

                    ' 记录盘点日志
                    sql = $"INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('盘点', '库存盘点', '{SafeSQL($"商品:{productCode} 数量差异:{diffQty} 金重差异:{diffWeight}")}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
                    DatabaseModule.ExecuteCommand(sql, MySQL_Write)
                End If
            Next

            CommitTransaction(trans)
            ShowSuccess("盘点数据已保存！")

            ' 添加系统日志
            AddSystemLog("盘点", "库存盘点完成", $"库房:{cmbWarehouse.SelectedItem}")

        Catch ex As Exception
            If trans IsNot Nothing Then RollbackTransaction(trans)
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        dgvList.Rows.Clear()
        lblSummary.Text = "共 0 条记录"
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If dgvList.Rows.Count = 0 Then
            ShowWarning("没有数据可导出！")
            Return
        End If

        Try
            Dim dt As New DataTable()
            For Each col As DataGridViewColumn In dgvList.Columns
                dt.Columns.Add(col.HeaderText)
            Next
            For Each row As DataGridViewRow In dgvList.Rows
                Dim dr As DataRow = dt.NewRow()
                For i As Integer = 0 To dgvList.Columns.Count - 1
                    dr(i) = If(row.Cells(i).Value, "")
                Next
                dt.Rows.Add(dr)
            Next
            ExportToExcel(dt, $"盘点数据_{DateTime.Now:yyyyMMddHHmmss}.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub
End Class

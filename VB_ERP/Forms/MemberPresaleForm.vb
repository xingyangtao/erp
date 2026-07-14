' ============================================================================
' 会员预购记录窗口
' 功能: 查看会员预售记录，支持汇总和导出
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class MemberPresaleForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private customerCode As String = ""
    Private lblMemberInfo As New Label()
    Private lblSummary As New Label()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnRefresh As New Button()

    Public Sub New(code As String)
        customerCode = code
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "会员预购记录"
        Me.Size = New Drawing.Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 60
        Me.Controls.Add(panelTop)

        lblMemberInfo.Text = $"会员编码：{customerCode}"
        lblMemberInfo.Location = New Drawing.Point(20, 10)
        lblMemberInfo.AutoSize = True
        lblMemberInfo.Font = New Drawing.Font("微软雅黑", 10, Drawing.FontStyle.Bold)
        panelTop.Controls.Add(lblMemberInfo)

        lblSummary.Text = "共 0 条记录"
        lblSummary.Location = New Drawing.Point(20, 35)
        lblSummary.AutoSize = True
        lblSummary.ForeColor = Drawing.Color.Blue
        panelTop.Controls.Add(lblSummary)

        btnRefresh.Text = "刷新"
        btnRefresh.Location = New Drawing.Point(200, 8)
        btnRefresh.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(btnRefresh)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(270, 8)
        btnExport.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(btnExport)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)

        InitGrid()
    End Sub

    Private Sub InitGrid()
        dgvList.Columns.Clear()
        Dim headers() As String = {"预售单号", "商品名称", "款号", "品类", "数量", "备注", "状态", "导购员", "时间"}
        Dim widths() As Integer = {130, 150, 100, 80, 60, 200, 80, 80, 140}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            col.Width = widths(i)
            dgvList.Columns.Add(col)
        Next
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadMemberInfo()
        LoadData()
    End Sub

    Private Sub LoadMemberInfo()
        Try
            Dim sql As String = $"SELECT name, tel FROM xipunum_erp_member WHERE customer_code='{SafeSQL(customerCode)}' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            If dt.Rows.Count > 0 Then
                lblMemberInfo.Text = $"会员：{GBKToUTF8(SafeString(dt.Rows(0)("name")))}  电话：{SafeString(dt.Rows(0)("tel"))}  编码：{customerCode}"
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim sql As String = $"SELECT b.presale_umber, a.product_name, d.item_number, e.title AS pinlei, " &
                               $"CAST(ROUND(a.quantity, 3) AS DECIMAL(10,3)) AS quantity, " &
                               $"a.remarks, b.state, c.name AS daogou, a.creationtime " &
                               $"FROM xipunum_erp_presale AS a " &
                               $"INNER JOIN xipunum_erp_presale_order AS b ON b.id = a.order_id " &
                               $"LEFT JOIN xipunum_erp_user AS c ON c.user = b.cjuser " &
                               $"LEFT JOIN xipunum_erp_shop AS d ON d.product_name = a.product_name " &
                               $"LEFT JOIN xipunum_erp_category AS e ON e.id = d.category_id " &
                               $"WHERE b.customer_code = '{SafeSQL(customerCode)}' ORDER BY a.id DESC"

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            dgvList.Rows.Clear()

            Dim totalQty As Decimal = 0
            Dim pendingCount As Integer = 0

            For Each row As DataRow In dt.Rows
                Dim qty As Decimal = SafeDecimal(row("quantity"))
                totalQty += qty

                If SafeString(row("state")) = "待处理" Then
                    pendingCount += 1
                End If

                dgvList.Rows.Add(
                    SafeString(row("presale_umber")),
                    GBKToUTF8(SafeString(row("product_name"))),
                    GBKToUTF8(SafeString(row("item_number"))),
                    GBKToUTF8(SafeString(row("pinlei"))),
                    qty.ToString("F3"),
                    GBKToUTF8(SafeString(row("remarks"))),
                    GBKToUTF8(SafeString(row("state"))),
                    GBKToUTF8(SafeString(row("daogou"))),
                    SafeString(row("creationtime"))
                )
            Next

            lblSummary.Text = $"共 {dt.Rows.Count} 条记录，总数量：{totalQty:F3}，待处理：{pendingCount}"
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadData()
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
            ExportToExcel(dt, $"会员预购记录_{customerCode}_{DateTime.Now:yyyyMMddHHmmss}.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub
End Class

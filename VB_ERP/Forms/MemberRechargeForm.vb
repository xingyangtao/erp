' ============================================================================
' 会员充值记录窗口
' 功能: 查看会员充值/存欠记录，支持汇总和导出
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class MemberRechargeForm
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
        Me.Text = "会员充值记录"
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
        Dim headers() As String = {"ID", "存/欠", "类型", "数值", "库房", "备注", "操作账户", "时间"}
        Dim widths() As Integer = {50, 60, 60, 100, 100, 200, 100, 140}

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
            Dim sql As String = $"SELECT a.id, a.cunqu, a.type, CAST(ROUND(a.number, 3) AS DECIMAL(10,3)) AS number, " &
                               $"d.title AS kufang_name, a.remarks, a.cjuser, a.creationtime " &
                               $"FROM xipunum_erp_member_cq AS a " &
                               $"LEFT JOIN xipunum_erp_type AS d ON d.id = a.kufang " &
                               $"WHERE a.customer_code = '{SafeSQL(customerCode)}' ORDER BY a.id DESC"

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            dgvList.Rows.Clear()

            Dim totalSave As Decimal = 0
            Dim totalOwe As Decimal = 0

            For Each row As DataRow In dt.Rows
                Dim number As Decimal = SafeDecimal(row("number"))
                Dim cunqu As String = GBKToUTF8(SafeString(row("cunqu")))

                If cunqu = "存" Then
                    totalSave += number
                Else
                    totalOwe += number
                End If

                dgvList.Rows.Add(
                    SafeString(row("id")),
                    cunqu,
                    GBKToUTF8(SafeString(row("type"))),
                    number.ToString("F3"),
                    GBKToUTF8(SafeString(row("kufang_name"))),
                    GBKToUTF8(SafeString(row("remarks"))),
                    SafeString(row("cjuser")),
                    SafeString(row("creationtime"))
                )
            Next

            lblSummary.Text = $"共 {dt.Rows.Count} 条记录，存：{totalSave:F3}，欠：{totalOwe:F3}，余额：{(totalSave - totalOwe):F3}"
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
            ExportToExcel(dt, $"会员充值记录_{customerCode}_{DateTime.Now:yyyyMMddHHmmss}.xlsx")
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub
End Class

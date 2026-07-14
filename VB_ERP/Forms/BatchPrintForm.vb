' ============================================================================
' 批量打印窗口
' 功能: 批量打印商品标签
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class BatchPrintForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtCode As New TextBox()
    Private WithEvents btnAdd As New Button()
    Private WithEvents btnPrint As New Button()
    Private WithEvents btnClear As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "批量打印"
        Me.Size = New Drawing.Size(800, 500)
        Me.StartPosition = FormStartPosition.CenterParent

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 60
        Me.Controls.Add(panelTop)

        Dim lblCode As New Label()
        lblCode.Text = "商品编码："
        lblCode.Location = New Drawing.Point(20, 20)
        lblCode.AutoSize = True
        panelTop.Controls.Add(lblCode)

        txtCode.Location = New Drawing.Point(90, 17)
        txtCode.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtCode)

        btnAdd.Text = "添加"
        btnAdd.Location = New Drawing.Point(310, 15)
        btnAdd.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnAdd)

        btnPrint.Text = "打印"
        btnPrint.Location = New Drawing.Point(410, 15)
        btnPrint.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnPrint)

        btnClear.Text = "清空"
        btnClear.Location = New Drawing.Point(510, 15)
        btnClear.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnClear)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        InitGrid()
    End Sub

    Private Sub InitGrid()
        dgvList.Columns.Clear()
        dgvList.Columns.Add("bianma", "商品编码")
        dgvList.Columns.Add("mingcheng", "商品名称")
        dgvList.Columns.Add("kuanhao", "款号")
        dgvList.Columns.Add("pinlei", "品类")
        dgvList.Columns.Add("guige", "规格")
        dgvList.Columns.Add("caizhi", "材质")
        dgvList.Columns.Add("danzhong", "单件重")
        dgvList.Columns.Add("jinzhong", "金重")
        dgvList.Columns("bianma").Width = 120
        dgvList.Columns("mingcheng").Width = 150
        dgvList.Columns("kuanhao").Width = 100
        dgvList.Columns("pinlei").Width = 80
        dgvList.Columns("guige").Width = 80
        dgvList.Columns("caizhi").Width = 80
        dgvList.Columns("danzhong").Width = 80
        dgvList.Columns("jinzhong").Width = 80
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If String.IsNullOrEmpty(txtCode.Text.Trim()) Then
            ShowWarning("请输入商品编码！")
            Return
        End If

        Try
            Dim sql As String = $"SELECT a.poduct_code, a.product_name, a.item_number, b.title AS pinlei, c.title AS guige, a.caizhi, a.single FROM xipunum_erp_shop AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id WHERE a.poduct_code='{SafeSQL(txtCode.Text.Trim())}' OR a.fu_code='{SafeSQL(txtCode.Text.Trim())}'"
            Dim dt As DataTable = ExecuteQuery(sql)

            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                dgvList.Rows.Add(
                    SafeString(row("poduct_code")),
                    SafeString(row("product_name")),
                    SafeString(row("item_number")),
                    SafeString(row("pinlei")),
                    SafeString(row("guige")),
                    SafeString(row("caizhi")),
                    SafeDecimal(row("single")),
                    0
                )
                txtCode.Text = ""
                txtCode.Focus()
            Else
                ShowWarning("未找到该商品！")
            End If
        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        If dgvList.Rows.Count = 0 Then
            ShowWarning("请先添加商品！")
            Return
        End If
        ShowSuccess("打印功能开发中...")
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        dgvList.Rows.Clear()
    End Sub
End Class

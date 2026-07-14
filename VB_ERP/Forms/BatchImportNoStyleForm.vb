' ============================================================================
' 批量导入无款窗口
' 功能: 批量导入商品数据（无款号）
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class BatchImportNoStyleForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private WithEvents btnImport As New Button()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnReset As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "批量导入无款"
        Me.Size = New Drawing.Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 50
        Me.Controls.Add(panelTop)

        btnImport.Text = "导入Excel"
        btnImport.Location = New Drawing.Point(20, 10)
        btnImport.Size = New Drawing.Size(100, 30)
        panelTop.Controls.Add(btnImport)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(140, 10)
        btnSave.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnSave)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(240, 10)
        btnReset.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnReset)

        dgvList.Dock = DockStyle.Fill
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)

        InitGrid()
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 初始化
    End Sub

    Private Sub InitGrid()
        dgvList.Columns.Clear()
        dgvList.Columns.Add("xuhao", "序号")
        dgvList.Columns.Add("poduct_code", "商品编码")
        dgvList.Columns.Add("product_name", "商品名称")
        dgvList.Columns.Add("pinlei", "品类")
        dgvList.Columns.Add("guige", "规格")
        dgvList.Columns.Add("kuanhao", "款号")
        dgvList.Columns.Add("caizhi", "材质")
        dgvList.Columns.Add("xiangqian", "镶嵌")
        dgvList.Columns.Add("quankou", "圈口")
        dgvList.Columns.Add("miankuan", "面宽")
        dgvList.Columns.Add("houdu", "厚度")
        dgvList.Columns.Add("danwei", "单位")
        dgvList.Columns.Add("danzhong", "单件重")
        dgvList.Columns.Add("shuliang", "数量")
        dgvList.Columns.Add("jinzhong", "金重")
        dgvList.Columns.Add("zhongliang", "重量")
        dgvList.Columns.Add("sunhao", "损耗")
        dgvList.Columns.Add("chengbenjia", "成本价")
        dgvList.Columns.Add("beizhu", "备注")

        dgvList.Columns("xuhao").Width = 50
        dgvList.Columns("poduct_code").Width = 120
        dgvList.Columns("product_name").Width = 150
        dgvList.Columns("pinlei").Width = 80
        dgvList.Columns("guige").Width = 80
        dgvList.Columns("kuanhao").Width = 100
        dgvList.Columns("caizhi").Width = 80
        dgvList.Columns("xiangqian").Width = 60
        dgvList.Columns("quankou").Width = 60
        dgvList.Columns("miankuan").Width = 60
        dgvList.Columns("houdu").Width = 60
        dgvList.Columns("danwei").Width = 50
        dgvList.Columns("danzhong").Width = 70
        dgvList.Columns("shuliang").Width = 50
        dgvList.Columns("jinzhong").Width = 70
        dgvList.Columns("zhongliang").Width = 70
        dgvList.Columns("sunhao").Width = 50
        dgvList.Columns("chengbenjia").Width = 80
        dgvList.Columns("beizhu").Width = 150
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        ShowSuccess("导入功能开发中...")
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If dgvList.Rows.Count = 0 Then
            ShowWarning("没有数据可保存！")
            Return
        End If

        ShowSuccess("保存功能开发中...")
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        dgvList.Rows.Clear()
    End Sub
End Class

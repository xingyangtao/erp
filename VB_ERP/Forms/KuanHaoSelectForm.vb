' ============================================================================
' 款号选择窗口
' 功能: 选择款号
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class KuanHaoSelectForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtSearch As New TextBox()
    Private WithEvents btnSearch As New Button()
    Private WithEvents btnSelect As New Button()
    Private WithEvents btnCancel As New Button()

    Public Property SelectedKuanHao As String = ""

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "选择款号"
        Me.Size = New Drawing.Size(800, 500)
        Me.StartPosition = FormStartPosition.CenterParent

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 50
        Me.Controls.Add(panelTop)

        Dim lblSearch As New Label()
        lblSearch.Text = "搜索："
        lblSearch.Location = New Drawing.Point(20, 15)
        lblSearch.AutoSize = True
        panelTop.Controls.Add(lblSearch)

        txtSearch.Location = New Drawing.Point(60, 12)
        txtSearch.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtSearch)

        btnSearch.Text = "搜索"
        btnSearch.Location = New Drawing.Point(280, 10)
        btnSearch.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnSearch)

        btnSelect.Text = "选择"
        btnSelect.Location = New Drawing.Point(380, 10)
        btnSelect.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnSelect)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(470, 10)
        btnCancel.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnCancel)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim sql As String = "SELECT a.id, a.kuanhao, a.title, a.caizhi, b.title AS pinlei, c.title AS guige FROM xipunum_erp_ksiamges AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id ORDER BY a.id DESC"
            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try
            Dim searchFilter As String = ""
            If Not String.IsNullOrEmpty(txtSearch.Text.Trim()) Then
                searchFilter = $" WHERE a.title LIKE '%{SafeSQL(txtSearch.Text)}%' OR a.kuanhao LIKE '%{SafeSQL(txtSearch.Text)}%'"
            End If

            Dim sql As String = $"SELECT a.id, a.kuanhao, a.title, a.caizhi, b.title AS pinlei, c.title AS guige FROM xipunum_erp_ksiamges AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id {searchFilter} ORDER BY a.id DESC"
            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择款号！")
            Return
        End If

        SelectedKuanHao = SafeString(dgvList.SelectedRows(0).Cells("kuanhao").Value)
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

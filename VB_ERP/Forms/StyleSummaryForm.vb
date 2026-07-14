' ============================================================================
' 款式数据汇总窗口
' 功能: 款号数据汇总分析
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class StyleSummaryForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtSearch As New TextBox()
    Private cmbCategory As New ComboBox()
    Private WithEvents btnSearch As New Button()
    Private WithEvents btnRefresh As New Button()
    Private WithEvents btnDetail As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "款式数据汇总"
        Me.Size = New Drawing.Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 60
        Me.Controls.Add(panelTop)

        Dim lblSearch As New Label()
        lblSearch.Text = "搜索："
        lblSearch.Location = New Drawing.Point(20, 18)
        lblSearch.AutoSize = True
        panelTop.Controls.Add(lblSearch)

        txtSearch.Location = New Drawing.Point(60, 15)
        txtSearch.Size = New Drawing.Size(200, 25)
        panelTop.Controls.Add(txtSearch)

        Dim lblCategory As New Label()
        lblCategory.Text = "品类："
        lblCategory.Location = New Drawing.Point(280, 18)
        lblCategory.AutoSize = True
        panelTop.Controls.Add(lblCategory)

        cmbCategory.Location = New Drawing.Point(320, 15)
        cmbCategory.Size = New Drawing.Size(120, 25)
        panelTop.Controls.Add(cmbCategory)

        btnSearch.Text = "查询"
        btnSearch.Location = New Drawing.Point(460, 13)
        btnSearch.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnSearch)

        btnDetail.Text = "详情"
        btnDetail.Location = New Drawing.Point(550, 13)
        btnDetail.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnDetail)

        btnRefresh.Text = "刷新"
        btnRefresh.Location = New Drawing.Point(640, 13)
        btnRefresh.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnRefresh)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadCategoryList()
        LoadData()
    End Sub

    Private Sub LoadCategoryList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, title FROM xipunum_erp_category ORDER BY id")
            cmbCategory.Items.Clear()
            cmbCategory.Items.Add("全部品类")
            For Each row As DataRow In dt.Rows
                cmbCategory.Items.Add(SafeString(row("title")))
            Next
            cmbCategory.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim sql As String = "SELECT a.id, a.kuanhao, a.title, a.caizhi, a.chengse, a.lingxiao, b.title AS pinlei, c.title AS guige FROM xipunum_erp_ksiamges AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id WHERE 1=1"

            If Not String.IsNullOrEmpty(txtSearch.Text.Trim()) Then
                sql &= $" AND (a.title LIKE '%{SafeSQL(txtSearch.Text)}%' OR a.kuanhao LIKE '%{SafeSQL(txtSearch.Text)}%')"
            End If

            sql &= " ORDER BY a.id DESC"

            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        LoadData()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadData()
    End Sub

    Private Sub btnDetail_Click(sender As Object, e As EventArgs) Handles btnDetail.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要查看的记录！")
            Return
        End If

        Dim id As String = SafeString(dgvList.SelectedRows(0).Cells("id").Value)
        Dim form As New StyleSummaryDetailForm(id)
        form.ShowDialog()
    End Sub
End Class

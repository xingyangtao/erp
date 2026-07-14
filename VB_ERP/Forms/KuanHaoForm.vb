' ============================================================================
' 款号管理窗口
' 功能: 款号信息管理，支持图片识别、分页加载
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class KuanHaoForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private cmbCategory As New ComboBox()
    Private cmbSpec As New ComboBox()
    Private txtSearch As New TextBox()
    Private WithEvents btnSearch As New Button()
    Private WithEvents btnAdd As New Button()
    Private WithEvents btnEdit As New Button()
    Private WithEvents btnDelete As New Button()
    Private WithEvents btnRefresh As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "款号管理"
        Me.Size = New Drawing.Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterParent

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 80
        Me.Controls.Add(panelTop)

        Dim lblCategory As New Label()
        lblCategory.Text = "品类："
        lblCategory.Location = New Drawing.Point(20, 15)
        lblCategory.AutoSize = True
        panelTop.Controls.Add(lblCategory)

        cmbCategory.Location = New Drawing.Point(60, 12)
        cmbCategory.Size = New Drawing.Size(120, 25)
        panelTop.Controls.Add(cmbCategory)

        Dim lblSpec As New Label()
        lblSpec.Text = "规格："
        lblSpec.Location = New Drawing.Point(200, 15)
        lblSpec.AutoSize = True
        panelTop.Controls.Add(lblSpec)

        cmbSpec.Location = New Drawing.Point(240, 12)
        cmbSpec.Size = New Drawing.Size(120, 25)
        panelTop.Controls.Add(cmbSpec)

        Dim lblSearch As New Label()
        lblSearch.Text = "搜索："
        lblSearch.Location = New Drawing.Point(380, 15)
        lblSearch.AutoSize = True
        panelTop.Controls.Add(lblSearch)

        txtSearch.Location = New Drawing.Point(420, 12)
        txtSearch.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtSearch)

        btnSearch.Text = "查询"
        btnSearch.Location = New Drawing.Point(590, 10)
        btnSearch.Size = New Drawing.Size(70, 28)
        panelTop.Controls.Add(btnSearch)

        btnRefresh.Text = "刷新"
        btnRefresh.Location = New Drawing.Point(680, 10)
        btnRefresh.Size = New Drawing.Size(70, 28)
        panelTop.Controls.Add(btnRefresh)

        btnAdd.Text = "添加"
        btnAdd.Location = New Drawing.Point(20, 45)
        btnAdd.Size = New Drawing.Size(70, 28)
        panelTop.Controls.Add(btnAdd)

        btnEdit.Text = "编辑"
        btnEdit.Location = New Drawing.Point(110, 45)
        btnEdit.Size = New Drawing.Size(70, 28)
        panelTop.Controls.Add(btnEdit)

        btnDelete.Text = "删除"
        btnDelete.Location = New Drawing.Point(200, 45)
        btnDelete.Size = New Drawing.Size(70, 28)
        panelTop.Controls.Add(btnDelete)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadCategoryList()
        LoadSpecList()
        LoadData()
    End Sub

    Private Sub LoadCategoryList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, title FROM xipunum_erp_category ORDER BY id")
            cmbCategory.Items.Clear()
            cmbCategory.Items.Add(New ComboBoxItem("", "全部品类"))
            For Each row As DataRow In dt.Rows
                cmbCategory.Items.Add(New ComboBoxItem(SafeString(row("id")), SafeString(row("title"))))
            Next
            cmbCategory.SelectedIndex = 0
        Catch ex As Exception
            ShowError("加载品类列表失败：" & ex.Message)
        End Try
    End Sub

    Private Sub LoadSpecList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, title FROM xipunum_erp_specs ORDER BY id")
            cmbSpec.Items.Clear()
            cmbSpec.Items.Add(New ComboBoxItem("", "全部规格"))
            For Each row As DataRow In dt.Rows
                cmbSpec.Items.Add(New ComboBoxItem(SafeString(row("id")), SafeString(row("title"))))
            Next
            cmbSpec.SelectedIndex = 0
        Catch ex As Exception
            ShowError("加载规格列表失败：" & ex.Message)
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim sql As String = "SELECT a.id, a.title, a.kuanhao, a.caizhi, a.chengse, a.lingxiao, b.title AS pinlei, c.title AS guige FROM xipunum_erp_ksiamges AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id WHERE 1=1"

            If cmbCategory.SelectedIndex > 0 Then
                Dim categoryId As String = DirectCast(cmbCategory.SelectedItem, ComboBoxItem).ID
                sql &= $" AND a.category_id='{categoryId}'"
            End If

            If cmbSpec.SelectedIndex > 0 Then
                Dim specId As String = DirectCast(cmbSpec.SelectedItem, ComboBoxItem).ID
                sql &= $" AND a.specification_id='{specId}'"
            End If

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

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim form As New KuanHaoAddForm()
        If form.ShowDialog() = DialogResult.OK Then
            LoadData()
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要编辑的记录！")
            Return
        End If
        Dim id As String = SafeString(dgvList.SelectedRows(0).Cells("id").Value)
        Dim form As New KuanHaoAddForm(id)
        If form.ShowDialog() = DialogResult.OK Then
            LoadData()
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要删除的记录！")
            Return
        End If

        If Not ConfirmAction("确定要删除吗？") Then Return

        Try
            For Each row As DataGridViewRow In dgvList.SelectedRows
                Dim id As String = SafeString(row.Cells("id").Value)
                ExecuteCommand($"DELETE FROM xipunum_erp_ksiamges WHERE id='{id}'")
            Next
            ShowSuccess("删除成功！")
            LoadData()
        Catch ex As Exception
            ShowError("删除失败：" & ex.Message)
        End Try
    End Sub

    Private Class ComboBoxItem
        Public Property ID As String
        Public Property Text As String
        Public Sub New(id As String, text As String)
            Me.ID = id
            Me.Text = text
        End Sub
        Public Overrides Function ToString() As String
            Return Text
        End Function
    End Class
End Class

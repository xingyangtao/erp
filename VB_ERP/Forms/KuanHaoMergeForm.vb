' ============================================================================
' 款号合并窗口
' 功能: 款号合并操作
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class KuanHaoMergeForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private WithEvents btnMerge As New Button()
    Private WithEvents btnRefresh As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "款号合并"
        Me.Size = New Drawing.Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 50
        Me.Controls.Add(panelTop)

        btnRefresh.Text = "刷新"
        btnRefresh.Location = New Drawing.Point(20, 10)
        btnRefresh.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnRefresh)

        btnMerge.Text = "合并"
        btnMerge.Location = New Drawing.Point(120, 10)
        btnMerge.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnMerge)

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

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadData()
    End Sub

    Private Sub btnMerge_Click(sender As Object, e As EventArgs) Handles btnMerge.Click
        If dgvList.SelectedRows.Count < 2 Then
            ShowWarning("请先选择要合并的款号（至少2个）！")
            Return
        End If

        If Not ConfirmAction("确定要合并吗？第一个选中的款号将作为主款号。") Then Return

        Try
            Dim mainId As String = SafeString(dgvList.SelectedRows(0).Cells("id").Value)
            Dim mainKuanhao As String = SafeString(dgvList.SelectedRows(0).Cells("kuanhao").Value)

            For i As Integer = 1 To dgvList.SelectedRows.Count - 1
                Dim mergeId As String = SafeString(dgvList.SelectedRows(i).Cells("id").Value)
                Dim mergeKuanhao As String = SafeString(dgvList.SelectedRows(i).Cells("kuanhao").Value)

                ' 更新商品表
                ExecuteCommand($"UPDATE xipunum_erp_shop SET item_number='{mainKuanhao}' WHERE item_number='{mergeKuanhao}'")

                ' 更新入库表
                ExecuteCommand($"UPDATE xipunum_erp_store SET item_number='{mainKuanhao}' WHERE item_number='{mergeKuanhao}'")

                ' 删除合并款号
                ExecuteCommand($"DELETE FROM xipunum_erp_ksiamges WHERE id='{mergeId}'")
            Next

            ShowSuccess("合并成功！")
            LoadData()
        Catch ex As Exception
            ShowError("合并失败：" & ex.Message)
        End Try
    End Sub
End Class

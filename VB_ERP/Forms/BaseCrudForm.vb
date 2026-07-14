' ============================================================================
' 通用CRUD管理窗口基类
' 用于基础设置模块的简单增删改查窗口
' ============================================================================

Public Class BaseCrudForm
    Inherits System.Windows.Forms.Form

    Protected dgvList As New DataGridView()
    Protected WithEvents btnAdd As New Button()
    Protected WithEvents btnEdit As New Button()
    Protected WithEvents btnDelete As New Button()
    Protected WithEvents btnRefresh As New Button()

    Protected Overridable Sub InitializeBaseUI(title As String)
        Me.Text = title
        Me.Size = New Drawing.Size(800, 500)
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

        btnAdd.Text = "添加"
        btnAdd.Location = New Drawing.Point(120, 10)
        btnAdd.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnAdd)

        btnEdit.Text = "编辑"
        btnEdit.Location = New Drawing.Point(220, 10)
        btnEdit.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnEdit)

        btnDelete.Text = "删除"
        btnDelete.Location = New Drawing.Point(320, 10)
        btnDelete.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnDelete)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Protected Sub InitGrid(headers() As String, widths() As Integer)
        dgvList.Columns.Clear()
        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            If i < widths.Length Then
                col.Width = widths(i)
            Else
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End If
            dgvList.Columns.Add(col)
        Next
    End Sub

    Protected Overridable Sub LoadData()
        ' Override in subclass
    End Sub

    Protected Overridable Sub AddRecord()
        ' Override in subclass
    End Sub

    Protected Overridable Sub EditRecord()
        ' Override in subclass
    End Sub

    Protected Overridable Sub DeleteRecord()
        ' Override in subclass
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadData()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        AddRecord()
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        EditRecord()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        DeleteRecord()
    End Sub
End Class

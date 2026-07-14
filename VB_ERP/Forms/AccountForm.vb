' ============================================================================
' 账户管理窗口
' 功能: 用户新增/编辑/删除/密码初始化
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class AccountForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private txtUsername As New TextBox()
    Private txtName As New TextBox()
    Private txtPassword As New TextBox()
    Private txtPhone As New TextBox()
    Private txtEmail As New TextBox()
    Private cmbDepartment As New ComboBox()
    Private cmbPost As New ComboBox()
    Private cmbType As New ComboBox()
    Private cmbState As New ComboBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnDelete As New Button()
    Private WithEvents btnReset As New Button()
    Private WithEvents btnInitPassword As New Button()
    Private WithEvents btnSearch As New Button()
    Private txtSearch As New TextBox()
    Private currentId As String = ""

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "账户管理"
        Me.Size = New Drawing.Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 180
        Me.Controls.Add(panelTop)

        ' 工具条
        Dim toolStrip As New ToolStrip()
        toolStrip.Dock = DockStyle.Top
        toolStrip.Height = 40
        panelTop.Controls.Add(toolStrip)

        AddToolButton(toolStrip, "保存")
        AddToolButton(toolStrip, "删除")
        AddToolButton(toolStrip, "重置")
        AddToolButton(toolStrip, "密码初始化")

        ' 搜索
        txtSearch.Location = New Drawing.Point(300, 10)
        txtSearch.Size = New Drawing.Size(150, 25)
        txtSearch.Text = "输入用户名/姓名搜索"
        AddHandler txtSearch.GotFocus, Sub() If txtSearch.Text = "输入用户名/姓名搜索" Then txtSearch.Text = ""
        AddHandler txtSearch.LostFocus, Sub() If String.IsNullOrEmpty(txtSearch.Text) Then txtSearch.Text = "输入用户名/姓名搜索"
        panelTop.Controls.Add(txtSearch)

        btnSearch.Text = "搜索"
        btnSearch.Location = New Drawing.Point(460, 10)
        btnSearch.Size = New Drawing.Size(60, 25)
        panelTop.Controls.Add(btnSearch)

        ' 第一行
        AddLabel(panelTop, "用户名：", 20, 50)
        txtUsername.Location = New Drawing.Point(80, 47)
        txtUsername.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtUsername)

        AddLabel(panelTop, "姓名：", 250, 50)
        txtName.Location = New Drawing.Point(300, 47)
        txtName.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtName)

        AddLabel(panelTop, "密码：", 470, 50)
        txtPassword.Location = New Drawing.Point(520, 47)
        txtPassword.Size = New Drawing.Size(150, 25)
        txtPassword.UseSystemPasswordChar = True
        panelTop.Controls.Add(txtPassword)

        ' 第二行
        AddLabel(panelTop, "电话：", 20, 85)
        txtPhone.Location = New Drawing.Point(80, 82)
        txtPhone.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtPhone)

        AddLabel(panelTop, "邮箱：", 250, 85)
        txtEmail.Location = New Drawing.Point(300, 82)
        txtEmail.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(txtEmail)

        AddLabel(panelTop, "类型：", 470, 85)
        cmbType.Location = New Drawing.Point(520, 82)
        cmbType.Size = New Drawing.Size(100, 25)
        cmbType.Items.AddRange(New String() {"商铺", "后台"})
        cmbType.SelectedIndex = 0
        panelTop.Controls.Add(cmbType)

        AddLabel(panelTop, "状态：", 640, 85)
        cmbState.Location = New Drawing.Point(690, 82)
        cmbState.Size = New Drawing.Size(100, 25)
        cmbState.Items.AddRange(New String() {"正常", "停用"})
        cmbState.SelectedIndex = 0
        panelTop.Controls.Add(cmbState)

        ' 第三行
        AddLabel(panelTop, "分组：", 20, 120)
        cmbDepartment.Location = New Drawing.Point(80, 117)
        cmbDepartment.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(cmbDepartment)

        AddLabel(panelTop, "岗位：", 250, 120)
        cmbPost.Location = New Drawing.Point(300, 117)
        cmbPost.Size = New Drawing.Size(150, 25)
        panelTop.Controls.Add(cmbPost)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        AddHandler dgvList.CellClick, AddressOf dgvList_CellClick
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub AddToolButton(toolStrip As ToolStrip, text As String)
        Dim btn As New ToolStripButton(text)
        btn.DisplayStyle = ToolStripItemDisplayStyle.Text
        AddHandler btn.Click, AddressOf ToolButton_Click
        toolStrip.Items.Add(btn)
    End Sub

    Private Sub AddLabel(parent As Panel, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadDepartmentList()
        LoadPostList()
        LoadData()
    End Sub

    Private Sub LoadDepartmentList()
        Try
            Dim dt As DataTable = DatabaseModule.ExecuteQuery("SELECT id, title FROM xipunum_erp_type WHERE type='商铺' AND superior='0' ORDER BY id")
            cmbDepartment.Items.Clear()
            cmbDepartment.Items.Add(New ComboBoxItem("", "请选择"))
            For Each row As DataRow In dt.Rows
                cmbDepartment.Items.Add(New ComboBoxItem(SafeString(row("id")), GBKToUTF8(SafeString(row("title")))))
            Next
            If cmbDepartment.Items.Count > 0 Then cmbDepartment.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadPostList()
        Try
            Dim dt As DataTable = DatabaseModule.ExecuteQuery("SELECT id, title FROM xipunum_erp_type WHERE superior>'0' AND type='商铺' ORDER BY id")
            cmbPost.Items.Clear()
            cmbPost.Items.Add(New ComboBoxItem("", "请选择"))
            For Each row As DataRow In dt.Rows
                cmbPost.Items.Add(New ComboBoxItem(SafeString(row("id")), GBKToUTF8(SafeString(row("title")))))
            Next
            If cmbPost.Items.Count > 0 Then cmbPost.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim searchFilter As String = ""
            If Not String.IsNullOrEmpty(txtSearch.Text) AndAlso txtSearch.Text <> "输入用户名/姓名搜索" Then
                searchFilter = $" AND (a.user LIKE '%{SafeSQL(txtSearch.Text)}%' OR a.name LIKE '%{SafeSQL(txtSearch.Text)}%')"
            End If

            Dim sql As String = $"SELECT a.id, a.user, a.name, a.tel, a.mailbox, b.title AS department, c.title AS post, a.type, a.state, a.creationtime " &
                               $"FROM xipunum_erp_user AS a " &
                               $"LEFT JOIN xipunum_erp_type AS b ON b.id = a.department " &
                               $"LEFT JOIN xipunum_erp_type AS c ON c.id = a.post " &
                               $"WHERE 1=1 {searchFilter} ORDER BY a.id"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)

            dgvList.Rows.Clear()
            For Each row As DataRow In dt.Rows
                dgvList.Rows.Add(
                    SafeString(row("id")),
                    SafeString(row("user")),
                    GBKToUTF8(SafeString(row("name"))),
                    SafeString(row("tel")),
                    SafeString(row("mailbox")),
                    GBKToUTF8(SafeString(row("department"))),
                    GBKToUTF8(SafeString(row("post"))),
                    SafeString(row("type")),
                    If(SafeString(row("state")) = "0", "正常", "停用"),
                    SafeString(row("creationtime"))
                )
            Next
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub dgvList_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        Dim row As DataGridViewRow = dgvList.Rows(e.RowIndex)
        currentId = SafeString(row.Cells("col0").Value)
        txtUsername.Text = SafeString(row.Cells("col1").Value)
        txtName.Text = SafeString(row.Cells("col2").Value)
        txtPhone.Text = SafeString(row.Cells("col3").Value)
        txtEmail.Text = SafeString(row.Cells("col4").Value)
        txtUsername.ReadOnly = True
        txtPassword.Text = ""
    End Sub

    Private Sub ToolButton_Click(sender As Object, e As EventArgs)
        Dim btn As ToolStripButton = DirectCast(sender, ToolStripButton)
        Select Case btn.Text
            Case "保存"
                SaveUser()
            Case "删除"
                DeleteUser()
            Case "重置"
                ResetForm()
            Case "密码初始化"
                InitPassword()
        End Select
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        LoadData()
    End Sub

    Private Sub SaveUser()
        If String.IsNullOrEmpty(txtUsername.Text.Trim()) Then
            ShowWarning("请输入用户名！")
            Return
        End If
        If String.IsNullOrEmpty(txtName.Text.Trim()) Then
            ShowWarning("请输入姓名！")
            Return
        End If

        ' 权限检查
        If Not HasOperationPermission("账户管理") Then
            ShowWarning("没有账户管理权限！")
            Return
        End If

        Try
            Dim departmentId As String = ""
            If cmbDepartment.SelectedIndex > 0 Then
                departmentId = DirectCast(cmbDepartment.SelectedItem, ComboBoxItem).ID
            End If
            Dim postId As String = ""
            If cmbPost.SelectedIndex > 0 Then
                postId = DirectCast(cmbPost.SelectedItem, ComboBoxItem).ID
            End If
            Dim state As String = If(cmbState.SelectedIndex = 1, "1", "0")

            Dim sql As String
            If String.IsNullOrEmpty(currentId) Then
                ' 新增
                If String.IsNullOrEmpty(txtPassword.Text) Then
                    ShowWarning("请输入密码！")
                    Return
                End If
                Dim encrypted As String = MD5Encrypt(txtPassword.Text, AuthDataNum)
                sql = $"INSERT INTO xipunum_erp_user (user, password, name, tel, mailbox, department, post, type, state, cjuser, creationtime) VALUES ('{SafeSQL(txtUsername.Text)}', '{SafeSQL(encrypted)}', '{SafeSQL(txtName.Text)}', '{SafeSQL(txtPhone.Text)}', '{SafeSQL(txtEmail.Text)}', '{departmentId}', '{postId}', '{cmbType.SelectedItem}', '{state}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
            Else
                ' 编辑
                sql = $"UPDATE xipunum_erp_user SET name='{SafeSQL(txtName.Text)}', tel='{SafeSQL(txtPhone.Text)}', mailbox='{SafeSQL(txtEmail.Text)}', department='{departmentId}', post='{postId}', type='{cmbType.SelectedItem}', state='{state}', updatetime='{GetOperationDate()}' WHERE id='{currentId}'"
                If Not String.IsNullOrEmpty(txtPassword.Text) Then
                    Dim encrypted As String = MD5Encrypt(txtPassword.Text, AuthDataNum)
                    sql = $"UPDATE xipunum_erp_user SET name='{SafeSQL(txtName.Text)}', password='{SafeSQL(encrypted)}', tel='{SafeSQL(txtPhone.Text)}', mailbox='{SafeSQL(txtEmail.Text)}', department='{departmentId}', post='{postId}', type='{cmbType.SelectedItem}', state='{state}', updatetime='{GetOperationDate()}' WHERE id='{currentId}'"
                End If
            End If

            DatabaseModule.ExecuteCommand(sql)
            AddSystemLog(If(String.IsNullOrEmpty(currentId), "添加", "修改"), If(String.IsNullOrEmpty(currentId), "添加账户", "修改账户"), "账户：" & txtUsername.Text)
            ShowSuccess("保存成功！")
            LoadData()
            ResetForm()
        Catch ex As Exception
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub

    Private Sub DeleteUser()
        If String.IsNullOrEmpty(currentId) Then
            ShowWarning("请先选择要删除的记录！")
            Return
        End If

        ' 权限检查
        If Not HasOperationPermission("账户管理") Then
            ShowWarning("没有账户管理权限！")
            Return
        End If

        If txtUsername.Text = "admin" Then
            ShowWarning("不能删除管理员账户！")
            Return
        End If

        If Not ConfirmAction("确定要删除该账户吗？") Then Return

        Try
            DatabaseModule.ExecuteCommand($"DELETE FROM xipunum_erp_user WHERE id='{currentId}'")
            AddSystemLog("删除", "删除账户", "账户：" & txtUsername.Text)
            ShowSuccess("删除成功！")
            LoadData()
            ResetForm()
        Catch ex As Exception
            ShowError("删除失败：" & ex.Message)
        End Try
    End Sub

    Private Sub InitPassword()
        If String.IsNullOrEmpty(currentId) Then
            ShowWarning("请先选择要初始化密码的账户！")
            Return
        End If

        ' 权限检查
        If Not HasOperationPermission("密码初始化") Then
            ShowWarning("没有密码初始化权限！")
            Return
        End If

        If Not ConfirmAction($"确定要初始化 {txtUsername.Text} 的密码吗？") Then Return

        Try
            Dim defaultPassword As String = "123456"
            Dim encrypted As String = MD5Encrypt(defaultPassword, AuthDataNum)
            DatabaseModule.ExecuteCommand($"UPDATE xipunum_erp_user SET password='{SafeSQL(encrypted)}', login='0', updatetime='{GetOperationDate()}' WHERE id='{currentId}'")
            AddSystemLog("初始化", "初始化密码", "账户：" & txtUsername.Text)
            ShowSuccess("密码初始化成功！默认密码为：123456")
        Catch ex As Exception
            ShowError("初始化失败：" & ex.Message)
        End Try
    End Sub

    Private Sub ResetForm()
        currentId = ""
        txtUsername.Text = ""
        txtUsername.ReadOnly = False
        txtName.Text = ""
        txtPassword.Text = ""
        txtPhone.Text = ""
        txtEmail.Text = ""
        If cmbDepartment.Items.Count > 0 Then cmbDepartment.SelectedIndex = 0
        If cmbPost.Items.Count > 0 Then cmbPost.SelectedIndex = 0
        cmbType.SelectedIndex = 0
        cmbState.SelectedIndex = 0
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

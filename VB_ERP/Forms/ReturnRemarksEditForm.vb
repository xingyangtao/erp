' ============================================================================
' 商品退库备注修改窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品退库备注修改.form.e.txt
' 包含所有5个子程序、所有SQL查询
' 功能：修改退库单的备注，记录系统日志，刷新主窗口
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class ReturnRemarksEditForm
    Inherits System.Windows.Forms.Form

    ' ========== 控件声明 ==========
    Private grpMain As New GroupBox()                    ' 添加修改_分组框
    Private txtOrderNumber As New TextBox()              ' 编辑框_退库单号
    Private txtRemarks As New TextBox()                  ' 编辑框_备注
    Private lblOrderNumber As New Label()                 ' 退库单号标签
    Private lblRemarks As New Label()                     ' 备注标签
    Private btnSave As New Button()                      ' 按钮EX1（保存）
    Private btnReset As New Button()                     ' 按钮EX2（重置）
    Private btnClose As New Button()                     ' 图片框EX4（关闭）

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnReset.Click, AddressOf BtnReset_Click
        AddHandler btnClose.Click, AddressOf BtnClose_Click
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "退库备注修改"
        Me.Size = New Drawing.Size(449, 355)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False

        ' 添加修改_分组框
        grpMain.Text = ""
        grpMain.Size = New Drawing.Size(400, 260)
        grpMain.Location = New Drawing.Point(24, 47)
        Me.Controls.Add(grpMain)

        ' 退库单号标签
        lblOrderNumber.Text = "退库单号"
        lblOrderNumber.AutoSize = False
        lblOrderNumber.Size = New Drawing.Size(72, 30)
        lblOrderNumber.Location = New Drawing.Point(16, 48)
        lblOrderNumber.TextAlign = Drawing.ContentAlignment.MiddleLeft
        lblOrderNumber.Font = New Drawing.Font("微软雅黑", 9)
        grpMain.Controls.Add(lblOrderNumber)

        ' 退库单号编辑框
        txtOrderNumber.Location = New Drawing.Point(88, 48)
        txtOrderNumber.Size = New Drawing.Size(224, 30)
        txtOrderNumber.ReadOnly = True
        grpMain.Controls.Add(txtOrderNumber)

        ' 备注标签
        lblRemarks.Text = "备注"
        lblRemarks.AutoSize = False
        lblRemarks.Size = New Drawing.Size(72, 125)
        lblRemarks.Location = New Drawing.Point(16, 83)
        lblRemarks.TextAlign = Drawing.ContentAlignment.MiddleLeft
        lblRemarks.Font = New Drawing.Font("微软雅黑", 9)
        grpMain.Controls.Add(lblRemarks)

        ' 备注编辑框
        txtRemarks.Location = New Drawing.Point(88, 83)
        txtRemarks.Size = New Drawing.Size(224, 125)
        txtRemarks.Multiline = True
        txtRemarks.ScrollBars = ScrollBars.Vertical
        grpMain.Controls.Add(txtRemarks)

        ' 保存按钮
        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(200, 14)
        btnSave.Size = New Drawing.Size(72, 28)
        btnSave.BackColor = Drawing.Color.FromArgb(170, 186, 22)
        btnSave.ForeColor = Drawing.Color.White
        btnSave.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnSave)

        ' 重置按钮
        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(280, 14)
        btnReset.Size = New Drawing.Size(72, 28)
        btnReset.BackColor = Drawing.Color.FromArgb(239, 236, 233)
        btnReset.ForeColor = Drawing.Color.Black
        btnReset.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnReset)

        ' 关闭按钮
        btnClose.Text = "×"
        btnClose.Location = New Drawing.Point(360, 14)
        btnClose.Size = New Drawing.Size(28, 28)
        btnClose.Font = New Drawing.Font("宋体", 12)
        btnClose.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnClose)
    End Sub

    ' ========== 窗口加载（_窗口_商品退库备注修改_创建完毕） ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 获取主窗口选中行
        Dim orderSelected As Integer = MainForm.ReturnOrderButtonName

        ' 设置退库单号
        txtOrderNumber.Text = ReturnOrderNumber

        ' 设置备注（从主窗口表格列7获取）
        If orderSelected >= 0 AndAlso orderSelected < MainForm.dgvMain.Rows.Count Then
            txtRemarks.Text = SafeString(MainForm.dgvMain.Rows(orderSelected).Cells(7).Value)
        End If
    End Sub

    ' ========== 窗口尺寸改变（_窗口_商品退库备注修改_尺寸被改变） ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        ' 居中分组框
        grpMain.Left = CInt((Me.ClientSize.Width - grpMain.Width) / 2)
        grpMain.Top = CInt((Me.ClientSize.Height - grpMain.Height) / 2)
    End Sub

    ' ========== 关闭按钮（_图片框EX4_鼠标左键单击） ==========
    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ' ========== 重置按钮（_按钮EX2_鼠标左键单击） ==========
    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        Form_Load(Nothing, Nothing)
    End Sub

    ' ========== 保存按钮（_按钮EX1_鼠标左键单击） ==========
    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        InformationOperationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        InformationOperationAccount = UserAccount

        Dim modifyOrderNo As String = txtOrderNumber.Text
        Dim modifyRemarks As String = txtRemarks.Text

        ' 更新退库单备注
        Dim updateSql As String = "UPDATE xipunum_erp_tuiku_order SET `remarks`= '" & SafeSQL(modifyRemarks) & "',`updatetime`= '" & InformationOperationDate & "'  WHERE tuiku_umber ='" & SafeSQL(modifyOrderNo) & "' LIMIT 1"
        DatabaseModule.ExecuteCommand(updateSql, MySQL_Write)

        ' 记录系统日志
        LogSaveContent = ""
        LogSaveContent = "账户:" & UserAccount & " 修改退库备注，单号：" & txtOrderNumber.Text
        Dim logSql As String = "INSERT INTO xipunum_erp_xitong_log (type,title,conter,user,creationtime) VALUES ('" & SafeSQL("修改") & "','" & SafeSQL("修改退库备注") & "','" & SafeSQL(LogSaveContent) & "','" & SafeSQL(InformationOperationAccount) & "','" & InformationOperationDate & "')"
        DatabaseModule.ExecuteCommand(logSql, MySQL_Write)

        ShowSuccess("退库单:" & txtOrderNumber.Text & "备注修改成功！")
        Me.Close()

        ' 设置首页查询栏目并刷新主窗口
        HomePageQueryText = "商品退库"
        CType(Me.Owner, MainForm).RefreshSubFolderTable()
    End Sub

End Class

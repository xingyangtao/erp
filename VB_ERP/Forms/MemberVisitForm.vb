' ============================================================================
' 会员回访添加窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_会员回访添加.form.e.txt
' 包含所有1个程序集变量、6个子程序、4个SQL查询
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class MemberVisitForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（1个） ==========
    Private localOrderSelected As Integer = -1       ' 局部_订单是否选中

    ' ========== 控件声明（对应易语言窗口控件） ==========
    ' 分组框
    Private WithEvents grpAddModify As New GroupBox()  ' 添加修改_分组框

    ' 编辑框
    Private txtMemberCode As New TextBox()             ' 编辑框_会员编码
    Private txtMemberName As New TextBox()             ' 编辑框_会员姓名
    Private txtPhone As New TextBox()                  ' 编辑框_联系电话
    Private txtBirthday As New TextBox()               ' 编辑框_出生日期
    Private txtVisitTitle As New TextBox()             ' 编辑框_回访标题
    Private txtVisitContent As New TextBox()           ' 编辑框_回访内容

    ' 按钮
    Private WithEvents btnSubmit As New Button()       ' 按钮EX1（保存）
    Private WithEvents btnReset As New Button()        ' 按钮EX2（重置）
    Private WithEvents btnClose As New Button()        ' 图片框EX4（关闭）

    ' ========== 构造函数 ==========
    ' 传入选中行号：-1表示添加模式，其他表示查看详情模式
    Public Sub New(Optional selectedRow As Integer = -1, Optional memberCode As String = "",
                   Optional memberName As String = "", Optional memberPhone As String = "",
                   Optional visitTitle As String = "", Optional visitContent As String = "")
        localOrderSelected = selectedRow
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler txtPhone.KeyDown, AddressOf txtPhone_KeyDown

        ' 存储传入数据用于查看模式
        If selectedRow <> -1 Then
            txtPhone.Text = memberPhone
            txtMemberName.Text = memberName
            txtMemberCode.Text = memberCode
            txtVisitTitle.Text = visitTitle
            txtVisitContent.Text = visitContent
        End If
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "会员回访添加"
        Me.Size = New Drawing.Size(700, 500)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 关闭按钮（对应图片框EX4）
        btnClose.Text = "X"
        btnClose.Location = New Drawing.Point(Me.Width - 35, 5)
        btnClose.Size = New Drawing.Size(30, 25)
        btnClose.BackColor = Drawing.Color.Red
        btnClose.ForeColor = Drawing.Color.White
        btnClose.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnClose)

        ' 添加修改分组框（对应添加修改_分组框）
        grpAddModify.Text = "添加回访"
        grpAddModify.Size = New Drawing.Size(490, 380)
        grpAddModify.Location = New Drawing.Point((Me.Width - grpAddModify.Width) \ 2, (Me.Height - grpAddModify.Height) \ 2)
        Me.Controls.Add(grpAddModify)

        ' 联系电话（对应编辑框_联系电话）
        AddLabelAndTextBox(grpAddModify, "联系电话：", txtPhone, 16, 20, 200)

        ' 会员编码（对应编辑框_会员编码）
        AddLabelAndTextBox(grpAddModify, "会员编码：", txtMemberCode, 16, 60, 200)
        txtMemberCode.ReadOnly = True

        ' 会员姓名（对应编辑框_会员姓名）
        AddLabelAndTextBox(grpAddModify, "会员姓名：", txtMemberName, 16, 100, 200)
        txtMemberName.ReadOnly = True

        ' 出生日期（对应编辑框_出生日期）
        AddLabelAndTextBox(grpAddModify, "出生日期：", txtBirthday, 16, 140, 200)
        txtBirthday.ReadOnly = True

        ' 回访标题（对应编辑框_回访标题）
        AddLabelAndTextBox(grpAddModify, "回访标题：", txtVisitTitle, 16, 180, 460)

        ' 回访内容（对应编辑框_回访内容）
        AddLabel(grpAddModify, "回访内容：", 16, 220)
        txtVisitContent.Location = New Drawing.Point(96, 220)
        txtVisitContent.Size = New Drawing.Size(380, 100)
        txtVisitContent.Multiline = True
        grpAddModify.Controls.Add(txtVisitContent)

        ' 保存按钮（对应按钮EX1）
        btnSubmit.Text = "保存"
        btnSubmit.Location = New Drawing.Point(150, 340)
        btnSubmit.Size = New Drawing.Size(100, 30)
        grpAddModify.Controls.Add(btnSubmit)

        ' 重置按钮（对应按钮EX2）
        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(280, 340)
        btnReset.Size = New Drawing.Size(100, 30)
        grpAddModify.Controls.Add(btnReset)
    End Sub

    Private Sub AddLabel(parent As Control, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    Private Sub AddLabelAndTextBox(parent As Control, labelText As String, txtBox As TextBox, x As Integer, y As Integer, width As Integer)
        AddLabel(parent, labelText, x, y)
        txtBox.Location = New Drawing.Point(x + 80, y)
        txtBox.Size = New Drawing.Size(width, 25)
        parent.Controls.Add(txtBox)
    End Sub

    ' ========== _窗口_会员回访添加_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 对应易语言：赋值(局部_订单是否选中, 窗口_主窗口.高级表格1.取光标行号())

        ' 对应易语言：如果(等于(局部_订单是否选中, -1)) → 添加回访模式
        If localOrderSelected = -1 Then
            grpAddModify.Text = "添加回访"
            txtPhone.ReadOnly = False
            txtPhone.Text = ""
            txtMemberName.Text = ""
            txtMemberCode.Text = ""
            txtVisitTitle.Text = ""
            txtVisitContent.Text = ""
            txtVisitTitle.ReadOnly = False
            txtVisitContent.ReadOnly = False
            btnSubmit.Visible = True
            btnReset.Visible = True
        Else
            ' 对应易语言：否则 → 回访详情模式
            grpAddModify.Text = "回访详情"
            txtPhone.ReadOnly = True
            ' 数据已通过构造函数传入
            txtVisitTitle.ReadOnly = True
            txtVisitContent.ReadOnly = True
            btnSubmit.Visible = True
            btnReset.Visible = True
        End If
    End Sub

    ' ========== _窗口_会员回访添加_尺寸被改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        ' 对应易语言：分组框居中
        grpAddModify.Left = (Me.Width - grpAddModify.Width) \ 2
        grpAddModify.Top = (Me.Height - grpAddModify.Height) \ 2
    End Sub

    ' ========== _图片框EX4_鼠标左键单击（关闭按钮） ==========
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ' ========== _按钮EX2_鼠标左键单击（重置按钮） ==========
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        ' 对应易语言：重新执行_窗口_会员回访添加_创建完毕
        localOrderSelected = -1
        Form_Load(Me, EventArgs.Empty)
    End Sub

    ' ========== _编辑框_联系电话_键盘事件 ==========
    Private Sub txtPhone_KeyDown(sender As Object, e As KeyEventArgs)
        ' 对应易语言：如果真(等于(键代码, 13)) → Enter键
        If e.KeyCode <> Keys.Enter Then Return

        ' 对应易语言：如果(不等于(编辑框_联系电话.内容, ""))
        If txtPhone.Text <> "" Then
            ' 对应易语言：如果(等于(取文本左边(编辑框_联系电话.内容, 1), "1"))
            If txtPhone.Text.Substring(0, 1) = "1" Then
                ' 对应易语言：如果(等于(取文本长度(编辑框_联系电话.内容), 11))
                If txtPhone.Text.Length = 11 Then
                    ' 对应易语言：SELECT * FROM xipunum_erp_member where tel='...'
                    Dim checkSql As String = "SELECT * FROM xipunum_erp_member where tel='" & SafeSQL(txtPhone.Text) & "'"
                    Dim checkDt As DataTable = DatabaseModule.ExecuteQuery(checkSql, MySQL_Read)

                    ' 对应易语言：如果真(等于(会员是否存在, 0))
                    If checkDt.Rows.Count = 0 Then
                        MessageBox.Show("此会员手机不存在！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        txtPhone.Text = ""
                        txtPhone.Focus()
                        txtPhone.ReadOnly = False
                        Return
                    End If

                    ' 对应易语言：获取会员信息
                    Dim memberInfoCode As String = ""
                    Dim memberInfoName As String = ""
                    Dim memberInfoBirthday As String = ""

                    Dim memberSql As String = "SELECT * FROM xipunum_erp_member where tel='" & SafeSQL(txtPhone.Text) & "'"
                    Dim memberDt As DataTable = DatabaseModule.ExecuteQuery(memberSql, MySQL_Read)

                    If memberDt.Rows.Count > 0 Then
                        memberInfoCode = SafeString(memberDt.Rows(0)("customer_code"))
                        memberInfoName = DatabaseModule.GBKToUTF8(SafeString(memberDt.Rows(0)("name")))
                        memberInfoBirthday = DatabaseModule.GBKToUTF8(SafeString(memberDt.Rows(0)("shengri")))
                    End If

                    txtMemberName.Text = memberInfoName
                    txtMemberCode.Text = memberInfoCode
                    txtBirthday.Text = memberInfoBirthday
                    txtPhone.ReadOnly = True
                    Return
                Else
                    MessageBox.Show("请输入正确的手机号码", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtPhone.Text = ""
                    txtPhone.Focus()
                    Return
                End If
            Else
                MessageBox.Show("请输入正确的手机号码", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPhone.Text = ""
                txtPhone.Focus()
                Return
            End If
        Else
            MessageBox.Show("会员联系电话不能为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPhone.Focus()
            Return
        End If
    End Sub

    ' ========== _按钮EX1_鼠标左键单击（保存按钮） ==========
    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Dim infoMemberCode As String = txtMemberCode.Text
        Dim infoVisitTime As String = DateTime.Now.ToString("yyyy-MM-dd")
        Dim infoVisitTitle As String = txtVisitTitle.Text
        Dim infoVisitContent As String = txtVisitContent.Text

        Dim globalInfoOperationDate As String = DateTime.Now.ToString("yyyy-MM-dd") & " " & DateTime.Now.ToString("HH:mm:ss")
        Dim globalInfoOperationAccount As String = UserAccount
        Dim globalLogSaveContent As String = ""

        ' 对应易语言：如果(不等于(编辑框_会员编码.内容, ""))
        If txtMemberCode.Text <> "" Then
            ' 对应易语言：如果(不等于(编辑框_回访标题.内容, ""))
            If txtVisitTitle.Text <> "" Then
                ' 对应易语言：如果(不等于(编辑框_回访标题.内容, "")) → 实际是检查回访内容
                If txtVisitContent.Text <> "" Then
                    ' 对应易语言：增加记录 INSERT INTO xipunum_erp_visit
                    Dim insertSql As String = "INSERT INTO xipunum_erp_visit (customer_code, returntitle, returnconter, returndata, cjuser, creationtime) VALUES ('" &
                        SafeSQL(infoMemberCode) & "','" & SafeSQL(infoVisitTitle) & "','" & SafeSQL(infoVisitContent) & "','" &
                        SafeSQL(infoVisitTime) & "','" & SafeSQL(globalInfoOperationAccount) & "','" & SafeSQL(globalInfoOperationDate) & "')"
                    DatabaseModule.ExecuteCommand(insertSql)

                    MessageBox.Show("会员:" & txtMemberName.Text & "回访信息添加成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' 对应易语言：日志保存
                    globalLogSaveContent = "账户:" & UserAccount & "  添加会员:" & txtMemberName.Text & " 会员编码:" & txtMemberCode.Text &
                        " 回访标题:" & txtVisitTitle.Text & " 回访内容:" & txtVisitContent.Text & " 回访信息"

                    Dim logSql As String = "INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('添加','添加回访信息','" &
                        SafeSQL(globalLogSaveContent) & "','" & SafeSQL(globalInfoOperationAccount) & "','" & SafeSQL(globalInfoOperationDate) & "')"
                    DatabaseModule.ExecuteCommand(logSql)

                    Me.Close()
                    Return
                Else
                    MessageBox.Show("回访内容不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtVisitContent.Focus()
                    Return
                End If
            Else
                MessageBox.Show("回访标题不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtVisitTitle.Focus()
                Return
            End If
        Else
            MessageBox.Show("回访客户不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPhone.Focus()
            Return
        End If
    End Sub

End Class

' ============================================================================
' 密码修改窗口
' 功能: 密码修改
' 对应易语言: 窗口程序集_窗口_密码修改
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.IO

Public Class PasswordChangeForm
    Inherits System.Windows.Forms.Form

    ' ========== UI控件声明 ==========
    ' 对应易语言控件: 添加修改_分组框
    Private grpMain As New GroupBox()
    ' 对应易语言控件: 编辑框_原始密码
    Private txtOldPassword As New TextBox()
    ' 对应易语言控件: 编辑框_新密码
    Private txtNewPassword As New TextBox()
    ' 对应易语言控件: 编辑框_确认密码
    Private txtConfirmPassword As New TextBox()
    ' 对应易语言控件: 按钮EX1 (确认修改)
    Private WithEvents btnSave As New Button()
    ' 对应易语言控件: 按钮EX2 (重置)
    Private WithEvents btnReset As New Button()
    ' 对应易语言控件: 图片框EX4 (关闭按钮)
    Private WithEvents btnClose As New Button()

    ' ========== 标签控件 ==========
    Private lblOld As New Label()
    Private lblNew As New Label()
    Private lblConfirm As New Label()

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        ' 对应易语言: _窗口_密码修改_创建完毕
        ResetForm()
    End Sub

    ' ========== 初始化UI布局 ==========
    ' 对应易语言窗口定义(.form.json)中的控件布局
    Private Sub InitializeUI()
        Me.Text = "密码修改"
        Me.Size = New Drawing.Size(400, 300)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' --- 关闭按钮 (对应易语言: 图片框EX4_鼠标左键单击 → 销毁窗口) ---
        btnClose.Text = "X"
        btnClose.Font = New Font("Microsoft YaHei", 10, FontStyle.Bold)
        btnClose.ForeColor = Color.Gray
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.FlatAppearance.BorderSize = 0
        btnClose.Size = New Drawing.Size(28, 28)
        btnClose.Location = New Drawing.Point(Me.ClientSize.Width - 36, 4)
        btnClose.Cursor = Cursors.Hand
        btnClose.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Me.Controls.Add(btnClose)

        ' --- 分组框 (对应易语言: 添加修改_分组框) ---
        grpMain.Text = "密码修改"
        grpMain.Size = New Drawing.Size(340, 220)
        grpMain.Location = New Drawing.Point((Me.ClientSize.Width - grpMain.Width) / 2, 20)
        grpMain.Anchor = AnchorStyles.None
        Me.Controls.Add(grpMain)

        ' --- 原始密码标签 ---
        lblOld.Text = "原密码："
        lblOld.Location = New Drawing.Point(20, 30)
        lblOld.AutoSize = True
        grpMain.Controls.Add(lblOld)

        ' --- 原始密码编辑框 (对应易语言: 编辑框_原始密码) ---
        txtOldPassword.Location = New Drawing.Point(100, 27)
        txtOldPassword.Size = New Drawing.Size(210, 25)
        txtOldPassword.UseSystemPasswordChar = True
        ' 对应易语言: 输入方式 = 位或(#输入模式_密码输入, #输入模式_禁止输入法)
        txtOldPassword.PasswordChar = "*"
        grpMain.Controls.Add(txtOldPassword)

        ' --- 新密码标签 ---
        lblNew.Text = "新密码："
        lblNew.Location = New Drawing.Point(20, 70)
        lblNew.AutoSize = True
        grpMain.Controls.Add(lblNew)

        ' --- 新密码编辑框 (对应易语言: 编辑框_新密码) ---
        txtNewPassword.Location = New Drawing.Point(100, 67)
        txtNewPassword.Size = New Drawing.Size(210, 25)
        txtNewPassword.UseSystemPasswordChar = True
        txtNewPassword.PasswordChar = "*"
        grpMain.Controls.Add(txtNewPassword)

        ' --- 确认密码标签 ---
        lblConfirm.Text = "确认密码："
        lblConfirm.Location = New Drawing.Point(20, 110)
        lblConfirm.AutoSize = True
        grpMain.Controls.Add(lblConfirm)

        ' --- 确认密码编辑框 (对应易语言: 编辑框_确认密码) ---
        txtConfirmPassword.Location = New Drawing.Point(100, 107)
        txtConfirmPassword.Size = New Drawing.Size(210, 25)
        txtConfirmPassword.UseSystemPasswordChar = True
        txtConfirmPassword.PasswordChar = "*"
        grpMain.Controls.Add(txtConfirmPassword)

        ' --- 确认按钮 (对应易语言: 按钮EX1_鼠标左键单击) ---
        btnSave.Text = "确认修改"
        btnSave.Location = New Drawing.Point(60, 160)
        btnSave.Size = New Drawing.Size(100, 35)
        btnSave.BackColor = Color.FromArgb(0xAABA, 0x16FF >> 8, 0x16FF And 0xFF)
        btnSave.FlatStyle = FlatStyle.Flat
        grpMain.Controls.Add(btnSave)

        ' --- 重置按钮 (对应易语言: 按钮EX2_鼠标左键单击) ---
        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(180, 160)
        btnReset.Size = New Drawing.Size(100, 35)
        btnReset.BackColor = Color.FromArgb(0xEFECE, 0xF9 >> 8, 0xF9 And 0xFF)
        btnReset.FlatStyle = FlatStyle.Flat
        grpMain.Controls.Add(btnReset)
    End Sub

    ' ========== 窗口创建完毕 / 重置表单 ==========
    ' 对应易语言: _窗口_密码修改_创建完毕
    ' 赋值(编辑框_原始密码.内容, "")
    ' 赋值(编辑框_新密码.内容, "")
    ' 赋值(编辑框_确认密码.内容, "")
    ' 赋值(编辑框_原始密码.输入方式, 位或(#输入模式_密码输入, #输入模式_禁止输入法))
    ' 赋值(编辑框_新密码.输入方式, 位或(#输入模式_密码输入, #输入模式_禁止输入法))
    ' 赋值(编辑框_确认密码.输入方式, 位或(#输入模式_密码输入, #输入模式_禁止输入法))
    Private Sub ResetForm()
        txtOldPassword.Text = ""
        txtNewPassword.Text = ""
        txtConfirmPassword.Text = ""
        txtOldPassword.UseSystemPasswordChar = True
        txtOldPassword.PasswordChar = "*"
        txtNewPassword.UseSystemPasswordChar = True
        txtNewPassword.PasswordChar = "*"
        txtConfirmPassword.UseSystemPasswordChar = True
        txtConfirmPassword.PasswordChar = "*"
    End Sub

    ' ========== 窗口尺寸改变 - 分组框居中 ==========
    ' 对应易语言: _窗口_密码修改_尺寸被改变
    ' 赋值(添加修改_分组框.左边, 相除(相减(窗口_密码修改.宽度, 添加修改_分组框.宽度), 2))
    ' 赋值(添加修改_分组框.顶边, 相除(相减(窗口_密码修改.高度, 添加修改_分组框.高度), 2))
    Private Sub PasswordChangeForm_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        grpMain.Location = New Drawing.Point(
            (Me.ClientSize.Width - grpMain.Width) / 2,
            (Me.ClientSize.Height - grpMain.Height) / 2
        )
        btnClose.Location = New Drawing.Point(Me.ClientSize.Width - 36, 4)
    End Sub

    ' ========== 关闭按钮 (对应易语言: _图片框EX4_鼠标左键单击) ===
    ' 窗口_密码修改.销毁()
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ' ========== 重置按钮 (对应易语言: _按钮EX2_鼠标左键单击) ===
    ' 窗口_密码修改._窗口_密码修改_创建完毕()
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        ResetForm()
    End Sub

    ' ========== 确认修改按钮 (对应易语言: _按钮EX1_鼠标左键单击) ===
    ' 这是核心业务逻辑，完全按照易语言的嵌套条件判断实现
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' 对应易语言局部变量:
        ' .局部变量 原始密码, 文本型
        ' .局部变量 新密码, 文本型
        ' .局部变量 判断是否存在集句柄, 整数型
        ' .局部变量 判断是否存在, 整数型
        Dim oldPassword As String
        Dim newPassword As String
        Dim recordCount As Integer

        ' 对应易语言: 如果真(或者(等于(全_MySQL读取, 0), 等于(全_MySQL写入, 0)))
        '   提示框Ex_添加消息("数据库连接无效！", 9, 1000, , 假, )
        '   返回()
        If DatabaseModule.MySQL_Read Is Nothing OrElse DatabaseModule.MySQL_Write Is Nothing Then
            ShowWarning("数据库连接无效！")
            Return
        End If
        If DatabaseModule.MySQL_Read.State <> ConnectionState.Open AndAlso DatabaseModule.MySQL_Write.State <> ConnectionState.Open Then
            ShowWarning("数据库连接无效！")
            Return
        End If

        ' 对应易语言: 赋值(全局_信息操作日期, 到文本(时间_格式化(取现行时间(), "yyyy-MM-dd", " hh:mm:ss", 真)))
        ' 对应易语言: 赋值(全局_信息操作账户, 全局_用户账户)
        ' 对应易语言: 赋值(全局_日志保存内容, "")
        LogOperationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        LogOperationAccount = UserAccount
        LogContent = ""

        ' 对应易语言: 赋值(原始密码, 到文本(取数据摘要(到字节集(相加(编辑框_原始密码.内容, 全_授权数据码))))
        ' 对应易语言: 赋值(新密码, 到文本(取数据摘要(到字节集(相加(编辑框_确认密码.内容, 全_授权数据码))))
        ' 注意: 易语言的"取数据摘要"是MD5，"全_授权数据码"对应VB.NET的AuthDataNum
        oldPassword = MD5Encrypt(txtOldPassword.Text, AuthDataNum)
        newPassword = MD5Encrypt(txtConfirmPassword.Text, AuthDataNum)

        ' ===== 开始嵌套条件判断（完全按照易语言逻辑顺序） =====

        ' 对应易语言: .如果(不等于(编辑框_原始密码.内容, ""))
        If txtOldPassword.Text <> "" Then
            ' 对应易语言: .如果(不等于(编辑框_新密码.内容, ""))
            If txtNewPassword.Text <> "" Then
                ' 对应易语言: .如果(大于或等于(取文本长度(编辑框_新密码.内容), 6))
                If txtNewPassword.Text.Length >= 6 Then
                    ' 对应易语言: .如果(不等于(编辑框_新密码.内容, "123456"))
                    If txtNewPassword.Text <> "123456" Then
                        ' 对应易语言: .如果(等于(编辑框_新密码.内容, 编辑框_确认密码.内容))
                        If txtNewPassword.Text = txtConfirmPassword.Text Then
                            ' 对应易语言: .如果(不等于(编辑框_原始密码.内容, 编辑框_确认密码.内容))
                            If txtOldPassword.Text <> txtConfirmPassword.Text Then
                                ' 对应易语言: 执行SQL语句(全_MySQL读取, "SELECT * FROM xipunum_erp_user where user ='全局_信息操作账户' and password ='原始密码' LIMIT 1")
                                Dim sql As String = $"SELECT * FROM xipunum_erp_user WHERE user='{SafeSQL(LogOperationAccount)}' AND password='{SafeSQL(oldPassword)}' LIMIT 1"
                                Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, DatabaseModule.MySQL_Read)

                                ' 对应易语言: 赋值(判断是否存在集句柄, 取记录集(全_MySQL读取))
                                ' 对应易语言: 赋值(判断是否存在, 取记录集行数(判断是否存在集句柄))
                                ' 对应易语言: 释放记录集(判断是否存在集句柄)
                                recordCount = dt.Rows.Count

                                ' 对应易语言: .如果(不等于(判断是否存在, 0))
                                If recordCount <> 0 Then
                                    ' 对应易语言: 执行SQL语句(全_MySQL写入, "UPDATE xipunum_erp_user SET password= '新密码', updatetime= '全局_信息操作日期' WHERE xipunum_erp_user.user ='全局_信息操作账户' LIMIT 1")
                                    Dim updateSql As String = $"UPDATE xipunum_erp_user SET password='{SafeSQL(newPassword)}', updatetime='{LogOperationDate}' WHERE user='{SafeSQL(LogOperationAccount)}' LIMIT 1"
                                    DatabaseModule.ExecuteCommand(updateSql, DatabaseModule.MySQL_Write)

                                    ' 对应易语言: 信息框("密码修改成功！", 0, , )
                                    MessageBox.Show("密码修改成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)

                                    ' 对应易语言: 提示框Ex_添加消息("密码修改成功！", 5, 1000, , 假, )
                                    ShowSuccess("密码修改成功！")

                                    ' 对应易语言注释: 清除本地记住的旧密码（避免注销后自动填错密码）
                                    ' 对应易语言: 写配置项(相加(取运行目录(), "\data\user.ini"), "user", "pass", "")
                                    Dim userIniPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "user.ini")
                                    WriteIniFile(userIniPath, "user", "pass", "")

                                    ' 对应易语言: 赋值(全局_日志保存内容, 相加("账户:", 全局_用户账户, " 修改账户密码"))
                                    LogContent = "账户:" & UserAccount & " 修改账户密码"

                                    ' 对应易语言: 增加记录(全_MySQL写入, "xipunum_erp_xitong_log", 
                                    '   "type='修改',title='修改账户密码',conter='全局_日志保存内容',user='全局_信息操作账户',creationtime='全局_信息操作日期'")
                                    Dim logSql As String = $"INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('修改', '修改账户密码', '{SafeSQL(LogContent)}', '{SafeSQL(LogOperationAccount)}', '{LogOperationDate}')"
                                    DatabaseModule.ExecuteCommand(logSql, DatabaseModule.MySQL_Write)

                                    ' 对应易语言: 窗口_密码修改.销毁()
                                    Me.Close()

                                    ' 对应易语言: 窗口_主窗口._注销_被选择()
                                    ' 在VB.NET中，关闭此窗口后触发注销（关闭主窗口，回到登录窗口）
                                    For Each openForm As Form In Application.OpenForms
                                        If TypeOf openForm Is MainForm Then
                                            openForm.Close()
                                            Return
                                        End If
                                    Next
                                    Return
                                Else
                                    ' 对应易语言: 提示框Ex_添加消息("请输入正确的原始密码！", 7, 1000, , 假, )
                                    ShowWarning("请输入正确的原始密码！")
                                    txtOldPassword.Focus()
                                    Return
                                End If
                            Else
                                ' 对应易语言: 提示框Ex_添加消息("新密码和原始密码不能相同！", 7, 1000, , 假, )
                                ShowWarning("新密码和原始密码不能相同！")
                                txtNewPassword.Focus()
                                Return
                            End If
                        Else
                            ' 对应易语言: 提示框Ex_添加消息("两次输入的密码不相同！", 7, 1000, , 假, )
                            ShowWarning("两次输入的密码不相同！")
                            txtNewPassword.Focus()
                            Return
                        End If
                    Else
                        ' 对应易语言: 提示框Ex_添加消息("新密码不能为123456！", 7, 1000, , 假, )
                        ShowWarning("新密码不能为123456！")
                        txtNewPassword.Focus()
                        Return
                    End If
                Else
                    ' 对应易语言: 提示框Ex_添加消息("密码不能少于6位数！", 7, 1000, , 假, )
                    ShowWarning("密码不能少于6位数！")
                    txtNewPassword.Focus()
                    Return
                End If
            Else
                ' 对应易语言: 提示框Ex_添加消息("新密码不能为空！", 7, 1000, , 假, )
                ShowWarning("新密码不能为空！")
                txtNewPassword.Focus()
                Return
            End If
        Else
            ' 对应易语言: 提示框Ex_添加消息("原始密码不能为空！", 7, 1000, , 假, )
            ShowWarning("原始密码不能为空！")
            txtOldPassword.Focus()
            Return
        End If
    End Sub

End Class

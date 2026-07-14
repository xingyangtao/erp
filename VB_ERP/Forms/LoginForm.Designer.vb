' ============================================================================
' 登录窗口设计器文件
' ============================================================================

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class LoginForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.btnLogin = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.chkRemember = New System.Windows.Forms.CheckBox()
        Me.lblWelcome = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.btnTogglePassword = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        ' lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("微软雅黑", 18.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitle.Location = New System.Drawing.Point(80, 30)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(240, 31)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "ERPV4 珠宝管理系统"
        '
        ' lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Font = New System.Drawing.Font("微软雅黑", 12.0!)
        Me.lblUsername.Location = New System.Drawing.Point(50, 100)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(74, 21)
        Me.lblUsername.TabIndex = 1
        Me.lblUsername.Text = "用户名："
        '
        ' lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Font = New System.Drawing.Font("微软雅黑", 12.0!)
        Me.lblPassword.Location = New System.Drawing.Point(50, 150)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(58, 21)
        Me.lblPassword.TabIndex = 2
        Me.lblPassword.Text = "密码："
        '
        ' txtUsername
        '
        Me.txtUsername.Font = New System.Drawing.Font("微软雅黑", 12.0!)
        Me.txtUsername.Location = New System.Drawing.Point(130, 97)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(220, 29)
        Me.txtUsername.TabIndex = 3
        '
        ' txtPassword
        '
        Me.txtPassword.Font = New System.Drawing.Font("微软雅黑", 12.0!)
        Me.txtPassword.Location = New System.Drawing.Point(130, 147)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(190, 29)
        Me.txtPassword.TabIndex = 4
        Me.txtPassword.UseSystemPasswordChar = True
        '
        ' btnTogglePassword
        '
        Me.btnTogglePassword.Font = New System.Drawing.Font("微软雅黑", 10.0!)
        Me.btnTogglePassword.Location = New System.Drawing.Point(322, 147)
        Me.btnTogglePassword.Name = "btnTogglePassword"
        Me.btnTogglePassword.Size = New System.Drawing.Size(28, 29)
        Me.btnTogglePassword.TabIndex = 11
        Me.btnTogglePassword.Text = "👁"
        Me.btnTogglePassword.UseVisualStyleBackColor = True
        '
        ' btnLogin
        '
        Me.btnLogin.Font = New System.Drawing.Font("微软雅黑", 12.0!)
        Me.btnLogin.Location = New System.Drawing.Point(130, 200)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(100, 35)
        Me.btnLogin.TabIndex = 5
        Me.btnLogin.Text = "登录"
        Me.btnLogin.UseVisualStyleBackColor = True
        '
        ' btnCancel
        '
        Me.btnCancel.Font = New System.Drawing.Font("微软雅黑", 12.0!)
        Me.btnCancel.Location = New System.Drawing.Point(250, 200)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(100, 35)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "取消"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        ' lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Font = New System.Drawing.Font("微软雅黑", 9.0!)
        Me.lblVersion.ForeColor = System.Drawing.Color.Gray
        Me.lblVersion.Location = New System.Drawing.Point(150, 260)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(100, 17)
        Me.lblVersion.TabIndex = 7
        Me.lblVersion.Text = "版本: V3.0.88"
        '
        ' chkRemember
        '
        Me.chkRemember.AutoSize = True
        Me.chkRemember.Font = New System.Drawing.Font("微软雅黑", 10.0!)
        Me.chkRemember.Location = New System.Drawing.Point(130, 180)
        Me.chkRemember.Name = "chkRemember"
        Me.chkRemember.Size = New System.Drawing.Size(90, 24)
        Me.chkRemember.TabIndex = 8
        Me.chkRemember.Text = "记住密码"
        Me.chkRemember.UseVisualStyleBackColor = True
        '
        ' lblWelcome
        '
        Me.lblWelcome.AutoSize = True
        Me.lblWelcome.Font = New System.Drawing.Font("微软雅黑", 12.0!)
        Me.lblWelcome.Location = New System.Drawing.Point(130, 240)
        Me.lblWelcome.Name = "lblWelcome"
        Me.lblWelcome.Size = New System.Drawing.Size(150, 21)
        Me.lblWelcome.TabIndex = 9
        Me.lblWelcome.Text = "欢迎来到莆阳科技！"
        '
        ' lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Font = New System.Drawing.Font("微软雅黑", 9.0!)
        Me.lblStatus.ForeColor = System.Drawing.Color.Blue
        Me.lblStatus.Location = New System.Drawing.Point(130, 280)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(100, 17)
        Me.lblStatus.TabIndex = 10
        Me.lblStatus.Text = "正在初始化..."
        '
        ' LoginForm
        '
        Me.AcceptButton = Me.btnLogin
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(400, 320)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.btnTogglePassword)
        Me.Controls.Add(Me.lblWelcome)
        Me.Controls.Add(Me.chkRemember)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnLogin)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.txtUsername)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.lblUsername)
        Me.Controls.Add(Me.lblTitle)
        Me.Name = "LoginForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ERPV4 珠宝管理系统 - 登录"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents lblTitle As Label
    Friend WithEvents lblUsername As Label
    Friend WithEvents lblPassword As Label
    Friend WithEvents txtUsername As TextBox
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents btnLogin As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents lblVersion As Label
    Friend WithEvents chkRemember As CheckBox
    Friend WithEvents lblWelcome As Label
    Friend WithEvents lblStatus As Label
    Friend WithEvents btnTogglePassword As Button

End Class

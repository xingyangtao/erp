' ============================================================================
' 等待窗口
' 功能: 显示等待动画
' ============================================================================

Public Class WaitForm
    Inherits System.Windows.Forms.Form

    Private lblMessage As New Label()
    Private progressBar As New ProgressBar()

    Public Sub New(Optional message As String = "请稍候...")
        InitializeUI(message)
    End Sub

    Private Sub InitializeUI(message As String)
        Me.Text = "请稍候"
        Me.Size = New Drawing.Size(300, 120)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ControlBox = False

        lblMessage.Text = message
        lblMessage.Location = New Drawing.Point(30, 20)
        lblMessage.Size = New Drawing.Size(230, 20)
        lblMessage.TextAlign = ContentAlignment.MiddleCenter
        Me.Controls.Add(lblMessage)

        progressBar.Location = New Drawing.Point(30, 50)
        progressBar.Size = New Drawing.Size(230, 25)
        progressBar.Style = ProgressBarStyle.Marquee
        Me.Controls.Add(progressBar)
    End Sub

    Public Sub SetMessage(message As String)
        lblMessage.Text = message
        Application.DoEvents()
    End Sub
End Class

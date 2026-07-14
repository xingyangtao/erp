' ============================================================================
' 提示消息窗口
' 功能: 消息提示
' ============================================================================

Public Class MessageForm
    Inherits System.Windows.Forms.Form

    Private lblMessage As New Label()
    Private WithEvents btnOK As New Button()
    Private messageType As Integer = 0

    Public Sub New(message As String, Optional type As Integer = 0)
        messageType = type
        InitializeUI(message)
    End Sub

    Private Sub InitializeUI(message As String)
        Me.Text = "提示"
        Me.Size = New Drawing.Size(350, 180)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        lblMessage.Text = message
        lblMessage.Location = New Drawing.Point(30, 30)
        lblMessage.Size = New Drawing.Size(280, 60)
        lblMessage.TextAlign = ContentAlignment.MiddleCenter
        Me.Controls.Add(lblMessage)

        btnOK.Text = "确定"
        btnOK.Location = New Drawing.Point(120, 100)
        btnOK.Size = New Drawing.Size(80, 30)
        Me.Controls.Add(btnOK)
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Public Overloads Shared Sub Show(message As String, Optional type As Integer = 0)
        Using form As New MessageForm(message, type)
            form.ShowDialog()
        End Using
    End Sub
End Class

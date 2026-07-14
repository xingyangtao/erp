' ============================================================================
' 确认对话框窗口
' 功能: 通用确认对话框
' ============================================================================

Public Class ConfirmDialogForm
    Inherits System.Windows.Forms.Form

    Private lblMessage As New Label()
    Private WithEvents btnYes As New Button()
    Private WithEvents btnNo As New Button()

    Public Sub New(message As String, Optional title As String = "确认")
        InitializeUI(message, title)
    End Sub

    Private Sub InitializeUI(message As String, title As String)
        Me.Text = title
        Me.Size = New Drawing.Size(400, 180)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        lblMessage.Text = message
        lblMessage.Location = New Drawing.Point(30, 30)
        lblMessage.Size = New Drawing.Size(330, 60)
        lblMessage.TextAlign = ContentAlignment.MiddleCenter
        Me.Controls.Add(lblMessage)

        btnYes.Text = "是"
        btnYes.Location = New Drawing.Point(100, 100)
        btnYes.Size = New Drawing.Size(80, 30)
        Me.Controls.Add(btnYes)

        btnNo.Text = "否"
        btnNo.Location = New Drawing.Point(200, 100)
        btnNo.Size = New Drawing.Size(80, 30)
        Me.Controls.Add(btnNo)
    End Sub

    Private Sub btnYes_Click(sender As Object, e As EventArgs) Handles btnYes.Click
        Me.DialogResult = DialogResult.Yes
        Me.Close()
    End Sub

    Private Sub btnNo_Click(sender As Object, e As EventArgs) Handles btnNo.Click
        Me.DialogResult = DialogResult.No
        Me.Close()
    End Sub

    Public Overloads Shared Function Show(message As String, Optional title As String = "确认") As Boolean
        Using form As New ConfirmDialogForm(message, title)
            Return form.ShowDialog() = DialogResult.Yes
        End Using
    End Function
End Class

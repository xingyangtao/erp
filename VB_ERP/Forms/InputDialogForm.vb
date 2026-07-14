' ============================================================================
' 输入对话框窗口
' 功能: 通用输入对话框
' ============================================================================

Public Class InputDialogForm
    Inherits System.Windows.Forms.Form

    Private lblPrompt As New Label()
    Private txtInput As New TextBox()
    Private WithEvents btnOK As New Button()
    Private WithEvents btnCancel As New Button()

    Public Property InputValue As String = ""

    Public Sub New(prompt As String, Optional title As String = "输入", Optional defaultValue As String = "")
        InitializeUI(prompt, title, defaultValue)
    End Sub

    Private Sub InitializeUI(prompt As String, title As String, defaultValue As String)
        Me.Text = title
        Me.Size = New Drawing.Size(400, 180)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        lblPrompt.Text = prompt
        lblPrompt.Location = New Drawing.Point(30, 20)
        lblPrompt.AutoSize = True
        Me.Controls.Add(lblPrompt)

        txtInput.Text = defaultValue
        txtInput.Location = New Drawing.Point(30, 50)
        txtInput.Size = New Drawing.Size(330, 25)
        Me.Controls.Add(txtInput)

        btnOK.Text = "确定"
        btnOK.Location = New Drawing.Point(100, 100)
        btnOK.Size = New Drawing.Size(80, 30)
        Me.Controls.Add(btnOK)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(200, 100)
        btnCancel.Size = New Drawing.Size(80, 30)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        InputValue = txtInput.Text
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Public Overloads Shared Function Show(prompt As String, Optional title As String = "输入", Optional defaultValue As String = "") As String
        Using form As New InputDialogForm(prompt, title, defaultValue)
            If form.ShowDialog() = DialogResult.OK Then
                Return form.InputValue
            End If
            Return Nothing
        End Using
    End Function
End Class

' ============================================================================
' 打印预览窗口
' 功能: 打印预览
' ============================================================================

Public Class PrintPreviewForm
    Inherits System.Windows.Forms.Form

    Private lblContent As New Label()
    Private WithEvents btnPrint As New Button()
    Private WithEvents btnCancel As New Button()

    Public Property PrintContent As String = ""

    Public Sub New(content As String)
        PrintContent = content
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        Me.Text = "打印预览"
        Me.Size = New Drawing.Size(600, 500)
        Me.StartPosition = FormStartPosition.CenterParent

        lblContent.Text = PrintContent
        lblContent.Location = New Drawing.Point(30, 30)
        lblContent.Size = New Drawing.Size(530, 350)
        lblContent.TextAlign = ContentAlignment.TopLeft
        lblContent.BorderStyle = BorderStyle.FixedSingle
        lblContent.AutoEllipsis = False
        Me.Controls.Add(lblContent)

        btnPrint.Text = "打印"
        btnPrint.Location = New Drawing.Point(180, 410)
        btnPrint.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnPrint)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(300, 410)
        btnCancel.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        ShowSuccess("打印功能开发中...")
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

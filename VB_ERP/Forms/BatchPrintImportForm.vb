' ============================================================================
' 批量打印导入编码窗口
' 功能: 批量导入打印商品编码
' ============================================================================

Public Class BatchPrintImportForm
    Inherits System.Windows.Forms.Form

    Private txtCodes As New TextBox()
    Private WithEvents btnImport As New Button()
    Private WithEvents btnCancel As New Button()

    Public Property ImportedCodes As New List(Of String)

    Public Sub New()
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        Me.Text = "批量打印导入编码"
        Me.Size = New Drawing.Size(500, 400)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblTip As New Label()
        lblTip.Text = "请输入商品编码（每行一个）："
        lblTip.Location = New Drawing.Point(30, 20)
        lblTip.AutoSize = True
        Me.Controls.Add(lblTip)

        txtCodes.Location = New Drawing.Point(30, 50)
        txtCodes.Size = New Drawing.Size(420, 250)
        txtCodes.Multiline = True
        txtCodes.ScrollBars = ScrollBars.Vertical
        Me.Controls.Add(txtCodes)

        btnImport.Text = "导入"
        btnImport.Location = New Drawing.Point(150, 320)
        btnImport.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnImport)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(270, 320)
        btnCancel.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        If String.IsNullOrEmpty(txtCodes.Text.Trim()) Then
            ShowWarning("请输入商品编码！")
            Return
        End If

        Dim codes As String() = txtCodes.Text.Split(New Char() {vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
        ImportedCodes.Clear()

        For Each code As String In codes
            Dim trimCode As String = code.Trim()
            If Not String.IsNullOrEmpty(trimCode) Then
                ImportedCodes.Add(trimCode)
            End If
        Next

        If ImportedCodes.Count = 0 Then
            ShowWarning("未找到有效的编码！")
            Return
        End If

        ShowSuccess($"成功导入 {ImportedCodes.Count} 个编码")
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

' ============================================================================
' 款式数据汇总图片预览窗口
' 功能: 款式数据汇总图片预览
' ============================================================================

Public Class StyleSummaryImageForm
    Inherits System.Windows.Forms.Form

    Private picImage As New PictureBox()
    Private imageUrl As String = ""

    Public Sub New(imageUrl As String)
        Me.imageUrl = imageUrl
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "款式数据汇总图片预览"
        Me.Size = New Drawing.Size(600, 500)
        Me.StartPosition = FormStartPosition.CenterParent

        picImage.Dock = DockStyle.Fill
        picImage.SizeMode = PictureBoxSizeMode.Zoom
        Me.Controls.Add(picImage)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        If Not String.IsNullOrEmpty(imageUrl) Then
            Try
                picImage.Load(imageUrl)
            Catch ex As Exception
                ShowError("加载图片失败：" & ex.Message)
            End Try
        End If
    End Sub
End Class

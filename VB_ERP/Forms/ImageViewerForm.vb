' ============================================================================
' 图片查看器窗口
' 功能: 查看图片
' ============================================================================

Public Class ImageViewerForm
    Inherits System.Windows.Forms.Form

    Private picImage As New PictureBox()
    Private WithEvents btnZoomIn As New Button()
    Private WithEvents btnZoomOut As New Button()
    Private WithEvents btnReset As New Button()
    Private currentZoom As Double = 1.0

    Public Sub New(imageUrl As String)
        InitializeUI()
        LoadImage(imageUrl)
    End Sub

    Private Sub InitializeUI()
        Me.Text = "图片查看器"
        Me.Size = New Drawing.Size(800, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 50
        Me.Controls.Add(panelTop)

        btnZoomIn.Text = "放大"
        btnZoomIn.Location = New Drawing.Point(20, 10)
        btnZoomIn.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnZoomIn)

        btnZoomOut.Text = "缩小"
        btnZoomOut.Location = New Drawing.Point(120, 10)
        btnZoomOut.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnZoomOut)

        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(220, 10)
        btnReset.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnReset)

        picImage.Dock = DockStyle.Fill
        picImage.SizeMode = PictureBoxSizeMode.Zoom
        Me.Controls.Add(picImage)
    End Sub

    Private Sub LoadImage(url As String)
        Try
            picImage.Load(url)
        Catch ex As Exception
            ShowError("加载图片失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnZoomIn_Click(sender As Object, e As EventArgs) Handles btnZoomIn.Click
        currentZoom *= 1.2
        picImage.SizeMode = PictureBoxSizeMode.CenterImage
        picImage.Width = CInt(picImage.Image.Width * currentZoom)
        picImage.Height = CInt(picImage.Image.Height * currentZoom)
    End Sub

    Private Sub btnZoomOut_Click(sender As Object, e As EventArgs) Handles btnZoomOut.Click
        currentZoom /= 1.2
        picImage.SizeMode = PictureBoxSizeMode.CenterImage
        picImage.Width = CInt(picImage.Image.Width * currentZoom)
        picImage.Height = CInt(picImage.Image.Height * currentZoom)
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        currentZoom = 1.0
        picImage.SizeMode = PictureBoxSizeMode.Zoom
    End Sub
End Class

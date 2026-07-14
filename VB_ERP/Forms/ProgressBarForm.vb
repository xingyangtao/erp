' ============================================================================
' 进度条窗口
' 功能: 进度显示
' ============================================================================

Public Class ProgressBarForm
    Inherits System.Windows.Forms.Form

    Private progressBar As New ProgressBar()
    Private lblStatus As New Label()

    Public Sub New(Optional title As String = "处理中...")
        InitializeUI(title)
    End Sub

    Private Sub InitializeUI(title As String)
        Me.Text = title
        Me.Size = New Drawing.Size(350, 120)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        lblStatus.Text = "正在处理..."
        lblStatus.Location = New Drawing.Point(20, 15)
        lblStatus.Size = New Drawing.Size(300, 20)
        Me.Controls.Add(lblStatus)

        progressBar.Location = New Drawing.Point(20, 45)
        progressBar.Size = New Drawing.Size(300, 25)
        progressBar.Minimum = 0
        progressBar.Maximum = 100
        Me.Controls.Add(progressBar)
    End Sub

    Public Sub SetProgress(value As Integer, Optional status As String = "")
        progressBar.Value = Math.Min(value, 100)
        If Not String.IsNullOrEmpty(status) Then
            lblStatus.Text = status
        End If
        Application.DoEvents()
    End Sub

    Public Sub SetStatus(status As String)
        lblStatus.Text = status
        Application.DoEvents()
    End Sub

    ' ========== 兼容属性 ==========
    Public Property LabelText As String
        Get
            Return lblStatus.Text
        End Get
        Set(value As String)
            lblStatus.Text = value
            Application.DoEvents()
        End Set
    End Property

    Public Property MaxValue As Integer
        Get
            Return progressBar.Maximum
        End Get
        Set(value As Integer)
            progressBar.Maximum = value
        End Set
    End Property

    Public Shadows Property Value As Integer
        Get
            Return progressBar.Value
        End Get
        Set(value As Integer)
            progressBar.Value = Math.Min(value, progressBar.Maximum)
            Application.DoEvents()
        End Set
    End Property
End Class

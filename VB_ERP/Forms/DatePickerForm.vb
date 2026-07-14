' ============================================================================
' 日期选择器窗口
' 功能: 日期选择
' ============================================================================

Public Class DatePickerForm
    Inherits System.Windows.Forms.Form

    Private monthCalendar As New MonthCalendar()
    Public Property SelectedDate As String = ""

    Public Sub New()
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        Me.Text = "选择日期"
        Me.Size = New Drawing.Size(260, 230)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        monthCalendar.Location = New Drawing.Point(10, 10)
        monthCalendar.MaxSelectionCount = 1
        AddHandler monthCalendar.DateSelected, AddressOf monthCalendar_DateSelected
        Me.Controls.Add(monthCalendar)
    End Sub

    Private Sub monthCalendar_DateSelected(sender As Object, e As DateRangeEventArgs)
        SelectedDate = e.Start.ToString("yyyy-MM-dd")
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Public Overloads Function ShowDialog() As String
        If MyBase.ShowDialog() = DialogResult.OK Then
            Return SelectedDate
        End If
        Return ""
    End Function
End Class

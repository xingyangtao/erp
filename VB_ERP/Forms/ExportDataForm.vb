' ============================================================================
' 数据导出窗口
' 功能: 选择导出格式并导出数据
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class ExportDataForm
    Inherits System.Windows.Forms.Form

    Private cmbFormat As New ComboBox()
    Private txtFileName As New TextBox()
    Private WithEvents btnExport As New Button()
    Private WithEvents btnCancel As New Button()
    Private exportData As DataTable = Nothing

    Public Sub New(data As DataTable)
        exportData = data
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        Me.Text = "数据导出"
        Me.Size = New Drawing.Size(400, 200)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblFormat As New Label()
        lblFormat.Text = "导出格式："
        lblFormat.Location = New Drawing.Point(30, 20)
        lblFormat.AutoSize = True
        Me.Controls.Add(lblFormat)

        cmbFormat.Location = New Drawing.Point(110, 17)
        cmbFormat.Size = New Drawing.Size(150, 25)
        cmbFormat.Items.AddRange(New String() {"Excel (.xlsx)", "CSV (.csv)", "文本 (.txt)"})
        cmbFormat.SelectedIndex = 0
        Me.Controls.Add(cmbFormat)

        Dim lblFileName As New Label()
        lblFileName.Text = "文件名："
        lblFileName.Location = New Drawing.Point(30, 60)
        lblFileName.AutoSize = True
        Me.Controls.Add(lblFileName)

        txtFileName.Location = New Drawing.Point(110, 57)
        txtFileName.Size = New Drawing.Size(230, 25)
        txtFileName.Text = "导出数据_" & DateTime.Now.ToString("yyyyMMdd")
        Me.Controls.Add(txtFileName)

        btnExport.Text = "导出"
        btnExport.Location = New Drawing.Point(110, 110)
        btnExport.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnExport)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(230, 110)
        btnCancel.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If exportData Is Nothing OrElse exportData.Rows.Count = 0 Then
            ShowWarning("没有数据可导出！")
            Return
        End If

        Try
            Dim fileName As String = txtFileName.Text.Trim()
            If String.IsNullOrEmpty(fileName) Then
                ShowWarning("请输入文件名！")
                Return
            End If

            Select Case cmbFormat.SelectedIndex
                Case 0
                    ExportToExcel(exportData, fileName & ".xlsx")
                Case 1
                    ExportToCSV(exportData, fileName & ".csv")
                Case 2
                    ExportToText(exportData, fileName & ".txt")
            End Select

            ShowSuccess("导出成功！")
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            ShowError("导出失败：" & ex.Message)
        End Try
    End Sub

    Private Sub ExportToCSV(data As DataTable, fileName As String)
        Dim exportDir As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "biaoge")
        If Not IO.Directory.Exists(exportDir) Then
            IO.Directory.CreateDirectory(exportDir)
        End If

        Dim filePath As String = IO.Path.Combine(exportDir, fileName)
        Dim sb As New Text.StringBuilder()

        ' 写入表头
        Dim headers As New List(Of String)()
        For Each col As DataColumn In data.Columns
            headers.Add(col.ColumnName)
        Next
        sb.AppendLine(String.Join(",", headers))

        ' 写入数据
        For Each row As DataRow In data.Rows
            Dim values As New List(Of String)()
            For Each col As DataColumn In data.Columns
                values.Add("""" & row(col).ToString().Replace("""", """""") & """")
            Next
            sb.AppendLine(String.Join(",", values))
        Next

        IO.File.WriteAllText(filePath, sb.ToString(), System.Text.Encoding.UTF8)

        ' 打开文件
        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub

    Private Sub ExportToText(data As DataTable, fileName As String)
        Dim exportDir As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "biaoge")
        If Not IO.Directory.Exists(exportDir) Then
            IO.Directory.CreateDirectory(exportDir)
        End If

        Dim filePath As String = IO.Path.Combine(exportDir, fileName)
        Dim sb As New Text.StringBuilder()

        ' 写入表头
        Dim headers As New List(Of String)()
        For Each col As DataColumn In data.Columns
            headers.Add(col.ColumnName.PadRight(20))
        Next
        sb.AppendLine(String.Join("", headers))
        sb.AppendLine(New String("-"c, headers.Count * 20))

        ' 写入数据
        For Each row As DataRow In data.Rows
            Dim values As New List(Of String)()
            For Each col As DataColumn In data.Columns
                values.Add(row(col).ToString().PadRight(20))
            Next
            sb.AppendLine(String.Join("", values))
        Next

        IO.File.WriteAllText(filePath, sb.ToString(), System.Text.Encoding.UTF8)

        ' 打开文件
        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

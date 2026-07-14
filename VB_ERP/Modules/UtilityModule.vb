' ============================================================================
' 工具函数模块
' 提供通用的工具函数
' ============================================================================

Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports Newtonsoft.Json

Module UtilityModule

    ' ========== 文件操作 ==========
    Public Function ReadConfigValue(section As String, key As String, defaultValue As String) As String
        Dim configPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "config.ini")
        Return ReadIniFile(configPath, section, key, defaultValue)
    End Function

    Public Function ReadIniFile(filePath As String, section As String, key As String, defaultValue As String) As String
        If Not File.Exists(filePath) Then Return defaultValue

        Dim lines As String() = File.ReadAllLines(filePath)
        Dim currentSection As String = ""

        For Each line As String In lines
            Dim trimmedLine As String = line.Trim()
            If trimmedLine.StartsWith("[") AndAlso trimmedLine.EndsWith("]") Then
                currentSection = trimmedLine.Substring(1, trimmedLine.Length - 2)
            ElseIf currentSection = section AndAlso trimmedLine.Contains("=") Then
                Dim parts As String() = trimmedLine.Split(New Char() {"="c}, 2)
                If parts(0).Trim() = key Then
                    Return parts(1).Trim()
                End If
            End If
        Next

        Return defaultValue
    End Function

    Public Sub WriteConfigValue(section As String, key As String, value As String)
        Dim configPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "config.ini")
        WriteIniFile(configPath, section, key, value)
    End Sub

    Public Sub WriteIniFile(filePath As String, section As String, key As String, value As String)
        Dim lines As New List(Of String)()
        Dim currentSection As String = ""
        Dim found As Boolean = False

        If File.Exists(filePath) Then
            lines.AddRange(File.ReadAllLines(filePath))
        End If

        For i As Integer = 0 To lines.Count - 1
            Dim trimmedLine As String = lines(i).Trim()
            If trimmedLine.StartsWith("[") AndAlso trimmedLine.EndsWith("]") Then
                currentSection = trimmedLine.Substring(1, trimmedLine.Length - 2)
            ElseIf currentSection = section AndAlso trimmedLine.Contains("=") Then
                Dim parts As String() = trimmedLine.Split(New Char() {"="c}, 2)
                If parts(0).Trim() = key Then
                    lines(i) = $"{key}={value}"
                    found = True
                    Exit For
                End If
            End If
        Next

        If Not found Then
            Dim sectionFound As Boolean = False
            For i As Integer = 0 To lines.Count - 1
                If lines(i).Trim() = $"[{section}]" Then
                    lines.Insert(i + 1, $"{key}={value}")
                    sectionFound = True
                    Exit For
                End If
            Next
            If Not sectionFound Then
                lines.Add("")
                lines.Add($"[{section}]")
                lines.Add($"{key}={value}")
            End If
        End If

        Dim dirPath As String = Path.GetDirectoryName(filePath)
        If Not Directory.Exists(dirPath) Then
            Directory.CreateDirectory(dirPath)
        End If

        File.WriteAllLines(filePath, lines.ToArray())
    End Sub

    ' ========== JSON操作 ==========
    Public Function ToJson(obj As Object) As String
        Return JsonConvert.SerializeObject(obj, Formatting.None)
    End Function

    Public Function FromJson(Of T)(json As String) As T
        Return JsonConvert.DeserializeObject(Of T)(json)
    End Function

    ' ========== 日期操作 ==========
    Public Function FormatDate(dateValue As DateTime, format As String) As String
        Return dateValue.ToString(format)
    End Function

    Public Function ParseDate(dateStr As String) As DateTime
        Return DateTime.Parse(dateStr)
    End Function

    Public Function GetMonthStart() As String
        Return New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).ToString("yyyy-MM-dd")
    End Function

    Public Function GetMonthEnd() As String
        Dim monthStart As New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        Return monthStart.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
    End Function

    ' ========== 数值格式化 ==========
    Public Function FormatDecimal2(value As Decimal) As String
        Return value.ToString("F2")
    End Function

    Public Function FormatDecimal3(value As Decimal) As String
        Return value.ToString("F3")
    End Function

    Public Function SafeDecimal(value As Object) As Decimal
        If value Is Nothing OrElse value Is DBNull.Value Then Return 0D
        Dim result As Decimal
        If Decimal.TryParse(value.ToString(), result) Then
            Return result
        End If
        Return 0D
    End Function

    Public Function SafeString(value As Object) As String
        If value Is Nothing OrElse value Is DBNull.Value Then Return ""
        Return value.ToString()
    End Function

    ' ========== 表格操作 ==========
    Public Sub LoadDataGridView(dgv As DataGridView, dt As DataTable)
        dgv.DataSource = dt
        dgv.AutoResizeColumns()
    End Sub

    Public Sub AddTextBoxColumn(dgv As DataGridView, headerText As String, dataPropertyName As String,
                                 width As Integer, Optional alignment As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleLeft)
        Dim col As New DataGridViewTextBoxColumn()
        col.HeaderText = headerText
        col.DataPropertyName = dataPropertyName
        col.Width = width
        col.DefaultCellStyle.Alignment = alignment
        dgv.Columns.Add(col)
    End Sub

    ' ========== 导出Excel ==========
    Public Sub ExportToExcel(dt As DataTable, fileName As String)
        Try
            Using package As New OfficeOpenXml.ExcelPackage()
                Dim ws As OfficeOpenXml.ExcelWorksheet = package.Workbook.Worksheets.Add("Sheet1")

                ' 写入表头
                For col As Integer = 0 To dt.Columns.Count - 1
                    ws.Cells(1, col + 1).Value = dt.Columns(col).ColumnName
                Next

                ' 写入数据
                For row As Integer = 0 To dt.Rows.Count - 1
                    For col As Integer = 0 To dt.Columns.Count - 1
                        ws.Cells(row + 2, col + 1).Value = dt.Rows(row)(col)
                    Next
                Next

                ' 自动调整列宽
                ws.Cells.AutoFitColumns()

                ' 保存文件
                Dim exportDir As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "biaoge")
                If Not Directory.Exists(exportDir) Then
                    Directory.CreateDirectory(exportDir)
                End If

                Dim filePath As String = Path.Combine(exportDir, fileName)
                Dim file As New FileInfo(filePath)
                package.SaveAs(file)

                ' 打开文件
                Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
            End Using
        Catch ex As Exception
            MessageBox.Show("导出Excel失败：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ========== HTTP下载图片 ==========
    Public Async Function DownloadImage(url As String, savePath As String) As Task(Of Boolean)
        Try
            Using client As New Net.Http.HttpClient()
                Dim response = Await client.GetAsync(url)
                response.EnsureSuccessStatusCode()
                Dim bytes = Await response.Content.ReadAsByteArrayAsync()
                IO.File.WriteAllBytes(savePath, bytes)
                Return True
            End Using
        Catch
            Return False
        End Try
    End Function

    ' ========== 图片上传 ==========
    Public Async Function UploadImage(filePath As String, uploadUrl As String) As Task(Of String)
        Try
            Using client As New Net.Http.HttpClient()
                Using content As New Net.Http.MultipartFormDataContent()
                    Dim fileBytes = IO.File.ReadAllBytes(filePath)
                    Dim fileContent As New Net.Http.ByteArrayContent(fileBytes)
                    content.Add(fileContent, "file", IO.Path.GetFileName(filePath))
                    Dim response = Await client.PostAsync(uploadUrl, content)
                    response.EnsureSuccessStatusCode()
                    Return Await response.Content.ReadAsStringAsync()
                End Using
            End Using
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ' ========== 添加系统日志 ==========
    Public Sub AddSystemLog(logType As String, title As String, content As String)
        Dim sql As String = $"INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('{SafeSQL(logType)}', '{SafeSQL(title)}', '{SafeSQL(content)}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
        ExecuteCommand(sql)
    End Sub

    ' ========== 消息提示 ==========
    Public Sub ShowSuccess(message As String)
        MessageBox.Show(message, "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub ShowError(message As String)
        MessageBox.Show(message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Public Sub ShowWarning(message As String)
        MessageBox.Show(message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Public Sub ShowInfo(message As String)
        MessageBox.Show(message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Function ConfirmAction(message As String) As Boolean
        Return MessageBox.Show(message, "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes
    End Function

End Module

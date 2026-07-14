' ============================================================================
' 数据库连接模块
' 负责MySQL数据库连接、读写分离、SQL执行、事务处理
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.Text

Module DatabaseModule

    ' ========== 数据库连接对象 ==========
    Public MySQL_Auth As MySqlConnection          ' 授权库连接
    Public MySQL_Write As MySqlConnection         ' 业务写库连接
    Public MySQL_Read As MySqlConnection          ' 业务读库连接（通用）
    Public MySQL_ReadReport As MySqlConnection    ' 报表专用读连接
    Public MySQL_ReadTask As MySqlConnection      ' 任务专用读连接
    Public MySQL_ReadPrint As MySqlConnection     ' 打印专用读连接
    Public MySQL_ReadOrder As MySqlConnection     ' 订单查询专用读连接

    ' ========== 连接授权库 ==========
    Public Function ConnectAuthDB() As Boolean
        Try
            Dim connStr As String = "Server=127.0.0.1;Port=3306;Database=erpshouquan;Uid=erpshouquan;Pwd=erpshouquan;SslMode=None;CharSet=utf8mb4;"
            MySQL_Auth = New MySqlConnection(connStr)
            MySQL_Auth.Open()
            Return True
        Catch ex As Exception
            MessageBox.Show("授权库连接失败：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' ========== 连接业务库（读写分离） ==========
    Public Function ConnectBusinessDB(server As String, port As String, database As String,
                                       username As String, password As String) As Boolean
        Try
            Dim connStr As String = $"Server={server};Port={port};Database={database};Uid={username};Pwd={password};SslMode=None;CharSet=utf8mb4;"

            ' 创建写库连接
            MySQL_Write = New MySqlConnection(connStr)
            MySQL_Write.Open()

            ' 创建读库连接（通用）
            MySQL_Read = New MySqlConnection(connStr)
            MySQL_Read.Open()

            ' 创建报表专用读连接
            MySQL_ReadReport = New MySqlConnection(connStr)
            MySQL_ReadReport.Open()

            ' 创建任务专用读连接
            MySQL_ReadTask = New MySqlConnection(connStr)
            MySQL_ReadTask.Open()

            ' 创建打印专用读连接
            MySQL_ReadPrint = New MySqlConnection(connStr)
            MySQL_ReadPrint.Open()

            ' 创建订单查询专用读连接
            MySQL_ReadOrder = New MySqlConnection(connStr)
            MySQL_ReadOrder.Open()

            Return True
        Catch ex As Exception
            MessageBox.Show("业务库连接失败：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' ========== 执行SQL查询（返回DataTable） ==========
    Public Function ExecuteQuery(sql As String, Optional conn As MySqlConnection = Nothing) As DataTable
        If conn Is Nothing Then conn = MySQL_Read
        Dim dt As New DataTable()
        Try
            EnsureConnectionOpen(conn)
            Dim adapter As New MySqlDataAdapter(sql, conn)
            adapter.Fill(dt)
        Catch ex As Exception
            LogError("ExecuteQuery", sql, ex)
        End Try
        Return dt
    End Function

    ' ========== 执行SQL命令（INSERT/UPDATE/DELETE） ==========
    Public Function ExecuteCommand(sql As String, Optional conn As MySqlConnection = Nothing) As Integer
        If conn Is Nothing Then conn = MySQL_Write
        Try
            EnsureConnectionOpen(conn)
            Dim cmd As New MySqlCommand(sql, conn)
            Return cmd.ExecuteNonQuery()
        Catch ex As Exception
            LogError("ExecuteCommand", sql, ex)
            Return -1
        End Try
    End Function

    ' ========== 执行SQL查询（返回单个值） ==========
    Public Function ExecuteScalar(sql As String, Optional conn As MySqlConnection = Nothing) As Object
        If conn Is Nothing Then conn = MySQL_Read
        Try
            EnsureConnectionOpen(conn)
            Dim cmd As New MySqlCommand(sql, conn)
            Return cmd.ExecuteScalar()
        Catch ex As Exception
            LogError("ExecuteScalar", sql, ex)
            Return Nothing
        End Try
    End Function

    ' ========== 开始事务 ==========
    Public Function BeginTransaction(Optional conn As MySqlConnection = Nothing) As MySqlTransaction
        If conn Is Nothing Then conn = MySQL_Write
        EnsureConnectionOpen(conn)
        Return conn.BeginTransaction()
    End Function

    ' ========== 提交事务 ==========
    Public Sub CommitTransaction(trans As MySqlTransaction)
        Try
            trans.Commit()
        Catch ex As Exception
            LogError("CommitTransaction", "", ex)
            Throw
        End Try
    End Sub

    ' ========== 回滚事务 ==========
    Public Sub RollbackTransaction(trans As MySqlTransaction)
        Try
            trans.Rollback()
        Catch ex As Exception
            LogError("RollbackTransaction", "", ex)
        End Try
    End Sub

    ' ========== SQL文本安全处理（防注入） ==========
    Public Function SafeSQL(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Return text.Replace("'", "''")
    End Function

    ' ========== UTF8转GBK ==========
    Public Function UTF8ToGBK(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Try
            Dim utf8 As Encoding = Encoding.UTF8
            Dim gbk As Encoding = Encoding.GetEncoding("GBK")
            Dim utf8Bytes As Byte() = utf8.GetBytes(text)
            Dim gbkBytes As Byte() = Encoding.Convert(utf8, gbk, utf8Bytes)
            Return gbk.GetString(gbkBytes)
        Catch
            Return text
        End Try
    End Function

    ' ========== GBK转UTF8 ==========
    Public Function GBKToUTF8(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Try
            Dim utf8 As Encoding = Encoding.UTF8
            Dim gbk As Encoding = Encoding.GetEncoding("GBK")
            Dim gbkBytes As Byte() = gbk.GetBytes(text)
            Dim utf8Bytes As Byte() = Encoding.Convert(gbk, utf8, gbkBytes)
            Return utf8.GetString(utf8Bytes)
        Catch
            Return text
        End Try
    End Function

    ' ========== 确保连接已打开 ==========
    Private Sub EnsureConnectionOpen(conn As MySqlConnection)
        If conn.State <> ConnectionState.Open Then
            conn.Open()
        End If
    End Sub

    ' ========== 记录错误日志 ==========
    Private Sub LogError(method As String, sql As String, ex As Exception)
        Dim logPath As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "error.log")
        Dim logEntry As String = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {method}: {ex.Message}" & vbCrLf &
                                  $"SQL: {sql}" & vbCrLf &
                                  $"StackTrace: {ex.StackTrace}" & vbCrLf & vbCrLf
        IO.File.AppendAllText(logPath, logEntry)
    End Sub

    ' ========== 关闭所有连接 ==========
    Public Sub CloseAllConnections()
        CloseConnection(MySQL_Auth)
        CloseConnection(MySQL_Write)
        CloseConnection(MySQL_Read)
        CloseConnection(MySQL_ReadReport)
        CloseConnection(MySQL_ReadTask)
        CloseConnection(MySQL_ReadPrint)
        CloseConnection(MySQL_ReadOrder)
    End Sub

    Private Sub CloseConnection(conn As MySqlConnection)
        Try
            If conn IsNot Nothing AndAlso conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        Catch
        End Try
    End Sub

End Module

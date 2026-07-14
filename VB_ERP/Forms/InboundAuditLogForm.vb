' ============================================================================
' 入库审核日志窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_入库审核日志.form.e.txt
' 包含3个程序集变量、5个子程序
' 功能：显示入库订单的操作日志（xipunum_erp_store_log + xipunum_erp_user联查）
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class InboundAuditLogForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（3个） ==========
    Private mainWindowRowIndex As Integer = -1     ' 局部_主窗口行号
    Private orderId As String = ""                 ' 局部_订单id
    Private orderCode As String = ""               ' 局部_订单编码

    ' ========== 控件声明 ==========
    Private grpMain As New GroupBox()              ' 添加修改_分组框
    Private dgvLog As New DataGridView()           ' 超级列表框EX1
    Private btnClose As New Button()               ' 图片框EX4（关闭）
    Private btnRefresh As New Button()             ' 刷新按钮

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "入库审核日志"
        Me.Size = New Drawing.Size(900, 650)
        Me.StartPosition = FormStartPosition.CenterParent

        ' 分组框
        grpMain.Dock = DockStyle.Fill
        grpMain.Text = "入库订单操作日志"
        grpMain.Padding = New System.Windows.Forms.Padding(5)
        Me.Controls.Add(grpMain)

        ' 顶部按钮区
        Dim panelBtns As New Panel()
        panelBtns.Dock = DockStyle.Top
        panelBtns.Height = 40
        grpMain.Controls.Add(panelBtns)

        btnRefresh.Text = "刷新"
        btnRefresh.Location = New Drawing.Point(10, 8)
        btnRefresh.Size = New Drawing.Size(80, 28)
        panelBtns.Controls.Add(btnRefresh)

        btnClose.Text = "关闭"
        btnClose.Location = New Drawing.Point(100, 8)
        btnClose.Size = New Drawing.Size(80, 28)
        panelBtns.Controls.Add(btnClose)

        ' DataGridView
        dgvLog.Dock = DockStyle.Fill
        dgvLog.ReadOnly = True
        dgvLog.AllowUserToAddRows = False
        dgvLog.RowHeadersVisible = False
        dgvLog.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvLog.BackgroundColor = System.Drawing.Color.White
        grpMain.Controls.Add(dgvLog)

        ' 设置表头
        dgvLog.Columns.Add("seq", "序号")
        dgvLog.Columns.Add("conter", "操作类型")
        dgvLog.Columns.Add("user", "操作账户")
        dgvLog.Columns.Add("creationtime", "操作时间")
        dgvLog.Columns(0).Width = 60
        dgvLog.Columns(1).Width = 200
        dgvLog.Columns(2).Width = 200
        dgvLog.Columns(3).Width = 180

        ' 事件绑定
        AddHandler btnClose.Click, AddressOf BtnClose_Click
        AddHandler btnRefresh.Click, AddressOf BtnRefresh_Click
        AddHandler dgvLog.CellDoubleClick, AddressOf DgvLog_CellDoubleClick
    End Sub

    ' ========== _窗口_入库审核日志_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 从全局变量获取订单ID和编码（由调用方设置）
        If orderId = "" Then
            ShowWarning("入库订单ID为空，无法加载日志！")
            Me.Close()
            Return
        End If
        If orderCode = "" Then
            orderCode = "未知订单"
        End If

        grpMain.Text = "入库订单:" & orderCode & " 操作日志"
        LoadData()
    End Sub

    ' ========== 设置订单信息 ==========
    Public Sub SetOrderInfo(id As String, code As String)
        orderId = id
        orderCode = code
    End Sub

    ' ========== _窗口_入库审核日志_尺寸被改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        ' 分组框已经Dock=Fill，不需要手动调整
    End Sub

    ' ========== 子程序_加载数据 ==========
    Private Sub LoadData()
        dgvLog.Rows.Clear()

        If orderId = "" Then Return

        Try
            Dim sql As String = "SELECT a.conter AS conter,a.user AS user," &
                "COALESCE(b.name,'') AS name,a.creationtime AS creationtime " &
                "FROM xipunum_erp_store_log AS a LEFT JOIN xipunum_erp_user AS b ON b.user=a.user " &
                "WHERE a.order_id='" & SafeSQL(orderId) & "' ORDER BY a.creationtime DESC"
            Dim dt As DataTable = ExecuteQuery(sql, MySQL_Read)

            If dt Is Nothing Then
                ShowError("加载入库日志失败，请检查数据库连接！")
                Return
            End If

            If dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    Dim conter As String = SafeString(dt.Rows(i)("conter"))
                    Dim userAccount As String = SafeString(dt.Rows(i)("user"))
                    Dim userName As String = SafeString(dt.Rows(i)("name"))
                    Dim creationtime As String = SafeString(dt.Rows(i)("creationtime"))

                    ' 显示账户格式：account(name) 或 account
                    Dim displayUser As String = ""
                    If userName <> "" Then
                        displayUser = userAccount & "(" & userName & ")"
                    Else
                        displayUser = userAccount
                    End If

                    ' 序号：补零
                    Dim seq As String = (i + 1).ToString().PadLeft(dt.Rows.Count.ToString().Length, "0"c)

                    dgvLog.Rows.Add(seq, conter, displayUser, creationtime)
                Next
            Else
                ShowSuccess("该入库订单暂无操作日志！")
            End If
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    ' ========== _图片框EX4_鼠标左键单击（关闭） ==========
    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ' ========== 子程序_刷新日志 ==========
    Private Sub BtnRefresh_Click(sender As Object, e As EventArgs)
        If orderId = "" Then Return
        LoadData()
    End Sub

    ' ========== _超级列表框EX1_项目左键双击 ==========
    Private Sub DgvLog_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        Dim copyContent As String = SafeString(dgvLog.Rows(e.RowIndex).Cells(1).Value)
        If copyContent = "" Then Return
        Clipboard.SetText(copyContent)
        ShowSuccess("已复制: " & copyContent)
    End Sub

End Class

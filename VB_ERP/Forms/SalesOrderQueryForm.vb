' ============================================================================
' 商品销售订单查询窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品销售订单查询.form.e.txt
' 功能: 查询商品销售记录，支持单击选择和批量多选，选择后回调客退窗口
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class SalesOrderQueryForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量 ==========
    Private listControlHandle As Integer = 0          ' 副编码_超级列表框组件句柄

    ' ========== 公共属性 ==========
    Public Property ProductCode As String = ""         ' 接收客退窗口传递的商品编码
    Public Property ParentReturnForm As SalesReturnForm = Nothing  ' 父客退窗口引用

    ' ========== 控件声明 ==========
    Private dgvList As New DataGridView()              ' 副编码_超级列表框EX
    Private txtSalesNumber As New TextBox()            ' 编辑框_销售单号
    Private txtBatchCount As New TextBox()             ' 编辑框_批量多选
    Private WithEvents btnBatchSelect As New Button()  ' 按钮EX_批量多选
    Private WithEvents btnClose As New Button()        ' 图片框EX4 (关闭)
    Private WithEvents btnSelect As New Button()       ' 按钮EX1 (选择)
    Private WithEvents btnQuery As New Button()        ' 按钮EX4 (查询)
    Private grpMain As New GroupBox()                  ' 添加修改_分组框

    ' ========== 初始化 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
    End Sub

    Private Sub InitializeUI()
        Me.Text = "商品销售订单查询"
        Me.Size = New Drawing.Size(900, 670)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.Sizable

        ' 分组框
        grpMain.Text = ""
        grpMain.Size = New Drawing.Size(744, 568)
        grpMain.Location = New Drawing.Point(0, 0)
        Me.Controls.Add(grpMain)

        ' DataGridView - 副编码_超级列表框EX (9列)
        dgvList.Location = New Drawing.Point(-8, 57)
        dgvList.Size = New Drawing.Size(744, 455)
        dgvList.AllowUserToAddRows = False
        dgvList.AllowUserToDeleteRows = False
        dgvList.ReadOnly = True
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvList.MultiSelect = False
        dgvList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvList.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        AddHandler dgvList.CellClick, AddressOf DgvList_CellClick

        ' 定义9列
        dgvList.Columns.Add("col0", "序号")
        dgvList.Columns.Add("col1", "商品编码")
        dgvList.Columns.Add("col2", "副编码")
        dgvList.Columns.Add("col3", "销售单号")
        dgvList.Columns.Add("col4", "数量")
        dgvList.Columns.Add("col5", "金重")
        dgvList.Columns.Add("col6", "销售金额")
        dgvList.Columns.Add("col7", "销售时间")
        dgvList.Columns.Add("col8", "单据id")

        dgvList.Columns("col0").Width = 50
        dgvList.Columns("col1").Width = 120
        dgvList.Columns("col2").Width = 120
        dgvList.Columns("col3").Width = 140
        dgvList.Columns("col4").Width = 70
        dgvList.Columns("col5").Width = 70
        dgvList.Columns("col6").Width = 80
        dgvList.Columns("col7").Width = 140
        dgvList.Columns("col8").Width = 60

        ' 隐藏单据id列
        dgvList.Columns("col8").Visible = False

        grpMain.Controls.Add(dgvList)

        ' 关闭按钮 - 图片框EX4
        btnClose.Text = "X"
        btnClose.Location = New Drawing.Point(488, 8)
        btnClose.Size = New Drawing.Size(20, 20)
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.BackColor = Drawing.Color.Transparent
        btnClose.FlatAppearance.BorderSize = 0
        grpMain.Controls.Add(btnClose)

        ' 选择按钮 - 按钮EX1
        btnSelect.Text = "选择"
        btnSelect.Location = New Drawing.Point(280, 520)
        btnSelect.Size = New Drawing.Size(64, 30)
        btnSelect.Visible = False
        grpMain.Controls.Add(btnSelect)

        ' 销售单号编辑框
        txtSalesNumber.Location = New Drawing.Point(344, 520)
        txtSalesNumber.Size = New Drawing.Size(208, 30)
        txtSalesNumber.Visible = False
        grpMain.Controls.Add(txtSalesNumber)

        ' 查询按钮 - 按钮EX4
        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(552, 520)
        btnQuery.Size = New Drawing.Size(40, 30)
        btnQuery.Visible = False
        grpMain.Controls.Add(btnQuery)

        ' 批量多选数量编辑框
        txtBatchCount.Location = New Drawing.Point(592, 520)
        txtBatchCount.Size = New Drawing.Size(72, 30)
        txtBatchCount.Visible = False
        grpMain.Controls.Add(txtBatchCount)

        ' 批量多选按钮
        btnBatchSelect.Text = "批量多选"
        btnBatchSelect.Location = New Drawing.Point(664, 521)
        btnBatchSelect.Size = New Drawing.Size(72, 28)
        btnBatchSelect.Visible = False
        grpMain.Controls.Add(btnBatchSelect)
    End Sub

    ' ========== _窗口_商品销售订单查询_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 设置批量多选默认值
        txtBatchCount.Text = "0"

        ' 清空销售单号
        txtSalesNumber.Text = ""

        ' 清空列表
        dgvList.Rows.Clear()

        ' 执行SQL查询商品销售记录
        LoadSalesRecords()
    End Sub

    ' ========== 加载销售记录 ==========
    Private Sub LoadSalesRecords()
        Try
            Dim sql As String = "SELECT b.poduct_code as bianma, b.fu_code as fubianma, " &
                "c.settlement_number as danhao, " &
                "CAST(ROUND(a.quantity, 2) AS DECIMAL(10, 2)) as quantity, " &
                "CAST(ROUND(a.net_weight, 3) AS DECIMAL(10, 3)) as jinzhong, " &
                "CAST(ROUND(a.settlement, 2) AS DECIMAL(10, 2)) as jine, " &
                "a.creationtime as creationtime, c.id as danjuid " &
                "FROM xipunum_erp_outbound AS a " &
                "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
                "INNER JOIN xipunum_erp_outbound_order AS c ON c.id = a.order_id " &
                "WHERE (a.poduct_code = '" & SafeSQL(ProductCode) & "' " &
                "OR b.fu_code = '" & SafeSQL(ProductCode) & "') " &
                "AND a.kufang = '" & SafeSQL(UserDepartment) & "' " &
                "AND a.xsstate = '0' " &
                "ORDER BY a.creationtime DESC"

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            Dim rowCount As Integer = dt.Rows.Count
            Dim numLength As Integer = rowCount.ToString().Length

            For i As Integer = 0 To rowCount - 1
                Dim row As DataGridViewRow = dgvList.Rows(dgvList.Rows.Add())
                Dim rowNum As Integer = i + 1
                Dim seqStr As String = "000" & rowNum.ToString()
                If seqStr.Length > numLength Then
                    seqStr = seqStr.Substring(seqStr.Length - numLength)
                End If

                row.Cells(0).Value = seqStr
                row.Cells(1).Value = SafeString(dt.Rows(i)("bianma"))
                row.Cells(2).Value = SafeString(dt.Rows(i)("fubianma"))
                row.Cells(3).Value = SafeString(dt.Rows(i)("danhao"))
                row.Cells(4).Value = SafeString(dt.Rows(i)("quantity"))
                row.Cells(5).Value = SafeString(dt.Rows(i)("jinzhong"))
                row.Cells(6).Value = SafeString(dt.Rows(i)("jine"))
                row.Cells(7).Value = SafeString(dt.Rows(i)("creationtime"))
                row.Cells(8).Value = SafeString(dt.Rows(i)("danjuid"))
            Next
        Catch ex As Exception
            ShowError("加载销售记录失败：" & ex.Message)
        End Try
    End Sub

    ' ========== _副编码_超级列表框EX_项目左键单击 ==========
    Private Sub DgvList_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return

        ' 清空客退窗口的商品编码和订单id
        If ParentReturnForm IsNot Nothing Then
            ParentReturnForm.dataProductCode = ""
            ParentReturnForm.dataProductOrderId = ""

            ' 取商品编码(列1)和订单id(列8)
            ParentReturnForm.dataProductCode = SafeString(dgvList.Rows(e.RowIndex).Cells(1).Value)
            ParentReturnForm.dataProductOrderId = SafeString(dgvList.Rows(e.RowIndex).Cells(8).Value)

            ' 调用客退窗口的读取数据方法
            ParentReturnForm.ReadProductData()
        End If

        ' 销毁窗口
        Me.Close()
    End Sub

    ' ========== _窗口_商品销售订单查询_尺寸被改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        ' 居中分组框
        grpMain.Left = (Me.Width - grpMain.Width) \ 2
        grpMain.Top = (Me.Height - grpMain.Height) \ 2
    End Sub

    ' ========== _按钮EX_批量多选_鼠标左键单击 ==========
    Private Sub BtnBatchSelect_Click(sender As Object, e As EventArgs) Handles btnBatchSelect.Click
        ' 验证销售单号不能为空
        If txtSalesNumber.Text.Trim() = "" Then
            ShowWarning("销售单号不能为空！")
            txtSalesNumber.Focus()
            Return
        End If

        Dim customerReturnNumber As String = txtSalesNumber.Text.Trim()
        Dim selectedCount As Integer = 0
        Dim batchLimit As Integer = 0
        Integer.TryParse(txtBatchCount.Text.Trim(), batchLimit)

        Dim queryCount As Integer = dgvList.Rows.Count

        ' 从后往前遍历（与易语言一致：相减(查询信息数量, 商品查找计次)）
        For i As Integer = 1 To queryCount
            Dim rowIndex As Integer = queryCount - i
            If rowIndex < 0 Then Exit For

            If selectedCount < batchLimit Then
                ' 匹配销售单号（列3）
                If SafeString(dgvList.Rows(rowIndex).Cells(3).Value) = customerReturnNumber Then
                    selectedCount += 1

                    ' 设置客退窗口的商品编码和订单id
                    If ParentReturnForm IsNot Nothing Then
                        ParentReturnForm.dataProductCode = ""
                        ParentReturnForm.dataProductOrderId = ""

                        ParentReturnForm.dataProductCode = SafeString(dgvList.Rows(rowIndex).Cells(1).Value)
                        ParentReturnForm.dataProductOrderId = SafeString(dgvList.Rows(rowIndex).Cells(8).Value)

                        ' 调用客退窗口的读取数据方法
                        ParentReturnForm.ReadProductData()
                    End If
                End If
            End If
        Next

        ' 销毁窗口
        Me.Close()
    End Sub

    ' ========== _图片框EX4_鼠标左键单击 (关闭) ==========
    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ' ========== 辅助方法 ==========
    Private Function SafeString(ByVal value As Object) As String
        If value Is Nothing OrElse IsDBNull(value) Then Return ""
        Return value.ToString()
    End Function

    Private Function SafeSQL(ByVal value As String) As String
        If String.IsNullOrEmpty(value) Then Return ""
        Return value.Replace("'", "''")
    End Function

    Private Sub ShowWarning(ByVal msg As String)
        MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub ShowError(ByVal msg As String)
        MessageBox.Show(msg, "错误", MessageBoxButtons.OK, MessageBoxIcon.[Error])
    End Sub
End Class

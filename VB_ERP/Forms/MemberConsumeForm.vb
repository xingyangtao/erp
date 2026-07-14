' ============================================================================
' 会员订单消费数据窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_会员订单消费数据.form.e.txt
' 包含所有3个程序集变量、3个子程序、2个SQL查询
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class MemberConsumeForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（3个） ==========
    Private localOrderSelected As Integer = -1       ' 局部_订单是否选中
    Private localMemberCode As String = ""           ' 局部_会员信息编码

    ' ========== 控件声明（对应易语言窗口控件） ==========
    ' 超级列表框（ListView/DataGridView）
    Private WithEvents dgvList As New DataGridView() ' 超级列表框EX

    ' ========== 构造函数 ==========
    Public Sub New(Optional memberCode As String = "", Optional selectedRow As Integer = -1)
        localOrderSelected = selectedRow
        localMemberCode = memberCode
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "会员订单消费数据"
        Me.Size = New Drawing.Size(1200, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' 超级列表框EX → DataGridView
        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.AllowUserToDeleteRows = False
        dgvList.RowHeadersVisible = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    ' ========== _超级列表框EX_加载表头 ==========
    Private Sub LoadListViewHeader()
        ' 对应易语言：局部_表格头 = {"序号", "出库单号", "销售时间", "预售单号", "导购员", "销售金额", "实收金额", "销售克价", "结算净重", "成本工费", "参考工费", "销售工费", "销售附加费", "折扣"}
        dgvList.Columns.Clear()

        Dim headers() As String = {"序号", "出库单号", "销售时间", "预售单号", "导购员", "销售金额", "实收金额", "销售克价", "结算净重", "成本工费", "参考工费", "销售工费", "销售附加费", "折扣"}
        Dim widths() As Integer = {50, 140, 140, 140, 100, -1, -1, -1, -1, -1, -1, -1, -1, -1}
        Dim alignments() As Integer = {1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2}

        For i As Integer = 0 To headers.Length - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = headers(i)
            col.Name = "col" & i
            If widths(i) > 0 Then col.Width = widths(i)
            col.AutoSizeMode = If(widths(i) = -1, DataGridViewAutoSizeColumnMode.Fill, DataGridViewAutoSizeColumnMode.None)

            ' 对应易语言：对齐方式（1=左对齐，2=右对齐）
            If alignments(i) = 2 Then
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Else
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End If
            dgvList.Columns.Add(col)
        Next
    End Sub

    ' ========== _窗口_会员订单消费数据_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 对应易语言：获取选中会员信息编码
        localMemberCode = ""
        ' VB.NET中通过构造函数参数传入

        ' 对应易语言：_超级列表框EX_加载表头()
        LoadListViewHeader()

        ' 对应易语言：赋值(窗口Ex1.标题, 相加("会员:", 局部_会员信息编码, "消费记录"))
        Me.Text = "会员:" & localMemberCode & "消费记录"

        ' 对应易语言：_高级表格1_消费加载()
        LoadConsumeData()
    End Sub

    ' ========== _高级表格1_消费加载 ==========
    Private Sub LoadConsumeData()
        dgvList.Rows.Clear()

        ' 对应易语言：如果(等于(全局_账户查看权限, "全部")) → 全部权限的SQL
        Dim sql As String = ""

        If UserPermission = "全部" Then
            ' 对应易语言：SELECT a.settlement_number AS asettlement_number,a.presale_number AS apresale_number,b.xiao_amount AS bxiao_amount,b.settlement AS bsettlement,b.gold_price AS bgold_price,b.net_weight AS bnet_weight,b.basic_cost AS bbasic_cost,b.premium_cost AS bpremium_cost,b.sales_cost AS bsales_cost,b.sales_surcharge AS bsales_surcharge,b.zhekou AS bzhekou,b.creationtime AS bcreationtime,c.name as daogou FROM xipunum_erp_outbound_order AS a left JOIN xipunum_erp_outbound AS b ON b.order_id = a.id LEFT JOIN xipunum_erp_user AS c ON c.user = b.shopping_guide WHERE a.customer_code = '...' ORDER BY a.id DESC
            sql = "SELECT a.settlement_number AS asettlement_number,a.presale_number AS apresale_number," &
                "b.xiao_amount AS bxiao_amount,b.settlement AS bsettlement,b.gold_price AS bgold_price," &
                "b.net_weight AS bnet_weight,b.basic_cost AS bbasic_cost,b.premium_cost AS bpremium_cost," &
                "b.sales_cost AS bsales_cost,b.sales_surcharge AS bsales_surcharge,b.zhekou AS bzhekou," &
                "b.creationtime AS bcreationtime,c.name as daogou " &
                "FROM xipunum_erp_outbound_order AS a " &
                "left JOIN xipunum_erp_outbound AS b ON b.order_id = a.id " &
                "LEFT JOIN xipunum_erp_user AS c ON c.user = b.shopping_guide " &
                "WHERE a.customer_code = '" & SafeSQL(localMemberCode) & "' ORDER BY a.id DESC"
        Else
            ' 对应易语言：否则 → 带店铺权限的SQL（AND b.kufang='全局_账户分组id'）
            sql = "SELECT a.settlement_number AS asettlement_number,a.presale_number AS apresale_number," &
                "b.xiao_amount AS bxiao_amount,b.settlement AS bsettlement,b.gold_price AS bgold_price," &
                "b.net_weight AS bnet_weight,b.basic_cost AS bbasic_cost,b.premium_cost AS bpremium_cost," &
                "b.sales_cost AS bsales_cost,b.sales_surcharge AS bsales_surcharge,b.zhekou AS bzhekou," &
                "b.creationtime AS bcreationtime,c.name as daogou " &
                "FROM xipunum_erp_outbound_order AS a " &
                "left JOIN xipunum_erp_outbound AS b ON b.order_id = a.id " &
                "LEFT JOIN xipunum_erp_user AS c ON c.user = b.shopping_guide " &
                "WHERE a.customer_code = '" & SafeSQL(localMemberCode) & "' and b.kufang='" & SafeSQL(UserDepartment) & "' ORDER BY a.id DESC"
        End If

        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)

        ' 对应易语言：如果真(大于(数据详情数量, 0))
        If dt.Rows.Count > 0 Then
            ' 对应易语言：插入行数据
            Dim dataCount As Integer = dt.Rows.Count
            Dim paddedCount As String = CStr(dataCount)

            For i As Integer = 0 To dataCount - 1
                Dim row As DataRow = dt.Rows(i)

                ' 对应易语言：序号 = 右补零（如"001", "002"等）
                Dim seqNum As String = ("000" & CStr(i + 1))
                seqNum = seqNum.Substring(seqNum.Length - paddedCount.Length)

                dgvList.Rows.Add(
                    seqNum,
                    DatabaseModule.GBKToUTF8(SafeString(row("asettlement_number"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("bcreationtime"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("apresale_number"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("daogou"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("bxiao_amount"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("bsettlement"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("bgold_price"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("bnet_weight"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("bbasic_cost"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("bpremium_cost"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("bsales_cost"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("bsales_surcharge"))),
                    DatabaseModule.GBKToUTF8(SafeString(row("bzhekou")))
                )
            Next
        End If
    End Sub

End Class

' ============================================================================
' 会员信息合并窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_会员信息合并.form.e.txt
' 包含所有3个程序集变量、10个子程序、大量SQL操作
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class MemberMergeForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（3个） ==========
    Private setRow As Integer = 0               ' 集_行号
    Private setCol As Integer = 0               ' 集_列号

    ' ========== 控件声明（对应易语言窗口控件） ==========
    ' 分组框
    Private WithEvents grpHeader As New GroupBox()     ' 分组框_头部

    ' 编辑框
    Private WithEvents txtSearchInfo As New TextBox()  ' 查找信息编辑框

    ' 按钮
    Private WithEvents btnSearch As New Button()       ' 按钮_查询
    Private WithEvents btnReset As New Button()        ' 按钮_重置
    Private WithEvents btnMerge As New Button()        ' 按钮_合并

    ' 高级表格（DataGridView）
    Private WithEvents dgvMembers As New DataGridView() ' 高级表格1

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "会员信息合并"
        Me.Size = New Drawing.Size(1100, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MaximizeBox = False

        ' 分组框_头部
        grpHeader.Text = ""
        grpHeader.Location = New Drawing.Point(10, 10)
        grpHeader.Size = New Drawing.Size(Me.ClientSize.Width - 20, 45)
        Me.Controls.Add(grpHeader)

        ' 查找信息编辑框
        txtSearchInfo.Location = New Drawing.Point(10, 15)
        txtSearchInfo.Size = New Drawing.Size(400, 25)
        grpHeader.Controls.Add(txtSearchInfo)
        AddHandler txtSearchInfo.KeyDown, AddressOf txtSearchInfo_KeyDown

        ' 按钮_查询
        btnSearch.Text = "查询"
        btnSearch.Location = New Drawing.Point(420, 13)
        btnSearch.Size = New Drawing.Size(80, 30)
        grpHeader.Controls.Add(btnSearch)

        ' 按钮_重置
        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(510, 13)
        btnReset.Size = New Drawing.Size(80, 30)
        grpHeader.Controls.Add(btnReset)

        ' 按钮_合并
        btnMerge.Text = "合并"
        btnMerge.Location = New Drawing.Point(600, 13)
        btnMerge.Size = New Drawing.Size(80, 30)
        grpHeader.Controls.Add(btnMerge)

        ' 高级表格1（DataGridView）
        SetupDataGridView()
        dgvMembers.Location = New Drawing.Point(5, 65)
        dgvMembers.Size = New Drawing.Size(Me.ClientSize.Width - 10, Me.ClientSize.Height - 75)
        Me.Controls.Add(dgvMembers)
    End Sub

    ' ========== 设置DataGridView表头 ==========
    Private Sub SetupDataGridView()
        ' 对应易语言：高级表格1_加载表头
        dgvMembers.AllowUserToAddRows = False
        dgvMembers.AllowUserToDeleteRows = False
        dgvMembers.RowHeadersVisible = False
        dgvMembers.SelectionMode = DataGridViewSelectionMode.CellSelect
        dgvMembers.MultiSelect = False

        ' 对应易语言：局部_表格头
        dgvMembers.Columns.Clear()

        Dim colIndex As New DataGridViewTextBoxColumn()       ' 序号
        colIndex.HeaderText = "序号" : colIndex.DataPropertyName = "序号" : colIndex.Width = 50
        dgvMembers.Columns.Add(colIndex)

        Dim colCode As New DataGridViewTextBoxColumn()        ' 客户编码
        colCode.HeaderText = "客户编码" : colCode.DataPropertyName = "客户编码" : colCode.Width = 130
        dgvMembers.Columns.Add(colCode)

        Dim colMemberId As New DataGridViewTextBoxColumn()    ' 会员ID
        colMemberId.HeaderText = "会员ID" : colMemberId.DataPropertyName = "会员ID" : colMemberId.Width = 65
        dgvMembers.Columns.Add(colMemberId)

        Dim colName As New DataGridViewTextBoxColumn()        ' 会员姓名
        colName.HeaderText = "会员姓名" : colName.DataPropertyName = "会员姓名" : colName.Width = 80
        dgvMembers.Columns.Add(colName)

        Dim colTel As New DataGridViewTextBoxColumn()         ' 会员电话
        colTel.HeaderText = "会员电话" : colTel.DataPropertyName = "会员电话" : colTel.Width = 110
        dgvMembers.Columns.Add(colTel)

        Dim colJieLiao As New DataGridViewTextBoxColumn()     ' 结料(g)
        colJieLiao.HeaderText = "结料(g)" : colJieLiao.DataPropertyName = "结料" : colJieLiao.Width = 150
        dgvMembers.Columns.Add(colJieLiao)

        Dim colJieYuan As New DataGridViewTextBoxColumn()     ' 结款(元)
        colJieYuan.HeaderText = "结款(元)" : colJieYuan.DataPropertyName = "结款" : colJieYuan.Width = 150
        dgvMembers.Columns.Add(colJieYuan)

        Dim colBirthday As New DataGridViewTextBoxColumn()    ' 出生年月
        colBirthday.HeaderText = "出生年月" : colBirthday.DataPropertyName = "出生年月" : colBirthday.Width = 120
        dgvMembers.Columns.Add(colBirthday)

        Dim colAddress As New DataGridViewTextBoxColumn()     ' 联系地址
        colAddress.HeaderText = "联系地址" : colAddress.DataPropertyName = "联系地址" : colAddress.Width = 300
        dgvMembers.Columns.Add(colAddress)

        Dim colCreateTime As New DataGridViewTextBoxColumn()  ' 创建时间
        colCreateTime.HeaderText = "创建时间" : colCreateTime.DataPropertyName = "创建时间" : colCreateTime.Width = 140
        dgvMembers.Columns.Add(colCreateTime)

        Dim colMain As New DataGridViewCheckBoxColumn()       ' 主会员（选择型）
        colMain.HeaderText = "主会员" : colMain.Width = 65
        dgvMembers.Columns.Add(colMain)

        Dim colMerge As New DataGridViewCheckBoxColumn()      ' 合并会员（选择型）
        colMerge.HeaderText = "合并会员" : colMerge.Width = 65
        dgvMembers.Columns.Add(colMerge)

        Dim colDelete As New DataGridViewButtonColumn()       ' 操作（组件型-删除按钮）
        colDelete.HeaderText = "操作" : colDelete.Text = "删除" : colDelete.Width = 65
        colDelete.UseColumnTextForButtonValue = True
        dgvMembers.Columns.Add(colDelete)

        ' 对应易语言：置只读方式
        For i As Integer = 0 To 9
            dgvMembers.Columns(i).ReadOnly = True
        Next
        dgvMembers.Columns(10).ReadOnly = False
        dgvMembers.Columns(11).ReadOnly = False
        dgvMembers.Columns(12).ReadOnly = True

        AddHandler dgvMembers.CellContentClick, AddressOf dgvMembers_CellContentClick
        AddHandler dgvMembers.CellClick, AddressOf dgvMembers_CellClick_HandleDelete
    End Sub

    ' ========== _窗口_会员信息合并_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 对应易语言：高级表格1_加载表头()
        SetupDataGridView()
        btnMerge.Enabled = False
    End Sub

    ' ========== _窗口_会员信息合并_尺寸被改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        Dim nWidth As Integer = Me.ClientSize.Width
        Dim nHeight As Integer = Me.ClientSize.Height

        ' 对应易语言：分组框_头部尺寸调整
        grpHeader.Width = nWidth - 20
        grpHeader.Left = 10
        grpHeader.Top = 10
        grpHeader.Height = 45

        ' 对应易语言：高级表格1尺寸调整
        dgvMembers.Width = nWidth - 10
        dgvMembers.Left = 5
        dgvMembers.Top = 65
        dgvMembers.Height = nHeight - 75
    End Sub

    ' ========== 子程序_删除表格 ==========
    Private Function DeleteAllGridRows() As Boolean
        ' 对应易语言：删除所有数据行（保留表头）
        dgvMembers.Rows.Clear()
        Return True
    End Function

    ' ========== _高级表格1_光标位置改变（选择逻辑） ==========
    Private Sub dgvMembers_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return

        ' 对应易语言：如果真(等于(选中数据列数, 10)) → 点击"主会员"列
        If e.ColumnIndex = 10 Then
            ' 对应易语言：取消所有其他行的主会员选中，选中当前行
            For i As Integer = 0 To dgvMembers.Rows.Count - 1
                dgvMembers.Rows(i).Cells(10).Value = False
            Next
            dgvMembers.Rows(e.RowIndex).Cells(10).Value = True
            ' 对应易语言：同时取消当前行的合并会员选中
            dgvMembers.Rows(e.RowIndex).Cells(11).Value = False
        End If

        ' 对应易语言：如果真(等于(选中数据列数, 11)) → 点击"合并会员"列
        If e.ColumnIndex = 11 Then
            ' 对应易语言：如果主会员选中，取消主会员选中
            If CBool(dgvMembers.Rows(e.RowIndex).Cells(10).Value) Then
                dgvMembers.Rows(e.RowIndex).Cells(10).Value = False
            End If

            ' 对应易语言：切换合并会员选中状态
            Dim currentVal As Boolean = CBool(dgvMembers.Rows(e.RowIndex).Cells(11).Value)
            dgvMembers.Rows(e.RowIndex).Cells(11).Value = Not currentVal
        End If
    End Sub

    ' ========== _高级表格1_按钮被点击（删除行） ==========
    Private Sub dgvMembers_CellClick_HandleDelete(sender As Object, e As DataGridViewCellEventArgs)
        ' 对应易语言：_高级表格1_按钮被点击（删除按钮）
        If e.ColumnIndex = 12 AndAlso e.RowIndex >= 0 Then
            dgvMembers.Rows.RemoveAt(e.RowIndex)
            ' 对应易语言：重新编号序号
            For i As Integer = 0 To dgvMembers.Rows.Count - 1
                dgvMembers.Rows(i).Cells(0).Value = i.ToString()
            Next
        End If
    End Sub

    ' ========== _按钮_查询_被单击 ==========
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        ' 对应易语言：如果真(等于(查找信息编辑框.内容, ""))
        If txtSearchInfo.Text = "" Then
            MessageBox.Show("会员信息不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSearchInfo.Focus()
            Return
        End If

        btnMerge.Enabled = False
        MergeSearchData()
    End Sub

    ' ========== _按钮_重置_被单击 ==========
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        ' 对应易语言：重新执行_窗口_会员信息合并_创建完毕
        Form_Load(Me, EventArgs.Empty)
    End Sub

    ' ========== _查找信息编辑框_按下某键 ==========
    Private Sub txtSearchInfo_KeyDown(sender As Object, e As KeyEventArgs)
        ' 对应易语言：如果真(等于(键代码, #回车键))
        If e.KeyCode = Keys.Enter Then
            btnSearch_Click(btnSearch, EventArgs.Empty)
        End If
    End Sub

    ' ========== _高级表格1_合并数据（查询并填充表格） ==========
    Private Sub MergeSearchData()
        DeleteAllGridRows()

        Dim searchText As String = txtSearchInfo.Text

        ' 对应易语言中的完整会员查询SQL
        Dim sql As String = "SELECT a.memberid as amemberid,a.customer_code AS acustomer_code,a.NAME AS aname,a.tel AS atel," &
            "ROUND(IFNULL(c.cun_number, 0) - IFNULL(c.qian_number, 0), 3) AS jieyuliao," &
            "ROUND(IFNULL(d.cun_number, 0) - IFNULL(d.qian_number, 0), 2) AS jieyuyuan," &
            "a.shengri AS ashengri,a.dizhi AS adizhi,a.cjuser AS acjuser,a.creationtime AS acreationtime,b.NAME AS bname " &
            "FROM xipunum_erp_member AS a " &
            "INNER JOIN xipunum_erp_user AS b ON b.USER = a.cjuser " &
            "LEFT JOIN ( SELECT customer_code, SUM(CASE WHEN cunqu = '存' AND type = '料' THEN number ELSE 0 END) AS cun_number, SUM(CASE WHEN cunqu = '欠' AND type = '料' THEN number ELSE 0 END) AS qian_number FROM xipunum_erp_member_cq WHERE kufang in (" & UserShopPermission & ") AND type = '料' GROUP BY customer_code ) AS c ON c.customer_code = a.customer_code " &
            "LEFT JOIN ( SELECT customer_code, SUM(CASE WHEN cunqu = '存' AND type = '元' THEN number ELSE 0 END) AS cun_number, SUM(CASE WHEN cunqu = '欠' AND type = '元' THEN number ELSE 0 END) AS qian_number FROM xipunum_erp_member_cq WHERE kufang in (" & UserShopPermission & ") AND type = '元' GROUP BY customer_code ) AS d ON d.customer_code = a.customer_code " &
            "WHERE a.customer_code IN ( SELECT DISTINCT customer_code FROM (SELECT customer_code FROM xipunum_erp_outbound_order WHERE cjuser IN " & GlobalViewSQL & " UNION SELECT customer_code FROM xipunum_erp_retreat_order WHERE cjuser IN " & GlobalViewSQL & " UNION SELECT customer_code FROM xipunum_erp_presale_order WHERE cjuser IN " & GlobalViewSQL & " UNION SELECT customer_code FROM xipunum_erp_member WHERE cjuser IN " & GlobalViewSQL & " ) AS combined WHERE customer_code != '' AND customer_code IS NOT NULL ) " &
            "AND (a.customer_code LIKE '%" & SafeSQL(searchText) & "%' OR a.NAME LIKE '%" & SafeSQL(searchText) & "%' OR a.tel LIKE '%" & SafeSQL(searchText) & "%' OR a.dizhi LIKE '%" & SafeSQL(searchText) & "%') ORDER BY a.creationtime asc"

        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)

        ' 对应易语言：进度条（VB.NET简化为直接填充）
        Dim rowIndex As Integer = 0
        For Each row As DataRow In dt.Rows
            dgvMembers.Rows.Add()
            dgvMembers.Rows(rowIndex).Cells(0).Value = rowIndex.ToString()
            dgvMembers.Rows(rowIndex).Cells(1).Value = DatabaseModule.GBKToUTF8(SafeString(row("acustomer_code")))
            dgvMembers.Rows(rowIndex).Cells(2).Value = DatabaseModule.GBKToUTF8(SafeString(row("amemberid")))
            dgvMembers.Rows(rowIndex).Cells(3).Value = DatabaseModule.GBKToUTF8(SafeString(row("aname")))
            dgvMembers.Rows(rowIndex).Cells(4).Value = DatabaseModule.GBKToUTF8(SafeString(row("atel")))
            dgvMembers.Rows(rowIndex).Cells(5).Value = DatabaseModule.GBKToUTF8(SafeString(row("jieyuliao")))
            dgvMembers.Rows(rowIndex).Cells(6).Value = DatabaseModule.GBKToUTF8(SafeString(row("jieyuyuan")))
            dgvMembers.Rows(rowIndex).Cells(7).Value = DatabaseModule.GBKToUTF8(SafeString(row("ashengri")))
            dgvMembers.Rows(rowIndex).Cells(8).Value = DatabaseModule.GBKToUTF8(SafeString(row("adizhi")))
            dgvMembers.Rows(rowIndex).Cells(9).Value = DatabaseModule.GBKToUTF8(SafeString(row("acreationtime")))
            dgvMembers.Rows(rowIndex).Cells(10).Value = False
            dgvMembers.Rows(rowIndex).Cells(11).Value = False
            rowIndex += 1
        Next

        ' 对应易语言：置只读方式和对齐方式
        txtSearchInfo.Text = ""
        txtSearchInfo.Focus()

        ' 对应易语言：如果行数>1，启用合并按钮
        If dgvMembers.Rows.Count >= 2 Then
            btnMerge.Enabled = True
        End If
    End Sub

    ' ========== _按钮_合并_被单击 ==========
    Private Sub btnMerge_Click(sender As Object, e As EventArgs) Handles btnMerge.Click
        ' 对应易语言：如果真(小于或等于(高级表格1.行数, 2))
        If dgvMembers.Rows.Count <= 2 Then
            MessageBox.Show("合并款号数量不能小于2！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 对应易语言：查找主会员（选中主会员列的行）
        Dim queryMainMember As String = ""
        For i As Integer = 0 To dgvMembers.Rows.Count - 1
            If CBool(dgvMembers.Rows(i).Cells(10).Value) Then
                queryMainMember = DatabaseModule.UTF8ToGBK(SafeString(dgvMembers.Rows(i).Cells(1).Value))
                Exit For
            End If
        Next

        If queryMainMember = "" Then
            MessageBox.Show("请选择合并数据主会员！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 对应易语言：计算合并会员数量
        Dim mergeCount As Integer = 0
        For i As Integer = 0 To dgvMembers.Rows.Count - 1
            If CBool(dgvMembers.Rows(i).Cells(11).Value) Then
                mergeCount += 1
            End If
        Next

        If mergeCount = 0 Then
            MessageBox.Show("请选择合并数据需要合并的会员！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 对应易语言：信息框确认（返回1=否，返回其他=是）
        Dim result As DialogResult = MessageBox.Show("确定合并所选的会员数据？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If result = DialogResult.No Then Return

        Dim globalInfoOperationDate As String = DateTime.Now.ToString("yyyy-MM-dd") & " " & DateTime.Now.ToString("HH:mm:ss")
        Dim globalInfoOperationAccount As String = UserAccount

        ' 对应易语言：遍历合并会员行
        For i As Integer = 0 To dgvMembers.Rows.Count - 1
            If CBool(dgvMembers.Rows(i).Cells(11).Value) Then
                Dim modifyMemberInfo As String = SafeString(dgvMembers.Rows(i).Cells(1).Value)
                Dim modifyMemberId As String = SafeString(dgvMembers.Rows(i).Cells(2).Value)

                ' ======== 对应易语言：合并积分数据 ========
                Dim scoreLogSql As String = "SELECT * FROM xipunum_erp_member_score_log WHERE customer_code='" & SafeSQL(modifyMemberInfo) & "' order by id desc"
                Dim scoreLogDt As DataTable = DatabaseModule.ExecuteQuery(scoreLogSql, MySQL_Read)

                Dim scoreLogInsert As String = ""
                For Each scoreRow As DataRow In scoreLogDt.Rows
                    Dim scoreId As String = SafeString(scoreRow("id"))
                    Dim scoreNumber As String = SafeString(scoreRow("settlement_number"))

                    DatabaseModule.ExecuteCommand("UPDATE xipunum_erp_member_score_log SET customer_code= '" & SafeSQL(queryMainMember) & "' WHERE id ='" & SafeSQL(scoreId) & "' LIMIT 1")

                    Dim logContent As String = "账户:" & UserAccount & " 合并积分订单:" & scoreNumber & " 积分数据从会员:" & modifyMemberInfo & "->" & queryMainMember
                    scoreLogInsert &= "( '修改','合并会员信息',' " & logContent & "','" & globalInfoOperationAccount & "','" & globalInfoOperationDate & "' ),"
                Next

                If scoreLogInsert <> "" Then
                    scoreLogInsert = scoreLogInsert.Substring(0, scoreLogInsert.Length - 1)
                    DatabaseModule.ExecuteCommand("INSERT INTO xipunum_erp_xitong_log ( type, title, conter, user, creationtime ) VALUES " & scoreLogInsert & ";")
                End If

                ' ======== 对应易语言：合并存取数据 ========
                Dim cqSql As String = "SELECT * FROM xipunum_erp_member_cq WHERE customer_code='" & SafeSQL(modifyMemberInfo) & "' order by id desc"
                Dim cqDt As DataTable = DatabaseModule.ExecuteQuery(cqSql, MySQL_Read)

                Dim cqInsert As String = ""
                For Each cqRow As DataRow In cqDt.Rows
                    Dim cqId As String = SafeString(cqRow("id"))
                    DatabaseModule.ExecuteCommand("UPDATE xipunum_erp_member_cq SET customer_code= '" & SafeSQL(queryMainMember) & "',updatetime= '" & SafeSQL(globalInfoOperationDate) & "'  WHERE id ='" & SafeSQL(cqId) & "' LIMIT 1")
                Next

                Dim cqLogContent As String = "账户:" & UserAccount & " 合并会员存取数据 从会员:" & modifyMemberInfo & "->" & queryMainMember
                cqInsert = "( '修改','合并会员信息',' " & cqLogContent & "','" & globalInfoOperationAccount & "','" & globalInfoOperationDate & "' )"
                DatabaseModule.ExecuteCommand("INSERT INTO xipunum_erp_xitong_log ( type, title, conter, user, creationtime ) VALUES " & cqInsert & ";")

                ' ======== 对应易语言：合并预售数据 ========
                Dim presaleSql As String = "SELECT * FROM xipunum_erp_presale_order WHERE customer_code='" & SafeSQL(modifyMemberInfo) & "' order by id desc"
                Dim presaleDt As DataTable = DatabaseModule.ExecuteQuery(presaleSql, MySQL_Read)

                Dim presaleInsert As String = ""
                For Each presaleRow As DataRow In presaleDt.Rows
                    Dim presaleId As String = SafeString(presaleRow("id"))
                    Dim presaleNumber As String = SafeString(presaleRow("presale_umber"))

                    DatabaseModule.ExecuteCommand("UPDATE xipunum_erp_presale_order SET customer_code= '" & SafeSQL(queryMainMember) & "',updatetime= '" & SafeSQL(globalInfoOperationDate) & "'  WHERE id ='" & SafeSQL(presaleId) & "' LIMIT 1")

                    Dim logContent As String = "账户:" & UserAccount & " 合并预售订单:" & presaleNumber & " 预售数据从会员:" & modifyMemberInfo & "->" & queryMainMember
                    presaleInsert &= "( '修改','合并会员信息',' " & logContent & "','" & globalInfoOperationAccount & "','" & globalInfoOperationDate & "' ),"
                Next

                If presaleInsert <> "" Then
                    presaleInsert = presaleInsert.Substring(0, presaleInsert.Length - 1)
                    DatabaseModule.ExecuteCommand("INSERT INTO xipunum_erp_xitong_log ( type, title, conter, user, creationtime ) VALUES " & presaleInsert & ";")
                End If

                ' ======== 对应易语言：合并回收数据 ========
                Dim retreatSql As String = "SELECT * FROM xipunum_erp_retreat_order WHERE customer_code='" & SafeSQL(modifyMemberInfo) & "' order by id desc"
                Dim retreatDt As DataTable = DatabaseModule.ExecuteQuery(retreatSql, MySQL_Read)

                Dim retreatInsert As String = ""
                For Each retreatRow As DataRow In retreatDt.Rows
                    Dim retreatId As String = SafeString(retreatRow("id"))
                    Dim retreatNumber As String = SafeString(retreatRow("retrea_umber"))

                    DatabaseModule.ExecuteCommand("UPDATE xipunum_erp_retreat_order SET customer_code= '" & SafeSQL(queryMainMember) & "',updatetime= '" & SafeSQL(globalInfoOperationDate) & "'  WHERE id ='" & SafeSQL(retreatId) & "' LIMIT 1")

                    Dim logContent As String = "账户:" & UserAccount & " 合并回收订单:" & retreatNumber & " 回收数据从会员:" & modifyMemberInfo & "->" & queryMainMember
                    retreatInsert &= "( '修改','合并会员信息',' " & logContent & "','" & globalInfoOperationAccount & "','" & globalInfoOperationDate & "' ),"
                Next

                If retreatInsert <> "" Then
                    retreatInsert = retreatInsert.Substring(0, retreatInsert.Length - 1)
                    DatabaseModule.ExecuteCommand("INSERT INTO xipunum_erp_xitong_log ( type, title, conter, user, creationtime ) VALUES " & retreatInsert & ";")
                End If

                ' ======== 对应易语言：合并销售数据 ========
                Dim outboundSql As String = "SELECT * FROM xipunum_erp_outbound_order WHERE customer_code='" & SafeSQL(modifyMemberInfo) & "' order by id desc"
                Dim outboundDt As DataTable = DatabaseModule.ExecuteQuery(outboundSql, MySQL_Read)

                Dim outboundInsert As String = ""
                For Each outboundRow As DataRow In outboundDt.Rows
                    Dim outboundId As String = SafeString(outboundRow("id"))
                    Dim outboundNumber As String = SafeString(outboundRow("settlement_number"))

                    DatabaseModule.ExecuteCommand("UPDATE xipunum_erp_outbound_order SET customer_code= '" & SafeSQL(queryMainMember) & "',updatetime= '" & SafeSQL(globalInfoOperationDate) & "'  WHERE id ='" & SafeSQL(outboundId) & "' LIMIT 1")

                    Dim logContent As String = "账户:" & UserAccount & " 合并销售订单:" & outboundNumber & " 回收数据从会员:" & modifyMemberInfo & "->" & queryMainMember
                    outboundInsert &= "( '修改','合并会员信息',' " & logContent & "','" & globalInfoOperationAccount & "','" & globalInfoOperationDate & "' ),"
                Next

                If outboundInsert <> "" Then
                    outboundInsert = outboundInsert.Substring(0, outboundInsert.Length - 1)
                    DatabaseModule.ExecuteCommand("INSERT INTO xipunum_erp_xitong_log ( type, title, conter, user, creationtime ) VALUES " & outboundInsert & ";")
                End If

                ' ======== 对应易语言：合并线上数据 ========
                Dim onlineSql As String = "SELECT * FROM xipunum_erp_online_order WHERE customer_code='" & SafeSQL(modifyMemberInfo) & "' order by id desc"
                Dim onlineDt As DataTable = DatabaseModule.ExecuteQuery(onlineSql, MySQL_Read)

                Dim onlineInsert As String = ""
                For Each onlineRow As DataRow In onlineDt.Rows
                    Dim onlineId As String = SafeString(onlineRow("id"))
                    Dim onlineNumber As String = SafeString(onlineRow("order_no"))

                    DatabaseModule.ExecuteCommand("UPDATE xipunum_erp_online_order SET customer_code= '" & SafeSQL(queryMainMember) & "' WHERE id ='" & SafeSQL(onlineId) & "' LIMIT 1")

                    Dim logContent As String = "账户:" & UserAccount & " 合并线上订单:" & onlineNumber & " 线上数据从会员:" & modifyMemberInfo & "->" & queryMainMember
                    onlineInsert &= "( '修改','合并会员信息',' " & logContent & "','" & globalInfoOperationAccount & "','" & globalInfoOperationDate & "' ),"
                Next

                If onlineInsert <> "" Then
                    onlineInsert = onlineInsert.Substring(0, onlineInsert.Length - 1)
                    DatabaseModule.ExecuteCommand("INSERT INTO xipunum_erp_xitong_log ( type, title, conter, user, creationtime ) VALUES " & onlineInsert & ";")
                End If

                ' ======== 对应易语言：删除合并后会员账户 ========
                DatabaseModule.ExecuteCommand("delete from xipunum_erp_member where customer_code= '" & SafeSQL(modifyMemberInfo) & "' and memberid= '" & SafeSQL(modifyMemberId) & "'")

                Dim deleteLogContent As String = "账户:" & UserAccount & " 删除合并后会员账户:" & modifyMemberInfo & " id:" & modifyMemberId
                DatabaseModule.ExecuteCommand("INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('删除','合并会员信息','" & SafeSQL(deleteLogContent) & "','" & SafeSQL(globalInfoOperationAccount) & "','" & SafeSQL(globalInfoOperationDate) & "')")
            End If
        Next

        MessageBox.Show("合并会员至" & queryMainMember & "成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
        btnReset_Click(btnReset, EventArgs.Empty)
    End Sub

End Class

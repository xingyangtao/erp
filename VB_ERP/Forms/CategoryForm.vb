' ============================================================================
' 商品品类管理窗口
' 功能: 管理商品品类的新增、修改
' 100% 匹配易语言 窗口_商品品类管理 功能
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports VB_ERP.Modules

Public Class CategoryForm
    Inherits System.Windows.Forms.Form

#Region "程序集变量"

    Private 集_操作模式 As String = ""        ' "新增" 或 "修改"
    Private 集_操作id As String = ""          ' 修改时的 id

    ' 列表
    Private 超级列表框Ex_品类管理 As New ListView()

#End Region

#Region "控件声明"

    Private WithEvents 按钮EX_查询 As New Button()
    Private WithEvents 按钮EX_重置 As New Button()
    Private WithEvents 按钮EX_保存 As New Button()
    Private WithEvents 按钮EX4 As New Button()        ' 新增

    Private 编辑框_品类名称 As New TextBox()
    Private 编辑框_积分比例 As New TextBox()

    ' 材质组合框
    Private 组合框_材质 As New ComboBox()

    ' 成色多选框
    Private 选择框_一号色 As New CheckBox()
    Private 选择框_二号色 As New CheckBox()
    Private 选择框_三号色 As New CheckBox()
    Private 选择框_四号色 As New CheckBox()
    Private 选择框_五号色 As New CheckBox()

    Private WithEvents 单选框_镶嵌 As New RadioButton()
    Private WithEvents 单选框_非镶嵌 As New RadioButton()
    Private WithEvents 单选框_是 As New RadioButton()
    Private WithEvents 单选框_否 As New RadioButton()
    Private WithEvents 单选框_必填 As New RadioButton()
    Private WithEvents 单选框_非必填 As New RadioButton()

    Private 图片框EX1 As New PictureBox()        ' 材质容器 (含成色checkbox)
    Private 图片框EX3 As New PictureBox()        ' 镶嵌容器 (含是/否)

    ' 分组框
    Private 添加修改_分组框 As New GroupBox()

#End Region

#Region "初始化"

    Public Sub New()
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        Me.Text = "商品品类管理"
        Me.Size = New Drawing.Size(1246, 659)
        Me.StartPosition = FormStartPosition.CenterParent

        ' 主面板
        添加修改_分组框.Text = "品类管理"
        添加修改_分组框.Size = New Drawing.Size(1220, 620)
        添加修改_分组框.Location = New Drawing.Point(10, 10)
        Me.Controls.Add(添加修改_分组框)

        ' 列表
        超级列表框Ex_品类管理.View = View.Details
        超级列表框Ex_品类管理.FullRowSelect = True
        超级列表框Ex_品类管理.GridLines = True
        超级列表框Ex_品类管理.Columns.Add("ID", 50)
        超级列表框Ex_品类管理.Columns.Add("品类名称", 120)
        超级列表框Ex_品类管理.Columns.Add("材质", 150)
        超级列表框Ex_品类管理.Columns.Add("成色", 200)
        超级列表框Ex_品类管理.Columns.Add("镶嵌", 60)
        超级列表框Ex_品类管理.Columns.Add("多数量", 60)
        超级列表框Ex_品类管理.Columns.Add("克价/积分", 80)
        超级列表框Ex_品类管理.Columns.Add("积分比例", 80)
        超级列表框Ex_品类管理.Location = New Drawing.Point(10, 80)
        超级列表框Ex_品类管理.Size = New Drawing.Size(800, 530)
        添加修改_分组框.Controls.Add(超级列表框Ex_品类管理)

        ' 右侧编辑面板
        Dim editX As Integer = 820
        Dim editY As Integer = 80

        ' 按钮EX4 - 新增
        按钮EX4.Text = "新增"
        按钮EX4.Location = New Drawing.Point(editX, editY)
        按钮EX4.Size = New Drawing.Size(72, 30)
        添加修改_分组框.Controls.Add(按钮EX4)

        ' 按钮EX5 - 删除/重置
        ' (实际 EasyLanguage 用 重置 按钮刷新)

        ' 按钮EX_查询
        按钮EX_查询.Text = "查询"
        按钮EX_查询.Location = New Drawing.Point(editX + 77, editY)
        按钮EX_查询.Size = New Drawing.Size(72, 30)
        添加修改_分组框.Controls.Add(按钮EX_查询)

        ' 按钮EX_重置
        按钮EX_重置.Text = "重置"
        按钮EX_重置.Location = New Drawing.Point(editX + 154, editY)
        按钮EX_重置.Size = New Drawing.Size(72, 30)
        添加修改_分组框.Controls.Add(按钮EX_重置)

        ' 品类名称
        Dim y1 As Integer = editY + 35
        AddEditLabel("品类名称:", editX, y1)
        编辑框_品类名称.Location = New Drawing.Point(editX + 77, y1)
        编辑框_品类名称.Size = New Drawing.Size(170, 25)
        添加修改_分组框.Controls.Add(编辑框_品类名称)

        ' 材质选择
        Dim y2 As Integer = y1 + 30
        AddEditLabel("材质:", editX, y2)
        组合框_材质.Location = New Drawing.Point(editX + 77, y2)
        组合框_材质.Size = New Drawing.Size(170, 25)
        组合框_材质.DropDownStyle = ComboBoxStyle.DropDown
        添加修改_分组框.Controls.Add(组合框_材质)

        ' 成色复选框
        Dim y3 As Integer = y2 + 30
        ' (注意: EasyLanguage 中成色在图片框EX1上，但这里直接放panel)
        Dim colorPanel As New Panel()
        colorPanel.Location = New Drawing.Point(editX + 77, y3)
        colorPanel.Size = New Drawing.Size(200, 90)
        添加修改_分组框.Controls.Add(colorPanel)

        选择框_一号色.Text = "一号色"
        选择框_一号色.Location = New Drawing.Point(5, 5)
        colorPanel.Controls.Add(选择框_一号色)
        选择框_二号色.Text = "二号色"
        选择框_二号色.Location = New Drawing.Point(80, 5)
        colorPanel.Controls.Add(选择框_二号色)
        选择框_三号色.Text = "三号色"
        选择框_三号色.Location = New Drawing.Point(5, 30)
        colorPanel.Controls.Add(选择框_三号色)
        选择框_四号色.Text = "四号色"
        选择框_四号色.Location = New Drawing.Point(80, 30)
        colorPanel.Controls.Add(选择框_四号色)
        选择框_五号色.Text = "五号色"
        选择框_五号色.Location = New Drawing.Point(5, 55)
        colorPanel.Controls.Add(选择框_五号色)

        ' 镶嵌选择
        Dim y4 As Integer = y3 + 95
        AddEditLabel("镶嵌:", editX, y4)
        单选框_镶嵌.Text = "镶嵌"
        单选框_镶嵌.Location = New Drawing.Point(editX + 50, y4)
        单选框_镶嵌.Size = New Drawing.Size(60, 20)
        添加修改_分组框.Controls.Add(单选框_镶嵌)
        单选框_非镶嵌.Text = "非镶嵌"
        单选框_非镶嵌.Location = New Drawing.Point(editX + 110, y4)
        单选框_非镶嵌.Size = New Drawing.Size(70, 20)
        添加修改_分组框.Controls.Add(单选框_非镶嵌)

        ' 多数量
        Dim y5 As Integer = y4 + 30
        AddEditLabel("多数量:", editX, y5)
        单选框_是.Text = "是"
        单选框_是.Location = New Drawing.Point(editX + 60, y5)
        单选框_是.Size = New Drawing.Size(50, 20)
        添加修改_分组框.Controls.Add(单选框_是)
        单选框_否.Text = "否"
        单选框_否.Location = New Drawing.Point(editX + 110, y5)
        单选框_否.Size = New Drawing.Size(50, 20)
        添加修改_分组框.Controls.Add(单选框_否)

        ' 克价必填
        Dim y6 As Integer = y5 + 30
        AddEditLabel("克价必填:", editX, y6)
        单选框_必填.Text = "必填"
        单选框_必填.Location = New Drawing.Point(editX + 65, y6)
        单选框_必填.Size = New Drawing.Size(55, 20)
        添加修改_分组框.Controls.Add(单选框_必填)
        单选框_非必填.Text = "非必填"
        单选框_非必填.Location = New Drawing.Point(editX + 120, y6)
        单选框_非必填.Size = New Drawing.Size(70, 20)
        添加修改_分组框.Controls.Add(单选框_非必填)

        ' 积分比例
        Dim y7 As Integer = y6 + 30
        AddEditLabel("积分比例:", editX, y7)
        编辑框_积分比例.Location = New Drawing.Point(editX + 77, y7)
        编辑框_积分比例.Size = New Drawing.Size(112, 25)
        添加修改_分组框.Controls.Add(编辑框_积分比例)

        ' 保存按钮
        按钮EX_保存.Text = "保存"
        按钮EX_保存.Location = New Drawing.Point(editX + 77, y7 + 40)
        按钮EX_保存.Size = New Drawing.Size(80, 30)
        添加修改_分组框.Controls.Add(按钮EX_保存)

        ' 事件
        AddHandler Me.Load, AddressOf 窗口_商品品类管理_创建完毕
        AddHandler Me.Resize, AddressOf 窗口_商品品类管理_尺寸被改变
        AddHandler 超级列表框Ex_品类管理.Click, AddressOf 列表项目单击
    End Sub

    Private Sub AddEditLabel(text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y + 4)
        lbl.AutoSize = True
        lbl.TextAlign = ContentAlignment.MiddleRight
        添加修改_分组框.Controls.Add(lbl)
    End Sub

#End Region

#Region "创建完毕 & 尺寸"

    Private Sub 窗口_商品品类管理_创建完毕(sender As Object, e As EventArgs)
        重置状态()
        列表加载()
    End Sub

    Private Sub 窗口_商品品类管理_尺寸被改变(sender As Object, e As EventArgs)
        If 添加修改_分组框 IsNot Nothing Then
            添加修改_分组框.Width = Me.ClientSize.Width - 20
            添加修改_分组框.Height = Me.ClientSize.Height - 20
            超级列表框Ex_品类管理.Height = 添加修改_分组框.Height - 90
        End If
    End Sub

#End Region

#Region "列表操作"

    Private Sub 列表加载()
        超级列表框Ex_品类管理.Items.Clear()

        Dim sql As String = "SELECT * FROM xipunum_erp_category where 1=1 order by id ASC"
        Dim dt As DataTable = DatabaseModule.MySQL_Read(sql)
        If dt Is Nothing Then Return

        For Each row As DataRow In dt.Rows
            Dim lvi As New ListViewItem(SafeString(row("id")))
            lvi.SubItems.Add(SafeString(row("title")))
            lvi.SubItems.Add(SafeString(row("caizhiid")))
            lvi.SubItems.Add(SafeString(row("chengse")))
            lvi.SubItems.Add(SafeString(row("xiangqian")))
            lvi.SubItems.Add(SafeString(row("shuliang")))
            lvi.SubItems.Add(SafeString(row("kejia")))
            lvi.SubItems.Add(SafeString(row("jifen")))
            超级列表框Ex_品类管理.Items.Add(lvi)
        Next
    End Sub

    Private Sub 列表项目单击(sender As Object, e As EventArgs)
        If 超级列表框Ex_品类管理.SelectedItems.Count = 0 Then Return

        Dim lvi As ListViewItem = 超级列表框Ex_品类管理.SelectedItems(0)
        集_操作id = lvi.Text
        集_操作模式 = "修改"

        ' 填充编辑区域
        编辑框_品类名称.Text = lvi.SubItems(1).Text

        ' 材质 (caizhiid 是 | 分隔的字符串)
        Dim caizhi As String = lvi.SubItems(2).Text
        组合框_材质.Text = If(caizhi.Contains("|"), caizhi.Replace("|", "|"), caizhi)

        ' 成色
        Dim chengse As String = lvi.SubItems(3).Text
        ' 重置成色复选框
        选择框_一号色.Checked = False
        选择框_二号色.Checked = False
        选择框_三号色.Checked = False
        选择框_四号色.Checked = False
        选择框_五号色.Checked = False

        If Not String.IsNullOrEmpty(chengse) Then
            Dim chengseArray As String() = chengse.Split("|"c)
            If chengseArray.Length >= 1 Then 选择框_一号色.Checked = True
            If chengseArray.Length >= 2 Then 选择框_二号色.Checked = True
            If chengseArray.Length >= 3 Then 选择框_三号色.Checked = True
            If chengseArray.Length >= 4 Then 选择框_四号色.Checked = True
            If chengseArray.Length >= 5 Then 选择框_五号色.Checked = True
        End If

        ' 镶嵌
        Dim xiangqian As String = lvi.SubItems(4).Text
        If xiangqian = "镶嵌" Then
            单选框_镶嵌.Checked = True
            单选框_非镶嵌.Checked = False
        Else
            单选框_镶嵌.Checked = False
            单选框_非镶嵌.Checked = True
        End If

        ' 多数量
        Dim shuliang As String = lvi.SubItems(5).Text
        If shuliang = "0" Then
            单选框_是.Checked = False
            单选框_否.Checked = True
        Else
            单选框_是.Checked = True
            单选框_否.Checked = False
        End If

        ' 克价必填
        Dim kejia As String = lvi.SubItems(6).Text
        If kejia = "0" Then
            单选框_必填.Checked = False
            单选框_非必填.Checked = True
        Else
            单选框_必填.Checked = True
            单选框_非必填.Checked = False
        End If

        ' 积分比例
        编辑框_积分比例.Text = lvi.SubItems(7).Text
    End Sub

#End Region

#Region "按钮事件"

    Private Sub 按钮EX4_Click(sender As Object, e As EventArgs) Handles 按钮EX4.Click
        ' 新增模式
        集_操作模式 = "新增"
        重置状态()
    End Sub

    Private Sub 按钮EX_查询_Click(sender As Object, e As EventArgs) Handles 按钮EX_查询.Click
        Dim keyword As String = 编辑框_品类名称.Text.Trim()
        If String.IsNullOrEmpty(keyword) Then
            列表加载()
            Return
        End If

        超级列表框Ex_品类管理.Items.Clear()
        Dim sql As String = $"SELECT * FROM xipunum_erp_category WHERE title like '%{SafeSQL(keyword)}%' order by id ASC"
        Dim dt As DataTable = DatabaseModule.MySQL_Read(sql)
        If dt Is Nothing Then Return

        For Each row As DataRow In dt.Rows
            Dim lvi As New ListViewItem(SafeString(row("id")))
            lvi.SubItems.Add(SafeString(row("title")))
            lvi.SubItems.Add(SafeString(row("caizhiid")))
            lvi.SubItems.Add(SafeString(row("chengse")))
            lvi.SubItems.Add(SafeString(row("xiangqian")))
            lvi.SubItems.Add(SafeString(row("shuliang")))
            lvi.SubItems.Add(SafeString(row("kejia")))
            lvi.SubItems.Add(SafeString(row("jifen")))
            超级列表框Ex_品类管理.Items.Add(lvi)
        Next
    End Sub

    Private Sub 按钮EX_重置_Click(sender As Object, e As EventArgs) Handles 按钮EX_重置.Click
        重置状态()
        列表加载()
    End Sub

    Private Sub 按钮EX_保存_Click(sender As Object, e As EventArgs) Handles 按钮EX_保存.Click
        Dim title As String = 编辑框_品类名称.Text.Trim()
        If String.IsNullOrEmpty(title) Then
            ShowWarning("品类名称不能为空！")
            编辑框_品类名称.Focus()
            Return
        End If

        Dim caizhi As String = 组合框_材质.Text.Trim()

        ' 成色 (用 | 连接选中的)
        Dim chengseList As New List(Of String)()
        If 选择框_一号色.Checked Then chengseList.Add("一号色")
        If 选择框_二号色.Checked Then chengseList.Add("二号色")
        If 选择框_三号色.Checked Then chengseList.Add("三号色")
        If 选择框_四号色.Checked Then chengseList.Add("四号色")
        If 选择框_五号色.Checked Then chengseList.Add("五号色")

        ' 如果没有勾选，使用成色标签文本
        Dim chengse As String
        If chengseList.Count > 0 Then
            chengse = String.Join("|", chengseList)
        Else
            ' 默认全选（根据 EasyLanguage 逻辑，如果都不选就全有）
            chengse = "一号色|二号色|三号色|四号色|五号色"
        End If

        Dim xiangqian As String = If(单选框_镶嵌.Checked, "镶嵌", "非镶嵌")
        Dim shuliang As String = If(单选框_是.Checked, "1", "0")
        Dim kejia As String = If(单选框_必填.Checked, "1", "0")
        Dim jifen As String = 编辑框_积分比例.Text.Trim()
        If String.IsNullOrEmpty(jifen) Then jifen = "0"

        Dim 信息操作日期 As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Dim 信息操作账户 As String = GlobalVariables.全局_用户账户

        If 集_操作模式 = "新增" Then
            ' 检查重复
            Dim sqlCheck As String = $"SELECT id FROM xipunum_erp_category WHERE title='{SafeSQL(title)}' LIMIT 1"
            Dim dtCheck As DataTable = DatabaseModule.MySQL_Read(sqlCheck)
            If dtCheck IsNot Nothing AndAlso dtCheck.Rows.Count > 0 Then
                ShowWarning("该品类名称已存在！")
                Return
            End If

            Dim sqlInsert As String = $"INSERT INTO xipunum_erp_category (title,caizhiid,chengse,xiangqian,shuliang,kejia,jifen,cjuser,creationtime,updatetime) VALUES ('{SafeSQL(title)}','{SafeSQL(caizhi)}','{SafeSQL(chengse)}','{SafeSQL(xiangqian)}','{SafeSQL(shuliang)}','{SafeSQL(kejia)}','{SafeSQL(jifen)}','{SafeSQL(信息操作账户)}','{SafeSQL(信息操作日期)}','{SafeSQL(信息操作日期)}')"
            DatabaseModule.MySQL_Write(sqlInsert)

            ShowInfo("新增品类成功！")
        ElseIf 集_操作模式 = "修改" Then
            If String.IsNullOrEmpty(集_操作id) Then
                ShowWarning("请先选择要修改的品类！")
                Return
            End If

            ' 检查重复（排除自己）
            Dim sqlCheck As String = $"SELECT id FROM xipunum_erp_category WHERE title='{SafeSQL(title)}' AND id<>'{SafeSQL(集_操作id)}' LIMIT 1"
            Dim dtCheck As DataTable = DatabaseModule.MySQL_Read(sqlCheck)
            If dtCheck IsNot Nothing AndAlso dtCheck.Rows.Count > 0 Then
                ShowWarning("该品类名称已存在！")
                Return
            End If

            Dim sqlUpdate As String = $"UPDATE xipunum_erp_category SET title='{SafeSQL(title)}',caizhiid='{SafeSQL(caizhi)}',chengse='{SafeSQL(chengse)}',xiangqian='{SafeSQL(xiangqian)}',shuliang='{SafeSQL(shuliang)}',kejia='{SafeSQL(kejia)}',jifen='{SafeSQL(jifen)}',cjuser='{SafeSQL(信息操作账户)}',updatetime='{SafeSQL(信息操作日期)}' WHERE id='{SafeSQL(集_操作id)}'"
            DatabaseModule.MySQL_Write(sqlUpdate)

            ShowInfo("修改品类成功！")
        End If

        ' 系统日志
        Dim 日志内容 As String = $"账户:{GlobalVariables.全局_用户账户} {集_操作模式}品类，品类名称:{title}"
        Dim sqlLog As String = $"INSERT INTO xipunum_erp_xitong_log SET type='{SafeSQL(集_操作模式)}',title='品类管理',conter='{SafeSQL(日志内容)}',user='{SafeSQL(信息操作账户)}',creationtime='{SafeSQL(信息操作日期)}'"
        DatabaseModule.MySQL_Write(sqlLog)

        重置状态()
        列表加载()
    End Sub

#End Region

#Region "单选框事件"

    Private Sub 单选框_镶嵌_CheckedChanged(sender As Object, e As EventArgs) Handles 单选框_镶嵌.CheckedChanged
        If Not 单选框_镶嵌.Checked Then
            单选框_非镶嵌.Checked = True
        Else
            单选框_非镶嵌.Checked = False
        End If
    End Sub

    Private Sub 单选框_非镶嵌_CheckedChanged(sender As Object, e As EventArgs) Handles 单选框_非镶嵌.CheckedChanged
        If Not 单选框_非镶嵌.Checked Then
            单选框_镶嵌.Checked = True
        Else
            单选框_镶嵌.Checked = False
        End If
    End Sub

    Private Sub 单选框_是_CheckedChanged(sender As Object, e As EventArgs) Handles 单选框_是.CheckedChanged
        If 单选框_是.Checked Then
            单选框_否.Checked = False
        End If
    End Sub

    Private Sub 单选框_否_CheckedChanged(sender As Object, e As EventArgs) Handles 单选框_否.CheckedChanged
        If 单选框_否.Checked Then
            单选框_是.Checked = False
        End If
    End Sub

    Private Sub 单选框_必填_CheckedChanged(sender As Object, e As EventArgs) Handles 单选框_必填.CheckedChanged
        If 单选框_必填.Checked Then
            单选框_非必填.Checked = False
        End If
    End Sub

    Private Sub 单选框_非必填_CheckedChanged(sender As Object, e As EventArgs) Handles 单选框_非必填.CheckedChanged
        If 单选框_非必填.Checked Then
            单选框_必填.Checked = False
        End If
    End Sub

#End Region

#Region "工具方法"

    Private Sub 重置状态()
        集_操作模式 = "新增"
        集_操作id = ""
        编辑框_品类名称.Text = ""
        组合框_材质.Text = ""
        选择框_一号色.Checked = False
        选择框_二号色.Checked = False
        选择框_三号色.Checked = False
        选择框_四号色.Checked = False
        选择框_五号色.Checked = False
        单选框_镶嵌.Checked = False
        单选框_非镶嵌.Checked = True
        单选框_是.Checked = False
        单选框_否.Checked = True
        单选框_必填.Checked = False
        单选框_非必填.Checked = True
        编辑框_积分比例.Text = ""
    End Sub

    Private Shared Function SafeSQL(input As String) As String
        If String.IsNullOrEmpty(input) Then Return ""
        Return input.Replace("'", "\'")
    End Function

    Private Shared Function SafeString(obj As Object) As String
        If obj Is Nothing OrElse IsDBNull(obj) Then Return ""
        Return obj.ToString()
    End Function

    Private Sub ShowWarning(msg As String)
        MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub ShowInfo(msg As String)
        MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

#End Region

End Class

' ============================================================================
' 商品规格管理窗口
' 功能: 管理商品规格的新增、修改
' 100% 匹配易语言 窗口_商品规格管理 功能
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports VB_ERP.Modules

Public Class SpecForm
    Inherits System.Windows.Forms.Form

#Region "程序集变量"

    Private 集_操作模式 As String = ""        ' "新增" 或 "修改"
    Private 集_操作id As String = ""          ' 修改时的 id

    ' 列表
    Private 超级列表框Ex_规格管理 As New ListView()

#End Region

#Region "控件声明"

    Private WithEvents 按钮EX_查询 As New Button()
    Private WithEvents 按钮EX_重置 As New Button()
    Private WithEvents 按钮EX_保存 As New Button()
    Private WithEvents 按钮EX4 As New Button()        ' 新增按钮

    Private 组合框_品类名称 As New ComboBox()
    Private 编辑框_规格简写 As New TextBox()
    Private 编辑框_规格排序 As New TextBox()

    Private WithEvents 单选框_是 As New RadioButton()
    Private WithEvents 单选框_否 As New RadioButton()

    Private 添加修改_分组框 As New GroupBox()

#End Region

#Region "初始化"

    Public Sub New()
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        Me.Text = "商品规格管理"
        Me.Size = New Drawing.Size(1246, 659)
        Me.StartPosition = FormStartPosition.CenterParent

        ' 主面板
        添加修改_分组框.Text = "规格管理"
        添加修改_分组框.Size = New Drawing.Size(1220, 620)
        添加修改_分组框.Location = New Drawing.Point(10, 10)
        Me.Controls.Add(添加修改_分组框)

        ' 列表
        超级列表框Ex_规格管理.View = View.Details
        超级列表框Ex_规格管理.FullRowSelect = True
        超级列表框Ex_规格管理.GridLines = True
        超级列表框Ex_规格管理.Columns.Add("ID", 50)
        超级列表框Ex_规格管理.Columns.Add("规格名称", 120)
        超级列表框Ex_规格管理.Columns.Add("规格简写", 80)
        超级列表框Ex_规格管理.Columns.Add("品类名称", 120)
        超级列表框Ex_规格管理.Columns.Add("多数量", 60)
        超级列表框Ex_规格管理.Columns.Add("排序", 60)
        超级列表框Ex_规格管理.Location = New Drawing.Point(8, 8)
        超级列表框Ex_规格管理.Size = New Drawing.Point(700, 600)
        添加修改_分组框.Controls.Add(超级列表框Ex_规格管理)

        ' 右侧编辑面板
        Dim editX As Integer = 720
        Dim editY As Integer = 18

        ' 按钮EX4 - 新增
        按钮EX4.Text = "新增"
        按钮EX4.Location = New Drawing.Point(editX, editY)
        按钮EX4.Size = New Drawing.Size(72, 30)
        添加修改_分组框.Controls.Add(按钮EX4)

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
        Dim y1 As Integer = editY + 40
        AddEditLabel("品类名称:", editX, y1)
        组合框_品类名称.Location = New Drawing.Point(editX + 77, y1)
        组合框_品类名称.Size = New Drawing.Size(155, 25)
        组合框_品类名称.DropDownStyle = ComboBoxStyle.DropDownList
        添加修改_分组框.Controls.Add(组合框_品类名称)

        ' 规格简写
        Dim y2 As Integer = y1 + 35
        AddEditLabel("规格简写:", editX, y2)
        编辑框_规格简写.Location = New Drawing.Point(editX + 77, y2)
        编辑框_规格简写.Size = New Drawing.Size(112, 25)
        添加修改_分组框.Controls.Add(编辑框_规格简写)

        ' 规格排序
        Dim y3 As Integer = y2 + 35
        AddEditLabel("规格排序:", editX, y3)
        编辑框_规格排序.Location = New Drawing.Point(editX + 77, y3)
        编辑框_规格排序.Size = New Drawing.Size(112, 25)
        添加修改_分组框.Controls.Add(编辑框_规格排序)

        ' 多数量
        Dim y4 As Integer = y3 + 35
        AddEditLabel("多数量:", editX, y4)
        单选框_是.Text = "是"
        单选框_是.Location = New Drawing.Point(editX + 60, y4)
        单选框_是.Size = New Drawing.Size(50, 20)
        添加修改_分组框.Controls.Add(单选框_是)
        单选框_否.Text = "否"
        单选框_否.Location = New Drawing.Point(editX + 110, y4)
        单选框_否.Size = New Drawing.Size(50, 20)
        添加修改_分组框.Controls.Add(单选框_否)

        ' 保存按钮
        按钮EX_保存.Text = "保存"
        按钮EX_保存.Location = New Drawing.Point(editX + 77, y4 + 40)
        按钮EX_保存.Size = New Drawing.Size(80, 30)
        添加修改_分组框.Controls.Add(按钮EX_保存)

        ' 事件
        AddHandler Me.Load, AddressOf 窗口_商品规格管理_创建完毕
        AddHandler Me.Resize, AddressOf 窗口_商品规格管理_尺寸被改变
        AddHandler 超级列表框Ex_规格管理.Click, AddressOf 列表项目单击
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

    Private Sub 窗口_商品规格管理_创建完毕(sender As Object, e As EventArgs)
        ' 加载品类组合框
        加载品类列表()

        ' 加载规格列表
        列表加载()

        重置状态()
    End Sub

    Private Sub 窗口_商品规格管理_尺寸被改变(sender As Object, e As EventArgs)
        If 添加修改_分组框 IsNot Nothing Then
            添加修改_分组框.Width = Me.ClientSize.Width - 20
            添加修改_分组框.Height = Me.ClientSize.Height - 20
            超级列表框Ex_规格管理.Height = 添加修改_分组框.Height - 20
        End If
    End Sub

#End Region

#Region "列表操作"

    Private Sub 加载品类列表()
        组合框_品类名称.Items.Clear()

        Dim sql As String = "SELECT * FROM xipunum_erp_category where 1=1 order by id ASC"
        Dim dt As DataTable = DatabaseModule.MySQL_Read(sql)
        If dt Is Nothing Then Return

        For Each row As DataRow In dt.Rows
            Dim id As String = SafeString(row("id"))
            Dim title As String = SafeString(row("title"))
            组合框_品类名称.Items.Add(New ComboBoxItem(title, id))
        Next

        If 组合框_品类名称.Items.Count > 0 Then
            组合框_品类名称.SelectedIndex = 0
        End If
    End Sub

    Private Sub 列表加载()
        超级列表框Ex_规格管理.Items.Clear()

        Dim sql As String = "SELECT a.id AS aid,a.title AS atitle,a.jianxie AS ajianxie,b.title AS btitle,a.shuliang AS ashuliang,a.sort_order AS asort_order FROM xipunum_erp_specs AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id WHERE 1=1 ORDER BY a.sort_order ASC, a.id ASC"
        Dim dt As DataTable = DatabaseModule.MySQL_Read(sql)
        If dt Is Nothing Then Return

        For Each row As DataRow In dt.Rows
            Dim lvi As New ListViewItem(SafeString(row("aid")))
            lvi.SubItems.Add(SafeString(row("atitle")))
            lvi.SubItems.Add(SafeString(row("ajianxie")))
            lvi.SubItems.Add(SafeString(row("btitle")))
            lvi.SubItems.Add(SafeString(row("ashuliang")))
            lvi.SubItems.Add(SafeString(row("asort_order")))
            超级列表框Ex_规格管理.Items.Add(lvi)
        Next
    End Sub

    Private Sub 列表项目单击(sender As Object, e As EventArgs)
        If 超级列表框Ex_规格管理.SelectedItems.Count = 0 Then Return

        Dim lvi As ListViewItem = 超级列表框Ex_规格管理.SelectedItems(0)
        集_操作id = lvi.Text
        集_操作模式 = "修改"

        ' 填充编辑区域
        ' 匹配品类名称
        Dim 品类名称 As String = lvi.SubItems(3).Text
        If Not String.IsNullOrEmpty(品类名称) Then
            For i As Integer = 0 To 组合框_品类名称.Items.Count - 1
                Dim cbi As ComboBoxItem = CType(组合框_品类名称.Items(i), ComboBoxItem)
                If cbi.Text = 品类名称 Then
                    组合框_品类名称.SelectedIndex = i
                    Exit For
                End If
            Next
        End If

        编辑框_规格简写.Text = lvi.SubItems(2).Text

        Dim sortOrder As String = lvi.SubItems(5).Text
        If Not String.IsNullOrEmpty(sortOrder) AndAlso sortOrder <> "0" Then
            编辑框_规格排序.Text = sortOrder
        Else
            编辑框_规格排序.Text = lvi.Text ' 使用 id 作为排序
        End If

        ' 多数量
        Dim shuliang As String = lvi.SubItems(4).Text
        If shuliang = "0" Then
            单选框_是.Checked = False
            单选框_否.Checked = True
        Else
            单选框_是.Checked = True
            单选框_否.Checked = False
        End If
    End Sub

#End Region

#Region "按钮事件"

    Private Sub 按钮EX4_Click(sender As Object, e As EventArgs) Handles 按钮EX4.Click
        集_操作模式 = "新增"
        重置状态()
    End Sub

    Private Sub 按钮EX_查询_Click(sender As Object, e As EventArgs) Handles 按钮EX_查询.Click
        Dim keyword As String = 编辑框_规格简写.Text.Trim()
        If String.IsNullOrEmpty(keyword) Then
            列表加载()
            Return
        End If

        超级列表框Ex_规格管理.Items.Clear()
        Dim sql As String = $"SELECT a.id AS aid,a.title AS atitle,a.jianxie AS ajianxie,b.title AS btitle,a.shuliang AS ashuliang,a.sort_order AS asort_order FROM xipunum_erp_specs AS a LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id WHERE a.title LIKE '%{SafeSQL(keyword)}%' OR a.jianxie LIKE '%{SafeSQL(keyword)}%' ORDER BY a.sort_order ASC, a.id ASC"
        Dim dt As DataTable = DatabaseModule.MySQL_Read(sql)
        If dt Is Nothing Then Return

        For Each row As DataRow In dt.Rows
            Dim lvi As New ListViewItem(SafeString(row("aid")))
            lvi.SubItems.Add(SafeString(row("atitle")))
            lvi.SubItems.Add(SafeString(row("ajianxie")))
            lvi.SubItems.Add(SafeString(row("btitle")))
            lvi.SubItems.Add(SafeString(row("ashuliang")))
            lvi.SubItems.Add(SafeString(row("asort_order")))
            超级列表框Ex_规格管理.Items.Add(lvi)
        Next
    End Sub

    Private Sub 按钮EX_重置_Click(sender As Object, e As EventArgs) Handles 按钮EX_重置.Click
        重置状态()
        列表加载()
    End Sub

    Private Sub 按钮EX_保存_Click(sender As Object, e As EventArgs) Handles 按钮EX_保存.Click
        If 组合框_品类名称.SelectedIndex < 0 Then
            ShowWarning("请选择品类名称！")
            Return
        End If

        Dim cbi As ComboBoxItem = CType(组合框_品类名称.SelectedItem, ComboBoxItem)
        Dim categoryId As String = cbi.Value
        Dim specTitle As String = 编辑框_规格简写.Text.Trim()
        Dim jianxie As String = 编辑框_规格简写.Text.Trim()
        Dim sortOrder As String = 编辑框_规格排序.Text.Trim()
        Dim shuliang As String = If(单选框_是.Checked, "1", "0")

        ' 验证
        If String.IsNullOrEmpty(specTitle) Then
            ShowWarning("规格简写/名称不能为空！")
            编辑框_规格简写.Focus()
            Return
        End If

        If String.IsNullOrEmpty(sortOrder) Then
            sortOrder = "0"
        End If

        Dim 信息操作日期 As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Dim 信息操作账户 As String = GlobalVariables.全局_用户账户

        If 集_操作模式 = "新增" Then
            ' 检查重复
            Dim sqlCheck As String = $"SELECT id FROM xipunum_erp_specs WHERE title='{SafeSQL(specTitle)}' AND category_id='{SafeSQL(categoryId)}' LIMIT 1"
            Dim dtCheck As DataTable = DatabaseModule.MySQL_Read(sqlCheck)
            If dtCheck IsNot Nothing AndAlso dtCheck.Rows.Count > 0 Then
                ShowWarning("该规格名称在此品类下已存在！")
                Return
            End If

            Dim sqlInsert As String = $"INSERT INTO xipunum_erp_specs (title,category_id,jianxie,shuliang,sort_order,cjuser,creationtime,updatetime) VALUES ('{SafeSQL(specTitle)}','{SafeSQL(categoryId)}','{SafeSQL(jianxie)}','{SafeSQL(shuliang)}','{SafeSQL(sortOrder)}','{SafeSQL(信息操作账户)}','{SafeSQL(信息操作日期)}','{SafeSQL(信息操作日期)}')"
            DatabaseModule.MySQL_Write(sqlInsert)

            ShowInfo("新增规格成功！")
        ElseIf 集_操作模式 = "修改" Then
            If String.IsNullOrEmpty(集_操作id) Then
                ShowWarning("请先选择要修改的规格！")
                Return
            End If

            ' 检查重复（排除自己）
            Dim sqlCheck As String = $"SELECT id FROM xipunum_erp_specs WHERE title='{SafeSQL(specTitle)}' AND category_id='{SafeSQL(categoryId)}' AND id<>'{SafeSQL(集_操作id)}' LIMIT 1"
            Dim dtCheck As DataTable = DatabaseModule.MySQL_Read(sqlCheck)
            If dtCheck IsNot Nothing AndAlso dtCheck.Rows.Count > 0 Then
                ShowWarning("该规格名称在此品类下已存在！")
                Return
            End If

            Dim sqlUpdate As String = $"UPDATE xipunum_erp_specs SET title='{SafeSQL(specTitle)}',category_id='{SafeSQL(categoryId)}',jianxie='{SafeSQL(jianxie)}',shuliang='{SafeSQL(shuliang)}',sort_order='{SafeSQL(sortOrder)}',cjuser='{SafeSQL(信息操作账户)}',updatetime='{SafeSQL(信息操作日期)}' WHERE id='{SafeSQL(集_操作id)}'"
            DatabaseModule.MySQL_Write(sqlUpdate)

            ShowInfo("修改规格成功！")
        End If

        ' 系统日志
        Dim 日志内容 As String = $"账户:{GlobalVariables.全局_用户账户} {集_操作模式}规格，规格名称:{specTitle}"
        Dim sqlLog As String = $"INSERT INTO xipunum_erp_xitong_log SET type='{SafeSQL(集_操作模式)}',title='规格管理',conter='{SafeSQL(日志内容)}',user='{SafeSQL(信息操作账户)}',creationtime='{SafeSQL(信息操作日期)}'"
        DatabaseModule.MySQL_Write(sqlLog)

        重置状态()
        列表加载()
    End Sub

#End Region

#Region "单选框事件"

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

#End Region

#Region "工具方法"

    Private Sub 重置状态()
        集_操作模式 = "新增"
        集_操作id = ""
        编辑框_规格简写.Text = ""
        编辑框_规格排序.Text = ""
        单选框_是.Checked = False
        单选框_否.Checked = True
        If 组合框_品类名称.Items.Count > 0 Then
            组合框_品类名称.SelectedIndex = 0
        End If
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

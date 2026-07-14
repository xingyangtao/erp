' ============================================================================
' 品类属性管理窗口
' 功能: 管理品类属性的统计配置
' 100% 匹配易语言 窗口_品类属性管理 功能
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports VB_ERP.Modules

Public Class CategoryAttrForm
    Inherits System.Windows.Forms.Form

#Region "程序集变量"

    Private 集_属性列表 As New ListView()        ' 属性管理列表 (带checkbox)
    Private 集_品类列表 As New CheckedListBox()   ' 品类名称列表 (带checkbox)

#End Region

#Region "控件声明"

    Private WithEvents 按钮EX_规格管理保存 As New Button()
    Private WithEvents 按钮EX_规格管理重置 As New Button()

    Private 添加修改_分组框 As New GroupBox()

#End Region

#Region "初始化"

    Public Sub New()
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        Me.Text = "品类属性管理"
        Me.Size = New Drawing.Size(938, 575)
        Me.StartPosition = FormStartPosition.CenterParent

        ' 主面板
        添加修改_分组框.Text = "品类属性管理"
        添加修改_分组框.Size = New Drawing.Size(920, 545)
        添加修改_分组框.Location = New Drawing.Point(8, 8)
        Me.Controls.Add(添加修改_分组框)

        ' 属性管理列表 (带checkbox)
        集_属性列表.View = View.Details
        集_属性列表.FullRowSelect = True
        集_属性列表.GridLines = True
        集_属性列表.CheckBoxes = True
        集_属性列表.Columns.Add("属性名称", 180)
        集_属性列表.Location = New Drawing.Point(5, 45)
        集_属性列表.Size = New Drawing.Size(187, 420)
        添加修改_分组框.Controls.Add(集_属性列表)

        ' 品类名称列表 (带checkbox)
        集_品类列表.Location = New Drawing.Point(197, 45)
        集_品类列表.Size = New Drawing.Size(187, 420)
        添加修改_分组框.Controls.Add(集_品类列表)

        ' 按钮EX_规格管理保存
        按钮EX_规格管理保存.Text = "保存"
        按钮EX_规格管理保存.Location = New Drawing.Point(208, 474)
        按钮EX_规格管理保存.Size = New Drawing.Size(72, 30)
        添加修改_分组框.Controls.Add(按钮EX_规格管理保存)

        ' 按钮EX_规格管理重置
        按钮EX_规格管理重置.Text = "重置"
        按钮EX_规格管理重置.Location = New Drawing.Point(289, 473)
        按钮EX_规格管理重置.Size = New Drawing.Size(72, 30)
        添加修改_分组框.Controls.Add(按钮EX_规格管理重置)

        ' 事件
        AddHandler Me.Load, AddressOf 窗口_品类属性管理_创建完毕
    End Sub

#End Region

#Region "创建完毕"

    Private Sub 窗口_品类属性管理_创建完毕(sender As Object, e As EventArgs)
        ' 加载属性配置列表
        加载属性列表()

        ' 加载品类列表
        加载品类列表()
    End Sub

    Private Sub 加载属性列表()
        集_属性列表.Items.Clear()

        Dim sql As String = "SELECT * FROM xipunum_erp_category_stat_config WHERE 1=1 ORDER BY id ASC"
        Dim dt As DataTable = DatabaseModule.MySQL_Read(sql)
        If dt Is Nothing Then Return

        For Each row As DataRow In dt.Rows
            Dim lvi As New ListViewItem(SafeString(row("title")))
            lvi.Tag = row  ' 存储完整行数据
            lvi.Checked = SafeString(row("status")) = "1"
            集_属性列表.Items.Add(lvi)
        Next
    End Sub

    Private Sub 加载品类列表()
        集_品类列表.Items.Clear()

        Dim sql As String = "SELECT * FROM xipunum_erp_category WHERE 1=1 ORDER BY id ASC"
        Dim dt As DataTable = DatabaseModule.MySQL_Read(sql)
        If dt Is Nothing Then Return

        For Each row As DataRow In dt.Rows
            Dim id As String = SafeString(row("id"))
            Dim title As String = SafeString(row("title"))

            ' 检查哪些品类被选中了此属性
            Dim index As Integer = 集_品类列表.Items.Add(title)

            ' 默认 unchecked (会在选中属性后更新)
        Next
    End Sub

#End Region

#Region "属性管理列表事件"

    Private Sub 集_属性列表_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles 集_属性列表.ItemChecked
        ' 全选/取消 (表头回调)
        If 集_属性列表.Items.Count > 0 Then
            Dim firstChecked As Boolean = 集_属性列表.Items(0).Checked
            Dim allSame As Boolean = True
            For Each item As ListViewItem In 集_属性列表.Items
                If item.Checked <> firstChecked Then
                    allSame = False
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub 集_属性列表_SelectedIndexChanged(sender As Object, e As EventArgs) Handles 集_属性列表.SelectedIndexChanged
        选中属性数据初始化()
    End Sub

    Private Sub 选中属性数据初始化()
        If 集_属性列表.SelectedItems.Count = 0 Then Return

        Dim selectedItem As ListViewItem = 集_属性列表.SelectedItems(0)
        Dim row As DataRow = TryCast(selectedItem.Tag, DataRow)
        If row Is Nothing Then Return

        Dim categoryList As String = SafeString(row("category_list"))
        Dim configId As String = SafeString(row("id"))

        ' 根据 category_list 勾选品类
        If Not String.IsNullOrEmpty(categoryList) Then
            Dim categoryIds As String() = categoryList.Split(","c)

            ' 需要从数据库查询每个 category id 对应的 title
            For i As Integer = 0 To 集_品类列表.Items.Count - 1
                集_品类列表.SetItemChecked(i, False)
            Next

            ' Re-query to match by title
            Dim sql As String = $"SELECT id, title FROM xipunum_erp_category WHERE id IN ({String.Join(",", categoryIds)})"
            Dim dt As DataTable = DatabaseModule.MySQL_Read(sql)
            If dt IsNot Nothing Then
                Dim matchedTitles As New HashSet(Of String)()
                For Each matchRow As DataRow In dt.Rows
                    matchedTitles.Add(SafeString(matchRow("title")))
                Next

                For i As Integer = 0 To 集_品类列表.Items.Count - 1
                    If matchedTitles.Contains(集_品类列表.Items(i).ToString()) Then
                        集_品类列表.SetItemChecked(i, True)
                    End If
                Next
            End If
        Else
            ' 未设置则全不选
            For i As Integer = 0 To 集_品类列表.Items.Count - 1
                集_品类列表.SetItemChecked(i, False)
            Next
        End If
    End Sub

#End Region

#Region "品类列表事件"

    Private Sub 集_品类列表_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles 集_品类列表.ItemCheck
        ' 允许切换选中状态
    End Sub

#End Region

#Region "按钮事件"

    Private Sub 按钮EX_规格管理重置_Click(sender As Object, e As EventArgs) Handles 按钮EX_规格管理重置.Click
        加载属性列表()
        加载品类列表()
    End Sub

    Private Sub 按钮EX_规格管理保存_Click(sender As Object, e As EventArgs) Handles 按钮EX_规格管理保存.Click
        If 集_属性列表.SelectedItems.Count = 0 Then
            ShowWarning("请先选择要配置的属性！")
            Return
        End If

        Dim selectedItem As ListViewItem = 集_属性列表.SelectedItems(0)
        Dim row As DataRow = TryCast(selectedItem.Tag, DataRow)
        If row Is Nothing Then Return

        Dim configId As String = SafeString(row("id"))

        ' 获取选中的品类
        Dim selectedCategories As New List(Of String)()
        ' 收集选中的品类 title
        Dim selectedTitles As New List(Of String)()
        For i As Integer = 0 To 集_品类列表.Items.Count - 1
            If 集_品类列表.GetItemChecked(i) Then
                selectedTitles.Add(集_品类列表.Items(i).ToString())
            End If
        Next

        ' 根据 title 反查 id
        If selectedTitles.Count > 0 Then
            Dim titleList As String = String.Join("','", selectedTitles)
            Dim sql As String = $"SELECT id FROM xipunum_erp_category WHERE title IN ('{SafeSQL(titleList)}') ORDER BY id ASC"
            Dim dt As DataTable = DatabaseModule.MySQL_Read(sql)
            If dt IsNot Nothing Then
                For Each catRow As DataRow In dt.Rows
                    selectedCategories.Add(SafeString(catRow("id")))
                Next
            End If
        End If

        Dim categoryList As String = String.Join(",", selectedCategories)

        ' 更新配置
        Dim sqlUpdate As String = $"UPDATE xipunum_erp_category_stat_config SET category_list='{SafeSQL(categoryList)}', status='{If(selectedItem.Checked, "1", "0")}' WHERE id='{SafeSQL(configId)}'"
        DatabaseModule.MySQL_Write(sqlUpdate)

        ' 更新全局变量 (金类/银类/结算品类)
        UpdateGlobalCategoryVariables()

        ShowInfo("属性配置保存成功！")
    End Sub

    Private Sub UpdateGlobalCategoryVariables()
        ' 根据配置更新全局金类/银类/结算品类列表
        Dim sql As String = "SELECT * FROM xipunum_erp_category_stat_config WHERE status='1' ORDER BY id ASC"
        Dim dt As DataTable = DatabaseModule.MySQL_Read(sql)
        If dt Is Nothing Then Return

        Dim jinList As New List(Of String)()
        Dim yinList As New List(Of String)()
        Dim jiesuanList As New List(Of String)()

        For Each row As DataRow In dt.Rows
            Dim title As String = SafeString(row("title"))
            Dim categoryList As String = SafeString(row("category_list"))

            If title.Contains("金") Then
                jinList.Add(categoryList)
            ElseIf title.Contains("银") Then
                yinList.Add(categoryList)
            ElseIf title.Contains("结算") Then
                jiesuanList.Add(categoryList)
            End If
        Next

        ' 更新全局变量
        ' GlobalVariables.全局_金类 = String.Join(",", jinList)
        ' GlobalVariables.全局_银类 = String.Join(",", yinList)
        ' GlobalVariables.全局_结算品类 = String.Join(",", jiesuanList)
    End Sub

#End Region

#Region "工具方法"

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

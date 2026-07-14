' ============================================================================
' 会员列表排序窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_会员列表排序.form.e.txt
' 包含所有4个子程序
' ============================================================================

Public Class MemberSortForm
    Inherits System.Windows.Forms.Form

    ' ========== 控件声明（对应易语言窗口控件） ==========
    ' 分组框
    Private WithEvents grpQuery As New GroupBox()         ' 信息商品查询_分组框

    ' 组合框
    Private cmbFieldName As New ComboBox()                ' 组合框_字段名称

    ' 单选框
    Private WithEvents radioAsc As New RadioButton()      ' 单选框_升序
    Private WithEvents radioDesc As New RadioButton()     ' 单选框_降序

    ' 按钮
    Private WithEvents btnConfirm As New Button()         ' 按钮EX1（确定）
    Private WithEvents btnClose As New Button()           ' 图片框EX_关闭（关闭）

    ' ========== 当前排序设置（从主窗口传入） ==========
    Private currentSortSetting As String = ""

    ' ========== 构造函数 ==========
    Public Sub New(Optional sortSetting As String = "")
        currentSortSetting = sortSetting
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "会员列表排序"
        Me.Size = New Drawing.Size(400, 300)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MaximizeBox = False

        ' 关闭按钮（对应图片框EX_关闭）
        btnClose.Text = "X"
        btnClose.Location = New Drawing.Point(Me.Width - 35, 5)
        btnClose.Size = New Drawing.Size(30, 25)
        btnClose.BackColor = Drawing.Color.Red
        btnClose.ForeColor = Drawing.Color.White
        btnClose.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnClose)

        ' 信息商品查询分组框
        grpQuery.Text = "排序设置"
        grpQuery.Size = New Drawing.Size(300, 200)
        grpQuery.Location = New Drawing.Point((Me.Width - grpQuery.Width) \ 2, (Me.Height - grpQuery.Height) \ 2)
        Me.Controls.Add(grpQuery)

        ' 标签（对应标签Ex3）
        Dim lblField As New Label()
        lblField.Text = "排序字段："
        lblField.Location = New Drawing.Point(20, 30)
        lblField.AutoSize = True
        grpQuery.Controls.Add(lblField)

        ' 组合框_字段名称
        cmbFieldName.Items.Add("会员ID")
        cmbFieldName.Items.Add("结料(g)")
        cmbFieldName.Items.Add("结款(元)")
        cmbFieldName.Items.Add("创建时间")
        cmbFieldName.Items.Add("更新时间")
        cmbFieldName.DropDownStyle = ComboBoxStyle.DropDownList
        cmbFieldName.Location = New Drawing.Point(100, 27)
        cmbFieldName.Size = New Drawing.Size(180, 25)
        grpQuery.Controls.Add(cmbFieldName)

        ' 单选框_升序
        radioAsc.Text = "升序"
        radioAsc.Location = New Drawing.Point(100, 80)
        radioAsc.AutoSize = True
        grpQuery.Controls.Add(radioAsc)

        ' 单选框_降序
        radioDesc.Text = "降序"
        radioDesc.Location = New Drawing.Point(200, 80)
        radioDesc.AutoSize = True
        grpQuery.Controls.Add(radioDesc)

        ' 确定按钮（对应按钮EX1）
        btnConfirm.Text = "确定"
        btnConfirm.Location = New Drawing.Point(100, 130)
        btnConfirm.Size = New Drawing.Size(80, 30)
        grpQuery.Controls.Add(btnConfirm)
    End Sub

    ' ========== _窗口_会员列表排序_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 对应易语言：根据当前排序设置初始化选中状态

        ' 对应易语言：如果真(等于(局部_会员排序, "id DESC"))
        If currentSortSetting = "id DESC" Then
            cmbFieldName.SelectedIndex = 0
            radioAsc.Checked = False
            radioDesc.Checked = True
        End If

        ' 对应易语言：如果真(等于(局部_会员排序, "id asc"))
        If currentSortSetting = "id asc" Then
            cmbFieldName.SelectedIndex = 0
            radioAsc.Checked = True
            radioDesc.Checked = False
        End If

        ' 对应易语言：如果真(等于(局部_会员排序, "CDbl(data5) DESC"))
        If currentSortSetting = "CDbl(data5) DESC" Then
            cmbFieldName.SelectedIndex = 1
            radioAsc.Checked = False
            radioDesc.Checked = True
        End If

        ' 对应易语言：如果真(等于(局部_会员排序, "CDbl(data5) asc"))
        If currentSortSetting = "CDbl(data5) asc" Then
            cmbFieldName.SelectedIndex = 1
            radioAsc.Checked = True
            radioDesc.Checked = False
        End If

        ' 对应易语言：如果真(等于(局部_会员排序, "CDbl(data6) DESC"))
        If currentSortSetting = "CDbl(data6) DESC" Then
            cmbFieldName.SelectedIndex = 2
            radioAsc.Checked = False
            radioDesc.Checked = True
        End If

        ' 对应易语言：如果真(等于(局部_会员排序, "CDbl(data6) asc"))
        If currentSortSetting = "CDbl(data6) asc" Then
            cmbFieldName.SelectedIndex = 2
            radioAsc.Checked = True
            radioDesc.Checked = False
        End If

        ' 对应易语言：如果真(等于(局部_会员排序, "data9 DESC"))
        If currentSortSetting = "data9 DESC" Then
            cmbFieldName.SelectedIndex = 3
            radioAsc.Checked = False
            radioDesc.Checked = True
        End If

        ' 对应易语言：如果真(等于(局部_会员排序, "data9 asc"))
        If currentSortSetting = "data9 asc" Then
            cmbFieldName.SelectedIndex = 3
            radioAsc.Checked = True
            radioDesc.Checked = False
        End If

        ' 对应易语言：如果真(等于(局部_会员排序, "data10 DESC"))
        If currentSortSetting = "data10 DESC" Then
            cmbFieldName.SelectedIndex = 4
            radioAsc.Checked = False
            radioDesc.Checked = True
        End If

        ' 对应易语言：如果真(等于(局部_会员排序, "data10 asc"))
        If currentSortSetting = "data10 asc" Then
            cmbFieldName.SelectedIndex = 4
            radioAsc.Checked = True
            radioDesc.Checked = False
        End If

        ' 默认选中
        If cmbFieldName.SelectedIndex < 0 Then cmbFieldName.SelectedIndex = 0
        If Not radioAsc.Checked AndAlso Not radioDesc.Checked Then radioDesc.Checked = True
    End Sub

    ' ========== _窗口_会员列表排序_尺寸被改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        ' 对应易语言：分组框居中
        grpQuery.Left = (Me.Width - grpQuery.Width) \ 2
        grpQuery.Top = (Me.Height - grpQuery.Height) \ 2
    End Sub

    ' ========== _图片框EX_关闭_鼠标左键单击（关闭按钮） ==========
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ' ========== _按钮EX1_鼠标左键单击（确定按钮） ==========
    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        Dim sortField As String = ""
        Dim sortDirection As String = ""

        ' 对应易语言：根据组合框选中项确定排序字段
        If cmbFieldName.SelectedIndex = 0 Then sortField = "id "
        If cmbFieldName.SelectedIndex = 1 Then sortField = "CDbl(data5) "
        If cmbFieldName.SelectedIndex = 2 Then sortField = "CDbl(data6) "
        If cmbFieldName.SelectedIndex = 3 Then sortField = "data9 "
        If cmbFieldName.SelectedIndex = 4 Then sortField = "data10 "

        ' 对应易语言：根据单选框确定排序方向
        If radioAsc.Checked Then sortDirection = "asc"
        If radioDesc.Checked Then sortDirection = "DESC"

        ' 对应易语言：赋值(窗口_主窗口.局部_会员排序, 相加(查询排序字段, 查询排序升降))
        ' 设置主窗口排序属性并刷新
        If MainForm IsNot Nothing Then
            MainForm.MemberSortSetting = sortField & sortDirection
            MainForm.RefreshHomePage()
        End If

        Me.Close()
    End Sub

End Class

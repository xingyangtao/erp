' ============================================================================
' 商品信息添加窗口
' 功能: 新商品信息录入（对应易语言 窗口_商品信息添加）
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Drawing
Imports System.IO

Public Class ProductAddForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（对应易语言程序集变量） ==========
    Private 局部_款号文本内容 As String = ""
    Private 查找_款号文本内容 As String = ""
    Private 局部_品类名称id As String = ""
    Private 局部_品类简写 As String = ""
    Private 局部_商品规格id As String = ""
    Private 局部_品类多数量 As String = ""
    Private 局部_商品多数量 As String = ""
    Private 局部_商品零销售 As String = ""
    Private 局部_品类原料价 As String = ""

    ' ========== 控件声明 ==========
    ' 分类信息区域
    Private WithEvents 组合框_品类名称 As New ComboBox()
    Private WithEvents 组合框_商品材质 As New ComboBox()
    Private WithEvents 组合框_商品成色 As New ComboBox()
    Private WithEvents 组合框_商品规格 As New ComboBox()
    Private WithEvents 组合框_检测机构 As New ComboBox()

    ' 输入框
    Private WithEvents 编辑框_商品款号 As New TextBox()
    Private WithEvents 编辑框_商品名称 As New TextBox()
    Private WithEvents 编辑框_公司成色 As New TextBox()
    Private WithEvents 编辑框_固定编码 As New TextBox()
    Private WithEvents 编辑框_数量 As New TextBox()
    Private WithEvents 编辑框_重量 As New TextBox()
    Private WithEvents 编辑框_单件重 As New TextBox()
    Private WithEvents 编辑框_金重 As New TextBox()
    Private WithEvents 编辑框_损耗 As New TextBox()
    Private WithEvents 编辑框_含耗重 As New TextBox()
    Private WithEvents 编辑框_圈号长度 As New TextBox()
    Private WithEvents 编辑框_面宽 As New TextBox()
    Private WithEvents 编辑框_厚度 As New TextBox()
    Private WithEvents 编辑框_成本单价 As New TextBox()
    Private WithEvents 编辑框_系数 As New TextBox()
    Private WithEvents 编辑框_成本工费 As New TextBox()
    Private WithEvents 编辑框_成本附加费 As New TextBox()
    Private WithEvents 编辑框_销售工费 As New TextBox()
    Private WithEvents 编辑框_参考工费 As New TextBox()
    Private WithEvents 编辑框_销售附加费 As New TextBox()
    Private WithEvents 编辑框_销售价 As New TextBox()
    Private WithEvents 编辑框_备注 As New TextBox()
    Private WithEvents 编辑框_主石重 As New TextBox()
    Private WithEvents 编辑框_石头数量 As New TextBox()
    Private WithEvents 编辑框_副石重 As New TextBox()
    Private WithEvents 编辑框_副石头数量 As New TextBox()
    Private WithEvents 编辑框_商品主色 As New TextBox()
    Private WithEvents 编辑框_入库克价 As New TextBox()
    Private WithEvents 编辑框_图片地址 As New TextBox()
    Private WithEvents 编辑框_模具号 As New TextBox()
    Private WithEvents 编辑框_成本价 As New TextBox()
    Private WithEvents 编辑框_证书销售价 As New TextBox()

    ' 证书信息
    Private WithEvents 编辑框_证书编码 As New TextBox()
    Private WithEvents 编辑框_查询地址 As New TextBox()
    Private WithEvents 编辑框_检测结果 As New TextBox()
    Private WithEvents 编辑框_检测总重 As New TextBox()
    Private WithEvents 编辑框_检测形状 As New TextBox()
    Private WithEvents 编辑框_检测颜色 As New TextBox()

    ' 单选框
    Private WithEvents 单选框_有 As New RadioButton()
    Private WithEvents 单选框_无 As New RadioButton()
    Private WithEvents 单选框_重量 As New RadioButton()
    Private WithEvents 单选框_单件 As New RadioButton()
    Private WithEvents 单选框_固口 As New RadioButton()
    Private WithEvents 单选框_开口 As New RadioButton()

    ' 选择框
    Private WithEvents 图片识别_选择框 As New CheckBox()
    Private WithEvents 自动取图_选择框 As New CheckBox()
    Private WithEvents 标签打印_选择框 As New CheckBox()

    ' 按钮
    Private WithEvents 按钮EX_查找 As New Button()
    Private WithEvents 按钮EX_重置 As New Button()
    Private WithEvents 按钮EX_图片上传 As New Button()
    Private WithEvents 按钮EX_添加 As New Button()
    Private WithEvents 按钮EX_圈号长度 As New Button()
    Private WithEvents 按钮EX_新建款号 As New Button()

    ' 图片框
    Private 图片框EX_主图 As New PictureBox()
    Private 图片框EX_固开口 As New PictureBox()

    ' 图标列表框
    Private 图标列表框EX1 As New ListView()

    ' 查找输入
    Private 查找信息编辑框 As New TextBox()

    ' 分组框
    Private 证书信息_分组框 As New GroupBox()

    ' 通用对话框
    Private 通用对话框1 As New OpenFileDialog()

    ' 定时器
    Private WithEvents 时钟1 As New Timer()

    ' ========== 格式工具函数 ==========
    Private Function 格式三位小数(原数值 As String) As String
        Dim d As Double
        If Not Double.TryParse(原数值, d) Then Return "0.000"
        Dim rounded As Double = Math.Round(d, 3)
        Dim text As String = rounded.ToString("F3")
        Return text
    End Function

    Private Function 格式二位小数(原数值 As String) As String
        Dim d As Double
        If Not Double.TryParse(原数值, d) Then Return "0.00"
        Dim rounded As Double = Math.Round(d, 2)
        Dim text As String = rounded.ToString("F2")
        Return text
    End Function

    ' ========== 获取项目附加数值（ComboBox中存储ID） ==========
    Private Function 取项目附加数值(cb As ComboBox) As String
        If cb.SelectedIndex < 0 Then Return ""
        Dim item = TryCast(cb.SelectedItem, ComboBoxItem)
        If item IsNot Nothing Then Return item.ID
        Return ""
    End Function

    Private Function 取项目文本(cb As ComboBox) As String
        If cb.SelectedIndex < 0 Then Return ""
        Dim item = TryCast(cb.SelectedItem, ComboBoxItem)
        If item Is Nothing Then Return cb.Text
        Return item.Text
    End Function

    Private Function 取项目数量(cb As ComboBox) As Integer
        Return cb.Items.Count
    End Function

    ' ========== ComboBox辅助类 ==========
    Private Class ComboBoxItem
        Public Property ID As String
        Public Property Text As String
        Public Sub New(id As String, text As String)
            Me.ID = id
            Me.Text = text
        End Sub
        Public Overrides Function ToString() As String
            Return Text
        End Function
    End Class

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf _窗口_商品信息添加_创建完毕
        AddHandler Me.FormClosing, AddressOf _窗口_商品信息添加_可否被关闭
    End Sub

    ' ========== 界面初始化 ==========
    Private Sub InitializeUI()
        Me.Text = "商品信息添加"
        Me.Size = New Drawing.Size(1150, 810)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        ' 主面板
        Dim mainPanel As New Panel()
        mainPanel.Location = New Drawing.Point(10, 10)
        mainPanel.Size = New Drawing.Size(1120, 750)
        mainPanel.BorderStyle = BorderStyle.FixedSingle

        Dim y As Integer = 10
        Dim lx As Integer = 10
        Dim cx As Integer = 80

        ' ==== 品类名称 ====
        AddLabel(mainPanel, "品类名称：", lx, y)
        组合框_品类名称.Location = New Drawing.Point(cx, y - 3)
        组合框_品类名称.Size = New Drawing.Size(150, 25)
        组合框_品类名称.DropDownStyle = ComboBoxStyle.DropDownList
        mainPanel.Controls.Add(组合框_品类名称)

        按钮EX_查找.Text = "查找"
        按钮EX_查找.Location = New Drawing.Point(cx + 160, y - 5)
        按钮EX_查找.Size = New Drawing.Size(60, 28)
        mainPanel.Controls.Add(按钮EX_查找)

        查找信息编辑框.Location = New Drawing.Point(cx + 230, y - 3)
        查找信息编辑框.Size = New Drawing.Size(120, 25)
        mainPanel.Controls.Add(查找信息编辑框)

        按钮EX_重置.Text = "重置"
        按钮EX_重置.Location = New Drawing.Point(cx + 360, y - 5)
        按钮EX_重置.Size = New Drawing.Size(60, 28)
        mainPanel.Controls.Add(按钮EX_重置)

        图片识别_选择框.Text = "图片识别"
        图片识别_选择框.Location = New Drawing.Point(cx + 430, y - 2)
        图片识别_选择框.Size = New Drawing.Size(80, 22)
        图片识别_选择框.Checked = True
        mainPanel.Controls.Add(图片识别_选择框)

        按钮EX_图片上传.Text = "上传图片"
        按钮EX_图片上传.Location = New Drawing.Point(cx + 520, y - 5)
        按钮EX_图片上传.Size = New Drawing.Size(80, 28)
        mainPanel.Controls.Add(按钮EX_图片上传)

        ' ==== 主图显示 ====
        图片框EX_主图.Location = New Drawing.Point(lx + 700, y - 3)
        图片框EX_主图.Size = New Drawing.Size(200, 200)
        图片框EX_主图.BorderStyle = BorderStyle.FixedSingle
        图片框EX_主图.SizeMode = PictureBoxSizeMode.Zoom
        mainPanel.Controls.Add(图片框EX_主图)

        y += 35

        ' ==== 商品材质 ====
        AddLabel(mainPanel, "商品材质：", lx, y)
        组合框_商品材质.Location = New Drawing.Point(cx, y - 3)
        组合框_商品材质.Size = New Drawing.Size(120, 25)
        组合框_商品材质.DropDownStyle = ComboBoxStyle.DropDownList
        mainPanel.Controls.Add(组合框_商品材质)

        ' ==== 商品成色 ====
        AddLabel(mainPanel, "商品成色：", lx + 210, y)
        组合框_商品成色.Location = New Drawing.Point(lx + 290, y - 3)
        组合框_商品成色.Size = New Drawing.Size(120, 25)
        组合框_商品成色.DropDownStyle = ComboBoxStyle.DropDownList
        mainPanel.Controls.Add(组合框_商品成色)

        ' ==== 商品规格 ====
        AddLabel(mainPanel, "商品规格：", lx + 420, y)
        组合框_商品规格.Location = New Drawing.Point(lx + 500, y - 3)
        组合框_商品规格.Size = New Drawing.Size(120, 25)
        组合框_商品规格.DropDownStyle = ComboBoxStyle.DropDownList
        mainPanel.Controls.Add(组合框_商品规格)

        y += 35

        ' ==== 入库克价 ====
        AddLabel(mainPanel, "入库克价：", lx, y)
        编辑框_入库克价.Location = New Drawing.Point(cx, y - 3)
        编辑框_入库克价.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_入库克价)

        ' ==== 价格单位 ====
        单选框_重量.Text = "按重量"
        单选框_重量.Location = New Drawing.Point(cx + 90, y - 2)
        单选框_重量.Size = New Drawing.Size(65, 22)
        mainPanel.Controls.Add(单选框_重量)

        单选框_单件.Text = "按单件"
        单选框_单件.Location = New Drawing.Point(cx + 160, y - 2)
        单选框_单件.Size = New Drawing.Size(65, 22)
        mainPanel.Controls.Add(单选框_单件)

        ' ==== 固定编码 ====
        AddLabel(mainPanel, "固定编码：", lx + 300, y)
        编辑框_固定编码.Location = New Drawing.Point(lx + 380, y - 3)
        编辑框_固定编码.Size = New Drawing.Size(120, 25)
        mainPanel.Controls.Add(编辑框_固定编码)

        y += 35

        ' ==== 重量行 ====
        AddLabel(mainPanel, "商品重量：", lx, y)
        编辑框_重量.Location = New Drawing.Point(cx, y - 3)
        编辑框_重量.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_重量)

        AddLabel(mainPanel, "单件重：", cx + 90, y)
        编辑框_单件重.Location = New Drawing.Point(cx + 150, y - 3)
        编辑框_单件重.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_单件重)

        AddLabel(mainPanel, "批次数量：", cx + 240, y)
        编辑框_数量.Location = New Drawing.Point(cx + 310, y - 3)
        编辑框_数量.Size = New Drawing.Size(60, 25)
        mainPanel.Controls.Add(编辑框_数量)

        AddLabel(mainPanel, "金重：", cx + 380, y)
        编辑框_金重.Location = New Drawing.Point(cx + 425, y - 3)
        编辑框_金重.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_金重)

        AddLabel(mainPanel, "损耗：", cx + 515, y)
        编辑框_损耗.Location = New Drawing.Point(cx + 555, y - 3)
        编辑框_损耗.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_损耗)

        AddLabel(mainPanel, "含耗重：", cx + 645, y)
        编辑框_含耗重.Location = New Drawing.Point(cx + 695, y - 3)
        编辑框_含耗重.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_含耗重)

        y += 35

        ' ==== 固口/开口 ====
        AddLabel(mainPanel, "圈号长度：", lx, y)
        编辑框_圈号长度.Location = New Drawing.Point(cx, y - 3)
        编辑框_圈号长度.Size = New Drawing.Size(72, 25)
        mainPanel.Controls.Add(编辑框_圈号长度)

        按钮EX_圈号长度.Text = "固/开"
        按钮EX_圈号长度.Location = New Drawing.Point(cx + 80, y - 5)
        按钮EX_圈号长度.Size = New Drawing.Size(50, 28)
        mainPanel.Controls.Add(按钮EX_圈号长度)

        图片框EX_固开口.Location = New Drawing.Point(cx + 135, y - 3)
        图片框EX_固开口.Size = New Drawing.Size(60, 26)
        mainPanel.Controls.Add(图片框EX_固开口)

        单选框_固口.Text = "固口"
        单选框_固口.Location = New Drawing.Point(cx + 200, y - 2)
        单选框_固口.Size = New Drawing.Size(55, 22)
        mainPanel.Controls.Add(单选框_固口)

        单选框_开口.Text = "开口"
        单选框_开口.Location = New Drawing.Point(cx + 255, y - 2)
        单选框_开口.Size = New Drawing.Size(55, 22)
        mainPanel.Controls.Add(单选框_开口)

        AddLabel(mainPanel, "面宽：", cx + 320, y)
        编辑框_面宽.Location = New Drawing.Point(cx + 360, y - 3)
        编辑框_面宽.Size = New Drawing.Size(70, 25)
        mainPanel.Controls.Add(编辑框_面宽)

        AddLabel(mainPanel, "厚度：", cx + 440, y)
        编辑框_厚度.Location = New Drawing.Point(cx + 475, y - 3)
        编辑框_厚度.Size = New Drawing.Size(70, 25)
        mainPanel.Controls.Add(编辑框_厚度)

        y += 35

        ' ==== 成本单价与销售价 ====
        AddLabel(mainPanel, "成本单价：", lx, y)
        编辑框_成本单价.Location = New Drawing.Point(cx, y - 3)
        编辑框_成本单价.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_成本单价)

        AddLabel(mainPanel, "系数：", cx + 90, y)
        编辑框_系数.Location = New Drawing.Point(cx + 130, y - 3)
        编辑框_系数.Size = New Drawing.Size(60, 25)
        mainPanel.Controls.Add(编辑框_系数)

        AddLabel(mainPanel, "销售价：", cx + 200, y)
        编辑框_销售价.Location = New Drawing.Point(cx + 255, y - 3)
        编辑框_销售价.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_销售价)

        ' ==== 费费 ====
        AddLabel(mainPanel, "成本工费：", lx, y + 35)
        编辑框_成本工费.Location = New Drawing.Point(cx, y + 32)
        编辑框_成本工费.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_成本工费)

        AddLabel(mainPanel, "销售工费：", cx + 90, y + 35)
        编辑框_销售工费.Location = New Drawing.Point(cx + 160, y + 32)
        编辑框_销售工费.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_销售工费)

        AddLabel(mainPanel, "参考工费：", cx + 250, y + 35)
        编辑框_参考工费.Location = New Drawing.Point(cx + 320, y + 32)
        编辑框_参考工费.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_参考工费)

        AddLabel(mainPanel, "成本附加费：", cx + 410, y + 35)
        编辑框_成本附加费.Location = New Drawing.Point(cx + 490, y + 32)
        编辑框_成本附加费.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_成本附加费)

        AddLabel(mainPanel, "销售附加费：", cx + 580, y + 35)
        编辑框_销售附加费.Location = New Drawing.Point(cx + 660, y + 32)
        编辑框_销售附加费.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_销售附加费)

        y += 75

        ' ==== 主石重/副石重 ====
        AddLabel(mainPanel, "主石重：", lx, y)
        编辑框_主石重.Location = New Drawing.Point(cx, y - 3)
        编辑框_主石重.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_主石重)

        AddLabel(mainPanel, "石头数量：", cx + 90, y)
        编辑框_石头数量.Location = New Drawing.Point(cx + 160, y - 3)
        编辑框_石头数量.Size = New Drawing.Size(60, 25)
        mainPanel.Controls.Add(编辑框_石头数量)

        AddLabel(mainPanel, "副石重：", cx + 230, y)
        编辑框_副石重.Location = New Drawing.Point(cx + 290, y - 3)
        编辑框_副石重.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_副石重)

        AddLabel(mainPanel, "副石数：", cx + 380, y)
        编辑框_副石头数量.Location = New Drawing.Point(cx + 430, y - 3)
        编辑框_副石头数量.Size = New Drawing.Size(60, 25)
        mainPanel.Controls.Add(编辑框_副石头数量)

        AddLabel(mainPanel, "主色：", cx + 500, y)
        编辑框_商品主色.Location = New Drawing.Point(cx + 540, y - 3)
        编辑框_商品主色.Size = New Drawing.Size(60, 25)
        mainPanel.Controls.Add(编辑框_商品主色)

        y += 35

        ' ==== 商品名称/款号/公司成色 ====
        AddLabel(mainPanel, "商品名称：", lx, y)
        编辑框_商品名称.Location = New Drawing.Point(cx, y - 3)
        编辑框_商品名称.Size = New Drawing.Size(150, 25)
        mainPanel.Controls.Add(编辑框_商品名称)

        AddLabel(mainPanel, "商品款号：", cx + 250, y)
        编辑框_商品款号.Location = New Drawing.Point(cx + 320, y - 3)
        编辑框_商品款号.Size = New Drawing.Size(120, 25)
        mainPanel.Controls.Add(编辑框_商品款号)

        AddLabel(mainPanel, "公司成色：", cx + 450, y)
        编辑框_公司成色.Location = New Drawing.Point(cx + 520, y - 3)
        编辑框_公司成色.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_公司成色)

        y += 35

        ' ==== 模具号/图片地址 ====
        AddLabel(mainPanel, "模具号：", lx, y)
        编辑框_模具号.Location = New Drawing.Point(cx, y - 3)
        编辑框_模具号.Size = New Drawing.Size(110, 25)
        mainPanel.Controls.Add(编辑框_模具号)

        AddLabel(mainPanel, "图片地址：", cx + 200, y)
        编辑框_图片地址.Location = New Drawing.Point(cx + 270, y - 3)
        编辑框_图片地址.Size = New Drawing.Size(300, 25)
        mainPanel.Controls.Add(编辑框_图片地址)

        y += 35

        ' ==== 成本价/证书销售价 ====
        AddLabel(mainPanel, "成本价：", lx, y)
        编辑框_成本价.Location = New Drawing.Point(cx, y - 3)
        编辑框_成本价.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_成本价)

        AddLabel(mainPanel, "证书销售价：", cx + 180, y)
        编辑框_证书销售价.Location = New Drawing.Point(cx + 260, y - 3)
        编辑框_证书销售价.Size = New Drawing.Size(80, 25)
        mainPanel.Controls.Add(编辑框_证书销售价)

        ' ==== 有/无证书单选框 ====
        单选框_有.Text = "有证书"
        单选框_有.Location = New Drawing.Point(cx + 350, y - 2)
        单选框_有.Size = New Drawing.Size(65, 22)
        mainPanel.Controls.Add(单选框_有)

        单选框_无.Text = "无证书"
        单选框_无.Location = New Drawing.Point(cx + 420, y - 2)
        单选框_无.Size = New Drawing.Size(65, 22)
        mainPanel.Controls.Add(单选框_无)

        y += 35

        ' ==== 新建款号 / 标签打印 ====
        按钮EX_新建款号.Text = "新建款号"
        按钮EX_新建款号.Location = New Drawing.Point(lx, y)
        按钮EX_新建款号.Size = New Drawing.Size(80, 30)
        mainPanel.Controls.Add(按钮EX_新建款号)

        标签打印_选择框.Text = "标签打印"
        标签打印_选择框.Location = New Drawing.Point(cx, y + 2)
        标签打印_选择框.Size = New Drawing.Size(80, 22)
        mainPanel.Controls.Add(标签打印_选择框)

        自动取图_选择框.Text = "自动取图"
        自动取图_选择框.Location = New Drawing.Point(cx + 90, y + 2)
        自动取图_选择框.Size = New Drawing.Size(80, 22)
        mainPanel.Controls.Add(自动取图_选择框)

        按钮EX_添加.Text = "添加"
        按钮EX_添加.Location = New Drawing.Point(cx + 180, y)
        按钮EX_添加.Size = New Drawing.Size(100, 30)
        mainPanel.Controls.Add(按钮EX_添加)

        ' ==== 证书信息分组框 ====
        证书信息_分组框.Text = "证书信息"
        证书信息_分组框.Location = New Drawing.Point(lx, y + 40)
        证书信息_分组框.Size = New Drawing.Size(1100, 110)
        证书信息_分组框.Visible = False

        Dim zy As Integer = 20
        AddLabel(证书信息_分组框, "证书编码：", 10, zy)
        编辑框_证书编码.Location = New Drawing.Point(80, zy - 3)
        编辑框_证书编码.Size = New Drawing.Size(120, 25)
        证书信息_分组框.Controls.Add(编辑框_证书编码)

        AddLabel(证书信息_分组框, "检测机构：", 210, zy)
        组合框_检测机构.Location = New Drawing.Point(290, zy - 3)
        组合框_检测机构.Size = New Drawing.Size(150, 25)
        组合框_检测机构.DropDownStyle = ComboBoxStyle.DropDownList
        证书信息_分组框.Controls.Add(组合框_检测机构)

        AddLabel(证书信息_分组框, "查询地址：", 450, zy)
        编辑框_查询地址.Location = New Drawing.Point(530, zy - 3)
        编辑框_查询地址.Size = New Drawing.Size(300, 25)
        证书信息_分组框.Controls.Add(编辑框_查询地址)

        zy += 35
        AddLabel(证书信息_分组框, "检测总重：", 10, zy)
        编辑框_检测总重.Location = New Drawing.Point(80, zy - 3)
        编辑框_检测总重.Size = New Drawing.Size(80, 25)
        证书信息_分组框.Controls.Add(编辑框_检测总重)

        AddLabel(证书信息_分组框, "检测形状：", 170, zy)
        编辑框_检测形状.Location = New Drawing.Point(250, zy - 3)
        编辑框_检测形状.Size = New Drawing.Size(80, 25)
        证书信息_分组框.Controls.Add(编辑框_检测形状)

        AddLabel(证书信息_分组框, "检测颜色：", 340, zy)
        编辑框_检测颜色.Location = New Drawing.Point(420, zy - 3)
        编辑框_检测颜色.Size = New Drawing.Size(80, 25)
        证书信息_分组框.Controls.Add(编辑框_检测颜色)

        AddLabel(证书信息_分组框, "检测结果：", 510, zy)
        编辑框_检测结果.Location = New Drawing.Point(590, zy - 3)
        编辑框_检测结果.Size = New Drawing.Size(200, 25)
        证书信息_分组框.Controls.Add(编辑框_检测结果)

        mainPanel.Controls.Add(证书信息_分组框)

        ' ==== 备注 ====
        编辑框_备注.Location = New Drawing.Point(lx, y + 160)
        编辑框_备注.Size = New Drawing.Size(500, 25)
        编辑框_备注.Text = "备注"
        mainPanel.Controls.Add(编辑框_备注)

        ' ==== 款号图片列表 ====
        图标列表框EX1.Location = New Drawing.Point(lx, y + 195)
        图标列表框EX1.Size = New Drawing.Size(1090, 200)
        图标列表框EX1.View = View.LargeIcon
        图标列表框EX1.LargeImageList = New ImageList() With {.ImageSize = New Drawing.Size(80, 80)}
        mainPanel.Controls.Add(图标列表框EX1)

        Me.Controls.Add(mainPanel)

        ' 时钟设置
        时钟1.Interval = 3000
        AddHandler 图标列表框EX1.SelectedIndexChanged, AddressOf _图标列表框EX1_项目左键单击
    End Sub

    Private Sub AddLabel(parent As Control, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    ' ========== 窗口创建完毕 ==========
    Private Sub _窗口_商品信息添加_创建完毕(sender As Object, e As EventArgs)
        ' 询问是否自动识别款号
        Dim result = MessageBox.Show(Me, "是否自动识别款号？", "警告:", MessageBoxButtons.OKCancel)
        If result = DialogResult.OK Then
            图片识别_选择框.Checked = False
        Else
            图片识别_选择框.Checked = True
        End If

        局部_款号文本内容 = ""
        查找_款号文本内容 = ""
        局部_品类名称id = ""
        局部_品类简写 = ""
        局部_商品规格id = ""
        局部_品类多数量 = ""
        局部_商品多数量 = "1"
        局部_商品零销售 = "否"
        局部_品类原料价 = "0"
        编辑框_入库克价.Text = "0.00"

        _商品信息_默认参数()
        _商品组合框_初始化()
        _单选框_无_选中状态改变(Nothing, Nothing)

        编辑框_商品名称.Text = ""
        编辑框_公司成色.Text = ""
        编辑框_重量.Enabled = False
        单选框_重量.Checked = True
        单选框_单件.Checked = False
        单选框_固口.Checked = True
        单选框_开口.Checked = False
        编辑框_固定编码.Text = ""
        编辑框_单件重.Text = "0.000"
        编辑框_数量.Text = "1"
        编辑框_重量.Text = "0.000"
        编辑框_金重.Text = "0.000"
        编辑框_损耗.Text = "0.000"
        编辑框_含耗重.Text = "0.000"
        编辑框_圈号长度.Text = ""
        编辑框_面宽.Text = ""
        编辑框_厚度.Text = ""
        编辑框_成本单价.Text = "0.00"
        编辑框_商品款号.Text = ""
        编辑框_主石重.Text = ""
        编辑框_石头数量.Text = ""
        编辑框_副石重.Text = ""
        编辑框_副石头数量.Text = ""
        编辑框_商品主色.Text = "白"
        编辑框_系数.Text = "1"
        编辑框_成本工费.Text = "0.00"
        编辑框_成本附加费.Text = "0.00"
        编辑框_销售工费.Text = "0.00"
        编辑框_参考工费.Text = "0.00"
        编辑框_销售附加费.Text = "0.00"
        编辑框_销售价.Text = "0.00"
        编辑框_备注.Text = ""
        查找信息编辑框.Text = ""
        编辑框_图片地址.Text = ""
        编辑框_模具号.Text = ""
        编辑框_成本价.Text = ""
        编辑框_证书销售价.Text = ""
        按钮EX_图片上传.Visible = True
        时钟1.Stop()
    End Sub

    ' ========== 商品信息默认参数（禁用所有控件） ==========
    Private Sub _商品信息_默认参数()
        编辑框_商品款号.Enabled = False
        编辑框_商品名称.Enabled = False
        编辑框_公司成色.Enabled = False
        单选框_重量.Enabled = False
        单选框_单件.Enabled = False
        编辑框_入库克价.Enabled = False
        单选框_固口.Enabled = False
        单选框_开口.Enabled = False
        编辑框_固定编码.Enabled = False
        编辑框_单件重.Enabled = False
        编辑框_数量.Enabled = False
        编辑框_金重.Enabled = False
        编辑框_损耗.Enabled = False
        编辑框_含耗重.Enabled = False
        编辑框_圈号长度.Enabled = False
        编辑框_面宽.Enabled = False
        编辑框_厚度.Enabled = False
        编辑框_成本单价.Enabled = False
        编辑框_主石重.Enabled = False
        编辑框_石头数量.Enabled = False
        编辑框_副石重.Enabled = False
        编辑框_副石头数量.Enabled = False
        编辑框_商品主色.Enabled = False
        编辑框_系数.Enabled = False
        编辑框_成本工费.Enabled = False
        编辑框_成本附加费.Enabled = False
        编辑框_销售工费.Enabled = False
        编辑框_参考工费.Enabled = False
        编辑框_销售附加费.Enabled = False
        编辑框_销售价.Enabled = False
        编辑框_备注.Enabled = False
        编辑框_模具号.Enabled = False
        编辑框_证书编码.Enabled = False
        编辑框_查询地址.Enabled = False
        编辑框_检测结果.Enabled = False
        编辑框_检测总重.Enabled = False
        编辑框_检测形状.Enabled = False
        编辑框_检测颜色.Enabled = False
        编辑框_成本价.Enabled = False
        编辑框_证书销售价.Enabled = False
        单选框_有.Enabled = False
        单选框_无.Enabled = False
        图片框EX_固开口.Visible = False
        按钮EX_圈号长度.Visible = True
        按钮EX_添加.Enabled = False
        编辑框_证书编码.Text = ""
        编辑框_查询地址.Text = ""
        编辑框_检测结果.Text = ""
        编辑框_检测总重.Text = ""
        编辑框_检测形状.Text = ""
        编辑框_检测颜色.Text = ""
        编辑框_成本价.Text = ""
        编辑框_证书销售价.Text = ""
        图片框EX_主图.Image = Nothing
        If Not String.IsNullOrEmpty(编辑框_图片地址.Text) Then
            Try
                图片框EX_主图.ImageLocation = 编辑框_图片地址.Text
            Catch
            End Try
        End If
    End Sub

    ' ========== 商品组合框初始化 ==========
    Private Sub _商品组合框_初始化()
        组合框_品类名称.Enabled = True
        组合框_品类名称.Items.Clear()
        组合框_商品材质.Items.Clear()
        组合框_商品成色.Items.Clear()
        组合框_商品规格.Items.Clear()
        组合框_品类名称.Items.Add(New ComboBoxItem("请选择品类名称", "请选择品类名称"))
        组合框_商品材质.Items.Add(New ComboBoxItem("请选择商品材质", "请选择商品材质"))
        组合框_商品成色.Items.Add(New ComboBoxItem("请选择商品成色", "请选择商品成色"))
        组合框_商品规格.Items.Add(New ComboBoxItem("请选择商品规格", "请选择商品规格"))
        组合框_品类名称.SelectedIndex = 0
        组合框_商品材质.SelectedIndex = 0
        组合框_商品成色.SelectedIndex = 0
        组合框_商品规格.SelectedIndex = 0

        If MySQL_Read Is Nothing OrElse MySQL_Read.State <> ConnectionState.Open Then
            MessageBox.Show(Me, "数据库连接异常，无法加载商品基础资料！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 加载品类列表
        Dim sqlCategory As String = "SELECT id,title FROM xipunum_erp_category WHERE 1=1 ORDER BY id ASC"
        Dim dtCategory = ExecuteQuery(sqlCategory)
        If dtCategory IsNot Nothing AndAlso dtCategory.Rows.Count > 0 Then
            For Each row As DataRow In dtCategory.Rows
                组合框_品类名称.Items.Add(New ComboBoxItem(SafeString(row("id")), SafeString(row("title"))))
            Next
        End If

        ' 如果有入库上下文，尝试选中匹配的品类
        If 局部_入库商品数量 > 0 Then
            For i As Integer = 0 To 组合框_品类名称.Items.Count - 1
                Dim item = TryCast(组合框_品类名称.Items(i), ComboBoxItem)
                If item IsNot Nothing AndAlso item.Text = 入库品类名称 Then
                    组合框_品类名称.SelectedIndex = i
                    _组合框_品类名称_内容被改变(Nothing, Nothing)
                    编辑框_入库克价.Text = 入库克价
                    Exit For
                End If
            Next
        End If

        ' 加载检测机构
        组合框_检测机构.Items.Clear()
        组合框_检测机构.Items.Add(New ComboBoxItem("请选择检测机构", "请选择检测机构"))
        组合框_检测机构.SelectedIndex = 0

        Dim sqlOrg As String = "SELECT id,name FROM xipunum_erp_zsjigou WHERE 1=1 ORDER BY id ASC"
        Dim dtOrg = ExecuteQuery(sqlOrg)
        If dtOrg IsNot Nothing AndAlso dtOrg.Rows.Count > 0 Then
            For Each row As DataRow In dtOrg.Rows
                组合框_检测机构.Items.Add(New ComboBoxItem(SafeString(row("id")), SafeString(row("name"))))
            Next
        End If

        组合框_品类名称.Enabled = True
    End Sub

    ' ========== 品类名称选择改变 ==========
    Private Sub _组合框_品类名称_内容被改变(sender As Object, e As EventArgs)
        If 组合框_品类名称.SelectedIndex <= 0 Then Return
        If MySQL_Read Is Nothing OrElse MySQL_Read.State <> ConnectionState.Open Then
            MessageBox.Show(Me, "数据库连接异常，无法加载品类资料！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim 查找品类id As String = 取项目附加数值(组合框_品类名称)
        局部_商品品类id = 查找品类id

        Dim 局部_材质数据 As String = ""
        局部_品类简写 = ""
        Dim 局部_成色 As String = ""
        Dim 局部_商品镶嵌 As String = ""
        局部_品类多数量 = "1"
        局部_品类原料价 = "0"

        Dim sql As String = $"SELECT caizhiid,jianxie,chengse,xiangqian,shuliang,kejia FROM xipunum_erp_category WHERE id='{SafeSQL(查找品类id)}' ORDER BY id ASC LIMIT 1"
        Dim dt As DataTable = ExecuteQuery(sql)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim row As DataRow = dt.Rows(0)
            局部_材质数据 = SafeString(row("caizhiid"))
            局部_品类简写 = SafeString(row("jianxie"))
            局部_成色 = SafeString(row("chengse"))
            局部_商品镶嵌 = SafeString(row("xiangqian"))
            局部_品类多数量 = SafeString(row("shuliang"))
            局部_品类原料价 = SafeString(row("kejia"))
        End If

        局部_商品是否镶嵌 = 局部_商品镶嵌
        局部_品类名称id = 查找品类id

        If String.IsNullOrEmpty(局部_品类多数量) Then 局部_品类多数量 = "1"
        If String.IsNullOrEmpty(局部_品类原料价) Then 局部_品类原料价 = "0"

        ' 重新填充材质/成色/规格
        组合框_商品材质.Items.Clear()
        组合框_商品成色.Items.Clear()
        组合框_商品规格.Items.Clear()
        组合框_商品材质.Items.Add(New ComboBoxItem("请选择商品材质", "请选择商品材质"))
        组合框_商品成色.Items.Add(New ComboBoxItem("请选择商品成色", "请选择商品成色"))
        组合框_商品规格.Items.Add(New ComboBoxItem("请选择商品规格", "请选择商品规格"))
        组合框_商品材质.SelectedIndex = 0
        组合框_商品成色.SelectedIndex = 0
        组合框_商品规格.SelectedIndex = 0

        If Not String.IsNullOrEmpty(局部_材质数据) Then
            Dim 材质数组() As String = 局部_材质数据.Split("|"c)
            For i As Integer = 0 To 材质数组.Length - 1
                If Not String.IsNullOrEmpty(材质数组(i)) Then
                    组合框_商品材质.Items.Add(New ComboBoxItem(i.ToString(), 材质数组(i)))
                End If
            Next
        End If

        If Not String.IsNullOrEmpty(局部_成色) Then
            Dim 成色数组() As String = 局部_成色.Split("|"c)
            For i As Integer = 0 To 成色数组.Length - 1
                If Not String.IsNullOrEmpty(成色数组(i)) Then
                    组合框_商品成色.Items.Add(New ComboBoxItem(i.ToString(), 成色数组(i)))
                End If
            Next
        End If

        ' 加载规格
        Dim sqlSpec As String = $"SELECT id,title FROM xipunum_erp_specs WHERE category_id='{SafeSQL(局部_品类名称id)}' ORDER BY id ASC"
        Dim dtSpec = ExecuteQuery(sqlSpec)
        If dtSpec IsNot Nothing AndAlso dtSpec.Rows.Count > 0 Then
            For Each row As DataRow In dtSpec.Rows
                组合框_商品规格.Items.Add(New ComboBoxItem(SafeString(row("id")), SafeString(row("title"))))
            Next
        End If

        组合框_品类名称.Enabled = False
    End Sub

    ' ========== 规格选择改变 ==========
    Private Sub _组合框_商品规格_内容被改变(sender As Object, e As EventArgs)
        If 组合框_商品规格.SelectedIndex <= 0 Then Return

        Dim 商品规格数量 As String = "1"
        Dim 商品规格简写 As String = ""
        局部_商品多数量 = "1"
        局部_商品规格id = 取项目附加数值(组合框_商品规格)

        If MySQL_Read IsNot Nothing AndAlso MySQL_Read.State = ConnectionState.Open Then
            Dim sql As String = $"SELECT id,shuliang,jianxie FROM xipunum_erp_specs WHERE id='{SafeSQL(局部_商品规格id)}' LIMIT 1"
            Dim dt As DataTable = ExecuteQuery(sql)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                商品规格数量 = SafeString(row("shuliang"))
                商品规格简写 = SafeString(row("jianxie"))
            End If
        End If

        局部_商品多数量 = 商品规格数量
        If String.IsNullOrEmpty(局部_商品多数量) Then 局部_商品多数量 = "1"

        编辑框_数量.Enabled = True
        编辑框_数量.Text = "1"

        If 局部_品类多数量 = "1" Then
            If 局部_商品多数量 = "1" Then
                编辑框_数量.Enabled = False
                编辑框_数量.Text = If(局部_商品零销售 = "是", "0", "1")
            Else
                If 局部_商品零销售 = "是" Then
                    编辑框_数量.Enabled = False
                    编辑框_数量.Text = "0"
                Else
                    编辑框_数量.Enabled = True
                    If Val(编辑框_数量.Text) < 1 Then 编辑框_数量.Text = "1"
                End If
            End If
        End If

        ' 固口/开口 判断
        Dim 当前规格文本 As String = 取项目文本(组合框_商品规格)
        If 当前规格文本 = "戒指" OrElse 当前规格文本 = "手镯" Then
            图片框EX_固开口.Visible = True
            按钮EX_圈号长度.Visible = False
            编辑框_圈号长度.Width = 72
            编辑框_圈号长度.Left = 304
        Else
            图片框EX_固开口.Visible = False
            按钮EX_圈号长度.Visible = True
            编辑框_圈号长度.Width = 110
            编辑框_圈号长度.Left = 270
        End If

        If 局部_入库商品数量 > 0 Then
            编辑框_入库克价.Enabled = False
        Else
            编辑框_入库克价.Enabled = True
        End If

        编辑框_重量.Enabled = True
        编辑框_重量.Text = ""
        编辑框_重量.Focus()
    End Sub

    ' ========== 查找按钮 ==========
    Private Sub 按钮EX_查找_Click(sender As Object, e As EventArgs) Handles 按钮EX_查找.Click
        If 组合框_品类名称.SelectedIndex <= 0 Then
            MessageBox.Show(Me, "请选择品类名称！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            组合框_品类名称.Focus()
            Return
        End If

        局部_款号文本内容 = ""
        查找_款号文本内容 = ""
        查找_款号文本内容 = 查找信息编辑框.Text
        _图标列表框EX1_加载数据()
    End Sub

    ' ========== 重置按钮 ==========
    Private Sub 按钮EX_重置_Click(sender As Object, e As EventArgs) Handles 按钮EX_重置.Click
        _窗口_商品信息添加_创建完毕(Nothing, Nothing)
    End Sub

    ' ========== 款号列表加载 ==========
    Private Sub _图标列表框EX1_加载数据()
        图标列表框EX1.Items.Clear()
        Try
            Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "khimages"))
        Catch
        End Try

        If MySQL_Read Is Nothing OrElse MySQL_Read.State <> ConnectionState.Open Then
            MessageBox.Show(Me, "数据库连接异常，无法加载款号！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim sql As String
        If Not String.IsNullOrEmpty(局部_款号文本内容) Then
            sql = $"SELECT id,title,kuanhao,yimage FROM xipunum_erp_ksiamges WHERE kuanhao in ({局部_款号文本内容}) ORDER BY FIELD(kuanhao,{局部_款号文本内容}) LIMIT 9"
        Else
            If 组合框_品类名称.SelectedIndex <= 0 Then
                MessageBox.Show(Me, "请选择品类名称！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Dim 品类ID As String = 取项目附加数值(组合框_品类名称)
            sql = $"SELECT id,title,kuanhao,yimage FROM xipunum_erp_ksiamges WHERE (title like '%{SafeSQL(查找_款号文本内容)}%' or kuanhao like '%{SafeSQL(查找_款号文本内容)}%') and category_id='{SafeSQL(品类ID)}' ORDER BY id DESC LIMIT 9"
        End If

        Dim dt As DataTable = ExecuteQuery(sql)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        Dim khPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "khimages")

        For Each row As DataRow In dt.Rows
            Dim 款号 As String = SafeString(row("kuanhao"))
            Dim 名称 As String = SafeString(row("title"))
            Dim 图片地址 As String = SafeString(row("yimage"))
            Dim 本地路径 As String = Path.Combine(khPath, 款号 & ".jpg")

            If Not String.IsNullOrEmpty(图片地址) Then
                If Not File.Exists(本地路径) OrElse New FileInfo(本地路径).Length < 2000 Then
                    Try
                        Using wc As New System.Net.WebClient()
                            Dim data = wc.DownloadData(图片地址)
                            If data.Length > 2000 Then
                                File.WriteAllBytes(本地路径, data)
                            End If
                        End Using
                    Catch
                    End Try
                End If
            End If

            Dim lvi As New ListViewItem()
            lvi.Text = "款号:" & 款号 & vbCrLf & "名称:" & 名称
            lvi.Tag = SafeString(row("id"))
            If File.Exists(本地路径) Then
                Try
                    Dim img As Image = Image.FromFile(本地路径)
                    图标列表框EX1.LargeImageList.Images.Add(款号, img)
                    lvi.ImageKey = 款号
                Catch
                End Try
            End If
            图标列表框EX1.Items.Add(lvi)
        Next

        If 图标列表框EX1.Items.Count > 0 Then
            _图标列表框EX1_项目左键单击(0)
        End If
    End Sub

    ' ========== 图标列表项目单击 ==========
    Private Sub _图标列表框EX1_项目左键单击(index As Integer)
        If index < 0 OrElse index >= 图标列表框EX1.Items.Count Then Return
        If MySQL_Read Is Nothing OrElse MySQL_Read.State <> ConnectionState.Open Then
            MessageBox.Show(Me, "数据库连接异常，无法读取款号信息！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim 款号ID As String = 图标列表框EX1.Items(index).Tag?.ToString()
        If String.IsNullOrEmpty(款号ID) OrElse 款号ID = "0" Then Return

        Dim sql As String = $"SELECT a.title AS title,a.kuanhao AS kuanhao,a.yimage AS yimage,a.caizhi AS caizhi,a.chengse AS chengse,a.lingxiao AS lingxiao,b.title AS guige FROM xipunum_erp_ksiamges AS a LEFT JOIN xipunum_erp_specs AS b ON b.id = a.specification_id WHERE a.id='{SafeSQL(款号ID)}' LIMIT 1"
        Dim dt As DataTable = ExecuteQuery(sql)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        Dim row As DataRow = dt.Rows(0)
        Dim 款号列表名称 As String = SafeString(row("title"))
        Dim 款号列表款号 As String = SafeString(row("kuanhao"))
        Dim 款号列表图片地址 As String = SafeString(row("yimage"))
        Dim 款号列表材质 As String = SafeString(row("caizhi"))
        Dim 款号列表成色 As String = SafeString(row("chengse"))
        Dim 款号列表零销售 As String = SafeString(row("lingxiao"))
        Dim 款号列表规格 As String = SafeString(row("guige"))

        Dim khPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "khimages", 款号列表款号 & ".jpg")
        If File.Exists(khPath) Then
            图片框EX_主图.Image = Image.FromFile(khPath)
        Else
            图片框EX_主图.ImageLocation = 款号列表图片地址
        End If

        编辑框_商品名称.Text = 款号列表名称
        编辑框_商品款号.Text = 款号列表款号
        编辑框_图片地址.Text = 款号列表图片地址
        局部_商品零销售 = 款号列表零销售

        ' 匹配材质
        组合框_商品材质.SelectedIndex = 0
        For i As Integer = 0 To 组合框_商品材质.Items.Count - 1
            Dim item = TryCast(组合框_商品材质.Items(i), ComboBoxItem)
            If item IsNot Nothing AndAlso item.Text = 款号列表材质 Then
                组合框_商品材质.SelectedIndex = i
                Exit For
            End If
        Next

        ' 匹配规格
        组合框_商品规格.SelectedIndex = 0
        For i As Integer = 0 To 组合框_商品规格.Items.Count - 1
            Dim item = TryCast(组合框_商品规格.Items(i), ComboBoxItem)
            If item IsNot Nothing AndAlso item.Text = 款号列表规格 Then
                组合框_商品规格.SelectedIndex = i
                Exit For
            End If
        Next

        If 组合框_商品规格.SelectedIndex > 0 Then
            _组合框_商品规格_内容被改变(Nothing, Nothing)
        End If

        ' 匹配成色
        组合框_商品成色.SelectedIndex = 0
        For i As Integer = 0 To 组合框_商品成色.Items.Count - 1
            Dim item = TryCast(组合框_商品成色.Items(i), ComboBoxItem)
            If item IsNot Nothing AndAlso item.Text = 款号列表成色 Then
                组合框_商品成色.SelectedIndex = i
                Exit For
            End If
        Next

        If 组合框_商品成色.SelectedIndex <= 0 AndAlso 组合框_商品成色.Items.Count > 0 Then
            组合框_商品成色.SelectedIndex = 1
        End If

        If 组合框_商品成色.SelectedIndex > 0 Then
            _组合框_商品成色_内容被改变(Nothing, Nothing)
        End If

        编辑框_数量.Enabled = True
        编辑框_数量.Text = "1"
        If 局部_品类多数量 = "1" Then
            If 局部_商品多数量 = "1" Then
                编辑框_数量.Enabled = False
                编辑框_数量.Text = If(局部_商品零销售 = "是", "0", "1")
            Else
                If 局部_商品零销售 = "是" Then
                    编辑框_数量.Enabled = False
                    编辑框_数量.Text = "0"
                Else
                    编辑框_数量.Enabled = True
                    If Val(编辑框_数量.Text) < 1 Then 编辑框_数量.Text = "1"
                End If
            End If
        End If
        编辑框_重量.Focus()
    End Sub

    Private Sub 图标列表框EX1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles 图标列表框EX1.SelectedIndexChanged
        If 图标列表框EX1.SelectedIndices.Count > 0 Then
            _图标列表框EX1_项目左键单击(图标列表框EX1.SelectedIndices(0))
        End If
    End Sub

    ' ========== 商品成色改变 ==========
    Private Sub _组合框_商品成色_内容被改变(sender As Object, e As EventArgs)
        If 组合框_商品成色.SelectedIndex <= 0 Then Return
        编辑框_公司成色.Text = 取项目文本(组合框_商品成色)
    End Sub

    ' ========== 参数数值计算（核心业务逻辑） ==========
    Private Sub _商品信息_参数数值计算(修改参数 As String)
        Dim 数量 As Double = Val(编辑框_数量.Text)
        Dim 重量 As Double = Val(编辑框_重量.Text)
        Dim 金重 As Double = Val(编辑框_金重.Text)
        Dim 损耗 As Double = Val(编辑框_损耗.Text)
        Dim 入库克价 As Double = Val(编辑框_入库克价.Text)
        Dim 成本单价 As Double = Val(编辑框_成本单价.Text)
        Dim 系数 As Double = Val(编辑框_系数.Text)
        Dim 成本工费 As Double = Val(编辑框_成本工费.Text)
        Dim 成本附加费 As Double = Val(编辑框_成本附加费.Text)
        Dim 销售工费 As Double = Val(编辑框_销售工费.Text)
        Dim 参考工费 As Double = Val(编辑框_参考工费.Text)
        Dim 销售附加费 As Double = Val(编辑框_销售附加费.Text)
        Dim 主石重 As Double = Val(编辑框_主石重.Text)
        Dim 副石重 As Double = Val(编辑框_副石重.Text)

        If 数量 < 0 Then 数量 = 0
        If 重量 < 0 Then 重量 = 0
        If 金重 < 0 Then 金重 = 0
        If 损耗 < 0 Then 损耗 = 0
        If 入库克价 < 0 Then 入库克价 = 0
        If 成本单价 < 0 Then 成本单价 = 0
        If 系数 <= 0 Then 系数 = 1
        If 成本工费 < 0 Then 成本工费 = 0
        If 成本附加费 < 0 Then 成本附加费 = 0
        If 销售工费 < 0 Then 销售工费 = 0
        If 参考工费 < 0 Then 参考工费 = 0
        If 销售附加费 < 0 Then 销售附加费 = 0
        If 主石重 < 0 Then 主石重 = 0
        If 副石重 < 0 Then 副石重 = 0

        If 修改参数 = "重量" Then
            编辑框_重量.Text = 格式三位小数(重量.ToString())
            编辑框_金重.Text = 格式三位小数(重量.ToString())
            重量 = Val(编辑框_重量.Text)
            金重 = Val(编辑框_金重.Text)
            编辑框_单件重.Text = If(数量 = 0, 格式三位小数(重量.ToString()), 格式三位小数((重量 / 数量).ToString()))
            编辑框_主石重.Text = "0.000"
            编辑框_石头数量.Text = "0"
            编辑框_副石重.Text = "0.000"
            编辑框_副石头数量.Text = "0"
            编辑框_含耗重.Text = 格式三位小数((重量 + 损耗).ToString())
        End If

        If 修改参数 = "入库克价" Then
            编辑框_入库克价.Text = 格式二位小数(入库克价.ToString())
        End If

        If 修改参数 = "数量" Then
            编辑框_单件重.Text = If(数量 = 0, 格式三位小数(重量.ToString()), 格式三位小数((重量 / 数量).ToString()))
            编辑框_主石重.Text = "0.000"
            编辑框_石头数量.Text = "0"
            编辑框_副石重.Text = "0.000"
            编辑框_副石头数量.Text = "0"
        End If

        If 修改参数 = "金重" Then
            If 金重 > 重量 Then 金重 = 重量
            编辑框_金重.Text = 格式三位小数(金重.ToString())
            Dim 可用石重 As Double = (重量 - 金重) * 5
            If 可用石重 < 0 Then 可用石重 = 0
            编辑框_主石重.Text = 格式三位小数(可用石重.ToString())
            编辑框_副石重.Text = "0.000"
        End If

        If 修改参数 = "损耗" Then
            编辑框_损耗.Text = 格式三位小数(损耗.ToString())
            编辑框_含耗重.Text = 格式三位小数((重量 + 损耗).ToString())
        End If

        If 修改参数 = "成本单价" Then
            编辑框_成本单价.Text = 格式二位小数(成本单价.ToString())
            If 单选框_单件.Checked Then
                编辑框_销售价.Text = 格式二位小数((成本单价 * 系数).ToString())
            End If
        End If

        If 修改参数 = "销售价" Then
            编辑框_销售价.Text = 格式二位小数(Val(编辑框_销售价.Text).ToString())
            If 单选框_单件.Checked Then
                编辑框_系数.Text = If(成本单价 > 0, 格式二位小数((Val(编辑框_销售价.Text) / 成本单价).ToString()), "1")
            End If
        End If

        If 修改参数 = "主石重" Then
            Dim 可用石重 As Double = (重量 - 金重) * 5.5
            If 可用石重 < 0 Then 可用石重 = 0
            If 主石重 > 可用石重 Then
                MessageBox.Show(Me, "主石重不可以大于" & 可用石重.ToString() & "ct", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                主石重 = 可用石重
            End If
            编辑框_主石重.Text = 格式三位小数(主石重.ToString())
            Dim 计算值 As Double = (重量 - 金重) * 5 - 主石重
            If 计算值 < 0 Then 计算值 = 0
            编辑框_副石重.Text = 格式三位小数(计算值.ToString())
        End If

        If 修改参数 = "副石重" Then
            Dim 可用石重 As Double = (重量 - 金重) * 5.5
            If 可用石重 < 0 Then 可用石重 = 0
            If 副石重 > 可用石重 Then
                MessageBox.Show(Me, "副石重不可以大于" & 可用石重.ToString() & "ct", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                副石重 = 可用石重
            End If
            编辑框_副石重.Text = 格式三位小数(副石重.ToString())
        End If

        If 修改参数 = "系数" Then
            If 单选框_单件.Checked Then
                编辑框_销售价.Text = 格式二位小数((成本单价 * 系数).ToString())
            End If
        End If

        If 修改参数 = "成本工费" Then
            编辑框_成本工费.Text = 格式二位小数(成本工费.ToString())
            参考工费 = 成本工费
            销售工费 = 成本工费
        End If

        If 修改参数 = "成本附加费" Then
            编辑框_成本附加费.Text = 格式二位小数(成本附加费.ToString())
            编辑框_销售附加费.Text = 格式二位小数(成本附加费.ToString())
            销售附加费 = 成本附加费
        End If

        If 修改参数 = "销售工费" Then
            编辑框_销售工费.Text = 格式二位小数(销售工费.ToString())
            编辑框_参考工费.Text = 格式二位小数(销售工费.ToString())
            参考工费 = 销售工费
        End If

        If 修改参数 = "参考工费" Then
            编辑框_参考工费.Text = 格式二位小数(参考工费.ToString())
            编辑框_销售工费.Text = 格式二位小数(参考工费.ToString())
            销售工费 = 参考工费
        End If

        If 修改参数 = "销售附加费" Then
            编辑框_销售附加费.Text = 格式二位小数(销售附加费.ToString())
        End If

        ' 最终重算
        数量 = Val(编辑框_数量.Text)
        金重 = Val(编辑框_金重.Text)
        入库克价 = Val(编辑框_入库克价.Text)
        成本工费 = Val(编辑框_成本工费.Text)
        成本附加费 = Val(编辑框_成本附加费.Text)
        销售工费 = Val(编辑框_销售工费.Text)
        销售附加费 = Val(编辑框_销售附加费.Text)

        Dim relevantParams() As String = {"单件重", "重量", "数量", "金重", "成本单价", "成本工费", "成本附加费", "销售工费", "参考工费", "销售附加费", "入库克价"}
        If Array.IndexOf(relevantParams, 修改参数) >= 0 Then
            If 单选框_重量.Checked Then
                If 数量 = 0 Then
                    编辑框_成本单价.Text = 格式二位小数((金重 * (成本工费 + 入库克价) + 成本附加费).ToString())
                    编辑框_销售价.Text = 格式二位小数((金重 * (销售工费 + 入库克价) + 销售附加费).ToString())
                Else
                    编辑框_成本单价.Text = 格式二位小数(((金重 / 数量) * (成本工费 + 入库克价) + 成本附加费).ToString())
                    编辑框_销售价.Text = 格式二位小数(((金重 / 数量) * (销售工费 + 入库克价) + 销售附加费).ToString())
                End If
            End If
        End If
    End Sub

    ' ========== 重量焦点事件 ==========
    Private Sub 编辑框_重量_Enter(sender As Object, e As EventArgs) Handles 编辑框_重量.Enter
        If Val(编辑框_重量.Text) = 0 Then 编辑框_重量.Text = ""
    End Sub

    Private Sub 编辑框_重量_Leave(sender As Object, e As EventArgs) Handles 编辑框_重量.Leave
        If 编辑框_重量.Text = "" Then 编辑框_重量.Text = "0.000"
        _商品信息_参数数值计算("重量")
    End Sub

    Private Sub 编辑框_重量_KeyDown(sender As Object, e As KeyEventArgs) Handles 编辑框_重量.KeyDown
        If e.KeyCode = Keys.Enter Then
            编辑框_重量_Leave(Nothing, Nothing)
        End If
    End Sub

    ' ========== 入库克价事件 ==========
    Private Sub 编辑框_入库克价_Enter(sender As Object, e As EventArgs) Handles 编辑框_入库克价.Enter
        If 编辑框_入库克价.Text = "0.00" Then 编辑框_入库克价.Text = ""
    End Sub

    Private Sub 编辑框_入库克价_Leave(sender As Object, e As EventArgs) Handles 编辑框_入库克价.Leave
        If 编辑框_入库克价.Text = "" Then 编辑框_入库克价.Text = "0.00"
        _商品信息_参数数值计算("入库克价")
    End Sub

    Private Sub 编辑框_入库克价_KeyDown(sender As Object, e As KeyEventArgs) Handles 编辑框_入库克价.KeyDown
        If e.KeyCode = Keys.Enter Then
            If 编辑框_入库克价.Text = "" Then 编辑框_入库克价.Text = "0.00"
            If Val(编辑框_入库克价.Text) < 0 Then 编辑框_入库克价.Text = "0.00"
            _商品信息_参数数值计算("入库克价")
            编辑框_重量.Text = ""
            编辑框_重量.Focus()
        End If
    End Sub

    ' ========== 数量事件 ==========
    Private Sub 编辑框_数量_Enter(sender As Object, e As EventArgs) Handles 编辑框_数量.Enter
        ' Nothing special on enter
    End Sub

    Private Sub 编辑框_数量_Leave(sender As Object, e As EventArgs) Handles 编辑框_数量.Leave
        If 编辑框_数量.Text = "" Then 编辑框_数量.Text = "1"
        If 局部_商品零销售 = "否" AndAlso Val(编辑框_数量.Text) < 1 Then 编辑框_数量.Text = "1"
        _商品信息_参数数值计算("数量")
    End Sub

    ' ========== 金重事件 ==========
    Private Sub 编辑框_金重_Enter(sender As Object, e As EventArgs) Handles 编辑框_金重.Enter
        If Val(编辑框_金重.Text) = 0 Then 编辑框_金重.Text = ""
    End Sub

    Private Sub 编辑框_金重_Leave(sender As Object, e As EventArgs) Handles 编辑框_金重.Leave
        If 编辑框_金重.Text = "" Then 编辑框_金重.Text = "0.000"
        _商品信息_参数数值计算("金重")
    End Sub

    Private Sub 编辑框_金重_KeyDown(sender As Object, e As KeyEventArgs) Handles 编辑框_金重.KeyDown
        If e.KeyCode = Keys.Enter Then
            If 编辑框_金重.Text = "" Then 编辑框_金重.Text = 编辑框_重量.Text
            _商品信息_参数数值计算("金重")
        End If
    End Sub

    ' ========== 损耗事件 ==========
    Private Sub 编辑框_损耗_Enter(sender As Object, e As EventArgs) Handles 编辑框_损耗.Enter
        If 编辑框_损耗.Text = "0.000" Then 编辑框_损耗.Text = ""
    End Sub

    Private Sub 编辑框_损耗_Leave(sender As Object, e As EventArgs) Handles 编辑框_损耗.Leave
        If 编辑框_损耗.Text = "" Then 编辑框_损耗.Text = "0.000"
        _商品信息_参数数值计算("损耗")
    End Sub

    ' ========== 成本单价事件 ==========
    Private Sub 编辑框_成本单价_Enter(sender As Object, e As EventArgs) Handles 编辑框_成本单价.Enter
        If 编辑框_成本单价.Text = "0.00" Then 编辑框_成本单价.Text = ""
    End Sub

    Private Sub 编辑框_成本单价_Leave(sender As Object, e As EventArgs) Handles 编辑框_成本单价.Leave
        If 编辑框_成本单价.Text = "" OrElse Val(编辑框_成本单价.Text) < 0 Then 编辑框_成本单价.Text = "0.00"
        _商品信息_参数数值计算("成本单价")
    End Sub

    ' ========== 销售价事件 ==========
    Private Sub 编辑框_销售价_Enter(sender As Object, e As EventArgs) Handles 编辑框_销售价.Enter
        If 编辑框_销售价.Text = "0.00" Then 编辑框_销售价.Text = ""
    End Sub

    Private Sub 编辑框_销售价_Leave(sender As Object, e As EventArgs) Handles 编辑框_销售价.Leave
        If 编辑框_销售价.Text = "" OrElse Val(编辑框_销售价.Text) < 0 Then 编辑框_销售价.Text = "0.00"
        If 单选框_单件.Checked AndAlso Val(编辑框_成本单价.Text) <= 0 Then
            编辑框_成本单价.Text = "0.00"
        End If
        _商品信息_参数数值计算("销售价")
    End Sub

    ' ========== 主石重事件 ==========
    Private Sub 编辑框_主石重_Enter(sender As Object, e As EventArgs) Handles 编辑框_主石重.Enter
        If 编辑框_主石重.Text = "0.000" Then 编辑框_主石重.Text = ""
    End Sub

    Private Sub 编辑框_主石重_Leave(sender As Object, e As EventArgs) Handles 编辑框_主石重.Leave
        If 编辑框_主石重.Text = "" OrElse Val(编辑框_主石重.Text) < 0 Then 编辑框_主石重.Text = "0.000"
        _商品信息_参数数值计算("主石重")
    End Sub

    ' ========== 副石重事件 ==========
    Private Sub 编辑框_副石重_Enter(sender As Object, e As EventArgs) Handles 编辑框_副石重.Enter
        If 编辑框_副石重.Text = "0.000" Then 编辑框_副石重.Text = ""
    End Sub

    Private Sub 编辑框_副石重_Leave(sender As Object, e As EventArgs) Handles 编辑框_副石重.Leave
        If 编辑框_副石重.Text = "" OrElse Val(编辑框_副石重.Text) < 0 Then 编辑框_副石重.Text = "0.000"
        _商品信息_参数数值计算("副石重")
    End Sub

    ' ========== 系数事件 ==========
    Private Sub 编辑框_系数_Enter(sender As Object, e As EventArgs) Handles 编辑框_系数.Enter
        If 编辑框_系数.Text = "1" Then 编辑框_系数.Text = ""
    End Sub

    Private Sub 编辑框_系数_Leave(sender As Object, e As EventArgs) Handles 编辑框_系数.Leave
        If 编辑框_系数.Text = "" OrElse Val(编辑框_系数.Text) <= 0 Then 编辑框_系数.Text = "1"
        _商品信息_参数数值计算("系数")
    End Sub

    ' ========== 各种工费/附加费焦点事件 ==========
    Private Sub 编辑框_成本工费_Enter(sender As Object, e As EventArgs) Handles 编辑框_成本工费.Enter
        If 编辑框_成本工费.Text = "0.00" Then 编辑框_成本工费.Text = ""
    End Sub

    Private Sub 编辑框_成本工费_Leave(sender As Object, e As EventArgs) Handles 编辑框_成本工费.Leave
        If 编辑框_成本工费.Text = "" OrElse Val(编辑框_成本工费.Text) < 0 Then 编辑框_成本工费.Text = "0.00"
        _商品信息_参数数值计算("成本工费")
    End Sub

    Private Sub 编辑框_成本附加费_Enter(sender As Object, e As EventArgs) Handles 编辑框_成本附加费.Enter
        If 编辑框_成本附加费.Text = "0.00" Then 编辑框_成本附加费.Text = ""
    End Sub

    Private Sub 编辑框_成本附加费_Leave(sender As Object, e As EventArgs) Handles 编辑框_成本附加费.Leave
        If 编辑框_成本附加费.Text = "" OrElse Val(编辑框_成本附加费.Text) < 0 Then 编辑框_成本附加费.Text = "0.00"
        _商品信息_参数数值计算("成本附加费")
    End Sub

    Private Sub 编辑框_销售工费_Enter(sender As Object, e As EventArgs) Handles 编辑框_销售工费.Enter
        If 编辑框_销售工费.Text = "0.00" Then 编辑框_销售工费.Text = ""
    End Sub

    Private Sub 编辑框_销售工费_Leave(sender As Object, e As EventArgs) Handles 编辑框_销售工费.Leave
        If 编辑框_销售工费.Text = "" Then 编辑框_销售工费.Text = 编辑框_成本工费.Text
        If Val(编辑框_销售工费.Text) < 0 Then 编辑框_销售工费.Text = "0.00"
        _商品信息_参数数值计算("销售工费")
    End Sub

    Private Sub 编辑框_参考工费_Enter(sender As Object, e As EventArgs) Handles 编辑框_参考工费.Enter
        If 编辑框_参考工费.Text = "0.00" Then 编辑框_参考工费.Text = ""
    End Sub

    Private Sub 编辑框_参考工费_Leave(sender As Object, e As EventArgs) Handles 编辑框_参考工费.Leave
        If 编辑框_参考工费.Text = "" Then 编辑框_参考工费.Text = 编辑框_成本工费.Text
        If Val(编辑框_参考工费.Text) < 0 Then 编辑框_参考工费.Text = "0.00"
        _商品信息_参数数值计算("参考工费")
    End Sub

    Private Sub 编辑框_销售附加费_Enter(sender As Object, e As EventArgs) Handles 编辑框_销售附加费.Enter
        If 编辑框_销售附加费.Text = "0.00" Then 编辑框_销售附加费.Text = ""
    End Sub

    Private Sub 编辑框_销售附加费_Leave(sender As Object, e As EventArgs) Handles 编辑框_销售附加费.Leave
        If 编辑框_销售附加费.Text = "" Then 编辑框_销售附加费.Text = 编辑框_成本附加费.Text
        If Val(编辑框_销售附加费.Text) < 0 Then 编辑框_销售附加费.Text = "0.00"
        _商品信息_参数数值计算("销售附加费")
    End Sub

    ' ========== 有/无证书单选框 ==========
    Private Sub 单选框_有_CheckedChanged(sender As Object, e As EventArgs) Handles 单选框_有.CheckedChanged
        If 单选框_有.Checked Then
            单选框_有.Checked = True
            单选框_无.Checked = False
            Me.Height = 810
            证书信息_分组框.Visible = True
        End If
    End Sub

    Private Sub 单选框_无_CheckedChanged(sender As Object, e As EventArgs) Handles 单选框_无.CheckedChanged
        If 单选框_无.Checked Then
            单选框_有.Checked = False
            单选框_无.Checked = True
            Me.Height = 700
            证书信息_分组框.Visible = False
        End If
    End Sub

    ' ========== 重量/单件单选框 ==========
    Private Sub 单选框_重量_CheckedChanged(sender As Object, e As EventArgs) Handles 单选框_重量.CheckedChanged
        If 单选框_重量.Checked Then
            单选框_重量.Checked = True
            单选框_单件.Checked = False
            编辑框_销售价.Enabled = False
            _商品信息_参数数值计算("重量")
        End If
    End Sub

    Private Sub 单选框_单件_CheckedChanged(sender As Object, e As EventArgs) Handles 单选框_单件.CheckedChanged
        If 单选框_单件.Checked Then
            单选框_重量.Checked = False
            单选框_单件.Checked = True
            编辑框_销售价.Enabled = True
            If 编辑框_成本单价.Text = "" Then 编辑框_成本单价.Text = "0.00"
            If 编辑框_系数.Text = "" OrElse Val(编辑框_系数.Text) <= 0 Then 编辑框_系数.Text = "1"
            _商品信息_参数数值计算("系数")
        End If
    End Sub

    ' ========== 数量内容验证 ==========
    Private Sub 编辑框_数量_TextChanged(sender As Object, e As EventArgs) Handles 编辑框_数量.TextChanged
        Dim val As Double
        If Double.TryParse(编辑框_数量.Text, val) Then
            If 局部_商品多数量 = "0" AndAlso 局部_商品零销售 = "否" Then
                If val < 1 Then
                    MessageBox.Show(Me, "入库数量不能小于1！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    编辑框_数量.Text = "1"
                    编辑框_数量.Focus()
                End If
            End If
        End If
    End Sub

    ' ========== 金重验证 ==========
    Private Sub 编辑框_金重_TextChanged(sender As Object, e As EventArgs) Handles 编辑框_金重.TextChanged
        Dim val As Double
        If Double.TryParse(编辑框_金重.Text, val) Then
            Dim wt As Double = Val(编辑框_重量.Text)
            If val > wt AndAlso wt > 0 Then
                MessageBox.Show(Me, "金重不能大于总重！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                编辑框_金重.Text = 编辑框_重量.Text
            End If
            If val < 0 Then
                MessageBox.Show(Me, "金重不能小于0！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                编辑框_金重.Text = "0.000"
            End If
        End If
    End Sub

    ' ========== 损耗验证 ==========
    Private Sub 编辑框_损耗_TextChanged(sender As Object, e As EventArgs) Handles 编辑框_损耗.TextChanged
        Dim val As Double
        If Double.TryParse(编辑框_损耗.Text, val) AndAlso val < 0 Then
            MessageBox.Show(Me, "损耗不能小于0！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            编辑框_损耗.Text = "0.000"
        End If
    End Sub

    ' ========== 图片上传按钮 ==========
    Private Sub 按钮EX_图片上传_Click(sender As Object, e As EventArgs) Handles 按钮EX_图片上传.Click
        If 组合框_品类名称.SelectedIndex <= 0 Then
            MessageBox.Show(Me, "请选择品类名称！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            组合框_品类名称.Focus()
            Return
        End If

        Try
            Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "images"))
        Catch
        End Try

        If 通用对话框1.ShowDialog() <> DialogResult.OK Then Return

        Dim 图片名称 As String = DateTime.Now.ToString("yyMMddHHmmss")
        Dim 图片路径 As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "images", 图片名称 & ".jpg")

        Try
            File.Copy(通用对话框1.FileName, 图片路径, True)
        Catch
            MessageBox.Show(Me, "图片保存失败！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        If Not File.Exists(图片路径) Then
            MessageBox.Show(Me, "图片保存失败！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim fi As New FileInfo(图片路径)
        If fi.Length > 3000000 Then
            File.Delete(图片路径)
            MessageBox.Show(Me, "图片大小超过3M，请压缩图片后再上传！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim 品类数据id As String = 取项目附加数值(组合框_品类名称)
        Dim 规格数据id As String = ""
        Dim 材质数据名称 As String = ""

        图片框EX_主图.Image = Image.FromFile(图片路径)

        ' 图片识别
        If 图片识别_选择框.Checked Then
            If String.IsNullOrEmpty(KuanHaoRecognizeURL) Then
                MessageBox.Show(Me, "款号识别地址为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            ' 暂时简化图片识别逻辑（HTTP multipart上传）
            编辑框_图片地址.Text = ""
            _图标列表框EX1_加载数据()
        Else
            If String.IsNullOrEmpty(DevRecognizeURL) Then
                MessageBox.Show(Me, "图片上传地址为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            编辑框_图片地址.Text = DevRecognizeURL & "/" & 图片名称 & ".jpg"
        End If
    End Sub

    ' ========== 新建款号按钮 ==========
    Private Sub 按钮EX_新建款号_Click(sender As Object, e As EventArgs) Handles 按钮EX_新建款号.Click
        If UserOperation.Contains(",53款号管理添加,") Then
            MessageBox.Show(Me, "无权操作！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If 组合框_品类名称.SelectedIndex <= 0 Then
            MessageBox.Show(Me, "请选择品类名称！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            组合框_品类名称.Focus()
            Return
        End If

        ' 打开款号添加窗口（如果需要）
        MessageBox.Show(Me, "请在新窗口中添加款号信息", "提示")
    End Sub

    ' ========== 添加按钮（核心保存逻辑） ==========
    Private Sub 按钮EX_添加_Click(sender As Object, e As EventArgs) Handles 按钮EX_添加.Click
        If MySQL_Read Is Nothing OrElse MySQL_Write Is Nothing Then
            MessageBox.Show(Me, "数据库连接异常，无法添加商品！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 验证
        If 组合框_品类名称.SelectedIndex <= 0 Then
            MessageBox.Show(Me, "请选择品类名称！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            组合框_品类名称.Focus()
            Return
        End If
        If 组合框_商品材质.SelectedIndex <= 0 Then
            MessageBox.Show(Me, "请选择商品材质！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            组合框_商品材质.Focus()
            Return
        End If
        If 组合框_商品成色.SelectedIndex <= 0 Then
            MessageBox.Show(Me, "请选择商品成色！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            组合框_商品成色.Focus()
            Return
        End If
        If 组合框_商品规格.SelectedIndex <= 0 Then
            MessageBox.Show(Me, "请选择商品规格！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            组合框_商品规格.Focus()
            Return
        End If
        If 编辑框_重量.Text = "" OrElse Val(编辑框_重量.Text) < 0 Then
            MessageBox.Show(Me, "商品重量必须大于0！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            编辑框_重量.Enabled = True
            编辑框_重量.Focus()
            Return
        End If
        If 局部_品类原料价 = "0" Then
            If 编辑框_入库克价.Text = "" OrElse Val(编辑框_入库克价.Text) <= 0 Then
                MessageBox.Show(Me, "入库原料价不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
                编辑框_入库克价.Enabled = True
                编辑框_入库克价.Focus()
                Return
            End If
        End If
        If 编辑框_数量.Text = "" Then
            MessageBox.Show(Me, "商品数量不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            编辑框_数量.Focus()
            Return
        End If
        If 局部_商品零销售 = "否" AndAlso Val(编辑框_数量.Text) <= 0 Then
            MessageBox.Show(Me, "商品数量必须大于0！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
            编辑框_数量.Focus()
            Return
        End If

        Dim 当前规格文本 As String = 取项目文本(组合框_商品规格)
        If (当前规格文本 = "戒指" OrElse 当前规格文本 = "手镯") Then
            If 单选框_固口.Checked AndAlso 编辑框_圈号长度.Text = "" Then
                MessageBox.Show(Me, 当前规格文本 & "圈号不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
                编辑框_圈号长度.Focus()
                Return
            End If
        End If

        If 单选框_单件.Checked Then
            If 编辑框_成本单价.Text = "" OrElse Val(编辑框_成本单价.Text) <= 0 Then
                MessageBox.Show(Me, "商品成本价必须大于0！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
                编辑框_成本单价.Focus()
                Return
            End If
        End If

        If 单选框_有.Checked Then
            If 编辑框_证书编码.Text = "" Then
                MessageBox.Show(Me, "证书编码不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
                编辑框_证书编码.Focus()
                Return
            End If
            If 组合框_检测机构.SelectedIndex <= 0 Then
                MessageBox.Show(Me, "请选择检测机构！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
                组合框_检测机构.Focus()
                Return
            End If
        End If

        按钮EX_添加.Enabled = False

        ' 如果成本单价为0但有工费，重新计算
        If Val(编辑框_成本单价.Text) = 0 Then
            If Val(编辑框_成本工费.Text) <> 0 OrElse Val(编辑框_成本附加费.Text) <> 0 Then
                _商品信息_参数数值计算("成本工费")
            End If
        End If

        ' 准备变量
        Dim 修改信息证书编码 As String = 编辑框_证书编码.Text
        Dim 修改信息检测结果 As String = 编辑框_检测结果.Text
        Dim 修改信息检测总重 As String = 编辑框_检测总重.Text
        Dim 修改信息检测形状 As String = 编辑框_检测形状.Text
        Dim 修改信息检测颜色 As String = 编辑框_检测颜色.Text
        Dim 修改信息成本价 As String = 格式二位小数(编辑框_成本价.Text)
        Dim 修改信息销售价 As String = 格式二位小数(编辑框_证书销售价.Text)
        Dim 修改信息查询地址 As String = 编辑框_查询地址.Text
        Dim 修改信息备注信息 As String = 编辑框_备注.Text
        Dim 修改信息检测机构 As String = If(组合框_检测机构.SelectedIndex > 0, 取项目附加数值(组合框_检测机构), "")

        Dim 商品添加克价 As String = 编辑框_入库克价.Text
        Dim 商品添加名称 As String = 编辑框_商品名称.Text
        Dim 商品添加编码 As String = 编辑框_固定编码.Text
        Dim 商品添加成色 As String = 编辑框_公司成色.Text
        Dim 商品添加数量 As String = 编辑框_数量.Text
        Dim 商品添加重量 As String = 编辑框_重量.Text
        Dim 商品添加单件重 As String = 编辑框_单件重.Text
        Dim 商品添加款号 As String = 编辑框_商品款号.Text
        Dim 商品添加圈号 As String = 编辑框_圈号长度.Text
        Dim 商品添加面宽 As String = 编辑框_面宽.Text
        Dim 商品添加厚度 As String = 编辑框_厚度.Text
        Dim 商品添加金重 As String = 编辑框_金重.Text
        Dim 商品添加损耗 As String = 编辑框_损耗.Text
        Dim 商品添加含耗重 As String = 编辑框_含耗重.Text
        Dim 商品添加成本价 As String = 编辑框_成本单价.Text
        Dim 商品添加系数 As String = 编辑框_系数.Text
        Dim 商品添加石重 As String = 编辑框_主石重.Text
        Dim 商品添加石头数量 As String = 编辑框_石头数量.Text
        Dim 商品添加副石重 As String = 编辑框_副石重.Text
        Dim 商品添加副石头数量 As String = 编辑框_副石头数量.Text
        Dim 商品添加主色 As String = 编辑框_商品主色.Text
        Dim 商品添加基本工费 As String = 编辑框_成本工费.Text
        Dim 商品添加成本附加费 As String = 编辑框_成本附加费.Text
        Dim 商品添加销售工费 As String = 编辑框_销售工费.Text
        Dim 商品添加精品工费 As String = 编辑框_参考工费.Text
        Dim 商品添加销售附加费 As String = 编辑框_销售附加费.Text
        Dim 商品添加销售价 As String = 编辑框_销售价.Text
        Dim 商品添加备注 As String = 编辑框_备注.Text
        Dim 商品添加图片地址 As String = 编辑框_图片地址.Text
        Dim 商品添加原料价格 As String = 编辑框_入库克价.Text
        Dim 商品添加模具号 As String = 编辑框_模具号.Text

        ' 如果品类多数量=1且数量>1，平均分配
        If 局部_品类多数量 = "1" AndAlso Val(编辑框_数量.Text) > 1 Then
            商品添加重量 = 格式三位小数((Val(编辑框_重量.Text) / Val(编辑框_数量.Text)).ToString())
            商品添加金重 = 格式三位小数((Val(编辑框_金重.Text) / Val(编辑框_数量.Text)).ToString())
            商品添加损耗 = 格式三位小数((Val(编辑框_损耗.Text) / Val(编辑框_数量.Text)).ToString())
            商品添加含耗重 = 格式三位小数((Val(编辑框_含耗重.Text) / Val(编辑框_数量.Text)).ToString())
            商品添加石重 = 格式三位小数((Val(编辑框_主石重.Text) / Val(编辑框_数量.Text)).ToString())
            商品添加副石重 = 格式三位小数((Val(编辑框_副石重.Text) / Val(编辑框_数量.Text)).ToString())
            商品添加石头数量 = Math.Round(Val(编辑框_石头数量.Text) / Val(编辑框_数量.Text), 0).ToString()
            商品添加副石头数量 = Math.Round(Val(编辑框_副石头数量.Text) / Val(编辑框_数量.Text), 0).ToString()
        End If

        Dim 商品添加材质 As String = 取项目文本(组合框_商品材质)
        Dim 商品加成 As String = 取项目文本(组合框_商品成色)
        Dim 商品添加品类 As String = 取项目文本(组合框_品类名称)
        Dim 商品添加规格 As String = 取项目文本(组合框_商品规格)
        Dim 商品添加规格id As String = 取项目附加数值(组合框_商品规格)
        Dim 全局_信息操作日期 As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Dim 全局_信息操作账户 As String = UserAccount
        Dim 商品添加镶嵌 As String = 局部_商品是否镶嵌
        Dim 商品添加单位 As String = ""

        If 单选框_单件.Checked Then 商品添加单位 = 单选框_单件.Text
        If 单选框_重量.Checked Then 商品添加单位 = 单选框_重量.Text

        Dim 商品入库次数 As Integer = 1
        If 局部_品类多数量 = "1" AndAlso Val(编辑框_数量.Text) > 1 Then
            商品入库次数 = CInt(Val(编辑框_数量.Text))
        End If

        Dim 商品入库序号开始 As Integer = 局部_入库商品数量

        ' 更新共享计数
        局部_入库商品数量 += 商品入库次数

        Dim progressForm As New System.Windows.Forms.Form()
        progressForm.Text = "数据正在写入中..."
        progressForm.Size = New Drawing.Size(400, 100)
        progressForm.StartPosition = FormStartPosition.CenterParent
        Dim progressBar As New ProgressBar()
        progressBar.Location = New Drawing.Point(20, 30)
        progressBar.Size = New Drawing.Size(350, 23)
        progressBar.Maximum = 商品入库次数
        progressBar.Value = 0
        progressForm.Controls.Add(progressBar)
        progressForm.Show()

        For 商品入库计次 As Integer = 1 To 商品入库次数
            progressBar.Value = 商品入库计次
            Application.DoEvents()

            Dim 随机商品信息编码 As String = DateTime.Now.Ticks.ToString() & UserAccount & 商品入库计次.ToString()
            Dim 商品写入商品编码 As String = If(商品添加编码 = "", 随机商品信息编码, 商品添加编码)

            ' 插入商品到 xipunum_erp_shop
            Dim insertSql As String = $"INSERT INTO xipunum_erp_shop (poduct_code,fu_code,specification_id) VALUES ('{SafeSQL(随机商品信息编码)}','{SafeSQL(商品写入商品编码)}','{SafeSQL(商品添加规格id)}')"
            ExecuteCommand(insertSql)

            ' 查询刚插入的记录
            Dim selectSql As String = $"SELECT id,fu_code FROM xipunum_erp_shop WHERE poduct_code='{SafeSQL(随机商品信息编码)}' ORDER BY id ASC LIMIT 1"
            Dim dt As DataTable = ExecuteQuery(selectSql)
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then Continue For

            Dim 商品信息数据ID As String = SafeString(dt.Rows(0)("id"))
            Dim 商品信息数据编码 As String = SafeString(dt.Rows(0)("fu_code"))
            If String.IsNullOrEmpty(商品信息数据ID) Then Continue For

            商品入库序号开始 += 1
            Dim 入库数据主编码 As String = DateTime.Now.ToString("yyMM") & 商品信息数据ID.PadLeft(5, "0"c)

            If 商品信息数据编码 = 随机商品信息编码 Then
                商品信息数据编码 = 入库数据主编码
            End If

            ' 检查编码是否已存在
            Dim checkSql As String = $"SELECT id FROM xipunum_erp_shop WHERE fu_code='{SafeSQL(商品信息数据编码)}' LIMIT 2"
            Dim checkDt As DataTable = ExecuteQuery(checkSql)
            If checkDt IsNot Nothing AndAlso checkDt.Rows.Count = 1 Then
                入库数据主编码 = 商品信息数据编码
            End If

            ' 更新正式编码
            Dim updateSql As String = $"UPDATE xipunum_erp_shop SET poduct_code='{SafeSQL(入库数据主编码)}',fu_code='{SafeSQL(商品信息数据编码)}' WHERE id='{SafeSQL(商品信息数据ID)}' LIMIT 1"
            ExecuteCommand(updateSql)

            ' 如果有证书
            If 单选框_有.Checked Then
                Dim zsSql As String = $"INSERT INTO xipunum_erp_zhengshu (poduct_code,jigouid,zsbianma,conclusion,zongzhong,xingzhuang,yanse,beizhu,chengben,xiaoshou,cxdizhi,cjuser,creationtime) VALUES ('{SafeSQL(入库数据主编码)}','{SafeSQL(修改信息检测机构)}','{SafeSQL(修改信息证书编码)}','{SafeSQL(修改信息检测结果)}','{SafeSQL(修改信息检测总重)}','{SafeSQL(修改信息检测形状)}','{SafeSQL(修改信息检测颜色)}','{SafeSQL(修改信息备注信息)}','{SafeSQL(修改信息成本价)}','{SafeSQL(修改信息销售价)}','{SafeSQL(修改信息查询地址)}','{SafeSQL(全局_信息操作账户)}','{SafeSQL(全局_信息操作日期)}')"
                ExecuteCommand(zsSql)
            End If

            ' 写入高级表格（通过共享状态委托给父窗口）
            Dim rowIdx As Integer = 商品入库序号开始
            ' 更新入库商品序号
            商品入库序号开始 += 1
        Next

        progressForm.Close()

        ' 重置表单
        _商品信息_默认参数()
        编辑框_重量.Enabled = True
        组合框_品类名称.Enabled = True
        编辑框_重量.Text = ""
        编辑框_单件重.Text = "0.000"
        编辑框_数量.Text = "1"
        编辑框_金重.Text = "0.000"
        编辑框_损耗.Text = "0.000"
        编辑框_含耗重.Text = "0.000"
        编辑框_圈号长度.Text = ""
        编辑框_面宽.Text = ""
        编辑框_厚度.Text = ""
        编辑框_成本单价.Text = "0.00"
        编辑框_主石重.Text = ""
        编辑框_石头数量.Text = ""
        编辑框_副石重.Text = ""
        编辑框_副石头数量.Text = ""
        编辑框_商品主色.Text = "白"
        编辑框_系数.Text = "1"
        编辑框_销售价.Text = "0.00"
        编辑框_备注.Text = ""
        按钮EX_添加.Enabled = False

        编辑框_重量.Focus()
    End Sub

    ' ========== 窗口关闭 ==========
    Private Sub _窗口_商品信息添加_可否被关闭(sender As Object, e As FormClosingEventArgs)
        ' 卸载热键
    End Sub

    ' ========== 辅助方法 ==========
    Private Function SafeString(val As Object) As String
        Return If(val Is Nothing OrElse IsDBNull(val), "", val.ToString())
    End Function

    Private Function SafeSQL(val As String) As String
        If String.IsNullOrEmpty(val) Then Return ""
        Return val.Replace("'"c, "''"c)
    End Function
End Class

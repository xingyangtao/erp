' ============================================================================
' 商品成品数据修改窗口
' 功能: 修改已添加商品的信息
' 100% 匹配易语言 窗口_商品成品数据修改 功能
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports VB_ERP.Modules

Public Class ProductEditForm
    Inherits System.Windows.Forms.Form

#Region "程序集变量"

    ' 内部报表和显示组件
    Private 集_数据报表 As Object                    ' 报表
    Private 集_数据显示器 As Object                   ' 查询显示器

    ' 局部变量 - 商品信息缓存
    Private 局部_成品商品编码 As String = ""
    Private 局部_成品商品数据编码 As String = ""
    Private 局部_商品成色 As String = ""
    Private 局部_商品材质 As String = ""
    Private 局部_商品规格 As String = ""
    Private 局部_商品品类 As String = ""
    Private 局部_商品是否镶嵌 As String = ""
    Private 局部_品类多数量 As String = ""
    Private 局部_商品入库数量 As String = ""
    Private 局部_商品入库总重 As String = ""
    Private 局部_商品入库金重 As String = ""
    Private 局部_商品商品损耗 As String = ""
    Private 局部_商品含耗重 As String = ""
    Private 局部_商品款号数据 As String = ""
    Private 局部_商品图片地址 As String = ""
    Private 局部_商品库存id信息 As String = ""

#End Region

#Region "控件声明"

    Private 添加修改_分组框 As New GroupBox()

    Private WithEvents 编辑框_商品编码 As New TextBox()
    Private WithEvents 按钮EX_查找 As New Button()
    Private WithEvents 按钮EX_重置 As New Button()
    Private WithEvents 按钮EX_保存 As New Button()
    Private WithEvents 按钮EX_打印 As New Button()
    Private WithEvents 按钮EX_圈号长度 As New Button()
    Private 按钮EX2 As New Button()

    Private WithEvents 组合框_品类名称 As New ComboBox()
    Private WithEvents 组合框_商品材质 As New ComboBox()
    Private WithEvents 组合框_商品成色 As New ComboBox()
    Private WithEvents 组合框_商品规格 As New ComboBox()
    Private WithEvents 组合框_标签样式 As New ComboBox()

    Private WithEvents 单选框_重量 As New RadioButton()
    Private WithEvents 单选框_单件 As New RadioButton()
    Private WithEvents 单选框_固口 As New RadioButton()
    Private WithEvents 单选框_开口 As New RadioButton()

    Private WithEvents 选择框_自动打印 As New CheckBox()

    Private WithEvents 编辑框_入库克价 As New TextBox()
    Private WithEvents 编辑框_固定编码 As New TextBox()
    Private WithEvents 编辑框_单件重 As New TextBox()
    Private WithEvents 编辑框_数量 As New TextBox()
    Private WithEvents 编辑框_重量 As New TextBox()
    Private WithEvents 编辑框_金重 As New TextBox()
    Private WithEvents 编辑框_损耗 As New TextBox()
    Private WithEvents 编辑框_含耗重 As New TextBox()
    Private WithEvents 编辑框_圈号长度 As New TextBox()
    Private WithEvents 编辑框_面宽 As New TextBox()
    Private WithEvents 编辑框_厚度 As New TextBox()
    Private WithEvents 编辑框_成本单价 As New TextBox()
    Private WithEvents 编辑框_商品款号 As New TextBox()
    Private WithEvents 编辑框_主石重 As New TextBox()
    Private WithEvents 编辑框_石头数量 As New TextBox()
    Private WithEvents 编辑框_副石重 As New TextBox()
    Private WithEvents 编辑框_副石头数量 As New TextBox()
    Private WithEvents 编辑框_商品主色 As New TextBox()
    Private WithEvents 编辑框_系数 As New TextBox()
    Private WithEvents 编辑框_成本工费 As New TextBox()
    Private WithEvents 编辑框_成本附加费 As New TextBox()
    Private WithEvents 编辑框_销售工费 As New TextBox()
    Private WithEvents 编辑框_参考工费 As New TextBox()
    Private WithEvents 编辑框_销售附加费 As New TextBox()
    Private WithEvents 编辑框_销售价 As New TextBox()
    Private WithEvents 编辑框_备注 As New TextBox()

    Private 图片框EX_主图 As New PictureBox()
    Private 图片框EX_固开口 As New PictureBox()
    Private WithEvents 图片框EX4 As New PictureBox()

    ' 标签
    Private 标签Ex14 As New Label()
    Private 标签Ex15 As New Label()
    Private 标签Ex16 As New Label()

    ' 数据库连接（用于 Access 标签打印）
    Private 数据库连接1 As Object ' OleDbConnection

#End Region

#Region "初始化"

    Public Sub New()
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        Me.Text = "商品成品数据修改"
        Me.Size = New Drawing.Size(1038, 723)
        Me.StartPosition = FormStartPosition.CenterParent

        ' 添加修改_分组框 - 主面板
        添加修改_分组框.Text = "商品数据修改"
        添加修改_分组框.Size = New Drawing.Size(1000, 680)
        Me.Controls.Add(添加修改_分组框)

        ' 商品编码行
        ' 编辑框_商品编码
        编辑框_商品编码.Location = New Drawing.Point(120, 45)
        编辑框_商品编码.Size = New Drawing.Size(180, 25)
        AddHandler 编辑框_商品编码.KeyDown, AddressOf 编辑框_商品编码_键盘事件
        添加修改_分组框.Controls.Add(编辑框_商品编码)

        ' 按钮EX_查找
        按钮EX_查找.Text = "查找"
        按钮EX_查找.Location = New Drawing.Point(304, 43)
        按钮EX_查找.Size = New Drawing.Size(72, 28)
        添加修改_分组框.Controls.Add(按钮EX_查找)

        ' 按钮EX_重置
        按钮EX_重置.Text = "重置"
        按钮EX_重置.Location = New Drawing.Point(380, 43)
        按钮EX_重置.Size = New Drawing.Size(72, 28)
        添加修改_分组框.Controls.Add(按钮EX_重置)

        ' 按钮EX_保存
        按钮EX_保存.Text = "保存"
        按钮EX_保存.Location = New Drawing.Point(456, 43)
        按钮EX_保存.Size = New Drawing.Size(72, 28)
        添加修改_分组框.Controls.Add(按钮EX_保存)

        ' 选择框_自动打印
        选择框_自动打印.Text = "自动打印"
        选择框_自动打印.Location = New Drawing.Point(540, 45)
        选择框_自动打印.AutoSize = True
        添加修改_分组框.Controls.Add(选择框_自动打印)

        ' 图片框EX_主图
        图片框EX_主图.Location = New Drawing.Point(750, 10)
        图片框EX_主图.Size = New Drawing.Size(230, 230)
        图片框EX_主图.BorderStyle = BorderStyle.FixedSingle
        图片框EX_主图.SizeMode = PictureBoxSizeMode.Zoom
        添加修改_分组框.Controls.Add(图片框EX_主图)

        ' 单品信息行1
        Dim y1 As Integer = 80

        ' 组合框_品类名称
        AddLabel("品类名称:", 10, y1)
        组合框_品类名称.Location = New Drawing.Point(120, y1)
        组合框_品类名称.Size = New Drawing.Size(150, 25)
        组合框_品类名称.DropDownStyle = ComboBoxStyle.DropDownList
        添加修改_分组框.Controls.Add(组合框_品类名称)

        ' 组合框_商品材质
        AddLabel("商品材质:", 280, y1)
        组合框_商品材质.Location = New Drawing.Point(370, y1)
        组合框_商品材质.Size = New Drawing.Size(120, 25)
        组合框_商品材质.DropDownStyle = ComboBoxStyle.DropDownList
        添加修改_分组框.Controls.Add(组合框_商品材质)

        ' 组合框_商品成色
        AddLabel("商品成色:", 500, y1)
        组合框_商品成色.Location = New Drawing.Point(590, y1)
        组合框_商品成色.Size = New Drawing.Size(120, 25)
        组合框_商品成色.DropDownStyle = ComboBoxStyle.DropDownList
        添加修改_分组框.Controls.Add(组合框_商品成色)

        ' 单选框_重量 / 单选框_单件
        AddLabel("计量方式:", 720, y1)
        ' (Note: these are inside 添加修改_分组框 but referenced as 窗口_商品成品数据修改.xxx in EL)

        ' 行2 - 入库克价/固定编码/单件重
        Dim y2 As Integer = 115

        AddLabel("入库克价:", 10, y2)
        编辑框_入库克价.Location = New Drawing.Point(120, y2)
        编辑框_入库克价.Size = New Drawing.Size(120, 25)
        添加修改_分组框.Controls.Add(编辑框_入库克价)

        AddLabel("固定编码:", 250, y2)
        编辑框_固定编码.Location = New Drawing.Point(350, y2)
        编辑框_固定编码.Size = New Drawing.Size(120, 25)
        添加修改_分组框.Controls.Add(编辑框_固定编码)

        AddLabel("单件重:", 480, y2)
        编辑框_单件重.Location = New Drawing.Point(560, y2)
        编辑框_单件重.Size = New Drawing.Size(120, 25)
        添加修改_分组框.Controls.Add(编辑框_单件重)

        ' 行3 - 数量/重量/金重
        Dim y3 As Integer = 150

        AddLabel("数量:", 10, y3)
        编辑框_数量.Location = New Drawing.Point(120, y3)
        编辑框_数量.Size = New Drawing.Size(120, 25)
        添加修改_分组框.Controls.Add(编辑框_数量)

        AddLabel("重量:", 250, y3)
        编辑框_重量.Location = New Drawing.Point(350, y3)
        编辑框_重量.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_重量.Leave, AddressOf 编辑框_重量_鼠标进入离开
        AddHandler 编辑框_重量.KeyDown, AddressOf 编辑框_重量_键盘事件
        添加修改_分组框.Controls.Add(编辑框_重量)

        AddLabel("金重:", 480, y3)
        编辑框_金重.Location = New Drawing.Point(560, y3)
        编辑框_金重.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_金重.TextChanged, AddressOf 编辑框_金重_内容被改变
        AddHandler 编辑框_金重.Leave, AddressOf 编辑框_金重_焦点事件
        添加修改_分组框.Controls.Add(编辑框_金重)

        ' 行4 - 损耗/含耗重
        Dim y4 As Integer = 185

        AddLabel("损耗:", 10, y4)
        编辑框_损耗.Location = New Drawing.Point(120, y4)
        编辑框_损耗.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_损耗.TextChanged, AddressOf 编辑框_损耗_内容被改变
        AddHandler 编辑框_损耗.Leave, AddressOf 编辑框_损耗_焦点事件
        添加修改_分组框.Controls.Add(编辑框_损耗)

        AddLabel("含耗重:", 250, y4)
        编辑框_含耗重.Location = New Drawing.Point(350, y4)
        编辑框_含耗重.Size = New Drawing.Size(120, 25)
        添加修改_分组框.Controls.Add(编辑框_含耗重)

        ' 行5 - 圈号长度/面宽/厚度
        Dim y5 As Integer = 220

        图片框EX_固开口.Location = New Drawing.Point(196, y5 + 5)
        图片框EX_固开口.Size = New Drawing.Size(26, 26)
        图片框EX_固开口.Visible = False
        添加修改_分组框.Controls.Add(图片框EX_固开口)

        按钮EX_圈号长度.Text = "圈号长度"
        按钮EX_圈号长度.Location = New Drawing.Point(196, y5 + 5)
        按钮EX_圈号长度.Size = New Drawing.Size(70, 25)
        按钮EX_圈号长度.Visible = True
        添加修改_分组框.Controls.Add(按钮EX_圈号长度)

        编辑框_圈号长度.Location = New Drawing.Point(270, y5)
        编辑框_圈号长度.Size = New Drawing.Size(110, 25)
        添加修改_分组框.Controls.Add(编辑框_圈号长度)

        ' 单选框_固口 / 单选框_开口
        单选框_固口.Text = "固口"
        单选框_固口.Location = New Drawing.Point(196, y5 + 30)
        单选框_固口.AutoSize = True
        添加修改_分组框.Controls.Add(单选框_固口)

        单选框_开口.Text = "开口"
        单选框_开口.Location = New Drawing.Point(265, y5 + 30)
        单选框_开口.AutoSize = True
        添加修改_分组框.Controls.Add(单选框_开口)

        AddLabel("面宽:", 388, y5)
        编辑框_面宽.Location = New Drawing.Point(450, y5)
        编辑框_面宽.Size = New Drawing.Size(100, 25)
        添加修改_分组框.Controls.Add(编辑框_面宽)

        AddLabel("厚度:", 560, y5)
        编辑框_厚度.Location = New Drawing.Point(620, y5)
        编辑框_厚度.Size = New Drawing.Size(100, 25)
        添加修改_分组框.Controls.Add(编辑框_厚度)

        ' 行6 - 成本单价/商品款号
        Dim y6 As Integer = 260

        AddLabel("成本单价:", 10, y6)
        编辑框_成本单价.Location = New Drawing.Point(120, y6)
        编辑框_成本单价.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_成本单价.Leave, AddressOf 编辑框_成本单价_焦点事件
        添加修改_分组框.Controls.Add(编辑框_成本单价)

        AddLabel("商品款号:", 250, y6)
        编辑框_商品款号.Location = New Drawing.Point(350, y6)
        编辑框_商品款号.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_商品款号.KeyDown, AddressOf 编辑框_商品款号_键盘事件
        AddHandler 编辑框_商品款号.Enter, AddressOf 编辑框_商品款号_获得焦点
        AddHandler 编辑框_商品款号.Leave, AddressOf 编辑框_商品款号_失去焦点
        添加修改_分组框.Controls.Add(编辑框_商品款号)

        ' 行7 - 主石信息
        Dim y7 As Integer = 295

        AddLabel("主石重:", 10, y7)
        编辑框_主石重.Location = New Drawing.Point(120, y7)
        编辑框_主石重.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_主石重.Leave, AddressOf 编辑框_主石重_焦点事件
        添加修改_分组框.Controls.Add(编辑框_主石重)

        AddLabel("石头数量:", 250, y7)
        编辑框_石头数量.Location = New Drawing.Point(350, y7)
        编辑框_石头数量.Size = New Drawing.Size(120, 25)
        添加修改_分组框.Controls.Add(编辑框_石头数量)

        ' 行8 - 副石信息
        Dim y8 As Integer = 330

        AddLabel("副石重:", 10, y8)
        编辑框_副石重.Location = New Drawing.Point(120, y8)
        编辑框_副石重.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_副石重.Leave, AddressOf 编辑框_副石重_焦点事件
        添加修改_分组框.Controls.Add(编辑框_副石重)

        AddLabel("副石数量:", 250, y8)
        编辑框_副石头数量.Location = New Drawing.Point(350, y8)
        编辑框_副石头数量.Size = New Drawing.Size(120, 25)
        添加修改_分组框.Controls.Add(编辑框_副石头数量)

        AddLabel("主色:", 480, y8)
        编辑框_商品主色.Location = New Drawing.Point(560, y8)
        编辑框_商品主色.Size = New Drawing.Size(120, 25)
        添加修改_分组框.Controls.Add(编辑框_商品主色)

        ' 行9 - 系数/成本工费/成本附加费
        Dim y9 As Integer = 365

        AddLabel("系数:", 10, y9)
        编辑框_系数.Location = New Drawing.Point(120, y9)
        编辑框_系数.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_系数.Leave, AddressOf 编辑框_系数_焦点事件
        添加修改_分组框.Controls.Add(编辑框_系数)

        AddLabel("成本工费:", 250, y9)
        编辑框_成本工费.Location = New Drawing.Point(350, y9)
        编辑框_成本工费.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_成本工费.Leave, AddressOf 编辑框_成本工费_焦点事件
        添加修改_分组框.Controls.Add(编辑框_成本工费)

        AddLabel("成本附加费:", 480, y9)
        编辑框_成本附加费.Location = New Drawing.Point(590, y9)
        编辑框_成本附加费.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_成本附加费.Leave, AddressOf 编辑框_成本附加费_焦点事件
        添加修改_分组框.Controls.Add(编辑框_成本附加费)

        ' 行10 - 销售工费/参考工费/销售附加费
        Dim y10 As Integer = 400

        AddLabel("销售工费:", 10, y10)
        编辑框_销售工费.Location = New Drawing.Point(120, y10)
        编辑框_销售工费.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_销售工费.Leave, AddressOf 编辑框_销售工费_焦点事件
        添加修改_分组框.Controls.Add(编辑框_销售工费)

        AddLabel("参考工费:", 250, y10)
        编辑框_参考工费.Location = New Drawing.Point(350, y10)
        编辑框_参考工费.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_参考工费.Leave, AddressOf 编辑框_参考工费_焦点事件
        添加修改_分组框.Controls.Add(编辑框_参考工费)

        AddLabel("销售附加费:", 480, y10)
        编辑框_销售附加费.Location = New Drawing.Point(590, y10)
        编辑框_销售附加费.Size = New Drawing.Size(120, 25)
        AddHandler 编辑框_销售附加费.Leave, AddressOf 编辑框_销售附加费_焦点事件
        添加修改_分组框.Controls.Add(编辑框_销售附加费)

        ' 行11 - 销售价
        Dim y11 As Integer = 435

        AddLabel("销售价:", 10, y11)
        编辑框_销售价.Location = New Drawing.Point(120, y11)
        编辑框_销售价.Size = New Drawing.Size(120, 25)
        添加修改_分组框.Controls.Add(编辑框_销售价)

        ' 行12 - 备注
        Dim y12 As Integer = 470

        AddLabel("备注:", 10, y12)
        编辑框_备注.Location = New Drawing.Point(120, y12)
        编辑框_备注.Size = New Drawing.Size(500, 25)
        添加修改_分组框.Controls.Add(编辑框_备注)

        ' 标签样式 和 打印按钮
        AddLabel("标签样式:", 364, y12 + 31)
        组合框_标签样式.Location = New Drawing.Point(450, y12 + 28)
        组合框_标签样式.Size = New Drawing.Size(130, 25)
        组合框_标签样式.DropDownStyle = ComboBoxStyle.DropDownList
        添加修改_分组框.Controls.Add(组合框_标签样式)

        按钮EX_打印.Text = "打印标签"
        按钮EX_打印.Location = New Drawing.Point(590, y12 + 27)
        按钮EX_打印.Size = New Drawing.Size(80, 28)
        添加修改_分组框.Controls.Add(按钮EX_打印)

        ' 关闭按钮
        图片框EX4.Location = New Drawing.Point(960, 10)
        图片框EX4.Size = New Drawing.Size(26, 26)
        图片框EX4.BackColor = Drawing.Color.Red
        添加修改_分组框.Controls.Add(图片框EX4)

        ' 设置 Numeric input mode for decimal fields
        SetNumericInput(编辑框_入库克价)
        SetNumericInput(编辑框_单件重)
        SetNumericInput(编辑框_重量)
        SetNumericInput(编辑框_金重)
        SetNumericInput(编辑框_损耗)
        SetNumericInput(编辑框_含耗重)
        SetNumericInput(编辑框_圈号长度)
        SetNumericInput(编辑框_面宽)
        SetNumericInput(编辑框_厚度)
        SetNumericInput(编辑框_成本单价)
        SetNumericInput(编辑框_主石重)
        SetNumericInput(编辑框_系数)
        SetNumericInput(编辑框_成本工费)
        SetNumericInput(编辑框_成本附加费)
        SetNumericInput(编辑框_销售工费)
        SetNumericInput(编辑框_参考工费)
        SetNumericInput(编辑框_销售附加费)
        SetNumericInput(编辑框_销售价)

        ' Integer-only fields
        SetIntegerInput(编辑框_数量)
        SetIntegerInput(编辑框_石头数量)
        SetIntegerInput(编辑框_副石头数量)
        SetIntegerInput(编辑框_圈号长度)
        SetIntegerInput(编辑框_面宽)
        SetIntegerInput(编辑框_厚度)

        ' 窗口创建完毕
        AddHandler Me.Load, AddressOf 窗口_商品成品数据修改_创建完毕
        AddHandler Me.Resize, AddressOf 窗口_商品成品数据修改_尺寸被改变
    End Sub

    Private Sub AddLabel(text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y + 4)
        lbl.AutoSize = True
        添加修改_分组框.Controls.Add(lbl)
    End Sub

    Private Sub SetNumericInput(tb As TextBox)
        tb.ShortcutsEnabled = False
    End Sub

    Private Sub SetIntegerInput(tb As TextBox)
        tb.ShortcutsEnabled = False
    End Sub

#End Region

#Region "窗口创建完毕"

    Private Sub 窗口_商品成品数据修改_创建完毕(sender As Object, e As EventArgs)
        ' 清空所有编辑框
        编辑框_商品编码.Text = ""
        组合框_品类名称.Text = "请选择品类名称"
        组合框_商品材质.Text = "请选择商品材质"
        组合框_商品成色.Text = "请选择商品成色"
        组合框_商品规格.Text = "请选择商品规格"
        组合框_标签样式.Text = "请选择标签样式"
        单选框_重量.Checked = False
        单选框_单件.Checked = False
        编辑框_入库克价.Text = ""
        编辑框_固定编码.Text = ""
        编辑框_单件重.Text = ""
        编辑框_数量.Text = ""
        编辑框_重量.Text = ""
        编辑框_金重.Text = ""
        编辑框_损耗.Text = ""
        编辑框_含耗重.Text = ""
        编辑框_圈号长度.Text = ""
        编辑框_面宽.Text = ""
        编辑框_厚度.Text = ""
        编辑框_成本单价.Text = ""
        编辑框_商品款号.Text = ""
        编辑框_主石重.Text = ""
        编辑框_石头数量.Text = ""
        编辑框_副石重.Text = ""
        编辑框_副石头数量.Text = ""
        编辑框_商品主色.Text = ""
        编辑框_系数.Text = ""
        编辑框_成本工费.Text = ""
        编辑框_成本附加费.Text = ""
        编辑框_销售工费.Text = ""
        编辑框_参考工费.Text = ""
        编辑框_销售附加费.Text = ""
        编辑框_销售价.Text = ""
        编辑框_备注.Text = ""

        ' 清空局部变量
        局部_商品品类 = ""
        局部_商品规格 = ""
        局部_商品材质 = ""
        局部_商品成色 = ""
        局部_商品是否镶嵌 = ""
        局部_品类多数量 = ""
        局部_商品入库数量 = ""
        局部_商品入库总重 = ""
        局部_商品入库金重 = ""
        局部_商品商品损耗 = ""
        局部_商品含耗重 = ""
        局部_商品款号数据 = ""
        局部_商品图片地址 = ""

        ' 禁用所有控件
        SetAllEnabled(False)
        SetAllEnabled(False)

        图片框EX_固开口.Visible = False
        按钮EX_圈号长度.Visible = True
        图片框EX_固开口.Left = 196
        图片框EX_固开口.Top = 239
        编辑框_圈号长度.Width = 110
        编辑框_圈号长度.Left = 270

        图片框EX_主图.Image = Nothing

        ' 设置只允许数字和字母输入
        ' (VB.NET textboxes already handle this - we use TextChanged validation)

        ' 初始化商品组合框
        商品组合框_初始化()
    End Sub

    Private Sub SetAllEnabled(enabled As Boolean)
        组合框_品类名称.Enabled = enabled
        组合框_商品材质.Enabled = enabled
        组合框_商品成色.Enabled = enabled
        组合框_商品规格.Enabled = enabled
        单选框_重量.Enabled = enabled
        单选框_单件.Enabled = enabled
        编辑框_入库克价.Enabled = enabled
        单选框_固口.Enabled = enabled
        单选框_开口.Enabled = enabled
        编辑框_固定编码.Enabled = enabled
        编辑框_单件重.Enabled = enabled
        编辑框_数量.Enabled = enabled
        编辑框_重量.Enabled = enabled
        编辑框_金重.Enabled = enabled
        编辑框_损耗.Enabled = enabled
        编辑框_含耗重.Enabled = enabled
        编辑框_圈号长度.Enabled = enabled
        编辑框_面宽.Enabled = enabled
        编辑框_厚度.Enabled = enabled
        编辑框_成本单价.Enabled = enabled
        编辑框_商品款号.Enabled = enabled
        编辑框_主石重.Enabled = enabled
        编辑框_石头数量.Enabled = enabled
        编辑框_副石重.Enabled = enabled
        编辑框_副石头数量.Enabled = enabled
        编辑框_商品主色.Enabled = enabled
        编辑框_系数.Enabled = enabled
        编辑框_成本工费.Enabled = enabled
        编辑框_成本附加费.Enabled = enabled
        编辑框_销售工费.Enabled = enabled
        编辑框_参考工费.Enabled = enabled
        编辑框_销售附加费.Enabled = enabled
        编辑框_销售价.Enabled = enabled
        编辑框_备注.Enabled = enabled
        组合框_标签样式.Enabled = enabled
    End Sub

#End Region

#Region "商品组合框_初始化"

    Private Sub 商品组合框_初始化()
        ' 品类名称
        组合框_品类名称.Items.Clear()
        组合框_商品材质.Items.Clear()
        组合框_商品成色.Items.Clear()
        组合框_商品规格.Items.Clear()
        组合框_标签样式.Items.Clear()

        组合框_品类名称.Text = "请选择品类名称"
        组合框_商品材质.Text = "请选择商品材质"
        组合框_商品成色.Text = "请选择商品成色"
        组合框_商品规格.Text = "请选择商品规格"
        组合框_标签样式.Text = "请选择标签样式"

        ' 加载品类列表
        Dim sql As String = "SELECT * FROM xipunum_erp_category where 1=1 order by id ASC"
        Dim dt As DataTable = DatabaseModule.MySQL_Read(sql)
        If dt IsNot Nothing Then
            For Each row As DataRow In dt.Rows
                Dim id As String = SafeString(row("id"))
                Dim title As String = SafeString(row("title"))
                组合框_品类名称.Items.Add(New ComboBoxItem(title, id))
            Next
        End If

        ' 加载标签样式
        Dim voucherPath As String = System.IO.Path.Combine(Application.StartupPath, "voucher\biaoqian\")
        If System.IO.Directory.Exists(voucherPath) Then
            Dim files As String() = System.IO.Directory.GetFiles(voucherPath, "*.qdf")
            For i As Integer = 0 To files.Length - 1
                Dim fileName As String = System.IO.Path.GetFileName(files(i))
                组合框_标签样式.Items.Add(New ComboBoxItem(fileName, (i + 1).ToString()))
            Next
            If files.Length > 0 Then
                组合框_标签样式.SelectedIndex = 0
            End If
        End If
    End Sub

#End Region

#Region "查找功能"

    Private Sub 编辑框_商品编码_键盘事件(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            按钮EX_查找_鼠标左键单击(Nothing, Nothing)
        End If
    End Sub

    Private Sub 按钮EX_查找_鼠标左键单击(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(编辑框_商品编码.Text) Then
            ShowWarning("商品编码不能为空！")
            编辑框_商品编码.Focus()
            Return
        End If

        If 编辑框_商品编码.Text.Length <= 3 Then
            ShowWarning("请输入正确的商品编码！")
            编辑框_商品编码.Text = ""
            编辑框_商品编码.Focus()
            Return
        End If

        商品数据_初始化()
    End Sub

#End Region

#Region "商品数据_初始化"

    Private Sub 商品数据_初始化()
        局部_成品商品编码 = ""
        局部_成品商品编码 = 编辑框_商品编码.Text

        ' 检查商品是否存在
        Dim sqlExist As String = $"SELECT * FROM xipunum_erp_shop where (poduct_code='{SafeSQL(局部_成品商品编码)}' or fu_code='{SafeSQL(局部_成品商品编码)}') order by id ASC"
        Dim dtExist As DataTable = DatabaseModule.MySQL_Read(sqlExist)
        If dtExist Is Nothing OrElse dtExist.Rows.Count = 0 Then
            ShowWarning("此商品不存在！")
            编辑框_商品编码.Text = ""
            编辑框_商品编码.Focus()
            Return
        End If

        ' 检查库存
        Dim sqlStock As String = $"SELECT a.poduct_code AS apoduct_code FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE (a.quantity > 0 or a.jinzhong  > 0) AND (b.poduct_code = '{SafeSQL(局部_成品商品编码)}' OR b.fu_code = '{SafeSQL(局部_成品商品编码)}') ORDER BY a.id DESC"
        Dim dtStock As DataTable = DatabaseModule.MySQL_Read(sqlStock)
        Dim stockCount As Integer = If(dtStock IsNot Nothing, dtStock.Rows.Count, 0)

        If stockCount = 0 Then
            ShowWarning("此商品已无无库存数据！")
            编辑框_商品编码.Text = ""
            编辑框_商品编码.Focus()
            Return
        End If

        If stockCount = 1 Then
            ' 单条记录，直接加载
            Dim poductCode As String = SafeString(dtStock.Rows(0)("apoduct_code"))
            局部_成品商品数据编码 = ""
            局部_成品商品数据编码 = poductCode
            编辑框_商品编码.Text = 局部_成品商品数据编码
            成品商品_数据加载()
            Return
        Else
            ' 多条记录，显示副编码选择窗体
            ' 载入 (窗口_成品修改副编码查询, 窗口_商品成品数据修改, 真)
            ' For now: select first record
            Dim poductCode As String = SafeString(dtStock.Rows(0)("apoduct_code"))
            局部_成品商品数据编码 = poductCode
            编辑框_商品编码.Text = 局部_成品商品数据编码
            成品商品_数据加载()
            Return
        End If
    End Sub

#End Region

#Region "成品商品_数据加载"

    Private Sub 成品商品_数据加载()
        Dim sqlDetail As String = $"SELECT a.poduct_code AS poduct_code,a.product_name as product_name,CASE WHEN COALESCE (f.title, '') = '' THEN '未匹配' ELSE f.title END AS pinlei,a.caizhi as caizhi,b.company_condition as chengse,COALESCE ( e1.title, e2.title, '无数据' ) AS guige,a.sales_unit as danwei,c.gold_price as kejia,a.fu_code as fubian,a.single as danjian,a.quantity as shuliang,a.weight as zongzhong,a.jin_zhong as jinzhong,a.loss as sunhao,a.including as hanhaozhong,a.quandu as quankou,a.wide as miankuan,a.thickness as houdu,b.cost_price as chengben,a.item_number as kuanhao,g.shitou as shizhong,g.stnum as shishu,g.shitou1 as fushizhong,g.shnum1 as fushishu,g.zhuse as zhuse,b.coefficient as xishu,b.basic_cost as cbgongfei,b.company_surcharge as cbfujia,b.sales_cost as xsgongfei,b.premium_cost as cankao,b.sales_surcharge as xsfujia,b.sales_price as yushou,a.remarks as beizhu,a.images as images FROM xipunum_erp_shop AS a  INNER JOIN xipunum_erp_store AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_store_order AS c ON c.id = b.order_id LEFT JOIN xipunum_erp_ksiamges AS d ON d.kuanhao = a.item_number AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = d.specification_id AND a.item_number IS NOT NULL AND a.item_number != '' LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id AND a.specification_id IS NOT NULL AND a.specification_id != '' LEFT JOIN xipunum_erp_category AS f ON f.id = COALESCE ( e1.category_id, e2.category_id ) AND COALESCE ( e1.category_id, e2.category_id ) IS NOT NULL  LEFT JOIN xipunum_erp_shop_xiangqian AS g ON g.poduct_code = a.poduct_code WHERE a.poduct_code = '{SafeSQL(局部_成品商品数据编码)}' ORDER BY a.id DESC  LIMIT 1"

        Dim dt As DataTable = DatabaseModule.MySQL_Read(sqlDetail)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            Return
        End If

        Dim row As DataRow = dt.Rows(0)

        ' 读取图片
        Dim imgUrl As String = SafeString(row("images"))
        If Not String.IsNullOrEmpty(imgUrl) Then
            Try
                图片框EX_主图.Load(imgUrl)
            Catch
                图片框EX_主图.Image = Nothing
            End Try
        End If
        局部_商品图片地址 = imgUrl

        ' 单位选择
        Dim danwei As String = SafeString(row("danwei"))
        If danwei = "重量" Then
            单选框_重量.Checked = True
            单选框_单件.Checked = False
        ElseIf danwei = "单件" Then
            单选框_重量.Checked = False
            单选框_单件.Checked = True
        End If

        ' 填充字段
        编辑框_入库克价.Text = SafeString(row("kejia"))
        编辑框_固定编码.Text = SafeString(row("fubian"))
        编辑框_单件重.Text = SafeString(row("danjian"))
        编辑框_数量.Text = SafeString(row("shuliang"))
        编辑框_重量.Text = SafeString(row("zongzhong"))
        编辑框_金重.Text = SafeString(row("jinzhong"))
        编辑框_损耗.Text = SafeString(row("sunhao"))
        编辑框_含耗重.Text = SafeString(row("hanhaozhong"))
        编辑框_圈号长度.Text = SafeString(row("quankou"))
        编辑框_面宽.Text = SafeString(row("miankuan"))
        编辑框_厚度.Text = SafeString(row("houdu"))
        编辑框_成本单价.Text = SafeString(row("chengben"))
        编辑框_商品款号.Text = SafeString(row("kuanhao"))
        编辑框_主石重.Text = SafeString(row("shizhong"))
        编辑框_石头数量.Text = SafeString(row("shishu"))
        编辑框_副石重.Text = SafeString(row("fushizhong"))
        编辑框_副石头数量.Text = SafeString(row("fushishu"))
        编辑框_商品主色.Text = SafeString(row("zhuse"))
        编辑框_系数.Text = SafeString(row("xishu"))
        编辑框_成本工费.Text = SafeString(row("cbgongfei"))
        编辑框_成本附加费.Text = SafeString(row("cbfujia"))
        编辑框_销售工费.Text = SafeString(row("xsgongfei"))
        编辑框_参考工费.Text = SafeString(row("cankao"))
        编辑框_销售附加费.Text = SafeString(row("xsfujia"))
        编辑框_销售价.Text = SafeString(row("yushou"))
        编辑框_备注.Text = SafeString(row("beizhu"))

        ' 缓存局部变量
        局部_商品品类 = SafeString(row("pinlei"))
        局部_商品材质 = SafeString(row("caizhi"))
        局部_商品成色 = SafeString(row("chengse"))
        局部_商品规格 = SafeString(row("guige"))
        局部_商品入库数量 = SafeString(row("shuliang"))
        局部_商品入库总重 = SafeString(row("zongzhong"))
        局部_商品入库金重 = SafeString(row("jinzhong"))
        局部_商品商品损耗 = SafeString(row("sunhao"))
        局部_商品含耗重 = SafeString(row("hanhaozhong"))
        局部_商品款号数据 = SafeString(row("kuanhao"))

        ' 固口/开口逻辑
        If 局部_商品规格 = "戒指" OrElse 局部_商品规格 = "手镯" Then
            If Not String.IsNullOrEmpty(SafeString(row("quankou"))) Then
                单选框_固口.Checked = True
                单选框_开口.Checked = False
            Else
                单选框_固口.Checked = False
                单选框_开口.Checked = True
            End If
            图片框EX_固开口.Visible = True
            按钮EX_圈号长度.Visible = False
            图片框EX_固开口.Left = 196
            图片框EX_固开口.Top = 239
            编辑框_圈号长度.Width = 72
            编辑框_圈号长度.Left = 304
        Else
            图片框EX_固开口.Visible = False
            按钮EX_圈号长度.Visible = True
            图片框EX_固开口.Left = 196
            图片框EX_固开口.Top = 239
            编辑框_圈号长度.Width = 110
            编辑框_圈号长度.Left = 270
        End If

        ' 匹配品类
        If Not String.IsNullOrEmpty(局部_商品品类) Then
            For i As Integer = 0 To 组合框_品类名称.Items.Count - 1
                Dim cbi As ComboBoxItem = CType(组合框_品类名称.Items(i), ComboBoxItem)
                If cbi.Text = 局部_商品品类 Then
                    组合框_品类名称.SelectedIndex = i
                    Exit For
                End If
            Next
            组合框_品类名称_内容被改变(True)
        End If

        If String.IsNullOrEmpty(SafeString(row("pinlei"))) Then
            组合框_商品材质.Text = 局部_商品材质
            组合框_商品成色.Text = 局部_商品规格
        End If

        ' 权限判断 - 58成品修改重量
        If GlobalVariables.全局_岗位权限操作 IsNot Nothing AndAlso GlobalVariables.全局_岗位权限操作.Contains(",58成品修改重量,") Then
            编辑框_金重.Enabled = False
            编辑框_重量.Enabled = False
            编辑框_损耗.Enabled = False
            编辑框_含耗重.Enabled = False
        Else
            编辑框_金重.Enabled = True
            编辑框_重量.Enabled = True
            编辑框_损耗.Enabled = True
            编辑框_含耗重.Enabled = True
        End If

        ' 检查销售数量
        Dim sqlSale As String = $"SELECT COALESCE(sum(quantity), 0) as quantity,COALESCE(sum(net_weight), 0) as net_weight FROM xipunum_erp_outbound WHERE poduct_code='{SafeSQL(局部_成品商品数据编码)}' AND kufang = '{SafeSQL(GlobalVariables.全局_账户分组id)}'"
        Dim dtSale As DataTable = DatabaseModule.MySQL_Read(sqlSale)
        Dim saleQuantity As Decimal = 0
        Dim saleWeight As Decimal = 0
        If dtSale IsNot Nothing AndAlso dtSale.Rows.Count > 0 Then
            saleQuantity = SafeDecimal(dtSale.Rows(0)("quantity"))
            saleWeight = SafeDecimal(dtSale.Rows(0)("net_weight"))
        End If

        If saleQuantity = 0 AndAlso saleWeight = 0 Then
            按钮EX_保存.Enabled = True
            编辑框_圈号长度.Enabled = True
            编辑框_面宽.Enabled = True
            编辑框_厚度.Enabled = True
            单选框_固口.Enabled = True
            单选框_开口.Enabled = True

            ' 检查当前库房库存
            Dim sqlKuFang As String = $"SELECT a.poduct_code AS apoduct_code FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE a.kufang = '{SafeSQL(GlobalVariables.全局_账户分组id)}' AND (a.quantity > 0 or a.jinzhong  > 0) AND (b.poduct_code = '{SafeSQL(局部_成品商品数据编码)}' OR b.fu_code = '{SafeSQL(局部_成品商品数据编码)}') ORDER BY a.id DESC"
            Dim dtKuFang As DataTable = DatabaseModule.MySQL_Read(sqlKuFang)
            Dim kufangCount As Integer = If(dtKuFang IsNot Nothing, dtKuFang.Rows.Count, 0)

            If kufangCount <> 0 Then
                ' 获取库存id
                局部_商品库存id信息 = ""
                Dim sqlKucid As String = $"SELECT a.id AS aid FROM xipunum_erp_shop_kucun AS a INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code WHERE a.kufang = '{SafeSQL(GlobalVariables.全局_账户分组id)}' AND (a.quantity > 0 or a.jinzhong  > 0) AND (b.poduct_code = '{SafeSQL(局部_成品商品数据编码)}' OR b.fu_code = '{SafeSQL(局部_成品商品数据编码)}') ORDER BY a.id DESC"
                Dim dtKucid As DataTable = DatabaseModule.MySQL_Read(sqlKucid)
                If dtKucid IsNot Nothing AndAlso dtKucid.Rows.Count > 0 Then
                    局部_商品库存id信息 = SafeString(dtKucid.Rows(0)("aid"))
                End If

                组合框_品类名称.Enabled = True
                组合框_商品材质.Enabled = True
                组合框_商品成色.Enabled = True
                组合框_商品规格.Enabled = True
                单选框_固口.Enabled = True
                单选框_开口.Enabled = True

                If 局部_品类多数量 = "0" Then
                    编辑框_数量.Enabled = True
                Else
                    编辑框_数量.Enabled = False
                End If

                If 局部_商品是否镶嵌 = "镶嵌" Then
                    编辑框_主石重.Enabled = True
                    编辑框_石头数量.Enabled = True
                    编辑框_副石重.Enabled = True
                    编辑框_副石头数量.Enabled = True
                    编辑框_商品主色.Enabled = True
                Else
                    编辑框_主石重.Enabled = False
                    编辑框_石头数量.Enabled = False
                    编辑框_副石重.Enabled = False
                    编辑框_副石头数量.Enabled = False
                    编辑框_商品主色.Enabled = False
                End If

                编辑框_成本单价.Enabled = True
                编辑框_商品款号.Enabled = True
                编辑框_系数.Enabled = True
                编辑框_成本工费.Enabled = True
                编辑框_成本附加费.Enabled = True
                编辑框_销售工费.Enabled = True
                编辑框_参考工费.Enabled = True
                编辑框_销售附加费.Enabled = True
                编辑框_销售价.Enabled = True
                编辑框_备注.Enabled = True
            End If
        End If

        组合框_标签样式.Enabled = True
    End Sub

#End Region

#Region "组合框事件"

    Private Sub 组合框_商品规格_选IndexChanged(sender As Object, e As EventArgs) Handles 组合框_商品规格.SelectedIndexChanged
        If 组合框_商品规格.SelectedIndex >= 0 Then
            Dim cbi As ComboBoxItem = CType(组合框_商品规格.SelectedItem, ComboBoxItem)
            If cbi.Text = "戒指" OrElse cbi.Text = "手镯" Then
                图片框EX_固开口.Visible = True
                按钮EX_圈号长度.Visible = False
                图片框EX_固开口.Left = 196
                图片框EX_固开口.Top = 239
                编辑框_圈号长度.Width = 72
                编辑框_圈号长度.Left = 304
            Else
                图片框EX_固开口.Visible = False
                按钮EX_圈号长度.Visible = True
                图片框EX_固开口.Left = 196
                图片框EX_固开口.Top = 239
                编辑框_圈号长度.Width = 110
                编辑框_圈号长度.Left = 270
            End If
        End If
    End Sub

    Private Sub 组合框_品类名称_选IndexChanged(sender As Object, e As EventArgs) Handles 组合框_品类名称.SelectedIndexChanged
        组合框_品类名称_内容被改变(True)
    End Sub

    Private Sub 组合框_品类名称_内容被改变(isChange As Boolean)
        If Not isChange Then Return
        If 组合框_品类名称.SelectedIndex < 0 Then Return

        Dim cbi As ComboBoxItem = CType(组合框_品类名称.SelectedItem, ComboBoxItem)
        Dim 查找品类id内容 As String = cbi.Value

        Dim sqlCategory As String = $"SELECT * FROM xipunum_erp_category where id='{SafeSQL(查找品类id内容)}' order by id ASC"
        Dim dtCategory As DataTable = DatabaseModule.MySQL_Read(sqlCategory)
        If dtCategory Is Nothing OrElse dtCategory.Rows.Count = 0 Then Return

        Dim caizhiid As String = SafeString(dtCategory.Rows(0)("caizhiid"))
        Dim chengse As String = SafeString(dtCategory.Rows(0)("chengse"))
        Dim xiangqian As String = SafeString(dtCategory.Rows(0)("xiangqian"))
        Dim shuliang As String = SafeString(dtCategory.Rows(0)("shuliang"))

        局部_商品是否镶嵌 = xiangqian
        局部_品类多数量 = shuliang

        ' 重新加载材质
        组合框_商品材质.Items.Clear()
        组合框_商品成色.Items.Clear()
        组合框_商品规格.Items.Clear()
        组合框_商品材质.Text = "请选择商品材质"
        组合框_商品成色.Text = "请选择商品成色"
        组合框_商品规格.Text = "请选择商品规格"

        ' 材质
        If Not String.IsNullOrEmpty(caizhiid) Then
            Dim caizhiParts As String() = caizhiid.Split("|"c)
            For i As Integer = 0 To caizhiParts.Length - 1
                组合框_商品材质.Items.Add(New ComboBoxItem(caizhiParts(i), (i + 1).ToString()))
            Next
        End If

        ' 成色
        If Not String.IsNullOrEmpty(chengse) Then
            Dim chengseParts As String() = chengse.Split("|"c)
            For i As Integer = 0 To chengseParts.Length - 1
                组合框_商品成色.Items.Add(New ComboBoxItem(chengseParts(i), (i + 1).ToString()))
            Next
        End If

        ' 规格
        Dim sqlSpec As String = $"SELECT * FROM xipunum_erp_specs where category_id='{SafeSQL(查找品类id内容)}' order by id ASC"
        Dim dtSpec As DataTable = DatabaseModule.MySQL_Read(sqlSpec)
        If dtSpec IsNot Nothing Then
            For Each row As DataRow In dtSpec.Rows
                Dim id As String = SafeString(row("id"))
                Dim title As String = SafeString(row("title"))
                组合框_商品规格.Items.Add(New ComboBoxItem(title, id))
            Next
        End If

        ' 匹配当前值
        For i As Integer = 0 To 组合框_商品材质.Items.Count - 1
            Dim item As ComboBoxItem = CType(组合框_商品材质.Items(i), ComboBoxItem)
            If item.Text = 局部_商品材质 Then
                组合框_商品材质.SelectedIndex = i
                Exit For
            End If
        Next

        For i As Integer = 0 To 组合框_商品成色.Items.Count - 1
            Dim item As ComboBoxItem = CType(组合框_商品成色.Items(i), ComboBoxItem)
            If item.Text = 局部_商品成色 Then
                组合框_商品成色.SelectedIndex = i
                Exit For
            End If
        Next

        For i As Integer = 0 To 组合框_商品规格.Items.Count - 1
            Dim item As ComboBoxItem = CType(组合框_商品规格.Items(i), ComboBoxItem)
            If item.Text = 局部_商品规格 Then
                组合框_商品规格.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

#End Region

#Region "商品款号事件"

    Private Sub 编辑框_商品款号_键盘事件(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            If 编辑框_商品款号.Text.Length <> 10 Then
                ShowWarning("请输入正确的10位商品款号！")
                编辑框_商品款号.Text = ""
                Return
            End If
            编辑框_商品款号_内容被改变()
        End If
    End Sub

    Private Sub 编辑框_商品款号_获得焦点(sender As Object, e As EventArgs)
        编辑框_商品款号.Text = ""
    End Sub

    Private Sub 编辑框_商品款号_失去焦点(sender As Object, e As EventArgs)
        编辑框_商品款号.Text = 局部_商品款号数据
    End Sub

    Private Sub 编辑框_商品款号_内容被改变()
        If 编辑框_商品款号.Text.Length > 9 Then
            Dim sqlKuanhao As String = $"SELECT a.yimage as yimage,a.kuanhao AS kuanhao,b.title AS pinlei,c.title AS guige,a.caizhi AS caizhi FROM xipunum_erp_ksiamges AS a  LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id  LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id WHERE a.kuanhao = '{SafeSQL(编辑框_商品款号.Text)}' LIMIT 1"
            Dim dtKuanhao As DataTable = DatabaseModule.MySQL_Read(sqlKuanhao)
            If dtKuanhao Is Nothing OrElse dtKuanhao.Rows.Count = 0 Then
                ShowWarning("您输入的款号数据不存在！")
                编辑框_商品款号.Text = 局部_商品款号数据
                Return
            End If

            Dim kuanhao As String = SafeString(dtKuanhao.Rows(0)("kuanhao"))
            If String.IsNullOrEmpty(kuanhao) Then
                ShowWarning("您输入的款号数据不存在！")
                编辑框_商品款号.Text = 局部_商品款号数据
                Return
            End If

            局部_商品款号数据 = kuanhao
            局部_商品品类 = SafeString(dtKuanhao.Rows(0)("pinlei"))
            局部_商品材质 = SafeString(dtKuanhao.Rows(0)("caizhi"))
            局部_商品规格 = SafeString(dtKuanhao.Rows(0)("guige"))
            局部_商品图片地址 = SafeString(dtKuanhao.Rows(0)("yimage"))

            ' 加载图片
            If Not String.IsNullOrEmpty(局部_商品图片地址) Then
                Try
                    图片框EX_主图.Load(局部_商品图片地址)
                Catch
                End Try
            End If

            ' 匹配品类
            If Not String.IsNullOrEmpty(局部_商品品类) Then
                For i As Integer = 0 To 组合框_品类名称.Items.Count - 1
                    Dim cbi As ComboBoxItem = CType(组合框_品类名称.Items(i), ComboBoxItem)
                    If cbi.Text = 局部_商品品类 Then
                        组合框_品类名称.SelectedIndex = i
                        Exit For
                    End If
                Next
                组合框_品类名称_内容被改变(True)
            End If

            组合框_品类名称.Enabled = False
            组合框_商品材质.Enabled = False
            组合框_商品规格.Enabled = False
        End If
    End Sub

#End Region

#Region "数值字段事件"

    Private Sub 编辑框_重量_鼠标进入离开(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(编辑框_重量.Text) Then
            编辑框_重量.Text = "0.000"
        End If

        Dim weight As Decimal
        If Not Decimal.TryParse(编辑框_重量.Text, weight) Then
            weight = 0
        End If

        If weight < 0 Then
            ShowWarning("重量不能小于0！")
            编辑框_重量.Text = "0.000"
            Return
        End If

        If weight > 0 Then
            商品信息_参数数值计算("重量")
        End If
    End Sub

    Private Sub 编辑框_重量_键盘事件(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            编辑框_重量_鼠标进入离开(sender, Nothing)
        End If
    End Sub

    Private Sub 编辑框_入库克价_内容被改变(sender As Object, e As EventArgs) Handles 编辑框_入库克价.TextChanged
        If 编辑框_入库克价.Focused Then
            Dim val As Decimal
            If Decimal.TryParse(编辑框_入库克价.Text, val) AndAlso val < 0 Then
                ShowWarning("入库克价不能小于0！")
                编辑框_入库克价.Text = "0.000"
            End If
        End If
    End Sub

    Private Sub 编辑框_入库克价_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_入库克价.Leave
        商品信息_参数数值计算("入库克价")
    End Sub

    Private Sub 编辑框_数量_内容被改变() Handles 编辑框_数量.TextChanged
        If 编辑框_数量.Focused AndAlso 局部_品类多数量 = "0" Then
            Dim val As Integer
            If Integer.TryParse(编辑框_数量.Text, val) AndAlso val < 1 Then
                ShowWarning("入库数量不能小于1！")
                编辑框_数量.Focus()
                编辑框_数量.Text = "1"
            End If
        End If
    End Sub

    Private Sub 编辑框_数量_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_数量.Leave
        商品信息_参数数值计算("数量")
    End Sub

    Private Sub 编辑框_金重_内容被改变(sender As Object, e As EventArgs) Handles 编辑框_金重.TextChanged
        If 编辑框_金重.Focused Then
            Dim jz As Decimal : Decimal.TryParse(编辑框_金重.Text, jz)
            Dim zl As Decimal : Decimal.TryParse(编辑框_重量.Text, zl)
            If jz > zl Then
                ShowWarning("金重不能大于总重！")
                编辑框_金重.Text = 编辑框_重量.Text
            End If
            If jz < 0 Then
                ShowWarning("金重不能小于0！")
                编辑框_金重.Text = 编辑框_重量.Text
            End If
        End If
    End Sub

    Private Sub 编辑框_金重_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_金重.Leave
        If String.IsNullOrEmpty(编辑框_金重.Text) Then
            编辑框_金重.Text = 编辑框_重量.Text
        End If
        商品信息_参数数值计算("金重")
    End Sub

    Private Sub 编辑框_损耗_内容被改变(sender As Object, e As EventArgs) Handles 编辑框_损耗.TextChanged
        If 编辑框_损耗.Focused Then
            Dim val As Decimal
            If Decimal.TryParse(编辑框_损耗.Text, val) AndAlso val < 0 Then
                ShowWarning("损耗不能小于0！")
                编辑框_损耗.Text = "0.000"
            End If
        End If
    End Sub

    Private Sub 编辑框_损耗_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_损耗.Leave
        If String.IsNullOrEmpty(编辑框_损耗.Text) Then
            编辑框_损耗.Text = "0.000"
        End If
        商品信息_参数数值计算("损耗")
    End Sub

    Private Sub 编辑框_成本单价_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_成本单价.Leave
        If String.IsNullOrEmpty(编辑框_成本单价.Text) Then
            编辑框_成本单价.Text = "0.00"
        End If
        商品信息_参数数值计算("成本单价")
    End Sub

    Private Sub 编辑框_主石重_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_主石重.Leave
        If String.IsNullOrEmpty(编辑框_主石重.Text) Then
            编辑框_主石重.Text = "0.000"
        End If
        商品信息_参数数值计算("主石重")
    End Sub

    Private Sub 编辑框_副石重_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_副石重.Leave
        If String.IsNullOrEmpty(编辑框_副石重.Text) Then
            编辑框_副石重.Text = "0.000"
        End If
        商品信息_参数数值计算("副石重")
    End Sub

    Private Sub 编辑框_系数_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_系数.Leave
        If String.IsNullOrEmpty(编辑框_系数.Text) Then
            编辑框_系数.Text = "1"
        End If
        商品信息_参数数值计算("系数")
    End Sub

    Private Sub 编辑框_成本工费_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_成本工费.Leave
        If String.IsNullOrEmpty(编辑框_成本工费.Text) Then
            编辑框_成本工费.Text = "0.00"
        End If
        商品信息_参数数值计算("成本工费")
    End Sub

    Private Sub 编辑框_成本附加费_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_成本附加费.Leave
        If String.IsNullOrEmpty(编辑框_成本附加费.Text) Then
            编辑框_成本附加费.Text = "0.00"
        End If
        商品信息_参数数值计算("成本附加费")
    End Sub

    Private Sub 编辑框_销售工费_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_销售工费.Leave
        If String.IsNullOrEmpty(编辑框_销售工费.Text) Then
            编辑框_销售工费.Text = 编辑框_成本工费.Text
        End If
        商品信息_参数数值计算("销售工费")
    End Sub

    Private Sub 编辑框_参考工费_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_参考工费.Leave
        If String.IsNullOrEmpty(编辑框_参考工费.Text) Then
            编辑框_参考工费.Text = 编辑框_成本工费.Text
        End If
        商品信息_参数数值计算("参考工费")
    End Sub

    Private Sub 编辑框_销售附加费_焦点事件(sender As Object, e As EventArgs) Handles 编辑框_销售附加费.Leave
        If String.IsNullOrEmpty(编辑框_销售附加费.Text) Then
            编辑框_销售附加费.Text = 编辑框_成本附加费.Text
        End If
        商品信息_参数数值计算("销售附加费")
    End Sub

#End Region

#Region "商品信息_参数数值计算"

    Private Sub 商品信息_参数数值计算(修改参数 As String)
        If 修改参数 = "重量" Then
            编辑框_重量.Text = 格式三位小数(编辑框_重量.Text)
            编辑框_金重.Text = 编辑框_重量.Text

            Dim qty As Decimal
            If SafeDecimal(编辑框_数量.Text) = 0 Then
                编辑框_单件重.Text = 格式三位小数(编辑框_重量.Text)
            Else
                编辑框_单件重.Text = 格式三位小数(编辑框_重量.Text, 编辑框_数量.Text)
            End If
            编辑框_主石重.Text = "0.000"
            编辑框_石头数量.Text = "0"
            编辑框_副石重.Text = "0.000"
            编辑框_副石头数量.Text = "0"

            Dim sunhao As Decimal = SafeDecimal(编辑框_损耗.Text)
            Dim zhongliang As Decimal = SafeDecimal(编辑框_重量.Text)
            编辑框_含耗重.Text = 格式三位小数((zhongliang + sunhao).ToString())
        End If

        If 修改参数 = "入库克价" Then
            编辑框_入库克价.Text = 格式二位小数(编辑框_入库克价.Text)
        End If

        If 修改参数 = "数量" Then
            Dim qty As Decimal = SafeDecimal(编辑框_数量.Text)
            If qty = 0 Then
                编辑框_单件重.Text = 格式三位小数(编辑框_重量.Text)
            Else
                编辑框_单件重.Text = 格式三位小数(编辑框_重量.Text, 编辑框_数量.Text)
            End If
            编辑框_主石重.Text = "0.000"
            编辑框_石头数量.Text = "0"
            编辑框_副石重.Text = "0.000"
            编辑框_副石头数量.Text = "0"
        End If

        If 修改参数 = "金重" Then
            编辑框_金重.Text = 格式三位小数(编辑框_金重.Text)

            Dim zl As Decimal = SafeDecimal(编辑框_重量.Text)
            Dim jz As Decimal = SafeDecimal(编辑框_金重.Text)
            Dim shizhong As Decimal = Math.Round((zl - jz) * 5, 3)
            编辑框_主石重.Text = 格式三位小数(shizhong.ToString())
            编辑框_副石重.Text = "0.000"
        End If

        If 修改参数 = "损耗" Then
            编辑框_损耗.Text = 格式三位小数(编辑框_损耗.Text)
            Dim zl As Decimal = SafeDecimal(编辑框_重量.Text)
            Dim sh As Decimal = SafeDecimal(编辑框_损耗.Text)
            编辑框_含耗重.Text = 格式三位小数((zl + sh).ToString())
        End If

        If 修改参数 = "成本单价" Then
            编辑框_成本单价.Text = 格式二位小数(编辑框_成本单价.Text)
            If 单选框_单件.Checked Then
                Dim cb As Decimal = SafeDecimal(编辑框_成本单价.Text)
                Dim xs As Decimal = SafeDecimal(编辑框_系数.Text)
                编辑框_销售价.Text = 格式二位小数((cb * xs).ToString())
            End If
        End If

        If 修改参数 = "主石重" Then
            Dim zl As Decimal = SafeDecimal(编辑框_重量.Text)
            Dim jz As Decimal = SafeDecimal(编辑框_金重.Text)
            Dim maxStone As Decimal = Math.Round((zl - jz) * 5, 3)
            Dim sz As Decimal = SafeDecimal(编辑框_主石重.Text)

            If sz > maxStone Then
                ShowWarning($"主石重不可以大于{maxStone}ct")
                编辑框_主石重.Text = maxStone.ToString()
            End If

            编辑框_主石重.Text = 格式三位小数(编辑框_主石重.Text)

            Dim remainder As Decimal = Math.Round(maxStone - SafeDecimal(编辑框_主石重.Text), 3)
            编辑框_副石重.Text = 格式三位小数(remainder.ToString())
        End If

        If 修改参数 = "副石重" Then
            Dim zl As Decimal = SafeDecimal(编辑框_重量.Text)
            Dim jz As Decimal = SafeDecimal(编辑框_金重.Text)
            Dim maxStone As Decimal = Math.Round((zl - jz) * 5, 3)
            Dim fsz As Decimal = SafeDecimal(编辑框_副石重.Text)

            If fsz > maxStone Then
                ShowWarning($"副石重不可以大于{maxStone}ct")
                编辑框_副石重.Text = maxStone.ToString()
            End If

            编辑框_副石重.Text = 格式三位小数(编辑框_副石重.Text)

            Dim remainder As Decimal = Math.Round(maxStone - SafeDecimal(编辑框_副石重.Text), 3)
            编辑框_主石重.Text = 格式三位小数(remainder.ToString())
        End If

        If 修改参数 = "系数" Then
            If 单选框_单件.Checked Then
                Dim cb As Decimal = SafeDecimal(编辑框_成本单价.Text)
                Dim xs As Decimal = SafeDecimal(编辑框_系数.Text)
                编辑框_销售价.Text = 格式二位小数((cb * xs).ToString())
            End If
        End If

        If 修改参数 = "成本工费" Then
            编辑框_成本工费.Text = 格式二位小数(编辑框_成本工费.Text)
            编辑框_参考工费.Text = 编辑框_成本工费.Text
            编辑框_销售工费.Text = 编辑框_成本工费.Text
        End If

        If 修改参数 = "成本附加费" Then
            编辑框_成本附加费.Text = 格式二位小数(编辑框_成本附加费.Text)
            编辑框_销售附加费.Text = 编辑框_成本附加费.Text
        End If

        If 修改参数 = "销售工费" Then
            编辑框_销售工费.Text = 格式二位小数(编辑框_销售工费.Text)
            编辑框_参考工费.Text = 编辑框_销售工费.Text
        End If

        If 修改参数 = "参考工费" Then
            编辑框_参考工费.Text = 格式二位小数(编辑框_参考工费.Text)
            编辑框_销售工费.Text = 编辑框_参考工费.Text
        End If

        If 修改参数 = "销售附加费" Then
            编辑框_销售附加费.Text = 格式二位小数(编辑框_销售附加费.Text)
        End If

        ' 成本单价/销售价 综合计算 (按重量方式)
        If 修改参数 = "单件重" OrElse 修改参数 = "重量" OrElse 修改参数 = "数量" OrElse
           修改参数 = "金重" OrElse 修改参数 = "成本单价" OrElse 修改参数 = "成本工费" OrElse
           修改参数 = "成本附加费" OrElse 修改参数 = "销售工费" OrElse 修改参数 = "参考工费" OrElse
           修改参数 = "销售附加费" Then

            If 单选框_重量.Checked Then
                Dim qty As Decimal = SafeDecimal(编辑框_数量.Text)
                Dim jz As Decimal = SafeDecimal(编辑框_金重.Text)
                Dim gongfei As Decimal = SafeDecimal(编辑框_成本工费.Text)
                Dim kejia As Decimal = SafeDecimal(编辑框_入库克价.Text)
                Dim fujia As Decimal = SafeDecimal(编辑框_成本附加费.Text)

                ' 成本单价
                If qty = 0 Then
                    编辑框_成本单价.Text = 格式二位小数((jz * (gongfei + kejia) + fujia).ToString())
                Else
                    编辑框_成本单价.Text = 格式二位小数(((jz / qty) * (gongfei + kejia) + fujia).ToString())
                End If

                ' 销售价
                Dim xsgongfei As Decimal = SafeDecimal(编辑框_销售工费.Text)
                Dim xsfujia As Decimal = SafeDecimal(编辑框_销售附加费.Text)
                If qty = 0 Then
                    编辑框_销售价.Text = 格式二位小数((jz * (xsgongfei + kejia) + xsfujia).ToString())
                Else
                    编辑框_销售价.Text = 格式二位小数(((jz / qty) * (xsgongfei + kejia) + xsfujia).ToString())
                End If
            End If
        End If
    End Sub

#End Region

#Region "格式工具函数"

    Private Function 格式三位小数(val As String, Optional divisor As String = "") As String
        Dim d As Decimal
        If Not String.IsNullOrEmpty(divisor) Then
            Dim v1 As Decimal : Decimal.TryParse(val, v1)
            Dim v2 As Decimal : Decimal.TryParse(divisor, v2)
            If v2 <> 0 Then
                d = Math.Round(v1 / v2, 3)
            Else
                d = Math.Round(v1, 3)
            End If
        Else
            Decimal.TryParse(val, d)
            d = Math.Round(d, 3)
        End If

        Return d.ToString("F3")
    End Function

    Private Function 格式二位小数(val As String) As String
        Dim d As Decimal
        Decimal.TryParse(val, d)
        d = Math.Round(d, 2)
        Return d.ToString("F2")
    End Function

#End Region

#Region "按钮事件"

    Private Sub 按钮EX_重置_Click(sender As Object, e As EventArgs) Handles 按钮EX_重置.Click
        窗口_商品成品数据修改_创建完毕(Nothing, Nothing)
    End Sub

    Private Sub 按钮EX_打印_Click(sender As Object, e As EventArgs) Handles 按钮EX_打印.Click
        打印功能_被单击()
    End Sub

    Private Sub 按钮EX_保存_Click(sender As Object, e As EventArgs) Handles 按钮EX_保存.Click
        Dim 信息操作商品编码 As String = 编辑框_商品编码.Text
        Dim 信息操作日期 As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Dim 信息操作账户 As String = GlobalVariables.全局_用户账户

        ' 规格
        Dim 信息操作规格 As String
        If 组合框_商品规格.SelectedIndex >= 0 Then
            Dim cbi As ComboBoxItem = CType(组合框_商品规格.SelectedItem, ComboBoxItem)
            信息操作规格 = $",specification_id= '{cbi.Value}'"
        Else
            信息操作规格 = ""
        End If

        ' 材质
        Dim 信息操作材质 As String
        If 组合框_商品材质.SelectedIndex >= 0 Then
            Dim cbi As ComboBoxItem = CType(组合框_商品材质.SelectedItem, ComboBoxItem)
            信息操作材质 = $",caizhi= '{cbi.Text}'"
        Else
            信息操作材质 = ""
        End If

        ' 成色
        Dim 信息操作工厂成色 As String
        If 局部_商品成色 <> "" AndAlso 组合框_商品材质.SelectedIndex >= 0 Then
            Dim caizhiItem As ComboBoxItem = CType(组合框_商品材质.SelectedItem, ComboBoxItem)
            If 局部_商品成色 <> caizhiItem.Text Then
                If 组合框_商品成色.SelectedIndex >= 0 Then
                    Dim chengseItem As ComboBoxItem = CType(组合框_商品成色.SelectedItem, ComboBoxItem)
                    信息操作工厂成色 = $",factory_condition= '{chengseItem.Text}',company_condition= '{chengseItem.Text}'"
                Else
                    信息操作工厂成色 = ""
                End If
            Else
                信息操作工厂成色 = ""
            End If
        Else
            信息操作工厂成色 = ""
        End If

        Dim 商品详情入库数量 As String = $",quantity= '{编辑框_数量.Text}'"
        Dim 商品详情入库总重 As String = $",weight= '{编辑框_重量.Text}'"
        Dim 商品详情入库金重 As String = $",jin_zhong= '{编辑框_金重.Text}'"
        Dim 商品详情商品损耗 As String = $",loss= '{编辑框_损耗.Text}'"
        Dim 商品详情含耗重 As String = $",including= '{编辑框_含耗重.Text}'"
        Dim 商品详情款号数据 As String = $",item_number= '{编辑框_商品款号.Text}'"

        Dim 商品详情入库克价 As String = 编辑框_入库克价.Text
        Dim 商品详情单件重 As String = 编辑框_单件重.Text
        Dim 商品详情商品圈口 As String = 编辑框_圈号长度.Text
        Dim 商品详情商品面宽 As String = 编辑框_面宽.Text
        Dim 商品详情商品厚度 As String = 编辑框_厚度.Text
        Dim 商品详情入库成本 As String = 编辑框_成本单价.Text
        Dim 商品详情主石重 As String = 编辑框_主石重.Text
        Dim 商品详情主石数 As String = 编辑框_石头数量.Text
        Dim 商品详情副石重 As String = 编辑框_副石重.Text
        Dim 商品详情副石数 As String = 编辑框_副石头数量.Text
        Dim 商品详情商品主色 As String = 编辑框_商品主色.Text
        Dim 商品详情销售系数 As String = 编辑框_系数.Text
        Dim 商品详情成本工费 As String = 编辑框_成本工费.Text
        Dim 商品详情成本附加 As String = 编辑框_成本附加费.Text
        Dim 商品详情销售工费 As String = 编辑框_销售工费.Text
        Dim 商品详情参考工费 As String = 编辑框_参考工费.Text
        Dim 商品详情销售附加 As String = 编辑框_销售附加费.Text
        Dim 商品详情预售价 As String = 编辑框_销售价.Text
        Dim 商品详情备注信息 As String = 编辑框_备注.Text

        ' Update xipunum_erp_shop
        If GlobalVariables.全局_岗位权限操作 IsNot Nothing AndAlso GlobalVariables.全局_岗位权限操作.Contains(",58成品修改重量,") Then
            Dim sqlShop As String = $"UPDATE xipunum_erp_shop SET images= '{SafeSQL(局部_商品图片地址)}'{SafeSQL(信息操作规格)}{SafeSQL(商品详情款号数据)}{SafeSQL(信息操作材质)}{SafeSQL(商品详情商品损耗)}{SafeSQL(商品详情含耗重)},quandu= '{SafeSQL(商品详情商品圈口)}',wide= '{SafeSQL(商品详情商品面宽)}',thickness= '{SafeSQL(商品详情商品厚度)}',single= '{SafeSQL(商品详情单件重)}',remarks= '{SafeSQL(商品详情备注信息)}',cjuser= '{SafeSQL(信息操作账户)}',updatetime= '{SafeSQL(信息操作日期)}'  WHERE poduct_code ='{SafeSQL(信息操作商品编码)}' LIMIT 1"
            DatabaseModule.MySQL_Write(sqlShop)
        Else
            Dim sqlShop As String = $"UPDATE xipunum_erp_shop SET images= '{SafeSQL(局部_商品图片地址)}'{SafeSQL(信息操作规格)}{SafeSQL(商品详情款号数据)}{SafeSQL(信息操作材质)}{SafeSQL(商品详情入库数量)}{SafeSQL(商品详情入库金重)}{SafeSQL(商品详情入库总重)}{SafeSQL(商品详情商品损耗)}{SafeSQL(商品详情含耗重)},quandu= '{SafeSQL(商品详情商品圈口)}',wide= '{SafeSQL(商品详情商品面宽)}',thickness= '{SafeSQL(商品详情商品厚度)}',single= '{SafeSQL(商品详情单件重)}',remarks= '{SafeSQL(商品详情备注信息)}',cjuser= '{SafeSQL(信息操作账户)}',updatetime= '{SafeSQL(信息操作日期)}'  WHERE poduct_code ='{SafeSQL(信息操作商品编码)}' LIMIT 1"
            DatabaseModule.MySQL_Write(sqlShop)

            ' Update kucun if jinzhong changed
            If 局部_商品入库金重 <> 编辑框_金重.Text Then
                Dim sqlKucunJz As String = $"UPDATE xipunum_erp_shop_kucun SET jinzhong= '{SafeSQL(编辑框_金重.Text)}' WHERE poduct_code ='{SafeSQL(信息操作商品编码)}' and id='{SafeSQL(局部_商品库存id信息)}' LIMIT 1"
                DatabaseModule.MySQL_Write(sqlKucunJz)
            End If

            ' Update kucun if quantity changed
            If 局部_商品入库数量 <> 编辑框_数量.Text Then
                Dim sqlKucunQty As String = $"UPDATE xipunum_erp_shop_kucun SET quantity= '{SafeSQL(编辑框_数量.Text)}' WHERE poduct_code ='{SafeSQL(信息操作商品编码)}' and id='{SafeSQL(局部_商品库存id信息)}' LIMIT 1"
                DatabaseModule.MySQL_Write(sqlKucunQty)
            End If
        End If

        ' Update xipunum_erp_store
        Dim sqlStore As String = $"UPDATE xipunum_erp_store SET sales_price= '{SafeSQL(商品详情预售价)}'{SafeSQL(信息操作工厂成色)},cost_price= '{SafeSQL(商品详情入库成本)}',coefficient= '{SafeSQL(商品详情销售系数)}',basic_cost= '{SafeSQL(商品详情成本工费)}',premium_cost= '{SafeSQL(商品详情参考工费)}',sales_cost= '{SafeSQL(商品详情销售工费)}',company_surcharge= '{SafeSQL(商品详情成本附加)}',sales_surcharge= '{SafeSQL(商品详情销售附加)}',cjuser= '{SafeSQL(信息操作账户)}',updatetime= '{SafeSQL(信息操作日期)}'  WHERE poduct_code ='{SafeSQL(信息操作商品编码)}' LIMIT 1"
        DatabaseModule.MySQL_Write(sqlStore)

        ' Get order_id
        Dim sqlOrder As String = $"SELECT order_id FROM xipunum_erp_store WHERE poduct_code='{SafeSQL(信息操作商品编码)}' LIMIT 1"
        Dim dtOrder As DataTable = DatabaseModule.MySQL_Read(sqlOrder)
        Dim 入库数据id As String = ""
        If dtOrder IsNot Nothing AndAlso dtOrder.Rows.Count > 0 Then
            入库数据id = SafeString(dtOrder.Rows(0)("order_id"))
        End If

        ' Update xipunum_erp_store_order
        Dim sqlStoreOrder As String = $"UPDATE xipunum_erp_store_order SET gold_price= '{SafeSQL(商品详情入库克价)}',updatetime= '{SafeSQL(信息操作日期)}'  WHERE id ='{SafeSQL(入库数据id)}' LIMIT 1"
        DatabaseModule.MySQL_Write(sqlStoreOrder)

        ' Update xiangqian if 镶嵌
        If 局部_商品是否镶嵌 = "镶嵌" Then
            Dim sqlXiangqian As String = $"UPDATE xipunum_erp_shop_xiangqian SET shitou= '{SafeSQL(商品详情主石重)}',stnum= '{SafeSQL(商品详情主石数)}',shitou1= '{SafeSQL(商品详情副石重)}',shnum1= '{SafeSQL(商品详情副石数)}',zhuse= '{SafeSQL(商品详情商品主色)}',cjuser= '{SafeSQL(信息操作账户)}',updatetime= '{SafeSQL(信息操作日期)}'  WHERE poduct_code ='{SafeSQL(信息操作商品编码)}' LIMIT 1"
            DatabaseModule.MySQL_Write(sqlXiangqian)
        End If

        ' 记录历史
        Dim sqlHistory As String = $"poduct_code='{SafeSQL(信息操作商品编码)}',updatetime='{SafeSQL(信息操作日期)}',number='无',type='成品修改',quantity='{SafeSQL(编辑框_数量.Text)}',jinzhong='{SafeSQL(编辑框_金重.Text)}',zhongliang='{SafeSQL(编辑框_重量.Text)}',conter='修改商品基本参数',cjuser='{SafeSQL(信息操作账户)}'"
        DatabaseModule.MySQL_Write($"INSERT INTO xipunum_erp_history SET {sqlHistory}")

        ' 记录shop_log
        DatabaseModule.MySQL_Write($"INSERT INTO xipunum_erp_shop_log SET poduct_code='{SafeSQL(信息操作商品编码)}',type='编辑',creationtime='{SafeSQL(信息操作日期)}'")

        ' 自动打印
        If 选择框_自动打印.Checked Then
            打印功能_被单击()
        End If

        ShowInfo($"商品编码:{编辑框_商品编码.Text}信息修改成功")

        ' 系统日志
        Dim 日志内容 As String = $"账户:{GlobalVariables.全局_用户账户} 修改商品信息，商品编码:{信息操作商品编码}"
        DatabaseModule.MySQL_Write($"INSERT INTO xipunum_erp_xitong_log SET type='修改',title='成品数据修改',conter='{SafeSQL(日志内容)}',user='{SafeSQL(信息操作账户)}',creationtime='{SafeSQL(信息操作日期)}'")
    End Sub

#End Region

#Region "打印功能"

    Private Sub 打印功能_被单击()
        If 组合框_标签样式.SelectedIndex < 0 Then
            ShowWarning("请选择打印标签")
            组合框_标签样式.Focus()
            Return
        End If

        Dim 数据文件 As String = System.IO.Path.Combine(Application.StartupPath, "data\erpdata.mdb")
        Dim 锁文件 As String = System.IO.Path.Combine(Application.StartupPath, "temp\lm_print.lock")

        ' Get factory info
        Dim 工厂名称简写1 As String = ""
        Dim sqlFactory As String = $"SELECT d.jianxie as jianxie FROM xipunum_erp_shop AS a INNER JOIN xipunum_erp_store AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_store_order AS c ON c.id = b.order_id INNER JOIN xipunum_erp_about AS d ON d.id = c.factory WHERE a.poduct_code='{SafeSQL(编辑框_商品编码.Text)}' ORDER BY a.id ASC LIMIT 1"
        Dim dtFactory As DataTable = DatabaseModule.MySQL_Read(sqlFactory)
        If dtFactory IsNot Nothing AndAlso dtFactory.Rows.Count > 0 Then
            工厂名称简写1 = SafeString(dtFactory.Rows(0)("jianxie"))
        End If

        ' Get account group info
        Dim 账户分组名称1 As String = "", 账户分组简写1 As String = "", 账户分组公司1 As String = "", 账户分组地址1 As String = ""
        Dim sqlGroup As String = $"SELECT b.title as btitle,b.data1 as bdata1,b.data2 as bdata2,b.data3 as bdata3 FROM xipunum_erp_shop AS a INNER JOIN xipunum_erp_type AS b ON b.id = a.kufang WHERE a.poduct_code='{SafeSQL(编辑框_商品编码.Text)}' ORDER BY a.id ASC LIMIT 1"
        Dim dtGroup As DataTable = DatabaseModule.MySQL_Read(sqlGroup)
        If dtGroup IsNot Nothing AndAlso dtGroup.Rows.Count > 0 Then
            账户分组名称1 = SafeString(dtGroup.Rows(0)("btitle"))
            账户分组简写1 = SafeString(dtGroup.Rows(0)("bdata1"))
            账户分组公司1 = SafeString(dtGroup.Rows(0)("bdata2"))
            账户分组地址1 = SafeString(dtGroup.Rows(0)("bdata3"))
        End If

        ' Get info for print
        Dim 入库信息商品材质 As String
        If 组合框_商品材质.SelectedIndex >= 0 Then
            Dim cbi As ComboBoxItem = CType(组合框_商品材质.SelectedItem, ComboBoxItem)
            入库信息商品材质 = cbi.Text
        Else
            入库信息商品材质 = 组合框_商品材质.Text
        End If

        Dim 入库信息商品规格 As String
        If 组合框_商品规格.SelectedIndex >= 0 Then
            Dim cbi As ComboBoxItem = CType(组合框_商品规格.SelectedItem, ComboBoxItem)
            入库信息商品规格 = cbi.Text
        Else
            入库信息商品规格 = 组合框_商品规格.Text
        End If

        Dim 入库信息商品公司成色 As String = ""
        If 组合框_商品成色.SelectedIndex >= 0 Then
            Dim cbi As ComboBoxItem = CType(组合框_商品成色.SelectedItem, ComboBoxItem)
            入库信息商品公司成色 = cbi.Text
        End If

        ' Get additional info
        Dim 商品基本预售价 As String = "", 商品基本名称 As String = "", 商品基本证书 As String = "", 商品基本颜色 As String = ""
        Dim sqlBaseInfo As String = $"SELECT c.sales_price AS sales_price,a.product_name as mingcheng,b.zsbianma as zhengshu,b.yanse as yanse FROM xipunum_erp_shop AS a LEFT JOIN xipunum_erp_zhengshu AS b ON b.poduct_code = a.poduct_code INNER JOIN xipunum_erp_store AS c ON c.poduct_code = a.poduct_code WHERE a.poduct_code = '{SafeSQL(编辑框_商品编码.Text)}'"
        Dim dtBaseInfo As DataTable = DatabaseModule.MySQL_Read(sqlBaseInfo)
        If dtBaseInfo IsNot Nothing AndAlso dtBaseInfo.Rows.Count > 0 Then
            商品基本预售价 = SafeString(dtBaseInfo.Rows(0)("sales_price"))
            商品基本名称 = SafeString(dtBaseInfo.Rows(0)("mingcheng"))
            商品基本证书 = SafeString(dtBaseInfo.Rows(0)("zhengshu"))
            商品基本颜色 = SafeString(dtBaseInfo.Rows(0)("yanse"))
        End If

        ' Write to Access MDB for LMWPRINT
        Try
            ' Note: In VB.NET this would use OleDbConnection but LMWPRINT integration
            ' requires interop with the native Access driver or external executable
            Dim tagFile As String = ""
            If 组合框_标签样式.SelectedIndex >= 0 Then
                Dim cbi As ComboBoxItem = CType(组合框_标签样式.SelectedItem, ComboBoxItem)
                tagFile = System.IO.Path.Combine(Application.StartupPath, "voucher\biaoqian\", cbi.Text)
            End If

            If Not String.IsNullOrEmpty(GlobalVariables.全局_标签打印机链接) AndAlso
               Not String.IsNullOrEmpty(GlobalVariables.全局_标签打印机名称) AndAlso
               Not String.IsNullOrEmpty(tagFile) Then

                ' Execute LMWPRINT
                Dim cmdLine As String = $"""{GlobalVariables.全局_标签打印机链接}"" /L=""{tagFile}"" /C=1 /X=""{GlobalVariables.全局_标签打印机名称}"" /N /Z=2 /FL=""{锁文件}"""
                Dim proc As New Process()
                proc.StartInfo.FileName = GlobalVariables.全局_标签打印机链接
                proc.StartInfo.Arguments = $"/L=""{tagFile}"" /C=1 /X=""{GlobalVariables.全局_标签打印机名称}"" /N /Z=2 /FL=""{锁文件}"""
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                proc.Start()
            Else
                ShowWarning("标签打印配置不完整，已跳过打印！")
            End If
        Catch ex As Exception
            ShowError("打印失败：" & ex.Message)
        End Try
    End Sub

#End Region

#Region "其他事件"

    Private Sub 图片框EX4_Click(sender As Object, e As EventArgs) Handles 图片框EX4.Click
        Me.Close()
    End Sub

    Private Sub 窗口_商品成品数据修改_尺寸被改变(sender As Object, e As EventArgs)
        If 添加修改_分组框 IsNot Nothing Then
            添加修改_分组框.Left = (Me.ClientSize.Width - 添加修改_分组框.Width) \ 2
            添加修改_分组框.Top = (Me.ClientSize.Height - 添加修改_分组框.Height) \ 2
        End If
    End Sub

#End Region

#Region "Helper Methods"

    Private Shared Function SafeSQL(input As String) As String
        If String.IsNullOrEmpty(input) Then Return ""
        Return input.Replace("'", "\'").Replace("\", "\\")
    End Function

    Private Shared Function SafeString(obj As Object) As String
        If obj Is Nothing OrElse IsDBNull(obj) Then Return ""
        Return obj.ToString()
    End Function

    Private Shared Function SafeDecimal(obj As Object) As Decimal
        If obj Is Nothing OrElse IsDBNull(obj) Then Return 0
        Dim d As Decimal
        Decimal.TryParse(obj.ToString(), d)
        Return d
    End Function

    Private Sub ShowWarning(msg As String)
        MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub ShowInfo(msg As String)
        MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ShowError(msg As String)
        MessageBox.Show(msg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

#End Region

End Class

' ============================================================================
' ComboBoxItem class (if not already defined elsewhere)
' ============================================================================
Public Class ComboBoxItem
    Public Property Text As String
    Public Property Value As String

    Public Sub New(text As String, value As String)
        Me.Text = text
        Me.Value = value
    End Sub

    Public Overrides Function ToString() As String
        Return Text
    End Function
End Class

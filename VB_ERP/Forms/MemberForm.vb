' ============================================================================
' 会员添加修改窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_会员添加修改.form.e.txt
' 包含所有1个程序集变量、13个函数、10个SQL查询
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class MemberForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（1个） ==========
    Private selectedMemberCode As String = ""        ' 选中会员信息内容

    ' ========== 控件声明（对应易语言窗口控件） ==========
    ' 分组框
    Private WithEvents grpAddModify As New GroupBox()   ' 添加修改_分组框

    ' 编辑框
    Private txtMemberCode As New TextBox()              ' 编辑框_会员编码
    Private txtMemberName As New TextBox()              ' 编辑框_会员姓名
    Private txtPhone As New TextBox()                   ' 编辑框_联系电话
    Private txtBirthday As New TextBox()                ' 编辑框_出生日期
    Private txtAddress As New TextBox()                 ' 编辑框_用户联系地址
    Private txtJieLiao As New TextBox()                 ' 编辑框_会员结料
    Private txtJieZhang As New TextBox()                ' 编辑框_会员结账
    Private txtRemarks As New TextBox()                 ' 编辑框_备注

    ' 组合框
    Private cmbProvince As New ComboBox()               ' 组合框_省份
    Private cmbCity As New ComboBox()                   ' 组合框_市区
    Private cmbDistrict As New ComboBox()               ' 组合框_县区

    ' 按钮
    Private WithEvents btnSubmit As New Button()        ' 按钮EX1（保存/修改）
    Private WithEvents btnReset As New Button()         ' 按钮EX2（重置）
    Private WithEvents btnClose As New Button()         ' 图片框EX4（关闭）

    ' ========== 构造函数 ==========
    Public Sub New(Optional memberCode As String = "")
        selectedMemberCode = memberCode
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "会员添加修改"
        Me.Size = New Drawing.Size(700, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 关闭按钮（对应图片框EX4）
        btnClose.Text = "X"
        btnClose.Location = New Drawing.Point(Me.Width - 35, 5)
        btnClose.Size = New Drawing.Size(30, 25)
        btnClose.BackColor = Drawing.Color.Red
        btnClose.ForeColor = Drawing.Color.White
        btnClose.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnClose)

        ' 添加修改分组框（对应添加修改_分组框）
        grpAddModify.Text = "添加会员信息"
        grpAddModify.Size = New Drawing.Size(490, 420)
        grpAddModify.Location = New Drawing.Point((Me.Width - grpAddModify.Width) \ 2, (Me.Height - grpAddModify.Height) \ 2)
        Me.Controls.Add(grpAddModify)

        ' 会员编码（对应编辑框_会员编码）
        AddLabelAndTextBox(grpAddModify, "会员编码：", txtMemberCode, 16, 30, 160)
        txtMemberCode.ReadOnly = True

        ' 会员姓名（对应编辑框_会员姓名）
        AddLabelAndTextBox(grpAddModify, "会员姓名：", txtMemberName, 16, 70, 160)

        ' 联系电话（对应编辑框_联系电话）
        AddLabelAndTextBox(grpAddModify, "联系电话：", txtPhone, 16, 110, 160)

        ' 出生日期（对应编辑框_出生日期）- 易语言中是编辑框，点击弹出日期框
        AddLabelAndTextBox(grpAddModify, "出生日期：", txtBirthday, 16, 150, 160)
        txtBirthday.ReadOnly = True
        AddHandler txtBirthday.Click, AddressOf txtBirthday_Click

        ' 省份（对应组合框_省份）
        AddLabel(grpAddModify, "省份：", 280, 150)
        cmbProvince.Location = New Drawing.Point(320, 147)
        cmbProvince.Size = New Drawing.Size(125, 25)
        cmbProvince.DropDownStyle = ComboBoxStyle.DropDownList
        grpAddModify.Controls.Add(cmbProvince)
        AddHandler cmbProvince.SelectedIndexChanged, AddressOf cmbProvince_SelectedIndexChanged

        ' 市区（对应组合框_市区）
        AddLabel(grpAddModify, "城市：", 280, 110)
        cmbCity.Location = New Drawing.Point(320, 107)
        cmbCity.Size = New Drawing.Size(125, 25)
        cmbCity.DropDownStyle = ComboBoxStyle.DropDownList
        grpAddModify.Controls.Add(cmbCity)
        AddHandler cmbCity.SelectedIndexChanged, AddressOf cmbCity_SelectedIndexChanged

        ' 县区（对应组合框_县区）
        AddLabel(grpAddModify, "区县：", 445, 110)
        cmbDistrict.Location = New Drawing.Point(320, 147)
        ' 县区位置调整到省份下方第二行区域
        cmbDistrict.Location = New Drawing.Point(320, 147)
        ' 实际布局：省/市/县排成一排
        cmbProvince.Location = New Drawing.Point(280, 147)
        cmbProvince.Size = New Drawing.Size(110, 25)
        cmbCity.Location = New Drawing.Point(395, 147)
        cmbCity.Size = New Drawing.Size(110, 25)
        cmbDistrict.Location = New Drawing.Point(280, 187)
        cmbDistrict.Size = New Drawing.Size(110, 25)
        grpAddModify.Controls.Add(cmbDistrict)
        AddHandler cmbDistrict.SelectedIndexChanged, AddressOf cmbDistrict_SelectedIndexChanged

        ' 联系地址（对应编辑框_用户联系地址）
        AddLabelAndTextBox(grpAddModify, "联系地址：", txtAddress, 16, 230, 460)

        ' 会员结料（对应编辑框_会员结料）
        AddLabelAndTextBox(grpAddModify, "会员结料(g)：", txtJieLiao, 16, 270, 150)
        txtJieLiao.Text = "0.000"

        ' 会员结账（对应编辑框_会员结账）
        AddLabelAndTextBox(grpAddModify, "会员结账(元)：", txtJieZhang, 16, 310, 150)
        txtJieZhang.Text = "0.00"

        ' 备注（对应编辑框_备注）
        AddLabelAndTextBox(grpAddModify, "备注：", txtRemarks, 16, 350, 460)

        ' 保存/修改按钮（对应按钮EX1）
        btnSubmit.Text = "保存"
        btnSubmit.Location = New Drawing.Point(160, 390)
        btnSubmit.Size = New Drawing.Size(100, 30)
        grpAddModify.Controls.Add(btnSubmit)

        ' 重置按钮（对应按钮EX2）
        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(280, 390)
        btnReset.Size = New Drawing.Size(100, 30)
        grpAddModify.Controls.Add(btnReset)
    End Sub

    Private Sub AddLabel(parent As Control, text As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Drawing.Point(x, y)
        lbl.AutoSize = True
        parent.Controls.Add(lbl)
    End Sub

    Private Sub AddLabelAndTextBox(parent As Control, labelText As String, txtBox As TextBox, x As Integer, y As Integer, width As Integer)
        AddLabel(parent, labelText, x, y)
        txtBox.Location = New Drawing.Point(x + 80, y)
        txtBox.Size = New Drawing.Size(width, 25)
        parent.Controls.Add(txtBox)
    End Sub

    ' ========== _窗口_会员添加修改_创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        selectedMemberCode = ""

        ' 对应易语言：如果主窗口高级表格1有选中行，获取选中会员信息内容
        ' VB.NET中通过构造函数参数传入

        If String.IsNullOrEmpty(selectedMemberCode) Then
            ' 添加模式
            grpAddModify.Text = "添加会员信息"
            btnSubmit.Text = "保存"
            txtMemberCode.Text = GenerateMemberCode()
            txtPhone.ReadOnly = False
            txtMemberName.Text = ""
            txtPhone.Text = ""
            txtAddress.Text = ""
            txtBirthday.Text = DateTime.Now.ToString("yyyy-MM-dd")
        Else
            ' 修改模式
            grpAddModify.Text = "编辑会员信息"
            btnSubmit.Text = "修改"
            txtMemberCode.Text = selectedMemberCode
            txtPhone.ReadOnly = True
            ' 修改模式下从主窗口表格获取数据（VB.NET通过构造函数传入memberCode，再从数据库查询）
            LoadMemberDataForEdit()
        End If

        txtJieLiao.Text = "0.000"
        txtJieZhang.Text = "0.00"
        txtRemarks.Text = ""

        ' 初始化省市区组合框
        cmbProvince.Items.Clear()
        cmbCity.Items.Clear()
        cmbDistrict.Items.Clear()
        LoadProvinceList()
    End Sub

    ' ========== 加载省份列表（对应易语言_窗口_会员添加修改_创建完毕中省份加载部分） ==========
    Private Sub LoadProvinceList()
        Try
            Dim sql As String = "SELECT id,name FROM `erp_city` where pid='0' and level='1' order by id ASC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbProvince.Items.Clear()
            For Each row As DataRow In dt.Rows
                Dim provinceId As String = SafeString(row("id"))
                Dim provinceName As String = DatabaseModule.GBKToUTF8(SafeString(row("name")))
                cmbProvince.Items.Add(New ComboBoxItem(provinceId, provinceName))
            Next
            If cmbProvince.Items.Count > 0 Then cmbProvince.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    ' ========== _组合框_省份_内容被改变 ==========
    Private Sub cmbProvince_SelectedIndexChanged(sender As Object, e As EventArgs)
        cmbCity.Items.Clear()
        cmbDistrict.Items.Clear()
        txtAddress.Text = ""

        If cmbProvince.SelectedIndex < 0 Then Return

        Dim provinceItem As ComboBoxItem = DirectCast(cmbProvince.SelectedItem, ComboBoxItem)
        Try
            Dim sql As String = "SELECT id,name FROM erp_city where pid='" & provinceItem.ID & "' and level='2' order by id ASC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbCity.Items.Clear()
            For Each row As DataRow In dt.Rows
                Dim cityId As String = SafeString(row("id"))
                Dim cityName As String = DatabaseModule.GBKToUTF8(SafeString(row("name")))
                cmbCity.Items.Add(New ComboBoxItem(cityId, cityName))
            Next
            If cmbCity.Items.Count > 0 Then cmbCity.SelectedIndex = 0

            ' 更新地址：省份名称
            txtAddress.Text = provinceItem.Text
        Catch ex As Exception
        End Try
    End Sub

    ' ========== _组合框_市区_内容被改变 ==========
    Private Sub cmbCity_SelectedIndexChanged(sender As Object, e As EventArgs)
        cmbDistrict.Items.Clear()

        If cmbCity.SelectedIndex < 0 Then Return

        Dim cityItem As ComboBoxItem = DirectCast(cmbCity.SelectedItem, ComboBoxItem)
        Try
            Dim sql As String = "SELECT id,name FROM erp_city where pid='" & cityItem.ID & "' and level='3' order by id ASC"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)
            cmbDistrict.Items.Clear()
            For Each row As DataRow In dt.Rows
                Dim districtId As String = SafeString(row("id"))
                Dim districtName As String = DatabaseModule.GBKToUTF8(SafeString(row("name")))
                cmbDistrict.Items.Add(New ComboBoxItem(districtId, districtName))
            Next
            If cmbDistrict.Items.Count > 0 Then cmbDistrict.SelectedIndex = 0

            ' 更新地址：省份名称 + 市区名称
            If cmbProvince.SelectedIndex >= 0 Then
                txtAddress.Text = DirectCast(cmbProvince.SelectedItem, ComboBoxItem).Text & cityItem.Text
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== _组合框_县区_内容被改变 ==========
    Private Sub cmbDistrict_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmbDistrict.SelectedIndex < 0 Then Return

        Dim districtItem As ComboBoxItem = DirectCast(cmbDistrict.SelectedItem, ComboBoxItem)

        ' 更新地址：省份名称 + 市区名称 + 县区名称
        If cmbProvince.SelectedIndex >= 0 AndAlso cmbCity.SelectedIndex >= 0 Then
            txtAddress.Text = DirectCast(cmbProvince.SelectedItem, ComboBoxItem).Text &
                              DirectCast(cmbCity.SelectedItem, ComboBoxItem).Text &
                              districtItem.Text
        End If
    End Sub

    ' ========== _编辑框_出生日期_鼠标左键单击 ==========
    Private Sub txtBirthday_Click(sender As Object, e As EventArgs)
        ' 对应易语言：弹出日期框
        Using dlg As New Form()
            dlg.Text = "选择日期"
            dlg.Size = New Drawing.Size(250, 200)
            dlg.StartPosition = FormStartPosition.CenterParent
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog
            dlg.MaximizeBox = False
            dlg.MinimizeBox = False

            Dim dtp As New DateTimePicker()
            dtp.Format = DateTimePickerFormat.Short
            dtp.Location = New Drawing.Point(30, 30)
            dtp.Size = New Drawing.Size(180, 25)
            If Not String.IsNullOrEmpty(txtBirthday.Text) Then
                Try
                    dtp.Value = DateTime.Parse(txtBirthday.Text)
                Catch
                    dtp.Value = DateTime.Now
                End Try
            End If
            dlg.Controls.Add(dtp)

            Dim btnOK As New Button()
            btnOK.Text = "确定"
            btnOK.Location = New Drawing.Point(50, 80)
            btnOK.Size = New Drawing.Size(80, 30)
            dlg.Controls.Add(btnOK)

            Dim btnCancel As New Button()
            btnCancel.Text = "取消"
            btnCancel.Location = New Drawing.Point(140, 80)
            btnCancel.Size = New Drawing.Size(80, 30)
            dlg.Controls.Add(btnCancel)

            AddHandler btnOK.Click, Sub(s2, e2)
                                        txtBirthday.Text = dtp.Value.ToString("yyyy-MM-dd")
                                        dlg.DialogResult = DialogResult.OK
                                        dlg.Close()
                                    End Sub
            AddHandler btnCancel.Click, Sub(s2, e2)
                                         dlg.DialogResult = DialogResult.Cancel
                                         dlg.Close()
                                     End Sub

            dlg.ShowDialog(Me)
        End Using
    End Sub

    ' ========== _窗口_会员添加修改_尺寸被改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        ' 对应易语言：分组框居中
        grpAddModify.Left = (Me.Width - grpAddModify.Width) \ 2
        grpAddModify.Top = (Me.Height - grpAddModify.Height) \ 2
    End Sub

    ' ========== _图片框EX4_鼠标左键单击（关闭按钮） ==========
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ' ========== _按钮EX1_鼠标左键单击（保存/修改按钮） ==========
    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        If btnSubmit.Text = "保存" Then
            MemberAddSaveData()
        End If

        If btnSubmit.Text = "修改" Then
            MemberModifySaveData()
        End If
    End Sub

    ' ========== _按钮EX2_鼠标左键单击（重置按钮） ==========
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        ' 对应易语言：重新执行_窗口_会员添加修改_创建完毕
        Form_Load(Me, EventArgs.Empty)
    End Sub

    ' ========== 加载会员数据（修改模式） ==========
    Private Sub LoadMemberDataForEdit()
        Try
            Dim sql As String = "SELECT * FROM xipunum_erp_member WHERE customer_code='" & SafeSQL(selectedMemberCode) & "' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql)

            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                txtMemberCode.Text = SafeString(row("customer_code"))
                txtMemberName.Text = DatabaseModule.GBKToUTF8(SafeString(row("name")))
                txtPhone.Text = SafeString(row("tel"))
                txtAddress.Text = DatabaseModule.GBKToUTF8(SafeString(row("dizhi")))
                txtBirthday.Text = SafeString(row("shengri"))
                If String.IsNullOrEmpty(txtBirthday.Text) Then
                    txtBirthday.Text = DateTime.Now.ToString("yyyy-MM-dd")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("加载会员数据失败：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ========== 生成会员编码（对应易语言编辑框_会员编码.内容赋值） ==========
    Private Function GenerateMemberCode() As String
        ' 对应易语言：编码转换(到字节集(到文本(相加("HY", 到文本(时间_格式化(取现行时间(), "yyyyMMdd", "hhmmss", 真)))), #编码_GBK, #编码_UTF_8))
        Dim raw As String = "HY" & DateTime.Now.ToString("yyyyMMdd") & DateTime.Now.ToString("HHmmss")
        Return DatabaseModule.UTF8ToGBK(raw)
    End Function

    ' ========== _会员添加_保存数据 ==========
    Private Sub MemberAddSaveData()
        Dim infoMemberCode As String = txtMemberCode.Text
        Dim infoMemberName As String = txtMemberName.Text
        Dim infoPhone As String = txtPhone.Text
        Dim infoBirthday As String = txtBirthday.Text
        Dim infoAddress As String = txtAddress.Text
        Dim infoJieLiao As String = txtJieLiao.Text
        Dim infoJieZhang As String = txtJieZhang.Text
        Dim infoRemarks As String = txtRemarks.Text

        Dim globalInfoOperationDate As String = DateTime.Now.ToString("yyyy-MM-dd") & " " & DateTime.Now.ToString("HH:mm:ss")
        Dim globalInfoOperationAccount As String = UserAccount
        Dim globalLogSaveContent As String = ""

        ' 对应易语言：如果(不等于(编辑框_联系电话.内容, ""))
        If infoPhone <> "" Then
            ' 对应易语言：如果(等于(取文本左边(编辑框_联系电话.内容, 1), "1"))
            If infoPhone.Substring(0, 1) = "1" Then
                ' 对应易语言：如果(等于(取文本长度(编辑框_联系电话.内容), 11))
                If infoPhone.Length = 11 Then
                    ' 对应易语言：SELECT * FROM xipunum_erp_member where name ='...' and tel ='...'
                    Dim checkSql As String = "SELECT * FROM xipunum_erp_member where name ='" & SafeSQL(infoMemberName) & "' and tel ='" & SafeSQL(infoPhone) & "'"
                    Dim checkDt As DataTable = DatabaseModule.ExecuteQuery(checkSql, MySQL_Read)

                    ' 对应易语言：如果(等于(会员电话是否存在, 0))
                    If checkDt.Rows.Count = 0 Then
                        ' 对应易语言：增加记录 INSERT
                        Dim insertSql As String = "INSERT INTO xipunum_erp_member (customer_code, name, tel, shengri, dizhi, cjuser, creationtime) VALUES ('" &
                            SafeSQL(infoMemberCode) & "','" & SafeSQL(infoMemberName) & "','" & SafeSQL(infoPhone) & "','" &
                            SafeSQL(infoBirthday) & "','" & SafeSQL(infoAddress) & "','" & SafeSQL(globalInfoOperationAccount) & "','" &
                            SafeSQL(globalInfoOperationDate) & "')"
                        DatabaseModule.ExecuteCommand(insertSql)

                        ' 对应易语言：SELECT id,customer_code FROM xipunum_erp_member where customer_code='...' order by id ASC LIMIT 1
                        Dim getMemberSql As String = "SELECT id,customer_code FROM xipunum_erp_member where customer_code='" & SafeSQL(infoMemberCode) & "' order by id ASC  LIMIT 1"
                        Dim memberDt As DataTable = DatabaseModule.ExecuteQuery(getMemberSql, MySQL_Read)

                        Dim infoCustomerID As String = ""
                        If memberDt.Rows.Count > 0 Then
                            infoCustomerID = SafeString(memberDt.Rows(0)("id"))
                        End If

                        ' 对应易语言：UPDATE xipunum_erp_member SET memberid= '1' + 右补零6位ID
                        Dim paddedID As String = ("00000000" & infoCustomerID)
                        Dim memberID As String = "1" & paddedID.Substring(paddedID.Length - 6)
                        Dim updateMemberIdSql As String = "UPDATE xipunum_erp_member SET memberid= '" & memberID & "' WHERE customer_code ='" & SafeSQL(infoMemberCode) & "' LIMIT 1"
                        DatabaseModule.ExecuteCommand(updateMemberIdSql)

                        MessageBox.Show("会员:" & txtMemberName.Text & "添加成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        ' 对应易语言：日志保存
                        globalLogSaveContent = ""
                        globalLogSaveContent = "账户:" & UserAccount & " 添加会员编码：" & txtMemberCode.Text & " 会员:" & txtMemberName.Text & " 联系电话:" & txtPhone.Text

                        Dim logSql As String = "INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('添加','添加会员数据','" &
                            SafeSQL(globalLogSaveContent) & "','" & SafeSQL(globalInfoOperationAccount) & "','" & SafeSQL(globalInfoOperationDate) & "')"
                        DatabaseModule.ExecuteCommand(logSql)

                        ' 对应易语言：_会员信息_存欠数据()
                        MemberStoreDebtData()

                        Me.Close()

                        ' 对应易语言：如果真(等于(全局_首页查询栏目文本, "会员列表"))
                        ' 在VB.NET中通过主窗体刷新实现
                        Return
                    Else
                        MessageBox.Show("此会员信息已存在！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        txtPhone.Focus()
                        Return
                    End If
                Else
                    MessageBox.Show("请输入正确的手机号码", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtPhone.Text = ""
                    txtPhone.Focus()
                    Return
                End If
            Else
                MessageBox.Show("请输入正确的手机号码", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPhone.Text = ""
                txtPhone.Focus()
                Return
            End If
        Else
            MessageBox.Show("会员联系电话不能为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPhone.Focus()
            Return
        End If

        Me.Close()
    End Sub

    ' ========== _会员修改_修改数据 ==========
    Private Sub MemberModifySaveData()
        Dim infoMemberCode As String = txtMemberCode.Text
        Dim infoMemberName As String = txtMemberName.Text
        Dim infoPhone As String = txtPhone.Text
        Dim infoBirthday As String = txtBirthday.Text
        Dim infoAddress As String = txtAddress.Text
        Dim infoJieLiao As String = txtJieLiao.Text
        Dim infoJieZhang As String = txtJieZhang.Text
        Dim infoRemarks As String = txtRemarks.Text

        Dim globalInfoOperationDate As String = DateTime.Now.ToString("yyyy-MM-dd") & " " & DateTime.Now.ToString("HH:mm:ss")
        Dim globalInfoOperationAccount As String = UserAccount
        Dim globalLogSaveContent As String = ""

        ' 对应易语言：UPDATE xipunum_erp_member SET name= '...',shengri= '...',dizhi= '...',cjuser= '...',updatetime= '...' WHERE customer_code ='...' LIMIT 1
        Dim updateSql As String = "UPDATE xipunum_erp_member SET name= '" & SafeSQL(infoMemberName) & "',shengri= '" & SafeSQL(infoBirthday) &
            "',dizhi= '" & SafeSQL(infoAddress) & "',cjuser= '" & SafeSQL(globalInfoOperationAccount) &
            "',updatetime= '" & SafeSQL(globalInfoOperationDate) & "'  WHERE customer_code ='" & SafeSQL(infoMemberCode) & "' LIMIT 1"
        DatabaseModule.ExecuteCommand(updateSql)

        MessageBox.Show("会员:" & txtMemberName.Text & "修改成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' 对应易语言：日志保存
        globalLogSaveContent = ""
        globalLogSaveContent = "账户:" & UserAccount & " 修改会员编码：" & txtMemberCode.Text & " 会员:" & txtMemberName.Text & " 联系电话:" & txtPhone.Text

        Dim logSql As String = "INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('修改','修改会员数据','" &
            SafeSQL(globalLogSaveContent) & "','" & SafeSQL(globalInfoOperationAccount) & "','" & SafeSQL(globalInfoOperationDate) & "')"
        DatabaseModule.ExecuteCommand(logSql)

        ' 对应易语言：_会员信息_存欠数据()
        MemberStoreDebtData()

        ' 对应易语言：查询会员存欠数据并更新主窗口表格
        Dim memberCunQianJieLiao As String = ""
        Dim memberCunQianJieYuan As String = ""

        ' 对应易语言中完整的会员存欠查询SQL
        Dim cunQianSql As String = "SELECT a.memberid as amemberid,a.customer_code AS acustomer_code,a.NAME AS aname,a.tel AS atel," &
            "ROUND(IFNULL(c.cun_number, 0) - IFNULL(c.qian_number, 0), 3) AS jieyuliao," &
            "ROUND(IFNULL(d.cun_number, 0) - IFNULL(d.qian_number, 0), 2) AS jieyuyuan," &
            "a.shengri AS ashengri,a.dizhi AS adizhi,a.cjuser AS acjuser,a.creationtime AS acreationtime,b.NAME AS bname " &
            "FROM xipunum_erp_member AS a " &
            "INNER JOIN xipunum_erp_user AS b ON b.USER = a.cjuser " &
            "LEFT JOIN ( SELECT customer_code, SUM(CASE WHEN cunqu = '存' AND type = '料' THEN number ELSE 0 END) AS cun_number, SUM(CASE WHEN cunqu = '欠' AND type = '料' THEN number ELSE 0 END) AS qian_number FROM xipunum_erp_member_cq WHERE kufang in (" & UserShopPermission & ") AND type = '料' GROUP BY customer_code ) AS c ON c.customer_code = a.customer_code " &
            "LEFT JOIN ( SELECT customer_code, SUM(CASE WHEN cunqu = '存' AND type = '元' THEN number ELSE 0 END) AS cun_number, SUM(CASE WHEN cunqu = '欠' AND type = '元' THEN number ELSE 0 END) AS qian_number FROM xipunum_erp_member_cq WHERE kufang in (" & UserShopPermission & ") AND type = '元' GROUP BY customer_code ) AS d ON d.customer_code = a.customer_code " &
            "WHERE a.customer_code IN ( SELECT DISTINCT customer_code FROM (SELECT customer_code FROM xipunum_erp_outbound_order WHERE cjuser IN " & GlobalViewSQL & " UNION SELECT customer_code FROM xipunum_erp_retreat_order WHERE cjuser IN " & GlobalViewSQL & " UNION SELECT customer_code FROM xipunum_erp_presale_order WHERE cjuser IN " & GlobalViewSQL & " UNION SELECT customer_code FROM xipunum_erp_member WHERE cjuser IN " & GlobalViewSQL & " ) AS combined WHERE customer_code != '' AND customer_code IS NOT NULL ) " &
            "AND a.customer_code='" & SafeSQL(infoMemberCode) & "' ORDER BY a.customer_code asc"

        Dim cunQianDt As DataTable = DatabaseModule.ExecuteQuery(cunQianSql, MySQL_Read)
        If cunQianDt.Rows.Count > 0 Then
            memberCunQianJieLiao = SafeString(cunQianDt.Rows(0)("jieyuliao"))
            memberCunQianJieYuan = SafeString(cunQianDt.Rows(0)("jieyuyuan"))
        End If

        ' 对应易语言：更新主窗口高级表格数据
        ' VB.NET中通过调用主窗口方法或事件来实现
        ' 如果有主窗口引用，更新对应行数据
        If MainForm IsNot Nothing Then
            ' 更新主窗口表格中的会员数据
            MainForm.UpdateMemberGridRow(selectedMemberCode, infoMemberName, memberCunQianJieLiao, memberCunQianJieYuan, infoBirthday, infoAddress)
        End If

        Me.Close()
    End Sub

    ' ========== _会员信息_存欠数据 ==========
    Private Sub MemberStoreDebtData()
        Dim infoMemberCode As String = txtMemberCode.Text
        Dim storeDebtLiaoStatus As String = ""    ' 存欠料状态信息
        Dim storeDebtLiaoData As String = ""       ' 存欠料数据信息
        Dim storeDebtYuanStatus As String = ""     ' 存欠款状态信息
        Dim storeDebtYuanData As String = ""       ' 存欠款数据信息

        Dim globalInfoOperationDate As String = DateTime.Now.ToString("yyyy-MM-dd") & " " & DateTime.Now.ToString("HH:mm:ss")
        Dim globalInfoOperationAccount As String = UserAccount

        ' 对应易语言：如果(大于(到数值(编辑框_会员结料.内容), 0))
        Dim jieLiaoValue As Decimal = SafeDecimal(txtJieLiao.Text)
        If jieLiaoValue > 0 Then
            storeDebtLiaoData = txtJieLiao.Text
            storeDebtLiaoStatus = "存"
        Else
            storeDebtLiaoData = CStr(Math.Abs(jieLiaoValue))
            storeDebtLiaoStatus = "欠"
        End If

        ' 对应易语言：如果(大于(到数值(编辑框_会员结账.内容), 0))
        Dim jieZhangValue As Decimal = SafeDecimal(txtJieZhang.Text)
        If jieZhangValue > 0 Then
            storeDebtYuanData = txtJieZhang.Text
            storeDebtYuanStatus = "存"
        Else
            storeDebtYuanData = CStr(Math.Abs(jieZhangValue))
            storeDebtYuanStatus = "欠"
        End If

        ' 对应易语言：如果真(不等于(到数值(存欠料数据信息), 0))
        If SafeDecimal(storeDebtLiaoData) <> 0 Then
            Dim insertLiaoSql As String = "INSERT INTO xipunum_erp_member_cq (customer_code, cunqu, type, number, remarks, kufang, cjuser, creationtime) VALUES ('" &
                SafeSQL(infoMemberCode) & "','" & storeDebtLiaoStatus & "','料','" & SafeSQL(storeDebtLiaoData) & "','" &
                SafeSQL(txtRemarks.Text) & "','" & SafeSQL(UserDepartment) & "','" & SafeSQL(globalInfoOperationAccount) & "','" &
                SafeSQL(globalInfoOperationDate) & "')"
            DatabaseModule.ExecuteCommand(insertLiaoSql)

            Dim globalLogSaveContent As String = ""
            globalLogSaveContent = "账户:" & UserAccount & " 操作 会员编码：" & txtMemberCode.Text & " 会员：" & txtMemberName.Text & " 存储（" & txtJieLiao.Text & ")g 料"

            Dim logLiaoSql As String = "INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('添加','会员存储数据','" &
                SafeSQL(globalLogSaveContent) & "','" & SafeSQL(globalInfoOperationAccount) & "','" & SafeSQL(globalInfoOperationDate) & "')"
            DatabaseModule.ExecuteCommand(logLiaoSql)
        End If

        ' 对应易语言：如果真(不等于(到数值(存欠款数据信息), 0))
        If SafeDecimal(storeDebtYuanData) <> 0 Then
            Dim insertYuanSql As String = "INSERT INTO xipunum_erp_member_cq (customer_code, cunqu, type, number, remarks, kufang, cjuser, creationtime) VALUES ('" &
                SafeSQL(infoMemberCode) & "','" & storeDebtYuanStatus & "','元','" & SafeSQL(storeDebtYuanData) & "','" &
                SafeSQL(txtRemarks.Text) & "','" & SafeSQL(UserDepartment) & "','" & SafeSQL(globalInfoOperationAccount) & "','" &
                SafeSQL(globalInfoOperationDate) & "')"
            DatabaseModule.ExecuteCommand(insertYuanSql)

            Dim globalLogSaveContent As String = "账户:" & UserAccount & " 操作 会员编码：" & txtMemberCode.Text & " 会员：" & txtMemberName.Text & " 存储（" & txtJieZhang.Text & ")g 元"

            Dim logYuanSql As String = "INSERT INTO xipunum_erp_xitong_log (type, title, conter, user, creationtime) VALUES ('添加','会员存储数据','" &
                SafeSQL(globalLogSaveContent) & "','" & SafeSQL(globalInfoOperationAccount) & "','" & SafeSQL(globalInfoOperationDate) & "')"
            DatabaseModule.ExecuteCommand(logYuanSql)
        End If
    End Sub

    ' ========== ComboBox项类 ==========
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

End Class

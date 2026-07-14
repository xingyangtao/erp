' ============================================================================
' 绩效信息添加窗口
' 功能: 添加/编辑绩效配置
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class PerformanceAddForm
    Inherits System.Windows.Forms.Form

    Private cmbType As New ComboBox()
    Private cmbCategory As New ComboBox()
    Private cmbPl As New ComboBox()
    Private cmbJszd As New ComboBox()
    Private cmbJsfw As New ComboBox()
    Private cmbDjs As New ComboBox()
    Private txtMin1 As New TextBox()
    Private txtCs1 As New TextBox()
    Private txtMin2 As New TextBox()
    Private txtCs2 As New TextBox()
    Private txtMin3 As New TextBox()
    Private txtCs3 As New TextBox()
    Private txtMin4 As New TextBox()
    Private txtCs4 As New TextBox()
    Private txtMin5 As New TextBox()
    Private txtCs5 As New TextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnCancel As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "添加绩效配置"
        Me.Size = New Drawing.Size(600, 500)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblType As New Label()
        lblType.Text = "岗位："
        lblType.Location = New Drawing.Point(30, 20)
        lblType.AutoSize = True
        Me.Controls.Add(lblType)

        cmbType.Location = New Drawing.Point(100, 17)
        cmbType.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbType)

        Dim lblCategory As New Label()
        lblCategory.Text = "品类："
        lblCategory.Location = New Drawing.Point(270, 20)
        lblCategory.AutoSize = True
        Me.Controls.Add(lblCategory)

        cmbCategory.Location = New Drawing.Point(320, 17)
        cmbCategory.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbCategory)

        Dim lblPl As New Label()
        lblPl.Text = "批零："
        lblPl.Location = New Drawing.Point(30, 60)
        lblPl.AutoSize = True
        Me.Controls.Add(lblPl)

        cmbPl.Location = New Drawing.Point(100, 57)
        cmbPl.Size = New Drawing.Size(100, 25)
        cmbPl.Items.AddRange(New String() {"批发", "零售"})
        cmbPl.SelectedIndex = 0
        Me.Controls.Add(cmbPl)

        Dim lblJszd As New Label()
        lblJszd.Text = "计算字段："
        lblJszd.Location = New Drawing.Point(220, 60)
        lblJszd.AutoSize = True
        Me.Controls.Add(lblJszd)

        cmbJszd.Location = New Drawing.Point(300, 57)
        cmbJszd.Size = New Drawing.Size(100, 25)
        cmbJszd.Items.AddRange(New String() {"总数量", "总重量", "总金额", "单件重"})
        cmbJszd.SelectedIndex = 0
        Me.Controls.Add(cmbJszd)

        Dim lblJsfw As New Label()
        lblJsfw.Text = "计算范围："
        lblJsfw.Location = New Drawing.Point(30, 100)
        lblJsfw.AutoSize = True
        Me.Controls.Add(lblJsfw)

        cmbJsfw.Location = New Drawing.Point(100, 97)
        cmbJsfw.Size = New Drawing.Size(100, 25)
        cmbJsfw.Items.AddRange(New String() {"店铺", "岗位", "个人"})
        cmbJsfw.SelectedIndex = 0
        Me.Controls.Add(cmbJsfw)

        Dim lblDjs As New Label()
        lblDjs.Text = "档次数："
        lblDjs.Location = New Drawing.Point(220, 100)
        lblDjs.AutoSize = True
        Me.Controls.Add(lblDjs)

        cmbDjs.Location = New Drawing.Point(300, 97)
        cmbDjs.Size = New Drawing.Size(80, 25)
        cmbDjs.Items.AddRange(New String() {"1", "2", "3", "4", "5"})
        cmbDjs.SelectedIndex = 0
        Me.Controls.Add(cmbDjs)

        ' 档次配置
        Dim lblTier As New Label()
        lblTier.Text = "档次配置（最小值/参数值）："
        lblTier.Location = New Drawing.Point(30, 140)
        lblTier.AutoSize = True
        Me.Controls.Add(lblTier)

        txtMin1.Location = New Drawing.Point(30, 170)
        txtMin1.Size = New Drawing.Size(80, 25)
        txtMin1.Text = "0"
        Me.Controls.Add(txtMin1)

        txtCs1.Location = New Drawing.Point(120, 170)
        txtCs1.Size = New Drawing.Size(80, 25)
        txtCs1.Text = "0"
        Me.Controls.Add(txtCs1)

        txtMin2.Location = New Drawing.Point(30, 200)
        txtMin2.Size = New Drawing.Size(80, 25)
        txtMin2.Text = "0"
        Me.Controls.Add(txtMin2)

        txtCs2.Location = New Drawing.Point(120, 200)
        txtCs2.Size = New Drawing.Size(80, 25)
        txtCs2.Text = "0"
        Me.Controls.Add(txtCs2)

        txtMin3.Location = New Drawing.Point(30, 230)
        txtMin3.Size = New Drawing.Size(80, 25)
        txtMin3.Text = "0"
        Me.Controls.Add(txtMin3)

        txtCs3.Location = New Drawing.Point(120, 230)
        txtCs3.Size = New Drawing.Size(80, 25)
        txtCs3.Text = "0"
        Me.Controls.Add(txtCs3)

        txtMin4.Location = New Drawing.Point(30, 260)
        txtMin4.Size = New Drawing.Size(80, 25)
        txtMin4.Text = "0"
        Me.Controls.Add(txtMin4)

        txtCs4.Location = New Drawing.Point(120, 260)
        txtCs4.Size = New Drawing.Size(80, 25)
        txtCs4.Text = "0"
        Me.Controls.Add(txtCs4)

        txtMin5.Location = New Drawing.Point(30, 290)
        txtMin5.Size = New Drawing.Size(80, 25)
        txtMin5.Text = "0"
        Me.Controls.Add(txtMin5)

        txtCs5.Location = New Drawing.Point(120, 290)
        txtCs5.Size = New Drawing.Size(80, 25)
        txtCs5.Text = "0"
        Me.Controls.Add(txtCs5)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(100, 350)
        btnSave.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnSave)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(220, 350)
        btnCancel.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadTypeList()
        LoadCategoryList()
    End Sub

    Private Sub LoadTypeList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, title FROM xipunum_erp_type WHERE superior>'0' AND type='商铺' ORDER BY id")
            cmbType.Items.Clear()
            For Each row As DataRow In dt.Rows
                cmbType.Items.Add(SafeString(row("title")))
            Next
            If cmbType.Items.Count > 0 Then cmbType.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadCategoryList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, title FROM xipunum_erp_category ORDER BY id")
            cmbCategory.Items.Clear()
            For Each row As DataRow In dt.Rows
                cmbCategory.Items.Add(SafeString(row("title")))
            Next
            If cmbCategory.Items.Count > 0 Then cmbCategory.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Dim sql As String = $"INSERT INTO xipunum_erp_performance (type_id, category_id, pl, jszd, jsfw, djs, min1, cs1, min2, cs2, min3, cs3, min4, cs4, min5, cs5, cjuser, creationtime) VALUES ('{cmbType.SelectedIndex}', '{cmbCategory.SelectedIndex}', '{cmbPl.SelectedItem}', '{cmbJszd.SelectedItem}', '{cmbJsfw.SelectedItem}', '{cmbDjs.SelectedItem}', '{txtMin1.Text}', '{txtCs1.Text}', '{txtMin2.Text}', '{txtCs2.Text}', '{txtMin3.Text}', '{txtCs3.Text}', '{txtMin4.Text}', '{txtCs4.Text}', '{txtMin5.Text}', '{txtCs5.Text}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
            ExecuteCommand(sql)
            ShowSuccess("保存成功！")
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

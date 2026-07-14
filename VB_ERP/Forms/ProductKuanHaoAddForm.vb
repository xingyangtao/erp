' ============================================================================
' 商品信息添加款号窗口
' 功能: 商品信息添加时的款号快速创建
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class ProductKuanHaoAddForm
    Inherits System.Windows.Forms.Form

    Private txtName As New TextBox()
    Private txtKuanHao As New TextBox()
    Private cmbCategory As New ComboBox()
    Private cmbSpec As New ComboBox()
    Private txtCaiZhi As New TextBox()
    Private txtImageUrl As New TextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnCancel As New Button()

    Public Property KuanHao As String = ""
    Public Overloads Property ProductName As String = ""
    Public Property ImageUrl As String = ""

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "商品信息添加款号"
        Me.Size = New Drawing.Size(450, 350)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblName As New Label()
        lblName.Text = "名称："
        lblName.Location = New Drawing.Point(30, 30)
        lblName.AutoSize = True
        Me.Controls.Add(lblName)

        txtName.Location = New Drawing.Point(100, 27)
        txtName.Size = New Drawing.Size(250, 25)
        Me.Controls.Add(txtName)

        Dim lblCategory As New Label()
        lblCategory.Text = "品类："
        lblCategory.Location = New Drawing.Point(30, 70)
        lblCategory.AutoSize = True
        Me.Controls.Add(lblCategory)

        cmbCategory.Location = New Drawing.Point(100, 67)
        cmbCategory.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbCategory)

        Dim lblSpec As New Label()
        lblSpec.Text = "规格："
        lblSpec.Location = New Drawing.Point(30, 110)
        lblSpec.AutoSize = True
        Me.Controls.Add(lblSpec)

        cmbSpec.Location = New Drawing.Point(100, 107)
        cmbSpec.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbSpec)

        Dim lblCaiZhi As New Label()
        lblCaiZhi.Text = "材质："
        lblCaiZhi.Location = New Drawing.Point(30, 150)
        lblCaiZhi.AutoSize = True
        Me.Controls.Add(lblCaiZhi)

        txtCaiZhi.Location = New Drawing.Point(100, 147)
        txtCaiZhi.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(txtCaiZhi)

        Dim lblImageUrl As New Label()
        lblImageUrl.Text = "图片地址："
        lblImageUrl.Location = New Drawing.Point(30, 190)
        lblImageUrl.AutoSize = True
        Me.Controls.Add(lblImageUrl)

        txtImageUrl.Location = New Drawing.Point(100, 187)
        txtImageUrl.Size = New Drawing.Size(250, 25)
        Me.Controls.Add(txtImageUrl)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(100, 240)
        btnSave.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnSave)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(220, 240)
        btnCancel.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadCategoryList()
        LoadSpecList()
    End Sub

    Private Sub LoadCategoryList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, title FROM xipunum_erp_category ORDER BY id")
            cmbCategory.Items.Clear()
            cmbCategory.Items.Add(New ComboBoxItem("", "请选择"))
            For Each row As DataRow In dt.Rows
                cmbCategory.Items.Add(New ComboBoxItem(SafeString(row("id")), SafeString(row("title"))))
            Next
            If cmbCategory.Items.Count > 0 Then cmbCategory.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadSpecList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, title FROM xipunum_erp_specs ORDER BY id")
            cmbSpec.Items.Clear()
            cmbSpec.Items.Add(New ComboBoxItem("", "请选择"))
            For Each row As DataRow In dt.Rows
                cmbSpec.Items.Add(New ComboBoxItem(SafeString(row("id")), SafeString(row("title"))))
            Next
            If cmbSpec.Items.Count > 0 Then cmbSpec.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If String.IsNullOrEmpty(txtName.Text.Trim()) Then
            ShowWarning("请输入名称！")
            Return
        End If

        Try
            Dim categoryId As String = ""
            If cmbCategory.SelectedIndex > 0 Then
                categoryId = DirectCast(cmbCategory.SelectedItem, ComboBoxItem).ID
            End If

            Dim specId As String = ""
            If cmbSpec.SelectedIndex > 0 Then
                specId = DirectCast(cmbSpec.SelectedItem, ComboBoxItem).ID
            End If

            ' 生成款号
            Dim kuanhaoCode As String = AuthShortCode & DateTime.Now.ToString("yyyyMMddHHmmss")

            Dim sql As String = $"INSERT INTO xipunum_erp_ksiamges (title, kuanhao, category_id, specification_id, caizhi, images, cjuser, creationtime) VALUES ('{SafeSQL(txtName.Text)}', '{kuanhaoCode}', '{categoryId}', '{specId}', '{SafeSQL(txtCaiZhi.Text)}', '{SafeSQL(txtImageUrl.Text)}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
            ExecuteCommand(sql)

            ' 更新正式款号
            sql = "SELECT LAST_INSERT_ID() AS id"
            Dim dt As DataTable = ExecuteQuery(sql)
            Dim newId As String = SafeString(dt.Rows(0)("id"))
            Dim formalKuanhao As String = AuthShortCode & newId.PadLeft(6, "0")
            sql = $"UPDATE xipunum_erp_ksiamges SET kuanhao='{formalKuanhao}' WHERE id='{newId}'"
            ExecuteCommand(sql)

            KuanHao = formalKuanhao
            ProductName = txtName.Text
            ImageUrl = txtImageUrl.Text

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

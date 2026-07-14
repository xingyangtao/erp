' ============================================================================
' 款号管理添加窗口
' 功能: 款号添加/修改
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class KuanHaoAddForm
    Inherits System.Windows.Forms.Form

    Private txtName As New TextBox()
    Private txtKuanHao As New TextBox()
    Private cmbCategory As New ComboBox()
    Private cmbSpec As New ComboBox()
    Private txtCaiZhi As New TextBox()
    Private cmbChengSe As New ComboBox()
    Private cmbLingXiao As New ComboBox()
    Private txtImageUrl As New TextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnCancel As New Button()
    Private editMode As Boolean = False
    Private editId As String = ""

    Public Sub New(Optional id As String = "")
        editId = id
        editMode = Not String.IsNullOrEmpty(id)
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = If(editMode, "编辑款号", "添加款号")
        Me.Size = New Drawing.Size(500, 450)
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

        Dim lblKuanHao As New Label()
        lblKuanHao.Text = "款号："
        lblKuanHao.Location = New Drawing.Point(30, 70)
        lblKuanHao.AutoSize = True
        Me.Controls.Add(lblKuanHao)

        txtKuanHao.Location = New Drawing.Point(100, 67)
        txtKuanHao.Size = New Drawing.Size(250, 25)
        Me.Controls.Add(txtKuanHao)

        Dim lblCategory As New Label()
        lblCategory.Text = "品类："
        lblCategory.Location = New Drawing.Point(30, 110)
        lblCategory.AutoSize = True
        Me.Controls.Add(lblCategory)

        cmbCategory.Location = New Drawing.Point(100, 107)
        cmbCategory.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbCategory)

        Dim lblSpec As New Label()
        lblSpec.Text = "规格："
        lblSpec.Location = New Drawing.Point(30, 150)
        lblSpec.AutoSize = True
        Me.Controls.Add(lblSpec)

        cmbSpec.Location = New Drawing.Point(100, 147)
        cmbSpec.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbSpec)

        Dim lblCaiZhi As New Label()
        lblCaiZhi.Text = "材质："
        lblCaiZhi.Location = New Drawing.Point(30, 190)
        lblCaiZhi.AutoSize = True
        Me.Controls.Add(lblCaiZhi)

        txtCaiZhi.Location = New Drawing.Point(100, 187)
        txtCaiZhi.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(txtCaiZhi)

        Dim lblChengSe As New Label()
        lblChengSe.Text = "成色："
        lblChengSe.Location = New Drawing.Point(30, 230)
        lblChengSe.AutoSize = True
        Me.Controls.Add(lblChengSe)

        cmbChengSe.Location = New Drawing.Point(100, 227)
        cmbChengSe.Size = New Drawing.Size(150, 25)
        cmbChengSe.Items.AddRange(New String() {"足金", "千足金", "万足金", "18K", "铂金", "银"})
        Me.Controls.Add(cmbChengSe)

        Dim lblLingXiao As New Label()
        lblLingXiao.Text = "零销售："
        lblLingXiao.Location = New Drawing.Point(30, 270)
        lblLingXiao.AutoSize = True
        Me.Controls.Add(lblLingXiao)

        cmbLingXiao.Location = New Drawing.Point(100, 267)
        cmbLingXiao.Size = New Drawing.Size(100, 25)
        cmbLingXiao.Items.AddRange(New String() {"否", "是"})
        cmbLingXiao.SelectedIndex = 0
        Me.Controls.Add(cmbLingXiao)

        Dim lblImageUrl As New Label()
        lblImageUrl.Text = "图片地址："
        lblImageUrl.Location = New Drawing.Point(30, 310)
        lblImageUrl.AutoSize = True
        Me.Controls.Add(lblImageUrl)

        txtImageUrl.Location = New Drawing.Point(100, 307)
        txtImageUrl.Size = New Drawing.Size(250, 25)
        Me.Controls.Add(txtImageUrl)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(100, 360)
        btnSave.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnSave)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(220, 360)
        btnCancel.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadCategoryList()
        LoadSpecList()
        If editMode Then
            LoadData()
        End If
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
            ShowError("加载品类列表失败：" & ex.Message)
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
            ShowError("加载规格列表失败：" & ex.Message)
        End Try
    End Sub

    Private Sub LoadData()
        Try
            Dim sql As String = $"SELECT * FROM xipunum_erp_ksiamges WHERE id='{editId}' LIMIT 1"
            Dim dt As DataTable = ExecuteQuery(sql)
            If dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                txtName.Text = SafeString(row("title"))
                txtKuanHao.Text = SafeString(row("kuanhao"))
                txtCaiZhi.Text = SafeString(row("caizhi"))
                txtImageUrl.Text = SafeString(row("images"))
                cmbChengSe.Text = SafeString(row("chengse"))
                cmbLingXiao.SelectedIndex = If(SafeString(row("lingxiao")) = "是", 1, 0)

                Dim categoryId As String = SafeString(row("category_id"))
                For i As Integer = 0 To cmbCategory.Items.Count - 1
                    If DirectCast(cmbCategory.Items(i), ComboBoxItem).ID = categoryId Then
                        cmbCategory.SelectedIndex = i
                        Exit For
                    End If
                Next

                Dim specId As String = SafeString(row("specification_id"))
                For i As Integer = 0 To cmbSpec.Items.Count - 1
                    If DirectCast(cmbSpec.Items(i), ComboBoxItem).ID = specId Then
                        cmbSpec.SelectedIndex = i
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
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

            Dim sql As String
            If editMode Then
                sql = $"UPDATE xipunum_erp_ksiamges SET title='{SafeSQL(txtName.Text)}', kuanhao='{SafeSQL(txtKuanHao.Text)}', category_id='{categoryId}', specification_id='{specId}', caizhi='{SafeSQL(txtCaiZhi.Text)}', chengse='{cmbChengSe.Text}', lingxiao='{cmbLingXiao.SelectedItem}', images='{SafeSQL(txtImageUrl.Text)}', updatetime='{GetOperationDate()}' WHERE id='{editId}'"
            Else
                sql = $"INSERT INTO xipunum_erp_ksiamges (title, kuanhao, category_id, specification_id, caizhi, chengse, lingxiao, images, cjuser, creationtime) VALUES ('{SafeSQL(txtName.Text)}', '{SafeSQL(txtKuanHao.Text)}', '{categoryId}', '{specId}', '{SafeSQL(txtCaiZhi.Text)}', '{cmbChengSe.Text}', '{cmbLingXiao.SelectedItem}', '{SafeSQL(txtImageUrl.Text)}', '{SafeSQL(UserAccount)}', '{GetOperationDate()}')"
            End If

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

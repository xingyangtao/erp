' ============================================================================
' 信息商品查询窗口
' 功能: 通用查询条件窗口
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class InfoProductQueryForm
    Inherits System.Windows.Forms.Form

    Private cmbShop As New ComboBox()
    Private cmbCategory As New ComboBox()
    Private cmbSpec As New ComboBox()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnCancel As New Button()

    Public Property QueryResult As DataTable = Nothing

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "信息商品查询"
        Me.Size = New Drawing.Size(450, 300)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblShop As New Label()
        lblShop.Text = "店铺："
        lblShop.Location = New Drawing.Point(30, 20)
        lblShop.AutoSize = True
        Me.Controls.Add(lblShop)

        cmbShop.Location = New Drawing.Point(100, 17)
        cmbShop.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbShop)

        Dim lblCategory As New Label()
        lblCategory.Text = "品类："
        lblCategory.Location = New Drawing.Point(30, 60)
        lblCategory.AutoSize = True
        Me.Controls.Add(lblCategory)

        cmbCategory.Location = New Drawing.Point(100, 57)
        cmbCategory.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbCategory)

        Dim lblSpec As New Label()
        lblSpec.Text = "规格："
        lblSpec.Location = New Drawing.Point(30, 100)
        lblSpec.AutoSize = True
        Me.Controls.Add(lblSpec)

        cmbSpec.Location = New Drawing.Point(100, 97)
        cmbSpec.Size = New Drawing.Size(150, 25)
        Me.Controls.Add(cmbSpec)

        Dim lblStart As New Label()
        lblStart.Text = "开始日期："
        lblStart.Location = New Drawing.Point(30, 140)
        lblStart.AutoSize = True
        Me.Controls.Add(lblStart)

        dtpStart.Location = New Drawing.Point(100, 137)
        dtpStart.Size = New Drawing.Size(150, 25)
        dtpStart.Format = DateTimePickerFormat.Short
        Me.Controls.Add(dtpStart)

        Dim lblEnd As New Label()
        lblEnd.Text = "结束日期："
        lblEnd.Location = New Drawing.Point(30, 180)
        lblEnd.AutoSize = True
        Me.Controls.Add(lblEnd)

        dtpEnd.Location = New Drawing.Point(100, 177)
        dtpEnd.Size = New Drawing.Size(150, 25)
        dtpEnd.Format = DateTimePickerFormat.Short
        Me.Controls.Add(dtpEnd)

        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(100, 220)
        btnQuery.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnQuery)

        btnCancel.Text = "取消"
        btnCancel.Location = New Drawing.Point(220, 220)
        btnCancel.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        LoadShopList()
        LoadCategoryList()
        LoadSpecList()
        dtpStart.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        dtpEnd.Value = DateTime.Now
    End Sub

    Private Sub LoadShopList()
        Try
            Dim sql As String = $"SELECT id, CASE WHEN id = '0' THEN '总库' ELSE title END AS title FROM xipunum_erp_type WHERE type = '商铺' AND superior = '0' AND id IN ({UserShopPermission}) ORDER BY id"
            Dim dt As DataTable = ExecuteQuery(sql)
            cmbShop.Items.Clear()
            cmbShop.Items.Add("全部")
            For Each row As DataRow In dt.Rows
                cmbShop.Items.Add(SafeString(row("title")))
            Next
            cmbShop.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadCategoryList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, title FROM xipunum_erp_category ORDER BY id")
            cmbCategory.Items.Clear()
            cmbCategory.Items.Add("全部品类")
            For Each row As DataRow In dt.Rows
                cmbCategory.Items.Add(SafeString(row("title")))
            Next
            cmbCategory.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadSpecList()
        Try
            Dim dt As DataTable = ExecuteQuery("SELECT id, title FROM xipunum_erp_specs ORDER BY id")
            cmbSpec.Items.Clear()
            cmbSpec.Items.Add("全部规格")
            For Each row As DataRow In dt.Rows
                cmbSpec.Items.Add(SafeString(row("title")))
            Next
            cmbSpec.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

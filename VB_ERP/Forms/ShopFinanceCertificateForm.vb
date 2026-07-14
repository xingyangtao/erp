' ============================================================================
' 报表店铺收支凭证窗口
' 功能: 收支凭证查看
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class ShopFinanceCertificateForm
    Inherits System.Windows.Forms.Form

    Private dgvList As New DataGridView()
    Private dtpStart As New DateTimePicker()
    Private dtpEnd As New DateTimePicker()
    Private WithEvents btnQuery As New Button()
    Private WithEvents btnView As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "报表店铺收支凭证"
        Me.Size = New Drawing.Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 60
        Me.Controls.Add(panelTop)

        Dim lblStart As New Label()
        lblStart.Text = "开始："
        lblStart.Location = New Drawing.Point(20, 18)
        lblStart.AutoSize = True
        panelTop.Controls.Add(lblStart)

        dtpStart.Location = New Drawing.Point(60, 15)
        dtpStart.Size = New Drawing.Size(130, 25)
        dtpStart.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpStart)

        Dim lblEnd As New Label()
        lblEnd.Text = "结束："
        lblEnd.Location = New Drawing.Point(200, 18)
        lblEnd.AutoSize = True
        panelTop.Controls.Add(lblEnd)

        dtpEnd.Location = New Drawing.Point(240, 15)
        dtpEnd.Size = New Drawing.Size(130, 25)
        dtpEnd.Format = DateTimePickerFormat.Short
        panelTop.Controls.Add(dtpEnd)

        btnQuery.Text = "查询"
        btnQuery.Location = New Drawing.Point(400, 13)
        btnQuery.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnQuery)

        btnView.Text = "查看凭证"
        btnView.Location = New Drawing.Point(490, 13)
        btnView.Size = New Drawing.Size(80, 30)
        panelTop.Controls.Add(btnView)

        dgvList.Dock = DockStyle.Fill
        dgvList.ReadOnly = True
        dgvList.AllowUserToAddRows = False
        dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.Controls.Add(dgvList)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        dtpStart.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        dtpEnd.Value = DateTime.Now
    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Try
            Dim startDate As String = dtpStart.Value.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim endDate As String = dtpEnd.Value.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00"

            Dim sql As String = $"SELECT a.id, a.type, a.title, a.amount, a.images, a.remarks, a.creationtime FROM xipunum_erp_finance AS a WHERE a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' AND a.kufang IN ({UserShopPermission}) AND a.images IS NOT NULL AND a.images != '' ORDER BY a.id DESC"
            Dim dt As DataTable = ExecuteQuery(sql)
            dgvList.DataSource = dt
            dgvList.AutoResizeColumns()
        Catch ex As Exception
            ShowError("查询失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnView_Click(sender As Object, e As EventArgs) Handles btnView.Click
        If dgvList.SelectedRows.Count = 0 Then
            ShowWarning("请先选择要查看的记录！")
            Return
        End If

        Dim images As String = SafeString(dgvList.SelectedRows(0).Cells("images").Value)
        If String.IsNullOrEmpty(images) Then
            ShowWarning("该记录没有凭证图片！")
            Return
        End If

        ' 打开图片预览
        Dim form As New KuanHaoImageForm(images)
        form.ShowDialog()
    End Sub
End Class

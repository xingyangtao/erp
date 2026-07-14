' ============================================================================
' 打印机设置窗口
' 功能: 打印机配置
' ============================================================================

Imports System.IO

Public Class PrinterSettingsForm
    Inherits System.Windows.Forms.Form

    Private cmbPrinters As New ComboBox()
    Private WithEvents btnSave As New Button()

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
    End Sub

    Private Sub InitializeUI()
        Me.Text = "打印机设置"
        Me.Size = New Drawing.Size(400, 200)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblPrinter As New Label()
        lblPrinter.Text = "选择打印机："
        lblPrinter.Location = New Drawing.Point(30, 30)
        lblPrinter.AutoSize = True
        Me.Controls.Add(lblPrinter)

        cmbPrinters.Location = New Drawing.Point(30, 60)
        cmbPrinters.Size = New Drawing.Size(330, 25)
        cmbPrinters.DropDownStyle = ComboBoxStyle.DropDownList
        Me.Controls.Add(cmbPrinters)

        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(30, 110)
        btnSave.Size = New Drawing.Size(100, 35)
        Me.Controls.Add(btnSave)
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 加载系统打印机列表
        For Each printer As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            cmbPrinters.Items.Add(printer)
        Next

        ' 加载保存的打印机
        Dim printPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "print.ini")
        If File.Exists(printPath) Then
            Dim savedPrinter As String = ReadConfigValue("print", "name", "")
            For i As Integer = 0 To cmbPrinters.Items.Count - 1
                If cmbPrinters.Items(i).ToString() = savedPrinter Then
                    cmbPrinters.SelectedIndex = i
                    Exit For
                End If
            Next
        End If

        If cmbPrinters.SelectedIndex < 0 AndAlso cmbPrinters.Items.Count > 0 Then
            cmbPrinters.SelectedIndex = 0
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If cmbPrinters.SelectedIndex < 0 Then
            ShowWarning("请选择打印机！")
            Return
        End If

        Try
            WriteConfigValue("print", "name", cmbPrinters.SelectedItem.ToString())
            ShowSuccess("保存成功！")
            Me.Close()
        Catch ex As Exception
            ShowError("保存失败：" & ex.Message)
        End Try
    End Sub
End Class

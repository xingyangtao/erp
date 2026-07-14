' ============================================================================
' 商品信息预售窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_商品信息预售.form.e.txt
' 包含所有8个程序集变量、17个子程序、所有SQL查询
' 两种模式：新建预售单（orderSelected=-1）、查看已有预售单
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.IO
Imports System.Net

Public Class PresaleForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（8个） ==========
    Private row As Integer = 0                           ' 集_行号
    Private col As Integer = 0                           ' 集_列号
    Private deleteBtn As New Button()                   ' 删除按钮（表格组件）
    Private uploadBtn As New Button()                   ' 图片上传按钮（表格组件）
    Private orderSelected As Integer = -1                ' 局部_订单是否选中
    Private localImagePath As String = ""                ' 局部_图片路径
    Private localImageResponse As String = ""             ' 局部_图片响应
    Private localImageName As String = ""                ' 局部_图片名称

    ' ========== 控件声明 ==========
    Private dgvProducts As New DataGridView()           ' 高级表格1
    Private txtOrderNumber As New TextBox()              ' 单据号_编辑框
    Private txtRemarks As New TextBox()                  ' 备注_编辑框
    Private txtMemberName As New TextBox()              ' 会员姓名_编辑框
    Private txtPhone As New TextBox()                   ' 联系电话_编辑框
    Private txtDeposit As New TextBox()                 ' 订金_编辑框
    Private txtSalesman As New TextBox()                ' 业务员_编辑框
    Private btnSave As New Button()                     ' 按钮_保存
    Private btnReset As New Button()                     ' 按钮_重置
    Private picImage As New PictureBox()                 ' 图片框1
    Private panelHeader As New Panel()                   ' 外形框_头部
    Private grpRemarks As New GroupBox()                 ' 分组框_备注
    Private lblTitle1 As New Label()                     ' 透明标签1
    Private lblTitle2 As New Label()                     ' 透明标签2
    Private lblMemberName As New Label()                 ' 透明标签10
    Private lblPhone As New Label()                      ' 透明标签11
    Private lblDeposit As New Label()                    ' 透明标签3
    Private lblSalesman As New Label()                   ' 透明标签12
    Private lblRemarksText As New Label()                ' 透明标签16
    Private ofdImage As New OpenFileDialog()             ' 通用对话框1

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler dgvProducts.SelectionChanged, AddressOf DgvProducts_SelectionChanged
        AddHandler dgvProducts.CellClick, AddressOf DgvProducts_CellClick
        AddHandler dgvProducts.CellEndEdit, AddressOf DgvProducts_CellEndEdit
        AddHandler txtPhone.TextChanged, AddressOf TxtPhone_TextChanged
        AddHandler txtPhone.KeyDown, AddressOf TxtPhone_KeyDown
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnReset.Click, AddressOf BtnReset_Click
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "商品信息预售"
        Me.Size = New Drawing.Size(1427, 800)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized

        ' 头部面板（外形框_头部）
        panelHeader.Dock = DockStyle.Top
        panelHeader.Height = 69
        panelHeader.BackColor = Drawing.Color.FromArgb(248, 248, 248)
        Me.Controls.Add(panelHeader)

        ' 透明标签1（标题）
        lblTitle1.Text = "商品信息预售"
        lblTitle1.Font = New Drawing.Font("微软雅黑", 14, Drawing.FontStyle.Bold)
        lblTitle1.AutoSize = True
        lblTitle1.Location = New Drawing.Point(5, 5)
        panelHeader.Controls.Add(lblTitle1)

        ' 透明标签2（单据号标签）
        lblTitle2.Text = "单据号："
        lblTitle2.AutoSize = True
        lblTitle2.Location = New Drawing.Point(0, 12)
        panelHeader.Controls.Add(lblTitle2)

        ' 单据号编辑框
        txtOrderNumber.Location = New Drawing.Point(60, 9)
        txtOrderNumber.Size = New Drawing.Size(200, 25)
        txtOrderNumber.ReadOnly = True
        panelHeader.Controls.Add(txtOrderNumber)

        ' 会员姓名标签
        lblMemberName.Text = "会员姓名："
        lblMemberName.AutoSize = True
        lblMemberName.Location = New Drawing.Point(7, 44)
        panelHeader.Controls.Add(lblMemberName)

        ' 会员姓名编辑框
        txtMemberName.Location = New Drawing.Point(67, 41)
        txtMemberName.Size = New Drawing.Size(100, 25)
        panelHeader.Controls.Add(txtMemberName)

        ' 联系电话标签
        lblPhone.Text = "联系电话："
        lblPhone.AutoSize = True
        lblPhone.Location = New Drawing.Point(185, 44)
        panelHeader.Controls.Add(lblPhone)

        ' 联系电话编辑框
        txtPhone.Location = New Drawing.Point(241, 41)
        txtPhone.Size = New Drawing.Size(130, 25)
        panelHeader.Controls.Add(txtPhone)

        ' 订金标签
        lblDeposit.Text = "订金："
        lblDeposit.AutoSize = True
        lblDeposit.Location = New Drawing.Point(383, 44)
        panelHeader.Controls.Add(lblDeposit)

        ' 订金编辑框
        txtDeposit.Location = New Drawing.Point(415, 41)
        txtDeposit.Size = New Drawing.Size(130, 25)
        panelHeader.Controls.Add(txtDeposit)

        ' 业务员标签
        lblSalesman.Text = "业务员："
        lblSalesman.AutoSize = True
        lblSalesman.Location = New Drawing.Point(563, 44)
        panelHeader.Controls.Add(lblSalesman)

        ' 业务员编辑框
        txtSalesman.Location = New Drawing.Point(623, 41)
        txtSalesman.Size = New Drawing.Size(100, 25)
        panelHeader.Controls.Add(txtSalesman)

        ' 高级表格1
        dgvProducts.Location = New Drawing.Point(5, 75)
        dgvProducts.Size = New Drawing.Size(975, 400)
        dgvProducts.AllowUserToAddRows = False
        dgvProducts.AllowUserToDeleteRows = False
        dgvProducts.RowHeadersVisible = False
        dgvProducts.BackgroundColor = Drawing.Color.White
        Me.Controls.Add(dgvProducts)

        ' 图片框1
        picImage.Location = New Drawing.Point(984, 75)
        picImage.Size = New Drawing.Size(400, 400)
        picImage.BorderStyle = BorderStyle.FixedSingle
        picImage.SizeMode = PictureBoxSizeMode.Zoom
        Me.Controls.Add(picImage)

        ' 按钮_重置
        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(1200, 120)
        btnReset.Size = New Drawing.Size(80, 32)
        btnReset.Visible = True
        Me.Controls.Add(btnReset)

        ' 按钮_保存
        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(1304, 120)
        btnSave.Size = New Drawing.Size(80, 32)
        btnSave.Visible = True
        Me.Controls.Add(btnSave)

        ' 分组框_备注
        grpRemarks.Text = ""
        grpRemarks.Location = New Drawing.Point(16, 544)
        grpRemarks.Size = New Drawing.Size(1288, 80)
        Me.Controls.Add(grpRemarks)

        ' 备注标签
        lblRemarksText.Text = "备注"
        lblRemarksText.AutoSize = True
        lblRemarksText.Location = New Drawing.Point(7, 27)
        grpRemarks.Controls.Add(lblRemarksText)

        ' 备注编辑框
        txtRemarks.Location = New Drawing.Point(35, 8)
        txtRemarks.Size = New Drawing.Size(1240, 67)
        txtRemarks.Multiline = True
        grpRemarks.Controls.Add(txtRemarks)

        ' 通用对话框
        ofdImage.Filter = "JPG文件 (*.jpg)|*.jpg|PNG文件 (*.png)|*.png|所有文件 (*.*)|*.*"

        ' 删除按钮配置
        deleteBtn.Text = "删除"
        deleteBtn.BackColor = Drawing.Color.LightCoral

        ' 图片上传按钮配置
        uploadBtn.Text = "图片上传"
    End Sub

    ' ========== 窗口加载（_窗口_商品信息预售_创建完毕） ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        ' 获取主窗口选中行
        orderSelected = MainForm.ReturnOrderButtonName

        ' 加载表头
        LoadTableHeader()

        ' 删除表格所有数据行
        ClearTableRows()

        ' 设置表格只读
        dgvProducts.ReadOnly = False

        If PresaleOrderNumber = "" Then
            ' 新建模式
            txtOrderNumber.Text = "YS" & DateTime.Now.ToString("yyyyMMdd") & "****"
            txtRemarks.Text = ""
            picImage.Image = Nothing
            txtMemberName.Text = ""
            txtMemberName.ReadOnly = False
            txtPhone.Text = ""
            txtPhone.ReadOnly = False
            txtDeposit.Text = ""
            txtDeposit.ReadOnly = False
            txtSalesman.Text = ""
            txtSalesman.ReadOnly = False
            btnReset.Visible = True
            btnSave.Visible = True
            AddPresaleDataRow()
            dgvProducts.SelectionMode = DataGridViewSelectionMode.CellSelect
        Else
            ' 查看模式
            txtOrderNumber.Text = PresaleOrderNumber
            txtRemarks.Text = SafeString(MainForm.dgvMain.Rows(orderSelected).Cells(8).Value)
            txtRemarks.ReadOnly = True
            txtMemberName.Text = SafeString(MainForm.dgvMain.Rows(orderSelected).Cells(3).Value)
            txtMemberName.ReadOnly = True
            txtPhone.Text = SafeString(MainForm.dgvMain.Rows(orderSelected).Cells(4).Value)
            txtPhone.ReadOnly = True
            txtDeposit.Text = SafeString(MainForm.dgvMain.Rows(orderSelected).Cells(5).Value)
            txtDeposit.ReadOnly = True
            txtSalesman.Text = SafeString(MainForm.dgvMain.Rows(orderSelected).Cells(9).Value)
            txtSalesman.ReadOnly = True
            btnReset.Visible = False
            btnSave.Visible = False
            LoadPresaleDetails()
            dgvProducts.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End If
    End Sub

    ' ========== 窗口尺寸改变（_窗口_商品信息预售_尺寸被改变） ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        If Me.WindowState = FormWindowState.Minimized Then Return

        Dim nWidth As Integer = Me.ClientSize.Width
        Dim nHeight As Integer = Me.ClientSize.Height

        panelHeader.Width = nWidth - 10
        panelHeader.Left = 5
        panelHeader.Top = 0

        dgvProducts.Width = nWidth - 415
        dgvProducts.Left = 5
        dgvProducts.Top = 75
        dgvProducts.Height = nHeight - 160

        picImage.Left = nWidth - 405
        picImage.Top = 75

        btnReset.Top = nHeight - 85
        btnReset.Left = nWidth - 170

        btnSave.Top = nHeight - 85
        btnSave.Left = nWidth - 85

        grpRemarks.Left = 5
        grpRemarks.Top = nHeight - 85
        grpRemarks.Width = nWidth - 10
        txtRemarks.Width = grpRemarks.Width - 40

        lblTitle1.Width = panelHeader.Width
        lblTitle1.Top = 5
        lblTitle2.Left = panelHeader.Width - 220
        lblTitle2.Top = 12
        txtOrderNumber.Left = panelHeader.Width - 130
        txtOrderNumber.Top = 12
    End Sub

    ' ========== 加载表头（_高级表格1_加载表头） ==========
    Private Sub LoadTableHeader()
        dgvProducts.Columns.Clear()

        Dim headers() As String
        Dim widths() As Integer

        If orderSelected = -1 Then
            ' 新建模式
            headers = {"序号", "商品名称", "数量", "备注", "图片地址", "图片", "操作"}
            widths = {50, 350, 100, 650, 0, 110, 150}
        Else
            ' 查看模式
            headers = {"序号", "商品名称", "数量", "备注", "图片地址", "图片", "预售时间"}
            widths = {50, 350, 100, 650, 0, 0, 150}
        End If

        For i As Integer = 0 To headers.Length - 1
            Dim dgvCol As New DataGridViewTextBoxColumn()
            dgvCol.HeaderText = headers(i)
            dgvCol.Name = "col" & i
            dgvCol.Width = widths(i)
            If i = 2 Then
                dgvCol.ValueType = GetType(Decimal)
            End If
            dgvProducts.Columns.Add(dgvCol)
        Next

        ' 新建模式时最后一列为操作列（按钮列）
        If orderSelected = -1 Then
            Dim btnCol As New DataGridViewButtonColumn()
            btnCol.HeaderText = "操作"
            btnCol.Name = "colAction"
            btnCol.Width = 150
            dgvProducts.Columns.RemoveAt(6)
            dgvProducts.Columns.Add(btnCol)
        End If
    End Sub

    ' ========== 删除表格所有数据行（子程序_删除表格） ==========
    Private Sub ClearTableRows()
        dgvProducts.Rows.Clear()
    End Sub

    ' ========== 光标位置改变（_高级表格1_光标位置改变） ==========
    Private Sub DgvProducts_SelectionChanged(sender As Object, e As EventArgs)
        If dgvProducts.CurrentRow Is Nothing Then Return
        If dgvProducts.CurrentRow.Index < 0 Then Return

        row = dgvProducts.CurrentRow.Index

        Dim imageFileName As String = SafeString(dgvProducts.Rows(row).Cells(4).Value)
        Dim imageId As String = SafeString(dgvProducts.Rows(row).Cells(5).Value)

        If imageFileName = "" OrElse imageId = "" OrElse imageId = "无" Then
            picImage.Image = Nothing
            Return
        End If

        Dim localPath As String = Application.StartupPath & "\images\presale\" & imageId & ".jpg"

        If Not File.Exists(localPath) Then
            Try
                If File.Exists(localPath) Then File.Delete(localPath)
                Dim imgBytes As Byte() = DownloadImage(imageFileName)
                If imgBytes IsNot Nothing AndAlso imgBytes.Length > 0 Then
                    Directory.CreateDirectory(Application.StartupPath & "\images\presale")
                    File.WriteAllBytes(localPath, imgBytes)
                End If
            Catch ex As Exception
            End Try
        End If

        Try
            If File.Exists(localPath) Then
                Using fs As New FileStream(localPath, FileMode.Open, FileAccess.Read)
                    picImage.Image = Drawing.Image.FromStream(fs)
                End Using
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' ========== 下载图片 ==========
    Private Function DownloadImage(url As String) As Byte()
        Try
            Using client As New WebClient()
                Return client.DownloadData(url)
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ' ========== 预售详情（_预售详情_被点击） ==========
    Private Sub LoadPresaleDetails()
        Dim sql As String = "SELECT a.id AS aid,a.product_name AS product_name,a.quantity AS auantity,a.remarks AS aremarks,a.images AS aimages,CASE WHEN a.images = '' THEN '无' ELSE '有' END AS tupian,a.creationtime AS acreationtime FROM xipunum_erp_presale AS a INNER JOIN xipunum_erp_presale_order AS b ON b.id = a.order_id AND b.presale_umber = '" & PresaleSQLSafe(txtOrderNumber.Text) & "' WHERE 1 = 1 ORDER BY a.id ASC"

        Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        Dim detailCount As Integer = dt.Rows.Count

        ' 显示进度条
        Dim progressForm As New ProgressBarForm()
        progressForm.LabelText = "数据正在加载中..."
        progressForm.MaxValue = detailCount
        progressForm.Value = 0
        progressForm.Show(Me)

        For i As Integer = 0 To detailCount - 1
            progressForm.LabelText = "正在获取所需数据...(" & CInt((i / detailCount) * 100) & "%)"

            Dim presaleId As String = SafeString(dt.Rows(i)("aid"))
            Dim productName As String = SafeString(dt.Rows(i)("product_name"))
            Dim quantity As String = SafeString(dt.Rows(i)("auantity"))
            Dim remarks As String = SafeString(dt.Rows(i)("aremarks"))
            Dim images As String = SafeString(dt.Rows(i)("aimages"))
            Dim creationtime As String = SafeString(dt.Rows(i)("acreationtime"))

            Dim rowIndex As Integer = dgvProducts.Rows.Add()
            dgvProducts.Rows(rowIndex).Cells(0).Value = (i + 1).ToString()
            dgvProducts.Rows(rowIndex).Cells(1).Value = productName
            dgvProducts.Rows(rowIndex).Cells(2).Value = quantity
            dgvProducts.Rows(rowIndex).Cells(3).Value = remarks
            dgvProducts.Rows(rowIndex).Cells(4).Value = images
            dgvProducts.Rows(rowIndex).Cells(5).Value = presaleId
            dgvProducts.Rows(rowIndex).Cells(6).Value = creationtime
            dgvProducts.Rows(rowIndex).ReadOnly = True

            ' 下载图片到本地
            Try
                Directory.CreateDirectory(Application.StartupPath & "\images")
                Directory.CreateDirectory(Application.StartupPath & "\images\presale")
                Dim localPath As String = Application.StartupPath & "\images\presale\" & presaleId & ".jpg"
                If File.Exists(localPath) Then File.Delete(localPath)
                If images <> "" Then
                    Dim imgBytes As Byte() = DownloadImage(images)
                    If imgBytes IsNot Nothing AndAlso imgBytes.Length > 0 Then
                        File.WriteAllBytes(localPath, imgBytes)
                    End If
                End If
            Catch ex As Exception
            End Try

            progressForm.Value = i + 1
            Application.DoEvents()
        Next

        progressForm.Close()

        ' 居中对齐
        For Each dgvRow As DataGridViewRow In dgvProducts.Rows
            dgvRow.Height = 28
        Next

        ' 显示第一行图片
        If detailCount > 0 Then
            Dim firstImageId As String = SafeString(dgvProducts.Rows(0).Cells(5).Value)
            Dim firstImagePath As String = Application.StartupPath & "\images\presale\" & firstImageId & ".jpg"
            If File.Exists(firstImagePath) Then
                Try
                    Using fs As New FileStream(firstImagePath, FileMode.Open, FileAccess.Read)
                        picImage.Image = Drawing.Image.FromStream(fs)
                    End Using
                Catch ex As Exception
                End Try
            End If
        End If
    End Sub

    ' ========== 联系电话内容改变（_联系电话_编辑框_内容被改变） ==========
    Private Sub TxtPhone_TextChanged(sender As Object, e As EventArgs)
        If txtPhone.Text.Length >= 11 Then
            Dim memberName As String = ""
            Dim sql As String = "SELECT name FROM xipunum_erp_member WHERE tel='" & PresaleSQLSafe(txtPhone.Text) & "' LIMIT 1"

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                memberName = SafeString(dt.Rows(0)("name"))
                txtMemberName.Text = memberName
            End If
        End If
    End Sub

    ' ========== 联系电话按键（_联系电话_编辑框_按下某键） ==========
    Private Sub TxtPhone_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If txtPhone.Text = "" Then
                ShowWarning("预售客户联系电话不能为空！")
                txtPhone.Text = ""
                txtPhone.Focus()
                Return
            End If
            TxtPhone_TextChanged(Nothing, Nothing)
            Return
        End If
    End Sub

    ' ========== 预售数据添加（_预售数据添加_被单击） ==========
    Private Sub AddPresaleDataRow()
        Dim rowCount As Integer = dgvProducts.Rows.Count

        If rowCount <= 1 Then
            ' 添加空行
            Dim rowIndex As Integer = dgvProducts.Rows.Add()
            dgvProducts.Rows(rowIndex).Cells(0).Value = (rowIndex + 1).ToString()
            dgvProducts.Rows(rowIndex).Cells(1).Value = ""
            dgvProducts.Rows(rowIndex).Cells(2).Value = 0
            dgvProducts.Rows(rowIndex).Cells(3).Value = ""
            dgvProducts.Rows(rowIndex).Cells(4).Value = ""
            dgvProducts.Rows(rowIndex).Cells(5).Value = "无"

            ' 添加合计行
            Dim totalRowIndex As Integer = dgvProducts.Rows.Add()
            dgvProducts.Rows(totalRowIndex).Cells(0).Value = ""
            dgvProducts.Rows(totalRowIndex).Cells(1).Value = "合计"
            dgvProducts.Rows(totalRowIndex).Cells(2).Value = "0"

            ' 设置合计行样式
            For i As Integer = 1 To 5
                dgvProducts.Rows(totalRowIndex).Cells(i).Style.BackColor = Drawing.Color.LightGray
                dgvProducts.Rows(totalRowIndex).Cells(i).ReadOnly = True
            Next
        Else
            ' 已有多行，在合计行前插入新行
            dgvProducts.Rows.RemoveAt(rowCount - 1)
            Dim newIndex As Integer = dgvProducts.Rows.Add()
            dgvProducts.Rows(newIndex).Cells(0).Value = (newIndex + 1).ToString()
            dgvProducts.Rows(newIndex).Cells(1).Value = ""
            dgvProducts.Rows(newIndex).Cells(2).Value = 0
            dgvProducts.Rows(newIndex).Cells(3).Value = ""
            dgvProducts.Rows(newIndex).Cells(4).Value = ""
            dgvProducts.Rows(newIndex).Cells(5).Value = "无"

            CalculateTotals()
        End If

        ' 重新编号
        For i As Integer = 0 To dgvProducts.Rows.Count - 2
            dgvProducts.Rows(i).Cells(0).Value = (i + 1).ToString()
        Next
    End Sub

    ' ========== 预售数据合计（_预售数据合计_被单击） ==========
    Private Sub CalculateTotals()
        Dim statRowCount As Integer = dgvProducts.Rows.Count - 2
        If statRowCount < 0 Then statRowCount = 0

        Dim totalQuantity As Decimal = 0

        ' 删除合计行
        If dgvProducts.Rows.Count > 0 Then
            dgvProducts.Rows.RemoveAt(dgvProducts.Rows.Count - 1)
        End If

        For i As Integer = 0 To statRowCount - 1
            totalQuantity += SafeDecimal(dgvProducts.Rows(i).Cells(2).Value)
        Next

        ' 添加合计行
        Dim totalRowIndex As Integer = dgvProducts.Rows.Add()
        dgvProducts.Rows(totalRowIndex).Cells(0).Value = ""
        dgvProducts.Rows(totalRowIndex).Cells(1).Value = "合计"
        dgvProducts.Rows(totalRowIndex).Cells(2).Value = totalQuantity.ToString()

        ' 设置合计行样式
        For i As Integer = 1 To 5
            dgvProducts.Rows(totalRowIndex).Cells(i).Style.BackColor = Drawing.Color.LightGray
            dgvProducts.Rows(totalRowIndex).Cells(i).ReadOnly = True
        Next
    End Sub

    ' ========== 表格按钮点击（_高级表格1_按钮被点击） ==========
    Private Sub DgvProducts_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return

        row = e.RowIndex
        col = e.ColumnIndex

        If col = 6 Then
            ' 操作列 - 图片上传
            UploadImageClicked()
        End If
    End Sub

    ' ========== 上传图片被点击（_上传图片_被点击） ==========
    Private Sub UploadImageClicked()
        Directory.CreateDirectory(Application.StartupPath & "\images")

        If ofdImage.ShowDialog() = DialogResult.OK Then
            localImageName = DateTime.Now.ToString("yyMMddHHmmss")

            ' 复制图片到本地
            Dim localPath As String = Application.StartupPath & "\images\" & localImageName & ".jpg"
            File.Copy(ofdImage.FileName, localPath, True)
            localImagePath = localPath

            ' 上传到服务器
            If Not UploadImage(localImagePath) OrElse localImageResponse = "" Then
                ShowWarning("图片上传失败！")
                picImage.Image = Nothing
                Return
            End If

            ShowSuccess("图片上传成功！")

            ' 显示图片
            Try
                Dim imgBytes As Byte() = DownloadImage(localImageResponse)
                If imgBytes IsNot Nothing AndAlso imgBytes.Length > 0 Then
                    Using ms As New MemoryStream(imgBytes)
                        picImage.Image = Drawing.Image.FromStream(ms)
                    End Using
                End If
            Catch ex As Exception
            End Try

            ' 更新表格数据
            dgvProducts.Rows(row).Cells(4).Value = localImageResponse
            dgvProducts.Rows(row).Cells(5).Value = localImageName

            ' 添加新行
            AddPresaleDataRow()
        End If
    End Sub

    ' ========== 上传图片函数（上传图片） ==========
    Private Function UploadImage(imagePath As String) As Boolean
        localImageResponse = ""

        Dim imgBytes As Byte() = File.ReadAllBytes(imagePath)
        If imgBytes.Length = 0 Then Return False

        Dim fileName As String = Path.GetFileName(imagePath)

        Try
            Dim uploadUrl As String = "http://erp.xipunum.com/uploads.php"
            Dim boundary As String = "----WebKitFormBoundaryvYHiHrIaFl4cMbTZ"

            Using client As New WebClient()
                client.Headers.Add("Content-Type", "multipart/form-data; boundary=" & boundary)

                ' 构建multipart/form-data请求体
                Dim sb As New StringBuilder()
                sb.AppendLine("--" & boundary)
                sb.AppendLine("Content-Disposition: form-data; name=""file""; filename=""" & fileName & """")
                sb.AppendLine("Content-Type: image/jpeg")
                sb.AppendLine()
                sb.AppendLine()

                Dim headerBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(sb.ToString())
                Dim footerBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(vbCrLf & "--" & boundary & "--" & vbCrLf)

                Dim dataBytes As New List(Of Byte())()
                dataBytes.Add(headerBytes)
                dataBytes.Add(imgBytes)
                dataBytes.Add(footerBytes)

                Dim totalLength As Integer = 0
                For Each b As Byte() In dataBytes
                    totalLength += b.Length
                Next

                Dim finalBytes(totalLength - 1) As Byte
                Dim offset As Integer = 0
                For Each b As Byte() In dataBytes
                    Array.Copy(b, 0, finalBytes, offset, b.Length)
                    offset += b.Length
                Next

                Dim responseBytes As Byte() = client.UploadData(uploadUrl, "POST", finalBytes)
                Dim response As String = System.Text.Encoding.UTF8.GetString(responseBytes)

                If Not response.Contains(".jpg") OrElse Not response.Contains("uploads/") Then
                    ShowWarning("款式图片上传失败！")
                    Return False
                End If

                localImageResponse = "http://erp.xipunum.com/" & response.Trim()
                Return True
            End Using
        Catch ex As Exception
            ShowWarning("款式图片上传失败！")
            Return False
        End Try
    End Function

    ' ========== 结束编辑（_高级表格1_结束编辑） ==========
    Private Sub DgvProducts_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.ColumnIndex = 2 Then
            CalculateTotals()
        End If
    End Sub

    ' ========== 重置按钮（_按钮_重置_被单击） ==========
    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        Form_Load(Nothing, Nothing)
    End Sub

    ' ========== 保存按钮（_按钮_保存_被单击） ==========
    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        InformationOperationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        InformationOperationAccount = UserAccount

        Dim productCount As Integer = dgvProducts.Rows.Count - 3
        If productCount <= 0 Then
            ShowWarning("预售商品不能为空！")
            Return
        End If

        If txtPhone.Text = "" Then
            ShowWarning("客户联系电话不能为空！")
            txtPhone.Text = ""
            txtPhone.Focus()
            Return
        End If

        If txtPhone.Text.Length <> 11 Then
            ShowWarning("请输入正确的客户手机号码！")
            txtPhone.Text = ""
            txtPhone.Focus()
            Return
        End If

        If txtDeposit.Text = "" Then
            ShowWarning("订金不能为空！")
            txtDeposit.Text = ""
            txtDeposit.Focus()
            Return
        End If

        If SafeDecimal(txtDeposit.Text) <= 0 Then
            ShowWarning("订金不能小于等于0元！")
            txtDeposit.Focus()
            Return
        End If

        ' 生成预售单号（附加时间戳和用户账户）
        Dim presaleOrderNo As String = ""
        presaleOrderNo = PresaleSQLSafe(txtOrderNumber.Text)
        presaleOrderNo = presaleOrderNo & CInt(DateTime.Now.Subtract(New DateTime(1970, 1, 1)).TotalSeconds).ToString() & UserAccount

        ' 检查单号是否存在
        Dim checkSql As String = "SELECT id FROM xipunum_erp_presale_order WHERE presale_umber='" & presaleOrderNo & "' LIMIT 1"
        Dim checkDt As DataTable = DatabaseModule.ExecuteQuery(checkSql, MySQL_Read)

        If checkDt IsNot Nothing AndAlso checkDt.Rows.Count > 0 Then
            ShowWarning("当前预售单号已存在，已重新生成单号，请再次点击保存！")
            txtOrderNumber.Text = "YS" & DateTime.Now.ToString("yyyyMMdd") & "****"
            Return
        End If

        ' 编码转换
        Dim presaleMemberName As String = txtMemberName.Text
        Dim presaleMemberPhone As String = txtPhone.Text
        Dim presaleMemberDeposit As String = txtDeposit.Text
        Dim presaleSalesman As String = txtSalesman.Text
        Dim presaleState As String = "待收货"
        Dim presaleRemarks As String = txtRemarks.Text

        ' 检查会员是否存在
        Dim memberCheckSql As String = "SELECT * FROM xipunum_erp_member where tel= '" & presaleMemberPhone & "'"
        Dim memberCheckDt As DataTable = DatabaseModule.ExecuteQuery(memberCheckSql, MySQL_Read)
        Dim memberExists As Integer = If(memberCheckDt IsNot Nothing, memberCheckDt.Rows.Count, 0)

        If memberExists = 0 Then
            ' 自动创建会员
            Dim memberCode As String = "HY" & DateTime.Now.ToString("yyyyMMddHHmmss")
            Dim insertMemberSql As String = "INSERT INTO xipunum_erp_member (customer_code,name,tel,cjuser,creationtime) VALUES ('" & SafeSQL(memberCode) & "','" & SafeSQL(presaleMemberName) & "','" & SafeSQL(presaleMemberPhone) & "','" & SafeSQL(InformationOperationAccount) & "','" & InformationOperationDate & "')"
            DatabaseModule.ExecuteCommand(insertMemberSql, MySQL_Write)
        End If

        ' 获取会员编码
        Dim customerCode As String = ""
        Dim memberSql As String = "SELECT customer_code FROM xipunum_erp_member where tel='" & presaleMemberPhone & "' order by id ASC LIMIT 1"
        Dim memberDt As DataTable = DatabaseModule.ExecuteQuery(memberSql, MySQL_Read)

        If memberDt IsNot Nothing AndAlso memberDt.Rows.Count > 0 Then
            customerCode = SafeString(memberDt.Rows(0)("customer_code"))
        End If

        If customerCode = "" Then
            ShowWarning("会员信息读取失败，请检查联系电话！")
            Return
        End If

        ' INSERT 预售主单
        Dim insertOrderSql As String = "INSERT INTO xipunum_erp_presale_order (presale_umber,customer_code,deposit,remarks,state,yewu,cjuser,creationtime) VALUES ('" & SafeSQL(presaleOrderNo) & "','" & SafeSQL(customerCode) & "','" & SafeSQL(presaleMemberDeposit) & "','" & SafeSQL(presaleRemarks) & "','" & SafeSQL(presaleState) & "','" & SafeSQL(presaleSalesman) & "','" & SafeSQL(InformationOperationAccount) & "','" & InformationOperationDate & "')"
        DatabaseModule.ExecuteCommand(insertOrderSql, MySQL_Write)

        ' 获取预售订单ID
        Dim orderId As String = ""
        Dim orderIdSql As String = "SELECT id FROM xipunum_erp_presale_order WHERE presale_umber='" & presaleOrderNo & "' ORDER BY id ASC LIMIT 1"
        Dim orderIdDt As DataTable = DatabaseModule.ExecuteQuery(orderIdSql, MySQL_Read)

        If orderIdDt IsNot Nothing AndAlso orderIdDt.Rows.Count > 0 Then
            orderId = SafeString(orderIdDt.Rows(0)("id"))
        End If

        If orderId = "" Then
            ShowWarning("预售主单保存失败，请重试！")
            Return
        End If

        ' INSERT 收款记录（类别='3'预售）
        Dim insertPaymentSql As String = "INSERT INTO xipunum_erp_shoukuan (leibie,settlement_number,xianjin,type,kufang,cjuser,creationtime) VALUES ('3','" & SafeSQL(presaleOrderNo) & "','" & SafeSQL(presaleMemberDeposit) & "','1','" & SafeSQL(UserDepartment) & "','" & SafeSQL(InformationOperationAccount) & "','" & InformationOperationDate & "')"
        DatabaseModule.ExecuteCommand(insertPaymentSql, MySQL_Write)

        ' 更新单据号（追加订单ID后4位）
        txtOrderNumber.Text = txtOrderNumber.Text.Substring(0, 10) & ("0000" & orderId).Substring(Math.Max(0, ("0000" & orderId).Length - 4))

        Dim updatedOrderNo As String = PresaleSQLSafe(txtOrderNumber.Text)
        Dim updateOrderSql As String = "UPDATE xipunum_erp_presale_order SET presale_umber='" & updatedOrderNo & "' WHERE id='" & orderId & "' LIMIT 1"
        DatabaseModule.ExecuteCommand(updateOrderSql, MySQL_Write)

        Dim updatePaymentSql As String = "UPDATE xipunum_erp_shoukuan SET settlement_number='" & updatedOrderNo & "' WHERE settlement_number='" & SafeSQL(presaleOrderNo) & "' AND leibie='3' LIMIT 1"
        DatabaseModule.ExecuteCommand(updatePaymentSql, MySQL_Write)

        ' 插入系统日志
        LogSaveContent = ""
        LogSaveContent = "账户:" & UserAccount & " 商品预售，预售单号:" & txtOrderNumber.Text
        Dim insertLogSql As String = "INSERT INTO xipunum_erp_xitong_log (type,title,conter,user,creationtime) VALUES ('" & SafeSQL("添加") & "','" & SafeSQL("商品预售") & "','" & SafeSQL(LogSaveContent) & "','" & SafeSQL(InformationOperationAccount) & "','" & InformationOperationDate & "')"
        DatabaseModule.ExecuteCommand(insertLogSql, MySQL_Write)

        ' 保存明细（显示进度条）
        Dim progressForm As New ProgressBarForm()
        progressForm.LabelText = "数据正在保存中..."
        progressForm.MaxValue = productCount
        progressForm.Value = 0
        progressForm.Show(Me)

        For i As Integer = 0 To productCount - 1
            progressForm.LabelText = "正在保存数据...(" & CInt((i / productCount) * 100) & "%)"

            Dim productName As String = SafeString(dgvProducts.Rows(i).Cells(1).Value)
            Dim quantity As String = SafeString(dgvProducts.Rows(i).Cells(2).Value)
            Dim remarks As String = SafeString(dgvProducts.Rows(i).Cells(3).Value)
            Dim images As String = SafeString(dgvProducts.Rows(i).Cells(4).Value)

            Dim insertDetailSql As String = "INSERT INTO xipunum_erp_presale (order_id,product_name,images,quantity,remarks,cjuser,creationtime) VALUES ('" & SafeSQL(orderId) & "','" & SafeSQL(productName) & "','" & SafeSQL(images) & "','" & SafeSQL(quantity) & "','" & SafeSQL(remarks) & "','" & SafeSQL(InformationOperationAccount) & "','" & InformationOperationDate & "')"
            DatabaseModule.ExecuteCommand(insertDetailSql, MySQL_Write)

            ' 插入明细日志
            LogSaveContent = ""
            LogSaveContent = "账户:" & UserAccount & " 商品预售，名称：" & productName
            Dim insertDetailLogSql As String = "INSERT INTO xipunum_erp_xitong_log (type,title,conter,user,creationtime) VALUES ('" & SafeSQL("添加") & "','" & SafeSQL("商品预售") & "','" & SafeSQL(LogSaveContent) & "','" & SafeSQL(InformationOperationAccount) & "','" & InformationOperationDate & "')"
            DatabaseModule.ExecuteCommand(insertDetailLogSql, MySQL_Write)

            progressForm.Value = i + 1
            Application.DoEvents()
            System.Threading.Thread.Sleep(50)
        Next

        progressForm.Close()

        ShowSuccess("预售商品操作完成！")
        Me.Close()

        ' 刷新主窗口
        If HomePageQueryText = "商品预售" Then
            CType(Me.Owner, MainForm).RefreshSubFolderTable()
        End If
    End Sub

    ' ========== SQL文本安全（_预售_SQL文本安全） ==========
    Private Function PresaleSQLSafe(text As String) As String
        Return text.Replace("'", "''")
    End Function

End Class

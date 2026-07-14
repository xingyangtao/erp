' ============================================================================
' 入库库房修改窗口
' 完全按照易语言源文件实现：窗口程序集_窗口_入库库房修改.form.e.txt
' 包含所有3个程序集变量、5个子程序、所有SQL查询
' 功能：修改入库订单的库房，同时更新商品、历史、临存表和日志
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Public Class InboundWarehouseEditForm
    Inherits System.Windows.Forms.Form

    ' ========== 程序集变量（3个） ==========
    Private mainRow As Integer = -1                        ' 局部_主窗口行号
    Private orderId As String = ""                        ' 局部_订单id
    Private isSaving As Boolean = False                   ' 局部_正在保存

    ' ========== 控件声明 ==========
    Private txtOrderCode As New TextBox()                 ' 编辑框_订单编码
    Private txtOriginalWarehouse As New TextBox()         ' 编辑框_原入库库房
    Private cmbNewWarehouse As New ComboBox()            ' 组合框新入库库房
    Private btnSave As New Button()                       ' 按钮EX1（保存）
    Private btnReset As New Button()                      ' 按钮EX2（重置）
    Private btnClose As New Button()                      ' 图片框EX4（关闭）
    Private grpEdit As New GroupBox()                     ' 添加修改_分组框

    ' 库房列表数据（ID + 名称）
    Private warehouseList As New List(Of (Id As String, Name As String))()

    ' ========== 构造函数 ==========
    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnReset.Click, AddressOf BtnReset_Click
        AddHandler btnClose.Click, AddressOf BtnClose_Click
    End Sub

    ' ========== 初始化界面 ==========
    Private Sub InitializeUI()
        Me.Text = "入库库房修改"
        Me.Size = New Drawing.Size(450, 280)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        ' 分组框
        grpEdit.Text = "添加修改"
        grpEdit.Size = New Drawing.Size(380, 200)
        grpEdit.Location = New Drawing.Point(25, 15)
        Me.Controls.Add(grpEdit)

        ' 订单编码标签
        Dim lblOrderCode As New Label()
        lblOrderCode.Text = "订单编码："
        lblOrderCode.AutoSize = True
        lblOrderCode.Location = New Drawing.Point(20, 30)
        grpEdit.Controls.Add(lblOrderCode)

        ' 订单编码编辑框（只读）
        txtOrderCode.Location = New Drawing.Point(100, 27)
        txtOrderCode.Size = New Drawing.Size(250, 25)
        txtOrderCode.ReadOnly = True
        grpEdit.Controls.Add(txtOrderCode)

        ' 原入库库房标签
        Dim lblOriginal As New Label()
        lblOriginal.Text = "原入库库房："
        lblOriginal.AutoSize = True
        lblOriginal.Location = New Drawing.Point(20, 65)
        grpEdit.Controls.Add(lblOriginal)

        ' 原入库库房编辑框（只读）
        txtOriginalWarehouse.Location = New Drawing.Point(100, 62)
        txtOriginalWarehouse.Size = New Drawing.Size(250, 25)
        txtOriginalWarehouse.ReadOnly = True
        grpEdit.Controls.Add(txtOriginalWarehouse)

        ' 新入库库房标签
        Dim lblNew As New Label()
        lblNew.Text = "新入库库房："
        lblNew.AutoSize = True
        lblNew.Location = New Drawing.Point(20, 100)
        grpEdit.Controls.Add(lblNew)

        ' 新入库库房组合框
        cmbNewWarehouse.Location = New Drawing.Point(100, 97)
        cmbNewWarehouse.Size = New Drawing.Size(250, 25)
        cmbNewWarehouse.DropDownStyle = ComboBoxStyle.DropDownList
        grpEdit.Controls.Add(cmbNewWarehouse)

        ' 保存按钮
        btnSave.Text = "保存"
        btnSave.Location = New Drawing.Point(50, 145)
        btnSave.Size = New Drawing.Size(90, 35)
        grpEdit.Controls.Add(btnSave)

        ' 重置按钮
        btnReset.Text = "重置"
        btnReset.Location = New Drawing.Point(155, 145)
        btnReset.Size = New Drawing.Size(90, 35)
        grpEdit.Controls.Add(btnReset)

        ' 关闭按钮
        btnClose.Text = "关闭"
        btnClose.Location = New Drawing.Point(260, 145)
        btnClose.Size = New Drawing.Size(90, 35)
        grpEdit.Controls.Add(btnClose)
    End Sub

    ' ========== 窗口创建完毕 ==========
    Private Sub Form_Load(sender As Object, e As EventArgs)
        isSaving = False

        ' 获取主窗口选中的行号
        Try
            Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
            If mainForm IsNot Nothing AndAlso mainForm.dgvMain IsNot Nothing AndAlso mainForm.dgvMain.CurrentCell IsNot Nothing Then
                mainRow = mainForm.dgvMain.CurrentCell.RowIndex
            End If
        Catch
        End Try

        If mainRow < 0 Then
            ShowWarning("请先在主窗口选择一条入库订单！")
            Me.Close()
            Return
        End If

        ' 从主窗口获取数据
        Try
            Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
            If mainForm IsNot Nothing Then
                ' 获取订单ID（列16）
                orderId = SafeString(mainForm.dgvMain.Rows(mainRow).Cells(16).Value)
                ' 获取订单编码（列2）
                txtOrderCode.Text = SafeString(mainForm.dgvMain.Rows(mainRow).Cells(2).Value)
                ' 获取原入库库房（列7）
                txtOriginalWarehouse.Text = SafeString(mainForm.dgvMain.Rows(mainRow).Cells(7).Value)
            End If
        Catch
        End Try

        If String.IsNullOrEmpty(txtOrderCode.Text) Then
            ShowWarning("入库单号为空，无法修改库房！")
            Me.Close()
            Return
        End If

        ' 加载库房列表
        LoadWarehouseList()
    End Sub

    ' ========== 加载库房列表 ==========
    Private Sub LoadWarehouseList()
        cmbNewWarehouse.Items.Clear()
        warehouseList.Clear()

        Try
            Dim sql As String = ""
            Dim shopPermission As String = UserShopPermission

            If String.IsNullOrEmpty(shopPermission) Then
                sql = "SELECT '0' AS akufang, '总库' AS btitle"
            Else
                sql = "SELECT id AS akufang, CASE WHEN id='0' THEN '总库' ELSE title END AS btitle FROM xipunum_erp_type WHERE type='商铺' AND superior='0' AND id IN (" & shopPermission & ") UNION ALL SELECT '0' AS akufang, '总库' AS btitle FROM dual WHERE '0' IN (" & shopPermission & ") ORDER BY akufang='0' DESC, akufang"
            End If

            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)

            For Each r As DataRow In dt.Rows
                Dim whId As String = SafeString(r("akufang"))
                Dim whName As String = SafeString(r("btitle"))
                warehouseList.Add((whId, whName))
                cmbNewWarehouse.Items.Add(whName)
            Next
        Catch
        End Try

        If cmbNewWarehouse.Items.Count > 0 Then
            cmbNewWarehouse.SelectedIndex = -1
        End If
    End Sub

    ' ========== 窗口尺寸改变 ==========
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        grpEdit.Left = CInt((Me.ClientSize.Width - grpEdit.Width) / 2)
        grpEdit.Top = CInt((Me.ClientSize.Height - grpEdit.Height) / 2)
    End Sub

    ' ========== 关闭按钮 ==========
    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        If isSaving Then Return
        Me.Close()
    End Sub

    ' ========== 保存按钮 ==========
    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        If isSaving Then Return

        If mainRow < 0 Then
            ShowWarning("请先选择入库订单！")
            Return
        End If

        If cmbNewWarehouse.SelectedIndex = -1 Then
            ShowWarning("请选择需要修改的入库库房！")
            Return
        End If

        Dim newWarehouseName As String = cmbNewWarehouse.SelectedItem.ToString()
        If newWarehouseName = txtOriginalWarehouse.Text Then
            ShowWarning("库房无任何修改！")
            Return
        End If

        Dim orderNumber As String = txtOrderCode.Text
        If String.IsNullOrEmpty(orderNumber) Then
            ShowWarning("入库单号不能为空！")
            Return
        End If

        isSaving = True
        InformationOperationDate = Now.ToString("yyyy-MM-dd HH:mm:ss")
        InformationOperationAccount = UserAccount

        ' 获取新库房ID
        Dim newWarehouseId As String = warehouseList(cmbNewWarehouse.SelectedIndex).Id

        ' 查询入库订单信息
        Dim creatorAccount As String = ""
        Dim orderIdVal As String = ""
        Dim creationTime As String = ""

        Try
            Dim sql As String = "SELECT cjuser,id,creationtime FROM xipunum_erp_store_order WHERE odd_numbers='" & SafeSQL(orderNumber) & "' LIMIT 1"
            Dim dt As DataTable = DatabaseModule.ExecuteQuery(sql, MySQL_Read)
            If dt.Rows.Count = 0 Then
                ShowWarning("未找到该入库订单，无法修改库房！")
                isSaving = False
                Return
            End If
            creatorAccount = SafeString(dt.Rows(0)("cjuser"))
            orderIdVal = SafeString(dt.Rows(0)("id"))
            creationTime = SafeString(dt.Rows(0)("creationtime"))
        Catch ex As Exception
            ShowError("查询入库订单失败：" & ex.Message)
            isSaving = False
            Return
        End Try

        If String.IsNullOrEmpty(orderIdVal) Then
            ShowWarning("入库订单ID为空，无法修改库房！")
            isSaving = False
            Return
        End If

        ' 查询入库明细
        Dim detailSql As String = "SELECT a.poduct_code AS bianam,a.creationtime AS creationtime FROM xipunum_erp_store AS a INNER JOIN xipunum_erp_store_order AS b ON b.id=a.order_id WHERE b.odd_numbers='" & SafeSQL(orderNumber) & "' ORDER BY a.id ASC"
        Dim detailDt As DataTable

        Try
            detailDt = DatabaseModule.ExecuteQuery(detailSql, MySQL_Read)
            If detailDt.Rows.Count = 0 Then
                ShowWarning("该入库单没有明细，无法修改库房！")
                isSaving = False
                Return
            End If
        Catch ex As Exception
            ShowError("查询入库明细失败：" & ex.Message)
            isSaving = False
            Return
        End Try

        ' 循环更新每个商品的库房
        For i As Integer = 0 To detailDt.Rows.Count - 1
            Dim productCode As String = SafeString(detailDt.Rows(i)("bianam"))
            Dim productCreateTime As String = SafeString(detailDt.Rows(i)("creationtime"))

            Try
                Dim updateSql As String = "UPDATE xipunum_erp_shop SET kufang='" & SafeSQL(newWarehouseId) & "' WHERE poduct_code='" & SafeSQL(productCode) & "' AND creationtime='" & SafeSQL(productCreateTime) & "' AND cjuser='" & SafeSQL(creatorAccount) & "' LIMIT 1"
                DatabaseModule.ExecuteCommand(updateSql, MySQL_Write)
            Catch
            End Try

            ' 插入日志
            LogSaveContent = "账户:" & UserAccount & " 修改商品编码：" & productCode & " 原库房：" & txtOriginalWarehouse.Text & "->" & newWarehouseName
            Try
                Dim logSql As String = "INSERT INTO xipunum_erp_xitong_log (type,title,conter,user,creationtime) VALUES ('修改','修改商品参数','" & SafeSQL(LogSaveContent) & "','" & SafeSQL(InformationOperationAccount) & "','" & SafeSQL(InformationOperationDate) & "')"
                DatabaseModule.ExecuteCommand(logSql, MySQL_Write)
            Catch
            End Try
        Next

        ' 更新历史表
        Try
            Dim updateHistorySql As String = "UPDATE xipunum_erp_history SET conter=REPLACE(conter,'" & SafeSQL(txtOriginalWarehouse.Text) & "','" & SafeSQL(newWarehouseName) & "') WHERE number='" & SafeSQL(orderNumber) & "'"
            DatabaseModule.ExecuteCommand(updateHistorySql, MySQL_Write)
        Catch
        End Try

        ' 更新临存表
        Try
            Dim updateLincunSql As String = "UPDATE xipunum_erp_shop_lincun SET kufang='" & SafeSQL(newWarehouseId) & "' WHERE rukuid='" & SafeSQL(orderIdVal) & "' AND state='0' AND cjuser='" & SafeSQL(creatorAccount) & "' AND creationtime='" & SafeSQL(creationTime) & "'"
            DatabaseModule.ExecuteCommand(updateLincunSql, MySQL_Write)
        Catch
        End Try

        ' 插入入库日志
        Try
            Dim storeLogSql As String = "INSERT INTO xipunum_erp_store_log (order_id,user,conter,creationtime) VALUES ('" & SafeSQL(orderId) & "','" & SafeSQL(InformationOperationAccount) & "','修改库房','" & SafeSQL(InformationOperationDate) & "')"
            DatabaseModule.ExecuteCommand(storeLogSql, MySQL_Write)
        Catch
        End Try

        ' 插入系统日志
        LogSaveContent = "账户:" & UserAccount & " 修改入库订单：" & orderNumber & " 库房由：" & txtOriginalWarehouse.Text & "->" & newWarehouseName
        Try
            Dim logSql As String = "INSERT INTO xipunum_erp_xitong_log (type,title,conter,user,creationtime) VALUES ('修改','修改入库订单','" & SafeSQL(LogSaveContent) & "','" & SafeSQL(InformationOperationAccount) & "','" & SafeSQL(InformationOperationDate) & "')"
            DatabaseModule.ExecuteCommand(logSql, MySQL_Write)
        Catch
        End Try

        ' 更新主窗口表格
        Try
            Dim mainForm As MainForm = TryCast(Application.OpenForms("MainForm"), MainForm)
            If mainForm IsNot Nothing AndAlso mainRow < mainForm.dgvMain.Rows.Count Then
                mainForm.dgvMain.Rows(mainRow).Cells(7).Value = newWarehouseName
            End If
        Catch
        End Try

        ShowSuccess("入库订单信息修改完成！")
        Me.Close()
    End Sub

    ' ========== 重置按钮 ==========
    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        Form_Load(Nothing, Nothing)
    End Sub

End Class

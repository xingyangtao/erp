$content = [System.IO.File]::ReadAllText('E:\ERPV4\VB_ERP\Forms\MainForm.vb')
$startMarker = "    ' ========== 页面路由 =========="
$endMarker = "    ' ========== 工具栏按钮点击 =========="
$start = $content.IndexOf($startMarker)
$end = $content.IndexOf($endMarker)

$newCode = @"
    ' ========== 页面路由 ==========
    Private Sub RouteToPage(pageName As String, pageId As String)
        Select Case pageName
            Case "首页"
                ShowHomePage()
            Case "系统设置", "商铺信息"
                ShowSystemSettings()
            Case "账户列表"
                ShowAccountList()
            Case "岗位分组"
                ShowRoleGroupPage()
            Case "商品列表"
                ShowProductList()
            Case "商品入库"
                ShowInboundList()
            Case "商品销售"
                ShowSalesList()
            Case "商品调拨"
                ShowTransferList()
            Case "商品退库"
                ShowReturnList()
            Case "商品回收"
                ShowRecoveryList()
            Case "商品预售"
                ShowPresaleList()
            Case "商品退货"
                ShowRefundList()
            Case "会员列表"
                ShowMemberList()
            Case "实时库存"
                ShowInventoryList()
            Case "日志记录"
                ShowLogPage()
            Case Else
                dgvMain.DataSource = Nothing
                dgvMain.Rows.Clear()
                dgvMain.Columns.Clear()
                dgvMain.Columns.Add("info", "提示")
                dgvMain.Rows.Add("页面 [" & pageName & "] (ID:" & pageId & ") 暂未实现")
        End Select
    End Sub

"@

$before = $content.Substring(0, $start)
$after = $content.Substring($end)
$result = $before + $newCode + $after
[System.IO.File]::WriteAllText('E:\ERPV4\VB_ERP\Forms\MainForm.vb', $result)
Write-Host "RouteToPage replaced successfully"

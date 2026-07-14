# -*- coding: utf-8 -*-
import sys

with open(r'E:\ERPV4\VB_ERP\Forms\MainForm.vb', 'r', encoding='utf-8') as f:
    content = f.read()

start_marker = "    ' ========== 页面路由 =========="
end_marker = "    ' ========== 工具栏按钮点击 =========="

start = content.index(start_marker)
end = content.index(end_marker)

new_code = """    ' ========== 页面路由 ==========
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

"""

with open(r'E:\ERPV4\VB_ERP\Forms\MainForm.vb', 'w', encoding='utf-8') as f:
    f.write(content[:start] + new_code + content[end:])

print('RouteToPage replaced successfully')

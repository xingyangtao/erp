' ============================================================================
' 商品入库共享状态模块
' 用于 ProductAddForm 和 ProductStockInForm 之间的数据共享
' ============================================================================

Module ProductStockInShared

    ' ========== 入库商品共享变量 ==========
    Public 局部_入库商品数量 As Integer = 0       ' 已入库商品数量
    Public 局部_商品品类id As String = ""          ' 当前商品品类ID
    Public 局部_商品是否镶嵌 As String = ""        ' 商品是否镶嵌（"镶嵌" 或 ""）
    Public 入库品类名称 As String = ""             ' 入库时选中的品类名称
    Public 入库克价 As String = ""                 ' 入库时选中的入库克价

End Module

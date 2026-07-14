' ============================================================================
' 商品报表查询模块
' 包含所有20个报表窗口的完整SQL查询逻辑
' ============================================================================

Module ReportQueryModule

    ' ========== 1. 商品销售报表 SQL ==========
    Public Function GetSalesReportSQL(startDate As String, endDate As String, shopPermission As String, mode As String) As String
        Select Case mode
            Case "订单"
                Return $"SELECT f.settlement_number, f.creationtime, f.customer_code, k.name AS customer_name, k.tel, f.presale_number, " &
                    "CAST(COALESCE(SUM(b.quantity),0) AS DECIMAL(30,2)) AS shuliang, " &
                    "CAST(COALESCE(SUM(b.net_weight),0) AS DECIMAL(30,3)) AS jinzhong, " &
                    "CAST(COALESCE(SUM(b.xiao_amount),0) AS DECIMAL(30,2)) AS xiaoshou, " &
                    "CAST(COALESCE(SUM(b.settlement),0) AS DECIMAL(30,2)) AS shishou " &
                    "FROM xipunum_erp_outbound AS b " &
                    "INNER JOIN xipunum_erp_outbound_order AS f ON f.id = b.order_id " &
                    "LEFT JOIN xipunum_erp_member AS k ON k.customer_code = f.customer_code " &
                    $"WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}' {shopPermission} " &
                    "GROUP BY f.id ORDER BY f.creationtime DESC"
            Case "明细"
                Return $"SELECT b.poduct_code, a.product_name, b.creationtime, f.settlement_number, " &
                    "h.title AS kufang, g.title AS gongchang, a.item_number AS kuanhao, " &
                    "d.title AS pinlei, e.title AS guige, a.caizhi, a.single AS danzhong, " &
                    "b.quantity AS shuliang, b.net_weight AS jinzhong, b.weight AS zhongliang, " &
                    "b.gold_price AS kejia, b.xiao_amount AS xiaoshoujia, b.settlement AS yingshou, " &
                    "c.basic_cost AS chengbengf, c.company_surcharge AS chengbenfj, c.cost_price AS chengben, " &
                    "b.sales_cost AS xiaoshougf, b.sales_surcharge AS xiaoshoufj, b.zhekou, " &
                    "m.name AS daogou, b.pling " &
                    "FROM xipunum_erp_outbound AS b " &
                    "INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code " &
                    "INNER JOIN xipunum_erp_outbound_order AS f ON f.id = b.order_id " &
                    "LEFT JOIN xipunum_erp_store AS c ON c.poduct_code = b.poduct_code " &
                    "LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id " &
                    "LEFT JOIN xipunum_erp_specs AS e ON e.id = a.specification_id " &
                    "LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang " &
                    "LEFT JOIN xipunum_erp_type AS g ON g.id = c.factory " &
                    "LEFT JOIN xipunum_erp_user AS m ON m.user = b.shopping_guide " &
                    $"WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}' {shopPermission} " &
                    "ORDER BY b.creationtime DESC"
            Case Else
                Return ""
        End Select
    End Function

    ' ========== 2. 商品入库报表 SQL ==========
    Public Function GetInboundReportSQL(startDate As String, endDate As String, shopPermission As String, mode As String) As String
        Select Case mode
            Case "订单"
                Return $"SELECT b.number, b.creationtime, g.title AS gongchang, " &
                    "CAST(COALESCE(SUM(a.quantity),0) AS DECIMAL(30,2)) AS shuliang, " &
                    "CAST(COALESCE(SUM(a.net_weight),0) AS DECIMAL(30,3)) AS jinzhong, " &
                    "CAST(COALESCE(SUM(a.weight),0) AS DECIMAL(30,3)) AS zhongliang, " &
                    "CAST(COALESCE(SUM(a.cost_price),0) AS DECIMAL(30,2)) AS chengben " &
                    "FROM xipunum_erp_store AS a " &
                    "INNER JOIN xipunum_erp_store_order AS b ON b.id = a.order_id " &
                    "LEFT JOIN xipunum_erp_type AS g ON g.id = b.factory " &
                    $"WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}' {shopPermission} " &
                    "GROUP BY b.id ORDER BY b.creationtime DESC"
            Case "明细"
                Return $"SELECT a.poduct_code, a.product_name, b.creationtime, b.number, " &
                    "g.title AS gongchang, h.title AS kufang, d.title AS pinlei, e.title AS guige, " &
                    "a.caizhi, a.single AS danzhong, a.quantity AS shuliang, " &
                    "a.net_weight AS jinzhong, a.weight AS zhongliang, a.cost_price AS chengben " &
                    "FROM xipunum_erp_store AS a " &
                    "INNER JOIN xipunum_erp_store_order AS b ON b.id = a.order_id " &
                    "LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id " &
                    "LEFT JOIN xipunum_erp_specs AS e ON e.id = a.specification_id " &
                    "LEFT JOIN xipunum_erp_type AS g ON g.id = b.factory " &
                    "LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang " &
                    $"WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}' {shopPermission} " &
                    "ORDER BY b.creationtime DESC"
            Case Else
                Return ""
        End Select
    End Function

    ' ========== 3. 商品退库报表 SQL ==========
    Public Function GetReturnReportSQL(startDate As String, endDate As String, shopPermission As String) As String
        Return $"SELECT b.retrea_umber, b.creationtime, " &
            "CAST(COALESCE(SUM(a.quantity),0) AS DECIMAL(30,2)) AS shuliang, " &
            "CAST(COALESCE(SUM(a.jin_zhong),0) AS DECIMAL(30,3)) AS jinzhong, " &
            "CAST(COALESCE(SUM(a.weight),0) AS DECIMAL(30,3)) AS zhongliang, " &
            "h.title AS kufang " &
            "FROM xipunum_erp_tuiku AS a " &
            "INNER JOIN xipunum_erp_tuiku_order AS b ON b.id = a.order_id " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang " &
            $"WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}' {shopPermission} " &
            "GROUP BY b.id ORDER BY b.creationtime DESC"
    End Function

    ' ========== 4. 商品调拨报表 SQL ==========
    Public Function GetTransferReportSQL(startDate As String, endDate As String, shopPermission As String) As String
        Return $"SELECT b.transfer_umber, b.creationtime, b.type, " &
            "hy.title AS yuankufang, hx.title AS xinkufang, " &
            "CAST(COALESCE(SUM(a.quantity),0) AS DECIMAL(30,2)) AS shuliang, " &
            "CAST(COALESCE(SUM(a.net_weight),0) AS DECIMAL(30,3)) AS jinzhong, " &
            "CAST(COALESCE(SUM(a.weight),0) AS DECIMAL(30,3)) AS zhongliang " &
            "FROM xipunum_erp_transfer AS a " &
            "INNER JOIN xipunum_erp_transfer_order AS b ON b.id = a.order_id " &
            "LEFT JOIN xipunum_erp_type AS hy ON hy.id = b.ykufang " &
            "LEFT JOIN xipunum_erp_type AS hx ON hx.id = b.xkufang " &
            $"WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}' {shopPermission} " &
            "GROUP BY b.id ORDER BY b.creationtime DESC"
    End Function

    ' ========== 5. 商品回收报表 SQL ==========
    Public Function GetRecoveryReportSQL(startDate As String, endDate As String, shopPermission As String) As String
        Return $"SELECT b.retrea_umber, b.creationtime, k.name AS khname, " &
            "CAST(COALESCE(SUM(a.jin_zhong),0) AS DECIMAL(30,3)) AS jinzhong, " &
            "CAST(COALESCE(SUM(a.retreat_amount),0) AS DECIMAL(30,2)) AS huishou " &
            "FROM xipunum_erp_retreat AS a " &
            "INNER JOIN xipunum_erp_retreat_order AS b ON b.id = a.order_id " &
            "LEFT JOIN xipunum_erp_member AS k ON k.customer_code = b.customer_code " &
            $"WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}' {shopPermission} " &
            "GROUP BY b.id ORDER BY b.creationtime DESC"
    End Function

    ' ========== 6. 商品查询报表 SQL ==========
    Public Function GetProductQueryReportSQL(searchText As String) As String
        Dim searchFilter As String = ""
        If Not String.IsNullOrEmpty(searchText) Then
            searchFilter = $" AND (a.poduct_code LIKE '%{SafeSQL(searchText)}%' OR a.product_name LIKE '%{SafeSQL(searchText)}%' OR a.item_number LIKE '%{SafeSQL(searchText)}%')"
        End If

        Return $"SELECT a.poduct_code, a.product_name, a.item_number AS kuanhao, " &
            "d.title AS pinlei, e.title AS guige, a.caizhi, a.single AS danzhong, " &
            "CAST(COALESCE(SUM(b.quantity),0) AS DECIMAL(30,2)) AS kucunshuliang, " &
            "CAST(COALESCE(SUM(b.jinzhong),0) AS DECIMAL(30,3)) AS kucunjinzhong " &
            "FROM xipunum_erp_shop AS a " &
            "LEFT JOIN xipunum_erp_shop_kucun AS b ON b.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id " &
            "LEFT JOIN xipunum_erp_specs AS e ON e.id = a.specification_id " &
            $"WHERE 1=1 {searchFilter} " &
            "GROUP BY a.poduct_code ORDER BY a.id DESC LIMIT 1000"
    End Function

    ' ========== 7. 月销售统计 SQL ==========
    Public Function GetMonthlySalesSQL(startDate As String, endDate As String, shopPermission As String) As String
        Return $"SELECT DATE_FORMAT(a.creationtime, '%Y-%m') AS yuefen, " &
            "CAST(COALESCE(SUM(a.quantity),0) AS DECIMAL(30,2)) AS xsshu, " &
            "CAST(COALESCE(SUM(a.net_weight),0) AS DECIMAL(30,3)) AS xszhong, " &
            "CAST(COALESCE(SUM(a.xiao_amount),0) AS DECIMAL(30,2)) AS xsjine, " &
            "CAST(COALESCE(SUM(a.settlement),0) AS DECIMAL(30,2)) AS shishou " &
            "FROM xipunum_erp_outbound AS a " &
            $"WHERE a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' {shopPermission} " &
            "GROUP BY DATE_FORMAT(a.creationtime, '%Y-%m') ORDER BY yuefen DESC"
    End Function

    ' ========== 8. 员工绩效 SQL ==========
    Public Function GetEmployeePerformanceSQL(startDate As String, endDate As String) As String
        Return $"SELECT h.title AS kufang, i.title AS gangwei, a.shopping_guide AS zhanghu, m.name AS daogou, a.pling AS piling, " &
            "CAST(COALESCE(SUM(a.quantity),0) AS DECIMAL(30,2)) AS xsshu, " &
            "CAST(COALESCE(SUM(a.net_weight),0) AS DECIMAL(30,3)) AS xszhong, " &
            "CAST(COALESCE(SUM(a.settlement),0) AS DECIMAL(30,2)) AS xsjine " &
            "FROM xipunum_erp_outbound AS a " &
            "INNER JOIN xipunum_erp_user AS m ON m.user = a.shopping_guide " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = m.department " &
            "LEFT JOIN xipunum_erp_type AS i ON i.id = m.post " &
            $"WHERE a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' " &
            "GROUP BY a.shopping_guide, a.pling ORDER BY h.title, i.title"
    End Function

    ' ========== 9. 店铺收支报表 SQL ==========
    Public Function GetShopFinanceSQL(startDate As String, endDate As String) As String
        Return $"SELECT DATE_FORMAT(a.creationtime, '%Y-%m-%d') AS date, h.title AS kufang, " &
            "SUM(CASE WHEN a.type='收入' THEN a.amount ELSE 0 END) AS shouru, " &
            "SUM(CASE WHEN a.type='支出' THEN a.amount ELSE 0 END) AS zhichu " &
            "FROM xipunum_erp_finance AS a " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = a.kufang " &
            $"WHERE a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' " &
            "GROUP BY DATE_FORMAT(a.creationtime, '%Y-%m-%d'), a.kufang ORDER BY date DESC"
    End Function

    ' ========== 10. 实时库存 SQL ==========
    Public Function GetRealtimeInventorySQL(shopPermission As String) As String
        Return $"SELECT a.poduct_code, a.product_name, a.item_number, " &
            "d.title AS pinlei, e.title AS guige, a.caizhi, a.single AS danzhong, " &
            "CAST(COALESCE(SUM(b.quantity),0) AS DECIMAL(30,2)) AS shuliang, " &
            "CAST(COALESCE(SUM(b.jinzhong),0) AS DECIMAL(30,3)) AS jinzhong, " &
            "h.title AS kufang " &
            "FROM xipunum_erp_shop AS a " &
            "INNER JOIN xipunum_erp_shop_kucun AS b ON b.poduct_code = a.poduct_code " &
            "LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id " &
            "LEFT JOIN xipunum_erp_specs AS e ON e.id = a.specification_id " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang " &
            $"WHERE (b.quantity > 0 OR b.jinzhong > 0) {shopPermission} " &
            "GROUP BY a.poduct_code, b.kufang ORDER BY a.id DESC"
    End Function

    ' ========== 11. 会员列表 SQL ==========
    Public Function GetMemberListSQL(searchText As String) As String
        Dim searchFilter As String = ""
        If Not String.IsNullOrEmpty(searchText) Then
            searchFilter = $" AND (a.customer_code LIKE '%{SafeSQL(searchText)}%' OR a.name LIKE '%{SafeSQL(searchText)}%' OR a.tel LIKE '%{SafeSQL(searchText)}%')"
        End If

        Return $"SELECT a.customer_code, a.memberid, a.name, a.tel, a.dizhi, a.creationtime " &
            "FROM xipunum_erp_member AS a " &
            $"WHERE 1=1 {searchFilter} ORDER BY a.creationtime DESC"
    End Function

    ' ========== 12. 款式汇总 SQL ==========
    Public Function GetStyleSummarySQL(searchText As String) As String
        Dim searchFilter As String = ""
        If Not String.IsNullOrEmpty(searchText) Then
            searchFilter = $" AND (a.title LIKE '%{SafeSQL(searchText)}%' OR a.kuanhao LIKE '%{SafeSQL(searchText)}%')"
        End If

        Return $"SELECT a.id, a.kuanhao, a.title, a.caizhi, a.chengse, " &
            "b.title AS pinlei, c.title AS guige " &
            "FROM xipunum_erp_ksiamges AS a " &
            "LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id " &
            "LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id " &
            $"WHERE 1=1 {searchFilter} ORDER BY a.id DESC"
    End Function

    ' ========== 13. 收银汇总 SQL ==========
    Public Function GetCashierSummarySQL(startDate As String, endDate As String) As String
        Return $"SELECT h.title AS kufang, b.name AS pay_name, " &
            "CAST(COALESCE(SUM(a.amount),0) AS DECIMAL(30,2)) AS amount " &
            "FROM xipunum_erp_shoukuan AS a " &
            "INNER JOIN xipunum_erp_outbound_order AS o ON o.id = a.order_id " &
            "LEFT JOIN xipunum_erp_pay AS b ON b.id = a.pay_type " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = o.kufang " &
            $"WHERE a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' " &
            "GROUP BY o.kufang, a.pay_type ORDER BY h.title"
    End Function

    ' ========== 14. 导购回收 SQL ==========
    Public Function GetGuideRecoverySQL(startDate As String, endDate As String) As String
        Return $"SELECT a.shopping_guide AS zhanghu, m.name AS daogou, h.title AS kufang, " &
            "CAST(COALESCE(SUM(a.jin_zhong),0) AS DECIMAL(30,3)) AS jinzhong, " &
            "CAST(COALESCE(SUM(a.retreat_amount),0) AS DECIMAL(30,2)) AS huishou " &
            "FROM xipunum_erp_retreat AS a " &
            "INNER JOIN xipunum_erp_user AS m ON m.user = a.shopping_guide " &
            "LEFT JOIN xipunum_erp_type AS h ON h.id = m.department " &
            $"WHERE a.creationtime >= '{startDate}' AND a.creationtime < '{endDate}' " &
            "GROUP BY a.shopping_guide ORDER BY h.title, m.name"
    End Function

    ' ========== 15. 对照报表 SQL ==========
    Public Function GetComparisonReportSQL(date1 As String, date2 As String) As String
        Return $"SELECT '入库' AS xiangmu, " &
            $"SUM(CASE WHEN DATE(a.creationtime) = '{date1}' THEN a.quantity ELSE 0 END) AS zuori_shuliang, " &
            $"SUM(CASE WHEN DATE(a.creationtime) = '{date1}' THEN a.net_weight ELSE 0 END) AS zuori_jinzhong, " &
            $"SUM(CASE WHEN DATE(a.creationtime) = '{date2}' THEN a.quantity ELSE 0 END) AS jinri_shuliang, " &
            $"SUM(CASE WHEN DATE(a.creationtime) = '{date2}' THEN a.net_weight ELSE 0 END) AS jinri_jinzhong " &
            "FROM xipunum_erp_store AS a " &
            $"WHERE DATE(a.creationtime) IN ('{date1}', '{date2}')"
    End Function

End Module

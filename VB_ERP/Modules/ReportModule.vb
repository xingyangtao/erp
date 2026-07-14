' ============================================================================
' 报表查询模块
' 包含所有报表的SQL查询逻辑
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data

Module ReportModule

    ' ========== 商品销售报表查询 ==========

    ' 订单视图查询（无订单编码时）
    Public Function GetSalesOrderView(startDate As String, endDate As String,
                                       shopPermission As String, specPermission As String,
                                       customerPermission As String, plingFilter As String,
                                       goldWeightFilter As String, searchText As String) As DataTable
        Dim sql As String = $"SELECT tol.djdingdanbian AS chukudanhao,
                                tol.chukutime AS chukutime,
                                tol.djkhbianma AS khbianma,
                                tol.djkhname AS khname,
                                tol.djkhtel AS khdiahua,
                                tol.djysdanhao AS ysdanhao,
                                CAST(ROUND(COALESCE(tol.djdingjin,0), 2) AS DECIMAL(30,2)) AS dingjin,
                                CAST(ROUND(COALESCE(sum(tol.xsshuliang),0), 2) AS DECIMAL(30,2)) AS shuliang,
                                CAST(ROUND(COALESCE(sum(tol.xsjinzhong),0), 3) AS DECIMAL(30,3)) AS jinzhong,
                                CAST(ROUND(COALESCE(tol.hsjinzhong,0), 3) AS DECIMAL(30,3)) AS hsjinzhong,
                                CAST(ROUND(COALESCE(sum(tol.zongzhong),0), 3) AS DECIMAL(30,3)) AS zhongliang,
                                CAST(ROUND(COALESCE(sum(tol.xiaoshoujia),0), 2) AS DECIMAL(30,2)) AS xiaoshou,
                                CAST(ROUND(COALESCE(sum(tol.yingshou),0), 2) AS DECIMAL(30,2)) AS yingshou,
                                CAST(ROUND(COALESCE(tol.hsjine,0), 2) AS DECIMAL(30,2)) AS hsjine,
                                CAST(ROUND(COALESCE(tol.sdjhishou,0), 2) AS DECIMAL(30,2)) AS shishou,
                                CAST(ROUND(COALESCE(sum(tol.chengbengf*tol.xsjinzhong),0), 2) AS DECIMAL(30,2)) AS cbgongfei,
                                CAST(ROUND(COALESCE(sum(tol.chengbenfj),0), 2) AS DECIMAL(30,2)) AS cbfujia,
                                CAST(ROUND(COALESCE(sum(tol.chengben),0), 2) AS DECIMAL(30,2)) AS chengben,
                                CAST(ROUND(COALESCE(sum(tol.xiaoshougf*tol.xsjinzhong),0), 2) AS DECIMAL(30,2)) AS xsgongfei,
                                CAST(ROUND(COALESCE(sum(tol.xiaoshoufj),0), 2) AS DECIMAL(30,2)) AS xsfujia,
                                tol.djyewu AS yewu,
                                CAST(ROUND(COALESCE(sum(tol.djshuijin),0), 2) AS DECIMAL(30,2)) AS shuie,
                                tol.djccjuser AS caozuo,
                                tol.djpiling AS piling,
                                CAST(ROUND(COALESCE(sum(tol.xiaoshougf*tol.xsjinzhong)-sum(tol.chengbengf*tol.xsjinzhong)-sum(tol.chengbenfj)+sum(tol.xiaoshoufj),0), 2) AS DECIMAL(30,2)) AS gflirun,
                                tol.pname AS pname,
                                tol.fxsfactory AS fxsfactory,
                                tol.youhui AS youhui
                            FROM (
                                SELECT COALESCE(e1.id, e2.id, '0') AS guigeid,
                                    b.kufang AS kufangid,
                                    b.creationtime AS chukutime,
                                    f.settlement_number AS djdingdanbian,
                                    f.customer_code as djkhbianma,
                                    k.`name` as djkhname,
                                    k.tel as djkhtel,
                                    f.presale_number as djysdanhao,
                                    o.deposit as djdingjin,
                                    l.jin_zhong as hsjinzhong,
                                    l.settlement as hsjine,
                                    f.salesman as djyewu,
                                    f.settlement as sdjhishou,
                                    f.taxamount as djshuijin,
                                    CONCAT(a.cjuser, '(', j.name, ')') AS djccjuser,
                                    f.pling as djpiling,
                                    b.poduct_code as bianma,
                                    a.product_name as mingcheng,
                                    CASE WHEN b.kufang = '0' THEN '总库' ELSE h.title END AS kufang,
                                    g.title as gongchang,
                                    a.item_number as kuanhao,
                                    CASE WHEN COALESCE(d.title, '') = '' THEN '未匹配' ELSE d.title END AS pinlei,
                                    COALESCE(e1.title, e2.title, '未匹配') AS guige,
                                    a.caizhi as caizhi,
                                    CAST(ROUND(COALESCE(a.single,0), 3) AS DECIMAL(30,3)) as danzhong,
                                    CAST(ROUND(COALESCE(b.quantity,0), 2) AS DECIMAL(30,2)) as xsshuliang,
                                    CAST(ROUND(COALESCE(b.net_weight,0), 3) AS DECIMAL(30,3)) as xsjinzhong,
                                    CAST(ROUND(COALESCE(CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN b.net_weight ELSE a.weight / NULLIF(a.quantity,0) * b.quantity END,0), 3) AS DECIMAL(30,3)) AS zongzhong,
                                    CAST(ROUND(COALESCE(b.gold_price,0), 2) AS DECIMAL(30,2)) as kejia,
                                    CAST(ROUND(COALESCE(b.xiao_amount,0), 2) AS DECIMAL(30,2)) as xiaoshoujia,
                                    CAST(ROUND(COALESCE(b.settlement,0), 2) AS DECIMAL(30,2)) as yingshou,
                                    CAST(ROUND(COALESCE(c.basic_cost,0), 2) AS DECIMAL(30,2)) as chengbengf,
                                    CAST(ROUND(COALESCE(c.company_surcharge,0), 2) AS DECIMAL(30,2)) as chengbenfj,
                                    CAST(ROUND(COALESCE(c.cost_price,0), 2) AS DECIMAL(30,2)) as chengben,
                                    CAST(ROUND(COALESCE(b.sales_cost,0), 2) AS DECIMAL(30,2)) as xiaoshougf,
                                    CAST(ROUND(COALESCE(b.sales_surcharge,0), 2) AS DECIMAL(30,2)) as xiaoshoufj,
                                    b.zhekou as zhekou,
                                    a.quandu as quankou,
                                    c.company_condition as chengse,
                                    a.sales_unit as danwei,
                                    m.`name` as daogou,
                                    b.pling as xqpiling,
                                    p.name AS pname,
                                    f.xsfactory AS fxsfactory,
                                    b.youhui AS youhui
                                FROM xipunum_erp_outbound AS b
                                INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code
                                INNER JOIN xipunum_erp_outbound_order AS f ON f.id = b.order_id
                                LEFT JOIN xipunum_erp_member AS k ON k.customer_code = f.customer_code
                                LEFT JOIN xipunum_erp_store AS c ON c.poduct_code = b.poduct_code
                                LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id
                                LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = (SELECT specification_id FROM xipunum_erp_ksiamges WHERE kuanhao = a.item_number AND a.item_number != '' LIMIT 1)
                                LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id
                                LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang
                                LEFT JOIN xipunum_erp_type AS g ON g.id = c.factory
                                LEFT JOIN xipunum_erp_user AS j ON j.user = a.cjuser
                                LEFT JOIN xipunum_erp_user AS m ON m.user = b.shopping_guide
                                LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = a.item_number
                                LEFT JOIN xipunum_erp_pay AS p ON p.id = b.pay_type
                                WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}'
                                    {shopPermission} {customerPermission} {plingFilter} {goldWeightFilter}
                            ) AS tol
                            {specPermission}
                            GROUP BY tol.djdingdanbian
                            ORDER BY tol.chukutime DESC"

        Return ExecuteQuery(sql)
    End Function

    ' 订单视图查询（有订单编码时）/ 明细视图查询
    Public Function GetSalesDetailView(startDate As String, endDate As String,
                                        shopPermission As String, specPermission As String,
                                        customerPermission As String, plingFilter As String,
                                        goldWeightFilter As String, searchText As String,
                                        Optional orderCode As String = "") As DataTable
        Dim orderFilter As String = ""
        If Not String.IsNullOrEmpty(orderCode) Then
            orderFilter = $" AND f.settlement_number='{SafeSQL(orderCode)}'"
        End If

        Dim searchFilter As String = ""
        If Not String.IsNullOrEmpty(searchText) Then
            searchFilter = $" AND (a.product_name LIKE '%{SafeSQL(searchText)}%' OR a.poduct_code LIKE '%{SafeSQL(searchText)}%' OR a.item_number LIKE '%{SafeSQL(searchText)}%')"
        End If

        Dim sql As String = $"SELECT COALESCE(e1.id, e2.id, '0') AS guigeid,
                                b.kufang AS kufangid,
                                b.creationtime AS chukutime,
                                f.settlement_number AS djdingdanbian,
                                f.customer_code as djkhbianma,
                                k.`name` as djkhname,
                                k.tel as djkhtel,
                                f.presale_number as djysdanhao,
                                o.deposit as djdingjin,
                                l.jin_zhong as hsjinzhong,
                                l.settlement as hsjine,
                                f.salesman as djyewu,
                                f.settlement as sdjhishou,
                                f.taxamount as djshuijin,
                                CONCAT(a.cjuser, '(', j.name, ')') AS djccjuser,
                                f.pling as djpiling,
                                b.poduct_code as bianma,
                                a.product_name as mingcheng,
                                CASE WHEN b.kufang = '0' THEN '总库' ELSE h.title END AS kufang,
                                g.title as gongchang,
                                a.item_number as kuanhao,
                                CASE WHEN COALESCE(d.title, '') = '' THEN '未匹配' ELSE d.title END AS pinlei,
                                COALESCE(e1.title, e2.title, '未匹配') AS guige,
                                a.caizhi as caizhi,
                                a.single as danzhong,
                                b.quantity as xsshuliang,
                                b.net_weight as xsjinzhong,
                                CAST(ROUND(CASE WHEN COALESCE(i.lingxiao, '') = '是' THEN b.net_weight ELSE a.weight / NULLIF(a.quantity, 0) * b.quantity END, 3) AS DECIMAL(30,3)) AS zongzhong,
                                b.gold_price as kejia,
                                b.xiao_amount as xiaoshoujia,
                                b.settlement as yingshou,
                                c.basic_cost as chengbengf,
                                c.company_surcharge as chengbenfj,
                                c.cost_price as chengben,
                                b.sales_cost as xiaoshougf,
                                b.sales_surcharge as xiaoshoufj,
                                b.zhekou as zhekou,
                                a.quandu as quankou,
                                c.company_condition as chengse,
                                a.sales_unit as danwei,
                                m.`name` as daogou,
                                b.pling as xqpiling,
                                p.name AS pname,
                                f.xsfactory AS fxsfactory,
                                b.youhui AS youhui
                            FROM xipunum_erp_outbound AS b
                            INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code
                            INNER JOIN xipunum_erp_outbound_order AS f ON f.id = b.order_id
                            LEFT JOIN xipunum_erp_member AS k ON k.customer_code = f.customer_code
                            LEFT JOIN xipunum_erp_store AS c ON c.poduct_code = b.poduct_code
                            LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id
                            LEFT JOIN xipunum_erp_specs AS e1 ON e1.id = (SELECT specification_id FROM xipunum_erp_ksiamges WHERE kuanhao = a.item_number AND a.item_number != '' LIMIT 1)
                            LEFT JOIN xipunum_erp_specs AS e2 ON e2.id = a.specification_id
                            LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang
                            LEFT JOIN xipunum_erp_type AS g ON g.id = c.factory
                            LEFT JOIN xipunum_erp_user AS j ON j.user = a.cjuser
                            LEFT JOIN xipunum_erp_user AS m ON m.user = b.shopping_guide
                            LEFT JOIN xipunum_erp_ksiamges AS i ON i.kuanhao = a.item_number
                            LEFT JOIN xipunum_erp_pay AS p ON p.id = b.pay_type
                            WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}'
                                {shopPermission} {customerPermission} {plingFilter} {goldWeightFilter} {orderFilter} {searchFilter}
                            ORDER BY b.creationtime DESC"

        Return ExecuteQuery(sql)
    End Function

    ' 天/月/年统计查询
    Public Function GetSalesDateSummary(startDate As String, endDate As String,
                                         shopPermission As String, dateGroup As String) As DataTable
        Dim sql As String = $"SELECT {dateGroup} AS date,
                                CAST(ROUND(COALESCE(SUM(b.quantity),0), 2) AS DECIMAL(30,2)) AS shuliang,
                                CAST(ROUND(COALESCE(SUM(b.net_weight),0), 3) AS DECIMAL(30,3)) AS jinzhong,
                                CAST(ROUND(COALESCE(SUM(b.xiao_amount),0), 2) AS DECIMAL(30,2)) AS xiaoshou,
                                CAST(ROUND(COALESCE(SUM(b.settlement),0), 2) AS DECIMAL(30,2)) AS yingshou,
                                CAST(ROUND(COALESCE(SUM(b.sales_cost * b.net_weight),0), 2) AS DECIMAL(30,2)) AS xsgongfei,
                                CAST(ROUND(COALESCE(SUM(b.sales_surcharge),0), 2) AS DECIMAL(30,2)) AS xsfujia
                            FROM xipunum_erp_outbound AS b
                            WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}'
                                {shopPermission}
                            GROUP BY {dateGroup}
                            ORDER BY date DESC"

        Return ExecuteQuery(sql)
    End Function

    ' 店铺汇总查询
    Public Function GetSalesShopSummary(startDate As String, endDate As String,
                                         shopPermission As String) As DataTable
        Dim sql As String = $"SELECT h.title AS kufang,
                                CAST(ROUND(COALESCE(SUM(b.quantity),0), 2) AS DECIMAL(30,2)) AS shuliang,
                                CAST(ROUND(COALESCE(SUM(b.net_weight),0), 3) AS DECIMAL(30,3)) AS jinzhong,
                                CAST(ROUND(COALESCE(SUM(b.xiao_amount),0), 2) AS DECIMAL(30,2)) AS xiaoshou,
                                CAST(ROUND(COALESCE(SUM(b.settlement),0), 2) AS DECIMAL(30,2)) AS yingshou,
                                CAST(ROUND(COALESCE(SUM(b.sales_cost * b.net_weight),0), 2) AS DECIMAL(30,2)) AS xsgongfei,
                                CAST(ROUND(COALESCE(SUM(b.sales_surcharge),0), 2) AS DECIMAL(30,2)) AS xsfujia
                            FROM xipunum_erp_outbound AS b
                            INNER JOIN xipunum_erp_type AS h ON h.id = b.kufang
                            WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}'
                                {shopPermission}
                            GROUP BY b.kufang, h.title
                            ORDER BY h.title"

        Return ExecuteQuery(sql)
    End Function

    ' 品类汇总查询
    Public Function GetSalesCategorySummary(startDate As String, endDate As String,
                                             shopPermission As String) As DataTable
        Dim sql As String = $"SELECT d.title AS pinlei,
                                CAST(ROUND(COALESCE(SUM(b.quantity),0), 2) AS DECIMAL(30,2)) AS shuliang,
                                CAST(ROUND(COALESCE(SUM(b.net_weight),0), 3) AS DECIMAL(30,3)) AS jinzhong,
                                CAST(ROUND(COALESCE(SUM(b.xiao_amount),0), 2) AS DECIMAL(30,2)) AS xiaoshou,
                                CAST(ROUND(COALESCE(SUM(b.settlement),0), 2) AS DECIMAL(30,2)) AS yingshou,
                                CAST(ROUND(COALESCE(SUM(b.sales_cost * b.net_weight),0), 2) AS DECIMAL(30,2)) AS xsgongfei,
                                CAST(ROUND(COALESCE(SUM(b.sales_surcharge),0), 2) AS DECIMAL(30,2)) AS xsfujia
                            FROM xipunum_erp_outbound AS b
                            INNER JOIN xipunum_erp_shop AS a ON a.poduct_code = b.poduct_code
                            LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id
                            WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}'
                                {shopPermission}
                            GROUP BY a.category_id, d.title
                            ORDER BY d.title"

        Return ExecuteQuery(sql)
    End Function

    ' ========== 商品入库报表查询 ==========

    Public Function GetInboundOrderView(startDate As String, endDate As String,
                                         shopPermission As String) As DataTable
        Dim sql As String = $"SELECT b.number AS chukudanhao,
                                b.creationtime AS chukutime,
                                b.is_xiangqian AS xiangqian,
                                b.delivery_number AS songhuodanhao,
                                d.title AS pinlei,
                                b.banchengpin AS banchengpin,
                                g.title AS gongchang,
                                b.source AS laiyuan,
                                b.settlement AS jiesuan,
                                h.title AS kufang,
                                CAST(ROUND(COALESCE(SUM(a.quantity),0), 2) AS DECIMAL(30,2)) AS shuliang,
                                CAST(ROUND(COALESCE(SUM(a.net_weight),0), 3) AS DECIMAL(30,3)) AS jinzhong,
                                CAST(ROUND(COALESCE(SUM(a.weight),0), 3) AS DECIMAL(30,3)) AS zhongliang,
                                CAST(ROUND(COALESCE(SUM(a.cost_price),0), 2) AS DECIMAL(30,2)) AS chengben,
                                CONCAT(b.cjuser, '(', j.name, ')') AS caozuo,
                                b.audit_time AS shenhetime,
                                b.state AS zhuangtai,
                                b.remarks AS beizhu
                            FROM xipunum_erp_store AS a
                            INNER JOIN xipunum_erp_store_order AS b ON b.id = a.order_id
                            LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id
                            LEFT JOIN xipunum_erp_type AS g ON g.id = b.factory
                            LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang
                            LEFT JOIN xipunum_erp_user AS j ON j.user = b.cjuser
                            WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}'
                                {shopPermission}
                            GROUP BY b.id
                            ORDER BY b.creationtime DESC"

        Return ExecuteQuery(sql)
    End Function

    Public Function GetInboundDetailView(startDate As String, endDate As String,
                                          shopPermission As String) As DataTable
        Dim sql As String = $"SELECT a.poduct_code AS bianma,
                                b.creationtime AS chukutime,
                                b.number AS chukudanhao,
                                a.product_name AS mingcheng,
                                a.is_xiangqian AS xiangqian,
                                a.item_number AS kuanhao,
                                g.title AS gongchang,
                                h.title AS kufang,
                                d.title AS pinlei,
                                e.title AS guige,
                                a.caizhi AS caizhi,
                                a.single AS danzhong,
                                CAST(ROUND(COALESCE(a.quantity,0), 2) AS DECIMAL(30,2)) AS shuliang,
                                CAST(ROUND(COALESCE(a.net_weight,0), 3) AS DECIMAL(30,3)) AS jinzhong,
                                CAST(ROUND(COALESCE(a.weight,0), 3) AS DECIMAL(30,3)) AS zhongliang,
                                CAST(ROUND(COALESCE(a.basic_cost,0), 2) AS DECIMAL(30,2)) AS chengbengf,
                                CAST(ROUND(COALESCE(a.company_surcharge,0), 2) AS DECIMAL(30,2)) AS chengbenfj,
                                CAST(ROUND(COALESCE(a.cost_price,0), 2) AS DECIMAL(30,2)) AS chengben,
                                CAST(ROUND(COALESCE(a.reference_labor,0), 2) AS DECIMAL(30,2)) AS cankaogf,
                                a.quandu AS quankou,
                                a.miankuan AS miankuan,
                                a.thickness AS houdu,
                                a.factory_condition AS gongchangchengse,
                                a.company_condition AS gongsichengse,
                                a.sales_unit AS danwei,
                                a.coefficient AS xishu,
                                CAST(ROUND(COALESCE(a.gold_price,0), 2) AS DECIMAL(30,2)) AS yuanliao
                            FROM xipunum_erp_store AS a
                            INNER JOIN xipunum_erp_store_order AS b ON b.id = a.order_id
                            LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id
                            LEFT JOIN xipunum_erp_specs AS e ON e.id = a.specification_id
                            LEFT JOIN xipunum_erp_type AS g ON g.id = b.factory
                            LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang
                            WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}'
                                {shopPermission}
                            ORDER BY b.creationtime DESC"

        Return ExecuteQuery(sql)
    End Function

    ' ========== 商品退库报表查询 ==========

    Public Function GetReturnOrderView(startDate As String, endDate As String,
                                        shopPermission As String) As DataTable
        Dim sql As String = $"SELECT b.retrea_umber AS chukudanhao,
                                b.creationtime AS chukutime,
                                CAST(ROUND(COALESCE(SUM(a.quantity),0), 2) AS DECIMAL(30,2)) AS shuliang,
                                CAST(ROUND(COALESCE(SUM(a.jin_zhong),0), 3) AS DECIMAL(30,3)) AS jinzhong,
                                CAST(ROUND(COALESCE(SUM(a.weight),0), 3) AS DECIMAL(30,3)) AS zhongliang,
                                CAST(ROUND(COALESCE(SUM(a.basic_cost),0), 2) AS DECIMAL(30,2)) AS chengbengf,
                                CAST(ROUND(COALESCE(SUM(a.company_surcharge),0), 2) AS DECIMAL(30,2)) AS chengbenfj,
                                CAST(ROUND(COALESCE(SUM(a.cost_price),0), 2) AS DECIMAL(30,2)) AS chengben,
                                b.state AS zhuangtai,
                                h.title AS kufang,
                                b.remarks AS beizhu,
                                CONCAT(b.cjuser, '(', j.name, ')') AS caozuo
                            FROM xipunum_erp_tuiku AS a
                            INNER JOIN xipunum_erp_tuiku_order AS b ON b.id = a.order_id
                            LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang
                            LEFT JOIN xipunum_erp_user AS j ON j.user = b.cjuser
                            WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}'
                                {shopPermission}
                            GROUP BY b.id
                            ORDER BY b.creationtime DESC"

        Return ExecuteQuery(sql)
    End Function

    ' ========== 商品调拨报表查询 ==========

    Public Function GetTransferOrderView(startDate As String, endDate As String,
                                          shopPermission As String) As DataTable
        Dim sql As String = $"SELECT b.transfer_umber AS chukudanhao,
                                b.creationtime AS chukutime,
                                b.type AS leixing,
                                hy.title AS yuankufang,
                                hx.title AS xinkufang,
                                CAST(ROUND(COALESCE(SUM(a.quantity),0), 2) AS DECIMAL(30,2)) AS shuliang,
                                CAST(ROUND(COALESCE(SUM(a.net_weight),0), 3) AS DECIMAL(30,3)) AS jinzhong,
                                CAST(ROUND(COALESCE(SUM(a.weight),0), 3) AS DECIMAL(30,3)) AS zhongliang,
                                CAST(ROUND(COALESCE(SUM(a.basic_cost),0), 2) AS DECIMAL(30,2)) AS chengbengf,
                                CAST(ROUND(COALESCE(SUM(a.company_surcharge),0), 2) AS DECIMAL(30,2)) AS chengbenfj,
                                CAST(ROUND(COALESCE(SUM(a.cost_price),0), 2) AS DECIMAL(30,2)) AS chengben,
                                CAST(ROUND(COALESCE(SUM(a.sales_surcharge),0), 2) AS DECIMAL(30,2)) AS xsfujia,
                                b.remarks AS beizhu,
                                CONCAT(b.cjuser, '(', j.name, ')') AS caozuo
                            FROM xipunum_erp_transfer AS a
                            INNER JOIN xipunum_erp_transfer_order AS b ON b.id = a.order_id
                            LEFT JOIN xipunum_erp_type AS hy ON hy.id = b.ykufang
                            LEFT JOIN xipunum_erp_type AS hx ON hx.id = b.xkufang
                            LEFT JOIN xipunum_erp_user AS j ON j.user = b.cjuser
                            WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}'
                                {shopPermission}
                            GROUP BY b.id
                            ORDER BY b.creationtime DESC"

        Return ExecuteQuery(sql)
    End Function

    ' ========== 商品回收报表查询 ==========

    Public Function GetRecoveryOrderView(startDate As String, endDate As String,
                                          shopPermission As String) As DataTable
        Dim sql As String = $"SELECT b.retrea_umber AS chukudanhao,
                                b.creationtime AS chukutime,
                                k.customer_code AS khbianma,
                                k.`name` AS khname,
                                k.tel AS khdianhua,
                                CAST(ROUND(COALESCE(SUM(a.jin_zhong),0), 3) AS DECIMAL(30,3)) AS jinzhong,
                                CAST(ROUND(COALESCE(SUM(a.weight),0), 3) AS DECIMAL(30,3)) AS zhongliang,
                                CAST(ROUND(COALESCE(SUM(a.other_amount),0), 2) AS DECIMAL(30,2)) AS qita,
                                CAST(ROUND(COALESCE(SUM(a.retreat_amount),0), 2) AS DECIMAL(30,2)) AS huishou,
                                CAST(ROUND(COALESCE(SUM(b.settlement),0), 2) AS DECIMAL(30,2)) AS yingfu,
                                CAST(ROUND(COALESCE(SUM(b.actual_amount),0), 2) AS DECIMAL(30,2)) AS shifu,
                                CAST(ROUND(COALESCE(b.tax_rate,0), 2) AS DECIMAL(30,2)) AS shuidian,
                                CAST(ROUND(COALESCE(b.tax_amount,0), 2) AS DECIMAL(30,2)) AS shuijin,
                                m.`name` AS yewu,
                                b.remarks AS beizhu,
                                CONCAT(b.cjuser, '(', j.name, ')') AS caozuo
                            FROM xipunum_erp_retreat AS a
                            INNER JOIN xipunum_erp_retreat_order AS b ON b.id = a.order_id
                            LEFT JOIN xipunum_erp_member AS k ON k.customer_code = b.customer_code
                            LEFT JOIN xipunum_erp_user AS j ON j.user = b.cjuser
                            LEFT JOIN xipunum_erp_user AS m ON m.user = b.salesman
                            WHERE b.creationtime >= '{startDate}' AND b.creationtime < '{endDate}'
                                {shopPermission}
                            GROUP BY b.id
                            ORDER BY b.creationtime DESC"

        Return ExecuteQuery(sql)
    End Function

    ' ========== 实时库存查询 ==========

    Public Function GetRealtimeInventory(shopPermission As String,
                                          categoryPermission As String,
                                          specPermission As String) As DataTable
        Dim sql As String = $"SELECT a.poduct_code AS bianma,
                                a.product_name AS mingcheng,
                                a.item_number AS kuanhao,
                                d.title AS pinlei,
                                e.title AS guige,
                                a.caizhi AS caizhi,
                                a.single AS danzhong,
                                CAST(ROUND(COALESCE(SUM(b.quantity),0), 2) AS DECIMAL(30,2)) AS shuliang,
                                CAST(ROUND(COALESCE(SUM(b.jinzhong),0), 3) AS DECIMAL(30,3)) AS jinzhong,
                                CAST(ROUND(COALESCE(SUM(a.weight),0), 3) AS DECIMAL(30,3)) AS zhongliang,
                                h.title AS kufang
                            FROM xipunum_erp_shop AS a
                            INNER JOIN xipunum_erp_shop_kucun AS b ON b.poduct_code = a.poduct_code
                            LEFT JOIN xipunum_erp_category AS d ON d.id = a.category_id
                            LEFT JOIN xipunum_erp_specs AS e ON e.id = a.specification_id
                            LEFT JOIN xipunum_erp_type AS h ON h.id = b.kufang
                            WHERE (b.quantity > 0 OR b.jinzhong > 0)
                                {shopPermission} {categoryPermission} {specPermission}
                            GROUP BY a.poduct_code, b.kufang
                            ORDER BY a.id DESC"

        Return ExecuteQuery(sql)
    End Function

    ' ========== 会员查询 ==========

    Public Function GetMemberList(searchText As String) As DataTable
        Dim searchFilter As String = ""
        If Not String.IsNullOrEmpty(searchText) Then
            searchFilter = $" AND (a.customer_code LIKE '%{SafeSQL(searchText)}%' OR a.memberid LIKE '%{SafeSQL(searchText)}%' OR a.name LIKE '%{SafeSQL(searchText)}%' OR a.tel LIKE '%{SafeSQL(searchText)}%')"
        End If

        Dim sql As String = $"SELECT a.customer_code, a.memberid, a.name, a.tel, 
                                a.shengri, a.dizhi, a.sex, a.creationtime
                            FROM xipunum_erp_member AS a
                            WHERE 1=1 {searchFilter}
                            ORDER BY a.creationtime DESC"

        Return ExecuteQuery(sql)
    End Function

    ' ========== 通用筛选条件构建 ==========

    Public Function BuildShopPermission(field As String) As String
        If UserPermission = "全部" Then
            Return ""
        ElseIf UserPermission = "店铺" Then
            Return $" AND {field} IN ({UserShopPermission})"
        ElseIf UserPermission = "岗位" Then
            Return $" AND {field} IN ({UserShopPermission})"
        ElseIf UserPermission = "个人" Then
            Return $" AND {field} = '{UserAccount}'"
        Else
            Return $" AND {field} = '-1'"
        End If
    End Function

    Public Function BuildSpecPermission() As String
        ' 规格权限在子查询中处理
        Return ""
    End Function

    Public Function BuildPlingFilter(pling As String) As String
        If String.IsNullOrEmpty(pling) Then
            Return ""
        End If
        Return $" AND b.pling='{pling}'"
    End Function

    Public Function BuildGoldWeightFilter(minWeight As String, maxWeight As String) As String
        If String.IsNullOrEmpty(minWeight) OrElse String.IsNullOrEmpty(maxWeight) Then
            Return ""
        End If
        Return $" AND b.net_weight >= '{minWeight}' AND b.net_weight <= '{maxWeight}'"
    End Function

End Module

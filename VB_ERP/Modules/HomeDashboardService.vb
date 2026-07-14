' ============================================================================
' 首页仪表盘数据服务
' 移植自：窗口程序集_首页自动任务 / 窗口程序集_窗口_主窗口._组件数据初始化
' ============================================================================

Imports System.Data
Imports System.Drawing
Imports MySql.Data.MySqlClient

Module HomeDashboardService

    Public ReadOnly PieColors As Color() = {
        Color.FromArgb(255, 105, 180),
        Color.FromArgb(30, 144, 255),
        Color.FromArgb(0, 0, 139),
        Color.FromArgb(128, 0, 128),
        Color.FromArgb(240, 230, 140),
        Color.FromArgb(0, 0, 255),
        Color.FromArgb(64, 224, 208),
        Color.FromArgb(255, 255, 224),
        Color.FromArgb(173, 216, 230),
        Color.FromArgb(0, 139, 139)
    }

    Public Class DashboardKpi
        Public Property SalesAmount As Decimal
        Public Property SalesQty As Decimal
        Public Property SalesWeight As Decimal
        Public Property RecoveryAmount As Decimal
        Public Property RecoveryWeight As Decimal
        Public Property ReceivedAmount As Decimal
        Public Property ActualWeight As Decimal
    End Class

    Public Class TrendPoint
        Public Property Label As String = ""
        Public Property Received As Decimal
        Public Property Sales As Decimal
        Public Property Recovery As Decimal
    End Class

    Public Class CategorySlice
        Public Property Name As String = ""
        Public Property Weight As Decimal
        Public Property Color As Color
    End Class

    Public Class SalesRankRow
        Public Property Rank As Integer
        Public Property ItemNumber As String = ""
        Public Property Category As String = ""
        Public Property Spec As String = ""
        Public Property Quantity As String = ""
        Public Property Amount As String = ""
    End Class

    Public Class DashboardSnapshot
        Public Property Kpi As New DashboardKpi()
        Public Property Trend As New List(Of TrendPoint)()
        Public Property Categories As New List(Of CategorySlice)()
        Public Property Ranking As New List(Of SalesRankRow)()
        Public Property TrendMax As Decimal
    End Class

    Private Class DateRange
        Public Property Start As String
        Public Property [End] As String
        Public Property MonthStart As String
    End Class

    Private Function GetTodayRange() As DateRange
        Dim today = DateTime.Today
        Return New DateRange With {
            .Start = today.ToString("yyyy-MM-dd") & " 00:00:00",
            .End = today.AddDays(1).ToString("yyyy-MM-dd") & " 00:00:00",
            .MonthStart = today.AddMonths(-1).ToString("yyyy-MM-dd") & " 00:00:00"
        }
    End Function

    Private Function QueryDataTable(sql As String) As DataTable
        Return DatabaseModule.ExecuteQuery(sql, MySQL_ReadTask)
    End Function

    Public Function LoadHistoricalTrend() As List(Of TrendPoint)
        Dim points As New List(Of TrendPoint)()
        Dim invoiceFilter = GetSalesInvoiceFilter()
        Dim maxValue As Decimal = 0

        For i As Integer = 1 To 30
            Dim dayStart = DateTime.Today.AddDays(-30 + i)
            Dim dayEnd = dayStart.AddDays(1)
            Dim label = dayStart.ToString("MM/dd")
            Dim startText = dayStart.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim endText = dayEnd.ToString("yyyy-MM-dd") & " 00:00:00"

            Dim sql = "SELECT CAST(ROUND(COALESCE(chuku.settlement, 0), 2) AS DECIMAL(30, 2)) AS xiaoshou," &
                      "CAST(ROUND(COALESCE(huishou.retreat_amount, 0), 2) AS DECIMAL(30, 2)) AS huishou," &
                      "CAST(ROUND(COALESCE(chuku.settlement - huishou.retreat_amount, 0), 2) AS DECIMAL(30, 2)) AS shishou " &
                      "FROM (SELECT CAST(ROUND(COALESCE(SUM(a.settlement), 0), 2) AS DECIMAL(30, 2)) AS settlement " &
                      "FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_outbound_order AS x ON x.id = a.order_id " &
                      $"WHERE a.creationtime >= '{startText}' AND a.creationtime < '{endText}'{invoiceFilter} AND a.shopping_guide IN {GlobalViewSQL}) AS chuku," &
                      "(SELECT CAST(ROUND(COALESCE(SUM(retreat_amount), 0), 2) AS DECIMAL(30, 2)) AS retreat_amount " &
                      "FROM xipunum_erp_retreat " &
                      $"WHERE creationtime >= '{startText}' AND creationtime < '{endText}' AND shopping_guide IN {GlobalViewSQL}) AS huishou"

            Dim dt = QueryDataTable(sql)
            Dim point As New TrendPoint With {.Label = label}
            If dt.Rows.Count > 0 Then
                Dim row = dt.Rows(0)
                point.Sales = SafeDecimal(row("xiaoshou"))
                point.Recovery = SafeDecimal(row("huishou"))
                point.Received = SafeDecimal(row("shishou"))
            End If

            maxValue = Math.Max(maxValue, Math.Max(point.Sales, point.Recovery))
            points.Add(point)
        Next

        Return points
    End Function

    Public Function CalculateTrendMax(points As List(Of TrendPoint), Optional salesAmount As Decimal = 0, Optional recoveryAmount As Decimal = 0) As Decimal
        Dim maxValue As Decimal = 0
        For Each p In points
            maxValue = Math.Max(maxValue, Math.Max(p.Sales, p.Recovery))
        Next
        maxValue = Math.Max(maxValue, Math.Max(salesAmount, recoveryAmount))
        If maxValue <= 0 Then Return 100
        Return Math.Ceiling(maxValue * 1.1D)
    End Function

    Public Function LoadTodayKpi() As DashboardKpi
        Dim range = GetTodayRange()
        Dim invoiceFilter = GetSalesInvoiceFilter()
        Dim kpi As New DashboardKpi()

        Dim sql = "SELECT chuku.settlement AS settlement, chuku.quantity AS quantity, chuku.net_weight AS net_weight," &
                  "huishou.retreat_amount AS retreat_amount, huishou.jin_zhong AS jin_zhong," &
                  "CAST(ROUND(COALESCE(chuku.settlement - huishou.retreat_amount, 0), 2) AS DECIMAL(30, 2)) AS shishou," &
                  "CAST(ROUND(COALESCE(chuku.net_weight - huishou.jin_zhong, 0), 3) AS DECIMAL(30, 3)) AS shijijin " &
                  "FROM (SELECT CAST(ROUND(COALESCE(SUM(a.settlement), 0), 2) AS DECIMAL(30, 2)) AS settlement," &
                  "CAST(ROUND(COALESCE(SUM(a.quantity), 0), 2) AS DECIMAL(30, 2)) AS quantity," &
                  "CAST(ROUND(COALESCE(SUM(a.net_weight), 0), 3) AS DECIMAL(30, 3)) AS net_weight " &
                  "FROM xipunum_erp_outbound AS a INNER JOIN xipunum_erp_outbound_order AS x ON x.id = a.order_id " &
                  $"WHERE a.creationtime >= '{range.Start}' AND a.creationtime < '{range.End}'{invoiceFilter} AND a.shopping_guide IN {GlobalViewSQL}) AS chuku," &
                  "(SELECT CAST(ROUND(COALESCE(SUM(retreat_amount), 0), 2) AS DECIMAL(30, 2)) AS retreat_amount," &
                  "CAST(ROUND(COALESCE(SUM(jin_zhong), 0), 3) AS DECIMAL(30, 3)) AS jin_zhong " &
                  "FROM xipunum_erp_retreat " &
                  $"WHERE creationtime >= '{range.Start}' AND creationtime < '{range.End}' AND shopping_guide IN {GlobalViewSQL}) AS huishou"

        Dim dt = QueryDataTable(sql)
        If dt.Rows.Count > 0 Then
            Dim row = dt.Rows(0)
            kpi.SalesAmount = SafeDecimal(row("settlement"))
            kpi.SalesQty = SafeDecimal(row("quantity"))
            kpi.SalesWeight = SafeDecimal(row("net_weight"))
            kpi.RecoveryAmount = SafeDecimal(row("retreat_amount"))
            kpi.RecoveryWeight = SafeDecimal(row("jin_zhong"))
            kpi.ReceivedAmount = SafeDecimal(row("shishou"))
            kpi.ActualWeight = SafeDecimal(row("shijijin"))
        End If

        Return kpi
    End Function

    Public Function LoadCategoryPie() As List(Of CategorySlice)
        Dim range = GetTodayRange()
        Dim invoiceFilter = GetSalesInvoiceFilter()
        Dim slices As New List(Of CategorySlice)()

        Dim sql = "SELECT CAST(ROUND(COALESCE(SUM(a.net_weight), 0), 3) AS DECIMAL(30, 3)) AS net_weight," &
                  "CASE WHEN b.item_number != '' THEN d.title WHEN b.specification_id != '' THEN f.title ELSE '未匹配' END AS pinlei " &
                  "FROM xipunum_erp_outbound AS a " &
                  "INNER JOIN xipunum_erp_outbound_order AS x ON x.id = a.order_id " &
                  "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
                  "LEFT JOIN xipunum_erp_ksiamges AS c ON c.kuanhao = b.item_number " &
                  "LEFT JOIN xipunum_erp_category AS d ON d.id = c.category_id " &
                  "LEFT JOIN xipunum_erp_specs AS e ON e.id = b.specification_id " &
                  "LEFT JOIN xipunum_erp_category AS f ON f.id = e.category_id " &
                  $"WHERE a.creationtime >= '{range.Start}' AND a.creationtime < '{range.End}'{invoiceFilter} AND a.shopping_guide IN {GlobalViewSQL} " &
                  "GROUP BY CASE WHEN b.item_number != '' THEN d.title WHEN b.specification_id != '' THEN f.title ELSE '未匹配' END " &
                  "HAVING net_weight > 0"

        Dim dt = QueryDataTable(sql)
        Dim index As Integer = 0
        For Each row As DataRow In dt.Rows
            Dim weight = SafeDecimal(row("net_weight"))
            If weight <= 0 Then Continue For
            slices.Add(New CategorySlice With {
                .Name = GBKToUTF8(SafeString(row("pinlei"))),
                .Weight = weight,
                .Color = PieColors(index Mod PieColors.Length)
            })
            index += 1
        Next

        Return slices
    End Function

    Public Function LoadMonthlyRanking() As List(Of SalesRankRow)
        Dim range = GetTodayRange()
        Dim invoiceFilter = GetSalesInvoiceFilter()
        Dim rows As New List(Of SalesRankRow)()

        Dim sql = "SELECT CASE WHEN b.item_number != '' THEN b.item_number ELSE '无款号' END AS item_number," &
                  "CASE WHEN b.item_number != '' THEN d.title WHEN b.specification_id != '' THEN f.title ELSE '未匹配' END AS pinlei," &
                  "CASE WHEN b.item_number != '' THEN e.title WHEN b.specification_id != '' THEN e.title ELSE '未匹配' END AS guige," &
                  "CAST(ROUND(COALESCE(SUM(a.quantity), 0), 2) AS DECIMAL(30, 2)) AS quantity," &
                  "CAST(ROUND(COALESCE(SUM(a.settlement), 0), 2) AS DECIMAL(30, 2)) AS settlement " &
                  "FROM xipunum_erp_outbound AS a " &
                  "INNER JOIN xipunum_erp_outbound_order AS x ON x.id = a.order_id " &
                  "INNER JOIN xipunum_erp_shop AS b ON b.poduct_code = a.poduct_code " &
                  "LEFT JOIN xipunum_erp_ksiamges AS c ON c.kuanhao = b.item_number " &
                  "LEFT JOIN xipunum_erp_category AS d ON d.id = c.category_id " &
                  "LEFT JOIN xipunum_erp_specs AS e ON e.id = b.specification_id " &
                  "LEFT JOIN xipunum_erp_category AS f ON f.id = e.category_id " &
                  $"WHERE a.creationtime >= '{range.MonthStart}' AND a.creationtime < '{range.End}'{invoiceFilter} AND a.shopping_guide IN {GlobalViewSQL} " &
                  "GROUP BY CASE WHEN b.item_number != '' AND b.item_number IS NOT NULL THEN CONCAT('item_', b.item_number) ELSE CONCAT('spec_', COALESCE(e.category_id, ''), '_', COALESCE(e.id, '')) END," &
                  "CASE WHEN b.item_number != '' THEN b.item_number ELSE '无款号' END," &
                  "CASE WHEN b.item_number != '' THEN d.title WHEN b.specification_id != '' THEN f.title ELSE '未匹配' END," &
                  "CASE WHEN b.item_number != '' THEN d.title WHEN b.specification_id != '' THEN e.title ELSE '未匹配' END " &
                  "ORDER BY SUM(a.quantity) DESC LIMIT 10"

        Dim dt = QueryDataTable(sql)
        Dim rank As Integer = 1
        For Each row As DataRow In dt.Rows
            rows.Add(New SalesRankRow With {
                .Rank = rank,
                .ItemNumber = GBKToUTF8(SafeString(row("item_number"))),
                .Category = GBKToUTF8(SafeString(row("pinlei"))),
                .Spec = GBKToUTF8(SafeString(row("guige"))),
                .Quantity = SafeString(row("quantity")),
                .Amount = SafeString(row("settlement"))
            })
            rank += 1
        Next

        Return rows
    End Function

    Public Function LoadDashboardSnapshot(existingTrend As List(Of TrendPoint)) As DashboardSnapshot
        Dim snapshot As New DashboardSnapshot()
        snapshot.Kpi = LoadTodayKpi()
        snapshot.Categories = LoadCategoryPie()
        snapshot.Ranking = LoadMonthlyRanking()

        If existingTrend IsNot Nothing AndAlso existingTrend.Count = 30 Then
            snapshot.Trend = existingTrend
            If snapshot.Trend.Count > 0 Then
                snapshot.Trend(29).Received = snapshot.Kpi.ReceivedAmount
                snapshot.Trend(29).Sales = snapshot.Kpi.SalesAmount
                snapshot.Trend(29).Recovery = snapshot.Kpi.RecoveryAmount
            End If
        Else
            snapshot.Trend = LoadHistoricalTrend()
        End If

        snapshot.TrendMax = CalculateTrendMax(snapshot.Trend, snapshot.Kpi.SalesAmount, snapshot.Kpi.RecoveryAmount)
        Return snapshot
    End Function

End Module

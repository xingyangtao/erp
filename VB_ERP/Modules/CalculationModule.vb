' ============================================================================
' 计算模块
' 实现所有价格、重量、折扣计算逻辑
' 完全按照易语言源文件实现
' ============================================================================

Module CalculationModule

    ' ========== 销售价格计算 ==========
    ' 按照易语言源文件：窗口程序集_窗口_商品销售出库.form.e.txt
    ' 赋值 (查找销售单价, 到文本 (相除 (相加 (相乘 (到数值 (销售数据销售金重), 相加 (到数值 (销售数据销售工费), 到数值 (销售数据销售克价))), 到数值 (销售数据销售附加费)), 到数值 (销售数据销售数量))))
    ' 赋值 (查找销售金额, 到文本 (相加 (相乘 (到数值 (销售数据销售金重), 相加 (到数值 (销售数据销售工费), 到数值 (销售数据销售克价))), 到数值 (销售数据销售附加费))))
    ' 赋值 (查找实收金额, 到文本 (四舍五入 (到数值 (查找销售金额), 0)))

    ' 计算销售单价
    Public Function CalculateSalesUnitPrice(goldWeight As Decimal, goldPrice As Decimal, salesCost As Decimal, salesSurcharge As Decimal, quantity As Decimal) As Decimal
        If quantity = 0 Then Return 0
        Return (goldWeight * (goldPrice + salesCost) + salesSurcharge) / quantity
    End Function

    ' 计算销售金额
    Public Function CalculateSalesAmount(goldWeight As Decimal, goldPrice As Decimal, salesCost As Decimal, salesSurcharge As Decimal) As Decimal
        Return goldWeight * (goldPrice + salesCost) + salesSurcharge
    End Function

    ' 计算实收金额（四舍五入）
    Public Function CalculateActualAmount(salesAmount As Decimal) As Decimal
        Return Math.Round(salesAmount, 0)
    End Function

    ' 计算工费利润
    Public Function CalculateCostProfit(salesCost As Decimal, goldWeight As Decimal, basicCost As Decimal, companySurcharge As Decimal, salesSurcharge As Decimal) As Decimal
        Return salesCost * goldWeight - basicCost * goldWeight - companySurcharge + salesSurcharge
    End Function

    ' ========== 回收价格计算 ==========
    ' 按照易语言源文件：窗口程序集_窗口_商品信息回收.form.e.txt
    ' 回收金额 = 金重 × 回收克价 + 其他费用
    ' 应付金额 = 回收金额
    ' 实付金额 = 应付金额 - 税率金额
    ' 税率金额 = (销售金额 - 回收金额 - 定金) × 税率 / 100

    ' 计算回收金额
    Public Function CalculateRecoveryAmount(goldWeight As Decimal, recoveryPrice As Decimal, otherFee As Decimal) As Decimal
        Return goldWeight * recoveryPrice + otherFee
    End Function

    ' 计算应付金额
    Public Function CalculatePayableAmount(recoveryAmount As Decimal) As Decimal
        Return recoveryAmount
    End Function

    ' 计算税率金额
    Public Function CalculateTaxAmount(salesAmount As Decimal, recoveryAmount As Decimal, deposit As Decimal, taxRate As Decimal) As Decimal
        Return (salesAmount - recoveryAmount - deposit) * taxRate / 100
    End Function

    ' 计算实付金额
    Public Function CalculateActualPayAmount(payableAmount As Decimal, taxAmount As Decimal) As Decimal
        Return payableAmount - taxAmount
    End Function

    ' ========== 会员存欠计算 ==========
    ' 按照易语言源文件：窗口程序集_窗口_商品销售出库.form.e.txt
    ' 存料 = SUM(存.type='料') - SUM(欠.type='料')
    ' 存款 = SUM(存.type='元') - SUM(欠.type='元')

    ' 计算存欠料
    Public Function CalculateStockMaterial(depositMaterial As Decimal, oweMaterial As Decimal) As Decimal
        Return depositMaterial - oweMaterial
    End Function

    ' 计算存欠款
    Public Function CalculateStockAmount(depositAmount As Decimal, oweAmount As Decimal) As Decimal
        Return depositAmount - oweAmount
    End Function

    ' ========== 重量计算 ==========
    ' 按照易语言源文件：窗口程序集_窗口_商品销售出库.form.e.txt
    ' 重量 = 单件重 × 数量
    ' 金重 = 重量 (如果是零销售)

    ' 计算重量
    Public Function CalculateWeight(singleWeight As Decimal, quantity As Decimal) As Decimal
        Return singleWeight * quantity
    End Function

    ' 计算金重（零销售时）
    Public Function CalculateGoldWeightForZeroSales(netWeight As Decimal) As Decimal
        Return netWeight
    End Function

    ' 计算金重（非零销售时）
    Public Function CalculateGoldWeight(weight As Decimal, quantity As Decimal, actualQuantity As Decimal) As Decimal
        If quantity = 0 Then Return 0
        Return weight / quantity * actualQuantity
    End Function

    ' ========== 折扣计算 ==========
    ' 按照易语言源文件：窗口程序集_窗口_商品销售出库.form.e.txt
    ' 折扣 = 附加费 / 原附加费

    ' 计算折扣
    Public Function CalculateDiscount(surcharge As Decimal, originalSurcharge As Decimal) As Decimal
        If originalSurcharge = 0 Then Return 1
        Return surcharge / originalSurcharge
    End Function

    ' 计算附加费（根据折扣）
    Public Function CalculateSurchargeWithDiscount(originalSurcharge As Decimal, discount As Decimal) As Decimal
        Return originalSurcharge * discount
    End Function

    ' ========== 批量修改计算 ==========
    ' 按照易语言源文件：窗口程序集_窗口_商品销售批量修改.form.e.txt
    ' 新附加费 = 原附加费 × 折扣

    ' 计算新附加费
    Public Function CalculateNewSurcharge(originalSurcharge As Decimal, discount As Decimal) As Decimal
        Return originalSurcharge * discount
    End Function

    ' ========== 绩效计算 ==========
    ' 按照易语言源文件：窗口程序集_窗口_报表员工绩效.form.e.txt
    ' 绩效 = (销售数量/重量/金额) × 绩效系数

    ' 计算绩效
    Public Function CalculatePerformance(value As Decimal, coefficient As Decimal) As Decimal
        Return value * coefficient
    End Function

    ' ========== 通用格式化函数 ==========

    ' 格式化为2位小数
    Public Function FormatDecimal2(value As Decimal) As String
        Return value.ToString("F2")
    End Function

    ' 格式化为3位小数
    Public Function FormatDecimal3(value As Decimal) As String
        Return value.ToString("F3")
    End Function

    ' 格式化为整数
    Public Function FormatInteger(value As Decimal) As String
        Return Math.Round(value, 0).ToString()
    End Function

End Module

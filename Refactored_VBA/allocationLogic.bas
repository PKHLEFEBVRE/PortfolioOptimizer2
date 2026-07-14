Option Explicit
Option Base 1

' =========================================================================
' REFACTORED DOWNSIDE BETA & CVaR EQUITY ALLOCATION MODULE
' =========================================================================

Sub adjustForMDD()
    InitializeGlobals

    Dim DSBeta As Double, DDFactor As Double
    Dim CVaRFactor As Double, benchCVaR As Double, portCVaR As Double
    Dim SemiDevFactor As Double, benchSemiDev As Double, portSemiDev As Double
    Dim HistoricalMDDFactor As Double, benchMDD As Double, portMDD As Double
    Dim finalRiskFactor As Double

    ComputeRiskFactors DSBeta, DDFactor, _
                       CVaRFactor, benchCVaR, portCVaR, _
                       SemiDevFactor, benchSemiDev, portSemiDev, _
                       HistoricalMDDFactor, benchMDD, portMDD

    finalRiskFactor = Application.WorksheetFunction.Average(DDFactor, CVaRFactor, SemiDevFactor, HistoricalMDDFactor)

    MsgBox "--- PORTFOLIO RISK FACTORS (" & Round(config.TargetRatio * 100, 0) & "% Target) ---" & vbCr & vbCr & _
           "1. DOWNSIDE BETA: " & Round(DSBeta, 2) & " (Limit: " & Round(DDFactor * 100, 1) & "%)" & vbCr & vbCr & _
           "2. TAIL RISK (" & Round(config.CVaRConfidence * 100, 0) & "% CVaR):" & vbCr & _
           "   Bench: " & Round(benchCVaR * 100, 2) & "%  |  Port: " & Round(portCVaR * 100, 2) & "%" & vbCr & _
           "   Limit: " & Round(CVaRFactor * 100, 1) & "%)" & vbCr & vbCr & _
           "3. VOLATILITY (Semi-Dev < 0%):" & vbCr & _
           "   Bench: " & Round(benchSemiDev * 100, 2) & "%  |  Port: " & Round(portSemiDev * 100, 2) & "%" & vbCr & _
           "   Limit: " & Round(SemiDevFactor * 100, 1) & "%)" & vbCr & vbCr & _
           "4. MAX DRAWDOWN:" & vbCr & _
           "   Bench: " & Round(benchMDD * 100, 2) & "%  |  Port: " & Round(portMDD * 100, 2) & "%" & vbCr & _
           "   Limit: " & Round(HistoricalMDDFactor * 100, 1) & "%)" & vbCr & vbCr & _
           "AVERAGED FINAL ALLOCATION: " & Round(finalRiskFactor * 100, 2) & "%", vbInformation, "Risk Metrics Updated"

    Dim adjustementFactor As Double
    adjustementFactor = 1#

    adjustementFactor = adjustForConvictions(DDFactor)

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(config.DashSheet)
    wsDash.Range("EqWeightingLimit").Value = adjustementFactor
End Sub

Private Function adjustForConvictions(ByVal DDFactor As Double) As Double
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim convictionsSum As Double

    Set wb = ThisWorkbook
    Set wsInput = wb.Sheets(config.DashSheet)

    Dim convData As Dictionary
    Set convData = convictionsUtils.getConvictions()

    Dim keysArray() As Variant
    keysArray = convData.Keys
    Dim totalKeys As Long
    totalKeys = UBound(keysArray) - LBound(keysArray) + 1

    Dim v() As Double
    Dim w() As Double
    ReDim v(1 To totalKeys)
    ReDim w(1 To totalKeys)

    Dim k As Long
    For k = LBound(keysArray) To UBound(keysArray)
        Dim val As Double
        Dim weight As Double
        val = Empty
        weight = 0#

        Dim assetData As Object
        Set assetData = convData(keysArray(k))

        val = assetData("TP")
        weight = assetData("conviction")
        If weight > 0# Then
            v(k - LBound(keysArray) + 1) = val
            w(k - LBound(keysArray) + 1) = weight
        End If
    Next k


    Dim adjustementFactor As Double
    adjustementFactor = 1#
    If UBound(v) > 0 Then
        adjustementFactor = 1 + solverUtils.SolveForT(w, v)
    End If

    If adjustementFactor > 1 Then
        adjustementFactor = 1#
    End If

    If adjustementFactor <= 0.1 Then
        adjustementFactor = 0.1
    End If

    adjustForConvictions = DDFactor * adjustementFactor
End Function

Private Sub ComputeRiskFactors(ByRef DSBeta As Double, ByRef DDFactor As Double, _
                               ByRef CVaRFactor As Double, ByRef benchCVaR As Double, ByRef portCVaR As Double, _
                               ByRef SemiDevFactor As Double, ByRef benchSemiDev As Double, ByRef portSemiDev As Double, _
                               ByRef HistoricalMDDFactor As Double, ByRef benchMDD As Double, ByRef portMDD As Double)
    Dim wb As Workbook
    Dim wsDash As Worksheet, wsChart As Worksheet, wsBench As Worksheet

    Set wb = ThisWorkbook
    Set wsDash = wb.Sheets(config.DashSheet)
    Set wsChart = wb.Sheets(config.ChartSheet)

    On Error Resume Next
    Set wsBench = wb.Sheets("O_SPY")
    If wsBench Is Nothing Then Set wsBench = wb.Sheets("O_QQQ")
    On Error GoTo 0

    If wsBench Is Nothing Then
        MsgBox "Cannot find benchmark sheet for risk limits.", vbCritical
        Exit Sub
    End If

    Dim meanColChart As Long
    meanColChart = GetColumnIndexByName(wsChart, "Mean")

    If meanColChart = 0 Then
        MsgBox "'Mean' column missing in ChartData.", vbCritical
        Exit Sub
    End If

    ' Pre-load and cache the data into PortfolioData class
    Dim matchedChartPrices() As Double, matchedBenchPrices() As Double
    Dim numMatched As Long
    numMatched = LoadAndAlignData(wsChart, meanColChart, wsBench, matchedChartPrices, matchedBenchPrices)

    If numMatched >= 20 Then
        Dim chartRets() As Double, benchRets() As Double
        chartRets = CalculateReturns(matchedChartPrices, numMatched)
        benchRets = CalculateReturns(matchedBenchPrices, numMatched)
        portDataCache.SetReturns chartRets, benchRets, numMatched
    End If

    ' 1. Downside Beta
    DSBeta = ComputeEmaDownsideBeta(config.HalfLife)
    If DSBeta > 0 And DSBeta > config.TargetRatio Then
        DDFactor = config.TargetRatio / DSBeta
    Else
        DDFactor = 1
    End If
    DDFactor = Application.WorksheetFunction.Min(DDFactor, 0.95)

    ' 2. Tail Risk (CVaR)
    CVaRFactor = ComputeCVaRFactor(benchCVaR, portCVaR, config.CVaRConfidence, config.TargetRatio)

    ' 3. Semi-Deviation
    SemiDevFactor = ComputeSemiDevFactor(benchSemiDev, portSemiDev, 0, config.TargetRatio)

    ' 4. Historical Max Drawdown
    HistoricalMDDFactor = ComputeHistoricalMDDFactor(benchMDD, portMDD, config.TargetRatio)

End Sub

Private Function ComputeCVaRFactor(ByRef benchCVaR As Double, ByRef portCVaR As Double, Optional confidence As Double = 0.95, Optional targetRatio As Double = 0.7) As Double
    Dim tempFactor As Double
    If portDataCache.IsLoaded And portDataCache.NumMatched >= 20 Then
        benchCVaR = CalculateHistoricalCVaR(portDataCache.BenchReturns, confidence)
        portCVaR = CalculateHistoricalCVaR(portDataCache.ChartReturns, confidence)

        Dim targetCVaR As Double
        targetCVaR = benchCVaR * targetRatio

        If portCVaR < 0 And targetCVaR < 0 Then
            tempFactor = targetCVaR / portCVaR
        Else
            tempFactor = 1
        End If
        ComputeCVaRFactor = Application.WorksheetFunction.Min(tempFactor, 0.95)
    Else
        benchCVaR = 0: portCVaR = 0
        ComputeCVaRFactor = 0.95
    End If
End Function

Private Function GetColumnIndexByName(ws As Worksheet, colName As String) As Long
    Dim c As Range
    Set c = ws.Rows(1).Find(What:=colName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not c Is Nothing Then
        GetColumnIndexByName = c.Column
    Else
        GetColumnIndexByName = 0
    End If
End Function

Private Function ComputeEmaDownsideBeta(halfLife As Double) As Double
    If Not portDataCache.IsLoaded Or portDataCache.NumMatched < 20 Then
        ComputeEmaDownsideBeta = 0
        Exit Function
    End If

    Dim chartRets() As Double, benchRets() As Double
    chartRets = portDataCache.ChartReturns
    benchRets = portDataCache.BenchReturns
    Dim numMatched As Long
    numMatched = portDataCache.NumMatched

    Dim emaWeights() As Double, sumW As Double, sumW2 As Double
    Dim numDownside As Long
    numDownside = CalculateDownsideEmaWeights(benchRets, halfLife, numMatched - 1, emaWeights, sumW, sumW2)

    If numDownside < 2 Or sumW <= 0 Then
        ComputeEmaDownsideBeta = 0
        Exit Function
    End If

    Dim meanBr As Double, meanPr As Double
    meanBr = CalculateWeightedMean(benchRets, emaWeights, numMatched - 1, sumW)
    meanPr = CalculateWeightedMean(chartRets, emaWeights, numMatched - 1, sumW)

    Dim cov As Double, var As Double
    If Not CalculateWeightedCovarianceAndVariance(benchRets, chartRets, emaWeights, numMatched - 1, meanBr, meanPr, sumW, sumW2, cov, var) Then
        ComputeEmaDownsideBeta = 0
        Exit Function
    End If

    ComputeEmaDownsideBeta = cov / var
End Function

Private Function LoadAndAlignData(wsChart As Worksheet, meanColChart As Long, wsBench As Worksheet, ByRef matchedChartPrices() As Double, ByRef matchedBenchPrices() As Double) As Long
    Dim lastRowChart As Long, lastRowBench As Long
    lastRowChart = wsChart.Cells(wsChart.Rows.count, "A").End(xlUp).Row
    lastRowBench = wsBench.Cells(wsBench.Rows.count, "A").End(xlUp).Row

    If lastRowChart < 2 Or lastRowBench < 2 Then
        LoadAndAlignData = 0
        Exit Function
    End If

    Dim chartDates() As Variant, chartPrices() As Variant
    chartDates = wsChart.Range("A2:A" & lastRowChart).Value
    chartPrices = wsChart.Range(wsChart.Cells(2, meanColChart), wsChart.Cells(lastRowChart, meanColChart)).Value

    Dim benchDates() As Variant, benchPrices() As Variant
    benchDates = wsBench.Range("A2:A" & lastRowBench).Value
    benchPrices = wsBench.Range("B2:B" & lastRowBench).Value

    ReDim matchedChartPrices(1 To Application.WorksheetFunction.Min(lastRowChart, lastRowBench))
    ReDim matchedBenchPrices(1 To UBound(matchedChartPrices))

    Dim i As Long, j As Long, count As Long
    count = 0
    j = 1

    For i = 1 To UBound(chartDates, 1)
        Dim cDate As Date
        If IsDate(chartDates(i, 1)) Then
            cDate = CDate(chartDates(i, 1))
            Do While j <= UBound(benchDates, 1) And CDate(benchDates(j, 1)) > cDate
                j = j + 1
            Loop

            If j <= UBound(benchDates, 1) And CDate(benchDates(j, 1)) = cDate Then
                Dim cp As Double, bp As Double
                cp = SafeCDbl(chartPrices(i, 1))
                bp = SafeCDbl(benchPrices(j, 1))

                If cp > 0 And bp > 0 Then
                    count = count + 1
                    matchedChartPrices(count) = cp
                    matchedBenchPrices(count) = bp
                End If
            End If
        End If
    Next i

    If count > 0 Then
        ReDim Preserve matchedChartPrices(1 To count)
        ReDim Preserve matchedBenchPrices(1 To count)
    Else
        Erase matchedChartPrices
        Erase matchedBenchPrices
    End If

    LoadAndAlignData = count
End Function

Private Function CalculateReturns(prices() As Double, count As Long) As Double()
    Dim rets() As Double
    ReDim rets(1 To count - 1)

    Dim i As Long
    For i = 1 To count - 1
        rets(i) = (prices(i) / prices(i + 1)) - 1
    Next i

    CalculateReturns = rets
End Function

Private Function CalculateDownsideEmaWeights(benchRets() As Double, halfLife As Double, numRets As Long, ByRef outWeights() As Double, ByRef outSumW As Double, ByRef outSumW2 As Double) As Long
    ReDim outWeights(1 To numRets)
    outSumW = 0
    outSumW2 = 0

    Dim lambda As Double
    lambda = Exp(-Log(2) / halfLife)

    Dim w As Double
    w = 1

    Dim i As Long
    Dim count As Long
    count = 0

    For i = 1 To numRets
        If benchRets(i) < 0 Then
            outWeights(i) = w
            outSumW = outSumW + w
            outSumW2 = outSumW2 + (w * w)
            count = count + 1
        Else
            outWeights(i) = 0
        End If
        w = w * lambda
    Next i

    CalculateDownsideEmaWeights = count
End Function

Private Function CalculateWeightedMean(rets() As Double, weights() As Double, numRets As Long, sumW As Double) As Double
    Dim i As Long
    Dim sumW_Ret As Double
    sumW_Ret = 0

    For i = 1 To numRets
        If weights(i) > 0 Then
            sumW_Ret = sumW_Ret + (weights(i) * rets(i))
        End If
    Next i

    CalculateWeightedMean = sumW_Ret / sumW
End Function

Private Function CalculateWeightedCovarianceAndVariance(benchRets() As Double, chartRets() As Double, weights() As Double, numRets As Long, meanBr As Double, meanPr As Double, sumW As Double, sumW2 As Double, ByRef outCov As Double, ByRef outVar As Double) As Boolean
    Dim sumWCov As Double, sumWVar As Double
    sumWCov = 0
    sumWVar = 0

    Dim i As Long
    Dim devBr As Double, devPr As Double

    For i = 1 To numRets
        If weights(i) > 0 Then
            devBr = benchRets(i) - meanBr
            devPr = chartRets(i) - meanPr

            sumWCov = sumWCov + (weights(i) * devBr * devPr)
            sumWVar = sumWVar + (weights(i) * devBr * devBr)
        End If
    Next i

    Dim denom As Double
    denom = sumW - (sumW2 / sumW)

    If denom <= 0 Or sumWVar = 0 Then
        CalculateWeightedCovarianceAndVariance = False
        Exit Function
    End If

    outCov = sumWCov / denom
    outVar = sumWVar / denom
    CalculateWeightedCovarianceAndVariance = True
End Function

Private Function SafeCDbl(val As Variant) As Double
    If IsEmpty(val) Or IsError(val) Then
        SafeCDbl = 0
    ElseIf IsNumeric(val) Then
        SafeCDbl = CDbl(val)
    Else
        SafeCDbl = 0
    End If
End Function

Private Function CalculateHistoricalCVaR(rets() As Double, confidence As Double) As Double
    Dim numRets As Long
    numRets = UBound(rets) - LBound(rets) + 1

    Dim sortedRets() As Double
    ReDim sortedRets(1 To numRets)
    Dim i As Long
    For i = 1 To numRets
        sortedRets(i) = rets(i)
    Next i

    Call QuickSortAscending(sortedRets, LBound(sortedRets), UBound(sortedRets))

    Dim varIndex As Long
    varIndex = Int(numRets * (1 - confidence))
    If varIndex < 1 Then varIndex = 1

    Dim sum As Double
    sum = 0
    For i = 1 To varIndex
        sum = sum + sortedRets(i)
    Next i

    CalculateHistoricalCVaR = sum / varIndex
End Function

Private Sub QuickSortAscending(arr() As Double, first As Long, last As Long)
    Dim pivot As Double, temp As Double
    Dim i As Long, j As Long

    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While arr(i) < pivot And i < last
            i = i + 1
        Loop
        Do While pivot < arr(j) And j > first
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If first < j Then QuickSortAscending arr, first, j
    If i < last Then QuickSortAscending arr, i, last
End Sub

Private Function ComputeSemiDevFactor(ByRef benchSemiDev As Double, ByRef portSemiDev As Double, Optional target As Double = 0, Optional targetRatio As Double = 0.7) As Double
    Dim tempFactor As Double
    If portDataCache.IsLoaded And portDataCache.NumMatched >= 20 Then
        benchSemiDev = CalculateSemiDeviation(portDataCache.BenchReturns, target)
        portSemiDev = CalculateSemiDeviation(portDataCache.ChartReturns, target)

        If portSemiDev > 0 And portSemiDev > (benchSemiDev * targetRatio) Then
            tempFactor = (benchSemiDev * targetRatio) / portSemiDev
        Else
            tempFactor = 1
        End If
        ComputeSemiDevFactor = Application.WorksheetFunction.Min(tempFactor, 0.95)
    Else
        benchSemiDev = 0: portSemiDev = 0
        ComputeSemiDevFactor = 0.95
    End If
End Function

Private Function ComputeHistoricalMDDFactor(ByRef benchMDD As Double, ByRef portMDD As Double, Optional targetRatio As Double = 0.7) As Double
    Dim tempFactor As Double
    If portDataCache.IsLoaded And portDataCache.NumMatched >= 20 Then
        benchMDD = CalculateMaximumDrawdown(portDataCache.BenchReturns)
        portMDD = CalculateMaximumDrawdown(portDataCache.ChartReturns)

        If portMDD < 0 And portMDD < (benchMDD * targetRatio) Then
            tempFactor = (benchMDD * targetRatio) / portMDD
        Else
            tempFactor = 1
        End If
        ComputeHistoricalMDDFactor = Application.WorksheetFunction.Min(tempFactor, 0.95)
    Else
        benchMDD = 0: portMDD = 0
        ComputeHistoricalMDDFactor = 0.95
    End If
End Function

Private Function CalculateSemiDeviation(rets() As Double, target As Double) As Double
    Dim i As Long, numRets As Long
    Dim sumSq As Double

    numRets = UBound(rets) - LBound(rets) + 1
    sumSq = 0

    For i = LBound(rets) To UBound(rets)
        If rets(i) < target Then
            sumSq = sumSq + ((rets(i) - target) ^ 2)
        End If
    Next i

    CalculateSemiDeviation = Sqr(sumSq / numRets)
End Function

Private Function CalculateMaximumDrawdown(rets() As Double) As Double
    Dim i As Long
    Dim peak As Double, current As Double, dd As Double, mdd As Double

    peak = 1
    current = 1
    mdd = 0

    For i = LBound(rets) To UBound(rets)
        current = current * (1 + rets(i))

        If current > peak Then
            peak = current
        End If

        dd = (current / peak) - 1

        If dd < mdd Then
            mdd = dd
        End If
    Next i

    CalculateMaximumDrawdown = mdd
End Function

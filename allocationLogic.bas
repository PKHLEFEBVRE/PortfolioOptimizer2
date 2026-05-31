Attribute VB_Name = "allocationLogic"
Option Explicit

' =========================================================================
' DOWNSIDE BETA & CVaR EQUITY ALLOCATION MODULE
' =========================================================================
Function ComputeFinalRiskFactor() As Double
    ' 1. Declare variables for all 4 metrics and their allocation factors
    Dim DSBeta As Double, DDFactor As Double
    Dim CVaRFactor As Double, benchCVaR As Double, portCVaR As Double
    Dim SemiDevFactor As Double, benchSemiDev As Double, portSemiDev As Double
    Dim HistoricalMDDFactor As Double, benchMDD As Double, portMDD As Double
    Dim finalRiskFactor As Double

    ' 2. Execute the isolated risk factor routing hub
    ComputeRiskFactors DSBeta, DDFactor, _
                       CVaRFactor, benchCVaR, portCVaR, _
                       SemiDevFactor, benchSemiDev, portSemiDev, _
                       HistoricalMDDFactor, benchMDD, portMDD, 256

    ' 3. Reconcile the constraints: Select the absolute strictest factor
    finalRiskFactor = Application.WorksheetFunction.Average(DDFactor, CVaRFactor, SemiDevFactor, HistoricalMDDFactor)

    ' 4. Clean dashboard readout via MsgBox (Optional, kept for user transparency)
    'MsgBox "--- PORTFOLIO RISK FACTORS (70% Target) ---" & vbCr & vbCr & _
    '       "1. DOWNSIDE BETA: " & Round(DSBeta, 2) & " (Limit: " & Round(DDFactor * 100, 1) & "%)" & vbCr & vbCr & _
    '       "2. TAIL RISK (95% CVaR):" & vbCr & _
    '       "   Bench: " & Round(benchCVaR * 100, 2) & "%  |  Port: " & Round(portCVaR * 100, 2) & "%" & vbCr & _
    '       "   Limit: " & Round(CVaRFactor * 100, 1) & "%)" & vbCr & vbCr & _
    '       "3. VOLATILITY (Semi-Dev < 0%):" & vbCr & _
    '       "   Bench: " & Round(benchSemiDev * 100, 2) & "%  |  Port: " & Round(portSemiDev * 100, 2) & "%" & vbCr & _
    '       "   Limit: " & Round(SemiDevFactor * 100, 1) & "%)" & vbCr & vbCr & _
    '       "4. HISTORICAL MAX DRAWDOWN:" & vbCr & _
    '       "   Bench: " & Round(benchMDD * 100, 2) & "%  |  Port: " & Round(portMDD * 100, 2) & "%" & vbCr & _
    '       "   Limit: " & Round(HistoricalMDDFactor * 100, 1) & "%)" & vbCr & _
    '       "----------------------------------------" & vbCr & vbCr & _
    '       "FINAL APPLIED RISK FACTOR: " & Round(finalRiskFactor * 100, 1) & "%", vbInformation, "Risk Adjustments"

    ComputeFinalRiskFactor = finalRiskFactor
End Function


Private Function adjustForConvictions(ByVal DDFactor As Double) As Double

    Dim convictions As Dictionary
    Set convictions = convictionsUtils.getConvictions

    Dim adjustementFactor As Double
    adjustementFactor = 1#

    Dim equityPositions As Dictionary
    Set equityPositions = DataUtils.GetEquityPositionsFromFund

    Dim pos As Variant

    For Each pos In equityPositions
        adjustementFactor = adjustementFactor + convictions(pos)("TP proba")
    Next pos

    adjustementFactor = adjustementFactor / equityPositions.count

    adjustForConvictions = DDFactor * adjustementFactor

End Function

' Main Function: Computes desired equity allocation based on 4 separate risk factors
Private Sub ComputeRiskFactors(ByRef DSBeta As Double, ByRef DDFactor As Double, _
                               ByRef CVaRFactor As Double, ByRef benchCVaR As Double, ByRef portCVaR As Double, _
                               ByRef SemiDevFactor As Double, ByRef benchSemiDev As Double, ByRef portSemiDev As Double, _
                               ByRef HistoricalMDDFactor As Double, ByRef benchMDD As Double, ByRef portMDD As Double, _
                               Optional halfLife As Double = 256)
    Dim wb As Workbook
    Dim wsChart As Worksheet
    Dim wsBench As Worksheet

    Set wb = ThisWorkbook
    On Error Resume Next
    Set wsChart = wb.Sheets("ChartData")
    Set wsBench = wb.Sheets("BenchData")
    On Error GoTo 0

    If wsChart Is Nothing Or wsBench Is Nothing Then
        Exit Sub
    End If

    Dim meanColChart As Long
    meanColChart = GetColumnIndexByName(wsChart, "MEAN")

    If meanColChart = 0 Then Exit Sub

    ' -------------------------------------------------------------
    ' 1. Downside Beta Logic
    ' -------------------------------------------------------------
    DSBeta = ComputeEmaDownsideBeta(wsChart, meanColChart, wsBench, halfLife)
    If DSBeta > 0 Then
        DDFactor = Application.WorksheetFunction.Min((1 / DSBeta) * 0.7, 0.95)
    Else
        DDFactor = 0.95
    End If

    ' -------------------------------------------------------------
    ' 2. Historical CVaR Logic (Abstracted)
    ' -------------------------------------------------------------
    CVaRFactor = ComputeCVaRFactor(wsChart, meanColChart, wsBench, benchCVaR, portCVaR, 0.95, 0.7)

    ' -------------------------------------------------------------
    ' 3. Semi-Deviation Logic (Abstracted)
    ' -------------------------------------------------------------
    SemiDevFactor = ComputeSemiDevFactor(wsChart, meanColChart, wsBench, benchSemiDev, portSemiDev, 0, 0.7)

    ' -------------------------------------------------------------
    ' 4. Historical Max Drawdown Logic (Abstracted)
    ' -------------------------------------------------------------
    HistoricalMDDFactor = ComputeHistoricalMDDFactor(wsChart, meanColChart, wsBench, benchMDD, portMDD, 0.7)

End Sub


Private Sub ComputeMDDFactor(ByRef DSBeta As Double, ByRef DDFactor As Double, Optional halfLife As Double = 256)
    Dim wb As Workbook
    Dim wsChart As Worksheet
    Dim wsBench As Worksheet

    Set wb = ThisWorkbook
    On Error Resume Next
    Set wsChart = wb.Sheets("ChartData")
    Set wsBench = wb.Sheets("BenchData")
    On Error GoTo 0

    If wsChart Is Nothing Or wsBench Is Nothing Then
        Exit Sub
    End If

    Dim meanColChart As Long
    meanColChart = GetColumnIndexByName(wsChart, "MEAN")

    If meanColChart = 0 Then
        Exit Sub
    End If

    DSBeta = ComputeEmaDownsideBeta(wsChart, meanColChart, wsBench, halfLife)
    DDFactor = Application.WorksheetFunction.Min((1 / DSBeta) * 0.7, 0.95)

End Sub

' Main Sub-Function: Computes the CVaR allocation factor
Private Function ComputeCVaRFactor(wsChart As Worksheet, meanColChart As Long, wsBench As Worksheet, ByRef benchCVaR As Double, ByRef portCVaR As Double, Optional confidence As Double = 0.95, Optional targetRatio As Double = 0.7) As Double
    Dim matchedChartPrices() As Double, matchedBenchPrices() As Double
    Dim numMatched As Long
    Dim tempFactor As Double

    ' Load and align the data
    numMatched = LoadAndAlignData(wsChart, meanColChart, wsBench, matchedChartPrices, matchedBenchPrices)

    ' Need sufficient data points to calculate a meaningful tail (e.g., at least 20 days)
    If numMatched >= 20 Then
        Dim chartRets() As Double, benchRets() As Double
        chartRets = CalculateReturns(matchedChartPrices, numMatched)
        benchRets = CalculateReturns(matchedBenchPrices, numMatched)

        ' Calculate CVaR for both portfolio and benchmark
        benchCVaR = CalculateHistoricalCVaR(benchRets, confidence)
        portCVaR = CalculateHistoricalCVaR(chartRets, confidence)

        ' Target Risk is 70% (targetRatio) of the Benchmark's tail risk
        Dim targetCVaR As Double
        targetCVaR = benchCVaR * targetRatio

        ' Calculate implied weight (Only scale if both have negative tails)
        If portCVaR < 0 And targetCVaR < 0 Then
            tempFactor = targetCVaR / portCVaR
        Else
            tempFactor = 1 ' No risk reduction needed if tail is positive
        End If

        ' Cap at 95% maximum exposure
        ComputeCVaRFactor = Application.WorksheetFunction.Min(tempFactor, 0.95)
    Else
        ' Fallback if not enough data
        benchCVaR = 0
        portCVaR = 0
        ComputeCVaRFactor = 0.95
    End If
End Function


' =========================================================================
' DD MATH HELPER FUNCTIONS
' =========================================================================

' Helper: Find Column Index by Name
Private Function GetColumnIndexByName(ws As Worksheet, colName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    GetColumnIndexByName = 0
    For c = 1 To lastCol
        If UCase(Trim(ws.Cells(1, c).Value)) = UCase(Trim(colName)) Then
            GetColumnIndexByName = c
            Exit Function
        End If
    Next c
End Function

' Orchestrator: EM Weighted Downside Beta matching dates
Private Function ComputeEmaDownsideBeta(wsChart As Worksheet, meanColChart As Long, wsBench As Worksheet, halfLife As Double) As Double
    Dim matchedChartPrices() As Double
    Dim matchedBenchPrices() As Double
    Dim numMatched As Long

    numMatched = LoadAndAlignData(wsChart, meanColChart, wsBench, matchedChartPrices, matchedBenchPrices)

    If numMatched < 2 Then
        ComputeEmaDownsideBeta = 1
        Exit Function
    End If

    Dim numRets As Long
    numRets = numMatched - 1

    Dim chartRets() As Double
    Dim benchRets() As Double
    chartRets = CalculateReturns(matchedChartPrices, numMatched)
    benchRets = CalculateReturns(matchedBenchPrices, numMatched)

    Dim weights() As Double
    Dim sumW As Double, sumW2 As Double
    Dim countDownside As Long

    countDownside = CalculateDownsideEmaWeights(benchRets, halfLife, numRets, weights, sumW, sumW2)

    If countDownside < 2 Or sumW = 0 Then
        ComputeEmaDownsideBeta = 1
        Exit Function
    End If

    Dim meanBr As Double, meanPr As Double
    meanBr = CalculateWeightedMean(benchRets, weights, numRets, sumW)
    meanPr = CalculateWeightedMean(chartRets, weights, numRets, sumW)

    Dim cov As Double, var As Double
    If CalculateWeightedCovarianceAndVariance(benchRets, chartRets, weights, numRets, meanBr, meanPr, sumW, sumW2, cov, var) Then
        ComputeEmaDownsideBeta = cov / var
    Else
        ComputeEmaDownsideBeta = 1
    End If
End Function

' Sub-Function: Load and Align Data
Private Function LoadAndAlignData(wsChart As Worksheet, meanColChart As Long, wsBench As Worksheet, ByRef matchedChartPrices() As Double, ByRef matchedBenchPrices() As Double) As Long
    Dim lastRowChart As Long, lastRowBench As Long
    lastRowChart = wsChart.Cells(wsChart.Rows.count, 1).End(xlUp).Row
    lastRowBench = wsBench.Cells(wsBench.Rows.count, 1).End(xlUp).Row

    If lastRowChart < 3 Or lastRowBench < 3 Then
        LoadAndAlignData = 0
        Exit Function
    End If

    Dim arrChartDates() As Variant, arrChartPrices() As Variant
    Dim arrBenchDates() As Variant, arrBenchPrices() As Variant

    arrChartDates = wsChart.Range(wsChart.Cells(1, 1), wsChart.Cells(lastRowChart, 1)).Value
    arrChartPrices = wsChart.Range(wsChart.Cells(1, meanColChart), wsChart.Cells(lastRowChart, meanColChart)).Value
    arrBenchDates = wsBench.Range(wsBench.Cells(1, 5), wsBench.Cells(lastRowBench, 5)).Value
    arrBenchPrices = wsBench.Range(wsBench.Cells(1, 8), wsBench.Cells(lastRowBench, 8)).Value

    Dim dictBench As Object
    Set dictBench = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 2 To lastRowBench
        If IsDate(arrBenchDates(r, 1)) Then
            dictBench(CDate(arrBenchDates(r, 1))) = SafeCDbl(arrBenchPrices(r, 1))
        End If
    Next r

    Dim countMatched As Long
    countMatched = 0
    ReDim matchedChartPrices(1 To lastRowChart)
    ReDim matchedBenchPrices(1 To lastRowChart)

    For r = 2 To lastRowChart
        If IsDate(arrChartDates(r, 1)) Then
            Dim d As Date
            d = CDate(arrChartDates(r, 1))
            If dictBench.Exists(d) Then
                countMatched = countMatched + 1
                matchedChartPrices(countMatched) = SafeCDbl(arrChartPrices(r, 1))
                matchedBenchPrices(countMatched) = dictBench(d)
            End If
        End If
    Next r

    If countMatched > 0 Then
        ReDim Preserve matchedChartPrices(1 To countMatched)
        ReDim Preserve matchedBenchPrices(1 To countMatched)
    End If

    LoadAndAlignData = countMatched
End Function

' Sub-Function: Calculate Returns
Private Function CalculateReturns(prices() As Double, count As Long) As Double()
    Dim rets() As Double
    Dim numRets As Long
    numRets = count - 1
    ReDim rets(1 To numRets)

    Dim i As Long
    For i = 1 To numRets
        If prices(i) <> 0 And prices(i + 1) <> 0 Then
            rets(i) = (prices(i + 1) / prices(i)) - 1
        Else
            rets(i) = 0
        End If
    Next i
    CalculateReturns = rets
End Function

' Sub-Function: Calculate EMA Weights for Downside
Private Function CalculateDownsideEmaWeights(benchRets() As Double, halfLife As Double, numRets As Long, ByRef outWeights() As Double, ByRef outSumW As Double, ByRef outSumW2 As Double) As Long
    Dim lambda As Double
    lambda = Exp(-Log(2) / halfLife)

    ReDim outWeights(1 To numRets)
    outSumW = 0
    outSumW2 = 0

    Dim countDownside As Long
    countDownside = 0
    Dim i As Long
    Dim w As Double

    For i = 1 To numRets
        If benchRets(i) < 0 Then
            countDownside = countDownside + 1
            w = lambda ^ (numRets - i)
            outWeights(i) = w
            outSumW = outSumW + w
            outSumW2 = outSumW2 + (w * w)
        Else
            outWeights(i) = 0
        End If
    Next i

    CalculateDownsideEmaWeights = countDownside
End Function

' Sub-Function: Calculate Weighted Mean
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

' Sub-Function: Calculate Weighted Covariance and Variance
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

' Safe conversion wrapper for Variants to Double
Private Function SafeCDbl(val As Variant) As Double
    If IsEmpty(val) Or IsError(val) Then
        SafeCDbl = 0
    ElseIf IsNumeric(val) Then
        SafeCDbl = CDbl(val)
    Else
        SafeCDbl = 0
    End If
End Function


' =========================================================================
' NEW CVaR MATH HELPER FUNCTIONS
' =========================================================================



' Abstracted Function: Computes the Semi-Deviation allocation factor
Private Function ComputeSemiDevFactor(wsChart As Worksheet, meanColChart As Long, wsBench As Worksheet, ByRef benchSemiDev As Double, ByRef portSemiDev As Double, Optional target As Double = 0, Optional targetRatio As Double = 0.7) As Double
    Dim matchedChartPrices() As Double, matchedBenchPrices() As Double
    Dim numMatched As Long
    Dim tempFactor As Double

    numMatched = LoadAndAlignData(wsChart, meanColChart, wsBench, matchedChartPrices, matchedBenchPrices)

    If numMatched >= 20 Then
        Dim chartRets() As Double, benchRets() As Double
        chartRets = CalculateReturns(matchedChartPrices, numMatched)
        benchRets = CalculateReturns(matchedBenchPrices, numMatched)

        benchSemiDev = CalculateSemiDeviation(benchRets, target)
        portSemiDev = CalculateSemiDeviation(chartRets, target)

        ' Worse volatility is higher (positive numbers)
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


' Abstracted Function: Computes the Absolute Maximum Drawdown allocation factor
Private Function ComputeHistoricalMDDFactor(wsChart As Worksheet, meanColChart As Long, wsBench As Worksheet, ByRef benchMDD As Double, ByRef portMDD As Double, Optional targetRatio As Double = 0.7) As Double
    Dim matchedChartPrices() As Double, matchedBenchPrices() As Double
    Dim numMatched As Long
    Dim tempFactor As Double

    numMatched = LoadAndAlignData(wsChart, meanColChart, wsBench, matchedChartPrices, matchedBenchPrices)

    If numMatched >= 20 Then
        Dim chartRets() As Double, benchRets() As Double
        chartRets = CalculateReturns(matchedChartPrices, numMatched)
        benchRets = CalculateReturns(matchedBenchPrices, numMatched)

        benchMDD = CalculateMaximumDrawdown(benchRets)
        portMDD = CalculateMaximumDrawdown(chartRets)

        ' Worse drawdown is more negative (smaller numbers mathematically)
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

' Sub-Function: Calculate Target Semi-Deviation (Returns < Target)
' Sub-Function: Calculate Absolute Maximum Drawdown (MDD)
Private Function CalculateMaximumDrawdown(rets() As Double) As Double
    Dim i As Long
    Dim peak As Double, current As Double, dd As Double, mdd As Double

    peak = 1
    current = 1
    mdd = 0

    ' Reconstruct a cumulative wealth index to find the peak-to-trough drop
    For i = LBound(rets) To UBound(rets)
        current = current * (1 + rets(i))

        If current > peak Then
            peak = current
        End If

        ' Current drawdown from the highest historical peak
        dd = (current / peak) - 1

        If dd < mdd Then
            mdd = dd
        End If
    Next i

    CalculateMaximumDrawdown = mdd
End Function

' Sub-Function: Calculate Historical CVaR
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

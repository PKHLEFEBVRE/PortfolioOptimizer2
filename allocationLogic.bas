Attribute VB_Name = "allocationLogic"
Option Explicit

' =========================================================================
' DOWNSIDE BETA & CVaR EQUITY ALLOCATION MODULE
' =========================================================================

Function ComputeFinalRiskFactor(portMean As SimulatedPortfolioCls, portBench As SimulatedPortfolioCls, targetRatio As Double) As Object
    Dim DSBeta As Double, DDFactor As Double
    Dim CVaRFactor As Double, benchCVaR As Double, portCVaR As Double
    Dim SemiDevFactor As Double, benchSemiDev As Double, portSemiDev As Double
    Dim HistoricalMDDFactor As Double, benchMDD As Double, portMDD As Double
    Dim finalRiskFactor As Double

    Dim meanEq() As Double
    meanEq = portMean.EquityCurve
    Dim benchEq() As Double
    benchEq = portBench.EquityCurve

    Dim nDays As Long
    nDays = UBound(meanEq, 1)

    Dim chartRets() As Double, benchRets() As Double
    ReDim chartRets(1 To nDays - 1)
    ReDim benchRets(1 To nDays - 1)

    Dim i As Long
    For i = 1 To nDays - 1
        If meanEq(i, 1) > 0 And meanEq(i + 1, 1) > 0 Then
            chartRets(i) = Log(meanEq(i + 1, 1) / meanEq(i, 1))
        Else
            chartRets(i) = 0
        End If

        If benchEq(i, 1) > 0 And benchEq(i + 1, 1) > 0 Then
            benchRets(i) = Log(benchEq(i + 1, 1) / benchEq(i, 1))
        Else
            benchRets(i) = 0
        End If
    Next i

    ' 1. Downside Beta
    DSBeta = ComputeEmaDownsideBeta(benchRets, chartRets, 256)
    If DSBeta > 0 Then
        DDFactor = Application.WorksheetFunction.Min((1 / DSBeta) * targetRatio, 0.95)
    Else
        DDFactor = 0.95
    End If

    ' 2. CVaR Factor
    benchCVaR = CalculateHistoricalCVaR(benchRets, 0.95)
    portCVaR = CalculateHistoricalCVaR(chartRets, 0.95)
    Dim targetCVaR As Double
    targetCVaR = benchCVaR * targetRatio
    If portCVaR < 0 And targetCVaR < 0 Then
        CVaRFactor = Application.WorksheetFunction.Min(targetCVaR / portCVaR, 0.95)
    Else
        CVaRFactor = 0.95
    End If

    ' 3. Semi-Dev Factor
    benchSemiDev = CalculateSemiDeviation(benchRets, 0)
    portSemiDev = CalculateSemiDeviation(chartRets, 0)
    If portSemiDev > 0 And portSemiDev > (benchSemiDev * targetRatio) Then
        SemiDevFactor = Application.WorksheetFunction.Min((benchSemiDev * targetRatio) / portSemiDev, 0.95)
    Else
        SemiDevFactor = 0.95
    End If

    ' 4. Historical MDD Factor
    benchMDD = CalculateMaximumDrawdown(benchRets)
    portMDD = CalculateMaximumDrawdown(chartRets)
    If portMDD < 0 And portMDD < (benchMDD * targetRatio) Then
        HistoricalMDDFactor = Application.WorksheetFunction.Min((benchMDD * targetRatio) / portMDD, 0.95)
    Else
        HistoricalMDDFactor = 0.95
    End If

    finalRiskFactor = Application.WorksheetFunction.Average(DDFactor, CVaRFactor, SemiDevFactor, HistoricalMDDFactor)

    Dim resDict As Object
    Set resDict = CreateObject("Scripting.Dictionary")
    resDict.Add "DDFactor", DDFactor
    resDict.Add "CVaRFactor", CVaRFactor
    resDict.Add "SemiDevFactor", SemiDevFactor
    resDict.Add "HistoricalMDDFactor", HistoricalMDDFactor
    resDict.Add "FinalRiskFactor", finalRiskFactor

    Set ComputeFinalRiskFactor = resDict
End Function

' =========================================================================
' DOWNSIDE MATH
' =========================================================================

Private Function ComputeEmaDownsideBeta(benchRets() As Double, chartRets() As Double, halfLife As Double) As Double
    Dim numRets As Long
    numRets = UBound(benchRets)
    If numRets < 2 Then
        ComputeEmaDownsideBeta = 1
        Exit Function
    End If

    Dim lambda As Double
    lambda = Exp(-Log(2) / halfLife)

    Dim weights() As Double
    ReDim weights(1 To numRets)
    Dim sumW As Double, sumW2 As Double
    sumW = 0: sumW2 = 0

    Dim countDownside As Long
    countDownside = 0
    Dim i As Long, w As Double

    For i = 1 To numRets
        If benchRets(i) < 0 Then
            countDownside = countDownside + 1
            w = lambda ^ (numRets - i)
            weights(i) = w
            sumW = sumW + w
            sumW2 = sumW2 + (w * w)
        Else
            weights(i) = 0
        End If
    Next i

    If countDownside < 2 Or sumW = 0 Then
        ComputeEmaDownsideBeta = 1
        Exit Function
    End If

    Dim meanBr As Double, meanPr As Double
    Dim sumW_RetB As Double, sumW_RetP As Double
    sumW_RetB = 0: sumW_RetP = 0

    For i = 1 To numRets
        If weights(i) > 0 Then
            sumW_RetB = sumW_RetB + (weights(i) * benchRets(i))
            sumW_RetP = sumW_RetP + (weights(i) * chartRets(i))
        End If
    Next i
    meanBr = sumW_RetB / sumW
    meanPr = sumW_RetP / sumW

    Dim sumWCov As Double, sumWVar As Double
    sumWCov = 0: sumWVar = 0
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
        ComputeEmaDownsideBeta = 1
    Else
        ComputeEmaDownsideBeta = sumWCov / sumWVar
    End If
End Function

Private Function CalculateHistoricalCVaR(rets() As Double, confidence As Double) As Double
    Dim numRets As Long
    numRets = UBound(rets)

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

    numRets = UBound(rets)
    sumSq = 0

    For i = 1 To numRets
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

    For i = 1 To UBound(rets)
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

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetricsCalculatorCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_mRet As Double
Private p_mVol As Double
Private p_mSharpe As Double
Private p_mMDD As Double
Private p_mLen As Integer
Private p_mVaR As Double
Private p_mCVaR As Double
Private p_mSemiDev As Double
Private p_mDownsideBeta As Double

Public Property Get Ret() As Double: Ret = p_mRet: End Property
Public Property Get Vol() As Double: Vol = p_mVol: End Property
Public Property Get Sharpe() As Double: Sharpe = p_mSharpe: End Property
Public Property Get MDD() As Double: MDD = p_mMDD: End Property
Public Property Get MDDLen() As Integer: MDDLen = p_mLen: End Property
Public Property Get VaR() As Double: VaR = p_mVaR: End Property
Public Property Get CVaR() As Double: CVaR = p_mCVaR: End Property
Public Property Get SemiDev() As Double: SemiDev = p_mSemiDev: End Property
Public Property Get DownsideBeta() As Double: DownsideBeta = p_mDownsideBeta: End Property

' We expect 1D arrays: prices(1 to n), dates(1 to n), logRets(1 to n-1)
Public Sub Calculate(prices() As Double, dates() As Date, logRets() As Double, _
                     useEWMA As Boolean, riskFreeRate As Double, varConfidence As Double, _
                     Optional benchRets As Variant, Optional calcType As String = "Asset")

    Dim nDays As Long
    nDays = UBound(prices) - LBound(prices) + 1

    If nDays < 2 Then Exit Sub

    Dim totalRet As Double
    totalRet = prices(UBound(prices)) / prices(LBound(prices))

    ' Annualized Return
    If calcType = "Asset" Then
        p_mRet = totalRet ^ (365 / (dates(UBound(dates)) - dates(LBound(dates)))) - 1
    Else
        p_mRet = (totalRet ^ (256 / (nDays - 1))) - 1
    End If

    ' Volatility
    Dim s As Double, meanL As Double
    Dim i As Long
    s = 0
    For i = LBound(logRets) To UBound(logRets)
        s = s + logRets(i)
    Next i
    meanL = s / (UBound(logRets) - LBound(logRets) + 1)

    If useEWMA Then
        Dim lambda1M As Double
        Dim varEWMA1M As Double
        Dim ewmaVols() As Double
        ReDim ewmaVols(1 To UBound(logRets))
        lambda1M = 0.5 ^ (1 / 21)

        For i = LBound(logRets) To UBound(logRets)
            Dim sqDev As Double
            sqDev = (logRets(i) - meanL) ^ 2

            If i = LBound(logRets) Then
                varEWMA1M = sqDev
            Else
                varEWMA1M = lambda1M * varEWMA1M + (1 - lambda1M) * sqDev
            End If

            ewmaVols(i - LBound(logRets) + 1) = Sqr(varEWMA1M) * Sqr(256)
        Next i

        Dim currentVol1M As Double
        Dim perc90Vol1M As Double
        currentVol1M = ewmaVols(UBound(ewmaVols))
        perc90Vol1M = Application.WorksheetFunction.Percentile_Inc(ewmaVols, 0.9)
        p_mVol = (currentVol1M + perc90Vol1M) / 2
    Else
        Dim sumSq As Double
        sumSq = 0
        For i = LBound(logRets) To UBound(logRets)
            sumSq = sumSq + (logRets(i) - meanL) ^ 2
        Next i
        ' (nDays - 2) logic was in portfolioUtils. We use UBound(logRets) - 1.
        ' To avoid division by zero:
        Dim denom As Long
        denom = UBound(logRets) - LBound(logRets)
        If denom < 1 Then denom = 1
        p_mVol = Sqr(sumSq / denom) * Sqr(256)
    End If

    ' Sharpe Ratio
    If p_mVol > 0 Then
        p_mSharpe = (p_mRet - riskFreeRate) / p_mVol
    Else
        p_mSharpe = 0
    End If

    ' Drawdown and Length
    Dim peak As Double
    Dim dd As Double
    Dim peakDay As Long, currLen As Integer, maxLen As Integer
    peak = -99999
    p_mMDD = 0
    maxLen = 0: currLen = 0

    For i = LBound(prices) To UBound(prices)
        If prices(i) > peak Then
            peak = prices(i)
            peakDay = i
            currLen = 0
        Else
            dd = (prices(i) / peak) - 1
            currLen = i - peakDay
            If dd < p_mMDD Then p_mMDD = dd
            If currLen > maxLen Then maxLen = currLen
        End If
    Next i
    p_mLen = maxLen

    ' Analytical VaR
    Dim Z As Double
    On Error Resume Next
    Z = Application.WorksheetFunction.NormSInv(varConfidence)
    On Error GoTo 0
    If Z = 0 Then Z = 1.645 ' default for 95% if error
    p_mVaR = p_mRet - Z * p_mVol

    ' Historical CVaR (95%)
    p_mCVaR = CalculateHistoricalCVaR(logRets, varConfidence)

    ' Semi-Deviation (target = 0)
    p_mSemiDev = CalculateSemiDeviation(logRets, 0)

    ' Downside Beta
    If Not IsMissing(benchRets) Then
        If IsArray(benchRets) Then
            p_mDownsideBeta = ComputeDownsideBeta(benchRets, logRets, 256)
        End If
    End If
End Sub

Private Function CalculateHistoricalCVaR(rets() As Double, confidence As Double) As Double
    Dim numRets As Long
    numRets = UBound(rets) - LBound(rets) + 1

    Dim sortedRets() As Double
    ReDim sortedRets(1 To numRets)
    Dim i As Long
    For i = 1 To numRets
        sortedRets(i) = rets(i - 1 + LBound(rets))
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

Private Function ComputeDownsideBeta(benchRets() As Double, portRets() As Double, halfLife As Double) As Double
    Dim lambda As Double
    lambda = Exp(-Log(2) / halfLife)

    Dim numRets As Long
    numRets = UBound(benchRets) - LBound(benchRets) + 1

    Dim weights() As Double
    ReDim weights(1 To numRets)
    Dim sumW As Double, sumW2 As Double
    sumW = 0: sumW2 = 0

    Dim countDownside As Long
    countDownside = 0
    Dim i As Long, w As Double

    For i = 1 To numRets
        If benchRets(i - 1 + LBound(benchRets)) < 0 Then
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
        ComputeDownsideBeta = 1
        Exit Function
    End If

    Dim meanBr As Double, meanPr As Double
    Dim sumW_RetB As Double, sumW_RetP As Double
    sumW_RetB = 0: sumW_RetP = 0
    For i = 1 To numRets
        If weights(i) > 0 Then
            sumW_RetB = sumW_RetB + (weights(i) * benchRets(i - 1 + LBound(benchRets)))
            sumW_RetP = sumW_RetP + (weights(i) * portRets(i - 1 + LBound(portRets)))
        End If
    Next i
    meanBr = sumW_RetB / sumW
    meanPr = sumW_RetP / sumW

    Dim sumWCov As Double, sumWVar As Double
    sumWCov = 0: sumWVar = 0
    Dim devBr As Double, devPr As Double

    For i = 1 To numRets
        If weights(i) > 0 Then
            devBr = benchRets(i - 1 + LBound(benchRets)) - meanBr
            devPr = portRets(i - 1 + LBound(portRets)) - meanPr

            sumWCov = sumWCov + (weights(i) * devBr * devPr)
            sumWVar = sumWVar + (weights(i) * devBr * devBr)
        End If
    Next i

    Dim denom As Double
    denom = sumW - (sumW2 / sumW)

    If denom <= 0 Or sumWVar = 0 Then
        ComputeDownsideBeta = 1
    Else
        ComputeDownsideBeta = sumWCov / sumWVar
    End If
End Function

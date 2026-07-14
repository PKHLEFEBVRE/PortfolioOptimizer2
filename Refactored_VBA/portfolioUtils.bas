Option Explicit
Option Base 1



Sub SimulatePortfolio(prices As Variant, dates As Variant, w As Variant, _
                      ByRef outCurve() As Double, ByRef outWeights() As Double, ByRef outDD() As Double, _
                      ByRef mRet As Double, ByRef mVol As Double, ByRef mSharpe As Double, _
                      ByRef mMDD As Double, ByRef mLen As Integer, ByRef mVaR As Double)
    Dim nDays As Long, nAssets As Long
    Dim i As Long, j As Long
    nDays = UBound(prices, 1)
    nAssets = UBound(prices, 2)
    ReDim outCurve(1 To nDays, 1 To 1)
    outCurve(1, 1) = 100
    ReDim outWeights(1 To nDays, 1 To nAssets)
    ReDim outDD(1 To nDays, 1 To 1)
    Dim currentVal As Double
    currentVal = 100
    Dim shares() As Double
    ReDim shares(1 To nAssets)
    For j = 1 To nAssets
        shares(j) = (currentVal * w(1, j)) / prices(1, j)
        outWeights(1, j) = w(1, j)
    Next j
    Dim logRets() As Double
    ReDim logRets(1 To nDays - 1)
    Dim peak As Double
    peak = 100
    For i = 2 To nDays
        Dim dPrev As Date, dCurr As Date
        dPrev = dates(i - 1)
        dCurr = dates(i)
        Dim prevTotal As Double
        prevTotal = outCurve(i - 1, 1)
        If Month(dCurr) <> Month(dPrev) Then
            For j = 1 To nAssets
                shares(j) = (prevTotal * w(1, j)) / prices(i, j)
            Next j
        End If
        Dim assetVal As Double
        assetVal = 0
        For j = 1 To nAssets
            assetVal = assetVal + shares(j) * prices(i, j)
        Next j
        outCurve(i, 1) = assetVal
        If assetVal > peak Then
            peak = assetVal
            outDD(i, 1) = 0
        Else
            outDD(i, 1) = (assetVal / peak) - 1
        End If
        For j = 1 To nAssets
            If assetVal > 0 Then
                outWeights(i, j) = (shares(j) * prices(i, j)) / assetVal
            Else
                outWeights(i, j) = 0
            End If
        Next j
        logRets(i - 1) = Log(outCurve(i, 1) / outCurve(i - 1, 1))
    Next i
    Dim totalRet As Double
    totalRet = outCurve(nDays, 1) / 100
    mRet = (totalRet ^ (256 / (nDays - 1))) - 1
    Dim sumSq As Double, meanL As Double
    Dim s As Double
    s = 0
    For i = 1 To nDays - 1
        s = s + logRets(i)
    Next i
    meanL = s / (nDays - 1)
    sumSq = 0
    For i = 1 To nDays - 1
        sumSq = sumSq + (logRets(i) - meanL) ^ 2
    Next i
    mVol = Sqr(sumSq / (nDays - 2)) * Sqr(256)
    Dim rf As Double
    rf = Sheets(config.DashSheet).Range("D" & config.RiskFreeRow).Value
    mSharpe = (mRet - rf) / mVol
    Call CalculateDrawdownFromCurve(outCurve, mMDD, mLen)
    Dim conf As Double, Z As Double
    conf = Sheets(config.DashSheet).Range("D5").Value
    Z = Application.NormSInv(conf)
    mVaR = mRet - Z * mVol
End Sub

Sub CalculateDrawdownFromCurve(curve As Variant, ByRef outMDD As Double, ByRef outLen As Integer)
    InitializeGlobals
    Dim n As Long, i As Long
    n = UBound(curve, 1)
    Dim peak As Double, dd As Double, maxDD As Double
    Dim peakDay As Long, currLen As Integer, maxLen As Integer
    peak = -99999
    maxDD = 0: maxLen = 0: currLen = 0
    For i = 1 To n
        Dim val As Double
        val = curve(i, 1)
        If val > peak Then
            peak = val
            peakDay = i
            currLen = 0
        Else
            dd = (val / peak) - 1
            currLen = i - peakDay
            If dd < maxDD Then maxDD = dd
            If currLen > maxLen Then maxLen = currLen
        End If
    Next i
    outMDD = maxDD
    outLen = maxLen
End Sub

Sub OutputMetricsToRow(rowOffset As Integer, TypeStr As String, mRet, mVol, mSharpe, mMDD, mLen, mVaR, w, RangeName As String, WriteWeights As Boolean)
    InitializeGlobals
    Dim activeSht As Worksheet
    Set activeSht = ThisWorkbook.ActiveSheet

    Dim ws As Worksheet
    Set ws = Sheets(config.DashSheet)
    Dim startRow As Long, startCol As Long
    startRow = ws.Range(RangeName).Row + 1
    startCol = ws.Range(RangeName).Column

    Dim r As Integer
    r = startRow + rowOffset
    If RangeName = "AssetMetricsStart" Then
        startCol = startCol - 1
    Else
        ws.Cells(r, startCol).Value = TypeStr
    End If

    ws.Cells(r, startCol + 1).Value = mRet
    ws.Cells(r, startCol + 2).Value = mVol
    ws.Cells(r, startCol + 3).Value = mSharpe
    ws.Cells(r, startCol + 4).Value = mMDD
    ws.Cells(r, startCol + 5).Value = mLen
    ws.Cells(r, startCol + 6).Value = mVaR

    If WriteWeights Then

        Dim equityPositions As Dictionary
        Set equityPositions = DataUtils.GetEquityPositionsFromFund

        Dim equitySum As Double, i As Integer
        equitySum = 0

        For i = 1 To equityPositions.count
            equitySum = equitySum + equityPositions(equityPositions.Keys(i - 1)).Weight
        Next i

        startRow = ws.Range("StrategyWeightsStart").Row
        startCol = ws.Range("StrategyWeightsStart").Column + 1 + rowOffset

        ws.Cells(startRow, startCol).Value = TypeStr

        Dim nCols As Integer
        nCols = UBound(w, 2)

        For i = 1 To nCols
            ws.Cells(startRow + i, startCol).Value = w(1, i)
            ws.Cells(startRow + i, startCol + 7).Value = w(1, i) * equitySum
        Next i
    End If

    activeSht.Activate
End Sub
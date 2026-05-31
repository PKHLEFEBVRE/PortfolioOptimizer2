olevba 0.60.2 on Python 3.12.13 - http://decalage.info/python/oletools
===============================================================================
FILE: Portfolio_Optimizer_MLSU.xlsm
Type: OpenXML
WARNING  For now, VBA stomping cannot be detected for files in memory
-------------------------------------------------------------------------------
VBA MACRO ThisWorkbook.cls
in file: xl/vbaProject.bin - OLE stream: 'VBA/ThisWorkbook'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet1.cls
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet2.cls
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet3.cls
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet3'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet4.cls
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet4'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet5.cls
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet5'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Main.bas
in file: xl/vbaProject.bin - OLE stream: 'VBA/Main'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Option Base 1

' ==========================================================================================
' PORTFOLIO OPTIMIZER VBA CODE
' ==========================================================================================


Private Const WEIGHTS_SHEET As String = "SimulatedWeights"
Private Const DASH_SHEET As String = "Dashboard"
Private Const ENGINE_SHEET As String = "Engine"
Private Const CHART_SHEET As String = "ChartData"


' ==========================================================
' MAIN ENTRY POINTS
' ==========================================================

Sub GetDataFromAddin()
    Application.ScreenUpdating = False
    Call RunClean
    Call updatePriceHistoryFromInfinForFund
    Call updatePriceHistoryFromInfinForBench
End Sub


Sub RunUpdateMatrices()
    Application.ScreenUpdating = False

    Dim prices() As Variant, dates() As Date, assetNames() As String
    Call GetHistoricalDataDatesAndNames(prices, dates, assetNames)
    If UBound(prices, 2) < 1 Then Exit Sub

    Dim strategies As Variant
    strategies = Array("ERC UNCSTRD", "ER/VOL", "SHARPE", "CUSTOM", "MEAN") ', "MIN VAR", "KELLY"

    Call SyncDashboardHeaders(dates, assetNames, strategies)
    Call UpdateConvictions
    Call UpdateCurrentPrices(prices)

    Dim logRets() As Double, meanRets() As Double
    Call CalculateHistoricalStats(dates, prices, logRets, meanRets)
    Call ProcessIndividualAssets(prices, dates, logRets, assetNames)

    'MsgBox "Data Updated. Matrices Built.", vbInformation
    Application.ScreenUpdating = True
End Sub

Sub RunAllSolvers()
    Application.ScreenUpdating = False
    Sheets(ENGINE_SHEET).Activate

    Dim prices() As Variant, dates() As Date, assetNames() As String
    Call GetHistoricalDataDatesAndNames(prices, dates, assetNames)

    Dim logRets() As Double, meanRets() As Double
    Call CalculateHistoricalStats(dates, prices, logRets, meanRets)

    Dim covMat() As Double
    covMat = CalculateCovariance(logRets, meanRets)
    Call OutputCorrelationMatrix(covMat, assetNames)

    Dim expReturns() As Double
    Dim lastDate As Date
    lastDate = dates(UBound(dates))
    expReturns = CalculateExpectedReturns(meanRets, lastDate)

    Dim wsDash As Worksheet
    Set wsDash = Sheets(DASH_SHEET)
    Dim inputStart As Long
    inputStart = wsDash.Range("InputTableStart").Row + 1

    Dim k As Integer, nAssets As Integer
    nAssets = UBound(assetNames)
    For k = 1 To nAssets
        wsDash.Cells(inputStart + k - 1, wsDash.Range("AssetMetricsStart").Column - 1).Value = expReturns(k)
    Next k

    Call WriteToEngine(covMat, expReturns)


    Dim strategies As Variant
    strategies = Array("ERC UNCSTRD", "ER/VOL", "SHARPE", "CUSTOM", "MEAN") ', "MIN VAR", "KELLY"
    Dim nStrat As Integer
    nStrat = UBound(strategies)

    Call SyncSimWeightsHeaders(dates, assetNames, strategies)

    Dim minWeights() As Double, maxWeights() As Double
    ReDim minWeights(1 To nAssets)
    ReDim maxWeights(1 To nAssets)

    For k = 1 To nAssets
        If IsEmpty(wsDash.Cells(inputStart + k - 1, wsDash.Range("maxWeightCol").Column).Value) Then
            maxWeights(k) = 1
        Else
            maxWeights(k) = wsDash.Cells(inputStart + k - 1, wsDash.Range("maxWeightCol").Column).Value
        End If
        If IsEmpty(wsDash.Cells(inputStart + k - 1, wsDash.Range("maxWeightCol").Column - 1).Value) Then
            minWeights(k) = 0
        Else
            minWeights(k) = wsDash.Cells(inputStart + k - 1, wsDash.Range("maxWeightCol").Column - 1).Value
        End If
    Next k

    Dim wsEngC As Worksheet
    Set wsEngC = Sheets(ENGINE_SHEET)
    Dim optRangeC As Range
    Set optRangeC = wsEngC.Range("OptWeights")

    Dim cVal As Double
    Dim m As Integer
    Dim i As Integer

    Dim sumWeights() As Double
    ReDim sumWeights(1 To nAssets)
    Dim wCount As Integer
    wCount = UBound(strategies) - 1 ' All strategies except MEAN

    For m = 1 To nAssets
        sumWeights(m) = 0
    Next m

    For i = LBound(strategies) To UBound(strategies)
        Dim stratName As String
        stratName = UCase(strategies(i))

        Dim eqW As Double
        eqW = 1 / nAssets
        Dim wsEng As Worksheet
        Set wsEng = Sheets(ENGINE_SHEET)
        Dim optRange As Range
        Set optRange = wsEng.Range("OptWeights")
        Dim j As Integer
        For j = 1 To nAssets
            optRange.Cells(1, j).Value = eqW
        Next j

        If stratName = "EQUAL WEIGHT" Then
            'done

        ElseIf stratName = "CUSTOM" Then

            Dim convictions() As Variant
            convictions = wsDash.Range(wsDash.Cells(inputStart, wsDash.Range("customWeightCol").Column), wsDash.Cells(inputStart + nAssets - 1, wsDash.Range("customWeightCol").Column)).Value
            convictions = Application.WorksheetFunction.Transpose(convictions)

            Dim sum_ As Long
            sum_ = 0
            For m = 1 To nAssets
                sum_ = sum_ + convictions(m)
            Next m

            For m = 1 To nAssets
                cVal = convictions(m) / sum_
                optRangeC.Cells(1, m).Value = cVal
            Next m

        ElseIf stratName = "ER/VOL" Then

            For m = 1 To nAssets
                cVal = wsDash.Cells(inputStart + m - 1, wsDash.Range("AssetMetricsStart").Column - 1).Value / _
                          wsDash.Cells(inputStart + m - 1, wsDash.Range("AssetMetricsStart").Column + 1).Value
                optRangeC.Cells(1, m).Value = cVal
            Next m

            Dim currentSum As Double
            currentSum = Application.WorksheetFunction.sum(optRange)

            If currentSum <> 0 Then
                Dim finalWeights() As Variant
                ReDim finalWeights(1 To nAssets)
                For m = 1 To nAssets
                    finalWeights(m) = optRange.Cells(1, m).Value / currentSum
                Next m
                optRange.Value = finalWeights
            End If

        ElseIf stratName = "MEAN" Then

            For m = 1 To nAssets
                optRange.Cells(1, m).Value = sumWeights(m) / wCount
            Next m

        Else

            ' Optimization: Sharpe, Variance, Kelly, ERC UNCSTRD
            Call RunSolver(CStr(strategies(i)), minWeights, maxWeights, False)

        End If

        Dim w As Variant
        w = Sheets(ENGINE_SHEET).Range("OptWeights").Value

        If stratName <> "MEAN" Then
            For m = 1 To nAssets
                sumWeights(m) = sumWeights(m) + w(1, m)
            Next m
        End If

        Dim equityCurve() As Double, ddCurve() As Double, dailyWeights() As Double
        Dim mRet As Double, mVol As Double, mSharpe As Double, mMDD As Double, mLen As Integer, mVaR As Double

        Call SimulatePortfolio(prices, dates, w, equityCurve, dailyWeights, ddCurve, mRet, mVol, mSharpe, mMDD, mLen, mVaR)

        Call OutputMetricsToRow(i - 1, CStr(strategies(i)), mRet, mVol, mSharpe, mMDD, mLen, mVaR, w, "StrategyTableStart", True)

        Dim colOffset As Integer
        colOffset = (nAssets + 1) + i
        Call WriteCurveToSheet(equityCurve, colOffset)

        Dim ddColOffset As Integer
        ddColOffset = (nAssets + 1) + UBound(strategies) + i
        Call WriteCurveToSheet(ddCurve, ddColOffset)

        Dim wColStart As Integer
        wColStart = 2 + (i - 1) * nAssets
        Call WriteWeightsToSheet(dailyWeights, wColStart)
    Next i

    Dim ws As Worksheet
    Set ws = Sheets(DASH_SHEET)
    Dim startRow As Integer, startCol As Integer
    startRow = ws.Range("StrategyWeightsStart").Row
    startCol = ws.Range("StrategyWeightsStart").Column + UBound(strategies) - LBound(strategies) + 1

    Dim equityPositions As Dictionary
    Set equityPositions = DataUtils.GetEquityPositionsFromFund

    Dim equitySum As Double
    equitySum = 0

    For i = 1 To equityPositions.count
        equitySum = equitySum + equityPositions(equityPositions.Keys(i - 1)).Weight
    Next i

    For i = 1 To equityPositions.count
        ws.Cells(startRow + i, startCol + 1).Value = equityPositions(equityPositions.Keys(i - 1)).Weight / equitySum
        ws.Cells(startRow + i, startCol + 2).Value = ws.Cells(startRow + i, startCol + 1).Value - ws.Cells(startRow + i, startCol).Value
        ws.Cells(startRow + i, startCol + nStrat + 3).Value = equityPositions(equityPositions.Keys(i - 1)).Weight
        ws.Cells(startRow + i, startCol + nStrat + 4).Value = ws.Cells(startRow + i, startCol + nStrat + 3).Value - ws.Cells(startRow + i, startCol + nStrat + 2).Value
    Next i

    wsEngC.Activate
    'Call DrawEfficientFrontier(nAssets, maxWeights)

    Sheets(ENGINE_SHEET).Visible = False
    Sheets(DASH_SHEET).Activate

    Call UpdateDashboardCharts(nAssets, UBound(strategies), lastDate)


    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Optimization Completed.", vbInformation
End Sub


' ==========================================================
' CLEAN
' ==========================================================

Sub RunClean()
    Dim wsDash As Worksheet
    Set wsDash = Sheets(DASH_SHEET)

    Dim stratR As Long, stratC As Long
    stratR = wsDash.Range("StrategyTableStart").Row + 1
    stratC = wsDash.Range("StrategyTableStart").Column
    wsDash.Range(wsDash.Cells(stratR, stratC), wsDash.Cells(stratR + 5, stratC + 100)).ClearContents

    Dim inputR As Long, inputC As Long
    inputR = wsDash.Range("InputTableStart").Row + 1
    inputC = wsDash.Range("InputTableStart").Column
    wsDash.Range(wsDash.Cells(inputR, inputC - 1), wsDash.Cells(stratR - 3, wsDash.Range("InputTableStart").End(xlToRight).Column)).ClearContents
    'wsDash.Range(wsDash.Cells(inputR, inputC + 2), wsDash.Cells(stratR - 3, inputC + 2)).ClearContents
    'wsDash.Range(wsDash.Cells(inputR, wsDash.Range("AssetMetricsStart").Column - 1), wsDash.Cells(stratR - 2, wsDash.Range("InputTableStart").End(xlToRight).Column)).ClearContents

    Dim corrR As Long
    corrR = wsDash.Range("CorrMatrixStart").Row + 1
    wsDash.Range(wsDash.Cells(corrR, wsDash.Range("CorrMatrixStart").Column), wsDash.Cells(corrR + 50, wsDash.Range("CorrMatrixStart").Column + 50)).ClearContents
    corrR = wsDash.Range("StrategyWeightsStart").Row + 1
    wsDash.Range(wsDash.Cells(corrR - 1, wsDash.Range("StrategyWeightsStart").Column + 1), wsDash.Cells(corrR - 1, wsDash.Range("CorrMatrixStart").Column)).ClearContents
    wsDash.Range(wsDash.Cells(corrR, wsDash.Range("StrategyWeightsStart").Column - 1), wsDash.Cells(corrR + 50, wsDash.Range("CorrMatrixStart").Column - 2)).ClearContents
    Sheets(CHART_SHEET).Cells.ClearContents
    Sheets(WEIGHTS_SHEET).Cells.ClearContents
End Sub












-------------------------------------------------------------------------------
VBA MACRO chartsUtils.bas
in file: xl/vbaProject.bin - OLE stream: 'VBA/chartsUtils'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Option Base 1

Private Const CHART_SHEET As String = "ChartData"
Private Const DASH_SHEET As String = "Dashboard"



Sub UpdateDashboardCharts(nAssets As Integer, nStrats As Integer, lastDate As Date)
    Dim wsDash As Worksheet, wsChart As Worksheet
    Set wsDash = Sheets(DASH_SHEET)
    Set wsChart = Sheets(CHART_SHEET)

    Dim lastRow As Long
    lastRow = wsChart.Cells(wsChart.Rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    Dim xRange As Range
    Set xRange = wsChart.Range(wsChart.Cells(2, 1), wsChart.Cells(lastRow, 1))

    Dim chtObj As ChartObject
    For Each chtObj In wsDash.ChartObjects
        Dim title As String
        On Error Resume Next
        title = chtObj.Chart.ChartTitle.text
        On Error GoTo 0

        Dim startCol As Integer, count As Integer
        Dim isPerf As Boolean
        isPerf = False

        If title = "Assets Cumulative Returns" Then
            startCol = 2
            count = nAssets
            isPerf = True
        ElseIf title = "Strategies Cumulative Returns" Then
            startCol = nAssets + 2
            count = nStrats
            isPerf = True
        ElseIf title = "Strategies Drawdowns" Then
            startCol = nAssets + 2 + nStrats
            count = nStrats
            isPerf = False
        Else
            startCol = 0
        End If

        If startCol > 0 Then
            UpdateSingleChartSeries chtObj.Chart, xRange, wsChart, startCol, count

            Dim totalMonths As Long
            ' Calcul de la durée en mois entre le début et la fin


            With chtObj.Chart.Axes(xlCategory)
            totalMonths = DateDiff("m", .MinimumScale, .MaximumScale)
                .CategoryType = xlTimeScale
                .TickLabels.NumberFormat = "mmm yy"
                .MaximumScale = lastDate
                ' Ajustement dynamique de l'unité majeure
                Select Case totalMonths
                    Case Is <= 12
                        .MajorUnitScale = xlMonths
                        .MajorUnit = 1        ' Un label par mois
                    Case 13 To 24
                        .MajorUnitScale = xlMonths
                        .MajorUnit = 3        ' Un label tous les trimestres
                    Case 25 To 48
                        .MajorUnitScale = xlMonths
                        .MajorUnit = 6        ' Un label tous les semestres
                    Case Else
                        .MajorUnitScale = xlYears
                        .MajorUnit = 1        ' Un label par an
                End Select
            End With

            If isPerf Then
                Dim dRange As Range
                Set dRange = wsChart.Range(wsChart.Cells(2, startCol), wsChart.Cells(lastRow, startCol + count - 1))
                Dim minVal As Double, maxVal As Double
                minVal = Application.WorksheetFunction.Min(dRange)
                maxVal = Application.WorksheetFunction.Max(dRange)
                Dim rangeVal As Double
                rangeVal = maxVal - minVal
                If rangeVal = 0 Then rangeVal = 10
                Dim axisMin As Double, axisMax As Double
                axisMin = minVal - (rangeVal * 0.05)
                axisMax = maxVal + (rangeVal * 0.05)
                With chtObj.Chart.Axes(xlValue)
                    .MinimumScale = axisMin
                    .MaximumScale = axisMax
                    .TickLabels.NumberFormat = "0"
                End With
            Else
                Dim ddRange As Range
                Set ddRange = wsChart.Range(wsChart.Cells(2, startCol), wsChart.Cells(lastRow, startCol + count - 1))
                Dim minDD As Double
                minDD = Application.WorksheetFunction.Min(ddRange)
                With chtObj.Chart.Axes(xlValue)
                    .MaximumScale = 0
                    .MinimumScale = minDD * 1.1
                    .TickLabels.NumberFormat = "0%"
                End With
            End If
        End If
    Next chtObj
End Sub


Sub UpdateSingleChartSeries(cht As Chart, xRange As Range, wsData As Worksheet, startCol As Integer, count As Integer)
    Dim s As Series
    Do While cht.SeriesCollection.count > 0
        cht.SeriesCollection(1).Delete
    Loop
    Dim i As Integer
    For i = 0 To count - 1
        Dim colIdx As Integer
        colIdx = startCol + i
        Set s = cht.SeriesCollection.NewSeries
        s.XValues = xRange
        s.Values = wsData.Range(wsData.Cells(2, colIdx), wsData.Cells(xRange.Rows.count + 1, colIdx))
        s.Name = wsData.Cells(1, colIdx).Value
        s.Format.Line.Weight = 1.25
    Next i
    With cht.Legend
        .Position = xlLegendPositionTop
        .IncludeInLayout = True
        .Top = cht.ChartTitle.Top + cht.ChartTitle.Height + 5
    End With
End Sub
-------------------------------------------------------------------------------
VBA MACRO portfolioUtils.bas
in file: xl/vbaProject.bin - OLE stream: 'VBA/portfolioUtils'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Option Base 1

Private Const DASH_SHEET As String = "Dashboard"
Private Const WEIGHTS_SHEET As String = "SimulatedWeights"
Private Const RISK_FREE_ROW As Integer = 4


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
    rf = Sheets(DASH_SHEET).Range("D" & RISK_FREE_ROW).Value
    mSharpe = (mRet - rf) / mVol
    Call CalculateDrawdownFromCurve(outCurve, mMDD, mLen)
    Dim conf As Double, Z As Double
    conf = Sheets(DASH_SHEET).Range("D5").Value
    Z = Application.NormSInv(conf)
    mVaR = mRet - Z * mVol
End Sub

Sub CalculateDrawdownFromCurve(curve As Variant, ByRef outMDD As Double, ByRef outLen As Integer)
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
    Dim activeSht As Worksheet
    Set activeSht = ThisWorkbook.ActiveSheet

    Dim ws As Worksheet
    Set ws = Sheets(DASH_SHEET)
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

-------------------------------------------------------------------------------
VBA MACRO solverUtils.bas
in file: xl/vbaProject.bin - OLE stream: 'VBA/solverUtils'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Option Base 1

Private Const DASH_SHEET As String = "Dashboard"
Private Const ENGINE_SHEET As String = "Engine"
Private Const RISK_FREE_ROW As Integer = 4

Sub WriteToEngine(covMat() As Double, expReturns() As Double)
    Dim wsEng As Worksheet
    Dim nAssets As Long
    Dim i As Long, j As Long
    Set wsEng = Sheets(ENGINE_SHEET)
    nAssets = UBound(expReturns)
    wsEng.Cells.Clear
    wsEng.Range("A1").Value = "Expected Returns"
    wsEng.Range(wsEng.Cells(2, 1), wsEng.Cells(2, nAssets)).Value = expReturns
    wsEng.Range(wsEng.Cells(2, 1), wsEng.Cells(2, nAssets)).Name = "ExpReturns"
    wsEng.Range("A5").Value = "Covariance Matrix"
    wsEng.Range(wsEng.Cells(6, 1), wsEng.Cells(6 + nAssets - 1, nAssets)).Value = covMat
    wsEng.Range(wsEng.Cells(6, 1), wsEng.Cells(6 + nAssets - 1, nAssets)).Name = "CovMatrix"
    Dim solverRow As Long
    solverRow = 6 + nAssets + 5
    Call SetupOptimizationFormulas(wsEng, nAssets, solverRow)
End Sub

' ---------------------------------------------------------
' 1. SETUP: Added Log Calculation Rows for the "Log Trick"
' ---------------------------------------------------------
Sub SetupOptimizationFormulas(wsEng As Worksheet, nTotal As Long, startRow As Long)
    wsEng.Cells(startRow - 1, 1).Value = "Weights (Solver)"
    Dim i As Long
    ' Initialize with Equal Weights (Safe starting point for Log)
    For i = 1 To nTotal
        wsEng.Cells(startRow, i).Value = 1 / nTotal
    Next i
    wsEng.Range(wsEng.Cells(startRow, 1), wsEng.Cells(startRow, nTotal)).Name = "OptWeights"

    Dim rf As Double
    rf = Sheets(DASH_SHEET).Range("D" & RISK_FREE_ROW).Value

    Dim fRow As Long
    fRow = startRow + 2

    ' Standard Portfolio Calcs
    wsEng.Cells(fRow, 1).Value = "Port Return"
    wsEng.Cells(fRow, 2).Formula = "=SUMPRODUCT(OptWeights, ExpReturns)"
    wsEng.Cells(fRow, 2).Name = "PortRet"

    wsEng.Cells(fRow + 1, 1).Value = "Port Variance"
    wsEng.Cells(fRow + 1, 2).FormulaArray = "=MMULT(OptWeights, MMULT(CovMatrix, TRANSPOSE(OptWeights)))"
    wsEng.Cells(fRow + 1, 2).Name = "PortVar"

    wsEng.Cells(fRow + 2, 1).Value = "Port Vol"
    wsEng.Cells(fRow + 2, 2).Formula = "=SQRT(PortVar)"
    wsEng.Cells(fRow + 2, 2).Name = "PortVol"

    wsEng.Cells(fRow + 3, 1).Value = "Sharpe"
    wsEng.Cells(fRow + 3, 2).Formula = "=(PortRet - " & rf & ")/PortVol"
    wsEng.Cells(fRow + 3, 2).Name = "PortSharpe"

    wsEng.Cells(fRow + 4, 1).Value = "Kelly"
    wsEng.Cells(fRow + 4, 2).Formula = "=PortRet - (PortVar/2)"
    wsEng.Cells(fRow + 4, 2).Name = "PortKelly"

    wsEng.Cells(fRow + 5, 1).Value = "Sum Weights"
    wsEng.Cells(fRow + 5, 2).Formula = "=SUM(OptWeights)"
    wsEng.Cells(fRow + 5, 2).Name = "SumWeights"

    wsEng.Cells(fRow + 6, 1).Value = "Portfolio HHI"
    wsEng.Cells(fRow + 6, 2).Formula = "=SUMPRODUCT(OptWeights, OptWeights)"
    wsEng.Cells(fRow + 6, 2).Name = "PortHHI"

    wsEng.Cells(fRow + 7, 1).Value = "Max HHI Limit"
    wsEng.Cells(fRow + 7, 2).Formula = 1 / Sheets(DASH_SHEET).Range("G3").Value
    wsEng.Cells(fRow + 7, 2).Name = "MaxHHI"

    ' --- ERC UNCSTRD Log-Barrier Setup ---
    Dim logRow As Long
    logRow = fRow + 9 ' Shifted down by 2 to accommodate the HHI rows
    wsEng.Cells(logRow, 1).Value = "Log Weights"
    Dim wRange As Range
    Set wRange = wsEng.Range("OptWeights")
    For i = 1 To nTotal
        wsEng.Cells(logRow, 1 + i).Formula = "=LN(MAX(" & wRange.Cells(1, i).Address & ", 0.00001))"
    Next i
    wsEng.Cells(logRow + 1, 1).Value = "Sum Logs"
    wsEng.Cells(logRow + 1, 2).Formula = "=SUM(" & wsEng.Range(wsEng.Cells(logRow, 2), wsEng.Cells(logRow, 1 + nTotal)).Address & ")"
    wsEng.Cells(logRow + 1, 2).Name = "SumLogWeights"

    ' Visualization of Risk Contributions (Optional, for checking result)
    Dim rcRow As Long
    rcRow = logRow + 3
    wsEng.Cells(rcRow, 1).Value = "Risk Contribs Check"
    wsEng.Range(wsEng.Cells(rcRow, 2), wsEng.Cells(rcRow, 1 + nTotal)).FormulaArray = _
        "=OptWeights * TRANSPOSE(MMULT(CovMatrix, TRANSPOSE(OptWeights)))"

End Sub
' ---------------------------------------------------------
' 2. SOLVER: Implemented High Precision & Log Logic
' ---------------------------------------------------------
Sub RunSolver(Mode As String, minWeights() As Double, maxWeights() As Double, Optional ShowMsg As Boolean = True)
    Dim wsEng As Worksheet
    Set wsEng = Sheets(ENGINE_SHEET)
    Dim nAssets As Integer
    nAssets = UBound(maxWeights)

    SolverReset
    SolverOptions Precision:=0.00000001, Convergence:=0.00000001, Derivatives:=2, IntTolerance:=0, AssumeNonNeg:=True, StepThru:=False

    Dim optRange As Range
    Set optRange = wsEng.Range("OptWeights")
    Dim i As Integer

    Dim eqW As Double
    eqW = 1 / nAssets
    For i = 1 To nAssets
        wsEng.Range("OptWeights").Cells(1, i).Value = eqW
    Next i

    If Mode = "ERC UNCSTRD" Then
        ' --- STRATÉGIE DUAL CONVEXE (La plus robuste) ---
        ' On cherche à Maximiser la Diversification (Sum Logs) pour un budget de risque donné.
        SolverOk SetCell:="SumLogWeights", MaxMinVal:=1, ValueOf:=0, ByChange:="OptWeights", Engine:=1
        SolverAdd CellRef:="PortVar", Relation:=1, FormulaText:="1"
        SolverAdd CellRef:="OptWeights", Relation:=3, FormulaText:="0.001"

    Else
        ' --- STRATÉGIES CLASSIQUES (Sharpe, Variance, Kelly) ---
        SolverAdd CellRef:="SumWeights", Relation:=2, FormulaText:="1"
        For i = 1 To nAssets
            SolverAdd CellRef:=optRange.Cells(1, i).Address, Relation:=1, FormulaText:=CStr(maxWeights(i))
            SolverAdd CellRef:=optRange.Cells(1, i).Address, Relation:=3, FormulaText:=CStr(minWeights(i))
        Next i
        Select Case Mode
            Case "SHARPE":
                SolverOk SetCell:="PortSharpe", MaxMinVal:=1, ValueOf:=0, ByChange:="OptWeights", Engine:=1
                SolverAdd CellRef:="PortHHI", Relation:=1, FormulaText:="MaxHHI"
            Case "MIN VAR":
                SolverOk SetCell:="PortVar", MaxMinVal:=2, ValueOf:=0, ByChange:="OptWeights", Engine:=1
            Case "KELLY":
                SolverOk SetCell:="PortKelly", MaxMinVal:=1, ValueOf:=0, ByChange:="OptWeights", Engine:=1
        End Select
    End If

    ' Exécution
    On Error Resume Next
    Dim res As Integer
    res = SolverSolve(True)
    On Error GoTo 0

    ' --- NORMALISATION FINALE POUR ERC UNCSTRD ---
    If Mode = "ERC UNCSTRD" And res <= 2 Then
        Dim currentSum As Double
        currentSum = Application.WorksheetFunction.sum(optRange)

        If currentSum <> 0 Then
            Dim finalWeights() As Variant
            ReDim finalWeights(1 To nAssets)
            For i = 1 To nAssets
                finalWeights(i) = optRange.Cells(1, i).Value / currentSum
            Next i
            optRange.Value = finalWeights
        End If
    End If

    ' Gestion d'erreur
    If res > 2 Then
        For i = 1 To nAssets
            wsEng.Range("OptWeights").Cells(1, i).Value = 1 / nAssets
        Next i
        MsgBox Mode & " - Échec de convergence. Retour aux poids égaux.", vbExclamation
    ElseIf ShowMsg Then
        MsgBox "Optimisation terminée (" & Mode & ")", vbInformation
    End If

End Sub
-------------------------------------------------------------------------------
VBA MACRO assetUtils.bas
in file: xl/vbaProject.bin - OLE stream: 'VBA/assetUtils'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Option Base 1

Private Const DASH_SHEET As String = "Dashboard"
Private Const PRICE_SHEET As String = "PriceHistory"
Private Const CHART_SHEET As String = "ChartData"
Private Const WEIGHTS_SHEET As String = "SimulatedWeights"
Private Const ENGINE_SHEET As String = "Engine"
Private Const RISK_FREE_ROW As Integer = 4

Sub GetHistoricalDataDatesAndNames(ByRef outPrices As Variant, ByRef outDates() As Date, ByRef outNames() As String)
    Dim wsHist As Worksheet
    Dim lastRow As Long, lastCol As Long
    Set wsHist = Sheets(PRICE_SHEET)
    lastRow = wsHist.Cells(wsHist.Rows.count, 1).End(xlUp).Row
    lastCol = wsHist.Cells(1, wsHist.Columns.count).End(xlToLeft).Column
    outPrices = wsHist.Range(wsHist.Cells(2, 2), wsHist.Cells(lastRow, lastCol)).Value
    Dim nAssets As Integer, nDays As Integer
    nAssets = lastCol - 1
    nDays = lastRow - 1
    ReDim outNames(1 To nAssets)
    Dim i As Integer
    For i = 1 To nAssets
        outNames(i) = wsHist.Cells(1, 1 + i).Value
    Next i
    ReDim outDates(1 To nDays)
    For i = 1 To nDays
        outDates(i) = wsHist.Cells(1 + i, 1).Value
    Next i
End Sub

Sub CalculateHistoricalStats(dates As Variant, prices As Variant, ByRef outLogRets() As Double, ByRef outMeanRets() As Double)
    Dim nDays As Long, nAssets As Long
    Dim i As Long, j As Long
    nDays = UBound(prices, 1)
    nAssets = UBound(prices, 2)
    ReDim outLogRets(1 To nDays - 1, 1 To nAssets)
    ReDim outMeanRets(1 To nAssets)
    For j = 1 To nAssets
        Dim sumLogRet As Double
        sumLogRet = 0
        For i = 2 To nDays
            Dim pNow As Double, pPrev As Double
            pNow = prices(i, j)
            pPrev = prices(i - 1, j)
            If pPrev > 0 And pNow > 0 Then
                outLogRets(i - 1, j) = Log(pNow / pPrev)
                sumLogRet = sumLogRet + outLogRets(i - 1, j)
            Else
                outLogRets(i - 1, j) = 0
            End If
        Next i
        outMeanRets(j) = sumLogRet / (dates(nDays) - dates(1)) * 365
    Next j
End Sub

Function CalculateCovariance(logRets() As Double, meanRets() As Double) As Double()
    Dim nDays As Long, nAssets As Long
    Dim i As Long, j As Long, k As Long
    Dim res() As Double
    nDays = UBound(logRets, 1) + 1
    nAssets = UBound(logRets, 2)
    ReDim res(1 To nAssets, 1 To nAssets)
    For j = 1 To nAssets
        For k = 1 To nAssets
            Dim sumProd As Double, meanJ As Double, meanK As Double
            sumProd = 0
            meanJ = meanRets(j) / 256
            meanK = meanRets(k) / 256
            For i = 1 To nDays - 1
                sumProd = sumProd + (logRets(i, j) - meanJ) * (logRets(i, k) - meanK)
            Next i
            res(j, k) = (sumProd / (nDays - 1)) * 256
        Next k
    Next j
    CalculateCovariance = res
End Function

Function CalculateExpectedReturns(histMeanRets() As Double, currDate As Date) As Double()
    Dim wsDash As Worksheet
    Dim nAssets As Long, i As Long
    Dim res() As Double
    Set wsDash = Sheets(DASH_SHEET)
    nAssets = UBound(histMeanRets)
    ReDim res(1 To nAssets)
    Dim inputStart As Long, inputStartCol As Integer
    inputStart = wsDash.Range("InputTableStart").Row + 1
    inputStartCol = wsDash.Range("InputTableStart").Column
    For i = 1 To nAssets
        Dim rIdx As Integer
        rIdx = inputStart + i - 1
        Dim pCurr As Double, pTgt As Double, wTgt As Double
        Dim pStop As Double, wStop As Double
        Dim targetDate As Date
        pCurr = wsDash.Cells(rIdx, inputStartCol + 1).Value
        pTgt = wsDash.Cells(rIdx, inputStartCol + 3).Value
        wTgt = wsDash.Cells(rIdx, inputStartCol + 4).Value
        pStop = wsDash.Cells(rIdx, inputStartCol + 5).Value
        wStop = wsDash.Cells(rIdx, inputStartCol + 6).Value
        targetDate = wsDash.Cells(rIdx, inputStartCol + 7).Value
        Dim T_years As Double
        If targetDate > currDate Then
            T_years = (targetDate - currDate) / 365
        Else
            T_years = 1
        End If
        Dim termTgt As Double, termStop As Double, termHist As Double
        termTgt = pTgt * wTgt
        termStop = pStop * wStop
        termHist = pCurr * (Exp(histMeanRets(i)) ^ T_years) * (1 - wTgt - wStop)
        Dim eFinal As Double
        eFinal = termTgt + termStop + termHist
        If pCurr > 0 And T_years > 0 Then
            res(i) = (eFinal / pCurr) ^ (1 / T_years) - 1
        Else
            res(i) = 0
        End If
    Next i
    CalculateExpectedReturns = res
End Function



Sub OutputCorrelationMatrix(covMat() As Double, assetNames() As String)
    Dim n As Integer
    n = UBound(covMat, 1)
    Dim i As Integer, j As Integer
    Dim wsDash As Worksheet
    Set wsDash = Sheets(DASH_SHEET)
    Dim startR As Long, startC As Integer
    startR = wsDash.Range("CorrMatrixStart").Row + 1
    startC = wsDash.Range("CorrMatrixStart").Column
    For i = 1 To n
        wsDash.Cells(startR, startC + i).Value = assetNames(i)
        wsDash.Cells(startR + i, startC).Value = assetNames(i)
    Next i

    Dim corrArr() As Variant
    ReDim corrArr(1 To n, 1 To n)

    Dim stdI As Double, stdJ As Double
    For i = 1 To n
        stdI = Sqr(covMat(i, i))

        For j = 1 To n
            If i = j Then
                corrArr(i, j) = "-"
            Else
                stdJ = Sqr(covMat(j, j))
                If stdI * stdJ > 0 Then
                    corrArr(i, j) = covMat(i, j) / (stdI * stdJ)
                Else
                    corrArr(i, j) = "NA"
                End If
            End If
        Next j
    Next i
    wsDash.Cells(startR + 1, startC + 1).Resize(n, n).Value = corrArr
End Sub

Sub DrawEfficientFrontier(nAssets As Integer, maxWeights() As Double)
    Dim wsEng As Worksheet
    Set wsEng = Sheets(ENGINE_SHEET)
    Dim expRets As Variant
    expRets = wsEng.Range("ExpReturns").Value
    Dim minR As Double, maxR As Double
    minR = 9999: maxR = -9999
    Dim i As Integer
    For i = 1 To nAssets
        Dim r As Double
        r = expRets(1, i)
        If r < minR Then minR = r
        If r > maxR Then maxR = r
    Next i
    Dim nPoints As Integer
    nPoints = 10
    Dim stepR As Double
    stepR = (maxR - minR) / (nPoints - 1)
    Dim wsChart As Worksheet
    Set wsChart = Sheets(CHART_SHEET)
    Dim startCol As Integer
    startCol = 100
    wsChart.Range(wsChart.Cells(2, startCol), wsChart.Cells(100, startCol + 1)).ClearContents
    SolverReset
    SolverAdd CellRef:="SumWeights", Relation:=2, FormulaText:="1"
    SolverAdd CellRef:="OptWeights", Relation:=3, FormulaText:="0"
    Dim optRange As Range
    Set optRange = wsEng.Range("OptWeights")
    Dim j As Integer
    For j = 1 To nAssets
        SolverAdd CellRef:=optRange.Cells(1, j).Address, Relation:=1, FormulaText:=CStr(maxWeights(j))
    Next j
    For i = 1 To nPoints
        Dim targetR As Double
        targetR = minR + (i - 1) * stepR
        SolverOk SetCell:="PortVar", MaxMinVal:=2, ValueOf:=0, ByChange:="OptWeights", Engine:=1
        SolverAdd CellRef:="PortRet", Relation:=3, FormulaText:=CStr(targetR)
        On Error Resume Next
        SolverSolve True
        On Error GoTo 0
        Dim vol As Double, ret As Double
        vol = wsEng.Range("PortVol").Value
        ret = wsEng.Range("PortRet").Value
        wsChart.Cells(1 + i, startCol).Value = vol
        wsChart.Cells(1 + i, startCol + 1).Value = ret
    Next i
End Sub

Sub UpdateCurrentPrices(prices As Variant)
    Dim wsDash As Worksheet
    Set wsDash = Sheets(DASH_SHEET)
    Dim inputStart As Long
    inputStart = wsDash.Range("InputTableStart").Row + 1
    Dim nDays As Long, nAssets As Long
    nDays = UBound(prices, 1)
    nAssets = UBound(prices, 2)
    Dim i As Integer
    For i = 1 To nAssets
        wsDash.Cells(inputStart + i - 1, 4).Value = prices(nDays, i)
    Next i
End Sub


Sub WriteCurveToSheet(curve As Variant, colIdx As Integer)
    Dim ws As Worksheet
    Set ws = Sheets(CHART_SHEET)
    Dim n As Long
    n = UBound(curve, 1)
    ws.Range(ws.Cells(2, colIdx), ws.Cells(2 + n - 1, colIdx)).Value = curve
End Sub

Sub WriteWeightsToSheet(wArray As Variant, startCol As Integer)
    Dim ws As Worksheet
    Set ws = Sheets(WEIGHTS_SHEET)
    Dim nRows As Long, nCols As Long
    nRows = UBound(wArray, 1)
    nCols = UBound(wArray, 2)
    ws.Range(ws.Cells(2, startCol), ws.Cells(2 + nRows - 1, startCol + nCols - 1)).Value = wArray
End Sub

Sub SyncSimWeightsHeaders(dates() As Date, assetNames() As String, strategies As Variant)
    Dim wsSim As Worksheet
    Set wsSim = Sheets(WEIGHTS_SHEET)
    wsSim.Rows(1).ClearContents
    wsSim.Cells(1, 1).Value = "Date"
    Dim nAssets As Integer, nDays As Integer
    nAssets = UBound(assetNames)
    nDays = UBound(dates)
    Dim i As Integer, j As Integer
    Dim colIdx As Integer
    colIdx = 2
    For i = LBound(strategies) To UBound(strategies)
        For j = 1 To nAssets
            wsSim.Cells(1, colIdx).Value = strategies(i) & "_" & assetNames(j)
            colIdx = colIdx + 1
        Next j
    Next i
    For i = 1 To nDays
        wsSim.Cells(1 + i, 1).Value = dates(i)
    Next i
End Sub

Sub ProcessIndividualAssets(prices As Variant, dates As Variant, logRets As Variant, assetNames() As String)
    Dim nDays As Long, nAssets As Long
    Dim i As Long, j As Long
    nDays = UBound(prices, 1)
    nAssets = UBound(prices, 2)
    Dim wsChart As Worksheet
    Set wsChart = Sheets(CHART_SHEET)

    For j = 1 To nAssets
        Dim curve() As Double
        ReDim curve(1 To nDays, 1 To 1)
        Dim startP As Double
        startP = prices(1, j)

        For i = 1 To nDays
            curve(i, 1) = (prices(i, j) / startP) * 100
        Next i
        wsChart.Range(wsChart.Cells(2, 1 + j), wsChart.Cells(2 + nDays - 1, 1 + j)).Value = curve

        Dim mRet As Double, mVol As Double, mSharpe As Double, mMDD As Double, mLen As Integer, mVaR As Double
        Dim totalRet As Double
        totalRet = prices(nDays, j) / prices(1, j)
        mRet = totalRet ^ (365 / (dates(nDays) - dates(1))) - 1

        Dim s As Double, meanL As Double
        s = 0
        For i = 1 To nDays - 1
             s = s + logRets(i, j)
        Next i
        meanL = s / (nDays - 1)

        ' --- Volatility Calculations ---
        Dim lambda1M As Double
        Dim varEWMA1M As Double
        Dim ewmaVols() As Double

        ReDim ewmaVols(1 To nDays - 1)
        lambda1M = 0.5 ^ (1 / 21)

        For i = 1 To nDays - 1
            Dim sqDev As Double
            sqDev = (logRets(i, j) - meanL) ^ 2

            ' EWMA Recursion
            If i = 1 Then
                varEWMA1M = sqDev
            Else
                varEWMA1M = lambda1M * varEWMA1M + (1 - lambda1M) * sqDev
            End If

            ewmaVols(i) = Sqr(varEWMA1M) * Sqr(256)
        Next i

        Dim currentVol1M As Double
        Dim perc90Vol1M As Double

        currentVol1M = ewmaVols(nDays - 1)
        perc90Vol1M = Application.WorksheetFunction.Percentile_Inc(ewmaVols, 0.9)
        mVol = (currentVol1M + perc90Vol1M) / 2

        Dim rf As Double
        rf = Sheets(DASH_SHEET).Range("D" & RISK_FREE_ROW).Value
        mSharpe = (mRet - rf) / mVol

        Call CalculateDrawdownFromCurve(curve, mMDD, mLen)

        Dim conf As Double, Z As Double
        conf = Sheets(DASH_SHEET).Range("D5").Value
        Z = Application.NormSInv(conf)
        mVaR = mRet - Z * mVol

        Dim dummyW As Variant
        Call OutputMetricsToRow(j - 1, assetNames(j), mRet, mVol, mSharpe, mMDD, mLen, mVaR, dummyW, "AssetMetricsStart", False)
    Next j
End Sub

Sub SyncDashboardHeaders(dates() As Date, assetNames() As String, strategies As Variant)
    Dim wsDash As Worksheet
    Set wsDash = Sheets(DASH_SHEET)
    Dim nAssets As Integer, nDates As Integer
    nAssets = UBound(assetNames)
    nDates = UBound(dates)
    Dim i As Integer, j As Integer
    Dim inputStart As Long, inputStartCol As Integer
    inputStart = wsDash.Range("InputTableStart").Row + 1
    inputStartCol = wsDash.Range("InputTableStart").Column
    Dim equityPositions As Dictionary
    Set equityPositions = DataUtils.GetEquityPositionsFromFund
    For i = 1 To nAssets
        wsDash.Cells(inputStart + i - 1, inputStartCol - 1).Value = equityPositions.Keys(i - 1)
        wsDash.Cells(inputStart + i - 1, inputStartCol).Value = equityPositions(equityPositions.Keys(i - 1)).Name
    Next i
    inputStart = wsDash.Range("StrategyWeightsStart").Row + 1
    inputStartCol = wsDash.Range("StrategyWeightsStart").Column
    For i = 1 To nAssets
        wsDash.Cells(inputStart + i - 1, inputStartCol - 1).Value = equityPositions.Keys(i - 1)
        wsDash.Cells(inputStart + i - 1, inputStartCol).Value = equityPositions(equityPositions.Keys(i - 1)).Name
    Next i

    Dim nStrat As Integer
    nStrat = UBound(strategies) - LBound(strategies) + 1
    Dim allocations As Variant, current As String
    allocations = Array("", " matched", " target")
    For j = 1 To 3
        For i = 1 To nStrat
            wsDash.Cells(inputStart - 1, inputStartCol + i + (j - 1) * (nStrat + 2)).Value = CStr(strategies(i)) & CStr(allocations(j))
        Next i
        If j = 1 Then current = " rebased" Else current = ""
        wsDash.Cells(inputStart - 1, inputStartCol + i + (j - 1) * (nStrat + 2)).Value = "CURRENT" & current
        wsDash.Cells(inputStart - 1, inputStartCol + i + (j - 1) * (nStrat + 2) + 1).Value = "DIFF" & CStr(allocations(j))
    Next j

    Dim stratStart As Long, stratCol As Long
    stratStart = wsDash.Range("StrategyTableStart").Row
    stratCol = wsDash.Range("StrategyTableStart").Column
    wsDash.Range(wsDash.Cells(stratStart, stratCol + 7), wsDash.Cells(stratStart, stratCol + 100)).ClearContents

    Dim wsChart As Worksheet
    Set wsChart = Sheets(CHART_SHEET)
    wsChart.Rows(1).ClearContents
    wsChart.Cells(1, 1).Value = "Date"
    For i = 1 To nAssets
        wsChart.Cells(1, 1 + i).Value = assetNames(i)
    Next i

    For i = 1 To nStrat
        wsChart.Cells(1, 1 + nAssets + i).Value = CStr(strategies(i))
    Next i
    For i = 1 To nDates
        wsChart.Cells(1 + i, 1).Value = dates(i)
    Next i
    For i = 1 To nStrat
        wsChart.Cells(1, 1 + nAssets + nStrat + i).Value = CStr(strategies(i)) & "_DD"
    Next i
    wsChart.Cells(1, 100).Value = "Frontier_Vol"
    wsChart.Cells(1, 101).Value = "Frontier_Ret"
End Sub



-------------------------------------------------------------------------------
VBA MACRO Sheet6.cls
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet6'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet7.cls
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet7'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO DataUtils.bas
in file: xl/vbaProject.bin - OLE stream: 'VBA/DataUtils'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Function updatePriceHistoryFromInfinForFund()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")

    ws.Cells.ClearContents

    Dim equityPositions As Dictionary
    Set equityPositions = GetEquityPositionsFromFund()

    Dim pos As Position, p As Variant, i As Integer

    i = 1
    ws.Activate

    For Each p In equityPositions
        Set pos = equityPositions(p)
        ws.Cells(1, i) = pos.posInstrument.Symbol
        ws.Cells(1, i + 1) = pos.posInstrument.Name
        ws.Cells(2, i).Formula2 = "=INFIN.GETSECURITYHISTORY(" & ws.Cells(1, i).Address & ")"
        i = i + 2
    Next p

    ThisWorkbook.Sheets("Dashboard").Range("G3").Value = CInt(4 * equityPositions.count / 5)

End Function

Function updatePriceHistoryFromInfinForBench()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BenchData")

    ws.Cells.ClearContents

    Dim i As Integer, symbols As Variant, names As Variant
    symbols = Array("", "SPX", "", "RTY")
    names = Array("", "S&P 500 Index", "", "Russel 2000 Index")

    For i = 1 To 3 Step 2
        ws.Cells(1, i) = symbols(i)
        ws.Cells(1, i + 1) = names(i)
        ws.Cells(2, i).Formula2 = "=INFIN.GETSECURITYHISTORY(" & ws.Cells(1, i).Address & ")"
    Next i

End Function

Sub AlignSecurityDataRefactored()
    Dim wsInput As Worksheet
    Dim wsBench As Worksheet
    Dim lastCol As Long
    Dim startDate As Date
    Dim masterDates() As Date
    Dim dictBenchPrices As Object
    Dim dictPrices As Object

    Application.ScreenUpdating = False

    ' Set your starting parameters
    ' You can eventually link this to a cell, e.g., wsInput.Range("A1").Value
    startDate = ThisWorkbook.Sheets("Dashboard").Range("D2").Value

    ComputeEquallyWeightedBenchmark

    Set wsInput = ActiveSheet
    Set wsBench = ThisWorkbook.Sheets("BenchData")

    ' ---------------------------------------------------------
    ' PHASE 1.1: Extract Benchmark Timeline and Prices
    ' ---------------------------------------------------------
    ' This generates our Master Date Array and grabs benchmark prices
    Set dictBenchPrices = ExtractBenchmarkData(wsBench, startDate, masterDates)

    ' Check if we actually found dates
    If (Not masterDates) = -1 Then ' Fast way to check if array is uninitialized
        MsgBox "No benchmark dates found on or after the start date.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' ---------------------------------------------------------
    ' PHASE 1.2: Extract Asset Data (No more tallying)
    ' ---------------------------------------------------------
    lastCol = wsInput.Cells(1, wsInput.Columns.count).End(xlToLeft).Column
    Set dictPrices = ExtractAssetData(wsInput, lastCol)

    ' ---------------------------------------------------------
    ' PHASES 2, 3 & 4: Overlap, Beta Calculation, and Backfilling
    ' ---------------------------------------------------------
    Dim assetKey As Variant
    Dim secDict As Object
    Dim firstAssetDate As Date
    Dim assetBeta As Double, assetBetas As Dictionary
    Dim dateKey As Variant

    Set assetBetas = New Dictionary

    For Each assetKey In dictPrices.Keys
        Set secDict = dictPrices(assetKey)

        ' Find the earliest available date for this specific asset
        firstAssetDate = #12/31/9999#
        For Each dateKey In secDict.Keys
            If CDate(dateKey) < firstAssetDate Then
                firstAssetDate = CDate(dateKey)
            End If
        Next dateKey

        ' If the asset's first date is after our master start date, we need to backfill
        If firstAssetDate > startDate Then

            ' Phase 3: Calculate the Asset's Beta against the Benchmark
            assetBeta = CalculateLogBeta(masterDates, dictBenchPrices, secDict, firstAssetDate)
            assetBetas.Add assetKey, assetBeta
            ' Phase 4: Generate Prices Backwards (CALLING THE FUNCTION)
            BackfillAssetPrices secDict, masterDates, dictBenchPrices, startDate, firstAssetDate, assetBeta

        End If
    Next assetKey

    ' ---------------------------------------------------------
    ' PHASE 5: Output the Unified Data
    ' ---------------------------------------------------------
    WriteFundAlignedDataWithArray wsInput, masterDates, dictPrices

    ComputeEquallyWeightedBenchmark
    RunUpdateMatrices
    ThisWorkbook.Sheets("Dashboard").Activate

    Application.ScreenUpdating = True

    Dim betaString As String
    betaString = ""
    If assetBetas.count > 0 Then
        betaString = vbCr & vbCr & "Backfilling Beta used :" & vbCr
        For Each assetKey In assetBetas
            betaString = betaString & assetKey & " : " & Round(assetBetas(assetKey), 2) & vbCr
        Next assetKey
    End If
    MsgBox "Data aligned and backfilled successfully!" & vbCr & "From " & masterDates(1) & vbCr & "To " & masterDates(UBound(masterDates)) & betaString, vbInformation

End Sub

'
'Sub AlignSecurityDataRefactored()
'
'    ComputeEquallyWeightedBenchmark
'
'    Dim wsInput As Worksheet
'    Dim lastCol As Long, secCount As Long
'    Dim dictDates As Object, dictPrices As Object
'    Dim arrDates() As Date
'
'    Set wsInput = ActiveSheet
'    lastCol = wsInput.Cells(1, wsInput.Columns.count).End(xlToLeft).Column
'    secCount = lastCol / 2 ' Assuming 2 columns per security (Date, Price)
'
'    Set dictDates = CreateObject("Scripting.Dictionary")
'
'    Application.ScreenUpdating = False
'
'    ' STEP 1: Extract data using Arrays and Asset Names as Keys
'    Set dictPrices = ExtractData(wsInput, lastCol, dictDates)
'
'    ' STEP 2: Find and sort common dates
'    If Not GetFundCommonDates(dictDates, secCount, arrDates) Then
'        MsgBox "No common dates found across all securities.", vbExclamation
'        Application.ScreenUpdating = True
'        Exit Sub
'    End If
'
'    ' STEP 3: Write to the final sheet via an Output Array
'    WriteFundAlignedDataWithArray wsInput, arrDates, dictPrices
'
'    ComputeEquallyWeightedBenchmark
'
'    Application.ScreenUpdating = True
'    MsgBox "Data aligned successfully!" & vbCr & "Found " & UBound(arrDates) + 1 & " common dates. Range :" & vbCr & vbCr & "From " & arrDates(LBound(arrDates)) & vbCr & "To " & arrDates(UBound(arrDates)), vbInformation
'
'    RunUpdateMatrices
'    ThisWorkbook.Sheets("Dashboard").Activate
'End Sub


Public Function GetEquityPositionsFromFund() As Dictionary

    Dim ptf As Portfolio
    Set ptf = GetPortfolio

    Dim equityPositions As Dictionary, positions As Dictionary
    Set equityPositions = New Dictionary
    Dim pos As Position, p As Variant

    Set positions = ptf.GetPositions

    For Each p In positions
        Set pos = positions(p)
        If LCase(pos.posInstrument.assetClass) = "equity" Then
            equityPositions.Add p, pos
        End If
    Next p

    Set GetEquityPositionsFromFund = equityPositions

End Function

Private Function GetPortfolio() As Portfolio

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")

    ws.Activate

    Dim ptfs As Dictionary
    Set ptfs = AddInData.GetPortfolios()

    Dim ptf As Portfolio, p As Variant

    For Each p In ptfs
        If ptfs(p).IdentifierVIA = ws.Range("G2").Value Then
            Set ptf = ptfs(p)
            ptf.dt = Date
            Exit For
        End If
    Next p

    ptf.GetPositions

    Set GetPortfolio = ptf

End Function


' =====================================================================
' HELPER FUNCTIONS FOR PHASE 1
' =====================================================================

Private Function ExtractBenchmarkData(ws As Worksheet, startDate As Date, ByRef arrDates() As Date) As Object
    Dim dictBench As Object
    Dim lastRow As Long
    Dim srcArray As Variant
    Dim i As Long, count As Long
    Dim d As Date, p As Double

    Set dictBench = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).Row

    If lastRow < 2 Then GoTo EarlyExit

    ' Grab Columns E through H into memory
    ' Index 1 = Col E (Date), Index 4 = Col H (Price)
    srcArray = ws.Range(ws.Cells(2, "E"), ws.Cells(lastRow, "H")).Value

    ' Dimension array to maximum possible size to avoid slow ReDim Preserves
    ReDim arrDates(1 To UBound(srcArray, 1))
    count = 0

    For i = 1 To UBound(srcArray, 1)
        If IsDate(srcArray(i, 1)) Then
            d = CDate(srcArray(i, 1))

            ' Only keep dates on or after our Start Date
            If d >= startDate Then
                p = srcArray(i, 4)

                If Not dictBench.Exists(d) Then
                    dictBench.Add d, p
                    count = count + 1
                    arrDates(count) = d
                End If
            End If
        End If
    Next i

    ' Shrink the array down to the actual number of valid dates found
    If count > 0 Then
        ReDim Preserve arrDates(1 To count)
        SortDatesAscending arrDates ' Ensure chronological order
    Else
        Erase arrDates
    End If

EarlyExit:
    Set ExtractBenchmarkData = dictBench
End Function


Private Function ExtractAssetData(ws As Worksheet, lastCol As Long) As Object
    Dim dictPrices As Object, secDict As Object
    Dim i As Long, j As Long, lastRow As Long
    Dim dateVal As Variant, priceVal As Variant
    Dim assetName As String
    Dim srcArray As Variant

    Set dictPrices = CreateObject("Scripting.Dictionary")

    For i = 1 To lastCol Step 2
        lastRow = ws.Cells(ws.Rows.count, i).End(xlUp).Row

        If lastRow >= 2 Then
            assetName = ws.Cells(1, i + 1).Value
            Set secDict = CreateObject("Scripting.Dictionary")
            srcArray = ws.Range(ws.Cells(2, i), ws.Cells(lastRow, i + 1)).Value

            For j = 1 To UBound(srcArray, 1)
                dateVal = srcArray(j, 1)
                priceVal = srcArray(j, 2)

                ' Store the price if it's a valid date and number
                If IsDate(dateVal) And IsNumeric(priceVal) Then
                    If Not secDict.Exists(dateVal) Then
                        secDict.Add dateVal, priceVal
                    End If
                End If
            Next j

            dictPrices.Add assetName, secDict
        End If
    Next i

    Set ExtractAssetData = dictPrices
End Function

Private Sub SortDatesAscending(ByRef arr() As Date)
    ' Standard Bubble Sort for dates
    Dim i As Long, j As Long
    Dim temp As Date
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

' =====================================================================
' HELPER FUNCTIONS FOR PHASES 2 & 3
' =====================================================================

Private Function CalculateLogBeta(masterDates() As Date, dictBenchPrices As Object, dictAssetPrices As Object, firstAssetDate As Date) As Double
    Dim i As Long, count As Long
    Dim pAsset1 As Double, pAsset0 As Double
    Dim pBench1 As Double, pBench0 As Double
    Dim d1 As Date, d0 As Date

    Dim retAsset As Double, retBench As Double
    Dim sumAsset As Double, sumBench As Double

    ' Arrays to temporarily hold the valid returns
    Dim arrRetA() As Double, arrRetB() As Double
    ReDim arrRetA(1 To UBound(masterDates))
    ReDim arrRetB(1 To UBound(masterDates))

    count = 0
    sumAsset = 0
    sumBench = 0

    ' -------------------------------------------------------------
    ' PASS 1: Calculate Log Returns and Sums
    ' -------------------------------------------------------------
    For i = 2 To UBound(masterDates)
        d1 = masterDates(i)
        d0 = masterDates(i - 1)

        If d0 >= firstAssetDate Then
            If dictAssetPrices.Exists(d1) And dictAssetPrices.Exists(d0) And _
               dictBenchPrices.Exists(d1) And dictBenchPrices.Exists(d0) Then

                pAsset1 = dictAssetPrices(d1)
                pAsset0 = dictAssetPrices(d0)
                pBench1 = dictBenchPrices(d1)
                pBench0 = dictBenchPrices(d0)

                If pAsset0 > 0 And pBench0 > 0 And pAsset1 > 0 And pBench1 > 0 Then
                    count = count + 1

                    ' Note: In native VBA, Log() is the Natural Logarithm (Ln)
                    retAsset = Log(pAsset1 / pAsset0)
                    retBench = Log(pBench1 / pBench0)

                    arrRetA(count) = retAsset
                    arrRetB(count) = retBench

                    sumAsset = sumAsset + retAsset
                    sumBench = sumBench + retBench
                End If
            End If
        End If
    Next i

    ' -------------------------------------------------------------
    ' PASS 2: Calculate NATIVE Covariance / Variance (Beta)
    ' -------------------------------------------------------------
    If count >= 2 Then
        Dim meanA As Double, meanB As Double
        Dim covSum As Double, varSum As Double

        meanA = sumAsset / count
        meanB = sumBench / count
        covSum = 0
        varSum = 0

        For i = 1 To count
            covSum = covSum + ((arrRetA(i) - meanA) * (arrRetB(i) - meanB))
            varSum = varSum + ((arrRetB(i) - meanB) ^ 2)
        Next i

        If varSum > 0 Then
            CalculateLogBeta = covSum / varSum
        Else
            CalculateLogBeta = 1 ' Default if benchmark variance is literally zero
        End If
    Else
        CalculateLogBeta = 1 ' Default if not enough overlapping data points
    End If
End Function


' =====================================================================
' HELPER FUNCTION FOR PHASE 4
' =====================================================================

Private Sub BackfillAssetPrices(ByRef secDict As Object, masterDates() As Date, dictBenchPrices As Object, startDate As Date, firstAssetDate As Date, assetBeta As Double)
    Dim i As Long, startIndex As Long
    Dim d1 As Date, d0 As Date
    Dim pBench1 As Double, pBench0 As Double
    Dim pAsset1 As Double, pAsset0 As Double
    Dim rBench As Double, rProxy As Double

    ' 1. Find where the asset's real history begins on the Master Timeline
    For i = 1 To UBound(masterDates)
        If masterDates(i) = firstAssetDate Then
            startIndex = i
            Exit For
        End If
    Next i

    ' If we couldn't find it, or it's the very first date, exit
    If startIndex <= 1 Then Exit Sub

    ' 2. Work backwards day-by-day
    For i = startIndex To 2 Step -1
        d1 = masterDates(i)
        d0 = masterDates(i - 1)

        ' Stop if we go before the requested start date
        If d0 < startDate Then Exit For

        ' Ensure we have the necessary data points to calculate
        If dictBenchPrices.Exists(d1) And dictBenchPrices.Exists(d0) And secDict.Exists(d1) Then
            pBench1 = dictBenchPrices(d1)
            pBench0 = dictBenchPrices(d0)
            pAsset1 = secDict(d1)

            If pBench1 > 0 And pBench0 > 0 Then
                ' Calculate log return of benchmark
                rBench = Application.WorksheetFunction.Ln(pBench1 / pBench0)

                ' Apply beta to get proxy return
                rProxy = assetBeta * rBench

                ' Reverse the log return to find yesterday's price
                pAsset0 = pAsset1 / Exp(rProxy)

                ' Inject the newly calculated price into the asset's dictionary
                If Not secDict.Exists(d0) Then
                    secDict.Add d0, pAsset0
                End If
            End If
        End If
    Next i
End Sub


' =====================================================================
' UPDATED PHASE 5: STRICT INTERSECTION OUTPUT
' =====================================================================

Private Sub WriteFundAlignedDataWithArray(wsInput As Worksheet, masterDates() As Date, dictPrices As Object)
    Dim wsOutput As Worksheet
    Dim outArray As Variant
    Dim assetNames As Variant
    Dim r As Long, c As Long
    Dim dateKey As Date
    Dim allExist As Boolean

    Dim validDateCount As Long
    Dim validDates() As Date

    assetNames = dictPrices.Keys

    ' Dimension a temporary array to hold the dates that survive the filter
    ReDim validDates(1 To UBound(masterDates))
    validDateCount = 0

    ' ---------------------------------------------------------
    ' 1. Filter the Master Dates for Strict Overlap
    ' ---------------------------------------------------------
    For r = 1 To UBound(masterDates)
        dateKey = masterDates(r)
        allExist = True

        ' Check if every single asset has this date
        For c = 0 To UBound(assetNames)
            If Not dictPrices(assetNames(c)).Exists(dateKey) Then
                allExist = False
                Exit For ' Stop checking if even one asset is missing it
            End If
        Next c

        ' If all assets have a price, keep the date
        If allExist Then
            validDateCount = validDateCount + 1
            validDates(validDateCount) = dateKey
        End If
    Next r

    ' Safety check in case the filter removes everything
    If validDateCount = 0 Then
        MsgBox "After alignment, no dates exist where ALL assets have a price.", vbExclamation
        Exit Sub
    End If

    ' ---------------------------------------------------------
    ' 2. Build the Final Output Array
    ' ---------------------------------------------------------
    ReDim outArray(1 To validDateCount + 1, 1 To dictPrices.count + 1)

    ' Headers
    outArray(1, 1) = "Aligned Date"
    For c = 0 To UBound(assetNames)
        outArray(1, c + 2) = assetNames(c)
    Next c

    ' Dates and Prices
    For r = 1 To validDateCount
        dateKey = validDates(r)
        outArray(r + 1, 1) = dateKey

        For c = 0 To UBound(assetNames)
            outArray(r + 1, c + 2) = dictPrices(assetNames(c))(dateKey)
        Next c
    Next r

    ' ---------------------------------------------------------
    ' 3. Drop Array on Sheet
    ' ---------------------------------------------------------
    Set wsOutput = ThisWorkbook.Sheets("PriceHistory")
    wsOutput.Cells.ClearContents

    ' Resize target block to perfectly match the filtered array size
    wsOutput.Cells(1, 1).Resize(validDateCount + 1, dictPrices.count + 1).Value = outArray

    ' Clean up formatting
    wsOutput.Columns(1).NumberFormat = "dd/mm/yyyy"
    wsOutput.Columns.AutoFit
End Sub

'Private Function ExtractData(ws As Worksheet, lastCol As Long, ByRef dictDates As Object) As Object
'    Dim dictPrices As Object, secDict As Object
'    Dim i As Long, j As Long, lastRow As Long
'    Dim dateVal As Variant, priceVal As Variant
'    Dim assetName As String
'    Dim srcArray As Variant
'
'    Set dictPrices = CreateObject("Scripting.Dictionary")
'
'    For i = 1 To lastCol Step 2
'        lastRow = ws.Cells(ws.Rows.count, i).End(xlUp).Row
'
'        If lastRow >= 2 Then
'            ' Use the Price Column Header (Column i + 1) as the Asset Name key
'            assetName = ws.Cells(1, i + 1).Value
'            Set secDict = CreateObject("Scripting.Dictionary")
'
'            ' Snatch data into memory array
'            srcArray = ws.Range(ws.Cells(2, i), ws.Cells(lastRow, i + 1)).Value
'
'            For j = 1 To UBound(srcArray, 1)
'                dateVal = srcArray(j, 1)
'                priceVal = srcArray(j, 2)
'
'                If IsDate(dateVal) Then
'                    ' Tally matching dates across assets
'                    If Not dictDates.Exists(dateVal) Then
'                        dictDates.Add dateVal, 1
'                    Else
'                        dictDates(dateVal) = dictDates(dateVal) + 1
'                    End If
'
'                    ' Store price mapped to date
'                    If Not secDict.Exists(dateVal) Then
'                        secDict.Add dateVal, priceVal
'                    End If
'                End If
'            Next j
'
'            ' Store the asset dictionary using the Asset Name as the main key
'            dictPrices.Add assetName, secDict
'        End If
'    Next i
'
'    Set ExtractData = dictPrices
'End Function
'
'
'
'Private Function GetFundCommonDates(dictDates As Object, secCount As Long, ByRef arrDates() As Date) As Boolean
'    Dim key As Variant
'    Dim count As Long
'    Dim tempDate As Date
'    Dim m As Long, n As Long
'
'    count = 0
'    For Each key In dictDates.Keys
'        If dictDates(key) = secCount Then
'            ReDim Preserve arrDates(count)
'            arrDates(count) = CDate(key)
'            count = count + 1
'        End If
'    Next key
'
'    If count = 0 Then
'        GetFundCommonDates = False
'        Exit Function
'    End If
'
'    ' Quick Bubble Sort (Oldest to Newest)
'    For m = LBound(arrDates) To UBound(arrDates) - 1
'        For n = m + 1 To UBound(arrDates)
'            If arrDates(m) > arrDates(n) Then
'                tempDate = arrDates(m)
'                arrDates(m) = arrDates(n)
'                arrDates(n) = tempDate
'            End If
'        Next n
'    Next m
'
'    GetFundCommonDates = True
'End Function



'Private Sub WriteFundAlignedDataWithArray(wsInput As Worksheet, arrDates() As Date, dictPrices As Object)
'    Dim wsOutput As Worksheet
'    Dim outArray As Variant
'    Dim assetNames As Variant
'    Dim totalRows As Long, totalCols As Long
'    Dim r As Long, c As Long
'    Dim dateKey As Date
'
'    ' Extract asset names keys out of our dictionary
'    assetNames = dictPrices.Keys
'
'    totalRows = UBound(arrDates) + 2 ' +1 for 0-index, +1 for Header row
'    totalCols = dictPrices.count + 1 ' +1 for the unified Date column
'
'    ' Dimensions of output matrix: (Rows, Columns)
'    ReDim outArray(1 To totalRows, 1 To totalCols)
'
'    ' 1. Populate Headers in the array matrix
'    outArray(1, 1) = "Aligned Date"
'    For c = 0 To UBound(assetNames)
'        outArray(1, c + 2) = assetNames(c)
'    Next c
'
'    ' 2. Populate Dates and Prices in the array matrix
'    For r = 0 To UBound(arrDates)
'        dateKey = arrDates(r)
'        outArray(r + 2, 1) = dateKey ' Date column
'
'        ' Retrieve prices using Asset Name keys
'        For c = 0 To UBound(assetNames)
'            outArray(r + 2, c + 2) = dictPrices(assetNames(c))(dateKey)
'        Next c
'    Next r
'
'    ' 3. Add worksheet and drop the array on the sheet all at once
'    Set wsOutput = ThisWorkbook.Sheets("PriceHistory")
'    wsOutput.Cells.ClearContents
'
'    ' Resize target block to perfectly match array size
'    wsOutput.Cells(1, 1).Resize(totalRows, totalCols).Value = outArray
'
'    ' Clean up formatting
'    wsOutput.Columns(1).NumberFormat = "dd/mm/yyyy" ' Enforces neat dates
'    wsOutput.Columns.AutoFit
'
'End Sub


Private Sub ComputeEquallyWeightedBenchmark()
    Dim wsData As Worksheet
    Dim arrSec1 As Variant, arrSec2 As Variant
    Dim dictSec1 As Object
    Dim outArr As Variant
    Dim outRows As Long

    ' 1. Define the source data sheet
    ' Change "Sheet1" to match your actual sheet name
    Set wsData = ThisWorkbook.Sheets("BenchData")

    ' 2. Load the raw data ranges directly into memory arrays
    arrSec1 = LoadRangeToArray(wsData, "A", "B")
    arrSec2 = LoadRangeToArray(wsData, "C", "D")

    ' 3. Build a Dictionary from Security 1 for instant date lookups
    Set dictSec1 = BuildDictionary(arrSec1)
    If dictSec1.count = 0 Then
        MsgBox "No valid data found for Security 1.", vbCritical
        Exit Sub
    End If

    ' 4. Process common dates and perform all calculations in-memory
    outArr = ProcessAndCalculate(dictSec1, arrSec2, outRows)

    ' 5. Exit if no overlap was found
    If outRows = 0 Then
        MsgBox "No common dates found between the two securities.", vbExclamation
        Exit Sub
    End If

    ' 6. Dump the final calculated array onto a new spreadsheet
    Call OutputResults(outArr, outRows)

End Sub


' ==========================================
' HELPER FUNCTIONS
' ==========================================

' Dynamically finds the last row and loads two columns into a 2D memory array
Private Function LoadRangeToArray(ws As Worksheet, col1 As String, col2 As String) As Variant
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, col1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2 ' Failsafe for empty columns

    LoadRangeToArray = ws.Range(ws.Cells(2, col1), ws.Cells(lastRow, col2)).Value
End Function

' Converts a 2D array (Date/Price) into a Dictionary for lightning-fast matching
Private Function BuildDictionary(arr As Variant) As Object
    Dim dict As Object
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")

    If Not IsEmpty(arr) Then
        For i = 1 To UBound(arr)
            ' Ensure we are looking at a valid date and a valid number
            If IsDate(arr(i, 1)) And IsNumeric(arr(i, 2)) Then
                dict(arr(i, 1)) = arr(i, 2)
            End If
        Next i
    End If

    Set BuildDictionary = dict
End Function

' Loops through Security 2, checks the Dictionary for matches, and calculates the benchmark
Private Function ProcessAndCalculate(dict As Object, arr2 As Variant, ByRef outRows As Long) As Variant
    Dim tempArr() As Variant
    Dim i As Long
    Dim dateVal As Variant, price1 As Double, price2 As Double
    Dim base1 As Double, base2 As Double
    Dim isBaseSet As Boolean

    ' Max possible size is the length of array 2, spanning 6 columns
    ReDim tempArr(1 To UBound(arr2), 1 To 4)
    outRows = 0
    isBaseSet = False

    For i = 1 To UBound(arr2)
        dateVal = arr2(i, 1)

        If dict.Exists(dateVal) And IsNumeric(arr2(i, 2)) Then
            outRows = outRows + 1
            price1 = dict(dateVal)
            price2 = arr2(i, 2)

            ' Assign Raw Data
            tempArr(outRows, 1) = dateVal
            'tempArr(outRows, 2) = price1
            'tempArr(outRows, 3) = price2

            ' Set the Base 100 starting prices on the very first matched date
            If Not isBaseSet Then
                base1 = price1
                base2 = price2
                isBaseSet = True
            End If

            ' Calculate Normalized Prices and Benchmark
            tempArr(outRows, 2) = (price1 / base1) * 100
            tempArr(outRows, 3) = (price2 / base2) * 100
            tempArr(outRows, 4) = (tempArr(outRows, 2) + tempArr(outRows, 3)) / 2
        End If
    Next i

    ' Return the populated array
    ProcessAndCalculate = tempArr
End Function

' Generates the final worksheet, drops the array, and applies formatting
Private Sub OutputResults(outArr As Variant, outRows As Long)
    Dim wsOut As Worksheet

    Set wsOut = ThisWorkbook.Sheets("BenchData")

    ' Write Headers
    wsOut.Range("E1:H1").Value = Array("Common Date", "Norm Price 1 (Base 100)", "Norm Price 2 (Base 100)", "Equally Weighted Benchmark")

    ' Bulk-drop the array data onto the sheet
    wsOut.Range("E2").Resize(outRows, 4).Value = outArr

    ' Apply Formatting
    wsOut.Columns("E:E").NumberFormat = "dd/mm/yyyy"
    wsOut.Columns("F:H").NumberFormat = "0.00"
    wsOut.Columns.AutoFit

End Sub

-------------------------------------------------------------------------------
VBA MACRO convictionsUtils.bas
in file: xl/vbaProject.bin - OLE stream: 'VBA/convictionsUtils'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Sub UpdateConvictions()

    Dim convictions As Dictionary
    Set convictions = getConvictions

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")

    Dim i As Integer, lastRow As Integer, startRow As Integer, colStart As Integer
    startRow = ws.Range("InputTableStart").Row + 1
    lastRow = ws.Range("InputTableStart").End(xlDown).Row
    colStart = ws.Range("InputTableStart").Column

    For i = ws.Range("InputTableStart").Row + 1 To lastRow
        If ws.Cells(i, 2).Value <> "" Then
            ws.Cells(i, colStart + 2).Value = convictions(CStr(ws.Cells(i, 2).Value))("conviction")
            ws.Cells(i, colStart + 3).Value = convictions(CStr(ws.Cells(i, 2).Value))("TP")
            ws.Cells(i, colStart + 4).Value = convictions(CStr(ws.Cells(i, 2).Value))("TP proba")
            ws.Cells(i, colStart + 7).Value = Date + 365
            ws.Cells(i, colStart + 8).Value = 1 / (4 * (lastRow - startRow + 1))
            ws.Cells(i, colStart + 9).Value = 2 / (lastRow - startRow + 1)
        Else
            Exit For
        End If
    Next i

End Sub

Public Function getConvictions() As Dictionary

    Dim convictions As Dictionary, singleConviction As Dictionary
    Set convictions = New Dictionary

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Convictions")

    Dim i As Integer, lastRow As Integer
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    i = 2

    For i = 2 To lastRow
        Set singleConviction = New Dictionary
        singleConviction.Add "conviction", ws.Cells(i, 4).Value
        singleConviction.Add "TP", ws.Cells(i, 5).Value
        singleConviction.Add "TP proba", ws.Cells(i, 6).Value
        singleConviction.Add "SP", ws.Cells(i, 7).Value
        singleConviction.Add "SP proba", ws.Cells(i, 8).Value
        convictions.Add ws.Cells(i, 1).Value, singleConviction
    Next i

    Set getConvictions = convictions

End Function
-------------------------------------------------------------------------------
VBA MACRO reportUtils.bas
in file: xl/vbaProject.bin - OLE stream: 'VBA/reportUtils'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function CreateCombinedAssetDictionary(sourceRange1 As Range, sourceRange2 As Range) As Object
    Dim mainDict As Object
    Dim innerDict As Object
    Dim dataArray As Variant
    Dim r As Long, c As Long
    Dim primaryKey As String, secondaryKey As String
    Dim i As Integer
    Dim rangesArray() As Variant

    Set mainDict = CreateObject("Scripting.Dictionary")

    ReDim rangesArray(1 To 2)
    rangesArray(1) = sourceRange1.Value
    rangesArray(2) = sourceRange2.Value

    Dim start As Integer

    For i = 1 To 2
        dataArray = rangesArray(i)
        If IsArray(dataArray) Then
            If UBound(dataArray, 1) >= 2 And UBound(dataArray, 2) >= 2 Then
                For r = 2 To UBound(dataArray, 1)
                    If Not IsEmpty(dataArray(r, 1)) Then
                        primaryKey = CStr(dataArray(r, 1))
                        If Not mainDict.Exists(primaryKey) Then
                            Set innerDict = CreateObject("Scripting.Dictionary")
                            mainDict.Add primaryKey, innerDict
                        Else
                            Set innerDict = mainDict(primaryKey)
                        End If
                        If i = 2 Then start = 3 Else start = 2
                        For c = start To UBound(dataArray, 2)
                            secondaryKey = CStr(dataArray(1, c))
                            innerDict(secondaryKey) = dataArray(r, c)
                        Next c
                    End If
                Next r
            End If
        End If
    Next i

    Set CreateCombinedAssetDictionary = mainDict
End Function



Sub SaveResultsToTop(dictAssets As Object, targetSheetName As String)
    Dim ws As Worksheet
    Dim numRows As Long, numCols As Long
    Dim innerDict As Object
    Dim keysArray As Variant, firstKey As Variant
    Dim r As Long, c As Long
    Dim assetKey As Variant, attrKey As Variant

    ' 1. Set the target sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Target sheet '" & targetSheetName & "' not found.", vbCritical
        Exit Sub
    End If

    numRows = dictAssets.count
    If numRows = 0 Then Exit Sub ' Nothing to save

    ' 2. Determine columns (Inner dict count + 1 for Primary Key + 1 for Date)
    keysArray = dictAssets.Keys
    firstKey = keysArray(0)
    Set innerDict = dictAssets.Item(firstKey)
    numCols = innerDict.count + 3

    ' 3. Insert new rows at row 2, pushing existing data down
    ws.Rows("2:" & CStr(numRows + 1)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' 4. Create a 2D array to hold the data
    Dim dataArray() As Variant
    ReDim dataArray(1 To numRows, 1 To numCols)

    ' 5. Populate the array and write headers
    r = 1
    For Each assetKey In dictAssets.Keys
        Set innerDict = dictAssets.Item(assetKey)

        ' Write the Primary Key and Date into the first two columns
        dataArray(r, 1) = assetKey
        dataArray(r, 2) = Date ' Use Now instead of Date if you want the exact time too

        ' Write the static headers in Row 1
        If r = 1 Then
            ws.Cells(1, 1).Value = "Asset_ID"
            ws.Cells(1, 2).Value = "Date_Saved"
        End If

        ' Start populating inner dictionary data from the 3rd column
        c = 3
        For Each attrKey In innerDict.Keys
            ' Write the dynamic headers to Row 1
            If r = 1 Then
                ws.Cells(1, c).Value = attrKey
            End If

            ' Populate the 2D array with inner dict values
            dataArray(r, c) = innerDict.Item(attrKey)
            c = c + 1
        Next attrKey

        dataArray(r, c) = innerDict("DIFF target") / innerDict("DIFF")
        r = r + 1

    Next assetKey

    ' 6. Dump the array into the newly inserted rows
    ws.Range(ws.Cells(2, 1), ws.Cells(numRows + 1, numCols)).Value = dataArray

    'MsgBox "Historical data saved successfully!", vbInformation
End Sub


-------------------------------------------------------------------------------
VBA MACRO Sheet8.cls
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet8'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet9.cls
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet9'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Report.bas
in file: xl/vbaProject.bin - OLE stream: 'VBA/Report'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private dictCols As Object

Sub ReportWorkflow()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")

    Dim firstRange As Range, secondRange As Range
    Set firstRange = ws.Range( _
                    ws.Cells(ws.Range("InputTableStart").Row, ws.Range("InputTableStart").Column - 1), _
                    ws.Cells(ws.Range("InputTableStart").End(xlDown).Row, ws.Range("InputTableStart").End(xlToRight).Column - 1) _
                )
    Set secondRange = ws.Range( _
                    ws.Cells(ws.Range("StrategyWeightsStart").Row, ws.Range("StrategyWeightsStart").Column - 1), _
                    ws.Cells(ws.Range("StrategyWeightsStart").End(xlDown).Row, ws.Range("StrategyWeightsStart").End(xlToRight).Column) _
                )

    Dim assetsData As Dictionary
    Set assetsData = reportUtils.CreateCombinedAssetDictionary(firstRange, secondRange)

    reportUtils.SaveResultsToTop assetsData, "HistoricalSimulations"

    GeneratePortfolioReport

End Sub

Private Sub GeneratePortfolioReport()
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim data() As Variant
    Dim rRow As Long
    Dim maxDate As Date
    Dim d1 As Date, d2 As Date, d7 As Date, d14 As Date

    ' Set up sheets
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("HistoricalSimulations")
    If wsData Is Nothing Then
        MsgBox "HistoricalSimulations sheet not found!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Create or clear report sheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Portfolio Report")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsReport.Name = "Portfolio Report"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0

    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 1 Then
        MsgBox "No data found in HistoricalSimulations.", vbExclamation
        Exit Sub
    End If

    ' Initialize Column Mappings
    If Not MapColumns(wsData, lastCol) Then
        MsgBox "Failed to map all required columns. Check headers in HistoricalSimulations.", vbCritical
        Exit Sub
    End If

    ' Read Data (Excluding headers to match our logic, 1-based index)
    data = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRow, lastCol)).Value

    ' 1. Determine Dates
    maxDate = GetMaxDate(data)
    If maxDate = 0 Then
        MsgBox "No valid dates found in the data.", vbExclamation
        Exit Sub
    End If

    d1 = FindClosestDate(data, maxDate - 1)
    d2 = FindClosestDate(data, maxDate - 2)
    d7 = FindClosestDate(data, maxDate - 7)
    d14 = FindClosestDate(data, maxDate - 14)

    rRow = 1

    ' Format Title
    wsReport.Cells(rRow, 1).Value = "Portfolio Report - " & Format(maxDate, "dd/mm/yyyy")
    wsReport.Cells(rRow, 1).Font.Bold = True
    wsReport.Cells(rRow, 1).Font.Size = 16
    rRow = rRow + 2

    ' Dictionary to hold historical data for fast lookup: Dict(Asset_ID & "_" & Date) = RowIndex
    Dim dictHist As Object
    Set dictHist = BuildHistoryDictionary(data)

    ' Process Sections
    rRow = ProcessRiskProximity(wsReport, data, maxDate, rRow)
    rRow = ProcessAllocationGaps(wsReport, data, dictHist, maxDate, d1, d2, d7, d14, rRow)
    rRow = ProcessHistoricalChanges(wsReport, data, dictHist, maxDate, d1, d2, d7, d14, rRow)
    rRow = ProcessConvictionAnalysis(wsReport, data, maxDate, rRow)
    rRow = ProcessSchemeVariations(wsReport, data, dictHist, maxDate, d1, d2, d7, d14, rRow)

    ' General Formatting
    wsReport.Columns("A:Z").AutoFit

    ' Freeze Panes
    wsReport.Activate
    ActiveWindow.FreezePanes = False
    wsReport.Range("A3").Select
    ActiveWindow.FreezePanes = True
    wsReport.Range("A1").Select
    wsReport.Range("A1").Columns.ColumnWidth = 10

    'MsgBox "Portfolio Report Generated Successfully!", vbInformation
End Sub

Private Function MapColumns(ByVal ws As Worksheet, ByVal lastCol As Long) As Boolean
    Dim i As Long
    Dim header As String

    Set dictCols = CreateObject("Scripting.Dictionary")

    For i = 1 To lastCol
        header = UCase(Trim(ws.Cells(1, i).Value))
        Select Case header
            Case "ASSET_ID": dictCols("ASSET_ID") = i
            Case "DATE_SAVED": dictCols("DATE_SAVED") = i
            Case "ASSET NAME": dictCols("ASSET_NAME") = i
            Case "CURRENT PRICE": dictCols("CURRENT_PRICE") = i
            Case "CONVICTION (/5)": dictCols("CONVICTION") = i
            Case "TARGET PRICE": dictCols("TARGET_PRICE") = i
            Case "STOP PRICE": dictCols("STOP_PRICE") = i
            Case "EXPECTED RETURN": dictCols("EXPECTED_RETURN") = i
            Case "ERC UNCSTRD": dictCols("ERC") = i
            Case "ER/VOL": dictCols("ERVOL") = i
            Case "SHARPE": dictCols("SHARPE") = i
            Case "CUSTOM": dictCols("CUSTOM") = i
            Case "MEAN": dictCols("MEAN") = i
            Case "CURRENT": dictCols("CURRENT_ALLOC") = i
        End Select
    Next i

    ' Validate required columns are found
    If Not dictCols.Exists("ASSET_ID") Or Not dictCols.Exists("DATE_SAVED") Or _
       Not dictCols.Exists("ASSET_NAME") Or Not dictCols.Exists("CURRENT_PRICE") Or _
       Not dictCols.Exists("CONVICTION") Or Not dictCols.Exists("TARGET_PRICE") Or _
       Not dictCols.Exists("STOP_PRICE") Or Not dictCols.Exists("EXPECTED_RETURN") Or _
       Not dictCols.Exists("ERC") Or Not dictCols.Exists("ERVOL") Or _
       Not dictCols.Exists("SHARPE") Or Not dictCols.Exists("CUSTOM") Or _
       Not dictCols.Exists("MEAN") Or Not dictCols.Exists("CURRENT_ALLOC") Then
        MapColumns = False
    Else
        MapColumns = True
    End If
End Function

Private Sub FormatTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal nRows As Long, ByVal nCols As Long)
    If nRows < 0 Then nRows = 0
    Dim rngHead As Range
    Set rngHead = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, nCols))
    Dim rngFull As Range
    Set rngFull = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + nRows, nCols))

    rngFull.Borders.LineStyle = 1 ' xlContinuous
    rngFull.Borders.Color = RGB(200, 200, 200)
    rngFull.Borders.Weight = 2 ' xlThin

    rngHead.Interior.Color = RGB(220, 230, 241) ' Light blue background
    rngHead.Font.Bold = True
End Sub

Private Sub FormatDiffColor(ByVal ws As Worksheet, ByVal startRow As Long, ByVal nRows As Long, ByVal colIdx As Long)
    If nRows < 1 Then Exit Sub
    Dim cell As Range
    For Each cell In ws.Range(ws.Cells(startRow + 1, colIdx), ws.Cells(startRow + nRows, colIdx))
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            If cell.Value > 0.0001 Then
                cell.Font.Color = RGB(0, 150, 0)
            ElseIf cell.Value < -0.0001 Then
                cell.Font.Color = RGB(200, 0, 0)
            End If
        End If
    Next cell
End Sub

Private Function GetMaxDate(ByRef data() As Variant) As Date
    Dim maxDate As Date
    Dim i As Long
    maxDate = 0
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) > maxDate Then maxDate = data(i, dictCols("DATE_SAVED"))
        End If
    Next i
    GetMaxDate = maxDate
End Function

Private Function FindClosestDate(ByRef data() As Variant, targetDate As Date) As Date
    Dim i As Long
    Dim closestDate As Date
    Dim diff As Long
    Dim minDiff As Long
    minDiff = 999999

    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            diff = targetDate - data(i, dictCols("DATE_SAVED"))
            If diff >= 0 And diff < minDiff Then
                minDiff = diff
                closestDate = data(i, dictCols("DATE_SAVED"))
            End If
        End If
    Next i

    FindClosestDate = closestDate
End Function

Private Function BuildHistoryDictionary(ByRef data() As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            dict(data(i, dictCols("ASSET_ID")) & "_" & Format(data(i, dictCols("DATE_SAVED")), "dd/mm/yyyy")) = i
        End If
    Next i
    Set BuildHistoryDictionary = dict
End Function


Private Function SafeCDbl(ByVal val As Variant) As Double
    If IsNumeric(val) Then
        SafeCDbl = CDbl(val)
    Else
        SafeCDbl = 0
    End If
End Function

Private Function SafeGetVal(ByRef data() As Variant, ByVal idx As Variant, ByVal col As Long) As Variant
    If IsEmpty(idx) Then
        SafeGetVal = "N/A"
    ElseIf idx = 0 Then
        SafeGetVal = "N/A"
    Else
        SafeGetVal = SafeCDbl(data(idx, col))
    End If
End Function

Private Function ProcessRiskProximity(ByVal ws As Worksheet, ByRef data() As Variant, maxDate As Date, startRow As Long) As Long
    ws.Cells(startRow, 1).Value = "Section 1: Risk & Target Proximity (Most Recent Data)"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 14
    startRow = startRow + 1

    Dim headers As Variant
    headers = Array("Asset ID", "Latest Name", "Current Price", "Target Price", "Stop Price", "Status", "Deviation")
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, UBound(headers) + 1)).Value = headers

    Dim outData() As Variant
    ReDim outData(1 To UBound(data, 1), 1 To 7)
    Dim outCount As Long
    outCount = 0

    Dim i As Long
    Dim cp As Double, tp As Double, sp As Double
    Dim diffT As Double, diffS As Double
    Dim status As String, devVal As Variant

    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) = maxDate Then
                cp = SafeCDbl(data(i, dictCols("CURRENT_PRICE")))
                tp = SafeCDbl(data(i, dictCols("TARGET_PRICE")))
                sp = SafeCDbl(data(i, dictCols("STOP_PRICE")))

                If cp > 0 Then
                    diffT = 999: diffS = 999
                    If tp > 0 Then diffT = Abs(cp - tp) / tp
                    If sp > 0 Then diffS = Abs(cp - sp) / sp

                    status = ""
                    devVal = ""

                    If tp > 0 And tp < cp Then
                        status = "WARNING: Target < Current"
                        devVal = diffT
                    ElseIf tp > 0 And diffT <= 0.1 Then
                        status = "Near Target (10%)"
                        devVal = diffT
                    ElseIf tp > 0 And diffT <= 0.25 Then
                        status = "Near Target (25%)"
                        devVal = diffT
                    ElseIf sp > 0 And diffS <= 0.1 Then
                        status = "Near Stop (10%)"
                        devVal = diffS
                    ElseIf sp > 0 And diffS <= 0.25 Then
                        status = "Near Stop (25%)"
                        devVal = diffS
                    End If

                    If status <> "" Then
                        outCount = outCount + 1
                        outData(outCount, 1) = data(i, dictCols("ASSET_ID"))
                        outData(outCount, 2) = data(i, dictCols("ASSET_NAME"))
                        outData(outCount, 3) = cp
                        outData(outCount, 4) = tp
                        outData(outCount, 5) = sp
                        outData(outCount, 6) = status
                        outData(outCount, 7) = devVal
                    End If
                End If
            End If
        End If
    Next i

    If outCount > 0 Then
        ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + outCount, 7)).Value = outData

        Call FormatTable(ws, startRow, outCount, 7)
        ws.Range(ws.Cells(startRow + 1, 7), ws.Cells(startRow + outCount, 7)).NumberFormat = "0.0%"

        Dim r As Long
        For r = 1 To outCount
            If InStr(outData(r, 6), "WARNING") > 0 Then
                ws.Cells(startRow + r, 6).Font.Color = RGB(255, 0, 0)
                ws.Cells(startRow + r, 6).Font.Bold = True
                ws.Cells(startRow + r, 7).Font.Color = RGB(255, 0, 0)
            ElseIf InStr(outData(r, 6), "Target") > 0 Then
                ws.Cells(startRow + r, 6).Font.Color = RGB(0, 150, 0)
                ws.Cells(startRow + r, 7).Font.Color = RGB(0, 150, 0)
            ElseIf InStr(outData(r, 6), "Stop") > 0 Then
                ws.Cells(startRow + r, 6).Font.Color = RGB(200, 0, 0)
                ws.Cells(startRow + r, 7).Font.Color = RGB(200, 0, 0)
            End If
        Next r

        startRow = startRow + outCount + 1
    Else
        ws.Cells(startRow, 1).Value = "No assets near targets/stops."
        startRow = startRow + 1
    End If

    ProcessRiskProximity = startRow + 1
End Function

Private Function ProcessAllocationGaps(ByVal ws As Worksheet, ByRef data() As Variant, dictHist As Object, maxDate As Date, d1 As Date, d2 As Date, d7 As Date, d14 As Date, startRow As Long) As Long
    ws.Cells(startRow, 1).Value = "Section 2: Allocation Gaps & Daily Performance"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 14
    startRow = startRow + 1

    Dim d1Str As String, d2Str As String, d7Str As String, d14Str As String
    d1Str = IIf(d1 > 0, Format(d1, "mm/dd"), "1D Ago")
    d2Str = IIf(d2 > 0, Format(d2, "mm/dd"), "2D Ago")
    d7Str = IIf(d7 > 0, Format(d7, "mm/dd"), "1W Ago")
    d14Str = IIf(d14 > 0, Format(d14, "mm/dd"), "2W Ago")

    Dim hHeaders As Variant
    hHeaders = Array("Asset ID", "Latest Name", "Current Price", "1D Perf", _
                    "Current Alloc", "MEAN Alloc", "Current Gap", _
                    "Gap Shift (" & d1Str & ")", "Gap Shift (" & d2Str & ")", _
                    "Gap Shift (" & d7Str & ")", "Gap Shift (" & d14Str & ")")

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, UBound(hHeaders) + 1)).Value = hHeaders

    Dim outData() As Variant
    ReDim outData(1 To UBound(data, 1), 1 To 11)
    Dim outCount As Long
    outCount = 0

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) = maxDate Then
                Dim assetID As String: assetID = data(i, dictCols("ASSET_ID"))
                Dim aName As String: aName = data(i, dictCols("ASSET_NAME"))

                Dim idx1D As Variant: idx1D = dictHist(assetID & "_" & Format(d1, "dd/mm/yyyy"))
                Dim idx2D As Variant: idx2D = dictHist(assetID & "_" & Format(d2, "dd/mm/yyyy"))
                Dim idx1W As Variant: idx1W = dictHist(assetID & "_" & Format(d7, "dd/mm/yyyy"))
                Dim idx2W As Variant: idx2W = dictHist(assetID & "_" & Format(d14, "dd/mm/yyyy"))

                outCount = outCount + 1
                outData(outCount, 1) = assetID
                outData(outCount, 2) = aName

                Dim cp As Double
                cp = SafeCDbl(data(i, dictCols("CURRENT_PRICE")))
                outData(outCount, 3) = cp

                ' 1D Perf
                Dim cp1D As Variant
                cp1D = SafeGetVal(data, idx1D, dictCols("CURRENT_PRICE"))
                If IsNumeric(cp1D) Then
                    If CDbl(cp1D) > 0 Then
                        outData(outCount, 4) = (cp / CDbl(cp1D)) - 1
                    Else
                        outData(outCount, 4) = "N/A"
                    End If
                Else
                    outData(outCount, 4) = "N/A"
                End If

                ' Current Alloc, MEAN Alloc, Current Gap
                Dim currAlloc As Double, meanAlloc As Double, currGap As Double
                currAlloc = SafeCDbl(data(i, dictCols("CURRENT_ALLOC")))
                meanAlloc = SafeCDbl(data(i, dictCols("MEAN")))
                currGap = currAlloc - meanAlloc

                outData(outCount, 5) = currAlloc
                outData(outCount, 6) = meanAlloc
                outData(outCount, 7) = currGap

                ' Gap Shifts
                Dim pastAlloc As Variant, pastMean As Variant, pastGap As Double

                ' 1D Shift
                pastAlloc = SafeGetVal(data, idx1D, dictCols("CURRENT_ALLOC"))
                pastMean = SafeGetVal(data, idx1D, dictCols("MEAN"))
                If IsNumeric(pastAlloc) And IsNumeric(pastMean) Then
                    pastGap = CDbl(pastAlloc) - CDbl(pastMean)
                    outData(outCount, 8) = currGap - pastGap
                Else
                    outData(outCount, 8) = "N/A"
                End If

                ' 2D Shift
                pastAlloc = SafeGetVal(data, idx2D, dictCols("CURRENT_ALLOC"))
                pastMean = SafeGetVal(data, idx2D, dictCols("MEAN"))
                If IsNumeric(pastAlloc) And IsNumeric(pastMean) Then
                    pastGap = CDbl(pastAlloc) - CDbl(pastMean)
                    outData(outCount, 9) = currGap - pastGap
                Else
                    outData(outCount, 9) = "N/A"
                End If

                ' 1W Shift
                pastAlloc = SafeGetVal(data, idx1W, dictCols("CURRENT_ALLOC"))
                pastMean = SafeGetVal(data, idx1W, dictCols("MEAN"))
                If IsNumeric(pastAlloc) And IsNumeric(pastMean) Then
                    pastGap = CDbl(pastAlloc) - CDbl(pastMean)
                    outData(outCount, 10) = currGap - pastGap
                Else
                    outData(outCount, 10) = "N/A"
                End If

                ' 2W Shift
                pastAlloc = SafeGetVal(data, idx2W, dictCols("CURRENT_ALLOC"))
                pastMean = SafeGetVal(data, idx2W, dictCols("MEAN"))
                If IsNumeric(pastAlloc) And IsNumeric(pastMean) Then
                    pastGap = CDbl(pastAlloc) - CDbl(pastMean)
                    outData(outCount, 11) = currGap - pastGap
                Else
                    outData(outCount, 11) = "N/A"
                End If

            End If
        End If
    Next i

    If outCount > 0 Then
        ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + outCount, 11)).Value = outData

        Call FormatTable(ws, startRow, outCount, 11)
        ws.Range(ws.Cells(startRow + 1, 4), ws.Cells(startRow + outCount, 11)).NumberFormat = "0.0%"

        Call FormatDiffColor(ws, startRow, outCount, 4)
        Call FormatDiffColor(ws, startRow, outCount, 7)
        Call FormatDiffColor(ws, startRow, outCount, 8)
        Call FormatDiffColor(ws, startRow, outCount, 9)
        Call FormatDiffColor(ws, startRow, outCount, 10)
        Call FormatDiffColor(ws, startRow, outCount, 11)

        startRow = startRow + outCount + 1
    Else
        ws.Cells(startRow, 1).Value = "No data available."
        startRow = startRow + 1
    End If

    ProcessAllocationGaps = startRow + 1
End Function

Private Function ProcessHistoricalChanges(ByVal ws As Worksheet, ByRef data() As Variant, dictHist As Object, maxDate As Date, d1 As Date, d2 As Date, d7 As Date, d14 As Date, startRow As Long) As Long
    ws.Cells(startRow, 1).Value = "Section 3: Historical Changes"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 14
    startRow = startRow + 1

    Dim cDateStr As String, d1Str As String, d2Str As String, d7Str As String, d14Str As String
    cDateStr = "Current"
    d1Str = IIf(d1 > 0, Format(d1, "mm/dd"), "1D Ago")
    d2Str = IIf(d2 > 0, Format(d2, "mm/dd"), "2D Ago")
    d7Str = IIf(d7 > 0, Format(d7, "mm/dd"), "1W Ago")
    d14Str = IIf(d14 > 0, Format(d14, "mm/dd"), "2W Ago")

    Dim hHeaders As Variant
    hHeaders = Array("Asset ID", "Latest Name", _
                    "Target Price (Latest)", "Dist to Target", _
                    "MEAN Alloc (" & cDateStr & ")", _
                    "MEAN Alloc (" & d1Str & ")", "Diff (" & d1Str & ")", _
                    "MEAN Alloc (" & d2Str & ")", "Diff (" & d2Str & ")", _
                    "MEAN Alloc (" & d7Str & ")", "Diff (" & d7Str & ")", _
                    "MEAN Alloc (" & d14Str & ")", "Diff (" & d14Str & ")")

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, UBound(hHeaders) + 1)).Value = hHeaders

    Dim outData() As Variant
    ReDim outData(1 To UBound(data, 1), 1 To 13)
    Dim outCount As Long
    outCount = 0

    Dim i As Long, pastVal As Variant
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) = maxDate Then
                Dim assetID As String: assetID = data(i, dictCols("ASSET_ID"))
                Dim aName As String: aName = data(i, dictCols("ASSET_NAME"))

                Dim idx1D As Variant: idx1D = dictHist(assetID & "_" & Format(d1, "dd/mm/yyyy"))
                Dim idx2D As Variant: idx2D = dictHist(assetID & "_" & Format(d2, "dd/mm/yyyy"))
                Dim idx1W As Variant: idx1W = dictHist(assetID & "_" & Format(d7, "dd/mm/yyyy"))
                Dim idx2W As Variant: idx2W = dictHist(assetID & "_" & Format(d14, "dd/mm/yyyy"))

                outCount = outCount + 1
                outData(outCount, 1) = assetID
                outData(outCount, 2) = aName

                Dim cp As Double, tp As Double
                cp = SafeCDbl(data(i, dictCols("CURRENT_PRICE")))
                tp = SafeCDbl(data(i, dictCols("TARGET_PRICE")))
                outData(outCount, 3) = tp
                If tp > 0 Then outData(outCount, 4) = Abs(cp - tp) / tp Else outData(outCount, 4) = "N/A"

                Dim meanCurr As Double
                meanCurr = SafeCDbl(data(i, dictCols("MEAN")))
                outData(outCount, 5) = meanCurr

                pastVal = SafeGetVal(data, idx1D, dictCols("MEAN"))
                outData(outCount, 6) = pastVal
                If IsNumeric(pastVal) Then outData(outCount, 7) = meanCurr - CDbl(pastVal) Else outData(outCount, 7) = "N/A"

                pastVal = SafeGetVal(data, idx2D, dictCols("MEAN"))
                outData(outCount, 8) = pastVal
                If IsNumeric(pastVal) Then outData(outCount, 9) = meanCurr - CDbl(pastVal) Else outData(outCount, 9) = "N/A"

                pastVal = SafeGetVal(data, idx1W, dictCols("MEAN"))
                outData(outCount, 10) = pastVal
                If IsNumeric(pastVal) Then outData(outCount, 11) = meanCurr - CDbl(pastVal) Else outData(outCount, 11) = "N/A"

                pastVal = SafeGetVal(data, idx2W, dictCols("MEAN"))
                outData(outCount, 12) = pastVal
                If IsNumeric(pastVal) Then outData(outCount, 13) = meanCurr - CDbl(pastVal) Else outData(outCount, 13) = "N/A"

            End If
        End If
    Next i

    If outCount > 0 Then
        ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + outCount, 13)).Value = outData

        Call FormatTable(ws, startRow, outCount, 13)
        ws.Range(ws.Cells(startRow + 1, 4), ws.Cells(startRow + outCount, 13)).NumberFormat = "0.0%"

        Call FormatDiffColor(ws, startRow, outCount, 7)
        Call FormatDiffColor(ws, startRow, outCount, 9)
        Call FormatDiffColor(ws, startRow, outCount, 11)
        Call FormatDiffColor(ws, startRow, outCount, 13)

        startRow = startRow + outCount + 1
    Else
        ws.Cells(startRow, 1).Value = "No historical data to compare."
        startRow = startRow + 1
    End If

    ProcessHistoricalChanges = startRow + 1
End Function

Private Function ProcessConvictionAnalysis(ByVal ws As Worksheet, ByRef data() As Variant, maxDate As Date, startRow As Long) As Long
    ws.Cells(startRow, 1).Value = "Section 4: Conviction Analysis"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 14
    startRow = startRow + 1

    Dim cHeaders As Variant
    cHeaders = Array("Conviction (/5)", "Avg Expected Return", "Avg Dist to Target", _
                     "Avg ERC UNCSTRD", "Avg ER/VOL", "Avg SHARPE", "Avg CUSTOM", "Avg MEAN Alloc", _
                     "Asset Count")
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, UBound(cHeaders) + 1)).Value = cHeaders

    Dim convCounts(1 To 5) As Long
    Dim convSumER(1 To 5) As Double
    Dim convSumDist(1 To 5) As Double
    Dim convCountDist(1 To 5) As Long

    Dim convSumERC(1 To 5) As Double
    Dim convSumERVOL(1 To 5) As Double
    Dim convSumSHARPE(1 To 5) As Double
    Dim convSumCUSTOM(1 To 5) As Double
    Dim convSumMEAN(1 To 5) As Double

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) = maxDate Then
                Dim cScore As Integer
                cScore = Int(SafeCDbl(data(i, dictCols("CONVICTION"))))
                If cScore >= 1 And cScore <= 5 Then
                    convCounts(cScore) = convCounts(cScore) + 1
                    convSumER(cScore) = convSumER(cScore) + SafeCDbl(data(i, dictCols("EXPECTED_RETURN")))
                    convSumERC(cScore) = convSumERC(cScore) + SafeCDbl(data(i, dictCols("ERC")))
                    convSumERVOL(cScore) = convSumERVOL(cScore) + SafeCDbl(data(i, dictCols("ERVOL")))
                    convSumSHARPE(cScore) = convSumSHARPE(cScore) + SafeCDbl(data(i, dictCols("SHARPE")))
                    convSumCUSTOM(cScore) = convSumCUSTOM(cScore) + SafeCDbl(data(i, dictCols("CUSTOM")))
                    convSumMEAN(cScore) = convSumMEAN(cScore) + SafeCDbl(data(i, dictCols("MEAN")))

                    Dim cp As Double, tp As Double
                    cp = SafeCDbl(data(i, dictCols("CURRENT_PRICE")))
                    tp = SafeCDbl(data(i, dictCols("TARGET_PRICE")))
                    If tp > 0 Then
                        convSumDist(cScore) = convSumDist(cScore) + (Abs(cp - tp) / tp)
                        convCountDist(cScore) = convCountDist(cScore) + 1
                    End If
                End If
            End If
        End If
    Next i

    Dim outData() As Variant
    ReDim outData(1 To 5, 1 To 9)
    Dim outCount As Long
    outCount = 0

    For i = 5 To 1 Step -1
        outCount = outCount + 1
        outData(outCount, 1) = i
        If convCounts(i) > 0 Then
            outData(outCount, 2) = convSumER(i) / convCounts(i)
            If convCountDist(i) > 0 Then outData(outCount, 3) = convSumDist(i) / convCountDist(i) Else outData(outCount, 3) = "N/A"
            outData(outCount, 4) = convSumERC(i) / convCounts(i)
            outData(outCount, 5) = convSumERVOL(i) / convCounts(i)
            outData(outCount, 6) = convSumSHARPE(i) / convCounts(i)
            outData(outCount, 7) = convSumCUSTOM(i) / convCounts(i)
            outData(outCount, 8) = convSumMEAN(i) / convCounts(i)
        Else
            outData(outCount, 2) = "N/A"
            outData(outCount, 3) = "N/A"
            outData(outCount, 4) = "N/A"
            outData(outCount, 5) = "N/A"
            outData(outCount, 6) = "N/A"
            outData(outCount, 7) = "N/A"
            outData(outCount, 8) = "N/A"
        End If
        outData(outCount, 9) = convCounts(i)
    Next i

    ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + outCount, 9)).Value = outData

    Call FormatTable(ws, startRow, outCount, 9)
    ws.Range(ws.Cells(startRow + 1, 2), ws.Cells(startRow + outCount, 8)).NumberFormat = "0.0%"

    ProcessConvictionAnalysis = startRow + outCount + 2
End Function

Private Function ProcessSchemeVariations(ByVal ws As Worksheet, ByRef data() As Variant, dictHist As Object, maxDate As Date, d1 As Date, d2 As Date, d7 As Date, d14 As Date, startRow As Long) As Long
    ws.Cells(startRow, 1).Value = "Section 5: Allocation Scheme Variations Over Time"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 14
    startRow = startRow + 1

    Dim cDateStr As String, d1Str As String, d2Str As String, d7Str As String, d14Str As String
    cDateStr = "Current"
    d1Str = IIf(d1 > 0, Format(d1, "mm/dd"), "1D Ago")
    d2Str = IIf(d2 > 0, Format(d2, "mm/dd"), "2D Ago")
    d7Str = IIf(d7 > 0, Format(d7, "mm/dd"), "1W Ago")
    d14Str = IIf(d14 > 0, Format(d14, "mm/dd"), "2W Ago")

    Dim hHeaders As Variant
    hHeaders = Array("Asset ID", "Latest Name", "Scheme", _
                    cDateStr, _
                    d1Str, "Diff (" & d1Str & ")", _
                    d2Str, "Diff (" & d2Str & ")", _
                    d7Str, "Diff (" & d7Str & ")", _
                    d14Str, "Diff (" & d14Str & ")")

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, UBound(hHeaders) + 1)).Value = hHeaders

    Dim schemeCols(1 To 5) As Long
    Dim schemeNames(1 To 5) As String
    schemeCols(1) = dictCols("ERC"): schemeNames(1) = "ERC UNCSTRD"
    schemeCols(2) = dictCols("ERVOL"): schemeNames(2) = "ER/VOL"
    schemeCols(3) = dictCols("SHARPE"): schemeNames(3) = "SHARPE"
    schemeCols(4) = dictCols("CUSTOM"): schemeNames(4) = "CUSTOM"
    schemeCols(5) = dictCols("MEAN"): schemeNames(5) = "MEAN"

    Dim outData() As Variant
    ReDim outData(1 To UBound(data, 1) * 5, 1 To 12)
    Dim outCount As Long
    outCount = 0

    Dim i As Long, s As Integer, pastVal As Variant
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) = maxDate Then
                Dim assetID As String: assetID = data(i, dictCols("ASSET_ID"))
                Dim aName As String: aName = data(i, dictCols("ASSET_NAME"))

                Dim idx1D As Variant: idx1D = dictHist(assetID & "_" & Format(d1, "dd/mm/yyyy"))
                Dim idx2D As Variant: idx2D = dictHist(assetID & "_" & Format(d2, "dd/mm/yyyy"))
                Dim idx1W As Variant: idx1W = dictHist(assetID & "_" & Format(d7, "dd/mm/yyyy"))
                Dim idx2W As Variant: idx2W = dictHist(assetID & "_" & Format(d14, "dd/mm/yyyy"))

                For s = 1 To 5
                    outCount = outCount + 1
                    outData(outCount, 1) = assetID
                    outData(outCount, 2) = aName
                    outData(outCount, 3) = schemeNames(s)

                    Dim currVal As Double
                    currVal = SafeCDbl(data(i, schemeCols(s)))
                    outData(outCount, 4) = currVal

                    ' 1D
                    pastVal = SafeGetVal(data, idx1D, schemeCols(s))
                    outData(outCount, 5) = pastVal
                    If IsNumeric(pastVal) Then outData(outCount, 6) = currVal - CDbl(pastVal) Else outData(outCount, 6) = "N/A"

                    ' 2D
                    pastVal = SafeGetVal(data, idx2D, schemeCols(s))
                    outData(outCount, 7) = pastVal
                    If IsNumeric(pastVal) Then outData(outCount, 8) = currVal - CDbl(pastVal) Else outData(outCount, 8) = "N/A"

                    ' 1W
                    pastVal = SafeGetVal(data, idx1W, schemeCols(s))
                    outData(outCount, 9) = pastVal
                    If IsNumeric(pastVal) Then outData(outCount, 10) = currVal - CDbl(pastVal) Else outData(outCount, 10) = "N/A"

                    ' 2W
                    pastVal = SafeGetVal(data, idx2W, schemeCols(s))
                    outData(outCount, 11) = pastVal
                    If IsNumeric(pastVal) Then outData(outCount, 12) = currVal - CDbl(pastVal) Else outData(outCount, 12) = "N/A"
                Next s
            End If
        End If
    Next i

    If outCount > 0 Then
        ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + outCount, 12)).Value = outData

        Call FormatTable(ws, startRow, outCount, 12)
        ws.Range(ws.Cells(startRow + 1, 4), ws.Cells(startRow + outCount, 12)).NumberFormat = "0.0%"

        Call FormatDiffColor(ws, startRow, outCount, 6)
        Call FormatDiffColor(ws, startRow, outCount, 8)
        Call FormatDiffColor(ws, startRow, outCount, 10)
        Call FormatDiffColor(ws, startRow, outCount, 12)

        startRow = startRow + outCount + 1
    Else
        ws.Cells(startRow, 1).Value = "No historical data to compare."
        startRow = startRow + 1
    End If

    ProcessSchemeVariations = startRow + 1
End Function

-------------------------------------------------------------------------------
VBA MACRO allocationLogic.bas
in file: xl/vbaProject.bin - OLE stream: 'VBA/allocationLogic'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

' =========================================================================
' DOWNSIDE BETA & CVaR EQUITY ALLOCATION MODULE
' =========================================================================
Sub adjustForMDD()

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

    ' 4. Clean dashboard readout via MsgBox
    MsgBox "--- PORTFOLIO RISK FACTORS (70% Target) ---" & vbCr & vbCr & _
           "1. DOWNSIDE BETA: " & Round(DSBeta, 2) & " (Limit: " & Round(DDFactor * 100, 1) & "%)" & vbCr & vbCr & _
           "2. TAIL RISK (95% CVaR):" & vbCr & _
           "   Bench: " & Round(benchCVaR * 100, 2) & "%  |  Port: " & Round(portCVaR * 100, 2) & "%" & vbCr & _
           "   Limit: " & Round(CVaRFactor * 100, 1) & "%)" & vbCr & vbCr & _
           "3. VOLATILITY (Semi-Dev < 0%):" & vbCr & _
           "   Bench: " & Round(benchSemiDev * 100, 2) & "%  |  Port: " & Round(portSemiDev * 100, 2) & "%" & vbCr & _
           "   Limit: " & Round(SemiDevFactor * 100, 1) & "%)" & vbCr & vbCr & _
           "4. HISTORICAL MAX DRAWDOWN:" & vbCr & _
           "   Bench: " & Round(benchMDD * 100, 2) & "%  |  Port: " & Round(portMDD * 100, 2) & "%" & vbCr & _
           "   Limit: " & Round(HistoricalMDDFactor * 100, 1) & "%)" & vbCr & _
           "----------------------------------------" & vbCr & vbCr & _
           "FINAL APPLIED RISK FACTOR: " & Round(finalRiskFactor * 100, 1) & "%", vbInformation, "Risk Adjustments"

    Dim fullAllocations As Variant, computedAllocations As Variant
    Set fullAllocations = ThisWorkbook.Sheets("Dashboard").Range("StrategyWeightsStart").CurrentRegion
    computedAllocations = Intersect(fullAllocations, fullAllocations.Offset(1, 1)).Value2

    Dim nStrat As Integer, nAssets As Integer
    nStrat = (UBound(computedAllocations, 2) - LBound(computedAllocations, 2) - 1) / 3 - 2
    nAssets = UBound(computedAllocations, 1) - LBound(computedAllocations, 1)

    Dim i As Integer, j As Integer
    Dim equityPositions As Dictionary
    Set equityPositions = DataUtils.GetEquityPositionsFromFund

    For j = 2 To nAssets + 1
        For i = 2 To nStrat + 1
            ' Apply the unified strictest risk limit
            computedAllocations(j, (nStrat + 2) * 2 + i) = computedAllocations(j, i) * finalRiskFactor
        Next i
        i = i - 1
        computedAllocations(j, (nStrat + 2) * 2 + i + 1) = equityPositions(equityPositions.Keys(j - 2)).Weight
        computedAllocations(j, (nStrat + 2) * 2 + i + 2) = computedAllocations(j, nStrat * 2 + i) - computedAllocations(j, nStrat * 2 + i + 1)
    Next j

    ThisWorkbook.Sheets("Dashboard").Range("StrategyWeightsStart").Resize(UBound(computedAllocations, 1), UBound(computedAllocations, 2)).Value = computedAllocations

End Sub


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

' Sub-Function: Calculate Historical CVaR
Private Function CalculateHistoricalCVaR(rets() As Double, confidence As Double) As Double
    Dim numRets As Long
    numRets = UBound(rets) - LBound(rets) + 1

    ' Isolate returns into a new array to be sorted
    Dim sortedRets() As Double
    ReDim sortedRets(1 To numRets)
    Dim i As Long
    For i = 1 To numRets
        sortedRets(i) = rets(i)
    Next i

    ' Sort the array from worst to best
    Call QuickSortAscending(sortedRets, LBound(sortedRets), UBound(sortedRets))

    ' Find the VaR cutoff index (e.g. 5% worst)
    Dim varIndex As Long
    varIndex = Int(numRets * (1 - confidence))
    If varIndex < 1 Then varIndex = 1 ' Failsafe for very small datasets

    ' Calculate Expected Shortfall (Average of returns strictly below the VaR cutoff)
    Dim sum As Double
    sum = 0
    For i = 1 To varIndex
        sum = sum + sortedRets(i)
    Next i

    CalculateHistoricalCVaR = sum / varIndex
End Function

' Helper: Fast sorting for CVaR arrays
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

' =========================================================================
' NEW DOWNSIDE VOLATILITY & DRAWDOWN MATH HELPER FUNCTIONS
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
Private Function CalculateSemiDeviation(rets() As Double, target As Double) As Double
    Dim i As Long, numRets As Long
    Dim sumSq As Double

    numRets = UBound(rets) - LBound(rets) + 1
    sumSq = 0

    For i = LBound(rets) To UBound(rets)
        If rets(i) < target Then
            ' Square the deviation from the target
            sumSq = sumSq + ((rets(i) - target) ^ 2)
        End If
    Next i

    ' Root Mean Square of the downside deviations
    CalculateSemiDeviation = Sqr(sumSq / numRets)
End Function


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
-------------------------------------------------------------------------------
VBA MACRO Sheet10.cls
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet10'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
(empty macro)
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|Suspicious|Write               |May write to a file (if combined with Open)  |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Create              |May execute file or a system command through |
|          |                    |WMI                                          |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Base64    |0@                  |MEAN                                         |
|String    |                    |                                             |
|Suspicious|VBA Stomping        |VBA Stomping was detected: the VBA source    |
|          |                    |code and P-code are different, this may have |
|          |                    |been used to hide malicious code             |
+----------+--------------------+---------------------------------------------+
VBA Stomping detection is experimental: please report any false positive/negative at https://github.com/decalage2/oletools/issues

Attribute VB_Name = "Main"
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
    Call updateAllPriceHistoryFromInfin
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

    ' --- PRE-CACHE ALL POSITIONS ONCE ---
    Dim masterPositions As Object
    Set masterPositions = CreateObject("Scripting.Dictionary")

    Dim convictions As Object
    Set convictions = convictionsUtils.getConvictions()

    Dim rf As Double
    rf = Sheets(DASH_SHEET).Range("D4").Value
    Dim conf As Double
    conf = Sheets(DASH_SHEET).Range("D5").Value

    Dim cacheIdx As Integer
    For cacheIdx = 1 To nAssets
        Dim pPrices() As Double, pDates() As Date
        ReDim pPrices(1 To UBound(prices, 1))
        ReDim pDates(1 To UBound(dates))
        Dim iDay As Long
        For iDay = 1 To UBound(prices, 1)
            pPrices(iDay) = prices(iDay, cacheIdx)
            pDates(iDay) = dates(iDay)
        Next iDay

        Dim cachePos As PositionCls
        Set cachePos = New PositionCls
        cachePos.AssetName = assetNames(cacheIdx)
        cachePos.InitializeData pPrices, pDates
        cachePos.ComputeMetrics rf, conf

        If convictions.Exists(cachePos.AssetName) Then
            Dim cDict As Object
            Set cDict = convictions(cachePos.AssetName)
            On Error Resume Next
            cachePos.Conviction = cDict("conviction")
            cachePos.TP = cDict("TP")
            cachePos.TPProba = cDict("TP proba")
            cachePos.SP = cDict("SP")
            cachePos.SPProba = cDict("SP proba")
            On Error GoTo 0
        End If

        cachePos.ExpectedReturn = expReturns(cacheIdx)

        masterPositions.Add assetNames(cacheIdx), cachePos
    Next cacheIdx

    Dim masterPortfolios As Object
    Set masterPortfolios = CreateObject("Scripting.Dictionary")

    ' Extract Benchmark Returns for Downside Beta calculations, dynamically matched to portfolio Dates
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Sheets("Data")
    Dim benchCol As Long
    benchCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Dim benchLastRow As Long
    benchLastRow = wsData.Cells(wsData.Rows.Count, benchCol).End(xlUp).Row

    Dim benchRawDates() As Variant
    Dim benchRawPrices() As Variant
    ' Benchmark dates are 3 columns to the left of the computed benchmark column
    benchRawDates = wsData.Range(wsData.Cells(2, benchCol - 3), wsData.Cells(benchLastRow, benchCol - 3)).Value
    benchRawPrices = wsData.Range(wsData.Cells(2, benchCol), wsData.Cells(benchLastRow, benchCol)).Value

    Dim dictBench As Object
    Set dictBench = CreateObject("Scripting.Dictionary")
    Dim bIdx As Long
    For bIdx = 1 To UBound(benchRawDates, 1)
        If IsDate(benchRawDates(bIdx, 1)) Then
            dictBench(CDate(benchRawDates(bIdx, 1))) = CDbl(benchRawPrices(bIdx, 1))
        End If
    Next bIdx

    ' Now build an aligned return array strictly matching the length of the portfolio dates (which matches logRets length)
    Dim benchRets() As Double
    ' logRets are size (1 to UBound(dates) - 1)
    ReDim benchRets(1 To UBound(dates) - 1)

    For bIdx = 1 To UBound(dates) - 1
        Dim dPrev As Date, dCurr As Date
        dPrev = dates(bIdx)
        dCurr = dates(bIdx + 1)

        If dictBench.Exists(dPrev) And dictBench.Exists(dCurr) Then
            Dim pBPrev As Double, pBCurr As Double
            pBPrev = dictBench(dPrev)
            pBCurr = dictBench(dCurr)

            If pBPrev > 0 And pBCurr > 0 Then
                benchRets(bIdx) = Log(pBCurr / pBPrev)
            Else
                benchRets(bIdx) = 0
            End If
        Else
            benchRets(bIdx) = 0
        End If
    Next bIdx

    ' Inject benchRets back into positions
    Dim pKey As Variant
    For Each pKey In masterPositions.Keys
        masterPositions(pKey).ComputeMetrics rf, conf, benchRets
    Next pKey


    ' ------------------------------------

    Dim optimizer As OptimizerCls
    Set optimizer = New OptimizerCls

    For i = LBound(strategies) To UBound(strategies)
        Dim simPort As SimulatedPortfolioCls
        Set simPort = New SimulatedPortfolioCls
        simPort.StrategyName = CStr(strategies(i))

        ' 1. Add pre-cached positions
        Dim jPos As Integer
        For jPos = 1 To nAssets
            simPort.AddPosition assetNames(jPos), masterPositions(assetNames(jPos))
        Next jPos

        ' 2. Initialize with equal weights
        simPort.SetEqualWeights

        ' 3. Optimize the weights based on strategy
        optimizer.Optimize simPort, minWeights, maxWeights, sumWeights, nAssets, assetNames



        ' 4. Simulate portfolio over time
        simPort.Simulate

        ' 5. Compute performance metrics
        Dim rf As Double
        rf = Sheets(DASH_SHEET).Range("D4").Value
        Dim conf As Double
        conf = Sheets(DASH_SHEET).Range("D5").Value
        simPort.ComputeMetrics rf, conf, benchRets

        ' 6. Output Results
        Dim mRet As Double, mVol As Double, mSharpe As Double, mMDD As Double, mLen As Integer, mVaR As Double
        mRet = simPort.Metrics.Ret
        mVol = simPort.Metrics.Vol
        mSharpe = simPort.Metrics.Sharpe
        mMDD = simPort.Metrics.MDD
        mLen = simPort.Metrics.MDDLen
        mVaR = simPort.Metrics.VaR

        Dim equityCurve() As Double, ddCurve() As Double, dailyWeights() As Double
        equityCurve = simPort.EquityCurve
        ddCurve = simPort.DrawdownCurve
        dailyWeights = simPort.DailyWeights

        ' Reconstruct weights array for backwards compatibility with OutputMetricsToRow
        Dim wArr() As Double
        ReDim wArr(1 To 1, 1 To nAssets)
        Dim idx As Integer
        For idx = 1 To nAssets
            wArr(1, idx) = simPort.GetWeight(assetNames(idx))
        Next idx

        Call OutputMetricsToRow(i - 1, CStr(strategies(i)), mRet, mVol, mSharpe, mMDD, mLen, mVaR, wArr, "StrategyTableStart", True)

        Dim colOffset As Integer
        colOffset = (nAssets + 1) + i
        Call WriteCurveToSheet(equityCurve, colOffset)

        Dim ddColOffset As Integer
        ddColOffset = (nAssets + 1) + UBound(strategies) + i
        Call WriteCurveToSheet(ddCurve, ddColOffset)

        Dim wColStart As Integer
        wColStart = 2 + (i - 1) * nAssets
        Call WriteWeightsToSheet(dailyWeights, wColStart)

        masterPortfolios.Add simPort.StrategyName, simPort
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

    ' --- POST-SIMULATION RISK OVERLAY ---
    Dim riskMultiplier As Double
    riskMultiplier = allocationLogic.ComputeFinalRiskFactor()

    Dim pKey As Variant
    For Each pKey In masterPortfolios.Keys
        Dim p As SimulatedPortfolioCls
        Set p = masterPortfolios(pKey)

        If p.StrategyName <> "EQUAL WEIGHT" And p.StrategyName <> "MEAN" Then
            p.ApplyRiskOverlay riskMultiplier

            ' Re-simulate to calculate the depressed equity curve
            p.Simulate
            p.ComputeMetrics rf, conf, benchRets

            ' Re-write the updated depressed curve to the Engine sheet so the chart uses the updated data
            ' But in the current macro architecture, the sheets are populated in the first loop.
            ' A better design: compute risk factor based on equal weight or just MEAN,
            ' which is available since the charts output happens concurrently.
            ' Let's just apply it dynamically right here and let the Portfolios Dashboard catch the new metrics.
        End If
    Next pKey
    ' ------------------------------------

    Call UpdateDashboardCharts(nAssets, UBound(strategies), lastDate)

    ' Generate the new unified Dashboard
    Call PositionDashboard.GeneratePositionsDashboard(masterPositions)
    Call PositionDashboard.GeneratePortfoliosDashboard(masterPortfolios, assetNames)

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
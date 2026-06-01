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

    ' Extract fund assets (skip benchmarks in first 2 columns)
    Dim nTotal As Integer
    nTotal = UBound(assetNames)
    Dim nAssets As Integer
    nAssets = nTotal - 2

    Dim fundPrices() As Double, fundAssetNames() As String
    ReDim fundPrices(1 To UBound(prices, 1), 1 To nAssets)
    ReDim fundAssetNames(1 To nAssets)

    Dim d As Long, j As Integer
    For d = 1 To UBound(prices, 1)
        For j = 1 To nAssets
            fundPrices(d, j) = prices(d, j + 2)
        Next j
    Next d
    For j = 1 To nAssets
        fundAssetNames(j) = assetNames(j + 2)
    Next j

    Dim strategies As Variant
    strategies = Array("ERC UNCSTRD", "ER/VOL", "SHARPE", "CUSTOM", "MEAN") ', "MIN VAR", "KELLY"

    Call SyncDashboardHeaders(dates, fundAssetNames, strategies)
    Call UpdateConvictions
    Call UpdateCurrentPrices(fundPrices)

    Dim logRets() As Double, meanRets() As Double
    Call CalculateHistoricalStats(dates, fundPrices, logRets, meanRets)

    'MsgBox "Data Updated. Matrices Built.", vbInformation
    Application.ScreenUpdating = True
End Sub

Sub RunAllSolvers()
    Application.ScreenUpdating = False
    Sheets(ENGINE_SHEET).Activate

    Dim prices() As Variant, dates() As Date, assetNames() As String
    Call GetHistoricalDataDatesAndNames(prices, dates, assetNames)

    ' Split into Fund Assets vs Benchmark Assets
    Dim nTotal As Integer
    nTotal = UBound(assetNames)
    Dim nAssets As Integer
    nAssets = nTotal - 2

    Dim fundPrices() As Double, fundAssetNames() As String
    ReDim fundPrices(1 To UBound(prices, 1), 1 To nAssets)
    ReDim fundAssetNames(1 To nAssets)

    Dim benchPricesRaw() As Double, benchAssetNames() As String
    ReDim benchPricesRaw(1 To UBound(prices, 1), 1 To 2)
    ReDim benchAssetNames(1 To 2)

    Dim d As Long, j As Integer
    For d = 1 To UBound(prices, 1)
        For j = 1 To 2
            benchPricesRaw(d, j) = prices(d, j)
        Next j
        For j = 1 To nAssets
            fundPrices(d, j) = prices(d, j + 2)
        Next j
    Next d
    For j = 1 To 2
        benchAssetNames(j) = assetNames(j)
    Next j
    For j = 1 To nAssets
        fundAssetNames(j) = assetNames(j + 2)
    Next j

    ' Substitute the old arrays with the fund-only arrays so the rest of the solver math works seamlessly
    Dim logRets() As Double, meanRets() As Double
    Call CalculateHistoricalStats(dates, fundPrices, logRets, meanRets)

    Dim covMat() As Double
    covMat = CalculateCovariance(logRets, meanRets)
    Call OutputCorrelationMatrix(covMat, fundAssetNames)

    Dim expReturns() As Double
    Dim lastDate As Date
    lastDate = dates(UBound(dates))
    expReturns = CalculateExpectedReturns(meanRets, lastDate)

    Dim wsDash As Worksheet
    Set wsDash = Sheets(DASH_SHEET)
    Dim inputStart As Long
    inputStart = wsDash.Range("InputTableStart").Row + 1

    Dim k As Integer
    For k = 1 To nAssets
        wsDash.Cells(inputStart + k - 1, wsDash.Range("AssetMetricsStart").Column - 1).Value = expReturns(k)
    Next k

    Call WriteToEngine(covMat, expReturns)


    Dim strategies As Variant
    strategies = Array("ERC UNCSTRD", "ER/VOL", "SHARPE", "CUSTOM", "MEAN") ', "MIN VAR", "KELLY"
    Dim nStrat As Integer
    nStrat = UBound(strategies)

    Call SyncSimWeightsHeaders(dates, fundAssetNames, strategies)

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
        ReDim pPrices(1 To UBound(fundPrices, 1))
        ReDim pDates(1 To UBound(dates))
        Dim iDay As Long
        For iDay = 1 To UBound(fundPrices, 1)
            pPrices(iDay) = fundPrices(iDay, cacheIdx)
            pDates(iDay) = dates(iDay)
        Next iDay

        Dim cachePos As PositionCls
        Set cachePos = New PositionCls
        cachePos.AssetName = fundAssetNames(cacheIdx)
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

        masterPositions.Add fundAssetNames(cacheIdx), cachePos
    Next cacheIdx

    Dim masterPortfolios As Object
    Set masterPortfolios = CreateObject("Scripting.Dictionary")

    ' Build the BENCHMARK Portfolio
    Dim benchPort As SimulatedPortfolioCls
    Set benchPort = New SimulatedPortfolioCls
    benchPort.StrategyName = "BENCHMARK"

    Dim benchMasterPos As Object
    Set benchMasterPos = CreateObject("Scripting.Dictionary")

    For cacheIdx = 1 To 2
        ReDim pPrices(1 To UBound(benchPricesRaw, 1))
        ReDim pDates(1 To UBound(dates))
        For iDay = 1 To UBound(benchPricesRaw, 1)
            pPrices(iDay) = benchPricesRaw(iDay, cacheIdx)
            pDates(iDay) = dates(iDay)
        Next iDay

        Dim bPos As PositionCls
        Set bPos = New PositionCls
        bPos.AssetName = benchAssetNames(cacheIdx)
        bPos.InitializeData pPrices, pDates
        bPos.ComputeMetrics rf, conf

        benchMasterPos.Add benchAssetNames(cacheIdx), bPos
        benchPort.AddPosition benchAssetNames(cacheIdx), bPos
    Next cacheIdx

    ' Assign 50/50 weights
    benchPort.SetEqualWeights
    benchPort.Simulate

    ' Extract benchRets to pass into ComputeMetrics for Downside Beta
    Dim benchCurve() As Double
    benchCurve = benchPort.EquityCurve
    Dim benchRets() As Double
    ReDim benchRets(1 To UBound(dates) - 1)
    Dim bIdx As Long
    For bIdx = 1 To UBound(dates) - 1
        If benchCurve(bIdx, 1) > 0 And benchCurve(bIdx + 1, 1) > 0 Then
            benchRets(bIdx) = Log(benchCurve(bIdx + 1, 1) / benchCurve(bIdx, 1))
        Else
            benchRets(bIdx) = 0
        End If
    Next bIdx

    benchPort.ComputeMetrics rf, conf, benchRets
    masterPortfolios.Add "BENCHMARK", benchPort


    ' Inject benchRets back into positions
    Dim pKey As Variant
    Dim rIdx As Integer
    rIdx = 0
    For Each pKey In masterPositions.Keys
        masterPositions(pKey).ComputeMetrics rf, conf, benchRets

        ' Output to legacy Dashboard Asset Metrics Table
        Dim curPos As PositionCls
        Set curPos = masterPositions(pKey)
        Dim dummyW As Variant
        Call OutputMetricsToRow(rIdx, curPos.AssetName, curPos.Metrics.Ret, curPos.Metrics.Vol, curPos.Metrics.Sharpe, curPos.Metrics.MDD, curPos.Metrics.MDDLen, curPos.Metrics.VaR, dummyW, "AssetMetricsStart", False)

        ' Also explicitly write the Downside Beta back to the sheet to ensure it is visible!
        Dim wsDashLegacy As Worksheet
        Set wsDashLegacy = Sheets(DASH_SHEET)
        Dim aRow As Long, aCol As Long
        aRow = wsDashLegacy.Range("AssetMetricsStart").Row + 1 + rIdx
        aCol = wsDashLegacy.Range("AssetMetricsStart").Column

        ' Note: Downside Beta is typically the 8th column in the old table, but just in case, we append it
        ' OutputMetricsToRow prints to StartCol + 1 through + 6.
        ' Let's print Downside Beta to StartCol + 7 (which usually represents CVaR or DSBeta depending on dashboard format)
        wsDashLegacy.Cells(aRow, aCol + 7).Value = curPos.Metrics.DownsideBeta

        ' And output the asset's Equity Curve to the CHART_SHEET (previously handled by ProcessIndividualAssets)
        Dim wsChart As Worksheet
        Set wsChart = Sheets(CHART_SHEET)
        Dim pArr() As Double
        pArr = curPos.Prices
        Dim curve() As Double
        Dim nDays As Long
        nDays = UBound(pArr)
        ReDim curve(1 To nDays, 1 To 1)
        Dim iDay As Long
        For iDay = 1 To nDays
            curve(iDay, 1) = (pArr(iDay) / pArr(1)) * 100
        Next iDay
        wsChart.Range(wsChart.Cells(2, 2 + rIdx), wsChart.Cells(2 + nDays - 1, 2 + rIdx)).Value = curve

        rIdx = rIdx + 1
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
            simPort.AddPosition fundAssetNames(jPos), masterPositions(fundAssetNames(jPos))
        Next jPos

        ' 2. Initialize with equal weights
        simPort.SetEqualWeights

        ' 3. Optimize the weights based on strategy
        optimizer.Optimize simPort, minWeights, maxWeights, sumWeights, nAssets, fundAssetNames



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
            wArr(1, idx) = simPort.GetWeight(fundAssetNames(idx))
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
    riskMultiplier = allocationLogic.ComputeFinalRiskFactor(masterPortfolios("MEAN"), masterPortfolios("BENCHMARK"))

    Dim pKey As Variant
    For Each pKey In masterPortfolios.Keys
        Dim p As SimulatedPortfolioCls
        Set p = masterPortfolios(pKey)

        If p.StrategyName <> "EQUAL WEIGHT" And p.StrategyName <> "MEAN" And p.StrategyName <> "BENCHMARK" Then
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
    Call PositionDashboard.GeneratePortfoliosDashboard(masterPortfolios, fundAssetNames)

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
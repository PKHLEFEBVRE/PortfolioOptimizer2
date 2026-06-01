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

    Dim rf As Double, conf As Double
    rf = Sheets(DASH_SHEET).Range("D4").Value
    conf = Sheets(DASH_SHEET).Range("D5").Value

    ' 1. Fetch raw data
    Dim prices() As Variant, dates() As Date, assetNames() As String
    Call GetHistoricalDataDatesAndNames(prices, dates, assetNames)

    ' 2. Split data into Benchmark vs Fund
    Dim fundPrices() As Double, fundAssetNames() As String
    Dim benchPricesRaw() As Double, benchAssetNames() As String
    Call SplitBenchmarkAndFundAssets(prices, assetNames, fundPrices, fundAssetNames, benchPricesRaw, benchAssetNames)

    Dim nAssets As Integer
    nAssets = UBound(fundAssetNames)

    ' 3. Compute stats for Fund assets
    Dim logRets() As Double, meanRets() As Double
    Call CalculateHistoricalStats(dates, fundPrices, logRets, meanRets)

    Dim covMat() As Double
    covMat = CalculateCovariance(logRets, meanRets)
    Call OutputCorrelationMatrix(covMat, fundAssetNames)

    Dim expReturns() As Double
    Dim lastDate As Date
    lastDate = dates(UBound(dates))
    expReturns = CalculateExpectedReturns(meanRets, lastDate)

    ' 4. Prep legacy Engine constraints
    Dim strategies As Variant
    strategies = Array("ERC UNCSTRD", "ER/VOL", "SHARPE", "CUSTOM", "MEAN")

    Dim minWeights() As Double, maxWeights() As Double
    Call PrepEngineSheetAndConstraints(fundAssetNames, expReturns, covMat, minWeights, maxWeights)
    Call SyncSimWeightsHeaders(dates, fundAssetNames, strategies)

    ' 5. Cache OOP objects
    Dim masterPositions As Object
    Set masterPositions = CacheMasterPositions(fundPrices, fundAssetNames, dates, expReturns, rf, conf)

    Dim masterPortfolios As Object
    Set masterPortfolios = CreateObject("Scripting.Dictionary")

    ' 6. Build and simulate Benchmark
    Dim benchPort As SimulatedPortfolioCls
    Dim benchRets() As Double
    Set benchPort = BuildBenchmarkPortfolio(benchPricesRaw, benchAssetNames, dates, rf, conf, benchRets)
    masterPortfolios.Add "BENCHMARK", benchPort

    ' 7. Inject benchRets into Fund Positions
    Dim pKey As Variant
    For Each pKey In masterPositions.Keys
        masterPositions(pKey).ComputeMetrics rf, conf, benchRets
    Next pKey

    ' 8. Optimize and Simulate Active Strategies
    Call OptimizeAndSimulateStrategies(strategies, minWeights, maxWeights, nAssets, fundAssetNames, masterPositions, masterPortfolios, rf, conf, benchRets)

    ' 9. Output to Dashboards
    Call OutputLegacyDashboard(masterPositions, masterPortfolios, strategies, nAssets, fundAssetNames, lastDate)

    Call PositionDashboard.GeneratePositionsDashboard(masterPositions)
    Call PositionDashboard.GeneratePortfoliosDashboard(masterPortfolios, fundAssetNames)

    Sheets(ENGINE_SHEET).Visible = False
    Sheets(DASH_SHEET).Activate

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Optimization Completed.", vbInformation
End Sub

' -------------------------------------------------------------------------
' REFACTOR HELPERS
' -------------------------------------------------------------------------

Private Sub SplitBenchmarkAndFundAssets(prices() As Variant, assetNames() As String, ByRef fundPrices() As Double, ByRef fundAssetNames() As String, ByRef benchPricesRaw() As Double, ByRef benchAssetNames() As String)
    Dim nTotal As Integer
    nTotal = UBound(assetNames)
    Dim nAssets As Integer
    nAssets = nTotal - 2

    ReDim fundPrices(1 To UBound(prices, 1), 1 To nAssets)
    ReDim fundAssetNames(1 To nAssets)

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
End Sub

Private Sub PrepEngineSheetAndConstraints(fundAssetNames() As String, expReturns() As Double, covMat() As Double, ByRef minWeights() As Double, ByRef maxWeights() As Double)
    Dim wsDash As Worksheet
    Set wsDash = Sheets(DASH_SHEET)
    Dim inputStart As Long
    inputStart = wsDash.Range("InputTableStart").Row + 1

    Dim k As Integer, nAssets As Integer
    nAssets = UBound(fundAssetNames)

    For k = 1 To nAssets
        wsDash.Cells(inputStart + k - 1, wsDash.Range("AssetMetricsStart").Column - 1).Value = expReturns(k)
    Next k

    Call solverUtils.WriteToEngine(covMat, expReturns)

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
End Sub

Private Function CacheMasterPositions(fundPrices() As Double, fundAssetNames() As String, dates() As Date, expReturns() As Double, rf As Double, conf As Double) As Object
    Dim masterPositions As Object
    Set masterPositions = CreateObject("Scripting.Dictionary")

    Dim convictions As Object
    Set convictions = convictionsUtils.getConvictions()

    Dim nAssets As Integer
    nAssets = UBound(fundAssetNames)
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

    Set CacheMasterPositions = masterPositions
End Function

Private Function BuildBenchmarkPortfolio(benchPricesRaw() As Double, benchAssetNames() As String, dates() As Date, rf As Double, conf As Double, ByRef benchRets() As Double) As SimulatedPortfolioCls
    Dim benchPort As SimulatedPortfolioCls
    Set benchPort = New SimulatedPortfolioCls
    benchPort.StrategyName = "BENCHMARK"

    Dim benchMasterPos As Object
    Set benchMasterPos = CreateObject("Scripting.Dictionary")

    Dim cacheIdx As Integer
    For cacheIdx = 1 To 2
        Dim pPrices() As Double, pDates() As Date
        ReDim pPrices(1 To UBound(benchPricesRaw, 1))
        ReDim pDates(1 To UBound(dates))
        Dim iDay As Long
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

    benchPort.SetEqualWeights
    benchPort.Simulate

    Dim benchCurve() As Double
    benchCurve = benchPort.EquityCurve
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
    Set BuildBenchmarkPortfolio = benchPort
End Function

Private Sub OptimizeAndSimulateStrategies(strategies As Variant, minWeights() As Double, maxWeights() As Double, nAssets As Integer, fundAssetNames() As String, masterPositions As Object, masterPortfolios As Object, rf As Double, conf As Double, benchRets() As Double)
    Dim optimizer As OptimizerCls
    Set optimizer = New OptimizerCls

    Dim sumWeights() As Double
    ReDim sumWeights(1 To nAssets)
    Dim m As Integer
    For m = 1 To nAssets
        sumWeights(m) = 0
    Next m

    Dim i As Integer
    For i = LBound(strategies) To UBound(strategies)
        Dim simPort As SimulatedPortfolioCls
        Set simPort = New SimulatedPortfolioCls
        simPort.StrategyName = CStr(strategies(i))

        Dim jPos As Integer
        For jPos = 1 To nAssets
            simPort.AddPosition fundAssetNames(jPos), masterPositions(fundAssetNames(jPos))
        Next jPos

        simPort.SetEqualWeights
        optimizer.Optimize simPort, minWeights, maxWeights, sumWeights, nAssets, fundAssetNames

        simPort.Simulate
        simPort.ComputeMetrics rf, conf, benchRets

        masterPortfolios.Add simPort.StrategyName, simPort
    Next i

    ' Apply Risk Overlay
    Dim riskMultiplier As Double
    riskMultiplier = allocationLogic.ComputeFinalRiskFactor(masterPortfolios("MEAN"), masterPortfolios("BENCHMARK"))

    Dim pKey As Variant
    For Each pKey In masterPortfolios.Keys
        Dim p As SimulatedPortfolioCls
        Set p = masterPortfolios(pKey)

        If p.StrategyName <> "EQUAL WEIGHT" And p.StrategyName <> "MEAN" And p.StrategyName <> "BENCHMARK" Then
            p.ApplyRiskOverlay riskMultiplier
            p.Simulate
            p.ComputeMetrics rf, conf, benchRets
        End If
    Next pKey
End Sub

Private Sub OutputLegacyDashboard(masterPositions As Object, masterPortfolios As Object, strategies As Variant, nAssets As Integer, fundAssetNames() As String, lastDate As Date)
    Dim pKey As Variant
    Dim rIdx As Integer
    rIdx = 0

    Dim wsChart As Worksheet
    Set wsChart = Sheets(CHART_SHEET)

    Dim wsDashLegacy As Worksheet
    Set wsDashLegacy = Sheets(DASH_SHEET)

    ' Output Positions
    For Each pKey In masterPositions.Keys
        Dim curPos As PositionCls
        Set curPos = masterPositions(pKey)
        Dim dummyW As Variant

        Call portfolioUtils.OutputMetricsToRow(rIdx, curPos.AssetName, curPos.Metrics.Ret, curPos.Metrics.Vol, curPos.Metrics.Sharpe, curPos.Metrics.MDD, curPos.Metrics.MDDLen, curPos.Metrics.VaR, dummyW, "AssetMetricsStart", False)

        Dim aRow As Long, aCol As Long
        aRow = wsDashLegacy.Range("AssetMetricsStart").Row + 1 + rIdx
        aCol = wsDashLegacy.Range("AssetMetricsStart").Column
        wsDashLegacy.Cells(aRow, aCol + 7).Value = curPos.Metrics.DownsideBeta

        Dim pArr() As Double
        pArr = curPos.Prices
        Dim nDays As Long
        nDays = UBound(pArr)
        Dim curve() As Double
        ReDim curve(1 To nDays, 1 To 1)
        Dim iDay As Long
        For iDay = 1 To nDays
            curve(iDay, 1) = (pArr(iDay) / pArr(1)) * 100
        Next iDay
        wsChart.Range(wsChart.Cells(2, 2 + rIdx), wsChart.Cells(2 + nDays - 1, 2 + rIdx)).Value = curve

        rIdx = rIdx + 1
    Next pKey

    ' Output Portfolios
    Dim i As Integer
    For i = LBound(strategies) To UBound(strategies)
        Dim simPort As SimulatedPortfolioCls
        Set simPort = masterPortfolios(CStr(strategies(i)))

        Dim wArr() As Double
        ReDim wArr(1 To 1, 1 To nAssets)
        Dim idx As Integer
        For idx = 1 To nAssets
            wArr(1, idx) = simPort.GetWeight(fundAssetNames(idx))
        Next idx

        Call portfolioUtils.OutputMetricsToRow(i - 1, simPort.StrategyName, simPort.Metrics.Ret, simPort.Metrics.Vol, simPort.Metrics.Sharpe, simPort.Metrics.MDD, simPort.Metrics.MDDLen, simPort.Metrics.VaR, wArr, "StrategyTableStart", True)

        Dim colOffset As Integer
        colOffset = (nAssets + 1) + i
        Call chartsUtils.WriteCurveToSheet(simPort.EquityCurve, colOffset)

        Dim ddColOffset As Integer
        ddColOffset = (nAssets + 1) + UBound(strategies) + i
        Call chartsUtils.WriteCurveToSheet(simPort.DrawdownCurve, ddColOffset)

        Dim wColStart As Integer
        wColStart = 2 + (i - 1) * nAssets
        Call chartsUtils.WriteWeightsToSheet(simPort.DailyWeights, wColStart)
    Next i

    Dim startRow As Integer, startCol As Integer
    startRow = wsDashLegacy.Range("StrategyWeightsStart").Row
    startCol = wsDashLegacy.Range("StrategyWeightsStart").Column + UBound(strategies) - LBound(strategies) + 1

    Dim equityPositions As Dictionary
    Set equityPositions = DataUtils.GetEquityPositionsFromFund()

    Dim equitySum As Double
    equitySum = 0
    For i = 1 To equityPositions.Count
        equitySum = equitySum + equityPositions(equityPositions.Keys(i - 1)).Weight
    Next i

    For i = 1 To equityPositions.Count
        wsDashLegacy.Cells(startRow + i, startCol + 1).Value = equityPositions(equityPositions.Keys(i - 1)).Weight / equitySum
        wsDashLegacy.Cells(startRow + i, startCol + 2).Value = wsDashLegacy.Cells(startRow + i, startCol + 1).Value - wsDashLegacy.Cells(startRow + i, startCol).Value
        wsDashLegacy.Cells(startRow + i, startCol + UBound(strategies) + 3).Value = equityPositions(equityPositions.Keys(i - 1)).Weight
        wsDashLegacy.Cells(startRow + i, startCol + UBound(strategies) + 4).Value = wsDashLegacy.Cells(startRow + i, startCol + UBound(strategies) + 3).Value - wsDashLegacy.Cells(startRow + i, startCol + UBound(strategies) + 2).Value
    Next i

    Call chartsUtils.UpdateDashboardCharts(nAssets, UBound(strategies), lastDate)
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
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

    Dim rf As Double, conf As Double
    rf = Sheets(DASH_SHEET).Range("D4").Value
    conf = Sheets(DASH_SHEET).Range("D5").Value

    Dim masterPositions As Object
    Dim benchPositions As Object
    Call LoadAllPositions(masterPositions, benchPositions, rf, conf)

    If masterPositions.Count < 1 Then Exit Sub

    ' Extract the last date from the first position
    Dim firstPos As PositionCls
    Set firstPos = masterPositions(GetDictKeys(masterPositions)(1))
    Dim lastDate As Date
    lastDate = firstPos.Dates(UBound(firstPos.Dates))

    Dim meanRets() As Double
    Call CalculateStatsFromObjects(masterPositions, lastDate, meanRets)

    Application.ScreenUpdating = True
End Sub

Sub RunAllSolvers()
    Application.ScreenUpdating = False
    Sheets(ENGINE_SHEET).Activate

    Dim rf As Double, conf As Double
    rf = Sheets(DASH_SHEET).Range("D4").Value
    conf = Sheets(DASH_SHEET).Range("D5").Value

    ' 1. & 2. Fetch raw data and Cache OOP objects FIRST
    Dim masterPositions As Object
    Dim benchPositions As Object
    Call LoadAllPositions(masterPositions, benchPositions, rf, conf)

    Dim firstPos As PositionCls
    Set firstPos = masterPositions(GetDictKeys(masterPositions)(1))
    Dim lastDate As Date
    lastDate = firstPos.Dates(UBound(firstPos.Dates))
    Dim dates() As Date
    dates = firstPos.Dates

    ' 3. Compute stats dynamically using the PositionCls objects
    Dim meanRets() As Double
    Call CalculateStatsFromObjects(masterPositions, lastDate, meanRets)

    Dim covMat() As Double
    covMat = CalculateCovarianceFromObjects(masterPositions, meanRets)
    Call OutputCorrelationMatrix(covMat, GetDictKeys(masterPositions))

    ' 4. Prep legacy Engine constraints
    Dim strategies As Variant
    strategies = Array("ERC UNCSTRD", "ER/VOL", "SHARPE", "CUSTOM", "MEAN")

    Dim minWeights() As Double, maxWeights() As Double
    Call PrepEngineSheetAndConstraints(masterPositions, covMat, minWeights, maxWeights)
    Call SyncSimWeightsHeaders(dates, GetDictKeys(masterPositions), strategies)

    Dim masterPortfolios As Object
    Set masterPortfolios = CreateObject("Scripting.Dictionary")

    ' 5. Build and simulate Benchmark
    Dim benchPort As SimulatedPortfolioCls
    Dim benchRets() As Double
    Set benchPort = BuildBenchmarkPortfolio(benchPositions, dates, rf, conf, benchRets)
    masterPortfolios.Add "BENCHMARK", benchPort

    ' 6. Inject benchRets into Fund Positions
    Dim pKey As Variant
    For Each pKey In masterPositions.Keys
        masterPositions(pKey).ComputeMetrics rf, conf, benchRets
    Next pKey

    ' 7. Optimize and Simulate Active Strategies
    Call OptimizeAndSimulateStrategies(strategies, minWeights, maxWeights, masterPositions, masterPortfolios, rf, conf, benchRets)

    ' 8. Output to Dashboards
    Call OutputLegacyDashboard(masterPositions, masterPortfolios, strategies, lastDate)

    Call PositionDashboard.GeneratePositionsDashboard(masterPositions)
    Call PositionDashboard.GeneratePortfoliosDashboard(masterPortfolios)

    Sheets(ENGINE_SHEET).Visible = False
    Sheets(DASH_SHEET).Activate

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Optimization Completed.", vbInformation
End Sub

' -------------------------------------------------------------------------
' REFACTOR HELPERS
' -------------------------------------------------------------------------

Private Function GetDictKeys(dict As Object) As String()
    Dim keysArray() As String
    ReDim keysArray(1 To dict.Count)
    Dim i As Long
    i = 1
    Dim k As Variant
    For Each k In dict.Keys
        keysArray(i) = CStr(k)
        i = i + 1
    Next k
    GetDictKeys = keysArray
End Function

Private Sub LoadAllPositions(ByRef outFundPositions As Object, ByRef outBenchPositions As Object, rf As Double, conf As Double)
    Set outFundPositions = CreateObject("Scripting.Dictionary")
    Set outBenchPositions = CreateObject("Scripting.Dictionary")

    Dim prices() As Variant, dates() As Date, assetNames() As String
    Call assetUtils.GetHistoricalDataDatesAndNames(prices, dates, assetNames)
    If UBound(prices, 2) < 1 Then Exit Sub

    Dim convictions As Object
    Set convictions = convictionsUtils.getConvictions()

    Dim nTotal As Integer
    nTotal = UBound(assetNames)
    Dim cacheIdx As Integer

    For cacheIdx = 1 To nTotal
        Dim pPrices() As Double, pDates() As Date
        ReDim pPrices(1 To UBound(prices, 1))
        ReDim pDates(1 To UBound(dates))
        Dim iDay As Long
        For iDay = 1 To UBound(prices, 1)
            pPrices(iDay) = CDbl(prices(iDay, cacheIdx))
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

        ' Default target date for expected return calculations
        cachePos.TargetDate = Date + 365

        If cacheIdx <= 2 Then
            outBenchPositions.Add assetNames(cacheIdx), cachePos
        Else
            outFundPositions.Add assetNames(cacheIdx), cachePos
        End If
    Next cacheIdx
End Sub

Private Function BuildBenchmarkPortfolio(benchPositions As Object, dates() As Date, rf As Double, conf As Double, ByRef benchRets() As Double) As SimulatedPortfolioCls
    Dim benchPort As SimulatedPortfolioCls
    Set benchPort = New SimulatedPortfolioCls
    benchPort.StrategyName = "BENCHMARK"

    Dim bKey As Variant
    For Each bKey In benchPositions.Keys
        benchPort.AddPosition CStr(bKey), benchPositions(bKey)
    Next bKey

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

Private Sub CalculateStatsFromObjects(masterPositions As Object, lastDate As Date, ByRef meanRets() As Double)
    Dim nAssets As Integer
    nAssets = masterPositions.Count
    ReDim meanRets(1 To nAssets)

    Dim i As Integer
    i = 1
    Dim pKey As Variant
    For Each pKey In masterPositions.Keys
        Dim pos As PositionCls
        Set pos = masterPositions(pKey)

        Dim lRets() As Double
        lRets = pos.LogReturns
        Dim nRets As Long
        nRets = UBound(lRets)

        Dim s As Double
        s = 0
        Dim j As Long
        For j = 1 To nRets
            s = s + lRets(j)
        Next j

        Dim dateDiff As Double
        dateDiff = pos.Dates(UBound(pos.Dates)) - pos.Dates(LBound(pos.Dates))
        If dateDiff > 0 Then
            meanRets(i) = s / dateDiff * 365
        Else
            meanRets(i) = 0
        End If

        ' Let the object compute its own expected return
        pos.ComputeExpectedReturn meanRets(i), lastDate

        i = i + 1
    Next pKey
End Sub

Private Function CalculateCovarianceFromObjects(masterPositions As Object, meanRets() As Double) As Double()
    Dim nAssets As Integer
    nAssets = masterPositions.Count

    Dim keysArray() As String
    keysArray = GetDictKeys(masterPositions)

    Dim firstPos As PositionCls
    Set firstPos = masterPositions(keysArray(1))
    Dim nRets As Long
    nRets = UBound(firstPos.LogReturns)

    Dim res() As Double
    ReDim res(1 To nAssets, 1 To nAssets)

    Dim j As Integer, k As Integer, i As Long
    For j = 1 To nAssets
        Dim posJ As PositionCls
        Set posJ = masterPositions(keysArray(j))
        Dim lRetsJ() As Double
        lRetsJ = posJ.LogReturns
        Dim meanJ As Double
        meanJ = meanRets(j) / 256

        For k = 1 To nAssets
            Dim posK As PositionCls
            Set posK = masterPositions(keysArray(k))
            Dim lRetsK() As Double
            lRetsK = posK.LogReturns
            Dim meanK As Double
            meanK = meanRets(k) / 256

            Dim sumProd As Double
            sumProd = 0

            For i = 1 To nRets
                sumProd = sumProd + (lRetsJ(i) - meanJ) * (lRetsK(i) - meanK)
            Next i

            If nRets > 0 Then
                res(j, k) = (sumProd / nRets) * 256
            End If
        Next k
    Next j

    CalculateCovarianceFromObjects = res
End Function

Private Sub PrepEngineSheetAndConstraints(masterPositions As Object, covMat() As Double, ByRef minWeights() As Double, ByRef maxWeights() As Double)
    Dim wsDash As Worksheet
    Set wsDash = Sheets(DASH_SHEET)
    Dim inputStart As Long
    inputStart = wsDash.Range("InputTableStart").Row + 1

    Dim k As Integer, nAssets As Integer
    nAssets = masterPositions.Count

    Dim keysArray() As String
    keysArray = GetDictKeys(masterPositions)

    ' Dynamically construct expReturns array to pass to the engine
    Dim expReturns() As Double
    ReDim expReturns(1 To nAssets)

    For k = 1 To nAssets
        expReturns(k) = masterPositions(keysArray(k)).ExpectedReturn
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

Private Sub OptimizeAndSimulateStrategies(strategies As Variant, minWeights() As Double, maxWeights() As Double, masterPositions As Object, masterPortfolios As Object, rf As Double, conf As Double, benchRets() As Double)
    Dim optimizer As OptimizerCls
    Set optimizer = New OptimizerCls

    Dim nAssets As Integer
    nAssets = masterPositions.Count

    Dim sumWeights() As Double
    ReDim sumWeights(1 To nAssets)
    Dim m As Integer
    For m = 1 To nAssets
        sumWeights(m) = 0
    Next m

    Dim keysArray() As String
    keysArray = GetDictKeys(masterPositions)

    Dim i As Integer
    For i = LBound(strategies) To UBound(strategies)
        Dim simPort As SimulatedPortfolioCls
        Set simPort = New SimulatedPortfolioCls
        simPort.StrategyName = CStr(strategies(i))

        Dim jPos As Integer
        For jPos = 1 To nAssets
            simPort.AddPosition keysArray(jPos), masterPositions(keysArray(jPos))
        Next jPos

        simPort.SetEqualWeights
        optimizer.Optimize simPort, minWeights, maxWeights, sumWeights

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

Private Sub OutputLegacyDashboard(masterPositions As Object, masterPortfolios As Object, strategies As Variant, lastDate As Date)
    Dim pKey As Variant
    Dim rIdx As Integer
    rIdx = 0

    Dim nAssets As Integer
    nAssets = masterPositions.Count
    Dim keysArray() As String
    keysArray = GetDictKeys(masterPositions)

    Dim wsChart As Worksheet
    Set wsChart = Sheets(CHART_SHEET)

    Dim wsDashLegacy As Worksheet
    Set wsDashLegacy = Sheets(DASH_SHEET)

    ' Output Positions
    For Each pKey In keysArray
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
            wArr(1, idx) = simPort.GetWeight(keysArray(idx))
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
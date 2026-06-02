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

    Dim rf As Double, conf As Double, targetRatio As Double, maxHHI As Double
    On Error Resume Next
    Dim wsNewDash As Worksheet
    Set wsNewDash = Sheets("Positions Dashboard")
    If Not wsNewDash Is Nothing Then
        rf = wsNewDash.Cells(2, 2).Value
        conf = wsNewDash.Cells(3, 2).Value
        targetRatio = wsNewDash.Cells(6, 2).Value
        maxHHI = wsNewDash.Cells(7, 2).Value
    End If
    On Error GoTo 0

    ' Fail-safe defaults if the Positions Dashboard is deleted or hasn't been created yet
    If rf = 0 Then rf = 0.02
    If conf = 0 Then conf = 0.95
    If targetRatio = 0 Then targetRatio = 0.7
    If maxHHI = 0 Then maxHHI = 4

    Dim masterPositions As Object
    Dim benchPositions As Object
    Call LoadAllPositions(masterPositions, benchPositions, rf, conf)

    If masterPositions.Count < 1 Then Exit Sub

    Dim pKey As Variant
    For Each pKey In masterPositions.Keys
        masterPositions(pKey).ComputeMetrics rf, conf
    Next pKey

    ' Extract the last date from the first position
    Dim firstPos As PositionCls
    Set firstPos = masterPositions(GetDictKeys(masterPositions)(1))
    Dim tmpDates() As Date
    tmpDates = firstPos.Dates
    Dim lastDate As Date
    lastDate = tmpDates(UBound(tmpDates))

    ' Expected returns are computed natively inside LoadAllPositions! No loop needed here.

    ' Output the bare positions dashboard so the user can tweak parameters
    Call PositionDashboard.GenerateDashboard(masterPositions, Nothing)

    Application.ScreenUpdating = True
End Sub

Sub RunAllSolvers()
    Application.ScreenUpdating = False
    Sheets(ENGINE_SHEET).Activate

    Dim rf As Double, conf As Double, targetRatio As Double, maxHHI As Double
    On Error Resume Next
    Dim wsNewDash As Worksheet
    Set wsNewDash = Sheets("Positions Dashboard")
    If Not wsNewDash Is Nothing Then
        rf = wsNewDash.Cells(2, 2).Value
        conf = wsNewDash.Cells(3, 2).Value
        targetRatio = wsNewDash.Cells(6, 2).Value
        maxHHI = wsNewDash.Cells(7, 2).Value
    End If
    On Error GoTo 0

    ' Fail-safe defaults if the Positions Dashboard is deleted or hasn't been created yet
    If rf = 0 Then rf = 0.02
    If conf = 0 Then conf = 0.95
    If targetRatio = 0 Then targetRatio = 0.7
    If maxHHI = 0 Then maxHHI = 4

    ' 1. & 2. Fetch raw data and Cache OOP objects FIRST
    Dim masterPositions As Object
    Dim benchPositions As Object
    Call LoadAllPositions(masterPositions, benchPositions, rf, conf)

    Dim firstPos As PositionCls
    Set firstPos = masterPositions(GetDictKeys(masterPositions)(1))
    Dim tmpDates() As Date
    tmpDates = firstPos.Dates
    Dim lastDate As Date
    lastDate = tmpDates(UBound(tmpDates))
    Dim dates() As Date
    dates = firstPos.Dates

    ' 3. Compute stats dynamically using the PositionCls objects
    Dim meanRets() As Double
    ReDim meanRets(1 To masterPositions.Count)
    Dim i As Integer
    i = 1
    Dim pKey As Variant
    For Each pKey In masterPositions.Keys
        meanRets(i) = masterPositions(pKey).HistoricalMeanReturn
        i = i + 1
    Next pKey

    Dim covMat() As Double
    covMat = CalculateCovarianceFromObjects(masterPositions, meanRets)

    ' 4. Prep legacy Engine constraints
    Dim strategies As Variant
    strategies = Array("ERC UNCSTRD", "ER/VOL", "SHARPE", "CUSTOM", "MEAN")

    Call PrepEngineSheetAndConstraints(masterPositions, covMat, maxHHI)
    Call SyncSimWeightsHeaders(dates, GetDictKeys(masterPositions), strategies)

    Dim masterPortfolios As Object
    Set masterPortfolios = CreateObject("Scripting.Dictionary")

    ' 5. Build and simulate Benchmark
    Dim benchPort As SimulatedPortfolioCls
    Set benchPort = BuildBenchmarkPortfolio(benchPositions, dates, rf, conf)
    masterPortfolios.Add "BENCHMARK", benchPort

    ' 6. Inject benchPort into Fund Positions
    Dim pKey As Variant
    For Each pKey In masterPositions.Keys
        masterPositions(pKey).ComputeMetrics rf, conf, benchPort
    Next pKey

    ' 7. Optimize and Simulate Active Strategies
    Call OptimizeAndSimulateStrategies(strategies, masterPositions, masterPortfolios, rf, conf, benchPort, targetRatio)

    ' 8. Output to Dashboards
    Call PositionDashboard.GenerateDashboard(masterPositions, masterPortfolios)

    Sheets(ENGINE_SHEET).Visible = False

    On Error Resume Next
    Sheets("Positions Dashboard").Activate
    On Error GoTo 0

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

        ' Set default bounds
        cachePos.MinWeight = 0
        cachePos.MaxWeight = 1

        ' Default target date for expected return calculations
        cachePos.TargetDate = Date + 365

        If cacheIdx <= 2 Then
            outBenchPositions.Add assetNames(cacheIdx), cachePos
        Else
            outFundPositions.Add assetNames(cacheIdx), cachePos
        End If
    Next cacheIdx

    ' AFTER all positions are initialized, overwrite the defaults using the user's manual dashboard inputs!
    Call PositionDashboard.GetDashboardInputs(outFundPositions)

    ' Compute Metrics and Expected Returns for all positions
    Dim pKey As Variant
    For Each pKey In outBenchPositions.Keys
        outBenchPositions(pKey).ComputeMetrics rf, conf
    Next pKey

    For Each pKey In outFundPositions.Keys
        outFundPositions(pKey).ComputeMetrics rf, conf
        outFundPositions(pKey).ComputeExpectedReturn
    Next pKey
End Sub

Private Function BuildBenchmarkPortfolio(benchPositions As Object, dates() As Date, rf As Double, conf As Double) As SimulatedPortfolioCls
    Dim benchPort As SimulatedPortfolioCls
    Set benchPort = New SimulatedPortfolioCls
    benchPort.StrategyName = "BENCHMARK"

    Dim bKey As Variant
    For Each bKey In benchPositions.Keys
        benchPort.AddPosition CStr(bKey), benchPositions(bKey)
    Next bKey

    benchPort.SetEqualWeights
    benchPort.Simulate

    benchPort.ComputeMetrics rf, conf
    Set BuildBenchmarkPortfolio = benchPort
End Function


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

Private Sub PrepEngineSheetAndConstraints(masterPositions As Object, covMat() As Double, maxHHI As Double)
    Dim k As Integer, nAssets As Integer
    nAssets = masterPositions.Count

    Dim keysArray() As String
    keysArray = GetDictKeys(masterPositions)

    ' Dynamically construct expReturns array to pass to the engine
    Dim expReturns() As Double
    ReDim expReturns(1 To nAssets)

    For k = 1 To nAssets
        expReturns(k) = masterPositions(keysArray(k)).ExpectedReturn
    Next k

    Call solverUtils.WriteToEngine(covMat, expReturns, maxHHI)
End Sub

Private Sub OptimizeAndSimulateStrategies(strategies As Variant, masterPositions As Object, masterPortfolios As Object, rf As Double, conf As Double, benchPort As SimulatedPortfolioCls, targetRatio As Double)
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
        optimizer.Optimize simPort, sumWeights

        simPort.Simulate
        simPort.ComputeMetrics rf, conf, benchPort

        masterPortfolios.Add simPort.StrategyName, simPort
    Next i

    ' Apply Risk Overlay
    Dim riskDict As Object
    Set riskDict = allocationLogic.ComputeFinalRiskFactor(masterPortfolios("MEAN"), masterPortfolios("BENCHMARK"), targetRatio)

    Dim pKey As Variant
    For Each pKey In masterPortfolios.Keys
        Dim p As SimulatedPortfolioCls
        Set p = masterPortfolios(pKey)

        If p.StrategyName <> "EQUAL WEIGHT" And p.StrategyName <> "MEAN" And p.StrategyName <> "BENCHMARK" Then
            p.ApplyRiskOverlay riskDict
            p.Simulate
            p.ComputeMetrics rf, conf, benchPort
        End If
    Next pKey
End Sub

' ==========================================================
' CLEAN
' ==========================================================

Sub RunClean()
    On Error Resume Next
    Sheets(CHART_SHEET).Cells.ClearContents
    Sheets(WEIGHTS_SHEET).Cells.ClearContents
    On Error GoTo 0
End Sub

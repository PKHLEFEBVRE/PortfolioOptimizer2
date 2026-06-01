Attribute VB_Name = "assetUtils"
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
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptimizerCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DASH_SHEET As String = "Dashboard"
Private Const ENGINE_SHEET As String = "Engine"

' Optimizes the weights in the provided portfolio based on its StrategyName.
' It modifies the portfolio's weights in place.
Public Sub Optimize(port As SimulatedPortfolioCls, minWeights() As Double, maxWeights() As Double, ByRef sumWeights() As Double, nAssets As Integer, assetNames() As String)
    Dim wsDash As Worksheet, wsEng As Worksheet
    Dim optRange As Range, optRangeC As Range
    Set wsDash = Sheets(DASH_SHEET)
    Set wsEng = Sheets(ENGINE_SHEET)
    Set optRange = wsEng.Range("OptWeights")
    Set optRangeC = wsEng.Range("OptWeights")

    Dim stratName As String
    stratName = UCase(port.StrategyName)

    Dim inputStart As Long
    inputStart = wsDash.Range("InputTableStart").Row + 1

    Dim cVal As Double
    Dim m As Integer, j As Integer
    Dim wCount As Integer
    wCount = 4 ' Hardcoded count of active optimization strategies contributing to MEAN. In original, UBound(strategies) - 1

    ' Always start the engine sheet with equal weights for the solver
    Dim eqW As Double
    eqW = 1# / nAssets
    For j = 1 To nAssets
        optRange.Cells(1, j).Value = eqW
    Next j

    If stratName = "EQUAL WEIGHT" Then
        ' Portfolio is already initialized with equal weights. Do nothing.

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
            port.SetWeight assetNames(m), cVal
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
            For m = 1 To nAssets
                port.SetWeight assetNames(m), optRange.Cells(1, m).Value / currentSum
            Next m
        End If

    ElseIf stratName = "MEAN" Then
        For m = 1 To nAssets
            port.SetWeight assetNames(m), sumWeights(m) / wCount
        Next m

    Else
        ' Solver-based Optimization: SHARPE, Variance, Kelly, ERC UNCSTRD
        Call solverUtils.RunSolver(stratName, minWeights, maxWeights, False)

        Dim w As Variant
        w = wsEng.Range("OptWeights").Value

        For m = 1 To nAssets
            port.SetWeight assetNames(m), CDbl(w(1, m))
        Next m
    End If

    ' Update the running sumWeights array for the MEAN strategy computation
    If stratName <> "MEAN" Then
        For m = 1 To nAssets
            sumWeights(m) = sumWeights(m) + port.GetWeight(assetNames(m))
        Next m
    End If
End Sub

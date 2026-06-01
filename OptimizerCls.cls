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
Public Sub Optimize(port As SimulatedPortfolioCls, minWeights() As Double, maxWeights() As Double, ByRef sumWeights() As Double)
    Dim wsEng As Worksheet
    Dim optRange As Range
    Set wsEng = Sheets(ENGINE_SHEET)
    Set optRange = wsEng.Range("OptWeights")

    Dim stratName As String
    stratName = UCase(port.StrategyName)

    Dim cVal As Double
    Dim m As Integer, j As Integer
    Dim wCount As Integer
    wCount = 4 ' Hardcoded count of active optimization strategies contributing to MEAN.

    ' Retrieve the asset names dynamically from the portfolio
    Dim nAssets As Integer
    Dim posKeys() As Variant

    posKeys = port.GetPositionNames()
    nAssets = UBound(posKeys)

    ' Always start the engine sheet with equal weights for the solver (if solver is used)
    Dim eqW As Double
    eqW = 1# / nAssets
    For j = 1 To nAssets
        optRange.Cells(1, j).Value = eqW
    Next j

    If stratName = "EQUAL WEIGHT" Then
        ' Portfolio is already initialized with equal weights. Do nothing.

    ElseIf stratName = "CUSTOM" Then
        Dim sum_ As Double
        sum_ = 0
        Dim pos As PositionCls

        ' Sum up all convictions
        For m = 1 To nAssets
            Set pos = port.GetPosition(CStr(posKeys(m)))
            If Not pos Is Nothing Then
                sum_ = sum_ + pos.Conviction
            End If
        Next m

        ' Assign normalized conviction weights
        For m = 1 To nAssets
            Set pos = port.GetPosition(CStr(posKeys(m)))
            If Not pos Is Nothing And sum_ > 0 Then
                cVal = pos.Conviction / sum_
                port.SetWeight CStr(posKeys(m)), cVal
            Else
                port.SetWeight CStr(posKeys(m)), 0
            End If
        Next m

    ElseIf stratName = "ER/VOL" Then
        Dim currentSum As Double
        currentSum = 0

        ' Temporarily store raw ER/VOL values
        Dim erVolArr() As Double
        ReDim erVolArr(1 To nAssets)

        For m = 1 To nAssets
            Set pos = port.GetPosition(CStr(posKeys(m)))
            If Not pos Is Nothing Then
                If pos.Metrics.Vol > 0 Then
                    erVolArr(m) = pos.ExpectedReturn / pos.Metrics.Vol
                Else
                    erVolArr(m) = 0
                End If
                currentSum = currentSum + erVolArr(m)
            End If
        Next m

        ' Normalize and assign
        For m = 1 To nAssets
            If currentSum <> 0 Then
                port.SetWeight CStr(posKeys(m)), erVolArr(m) / currentSum
            Else
                port.SetWeight CStr(posKeys(m)), 0
            End If
        Next m

    ElseIf stratName = "MEAN" Then
        For m = 1 To nAssets
            port.SetWeight CStr(posKeys(m)), sumWeights(m) / wCount
        Next m

    Else
        ' Solver-based Optimization: SHARPE, Variance, Kelly, ERC UNCSTRD
        Call solverUtils.RunSolver(stratName, minWeights, maxWeights, False)

        Dim w As Variant
        w = wsEng.Range("OptWeights").Value

        For m = 1 To nAssets
            port.SetWeight CStr(posKeys(m)), CDbl(w(1, m))
        Next m
    End If

    ' Update the running sumWeights array for the MEAN strategy computation
    If stratName <> "MEAN" Then
        For m = 1 To nAssets
            sumWeights(m) = sumWeights(m) + port.GetWeight(CStr(posKeys(m)))
        Next m
    End If
End Sub

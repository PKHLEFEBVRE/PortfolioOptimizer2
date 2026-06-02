VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimulatedPortfolioCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_StrategyName As String
Private p_Positions As Object ' Dictionary of PositionCls
Private p_Weights As Object   ' Dictionary of weights matching positions
Private p_EquityCurve() As Double
Private p_DailyWeights() As Double ' 2D array: (1 to nDays, 1 to nAssets)
Private p_DrawdownCurve() As Double
Private p_LogRets() As Double
Private p_Dates() As Date
Private p_Metrics As MetricsCalculatorCls

Public Property Get StrategyName() As String
    StrategyName = p_StrategyName
End Property

Public Property Let StrategyName(ByVal value As String)
    p_StrategyName = value
End Property

Public Property Get EquityCurve() As Double()
    EquityCurve = p_EquityCurve
End Property

Public Property Get DrawdownCurve() As Double()
    DrawdownCurve = p_DrawdownCurve
End Property

Public Property Get DailyWeights() As Double()
    DailyWeights = p_DailyWeights
End Property

Public Property Get LogReturns() As Double()
    LogReturns = p_LogRets
End Property

Public Property Get Metrics() As MetricsCalculatorCls
    Set Metrics = p_Metrics
End Property

Private p_DDFactor As Double
Private p_CVaRFactor As Double
Private p_SemiDevFactor As Double
Private p_HistoricalMDDFactor As Double
Private p_FinalRiskFactor As Double

Public Property Get DDFactor() As Double: DDFactor = p_DDFactor: End Property
Public Property Let DDFactor(ByVal value As Double): p_DDFactor = value: End Property

Public Property Get CVaRFactor() As Double: CVaRFactor = p_CVaRFactor: End Property
Public Property Let CVaRFactor(ByVal value As Double): p_CVaRFactor = value: End Property

Public Property Get SemiDevFactor() As Double: SemiDevFactor = p_SemiDevFactor: End Property
Public Property Let SemiDevFactor(ByVal value As Double): p_SemiDevFactor = value: End Property

Public Property Get HistoricalMDDFactor() As Double: HistoricalMDDFactor = p_HistoricalMDDFactor: End Property
Public Property Let HistoricalMDDFactor(ByVal value As Double): p_HistoricalMDDFactor = value: End Property

Public Property Get FinalRiskFactor() As Double: FinalRiskFactor = p_FinalRiskFactor: End Property
Public Property Let FinalRiskFactor(ByVal value As Double): p_FinalRiskFactor = value: End Property

Private Sub Class_Initialize()
    Set p_Positions = CreateObject("Scripting.Dictionary")
    Set p_Weights = CreateObject("Scripting.Dictionary")
    Set p_Metrics = New MetricsCalculatorCls
End Sub

Public Sub AddPosition(ByVal assetName As String, ByRef pos As PositionCls)
    If Not p_Positions.Exists(assetName) Then
        p_Positions.Add assetName, pos
    End If
End Sub

Public Sub SetEqualWeights()
    Dim nAssets As Long
    nAssets = p_Positions.Count
    If nAssets = 0 Then Exit Sub

    Dim eqW As Double
    eqW = 1# / nAssets

    p_Weights.RemoveAll

    Dim key As Variant
    For Each key In p_Positions.Keys
        p_Weights.Add key, eqW
    Next key
End Sub

Public Sub SetWeight(ByVal assetName As String, ByVal weight As Double)
    If p_Weights.Exists(assetName) Then
        p_Weights(assetName) = weight
    Else
        p_Weights.Add assetName, weight
    End If
End Sub

Public Function GetWeight(ByVal assetName As String) As Double
    If p_Weights.Exists(assetName) Then
        GetWeight = p_Weights(assetName)
    Else
        GetWeight = 0
    End If
End Function

Public Function GetPositionNames() As Variant()
    Dim keysArray() As Variant
    Dim n As Long
    n = p_Positions.Count
    If n = 0 Then
        GetPositionNames = Array()
        Exit Function
    End If

    ReDim keysArray(1 To n)
    Dim k As Variant
    Dim i As Long
    i = 1
    For Each k In p_Positions.Keys
        keysArray(i) = k
        i = i + 1
    Next k

    GetPositionNames = keysArray
End Function

Public Function GetPosition(ByVal assetName As String) As PositionCls
    If p_Positions.Exists(assetName) Then
        Set GetPosition = p_Positions(assetName)
    Else
        Set GetPosition = Nothing
    End If
End Function

Public Sub ApplyRiskOverlay(riskDict As Object)
    p_DDFactor = riskDict("DDFactor")
    p_CVaRFactor = riskDict("CVaRFactor")
    p_SemiDevFactor = riskDict("SemiDevFactor")
    p_HistoricalMDDFactor = riskDict("HistoricalMDDFactor")
    p_FinalRiskFactor = riskDict("FinalRiskFactor")

    Dim key As Variant
    For Each key In p_Weights.Keys
        p_Weights(key) = p_Weights(key) * p_FinalRiskFactor
    Next key
End Sub

Public Sub Simulate()
    If p_Positions.Count = 0 Then Exit Sub

    Dim keys As Variant
    keys = p_Positions.Keys

    Dim firstPos As PositionCls
    Set firstPos = p_Positions(keys(0))
    Dim nDays As Long
    nDays = UBound(firstPos.Prices)
    Dim nAssets As Long
    nAssets = p_Positions.Count

    ReDim p_EquityCurve(1 To nDays, 1 To 1)
    ReDim p_DrawdownCurve(1 To nDays, 1 To 1)
    ReDim p_DailyWeights(1 To nDays, 1 To nAssets)
    ReDim p_LogRets(1 To nDays - 1)
    ReDim p_Dates(1 To nDays)

    Dim i As Long, j As Long
    p_EquityCurve(1, 1) = 100

    Dim currentVal As Double
    currentVal = 100

    Dim shares() As Double
    ReDim shares(1 To nAssets)

    Dim prices() As Double
    ReDim prices(1 To nDays, 1 To nAssets)

    Dim sumW As Double
    sumW = 0

    For j = 1 To nAssets
        Dim pos As PositionCls
        Set pos = p_Positions(keys(j - 1))

        Dim pDates() As Date
        pDates = pos.Dates
        Dim pPrices() As Double
        pPrices = pos.Prices

        For i = 1 To nDays
            prices(i, j) = pPrices(i)
            If j = 1 Then p_Dates(i) = pDates(i)
        Next i

        Dim w As Double
        w = p_Weights(keys(j - 1))
        sumW = sumW + w

        shares(j) = (currentVal * w) / prices(1, j)
        p_DailyWeights(1, j) = w
    Next j

    Dim cash As Double
    cash = currentVal * (1 - sumW)

    Dim peak As Double
    peak = 100
    p_DrawdownCurve(1, 1) = 0

    For i = 2 To nDays
        Dim dPrev As Date, dCurr As Date
        dPrev = p_Dates(i - 1)
        dCurr = p_Dates(i)

        Dim prevTotal As Double
        prevTotal = p_EquityCurve(i - 1, 1)

        Dim assetVal As Double
        assetVal = cash ' Start with the cash balance
        For j = 1 To nAssets
            assetVal = assetVal + shares(j) * prices(i, j)
        Next j

        ' Rebalance monthly AFTER calculating the day's total value
        If Month(dCurr) <> Month(dPrev) Then
            cash = assetVal * (1 - sumW)
            For j = 1 To nAssets
                w = p_Weights(keys(j - 1))
                shares(j) = (assetVal * w) / prices(i, j)
            Next j
        End If

        p_EquityCurve(i, 1) = assetVal

        If assetVal > peak Then
            peak = assetVal
            p_DrawdownCurve(i, 1) = 0
        Else
            p_DrawdownCurve(i, 1) = (assetVal / peak) - 1
        End If

        For j = 1 To nAssets
            If assetVal > 0 Then
                p_DailyWeights(i, j) = (shares(j) * prices(i, j)) / assetVal
            Else
                p_DailyWeights(i, j) = 0
            End If
        Next j

        p_LogRets(i - 1) = Log(p_EquityCurve(i, 1) / p_EquityCurve(i - 1, 1))
    Next i
End Sub

Public Sub ComputeMetrics(riskFreeRate As Double, varConfidence As Double, Optional benchPort As SimulatedPortfolioCls = Nothing)
    Dim nDays As Long
    nDays = UBound(p_EquityCurve, 1)

    Dim eq1D() As Double
    ReDim eq1D(1 To nDays)
    Dim i As Long
    For i = 1 To nDays
        eq1D(i) = p_EquityCurve(i, 1)
    Next i

    ' Simulated Portfolios use Standard Deviation volatility by default in original code
    p_Metrics.Calculate eq1D, p_Dates, p_LogRets, False, riskFreeRate, varConfidence, benchPort, "Portfolio"
End Sub

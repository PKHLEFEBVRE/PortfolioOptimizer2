VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PositionCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_AssetName As String
Private p_Prices() As Double
Private p_Dates() As Date
Private p_LogRets() As Double
Private p_Metrics As MetricsCalculatorCls

Public Property Get AssetName() As String
    AssetName = p_AssetName
End Property

Public Property Let AssetName(ByVal value As String)
    p_AssetName = value
End Property

Public Property Get Prices() As Double()
    Prices = p_Prices
End Property

Public Property Get Dates() As Date()
    Dates = p_Dates
End Property

Public Property Get LogReturns() As Double()
    LogReturns = p_LogRets
End Property

Public Property Get Metrics() As MetricsCalculatorCls
    Set Metrics = p_Metrics
End Property

Private Sub Class_Initialize()
    Set p_Metrics = New MetricsCalculatorCls
End Sub

Public Sub InitializeData(pricesIn() As Double, datesIn() As Date)
    Dim i As Long
    Dim nDays As Long

    nDays = UBound(pricesIn) - LBound(pricesIn) + 1

    ReDim p_Prices(1 To nDays)
    ReDim p_Dates(1 To nDays)
    ReDim p_LogRets(1 To nDays - 1)

    For i = 1 To nDays
        p_Prices(i) = pricesIn(i - 1 + LBound(pricesIn))
        p_Dates(i) = datesIn(i - 1 + LBound(datesIn))

        If i > 1 Then
            Dim pNow As Double, pPrev As Double
            pNow = p_Prices(i)
            pPrev = p_Prices(i - 1)
            If pPrev > 0 And pNow > 0 Then
                p_LogRets(i - 1) = Log(pNow / pPrev)
            Else
                p_LogRets(i - 1) = 0
            End If
        End If
    Next i
End Sub

Public Sub ComputeMetrics(riskFreeRate As Double, varConfidence As Double, Optional benchRets As Variant)
    ' Individual assets use EWMA volatility by default in original code
    p_Metrics.Calculate p_Prices, p_Dates, p_LogRets, True, riskFreeRate, varConfidence, benchRets, "Asset"
End Sub

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

' Convictions Attributes
Private p_Conviction As Double
Private p_TP As Double
Private p_TPProba As Double
Private p_SP As Double
Private p_SPProba As Double
Private p_ExpectedReturn As Double
Private p_TargetDate As Date
Private p_MinWeight As Double
Private p_MaxWeight As Double
Private p_HistoricalMeanReturn As Double


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

Public Property Get Conviction() As Double
    Conviction = p_Conviction
End Property

Public Property Let Conviction(ByVal value As Double)
    p_Conviction = value
End Property

Public Property Get TP() As Double
    TP = p_TP
End Property

Public Property Let TP(ByVal value As Double)
    p_TP = value
End Property

Public Property Get TPProba() As Double
    TPProba = p_TPProba
End Property

Public Property Let TPProba(ByVal value As Double)
    p_TPProba = value
End Property

Public Property Get SP() As Double
    SP = p_SP
End Property

Public Property Let SP(ByVal value As Double)
    p_SP = value
End Property

Public Property Get SPProba() As Double
    SPProba = p_SPProba
End Property

Public Property Let SPProba(ByVal value As Double)
    p_SPProba = value
End Property

Public Property Get ExpectedReturn() As Double
    ExpectedReturn = p_ExpectedReturn
End Property

Public Property Let ExpectedReturn(ByVal value As Double)
    p_ExpectedReturn = value
End Property

Public Property Get TargetDate() As Date
    TargetDate = p_TargetDate
End Property

Public Property Let TargetDate(ByVal value As Date)
    p_TargetDate = value
End Property

Public Property Get MinWeight() As Double
    MinWeight = p_MinWeight
End Property

Public Property Let MinWeight(ByVal value As Double)
    p_MinWeight = value
End Property

Public Property Get MaxWeight() As Double
    MaxWeight = p_MaxWeight
End Property

Public Property Let MaxWeight(ByVal value As Double)
    p_MaxWeight = value
End Property

Public Property Get HistoricalMeanReturn() As Double
    HistoricalMeanReturn = p_HistoricalMeanReturn
End Property

Public Property Let HistoricalMeanReturn(ByVal value As Double)
    p_HistoricalMeanReturn = value
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

Public Sub ComputeMetrics(riskFreeRate As Double, varConfidence As Double, Optional benchPort As SimulatedPortfolioCls = Nothing)
    ' Individual assets use EWMA volatility by default in original code
    p_Metrics.Calculate p_Prices, p_Dates, p_LogRets, True, riskFreeRate, varConfidence, benchPort, "Asset"
End Sub

Public Sub ComputeExpectedReturn(currDate As Date)
    Dim pCurr As Double, pTgt As Double, wTgt As Double
    Dim pStop As Double, wStop As Double

    If UBound(p_Prices) >= LBound(p_Prices) Then
        pCurr = p_Prices(UBound(p_Prices))
    Else
        p_ExpectedReturn = 0
        Exit Sub
    End If

    ' 1. Calculate Historical Mean Return internally
    Dim nRets As Long
    nRets = UBound(p_LogRets)
    Dim s As Double
    s = 0
    Dim j As Long
    For j = 1 To nRets
        s = s + p_LogRets(j)
    Next j

    Dim dateDiff As Double
    dateDiff = p_Dates(UBound(p_Dates)) - p_Dates(LBound(p_Dates))
    If dateDiff > 0 Then
        p_HistoricalMeanReturn = s / dateDiff * 365
    Else
        p_HistoricalMeanReturn = 0
    End If

    ' 2. Calculate Expected Return
    pTgt = p_TP
    wTgt = p_TPProba
    pStop = p_SP
    wStop = p_SPProba

    Dim T_years As Double
    If p_TargetDate > currDate Then
        T_years = (p_TargetDate - currDate) / 365
    Else
        T_years = 1
    End If

    Dim termTgt As Double, termStop As Double, termHist As Double
    termTgt = pTgt * wTgt
    termStop = pStop * wStop
    termHist = pCurr * (Exp(p_HistoricalMeanReturn) ^ T_years) * (1 - wTgt - wStop)

    Dim eFinal As Double
    eFinal = termTgt + termStop + termHist

    If pCurr > 0 And T_years > 0 Then
        p_ExpectedReturn = (eFinal / pCurr) ^ (1 / T_years) - 1
    Else
        p_ExpectedReturn = 0
    End If
End Sub

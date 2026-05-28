VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDashSheet As String
Private pWeightsSheet As String
Private pEngineSheet As String
Private pChartSheet As String
Private pPriceSheet As String
Private pRiskFreeRow As Integer
Private pTargetRatio As Double
Private pCVaRConfidence As Double
Private pHalfLife As Double

Private Sub Class_Initialize()
    pDashSheet = "Dashboard"
    pWeightsSheet = "SimulatedWeights"
    pEngineSheet = "Engine"
    pChartSheet = "ChartData"
    pPriceSheet = "PriceHistory"
    pRiskFreeRow = 4
    pTargetRatio = 0.7
    pCVaRConfidence = 0.95
    pHalfLife = 256
End Sub

Public Property Get DashSheet() As String
    DashSheet = pDashSheet
End Property

Public Property Get WeightsSheet() As String
    WeightsSheet = pWeightsSheet
End Property

Public Property Get EngineSheet() As String
    EngineSheet = pEngineSheet
End Property

Public Property Get ChartSheet() As String
    ChartSheet = pChartSheet
End Property

Public Property Get PriceSheet() As String
    PriceSheet = pPriceSheet
End Property

Public Property Get RiskFreeRow() As Integer
    RiskFreeRow = pRiskFreeRow
End Property

Public Property Get TargetRatio() As Double
    TargetRatio = pTargetRatio
End Property

Public Property Get CVaRConfidence() As Double
    CVaRConfidence = pCVaRConfidence
End Property

Public Property Get HalfLife() As Double
    HalfLife = pHalfLife
End Property

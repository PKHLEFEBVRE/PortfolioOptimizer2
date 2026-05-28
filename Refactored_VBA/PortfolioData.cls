VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PortfolioData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pConfig As Config

Private pChartRets() As Double
Private pBenchRets() As Double
Private pNumMatched As Long
Private pIsLoaded As Boolean

Private pBenchMDD As Double
Private pPortMDD As Double
Private pBenchCVaR As Double
Private pPortCVaR As Double
Private pBenchSemiDev As Double
Private pPortSemiDev As Double
Private pDSBeta As Double

Private Sub Class_Initialize()
    pIsLoaded = False
    Set pConfig = New Config
End Sub

Public Sub Init(configObj As Config)
    Set pConfig = configObj
End Sub

Public Sub SetReturns(chartReturns() As Double, benchReturns() As Double, numMatched As Long)
    pChartRets = chartReturns
    pBenchRets = benchReturns
    pNumMatched = numMatched
    pIsLoaded = True
End Sub

Public Property Get ChartReturns() As Double()
    ChartReturns = pChartRets
End Property

Public Property Get BenchReturns() As Double()
    BenchReturns = pBenchRets
End Property

Public Property Get NumMatched() As Long
    NumMatched = pNumMatched
End Property

Public Property Get IsLoaded() As Boolean
    IsLoaded = pIsLoaded
End Property

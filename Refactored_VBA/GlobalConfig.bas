Option Explicit
Option Base 1

Public config As Config
Public portDataCache As PortfolioData

Public Sub InitializeGlobals()
    If config Is Nothing Then Set config = New Config
    If portDataCache Is Nothing Then
        Set portDataCache = New PortfolioData
        portDataCache.Init config
    End If
End Sub

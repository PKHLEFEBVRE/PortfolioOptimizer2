Option Explicit
Option Base 1

' =========================================================================
' REFACTORED DASHBOARD BUILDER MODULE
' =========================================================================

' Creates the base layout of the dashboard
Public Sub BuildMainDashboard()
    InitializeGlobals

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(config.DashSheet)

    Application.ScreenUpdating = False

    ' Title
    With wsDash.Range("B2")
        .Value = "Portfolio Optimizer Dashboard"
        .Font.Size = 16
        .Font.Bold = True
    End With

    ' Parameters Section
    CreateParametersSection wsDash, 4, 2

    ' Tables Section
    ' Creating standard tables starting at row 15
    CreateInputTable wsDash, 15, 2, "InputTable"
    CreateWeightsTable wsDash, 15, 8, "StrategyWeights"

    ' Adjust column widths
    wsDash.Columns("B:M").AutoFit

    Application.ScreenUpdating = True
    MsgBox "Dashboard layout has been built/refreshed.", vbInformation
End Sub

Private Sub CreateParametersSection(ws As Worksheet, startRow As Long, startCol As Long)
    ' Setup parameters for Config
    ws.Cells(startRow, startCol).Value = "Global Settings"
    ws.Cells(startRow, startCol).Font.Bold = True

    Dim labels As Variant
    labels = Array("Target Ratio", "CVaR Confidence", "Half Life", "Risk Free Rate Row", "Max HHI Limit", "Eq Weighting Limit")

    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        ws.Cells(startRow + i, startCol).Value = labels(i)
        ws.Cells(startRow + i, startCol + 1).Interior.Color = RGB(220, 230, 241) ' Light blue input color
        ws.Cells(startRow + i, startCol + 1).Borders.Weight = xlThin
    Next i

    ' Default values and named ranges
    ' These named ranges can later be read by the Config class
    ws.Cells(startRow + 1, startCol + 1).Value = 0.7
    ws.Cells(startRow + 1, startCol + 1).Name = "TargetRatioInput"

    ws.Cells(startRow + 2, startCol + 1).Value = 0.95
    ws.Cells(startRow + 2, startCol + 1).Name = "CVaRConfidenceInput"

    ws.Cells(startRow + 3, startCol + 1).Value = 256
    ws.Cells(startRow + 3, startCol + 1).Name = "HalfLifeInput"

    ws.Cells(startRow + 4, startCol + 1).Value = 4
    ws.Cells(startRow + 4, startCol + 1).Name = "RiskFreeRowInput"

    ws.Cells(startRow + 5, startCol + 1).Value = 0.2
    ws.Cells(startRow + 5, startCol + 1).Name = "MaxHHIInput"

    ws.Cells(startRow + 6, startCol + 1).Value = 1#
    ws.Cells(startRow + 6, startCol + 1).Name = "EqWeightingLimit"
End Sub

Private Sub CreateInputTable(ws As Worksheet, startRow As Long, startCol As Long, tableName As String)
    ' Setup a standard input table
    ws.Cells(startRow, startCol).Value = "Assets Inputs"
    ws.Cells(startRow, startCol).Font.Bold = True

    Dim headers As Variant
    headers = Array("Asset Ticker", "Include", "Min Weight", "Max Weight", "Conviction")

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(startRow + 1, startCol + i - 1)
            .Value = headers(i)
            .Font.Bold = True
            .Interior.Color = RGB(180, 198, 231)
            .Borders.Weight = xlThin
        End With
    Next i

    ' Name the range for easier macro targeting (like InputTableStart in Main.bas)
    On Error Resume Next
    ThisWorkbook.Names(tableName & "Start").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=tableName & "Start", RefersTo:=ws.Cells(startRow + 2, startCol + 1)
End Sub

Private Sub CreateWeightsTable(ws As Worksheet, startRow As Long, startCol As Long, tableName As String)
    ' Setup a standard weights output table
    ws.Cells(startRow, startCol).Value = "Strategy Outputs"
    ws.Cells(startRow, startCol).Font.Bold = True

    Dim headers As Variant
    headers = Array("Asset", "ERC UNCSTRD", "MIN VAR", "SHARPE", "KELLY")

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(startRow + 1, startCol + i - 1)
            .Value = headers(i)
            .Font.Bold = True
            .Interior.Color = RGB(198, 224, 180) ' Light green
            .Borders.Weight = xlThin
        End With
    Next i

    ' Name the range
    On Error Resume Next
    ThisWorkbook.Names(tableName & "Start").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=tableName & "Start", RefersTo:=ws.Cells(startRow + 2, startCol + 1)
End Sub

' Function to add a dynamic new section (e.g., custom table)
Public Sub AddCustomSection(sectionName As String, startRow As Long, startCol As Long, numCols As Long)
    InitializeGlobals

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(config.DashSheet)

    wsDash.Cells(startRow, startCol).Value = sectionName
    wsDash.Cells(startRow, startCol).Font.Bold = True
    wsDash.Cells(startRow, startCol).Font.Size = 12

    Dim i As Long
    For i = 1 To numCols
        With wsDash.Cells(startRow + 1, startCol + i - 1)
            .Value = "Header " & i
            .Interior.Color = RGB(255, 242, 204) ' Light yellow
            .Borders.Weight = xlThin
            .Font.Bold = True
        End With
    Next i

    MsgBox "Added new section: " & sectionName & " at Row " & startRow, vbInformation
End Sub

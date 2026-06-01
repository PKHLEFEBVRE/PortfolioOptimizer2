Attribute VB_Name = "PositionDashboard"
Option Explicit

Private Const DASH_SHEET As String = "Dashboard"

Public Sub GeneratePositionsDashboard(positionsDict As Object)
    Dim wsDash As Worksheet
    Dim sheetName As String
    sheetName = "Positions Dashboard"

    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If wsDash Is Nothing Then
        Set wsDash = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDash.Name = sheetName
    Else
        wsDash.Cells.Clear
    End If

    If positionsDict.Count = 0 Then Exit Sub

    ' Output to Sheet
    Dim headers As Variant
    headers = Array("Asset Name", "Current Price", "Return (Ann)", "Vol (Ann)", "Sharpe", _
                    "Max Drawdown", "Drawdown Len", "VaR (95%)", "CVaR (95%)", "Semi-Dev", "Downside Beta", _
                    "Conviction (/5)", "Target Price", "TP Proba", "Stop Price", "SP Proba", "Expected Return")

    wsDash.Range(wsDash.Cells(1, 1), wsDash.Cells(1, UBound(headers) + 1)).Value = headers

    Dim outData() As Variant
    ReDim outData(1 To positionsDict.Count, 1 To UBound(headers) + 1)

    Dim rowIdx As Long
    rowIdx = 1
    Dim key As Variant
    For Each key In positionsDict.Keys
        Dim p As PositionCls
        Set p = positionsDict(key)

        outData(rowIdx, 1) = p.AssetName

        Dim pArray() As Double
        pArray = p.Prices
        outData(rowIdx, 2) = pArray(UBound(pArray))

        outData(rowIdx, 3) = p.Metrics.Ret
        outData(rowIdx, 4) = p.Metrics.Vol
        outData(rowIdx, 5) = p.Metrics.Sharpe
        outData(rowIdx, 6) = p.Metrics.MDD
        outData(rowIdx, 7) = p.Metrics.MDDLen
        outData(rowIdx, 8) = p.Metrics.VaR
        outData(rowIdx, 9) = p.Metrics.CVaR
        outData(rowIdx, 10) = p.Metrics.SemiDev
        outData(rowIdx, 11) = p.Metrics.DownsideBeta

        outData(rowIdx, 12) = p.Conviction
        outData(rowIdx, 13) = p.TP
        outData(rowIdx, 14) = p.TPProba
        outData(rowIdx, 15) = p.SP
        outData(rowIdx, 16) = p.SPProba
        outData(rowIdx, 17) = p.ExpectedReturn

        rowIdx = rowIdx + 1
    Next key

    wsDash.Range(wsDash.Cells(2, 1), wsDash.Cells(positionsDict.Count + 1, UBound(headers) + 1)).Value = outData

    ' Format Table
    wsDash.Rows(1).Font.Bold = True
    wsDash.Rows(1).Interior.Color = RGB(220, 230, 241)

    ' Format Percentages
    Dim pctCols As Variant
    pctCols = Array(3, 4, 6, 8, 9, 10, 14, 16, 17)
    Dim c As Variant
    For Each c In pctCols
        wsDash.Range(wsDash.Cells(2, c), wsDash.Cells(positionsDict.Count + 1, c)).NumberFormat = "0.00%"
    Next c

    ' Format 2 decimal places
    Dim numCols As Variant
    numCols = Array(2, 5, 11, 12, 13, 15)
    For Each c In numCols
        wsDash.Range(wsDash.Cells(2, c), wsDash.Cells(positionsDict.Count + 1, c)).NumberFormat = "0.00"
    Next c

    wsDash.Columns.AutoFit
End Sub

Public Sub GeneratePortfoliosDashboard(portfoliosDict As Object)
    Dim wsDash As Worksheet
    Dim sheetName As String
    sheetName = "Positions Dashboard"

    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If wsDash Is Nothing Then Exit Sub ' Should have been created by GeneratePositionsDashboard
    If portfoliosDict.Count = 0 Then Exit Sub

    ' Extract asset names dynamically from the "MEAN" portfolio, which contains all fund assets
    Dim meanPort As SimulatedPortfolioCls
    If Not portfoliosDict.Exists("MEAN") Then Exit Sub
    Set meanPort = portfoliosDict("MEAN")

    Dim posKeys() As Variant
    posKeys = meanPort.GetPositionNames()

    Dim nAssets As Long
    nAssets = UBound(posKeys)

    ' Find the last row of the Positions table
    Dim lastRow As Long
    lastRow = wsDash.Cells(wsDash.Rows.Count, 1).End(xlUp).Row

    Dim startRow As Long
    startRow = lastRow + 3 ' Leave a gap

    ' Build Headers
    Dim headers() As Variant
    ReDim headers(1 To 10 + nAssets)

    headers(1) = "Portfolio Strategy"
    headers(2) = "Return (Ann)"
    headers(3) = "Vol (Ann)"
    headers(4) = "Sharpe"
    headers(5) = "Max Drawdown"
    headers(6) = "Drawdown Len"
    headers(7) = "VaR (95%)"
    headers(8) = "CVaR (95%)"
    headers(9) = "Semi-Dev"
    headers(10) = "Downside Beta"

    Dim i As Long
    For i = 1 To nAssets
        headers(10 + i) = "W_" & CStr(posKeys(i))
    Next i

    wsDash.Range(wsDash.Cells(startRow, 1), wsDash.Cells(startRow, UBound(headers))).Value = headers

    ' Build Data
    Dim outData() As Variant
    ReDim outData(1 To portfoliosDict.Count, 1 To UBound(headers))

    Dim rowIdx As Long
    rowIdx = 1
    Dim key As Variant
    For Each key In portfoliosDict.Keys
        Dim port As SimulatedPortfolioCls
        Set port = portfoliosDict(key)

        outData(rowIdx, 1) = port.StrategyName
        outData(rowIdx, 2) = port.Metrics.Ret
        outData(rowIdx, 3) = port.Metrics.Vol
        outData(rowIdx, 4) = port.Metrics.Sharpe
        outData(rowIdx, 5) = port.Metrics.MDD
        outData(rowIdx, 6) = port.Metrics.MDDLen
        outData(rowIdx, 7) = port.Metrics.VaR
        outData(rowIdx, 8) = port.Metrics.CVaR
        outData(rowIdx, 9) = port.Metrics.SemiDev
        outData(rowIdx, 10) = port.Metrics.DownsideBeta

        For i = 1 To nAssets
            outData(rowIdx, 10 + i) = port.GetWeight(CStr(posKeys(i)))
        Next i

        rowIdx = rowIdx + 1
    Next key

    wsDash.Range(wsDash.Cells(startRow + 1, 1), wsDash.Cells(startRow + portfoliosDict.Count, UBound(headers))).Value = outData

    ' Formatting
    wsDash.Rows(startRow).Font.Bold = True
    wsDash.Rows(startRow).Interior.Color = RGB(220, 230, 241)

    ' Format Percentages
    Dim pctCols As Variant
    pctCols = Array(2, 3, 5, 7, 8, 9)
    Dim c As Variant
    For Each c In pctCols
        wsDash.Range(wsDash.Cells(startRow + 1, c), wsDash.Cells(startRow + portfoliosDict.Count, c)).NumberFormat = "0.00%"
    Next c

    ' Format weight percentages
    wsDash.Range(wsDash.Cells(startRow + 1, 11), wsDash.Cells(startRow + portfoliosDict.Count, UBound(headers))).NumberFormat = "0.00%"

    ' Format Numbers
    Dim numCols As Variant
    numCols = Array(4, 10)
    For Each c In numCols
        wsDash.Range(wsDash.Cells(startRow + 1, c), wsDash.Cells(startRow + portfoliosDict.Count, c)).NumberFormat = "0.00"
    Next c

    wsDash.Columns.AutoFit
    wsDash.Activate
End Sub

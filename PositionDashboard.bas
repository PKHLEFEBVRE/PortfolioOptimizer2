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

    ' --- 1. Output Parameters Table at the top ---
    wsDash.Cells(1, 1).Value = "Parameters"
    wsDash.Cells(1, 1).Font.Bold = True
    wsDash.Cells(1, 1).Interior.Color = RGB(220, 230, 241)

    wsDash.Cells(2, 1).Value = "Risk-Free Rate:"
    wsDash.Cells(2, 2).Value = Sheets(DASH_SHEET).Range("D4").Value
    wsDash.Cells(2, 2).NumberFormat = "0.00%"

    wsDash.Cells(3, 1).Value = "Confidence Level:"
    wsDash.Cells(3, 2).Value = Sheets(DASH_SHEET).Range("D5").Value
    wsDash.Cells(3, 2).NumberFormat = "0.00%"

    Dim startRow As Long
    startRow = 6

    ' --- 2. Output Positions Table ---
    Dim headers As Variant
    headers = Array("Asset Name", "Current Price", "Return (Ann)", "Vol (Ann)", "Sharpe", _
                    "Max Drawdown", "Drawdown Len", "VaR (95%)", "CVaR (95%)", "Semi-Dev", "Downside Beta", _
                    "Conviction (/5)", "Target Price", "TP Proba", "Stop Price", "SP Proba", _
                    "Min Weight", "Max Weight", "Expected Return")

    wsDash.Range(wsDash.Cells(startRow, 1), wsDash.Cells(startRow, UBound(headers) + 1)).Value = headers

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
        outData(rowIdx, 17) = p.MinWeight
        outData(rowIdx, 18) = p.MaxWeight
        outData(rowIdx, 19) = p.ExpectedReturn

        rowIdx = rowIdx + 1
    Next key

    wsDash.Range(wsDash.Cells(startRow + 1, 1), wsDash.Cells(startRow + positionsDict.Count, UBound(headers) + 1)).Value = outData

    ' Format Positions Table
    wsDash.Rows(startRow).Font.Bold = True
    wsDash.Rows(startRow).Interior.Color = RGB(220, 230, 241)

    Dim pctCols As Variant
    pctCols = Array(3, 4, 6, 8, 9, 10, 14, 16, 17, 18, 19)
    Dim c As Variant
    For Each c In pctCols
        wsDash.Range(wsDash.Cells(startRow + 1, c), wsDash.Cells(startRow + positionsDict.Count, c)).NumberFormat = "0.00%"
    Next c

    Dim numCols As Variant
    numCols = Array(2, 5, 11, 12, 13, 15)
    For Each c In numCols
        wsDash.Range(wsDash.Cells(startRow + 1, c), wsDash.Cells(startRow + positionsDict.Count, c)).NumberFormat = "0.00"
    Next c
End Sub

Public Sub GeneratePortfoliosDashboard(portfoliosDict As Object)
    Dim wsDash As Worksheet
    Dim sheetName As String
    sheetName = "Positions Dashboard"

    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If wsDash Is Nothing Then Exit Sub
    If portfoliosDict.Count = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = wsDash.Cells(wsDash.Rows.Count, 1).End(xlUp).Row

    ' --- 3. Output Portfolios Metrics & Factors Table ---
    Dim mStartRow As Long
    mStartRow = lastRow + 4 ' Leave a gap

    Dim mHeaders As Variant
    mHeaders = Array("Portfolio Strategy", "Return (Ann)", "Vol (Ann)", "Sharpe", _
                     "Max Drawdown", "Drawdown Len", "VaR (95%)", "CVaR (95%)", "Semi-Dev", "Downside Beta", _
                     "CVaR Factor", "Drawdown Factor", "DS Beta Factor", "Semi-Dev Factor", "Final Risk Overlay")

    wsDash.Range(wsDash.Cells(mStartRow, 1), wsDash.Cells(mStartRow, UBound(mHeaders) + 1)).Value = mHeaders

    Dim mData() As Variant
    ReDim mData(1 To portfoliosDict.Count, 1 To UBound(mHeaders) + 1)

    Dim rowIdx As Long
    rowIdx = 1
    Dim key As Variant
    For Each key In portfoliosDict.Keys
        Dim port As SimulatedPortfolioCls
        Set port = portfoliosDict(key)

        mData(rowIdx, 1) = port.StrategyName
        mData(rowIdx, 2) = port.Metrics.Ret
        mData(rowIdx, 3) = port.Metrics.Vol
        mData(rowIdx, 4) = port.Metrics.Sharpe
        mData(rowIdx, 5) = port.Metrics.MDD
        mData(rowIdx, 6) = port.Metrics.MDDLen
        mData(rowIdx, 7) = port.Metrics.VaR
        mData(rowIdx, 8) = port.Metrics.CVaR
        mData(rowIdx, 9) = port.Metrics.SemiDev
        mData(rowIdx, 10) = port.Metrics.DownsideBeta

        If port.FinalRiskFactor > 0 Then
            mData(rowIdx, 11) = port.CVaRFactor
            mData(rowIdx, 12) = port.HistoricalMDDFactor
            mData(rowIdx, 13) = port.DDFactor
            mData(rowIdx, 14) = port.SemiDevFactor
            mData(rowIdx, 15) = port.FinalRiskFactor
        Else
            mData(rowIdx, 11) = "N/A"
            mData(rowIdx, 12) = "N/A"
            mData(rowIdx, 13) = "N/A"
            mData(rowIdx, 14) = "N/A"
            mData(rowIdx, 15) = "N/A"
        End If

        rowIdx = rowIdx + 1
    Next key

    wsDash.Range(wsDash.Cells(mStartRow + 1, 1), wsDash.Cells(mStartRow + portfoliosDict.Count, UBound(mHeaders) + 1)).Value = mData

    ' Format Metrics Table
    wsDash.Rows(mStartRow).Font.Bold = True
    wsDash.Rows(mStartRow).Interior.Color = RGB(220, 230, 241)

    Dim mPctCols As Variant
    mPctCols = Array(2, 3, 5, 7, 8, 9, 11, 12, 13, 14, 15)
    Dim c As Variant
    For Each c In mPctCols
        wsDash.Range(wsDash.Cells(mStartRow + 1, c), wsDash.Cells(mStartRow + portfoliosDict.Count, c)).NumberFormat = "0.00%"
    Next c

    Dim mNumCols As Variant
    mNumCols = Array(4, 10)
    For Each c In mNumCols
        wsDash.Range(wsDash.Cells(mStartRow + 1, c), wsDash.Cells(mStartRow + portfoliosDict.Count, c)).NumberFormat = "0.00"
    Next c


    ' --- 4. Output Portfolios Weights Table ---
    lastRow = wsDash.Cells(wsDash.Rows.Count, 1).End(xlUp).Row
    Dim wStartRow As Long
    wStartRow = lastRow + 4 ' Leave a gap

    Dim meanPort As SimulatedPortfolioCls
    If Not portfoliosDict.Exists("MEAN") Then Exit Sub
    Set meanPort = portfoliosDict("MEAN")

    Dim posKeys() As Variant
    posKeys = meanPort.GetPositionNames()
    Dim nAssets As Long
    nAssets = UBound(posKeys)

    Dim wHeaders() As Variant
    ReDim wHeaders(1 To 1 + nAssets)
    wHeaders(1) = "Portfolio Strategy"

    Dim i As Long
    For i = 1 To nAssets
        wHeaders(1 + i) = "W_" & CStr(posKeys(i))
    Next i

    wsDash.Range(wsDash.Cells(wStartRow, 1), wsDash.Cells(wStartRow, UBound(wHeaders))).Value = wHeaders

    Dim wData() As Variant
    ReDim wData(1 To portfoliosDict.Count, 1 To UBound(wHeaders))

    rowIdx = 1
    For Each key In portfoliosDict.Keys
        Set port = portfoliosDict(key)

        wData(rowIdx, 1) = port.StrategyName

        For i = 1 To nAssets
            wData(rowIdx, 1 + i) = port.GetWeight(CStr(posKeys(i)))
        Next i

        rowIdx = rowIdx + 1
    Next key

    wsDash.Range(wsDash.Cells(wStartRow + 1, 1), wsDash.Cells(wStartRow + portfoliosDict.Count, UBound(wHeaders))).Value = wData

    ' Format Weights Table
    wsDash.Rows(wStartRow).Font.Bold = True
    wsDash.Rows(wStartRow).Interior.Color = RGB(220, 230, 241)

    wsDash.Range(wsDash.Cells(wStartRow + 1, 2), wsDash.Cells(wStartRow + portfoliosDict.Count, UBound(wHeaders))).NumberFormat = "0.00%"

    wsDash.Columns.AutoFit
    wsDash.Activate
End Sub

Public Sub GetDashboardInputs(masterPositions As Object)
    Dim wsDash As Worksheet
    Dim sheetName As String
    sheetName = "Positions Dashboard"

    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If wsDash Is Nothing Then Exit Sub ' Dashboard hasn't been generated yet

    Dim lastRow As Long
    lastRow = wsDash.Cells(wsDash.Rows.Count, 1).End(xlUp).Row
    If lastRow < 7 Then Exit Sub ' Table is empty

    ' Headers are at Row 6
    ' Asset Name is Col 1.
    ' Overrides: Col 12 (Conviction), 13 (TP), 14 (TP Proba), 15 (SP), 16 (SP Proba), 17 (Min), 18 (Max)

    Dim r As Long
    For r = 7 To lastRow
        Dim aName As String
        aName = wsDash.Cells(r, 1).Value
        If aName = "" Then Exit For

        If masterPositions.Exists(aName) Then
            Dim pos As PositionCls
            Set pos = masterPositions(aName)

            ' Only update if the user typed something
            If Not IsEmpty(wsDash.Cells(r, 12).Value) Then pos.Conviction = CDbl(wsDash.Cells(r, 12).Value)
            If Not IsEmpty(wsDash.Cells(r, 13).Value) Then pos.TP = CDbl(wsDash.Cells(r, 13).Value)
            If Not IsEmpty(wsDash.Cells(r, 14).Value) Then pos.TPProba = CDbl(wsDash.Cells(r, 14).Value)
            If Not IsEmpty(wsDash.Cells(r, 15).Value) Then pos.SP = CDbl(wsDash.Cells(r, 15).Value)
            If Not IsEmpty(wsDash.Cells(r, 16).Value) Then pos.SPProba = CDbl(wsDash.Cells(r, 16).Value)

            If Not IsEmpty(wsDash.Cells(r, 17).Value) Then pos.MinWeight = CDbl(wsDash.Cells(r, 17).Value)
            If Not IsEmpty(wsDash.Cells(r, 18).Value) Then pos.MaxWeight = CDbl(wsDash.Cells(r, 18).Value)
        End If
    Next r
End Sub

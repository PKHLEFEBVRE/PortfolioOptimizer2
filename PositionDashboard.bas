Attribute VB_Name = "PositionDashboard"
Option Explicit

Public Sub GeneratePositionsDashboard(positionsDict As Object)
    Dim wsDash As Worksheet
    Dim sheetName As String
    sheetName = "Positions Dashboard"

    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ' If the sheet doesn't exist, create it and set default parameters
    If wsDash Is Nothing Then
        Set wsDash = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDash.Name = sheetName

        ' Set default parameters immediately upon creation so they aren't blank
        wsDash.Cells(2, 2).Value = 0.02   ' 2% Risk Free
        wsDash.Cells(3, 2).Value = 0.95   ' 95% Confidence
    End If

    If positionsDict.Count = 0 Then Exit Sub

    ' --- 1. Output Parameters Table at the top ---
    wsDash.Cells(1, 1).Value = "PARAMETERS"
    wsDash.Cells(1, 1).Font.Bold = True
    wsDash.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsDash.Cells(1, 1).Interior.Color = RGB(31, 73, 125) ' Dark Blue
    wsDash.Cells(1, 2).Interior.Color = RGB(31, 73, 125)

    wsDash.Cells(2, 1).Value = "Risk-Free Rate:"
    wsDash.Cells(2, 1).Font.Bold = True
    wsDash.Cells(2, 2).NumberFormat = "0.00%"

    wsDash.Cells(3, 1).Value = "Confidence Level:"
    wsDash.Cells(3, 1).Font.Bold = True
    wsDash.Cells(3, 2).NumberFormat = "0.00%"

    wsDash.Cells(4, 1).Value = "Start Date:"
    wsDash.Cells(4, 1).Font.Bold = True
    If IsEmpty(wsDash.Cells(4, 2).Value) Then wsDash.Cells(4, 2).Value = Date - 365 * 3
    wsDash.Cells(4, 2).NumberFormat = "dd/mm/yyyy"

    wsDash.Cells(5, 1).Value = "Portfolio ID (VIA):"
    wsDash.Cells(5, 1).Font.Bold = True
    If IsEmpty(wsDash.Cells(5, 2).Value) Then wsDash.Cells(5, 2).Value = "Enter ID Here"

    ' Draw borders around Parameters
    wsDash.Range("A1:B5").Borders.LineStyle = xlContinuous


    ' --- 2. Build Positions Table ---
    Dim startRow As Long
    startRow = 8

    wsDash.Cells(startRow - 2, 1).Value = "ASSET POSITIONS"
    wsDash.Cells(startRow - 2, 1).Font.Bold = True
    wsDash.Cells(startRow - 2, 1).Font.Size = 14

    ' Super Headers (Center Across Selection)
    wsDash.Range(wsDash.Cells(startRow - 1, 1), wsDash.Cells(startRow - 1, 2)).Value = "" ' Identifiers

    wsDash.Range(wsDash.Cells(startRow - 1, 3), wsDash.Cells(startRow - 1, 6)).Value = ""
    wsDash.Cells(startRow - 1, 3).Value = "PERFORMANCE"
    wsDash.Range(wsDash.Cells(startRow - 1, 3), wsDash.Cells(startRow - 1, 6)).HorizontalAlignment = xlCenterAcrossSelection
    wsDash.Range(wsDash.Cells(startRow - 1, 3), wsDash.Cells(startRow - 1, 6)).Font.Bold = True
    wsDash.Range(wsDash.Cells(startRow - 1, 3), wsDash.Cells(startRow - 1, 6)).Interior.Color = RGB(220, 230, 241)

    wsDash.Range(wsDash.Cells(startRow - 1, 7), wsDash.Cells(startRow - 1, 11)).Value = ""
    wsDash.Cells(startRow - 1, 7).Value = "RISK METRICS"
    wsDash.Range(wsDash.Cells(startRow - 1, 7), wsDash.Cells(startRow - 1, 11)).HorizontalAlignment = xlCenterAcrossSelection
    wsDash.Range(wsDash.Cells(startRow - 1, 7), wsDash.Cells(startRow - 1, 11)).Font.Bold = True
    wsDash.Range(wsDash.Cells(startRow - 1, 7), wsDash.Cells(startRow - 1, 11)).Interior.Color = RGB(253, 233, 217) ' Light Orange

    wsDash.Range(wsDash.Cells(startRow - 1, 12), wsDash.Cells(startRow - 1, 19)).Value = ""
    wsDash.Cells(startRow - 1, 12).Value = "USER CONSTRAINTS & EXPECTATIONS"
    wsDash.Range(wsDash.Cells(startRow - 1, 12), wsDash.Cells(startRow - 1, 19)).HorizontalAlignment = xlCenterAcrossSelection
    wsDash.Range(wsDash.Cells(startRow - 1, 12), wsDash.Cells(startRow - 1, 19)).Font.Bold = True
    wsDash.Range(wsDash.Cells(startRow - 1, 12), wsDash.Cells(startRow - 1, 19)).Interior.Color = RGB(235, 241, 222) ' Light Green

    ' Main Headers
    Dim headers As Variant
    headers = Array("Asset Name", "Current Price", "Return (Ann)", "Vol (Ann)", "Sharpe", "Expected Return", _
                    "Max Drawdown", "Drawdown Len", "VaR (95%)", "CVaR (95%)", "Semi-Dev", "Downside Beta", _
                    "Conviction (/5)", "Target Price", "TP Proba", "Stop Price", "SP Proba", _
                    "Min Weight", "Max Weight")

    wsDash.Range(wsDash.Cells(startRow, 1), wsDash.Cells(startRow, UBound(headers) + 1)).Value = headers
    wsDash.Range(wsDash.Cells(startRow, 1), wsDash.Cells(startRow, UBound(headers) + 1)).Font.Bold = True
    wsDash.Range(wsDash.Cells(startRow, 1), wsDash.Cells(startRow, UBound(headers) + 1)).Font.Color = RGB(255, 255, 255)
    wsDash.Range(wsDash.Cells(startRow, 1), wsDash.Cells(startRow, UBound(headers) + 1)).Interior.Color = RGB(31, 73, 125)
    wsDash.Range(wsDash.Cells(startRow, 1), wsDash.Cells(startRow, UBound(headers) + 1)).HorizontalAlignment = xlCenter

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
        outData(rowIdx, 6) = p.ExpectedReturn

        outData(rowIdx, 7) = p.Metrics.MDD
        outData(rowIdx, 8) = p.Metrics.MDDLen
        outData(rowIdx, 9) = p.Metrics.VaR
        outData(rowIdx, 10) = p.Metrics.CVaR
        outData(rowIdx, 11) = p.Metrics.SemiDev
        outData(rowIdx, 12) = p.Metrics.DownsideBeta

        outData(rowIdx, 13) = p.Conviction
        outData(rowIdx, 14) = p.TP
        outData(rowIdx, 15) = p.TPProba
        outData(rowIdx, 16) = p.SP
        outData(rowIdx, 17) = p.SPProba
        outData(rowIdx, 18) = p.MinWeight
        outData(rowIdx, 19) = p.MaxWeight

        rowIdx = rowIdx + 1
    Next key

    wsDash.Range(wsDash.Cells(startRow + 1, 1), wsDash.Cells(startRow + positionsDict.Count, UBound(headers) + 1)).Value = outData

    ' Draw Borders
    Dim dataRange As Range
    Set dataRange = wsDash.Range(wsDash.Cells(startRow - 1, 1), wsDash.Cells(startRow + positionsDict.Count, UBound(headers) + 1))
    dataRange.Borders.LineStyle = xlContinuous
    dataRange.HorizontalAlignment = xlCenter

    ' Format Percentages
    Dim pctCols As Variant
    pctCols = Array(3, 4, 6, 7, 9, 10, 11, 15, 17, 18, 19)
    Dim c As Variant
    For Each c In pctCols
        wsDash.Range(wsDash.Cells(startRow + 1, c), wsDash.Cells(startRow + positionsDict.Count, c)).NumberFormat = "0.00%"
    Next c

    ' Format 2 decimal places
    Dim numCols As Variant
    numCols = Array(2, 5, 12, 13, 14, 16)
    For Each c In numCols
        wsDash.Range(wsDash.Cells(startRow + 1, c), wsDash.Cells(startRow + positionsDict.Count, c)).NumberFormat = "0.00"
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

    If wsDash Is Nothing Then Exit Sub
    If portfoliosDict.Count = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = wsDash.Cells(wsDash.Rows.Count, 1).End(xlUp).Row

    ' --- 3. Output Portfolios Metrics Table ---
    Dim mStartRow As Long
    mStartRow = lastRow + 5

    wsDash.Cells(mStartRow - 2, 1).Value = "SIMULATED PORTFOLIOS"
    wsDash.Cells(mStartRow - 2, 1).Font.Bold = True
    wsDash.Cells(mStartRow - 2, 1).Font.Size = 14

    ' Super Headers
    wsDash.Range(wsDash.Cells(mStartRow - 1, 1), wsDash.Cells(mStartRow - 1, 1)).Value = ""
    wsDash.Cells(mStartRow - 1, 2).Value = "PERFORMANCE"
    wsDash.Range(wsDash.Cells(mStartRow - 1, 2), wsDash.Cells(mStartRow - 1, 4)).HorizontalAlignment = xlCenterAcrossSelection
    wsDash.Range(wsDash.Cells(mStartRow - 1, 2), wsDash.Cells(mStartRow - 1, 4)).Font.Bold = True
    wsDash.Range(wsDash.Cells(mStartRow - 1, 2), wsDash.Cells(mStartRow - 1, 4)).Interior.Color = RGB(220, 230, 241)

    wsDash.Cells(mStartRow - 1, 5).Value = "RISK METRICS & LIMITS"
    wsDash.Range(wsDash.Cells(mStartRow - 1, 5), wsDash.Cells(mStartRow - 1, 15)).HorizontalAlignment = xlCenterAcrossSelection
    wsDash.Range(wsDash.Cells(mStartRow - 1, 5), wsDash.Cells(mStartRow - 1, 15)).Font.Bold = True
    wsDash.Range(wsDash.Cells(mStartRow - 1, 5), wsDash.Cells(mStartRow - 1, 15)).Interior.Color = RGB(253, 233, 217)

    Dim mHeaders As Variant
    mHeaders = Array("Portfolio Strategy", "Return (Ann)", "Vol (Ann)", "Sharpe", _
                     "Max Drawdown", "Drawdown Len", "VaR (95%)", "CVaR (95%)", "Semi-Dev", "Downside Beta", _
                     "CVaR Factor", "MDD Factor", "DS Beta Factor", "Semi-Dev Factor", "Final Risk Overlay")

    wsDash.Range(wsDash.Cells(mStartRow, 1), wsDash.Cells(mStartRow, UBound(mHeaders) + 1)).Value = mHeaders
    wsDash.Range(wsDash.Cells(mStartRow, 1), wsDash.Cells(mStartRow, UBound(mHeaders) + 1)).Font.Bold = True
    wsDash.Range(wsDash.Cells(mStartRow, 1), wsDash.Cells(mStartRow, UBound(mHeaders) + 1)).Font.Color = RGB(255, 255, 255)
    wsDash.Range(wsDash.Cells(mStartRow, 1), wsDash.Cells(mStartRow, UBound(mHeaders) + 1)).Interior.Color = RGB(31, 73, 125)

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
    Dim mRange As Range
    Set mRange = wsDash.Range(wsDash.Cells(mStartRow - 1, 1), wsDash.Cells(mStartRow + portfoliosDict.Count, UBound(mHeaders) + 1))
    mRange.Borders.LineStyle = xlContinuous
    mRange.HorizontalAlignment = xlCenter

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
    wStartRow = lastRow + 5

    wsDash.Cells(wStartRow - 2, 1).Value = "TARGET WEIGHT ALLOCATIONS"
    wsDash.Cells(wStartRow - 2, 1).Font.Bold = True
    wsDash.Cells(wStartRow - 2, 1).Font.Size = 14

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
        wHeaders(1 + i) = CStr(posKeys(i))
    Next i

    wsDash.Range(wsDash.Cells(wStartRow, 1), wsDash.Cells(wStartRow, UBound(wHeaders))).Value = wHeaders
    wsDash.Range(wsDash.Cells(wStartRow, 1), wsDash.Cells(wStartRow, UBound(wHeaders))).Font.Bold = True
    wsDash.Range(wsDash.Cells(wStartRow, 1), wsDash.Cells(wStartRow, UBound(wHeaders))).Font.Color = RGB(255, 255, 255)
    wsDash.Range(wsDash.Cells(wStartRow, 1), wsDash.Cells(wStartRow, UBound(wHeaders))).Interior.Color = RGB(31, 73, 125)

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
    Dim wRange As Range
    Set wRange = wsDash.Range(wsDash.Cells(wStartRow, 1), wsDash.Cells(wStartRow + portfoliosDict.Count, UBound(wHeaders)))
    wRange.Borders.LineStyle = xlContinuous
    wRange.HorizontalAlignment = xlCenter

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
    If lastRow < 8 Then Exit Sub ' Table is empty

    ' User editable columns in the newly formatted table:
    ' 13 = Conviction
    ' 14 = Target Price
    ' 15 = TP Proba
    ' 16 = Stop Price
    ' 17 = SP Proba
    ' 18 = Min Weight
    ' 19 = Max Weight

    Dim r As Long
    For r = 9 To lastRow
        Dim aName As String
        aName = wsDash.Cells(r, 1).Value
        If aName = "" Then Exit For

        If masterPositions.Exists(aName) Then
            Dim pos As PositionCls
            Set pos = masterPositions(aName)

            ' Only update if the user typed something
            If IsNumeric(wsDash.Cells(r, 13).Value) And Not IsEmpty(wsDash.Cells(r, 13).Value) Then pos.Conviction = CDbl(wsDash.Cells(r, 13).Value)
            If IsNumeric(wsDash.Cells(r, 14).Value) And Not IsEmpty(wsDash.Cells(r, 14).Value) Then pos.TP = CDbl(wsDash.Cells(r, 14).Value)
            If IsNumeric(wsDash.Cells(r, 15).Value) And Not IsEmpty(wsDash.Cells(r, 15).Value) Then pos.TPProba = CDbl(wsDash.Cells(r, 15).Value)
            If IsNumeric(wsDash.Cells(r, 16).Value) And Not IsEmpty(wsDash.Cells(r, 16).Value) Then pos.SP = CDbl(wsDash.Cells(r, 16).Value)
            If IsNumeric(wsDash.Cells(r, 17).Value) And Not IsEmpty(wsDash.Cells(r, 17).Value) Then pos.SPProba = CDbl(wsDash.Cells(r, 17).Value)

            If IsNumeric(wsDash.Cells(r, 18).Value) And Not IsEmpty(wsDash.Cells(r, 18).Value) Then pos.MinWeight = CDbl(wsDash.Cells(r, 18).Value)
            If IsNumeric(wsDash.Cells(r, 19).Value) And Not IsEmpty(wsDash.Cells(r, 19).Value) Then pos.MaxWeight = CDbl(wsDash.Cells(r, 19).Value)
        End If
    Next r
End Sub

Attribute VB_Name = "PositionDashboard"
Option Explicit

Private Const DASH_SHEET As String = "Dashboard"

Public Sub GeneratePositionsDashboard()
    Application.ScreenUpdating = False

    ' 1. Fetch Historical Data
    Dim prices() As Variant, dates() As Date, assetNames() As String
    Call assetUtils.GetHistoricalDataDatesAndNames(prices, dates, assetNames)
    If UBound(prices, 2) < 1 Then Exit Sub

    ' 2. Fetch Convictions
    Dim convictions As Object ' Dictionary
    Set convictions = convictionsUtils.getConvictions()

    ' Fetch config for risk free rate and conf level
    Dim rf As Double
    rf = Sheets(DASH_SHEET).Range("D4").Value
    Dim conf As Double
    conf = Sheets(DASH_SHEET).Range("D5").Value

    ' 3. Calculate Mean Returns for Expected Return
    Dim logRets() As Double, meanRets() As Double
    Call assetUtils.CalculateHistoricalStats(dates, prices, logRets, meanRets)
    Dim expReturns() As Double
    Dim lastDate As Date
    lastDate = dates(UBound(dates))
    expReturns = assetUtils.CalculateExpectedReturns(meanRets, lastDate)

    Dim positionsDict As Object
    Set positionsDict = CreateObject("Scripting.Dictionary")

    Dim nDays As Long, nAssets As Long
    nDays = UBound(prices, 1)
    nAssets = UBound(prices, 2)

    Dim j As Long, i As Long

    ' 4. Instantiate PositionCls for each asset
    For j = 1 To nAssets
        Dim pos As PositionCls
        Set pos = New PositionCls

        pos.AssetName = assetNames(j)

        ' Arrays for class
        Dim pPrices() As Double, pDates() As Date
        ReDim pPrices(1 To nDays)
        ReDim pDates(1 To nDays)
        For i = 1 To nDays
            pPrices(i) = CDbl(prices(i, j))
            pDates(i) = dates(i)
        Next i

        pos.InitializeData pPrices, pDates
        pos.ComputeMetrics rf, conf

        ' Fetch convictions
        If convictions.Exists(pos.AssetName) Then
            Dim cDict As Object
            Set cDict = convictions(pos.AssetName)
            On Error Resume Next
            pos.Conviction = cDict("conviction")
            pos.TP = cDict("TP")
            pos.TPProba = cDict("TP proba")
            pos.SP = cDict("SP")
            pos.SPProba = cDict("SP proba")
            On Error GoTo 0
        End If

        ' Expected Return
        pos.ExpectedReturn = expReturns(j)

        positionsDict.Add pos.AssetName, pos
    Next j

    ' 5. Create or Clear Dashboard Sheet
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

    ' 6. Output to Sheet
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
        outData(rowIdx, 2) = p.Prices(UBound(p.Prices))
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

    ' 7. Format Table
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
    wsDash.Activate

    Application.ScreenUpdating = True
    MsgBox "Positions Dashboard generated successfully!", vbInformation
End Sub

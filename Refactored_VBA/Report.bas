Option Explicit
Option Base 1

Private dictCols As Object

Sub ReportWorkflow()
    InitializeGlobals

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")

    Dim firstRange As Range, secondRange As Range
    Set firstRange = ws.Range( _
                    ws.Cells(ws.Range("InputTableStart").Row, ws.Range("InputTableStart").Column - 1), _
                    ws.Cells(ws.Range("InputTableStart").End(xlDown).Row, ws.Range("InputTableStart").End(xlToRight).Column - 1) _
                )
    Set secondRange = ws.Range( _
                    ws.Cells(ws.Range("StrategyWeightsStart").Row, ws.Range("StrategyWeightsStart").Column - 1), _
                    ws.Cells(ws.Range("StrategyWeightsStart").End(xlDown).Row, ws.Range("StrategyWeightsStart").End(xlToRight).Column) _
                )

    Dim assetsData As Dictionary
    Set assetsData = reportUtils.CreateCombinedAssetDictionary(firstRange, secondRange)

    reportUtils.SaveResultsToTop assetsData, "HistoricalSimulations"

    GeneratePortfolioReport

End Sub

Private Sub GeneratePortfolioReport()
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim data() As Variant
    Dim rRow As Long
    Dim maxDate As Date
    Dim d1 As Date, d2 As Date, d7 As Date, d14 As Date

    ' Set up sheets
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("HistoricalSimulations")
    If wsData Is Nothing Then
        MsgBox "HistoricalSimulations sheet not found!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Create or clear report sheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Portfolio Report")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsReport.Name = "Portfolio Report"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0

    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 1 Then
        MsgBox "No data found in HistoricalSimulations.", vbExclamation
        Exit Sub
    End If

    ' Initialize Column Mappings
    If Not MapColumns(wsData, lastCol) Then
        MsgBox "Failed to map all required columns. Check headers in HistoricalSimulations.", vbCritical
        Exit Sub
    End If

    ' Read Data (Excluding headers to match our logic, 1-based index)
    data = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRow, lastCol)).Value

    ' 1. Determine Dates
    maxDate = GetMaxDate(data)
    If maxDate = 0 Then
        MsgBox "No valid dates found in the data.", vbExclamation
        Exit Sub
    End If

    d1 = FindClosestDate(data, maxDate - 1)
    d2 = FindClosestDate(data, maxDate - 2)
    d7 = FindClosestDate(data, maxDate - 7)
    d14 = FindClosestDate(data, maxDate - 14)

    rRow = 1

    ' Format Title
    wsReport.Cells(rRow, 1).Value = "Portfolio Report - " & Format(maxDate, "dd/mm/yyyy")
    wsReport.Cells(rRow, 1).Font.Bold = True
    wsReport.Cells(rRow, 1).Font.Size = 16
    rRow = rRow + 2

    ' Dictionary to hold historical data for fast lookup: Dict(Asset_ID & "_" & Date) = RowIndex
    Dim dictHist As Object
    Set dictHist = BuildHistoryDictionary(data)

    ' Process Sections
    rRow = ProcessRiskProximity(wsReport, data, maxDate, rRow)
    rRow = ProcessAllocationGaps(wsReport, data, dictHist, maxDate, d1, d2, d7, d14, rRow)
    rRow = ProcessHistoricalChanges(wsReport, data, dictHist, maxDate, d1, d2, d7, d14, rRow)
    rRow = ProcessConvictionAnalysis(wsReport, data, maxDate, rRow)
    rRow = ProcessSchemeVariations(wsReport, data, dictHist, maxDate, d1, d2, d7, d14, rRow)

    ' General Formatting
    wsReport.Columns("A:Z").AutoFit

    ' Freeze Panes
    wsReport.Activate
    ActiveWindow.FreezePanes = False
    wsReport.Range("A3").Select
    ActiveWindow.FreezePanes = True
    wsReport.Range("A1").Select
    wsReport.Range("A1").Columns.ColumnWidth = 10

    'MsgBox "Portfolio Report Generated Successfully!", vbInformation
End Sub

Private Function MapColumns(ByVal ws As Worksheet, ByVal lastCol As Long) As Boolean
    Dim i As Long
    Dim header As String

    Set dictCols = CreateObject("Scripting.Dictionary")

    For i = 1 To lastCol
        header = UCase(Trim(ws.Cells(1, i).Value))
        Select Case header
            Case "ASSET_ID": dictCols("ASSET_ID") = i
            Case "DATE_SAVED": dictCols("DATE_SAVED") = i
            Case "ASSET NAME": dictCols("ASSET_NAME") = i
            Case "CURRENT PRICE": dictCols("CURRENT_PRICE") = i
            Case "CONVICTION (/5)": dictCols("CONVICTION") = i
            Case "TARGET PRICE": dictCols("TARGET_PRICE") = i
            Case "STOP PRICE": dictCols("STOP_PRICE") = i
            Case "EXPECTED RETURN": dictCols("EXPECTED_RETURN") = i
            Case "ERC UNCSTRD": dictCols("ERC") = i
            Case "ER/VOL": dictCols("ERVOL") = i
            Case "SHARPE": dictCols("SHARPE") = i
            Case "CUSTOM": dictCols("CUSTOM") = i
            Case "MEAN": dictCols("MEAN") = i
            Case "CURRENT": dictCols("CURRENT_ALLOC") = i
        End Select
    Next i

    ' Validate required columns are found
    If Not dictCols.Exists("ASSET_ID") Or Not dictCols.Exists("DATE_SAVED") Or _
       Not dictCols.Exists("ASSET_NAME") Or Not dictCols.Exists("CURRENT_PRICE") Or _
       Not dictCols.Exists("CONVICTION") Or Not dictCols.Exists("TARGET_PRICE") Or _
       Not dictCols.Exists("STOP_PRICE") Or Not dictCols.Exists("EXPECTED_RETURN") Or _
       Not dictCols.Exists("ERC") Or Not dictCols.Exists("ERVOL") Or _
       Not dictCols.Exists("SHARPE") Or Not dictCols.Exists("CUSTOM") Or _
       Not dictCols.Exists("MEAN") Or Not dictCols.Exists("CURRENT_ALLOC") Then
        MapColumns = False
    Else
        MapColumns = True
    End If
End Function

Private Sub FormatTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal nRows As Long, ByVal nCols As Long)
    If nRows < 0 Then nRows = 0
    Dim rngHead As Range
    Set rngHead = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, nCols))
    Dim rngFull As Range
    Set rngFull = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + nRows, nCols))

    rngFull.Borders.LineStyle = 1 ' xlContinuous
    rngFull.Borders.Color = RGB(200, 200, 200)
    rngFull.Borders.Weight = 2 ' xlThin

    rngHead.Interior.Color = RGB(220, 230, 241) ' Light blue background
    rngHead.Font.Bold = True
End Sub

Private Sub FormatDiffColor(ByVal ws As Worksheet, ByVal startRow As Long, ByVal nRows As Long, ByVal colIdx As Long)
    If nRows < 1 Then Exit Sub
    Dim cell As Range
    For Each cell In ws.Range(ws.Cells(startRow + 1, colIdx), ws.Cells(startRow + nRows, colIdx))
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            If cell.Value > 0.0001 Then
                cell.Font.Color = RGB(0, 150, 0)
            ElseIf cell.Value < -0.0001 Then
                cell.Font.Color = RGB(200, 0, 0)
            End If
        End If
    Next cell
End Sub

Private Function GetMaxDate(ByRef data() As Variant) As Date
    Dim maxDate As Date
    Dim i As Long
    maxDate = 0
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) > maxDate Then maxDate = data(i, dictCols("DATE_SAVED"))
        End If
    Next i
    GetMaxDate = maxDate
End Function

Private Function FindClosestDate(ByRef data() As Variant, targetDate As Date) As Date
    Dim i As Long
    Dim closestDate As Date
    Dim diff As Long
    Dim minDiff As Long
    minDiff = 999999

    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            diff = targetDate - data(i, dictCols("DATE_SAVED"))
            If diff >= 0 And diff < minDiff Then
                minDiff = diff
                closestDate = data(i, dictCols("DATE_SAVED"))
            End If
        End If
    Next i

    FindClosestDate = closestDate
End Function

Private Function BuildHistoryDictionary(ByRef data() As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            dict(data(i, dictCols("ASSET_ID")) & "_" & Format(data(i, dictCols("DATE_SAVED")), "dd/mm/yyyy")) = i
        End If
    Next i
    Set BuildHistoryDictionary = dict
End Function


Private Function SafeCDbl(ByVal val As Variant) As Double
    If IsNumeric(val) Then
        SafeCDbl = CDbl(val)
    Else
        SafeCDbl = 0
    End If
End Function

Private Function SafeGetVal(ByRef data() As Variant, ByVal idx As Variant, ByVal col As Long) As Variant
    If IsEmpty(idx) Then
        SafeGetVal = "N/A"
    ElseIf idx = 0 Then
        SafeGetVal = "N/A"
    Else
        SafeGetVal = SafeCDbl(data(idx, col))
    End If
End Function

Private Function ProcessRiskProximity(ByVal ws As Worksheet, ByRef data() As Variant, maxDate As Date, startRow As Long) As Long
    ws.Cells(startRow, 1).Value = "Section 1: Risk & Target Proximity (Most Recent Data)"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 14
    startRow = startRow + 1

    Dim headers As Variant
    headers = Array("Asset ID", "Latest Name", "Current Price", "Target Price", "Stop Price", "Status", "Deviation")
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, UBound(headers) + 1)).Value = headers

    Dim outData() As Variant
    ReDim outData(1 To UBound(data, 1), 1 To 7)
    Dim outCount As Long
    outCount = 0

    Dim i As Long
    Dim cp As Double, tp As Double, sp As Double
    Dim diffT As Double, diffS As Double
    Dim status As String, devVal As Variant

    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) = maxDate Then
                cp = SafeCDbl(data(i, dictCols("CURRENT_PRICE")))
                tp = SafeCDbl(data(i, dictCols("TARGET_PRICE")))
                sp = SafeCDbl(data(i, dictCols("STOP_PRICE")))

                If cp > 0 Then
                    diffT = 999: diffS = 999
                    If tp > 0 Then diffT = Abs(cp - tp) / tp
                    If sp > 0 Then diffS = Abs(cp - sp) / sp

                    status = ""
                    devVal = ""

                    If tp > 0 And tp < cp Then
                        status = "WARNING: Target < Current"
                        devVal = diffT
                    ElseIf tp > 0 And diffT <= 0.1 Then
                        status = "Near Target (10%)"
                        devVal = diffT
                    ElseIf tp > 0 And diffT <= 0.25 Then
                        status = "Near Target (25%)"
                        devVal = diffT
                    ElseIf sp > 0 And diffS <= 0.1 Then
                        status = "Near Stop (10%)"
                        devVal = diffS
                    ElseIf sp > 0 And diffS <= 0.25 Then
                        status = "Near Stop (25%)"
                        devVal = diffS
                    End If

                    If status <> "" Then
                        outCount = outCount + 1
                        outData(outCount, 1) = data(i, dictCols("ASSET_ID"))
                        outData(outCount, 2) = data(i, dictCols("ASSET_NAME"))
                        outData(outCount, 3) = cp
                        outData(outCount, 4) = tp
                        outData(outCount, 5) = sp
                        outData(outCount, 6) = status
                        outData(outCount, 7) = devVal
                    End If
                End If
            End If
        End If
    Next i

    If outCount > 0 Then
        ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + outCount, 7)).Value = outData

        Call FormatTable(ws, startRow, outCount, 7)
        ws.Range(ws.Cells(startRow + 1, 7), ws.Cells(startRow + outCount, 7)).NumberFormat = "0.0%"

        Dim r As Long
        For r = 1 To outCount
            If InStr(outData(r, 6), "WARNING") > 0 Then
                ws.Cells(startRow + r, 6).Font.Color = RGB(255, 0, 0)
                ws.Cells(startRow + r, 6).Font.Bold = True
                ws.Cells(startRow + r, 7).Font.Color = RGB(255, 0, 0)
            ElseIf InStr(outData(r, 6), "Target") > 0 Then
                ws.Cells(startRow + r, 6).Font.Color = RGB(0, 150, 0)
                ws.Cells(startRow + r, 7).Font.Color = RGB(0, 150, 0)
            ElseIf InStr(outData(r, 6), "Stop") > 0 Then
                ws.Cells(startRow + r, 6).Font.Color = RGB(200, 0, 0)
                ws.Cells(startRow + r, 7).Font.Color = RGB(200, 0, 0)
            End If
        Next r

        startRow = startRow + outCount + 1
    Else
        ws.Cells(startRow, 1).Value = "No assets near targets/stops."
        startRow = startRow + 1
    End If

    ProcessRiskProximity = startRow + 1
End Function

Private Function ProcessAllocationGaps(ByVal ws As Worksheet, ByRef data() As Variant, dictHist As Object, maxDate As Date, d1 As Date, d2 As Date, d7 As Date, d14 As Date, startRow As Long) As Long
    ws.Cells(startRow, 1).Value = "Section 2: Allocation Gaps & Daily Performance"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 14
    startRow = startRow + 1

    Dim d1Str As String, d2Str As String, d7Str As String, d14Str As String
    d1Str = IIf(d1 > 0, Format(d1, "mm/dd"), "1D Ago")
    d2Str = IIf(d2 > 0, Format(d2, "mm/dd"), "2D Ago")
    d7Str = IIf(d7 > 0, Format(d7, "mm/dd"), "1W Ago")
    d14Str = IIf(d14 > 0, Format(d14, "mm/dd"), "2W Ago")

    Dim hHeaders As Variant
    hHeaders = Array("Asset ID", "Latest Name", "Current Price", "1D Perf", _
                    "Current Alloc", "MEAN Alloc", "Current Gap", _
                    "Gap Shift (" & d1Str & ")", "Gap Shift (" & d2Str & ")", _
                    "Gap Shift (" & d7Str & ")", "Gap Shift (" & d14Str & ")")

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, UBound(hHeaders) + 1)).Value = hHeaders

    Dim outData() As Variant
    ReDim outData(1 To UBound(data, 1), 1 To 11)
    Dim outCount As Long
    outCount = 0

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) = maxDate Then
                Dim assetID As String: assetID = data(i, dictCols("ASSET_ID"))
                Dim aName As String: aName = data(i, dictCols("ASSET_NAME"))

                Dim idx1D As Variant: idx1D = dictHist(assetID & "_" & Format(d1, "dd/mm/yyyy"))
                Dim idx2D As Variant: idx2D = dictHist(assetID & "_" & Format(d2, "dd/mm/yyyy"))
                Dim idx1W As Variant: idx1W = dictHist(assetID & "_" & Format(d7, "dd/mm/yyyy"))
                Dim idx2W As Variant: idx2W = dictHist(assetID & "_" & Format(d14, "dd/mm/yyyy"))

                outCount = outCount + 1
                outData(outCount, 1) = assetID
                outData(outCount, 2) = aName

                Dim cp As Double
                cp = SafeCDbl(data(i, dictCols("CURRENT_PRICE")))
                outData(outCount, 3) = cp

                ' 1D Perf
                Dim cp1D As Variant
                cp1D = SafeGetVal(data, idx1D, dictCols("CURRENT_PRICE"))
                If IsNumeric(cp1D) Then
                    If CDbl(cp1D) > 0 Then
                        outData(outCount, 4) = (cp / CDbl(cp1D)) - 1
                    Else
                        outData(outCount, 4) = "N/A"
                    End If
                Else
                    outData(outCount, 4) = "N/A"
                End If

                ' Current Alloc, MEAN Alloc, Current Gap
                Dim currAlloc As Double, meanAlloc As Double, currGap As Double
                currAlloc = SafeCDbl(data(i, dictCols("CURRENT_ALLOC")))
                meanAlloc = SafeCDbl(data(i, dictCols("MEAN")))
                currGap = currAlloc - meanAlloc

                outData(outCount, 5) = currAlloc
                outData(outCount, 6) = meanAlloc
                outData(outCount, 7) = currGap

                ' Gap Shifts
                Dim pastAlloc As Variant, pastMean As Variant, pastGap As Double

                ' 1D Shift
                pastAlloc = SafeGetVal(data, idx1D, dictCols("CURRENT_ALLOC"))
                pastMean = SafeGetVal(data, idx1D, dictCols("MEAN"))
                If IsNumeric(pastAlloc) And IsNumeric(pastMean) Then
                    pastGap = CDbl(pastAlloc) - CDbl(pastMean)
                    outData(outCount, 8) = currGap - pastGap
                Else
                    outData(outCount, 8) = "N/A"
                End If

                ' 2D Shift
                pastAlloc = SafeGetVal(data, idx2D, dictCols("CURRENT_ALLOC"))
                pastMean = SafeGetVal(data, idx2D, dictCols("MEAN"))
                If IsNumeric(pastAlloc) And IsNumeric(pastMean) Then
                    pastGap = CDbl(pastAlloc) - CDbl(pastMean)
                    outData(outCount, 9) = currGap - pastGap
                Else
                    outData(outCount, 9) = "N/A"
                End If

                ' 1W Shift
                pastAlloc = SafeGetVal(data, idx1W, dictCols("CURRENT_ALLOC"))
                pastMean = SafeGetVal(data, idx1W, dictCols("MEAN"))
                If IsNumeric(pastAlloc) And IsNumeric(pastMean) Then
                    pastGap = CDbl(pastAlloc) - CDbl(pastMean)
                    outData(outCount, 10) = currGap - pastGap
                Else
                    outData(outCount, 10) = "N/A"
                End If

                ' 2W Shift
                pastAlloc = SafeGetVal(data, idx2W, dictCols("CURRENT_ALLOC"))
                pastMean = SafeGetVal(data, idx2W, dictCols("MEAN"))
                If IsNumeric(pastAlloc) And IsNumeric(pastMean) Then
                    pastGap = CDbl(pastAlloc) - CDbl(pastMean)
                    outData(outCount, 11) = currGap - pastGap
                Else
                    outData(outCount, 11) = "N/A"
                End If

            End If
        End If
    Next i

    If outCount > 0 Then
        ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + outCount, 11)).Value = outData

        Call FormatTable(ws, startRow, outCount, 11)
        ws.Range(ws.Cells(startRow + 1, 4), ws.Cells(startRow + outCount, 11)).NumberFormat = "0.0%"

        Call FormatDiffColor(ws, startRow, outCount, 4)
        Call FormatDiffColor(ws, startRow, outCount, 7)
        Call FormatDiffColor(ws, startRow, outCount, 8)
        Call FormatDiffColor(ws, startRow, outCount, 9)
        Call FormatDiffColor(ws, startRow, outCount, 10)
        Call FormatDiffColor(ws, startRow, outCount, 11)

        startRow = startRow + outCount + 1
    Else
        ws.Cells(startRow, 1).Value = "No data available."
        startRow = startRow + 1
    End If

    ProcessAllocationGaps = startRow + 1
End Function

Private Function ProcessHistoricalChanges(ByVal ws As Worksheet, ByRef data() As Variant, dictHist As Object, maxDate As Date, d1 As Date, d2 As Date, d7 As Date, d14 As Date, startRow As Long) As Long
    ws.Cells(startRow, 1).Value = "Section 3: Historical Changes"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 14
    startRow = startRow + 1

    Dim cDateStr As String, d1Str As String, d2Str As String, d7Str As String, d14Str As String
    cDateStr = "Current"
    d1Str = IIf(d1 > 0, Format(d1, "mm/dd"), "1D Ago")
    d2Str = IIf(d2 > 0, Format(d2, "mm/dd"), "2D Ago")
    d7Str = IIf(d7 > 0, Format(d7, "mm/dd"), "1W Ago")
    d14Str = IIf(d14 > 0, Format(d14, "mm/dd"), "2W Ago")

    Dim hHeaders As Variant
    hHeaders = Array("Asset ID", "Latest Name", _
                    "Target Price (Latest)", "Dist to Target", _
                    "MEAN Alloc (" & cDateStr & ")", _
                    "MEAN Alloc (" & d1Str & ")", "Diff (" & d1Str & ")", _
                    "MEAN Alloc (" & d2Str & ")", "Diff (" & d2Str & ")", _
                    "MEAN Alloc (" & d7Str & ")", "Diff (" & d7Str & ")", _
                    "MEAN Alloc (" & d14Str & ")", "Diff (" & d14Str & ")")

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, UBound(hHeaders) + 1)).Value = hHeaders

    Dim outData() As Variant
    ReDim outData(1 To UBound(data, 1), 1 To 13)
    Dim outCount As Long
    outCount = 0

    Dim i As Long, pastVal As Variant
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) = maxDate Then
                Dim assetID As String: assetID = data(i, dictCols("ASSET_ID"))
                Dim aName As String: aName = data(i, dictCols("ASSET_NAME"))

                Dim idx1D As Variant: idx1D = dictHist(assetID & "_" & Format(d1, "dd/mm/yyyy"))
                Dim idx2D As Variant: idx2D = dictHist(assetID & "_" & Format(d2, "dd/mm/yyyy"))
                Dim idx1W As Variant: idx1W = dictHist(assetID & "_" & Format(d7, "dd/mm/yyyy"))
                Dim idx2W As Variant: idx2W = dictHist(assetID & "_" & Format(d14, "dd/mm/yyyy"))

                outCount = outCount + 1
                outData(outCount, 1) = assetID
                outData(outCount, 2) = aName

                Dim cp As Double, tp As Double
                cp = SafeCDbl(data(i, dictCols("CURRENT_PRICE")))
                tp = SafeCDbl(data(i, dictCols("TARGET_PRICE")))
                outData(outCount, 3) = tp
                If tp > 0 Then outData(outCount, 4) = Abs(cp - tp) / tp Else outData(outCount, 4) = "N/A"

                Dim meanCurr As Double
                meanCurr = SafeCDbl(data(i, dictCols("MEAN")))
                outData(outCount, 5) = meanCurr

                pastVal = SafeGetVal(data, idx1D, dictCols("MEAN"))
                outData(outCount, 6) = pastVal
                If IsNumeric(pastVal) Then outData(outCount, 7) = meanCurr - CDbl(pastVal) Else outData(outCount, 7) = "N/A"

                pastVal = SafeGetVal(data, idx2D, dictCols("MEAN"))
                outData(outCount, 8) = pastVal
                If IsNumeric(pastVal) Then outData(outCount, 9) = meanCurr - CDbl(pastVal) Else outData(outCount, 9) = "N/A"

                pastVal = SafeGetVal(data, idx1W, dictCols("MEAN"))
                outData(outCount, 10) = pastVal
                If IsNumeric(pastVal) Then outData(outCount, 11) = meanCurr - CDbl(pastVal) Else outData(outCount, 11) = "N/A"

                pastVal = SafeGetVal(data, idx2W, dictCols("MEAN"))
                outData(outCount, 12) = pastVal
                If IsNumeric(pastVal) Then outData(outCount, 13) = meanCurr - CDbl(pastVal) Else outData(outCount, 13) = "N/A"

            End If
        End If
    Next i

    If outCount > 0 Then
        ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + outCount, 13)).Value = outData

        Call FormatTable(ws, startRow, outCount, 13)
        ws.Range(ws.Cells(startRow + 1, 4), ws.Cells(startRow + outCount, 13)).NumberFormat = "0.0%"

        Call FormatDiffColor(ws, startRow, outCount, 7)
        Call FormatDiffColor(ws, startRow, outCount, 9)
        Call FormatDiffColor(ws, startRow, outCount, 11)
        Call FormatDiffColor(ws, startRow, outCount, 13)

        startRow = startRow + outCount + 1
    Else
        ws.Cells(startRow, 1).Value = "No historical data to compare."
        startRow = startRow + 1
    End If

    ProcessHistoricalChanges = startRow + 1
End Function

Private Function ProcessConvictionAnalysis(ByVal ws As Worksheet, ByRef data() As Variant, maxDate As Date, startRow As Long) As Long
    ws.Cells(startRow, 1).Value = "Section 4: Conviction Analysis"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 14
    startRow = startRow + 1

    Dim cHeaders As Variant
    cHeaders = Array("Conviction (/5)", "Avg Expected Return", "Avg Dist to Target", _
                     "Avg ERC UNCSTRD", "Avg ER/VOL", "Avg SHARPE", "Avg CUSTOM", "Avg MEAN Alloc", _
                     "Asset Count")
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, UBound(cHeaders) + 1)).Value = cHeaders

    Dim convCounts(1 To 5) As Long
    Dim convSumER(1 To 5) As Double
    Dim convSumDist(1 To 5) As Double
    Dim convCountDist(1 To 5) As Long

    Dim convSumERC(1 To 5) As Double
    Dim convSumERVOL(1 To 5) As Double
    Dim convSumSHARPE(1 To 5) As Double
    Dim convSumCUSTOM(1 To 5) As Double
    Dim convSumMEAN(1 To 5) As Double

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) = maxDate Then
                Dim cScore As Integer
                cScore = Int(SafeCDbl(data(i, dictCols("CONVICTION"))))
                If cScore >= 1 And cScore <= 5 Then
                    convCounts(cScore) = convCounts(cScore) + 1
                    convSumER(cScore) = convSumER(cScore) + SafeCDbl(data(i, dictCols("EXPECTED_RETURN")))
                    convSumERC(cScore) = convSumERC(cScore) + SafeCDbl(data(i, dictCols("ERC")))
                    convSumERVOL(cScore) = convSumERVOL(cScore) + SafeCDbl(data(i, dictCols("ERVOL")))
                    convSumSHARPE(cScore) = convSumSHARPE(cScore) + SafeCDbl(data(i, dictCols("SHARPE")))
                    convSumCUSTOM(cScore) = convSumCUSTOM(cScore) + SafeCDbl(data(i, dictCols("CUSTOM")))
                    convSumMEAN(cScore) = convSumMEAN(cScore) + SafeCDbl(data(i, dictCols("MEAN")))

                    Dim cp As Double, tp As Double
                    cp = SafeCDbl(data(i, dictCols("CURRENT_PRICE")))
                    tp = SafeCDbl(data(i, dictCols("TARGET_PRICE")))
                    If tp > 0 Then
                        convSumDist(cScore) = convSumDist(cScore) + (Abs(cp - tp) / tp)
                        convCountDist(cScore) = convCountDist(cScore) + 1
                    End If
                End If
            End If
        End If
    Next i

    Dim outData() As Variant
    ReDim outData(1 To 5, 1 To 9)
    Dim outCount As Long
    outCount = 0

    For i = 5 To 1 Step -1
        outCount = outCount + 1
        outData(outCount, 1) = i
        If convCounts(i) > 0 Then
            outData(outCount, 2) = convSumER(i) / convCounts(i)
            If convCountDist(i) > 0 Then outData(outCount, 3) = convSumDist(i) / convCountDist(i) Else outData(outCount, 3) = "N/A"
            outData(outCount, 4) = convSumERC(i) / convCounts(i)
            outData(outCount, 5) = convSumERVOL(i) / convCounts(i)
            outData(outCount, 6) = convSumSHARPE(i) / convCounts(i)
            outData(outCount, 7) = convSumCUSTOM(i) / convCounts(i)
            outData(outCount, 8) = convSumMEAN(i) / convCounts(i)
        Else
            outData(outCount, 2) = "N/A"
            outData(outCount, 3) = "N/A"
            outData(outCount, 4) = "N/A"
            outData(outCount, 5) = "N/A"
            outData(outCount, 6) = "N/A"
            outData(outCount, 7) = "N/A"
            outData(outCount, 8) = "N/A"
        End If
        outData(outCount, 9) = convCounts(i)
    Next i

    ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + outCount, 9)).Value = outData

    Call FormatTable(ws, startRow, outCount, 9)
    ws.Range(ws.Cells(startRow + 1, 2), ws.Cells(startRow + outCount, 8)).NumberFormat = "0.0%"

    ProcessConvictionAnalysis = startRow + outCount + 2
End Function

Private Function ProcessSchemeVariations(ByVal ws As Worksheet, ByRef data() As Variant, dictHist As Object, maxDate As Date, d1 As Date, d2 As Date, d7 As Date, d14 As Date, startRow As Long) As Long
    ws.Cells(startRow, 1).Value = "Section 5: Allocation Scheme Variations Over Time"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 14
    startRow = startRow + 1

    Dim cDateStr As String, d1Str As String, d2Str As String, d7Str As String, d14Str As String
    cDateStr = "Current"
    d1Str = IIf(d1 > 0, Format(d1, "mm/dd"), "1D Ago")
    d2Str = IIf(d2 > 0, Format(d2, "mm/dd"), "2D Ago")
    d7Str = IIf(d7 > 0, Format(d7, "mm/dd"), "1W Ago")
    d14Str = IIf(d14 > 0, Format(d14, "mm/dd"), "2W Ago")

    Dim hHeaders As Variant
    hHeaders = Array("Asset ID", "Latest Name", "Scheme", _
                    cDateStr, _
                    d1Str, "Diff (" & d1Str & ")", _
                    d2Str, "Diff (" & d2Str & ")", _
                    d7Str, "Diff (" & d7Str & ")", _
                    d14Str, "Diff (" & d14Str & ")")

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, UBound(hHeaders) + 1)).Value = hHeaders

    Dim schemeCols(1 To 5) As Long
    Dim schemeNames(1 To 5) As String
    schemeCols(1) = dictCols("ERC"): schemeNames(1) = "ERC UNCSTRD"
    schemeCols(2) = dictCols("ERVOL"): schemeNames(2) = "ER/VOL"
    schemeCols(3) = dictCols("SHARPE"): schemeNames(3) = "SHARPE"
    schemeCols(4) = dictCols("CUSTOM"): schemeNames(4) = "CUSTOM"
    schemeCols(5) = dictCols("MEAN"): schemeNames(5) = "MEAN"

    Dim outData() As Variant
    ReDim outData(1 To UBound(data, 1) * 5, 1 To 12)
    Dim outCount As Long
    outCount = 0

    Dim i As Long, s As Integer, pastVal As Variant
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, dictCols("DATE_SAVED"))) Then
            If data(i, dictCols("DATE_SAVED")) = maxDate Then
                Dim assetID As String: assetID = data(i, dictCols("ASSET_ID"))
                Dim aName As String: aName = data(i, dictCols("ASSET_NAME"))

                Dim idx1D As Variant: idx1D = dictHist(assetID & "_" & Format(d1, "dd/mm/yyyy"))
                Dim idx2D As Variant: idx2D = dictHist(assetID & "_" & Format(d2, "dd/mm/yyyy"))
                Dim idx1W As Variant: idx1W = dictHist(assetID & "_" & Format(d7, "dd/mm/yyyy"))
                Dim idx2W As Variant: idx2W = dictHist(assetID & "_" & Format(d14, "dd/mm/yyyy"))

                For s = 1 To 5
                    outCount = outCount + 1
                    outData(outCount, 1) = assetID
                    outData(outCount, 2) = aName
                    outData(outCount, 3) = schemeNames(s)

                    Dim currVal As Double
                    currVal = SafeCDbl(data(i, schemeCols(s)))
                    outData(outCount, 4) = currVal

                    ' 1D
                    pastVal = SafeGetVal(data, idx1D, schemeCols(s))
                    outData(outCount, 5) = pastVal
                    If IsNumeric(pastVal) Then outData(outCount, 6) = currVal - CDbl(pastVal) Else outData(outCount, 6) = "N/A"

                    ' 2D
                    pastVal = SafeGetVal(data, idx2D, schemeCols(s))
                    outData(outCount, 7) = pastVal
                    If IsNumeric(pastVal) Then outData(outCount, 8) = currVal - CDbl(pastVal) Else outData(outCount, 8) = "N/A"

                    ' 1W
                    pastVal = SafeGetVal(data, idx1W, schemeCols(s))
                    outData(outCount, 9) = pastVal
                    If IsNumeric(pastVal) Then outData(outCount, 10) = currVal - CDbl(pastVal) Else outData(outCount, 10) = "N/A"

                    ' 2W
                    pastVal = SafeGetVal(data, idx2W, schemeCols(s))
                    outData(outCount, 11) = pastVal
                    If IsNumeric(pastVal) Then outData(outCount, 12) = currVal - CDbl(pastVal) Else outData(outCount, 12) = "N/A"
                Next s
            End If
        End If
    Next i

    If outCount > 0 Then
        ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + outCount, 12)).Value = outData

        Call FormatTable(ws, startRow, outCount, 12)
        ws.Range(ws.Cells(startRow + 1, 4), ws.Cells(startRow + outCount, 12)).NumberFormat = "0.0%"

        Call FormatDiffColor(ws, startRow, outCount, 6)
        Call FormatDiffColor(ws, startRow, outCount, 8)
        Call FormatDiffColor(ws, startRow, outCount, 10)
        Call FormatDiffColor(ws, startRow, outCount, 12)

        startRow = startRow + outCount + 1
    Else
        ws.Cells(startRow, 1).Value = "No historical data to compare."
        startRow = startRow + 1
    End If

    ProcessSchemeVariations = startRow + 1
End Function
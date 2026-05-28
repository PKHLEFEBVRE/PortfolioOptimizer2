Option Explicit
Option Base 1

Function updatePriceHistoryFromInfinForFund()
    InitializeGlobals

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")

    ws.Cells.ClearContents

    Dim equityPositions As Dictionary
    Set equityPositions = GetEquityPositionsFromFund()

    Dim pos As Position, p As Variant, i As Integer

    i = 1
    ws.Activate

    For Each p In equityPositions
        Set pos = equityPositions(p)
        ws.Cells(1, i) = pos.posInstrument.Symbol
        ws.Cells(1, i + 1) = pos.posInstrument.Name
        ws.Cells(2, i).Formula2 = "=INFIN.GETSECURITYHISTORY(" & ws.Cells(1, i).Address & ")"
        i = i + 2
    Next p

    ThisWorkbook.Sheets("Dashboard").Range("G3").Value = CInt(4 * equityPositions.count / 5)

End Function

Function updatePriceHistoryFromInfinForBench()
    InitializeGlobals

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BenchData")

    ws.Cells.ClearContents

    Dim i As Integer, symbols As Variant, names As Variant
    symbols = Array("", "SPX", "", "RTY")
    names = Array("", "S&P 500 Index", "", "Russel 2000 Index")

    For i = 1 To 3 Step 2
        ws.Cells(1, i) = symbols(i)
        ws.Cells(1, i + 1) = names(i)
        ws.Cells(2, i).Formula2 = "=INFIN.GETSECURITYHISTORY(" & ws.Cells(1, i).Address & ")"
    Next i

End Function

Sub AlignSecurityDataRefactored()
    InitializeGlobals
    Dim wsInput As Worksheet
    Dim wsBench As Worksheet
    Dim lastCol As Long
    Dim startDate As Date
    Dim masterDates() As Date
    Dim dictBenchPrices As Object
    Dim dictPrices As Object

    Application.ScreenUpdating = False

    ' Set your starting parameters
    ' You can eventually link this to a cell, e.g., wsInput.Range("A1").Value
    startDate = ThisWorkbook.Sheets("Dashboard").Range("D2").Value

    ComputeEquallyWeightedBenchmark

    Set wsInput = ActiveSheet
    Set wsBench = ThisWorkbook.Sheets("BenchData")

    ' ---------------------------------------------------------
    ' PHASE 1.1: Extract Benchmark Timeline and Prices
    ' ---------------------------------------------------------
    ' This generates our Master Date Array and grabs benchmark prices
    Set dictBenchPrices = ExtractBenchmarkData(wsBench, startDate, masterDates)

    ' Check if we actually found dates
    If (Not Not masterDates) = 0 Then ' Fast way to check if array is uninitialized
        MsgBox "No benchmark dates found on or after the start date.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' ---------------------------------------------------------
    ' PHASE 1.2: Extract Asset Data (No more tallying)
    ' ---------------------------------------------------------
    lastCol = wsInput.Cells(1, wsInput.Columns.count).End(xlToLeft).Column
    Set dictPrices = ExtractAssetData(wsInput, lastCol)

    ' ---------------------------------------------------------
    ' PHASES 2, 3 & 4: Overlap, Beta Calculation, and Backfilling
    ' ---------------------------------------------------------
    Dim assetKey As Variant
    Dim secDict As Object
    Dim firstAssetDate As Date
    Dim assetBeta As Double, assetBetas As Dictionary
    Dim dateKey As Variant

    Set assetBetas = New Dictionary

    For Each assetKey In dictPrices.Keys
        Set secDict = dictPrices(assetKey)

        ' Find the earliest available date for this specific asset
        firstAssetDate = #12/31/9999#
        For Each dateKey In secDict.Keys
            If CDate(dateKey) < firstAssetDate Then
                firstAssetDate = CDate(dateKey)
            End If
        Next dateKey

        ' If the asset's first date is after our master start date, we need to backfill
        If firstAssetDate > startDate Then

            ' Phase 3: Calculate the Asset's Beta against the Benchmark
            assetBeta = CalculateLogBeta(masterDates, dictBenchPrices, secDict, firstAssetDate)
            assetBetas.Add assetKey, assetBeta
            ' Phase 4: Generate Prices Backwards (CALLING THE FUNCTION)
            BackfillAssetPrices secDict, masterDates, dictBenchPrices, startDate, firstAssetDate, assetBeta

        End If
    Next assetKey

    ' ---------------------------------------------------------
    ' PHASE 5: Output the Unified Data
    ' ---------------------------------------------------------
    WriteFundAlignedDataWithArray wsInput, masterDates, dictPrices

    ComputeEquallyWeightedBenchmark
    RunUpdateMatrices
    ThisWorkbook.Sheets("Dashboard").Activate

    Application.ScreenUpdating = True

    Dim betaString As String
    betaString = ""
    If assetBetas.count > 0 Then
        betaString = vbCr & vbCr & "Backfilling Beta used :" & vbCr
        For Each assetKey In assetBetas
            betaString = betaString & assetKey & " : " & Round(assetBetas(assetKey), 2) & vbCr
        Next assetKey
    End If
    MsgBox "Data aligned and backfilled successfully!" & vbCr & "From " & masterDates(1) & vbCr & "To " & masterDates(UBound(masterDates)) & betaString, vbInformation

End Sub

'
'Sub AlignSecurityDataRefactored()
'
'    ComputeEquallyWeightedBenchmark
'
'    Dim wsInput As Worksheet
'    Dim lastCol As Long, secCount As Long
'    Dim dictDates As Object, dictPrices As Object
'    Dim arrDates() As Date
'
'    Set wsInput = ActiveSheet
'    lastCol = wsInput.Cells(1, wsInput.Columns.count).End(xlToLeft).Column
'    secCount = lastCol / 2 ' Assuming 2 columns per security (Date, Price)
'
'    Set dictDates = CreateObject("Scripting.Dictionary")
'
'    Application.ScreenUpdating = False
'
'    ' STEP 1: Extract data using Arrays and Asset Names as Keys
'    Set dictPrices = ExtractData(wsInput, lastCol, dictDates)
'
'    ' STEP 2: Find and sort common dates
'    If Not GetFundCommonDates(dictDates, secCount, arrDates) Then
'        MsgBox "No common dates found across all securities.", vbExclamation
'        Application.ScreenUpdating = True
'        Exit Sub
'    End If
'
'    ' STEP 3: Write to the final sheet via an Output Array
'    WriteFundAlignedDataWithArray wsInput, arrDates, dictPrices
'
'    ComputeEquallyWeightedBenchmark
'
'    Application.ScreenUpdating = True
'    MsgBox "Data aligned successfully!" & vbCr & "Found " & UBound(arrDates) + 1 & " common dates. Range :" & vbCr & vbCr & "From " & arrDates(LBound(arrDates)) & vbCr & "To " & arrDates(UBound(arrDates)), vbInformation
'
'    RunUpdateMatrices
'    ThisWorkbook.Sheets("Dashboard").Activate
'End Sub


Public Function GetEquityPositionsFromFund() As Dictionary
    InitializeGlobals

    Dim ptf As Portfolio
    Set ptf = GetPortfolio

    Dim equityPositions As Dictionary, positions As Dictionary
    Set equityPositions = New Dictionary
    Dim pos As Position, p As Variant

    Set positions = ptf.GetPositions

    For Each p In positions
        Set pos = positions(p)
        If LCase(pos.posInstrument.assetClass) = "equity" Then
            equityPositions.Add p, pos
        End If
    Next p

    Set GetEquityPositionsFromFund = equityPositions

End Function

Private Function GetPortfolio() As Portfolio

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")

    ws.Activate

    Dim ptfs As Dictionary
    Set ptfs = AddInData.GetPortfolios()

    Dim ptf As Portfolio, p As Variant

    For Each p In ptfs
        If ptfs(p).IdentifierVIA = ws.Range("G2").Value Then
            Set ptf = ptfs(p)
            ptf.dt = Date
            Exit For
        End If
    Next p

    ptf.GetPositions

    Set GetPortfolio = ptf

End Function


' =====================================================================
' HELPER FUNCTIONS FOR PHASE 1
' =====================================================================

Private Function ExtractBenchmarkData(ws As Worksheet, startDate As Date, ByRef arrDates() As Date) As Object
    Dim dictBench As Object
    Dim lastRow As Long
    Dim srcArray As Variant
    Dim i As Long, count As Long
    Dim d As Date, p As Double

    Set dictBench = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).Row

    If lastRow < 2 Then GoTo EarlyExit

    ' Grab Columns E through H into memory
    ' Index 1 = Col E (Date), Index 4 = Col H (Price)
    srcArray = ws.Range(ws.Cells(2, "E"), ws.Cells(lastRow, "H")).Value

    ' Dimension array to maximum possible size to avoid slow ReDim Preserves
    ReDim arrDates(1 To UBound(srcArray, 1))
    count = 0

    For i = 1 To UBound(srcArray, 1)
        If IsDate(srcArray(i, 1)) Then
            d = CDate(srcArray(i, 1))

            ' Only keep dates on or after our Start Date
            If d >= startDate Then
                p = srcArray(i, 4)

                If Not dictBench.Exists(d) Then
                    dictBench.Add d, p
                    count = count + 1
                    arrDates(count) = d
                End If
            End If
        End If
    Next i

    ' Shrink the array down to the actual number of valid dates found
    If count > 0 Then
        ReDim Preserve arrDates(1 To count)
        SortDatesAscending arrDates ' Ensure chronological order
    Else
        Erase arrDates
    End If

EarlyExit:
    Set ExtractBenchmarkData = dictBench
End Function


Private Function ExtractAssetData(ws As Worksheet, lastCol As Long) As Object
    Dim dictPrices As Object, secDict As Object
    Dim i As Long, j As Long, lastRow As Long
    Dim dateVal As Variant, priceVal As Variant
    Dim assetName As String
    Dim srcArray As Variant

    Set dictPrices = CreateObject("Scripting.Dictionary")

    For i = 1 To lastCol Step 2
        lastRow = ws.Cells(ws.Rows.count, i).End(xlUp).Row

        If lastRow >= 2 Then
            assetName = ws.Cells(1, i + 1).Value
            Set secDict = CreateObject("Scripting.Dictionary")
            srcArray = ws.Range(ws.Cells(2, i), ws.Cells(lastRow, i + 1)).Value

            For j = 1 To UBound(srcArray, 1)
                dateVal = srcArray(j, 1)
                priceVal = srcArray(j, 2)

                ' Store the price if it's a valid date and number
                If IsDate(dateVal) And IsNumeric(priceVal) Then
                    If Not secDict.Exists(dateVal) Then
                        secDict.Add dateVal, priceVal
                    End If
                End If
            Next j

            dictPrices.Add assetName, secDict
        End If
    Next i

    Set ExtractAssetData = dictPrices
End Function

Private Sub SortDatesAscending(ByRef arr() As Date)
    ' Standard Bubble Sort for dates
    Dim i As Long, j As Long
    Dim temp As Date
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

' =====================================================================
' HELPER FUNCTIONS FOR PHASES 2 & 3
' =====================================================================

Private Function CalculateLogBeta(masterDates() As Date, dictBenchPrices As Object, dictAssetPrices As Object, firstAssetDate As Date) As Double
    Dim i As Long, count As Long
    Dim pAsset1 As Double, pAsset0 As Double
    Dim pBench1 As Double, pBench0 As Double
    Dim d1 As Date, d0 As Date

    Dim retAsset As Double, retBench As Double
    Dim sumAsset As Double, sumBench As Double

    ' Arrays to temporarily hold the valid returns
    Dim arrRetA() As Double, arrRetB() As Double
    ReDim arrRetA(1 To UBound(masterDates))
    ReDim arrRetB(1 To UBound(masterDates))

    count = 0
    sumAsset = 0
    sumBench = 0

    ' -------------------------------------------------------------
    ' PASS 1: Calculate Log Returns and Sums
    ' -------------------------------------------------------------
    For i = 2 To UBound(masterDates)
        d1 = masterDates(i)
        d0 = masterDates(i - 1)

        If d0 >= firstAssetDate Then
            If dictAssetPrices.Exists(d1) And dictAssetPrices.Exists(d0) And _
               dictBenchPrices.Exists(d1) And dictBenchPrices.Exists(d0) Then

                pAsset1 = dictAssetPrices(d1)
                pAsset0 = dictAssetPrices(d0)
                pBench1 = dictBenchPrices(d1)
                pBench0 = dictBenchPrices(d0)

                If pAsset0 > 0 And pBench0 > 0 And pAsset1 > 0 And pBench1 > 0 Then
                    count = count + 1

                    ' Note: In native VBA, Log() is the Natural Logarithm (Ln)
                    retAsset = Log(pAsset1 / pAsset0)
                    retBench = Log(pBench1 / pBench0)

                    arrRetA(count) = retAsset
                    arrRetB(count) = retBench

                    sumAsset = sumAsset + retAsset
                    sumBench = sumBench + retBench
                End If
            End If
        End If
    Next i

    ' -------------------------------------------------------------
    ' PASS 2: Calculate NATIVE Covariance / Variance (Beta)
    ' -------------------------------------------------------------
    If count >= 2 Then
        Dim meanA As Double, meanB As Double
        Dim covSum As Double, varSum As Double

        meanA = sumAsset / count
        meanB = sumBench / count
        covSum = 0
        varSum = 0

        For i = 1 To count
            covSum = covSum + ((arrRetA(i) - meanA) * (arrRetB(i) - meanB))
            varSum = varSum + ((arrRetB(i) - meanB) ^ 2)
        Next i

        If varSum > 0 Then
            CalculateLogBeta = covSum / varSum
        Else
            CalculateLogBeta = 1 ' Default if benchmark variance is literally zero
        End If
    Else
        CalculateLogBeta = 1 ' Default if not enough overlapping data points
    End If
End Function


' =====================================================================
' HELPER FUNCTION FOR PHASE 4
' =====================================================================

Private Sub BackfillAssetPrices(ByRef secDict As Object, masterDates() As Date, dictBenchPrices As Object, startDate As Date, firstAssetDate As Date, assetBeta As Double)
    Dim i As Long, startIndex As Long
    Dim d1 As Date, d0 As Date
    Dim pBench1 As Double, pBench0 As Double
    Dim pAsset1 As Double, pAsset0 As Double
    Dim rBench As Double, rProxy As Double

    ' 1. Find where the asset's real history begins on the Master Timeline
    For i = 1 To UBound(masterDates)
        If masterDates(i) = firstAssetDate Then
            startIndex = i
            Exit For
        End If
    Next i

    ' If we couldn't find it, or it's the very first date, exit
    If startIndex <= 1 Then Exit Sub

    ' 2. Work backwards day-by-day
    For i = startIndex To 2 Step -1
        d1 = masterDates(i)
        d0 = masterDates(i - 1)

        ' Stop if we go before the requested start date
        If d0 < startDate Then Exit For

        ' Ensure we have the necessary data points to calculate
        If dictBenchPrices.Exists(d1) And dictBenchPrices.Exists(d0) And secDict.Exists(d1) Then
            pBench1 = dictBenchPrices(d1)
            pBench0 = dictBenchPrices(d0)
            pAsset1 = secDict(d1)

            If pBench1 > 0 And pBench0 > 0 Then
                ' Calculate log return of benchmark
                rBench = Application.WorksheetFunction.Ln(pBench1 / pBench0)

                ' Apply beta to get proxy return
                rProxy = assetBeta * rBench

                ' Reverse the log return to find yesterday's price
                pAsset0 = pAsset1 / Exp(rProxy)

                ' Inject the newly calculated price into the asset's dictionary
                If Not secDict.Exists(d0) Then
                    secDict.Add d0, pAsset0
                End If
            End If
        End If
    Next i
End Sub


' =====================================================================
' UPDATED PHASE 5: STRICT INTERSECTION OUTPUT
' =====================================================================

Private Sub WriteFundAlignedDataWithArray(wsInput As Worksheet, masterDates() As Date, dictPrices As Object)
    Dim wsOutput As Worksheet
    Dim outArray As Variant
    Dim assetNames As Variant
    Dim r As Long, c As Long
    Dim dateKey As Date
    Dim allExist As Boolean

    Dim validDateCount As Long
    Dim validDates() As Date

    assetNames = dictPrices.Keys

    ' Dimension a temporary array to hold the dates that survive the filter
    ReDim validDates(1 To UBound(masterDates))
    validDateCount = 0

    ' ---------------------------------------------------------
    ' 1. Filter the Master Dates for Strict Overlap
    ' ---------------------------------------------------------
    For r = 1 To UBound(masterDates)
        dateKey = masterDates(r)
        allExist = True

        ' Check if every single asset has this date
        For c = 0 To UBound(assetNames)
            If Not dictPrices(assetNames(c)).Exists(dateKey) Then
                allExist = False
                Exit For ' Stop checking if even one asset is missing it
            End If
        Next c

        ' If all assets have a price, keep the date
        If allExist Then
            validDateCount = validDateCount + 1
            validDates(validDateCount) = dateKey
        End If
    Next r

    ' Safety check in case the filter removes everything
    If validDateCount = 0 Then
        MsgBox "After alignment, no dates exist where ALL assets have a price.", vbExclamation
        Exit Sub
    End If

    ' ---------------------------------------------------------
    ' 2. Build the Final Output Array
    ' ---------------------------------------------------------
    ReDim outArray(1 To validDateCount + 1, 1 To dictPrices.count + 1)

    ' Headers
    outArray(1, 1) = "Aligned Date"
    For c = 0 To UBound(assetNames)
        outArray(1, c + 2) = assetNames(c)
    Next c

    ' Dates and Prices
    For r = 1 To validDateCount
        dateKey = validDates(r)
        outArray(r + 1, 1) = dateKey

        For c = 0 To UBound(assetNames)
            outArray(r + 1, c + 2) = dictPrices(assetNames(c))(dateKey)
        Next c
    Next r

    ' ---------------------------------------------------------
    ' 3. Drop Array on Sheet
    ' ---------------------------------------------------------
    Set wsOutput = ThisWorkbook.Sheets("PriceHistory")
    wsOutput.Cells.ClearContents

    ' Resize target block to perfectly match the filtered array size
    wsOutput.Cells(1, 1).Resize(validDateCount + 1, dictPrices.count + 1).Value = outArray

    ' Clean up formatting
    wsOutput.Columns(1).NumberFormat = "dd/mm/yyyy"
    wsOutput.Columns.AutoFit
End Sub

'Private Function ExtractData(ws As Worksheet, lastCol As Long, ByRef dictDates As Object) As Object
'    Dim dictPrices As Object, secDict As Object
'    Dim i As Long, j As Long, lastRow As Long
'    Dim dateVal As Variant, priceVal As Variant
'    Dim assetName As String
'    Dim srcArray As Variant
'
'    Set dictPrices = CreateObject("Scripting.Dictionary")
'
'    For i = 1 To lastCol Step 2
'        lastRow = ws.Cells(ws.Rows.count, i).End(xlUp).Row
'
'        If lastRow >= 2 Then
'            ' Use the Price Column Header (Column i + 1) as the Asset Name key
'            assetName = ws.Cells(1, i + 1).Value
'            Set secDict = CreateObject("Scripting.Dictionary")
'
'            ' Snatch data into memory array
'            srcArray = ws.Range(ws.Cells(2, i), ws.Cells(lastRow, i + 1)).Value
'
'            For j = 1 To UBound(srcArray, 1)
'                dateVal = srcArray(j, 1)
'                priceVal = srcArray(j, 2)
'
'                If IsDate(dateVal) Then
'                    ' Tally matching dates across assets
'                    If Not dictDates.Exists(dateVal) Then
'                        dictDates.Add dateVal, 1
'                    Else
'                        dictDates(dateVal) = dictDates(dateVal) + 1
'                    End If
'
'                    ' Store price mapped to date
'                    If Not secDict.Exists(dateVal) Then
'                        secDict.Add dateVal, priceVal
'                    End If
'                End If
'            Next j
'
'            ' Store the asset dictionary using the Asset Name as the main key
'            dictPrices.Add assetName, secDict
'        End If
'    Next i
'
'    Set ExtractData = dictPrices
'End Function
'
'
'
'Private Function GetFundCommonDates(dictDates As Object, secCount As Long, ByRef arrDates() As Date) As Boolean
'    Dim key As Variant
'    Dim count As Long
'    Dim tempDate As Date
'    Dim m As Long, n As Long
'
'    count = 0
'    For Each key In dictDates.Keys
'        If dictDates(key) = secCount Then
'            ReDim Preserve arrDates(count)
'            arrDates(count) = CDate(key)
'            count = count + 1
'        End If
'    Next key
'
'    If count = 0 Then
'        GetFundCommonDates = False
'        Exit Function
'    End If
'
'    ' Quick Bubble Sort (Oldest to Newest)
'    For m = LBound(arrDates) To UBound(arrDates) - 1
'        For n = m + 1 To UBound(arrDates)
'            If arrDates(m) > arrDates(n) Then
'                tempDate = arrDates(m)
'                arrDates(m) = arrDates(n)
'                arrDates(n) = tempDate
'            End If
'        Next n
'    Next m
'
'    GetFundCommonDates = True
'End Function



'Private Sub WriteFundAlignedDataWithArray(wsInput As Worksheet, arrDates() As Date, dictPrices As Object)
'    Dim wsOutput As Worksheet
'    Dim outArray As Variant
'    Dim assetNames As Variant
'    Dim totalRows As Long, totalCols As Long
'    Dim r As Long, c As Long
'    Dim dateKey As Date
'
'    ' Extract asset names keys out of our dictionary
'    assetNames = dictPrices.Keys
'
'    totalRows = UBound(arrDates) + 2 ' +1 for 0-index, +1 for Header row
'    totalCols = dictPrices.count + 1 ' +1 for the unified Date column
'
'    ' Dimensions of output matrix: (Rows, Columns)
'    ReDim outArray(1 To totalRows, 1 To totalCols)
'
'    ' 1. Populate Headers in the array matrix
'    outArray(1, 1) = "Aligned Date"
'    For c = 0 To UBound(assetNames)
'        outArray(1, c + 2) = assetNames(c)
'    Next c
'
'    ' 2. Populate Dates and Prices in the array matrix
'    For r = 0 To UBound(arrDates)
'        dateKey = arrDates(r)
'        outArray(r + 2, 1) = dateKey ' Date column
'
'        ' Retrieve prices using Asset Name keys
'        For c = 0 To UBound(assetNames)
'            outArray(r + 2, c + 2) = dictPrices(assetNames(c))(dateKey)
'        Next c
'    Next r
'
'    ' 3. Add worksheet and drop the array on the sheet all at once
'    Set wsOutput = ThisWorkbook.Sheets("PriceHistory")
'    wsOutput.Cells.ClearContents
'
'    ' Resize target block to perfectly match array size
'    wsOutput.Cells(1, 1).Resize(totalRows, totalCols).Value = outArray
'
'    ' Clean up formatting
'    wsOutput.Columns(1).NumberFormat = "dd/mm/yyyy" ' Enforces neat dates
'    wsOutput.Columns.AutoFit
'
'End Sub


Private Sub ComputeEquallyWeightedBenchmark()
    Dim wsData As Worksheet
    Dim arrSec1 As Variant, arrSec2 As Variant
    Dim dictSec1 As Object
    Dim outArr As Variant
    Dim outRows As Long

    ' 1. Define the source data sheet
    ' Change "Sheet1" to match your actual sheet name
    Set wsData = ThisWorkbook.Sheets("BenchData")

    ' 2. Load the raw data ranges directly into memory arrays
    arrSec1 = LoadRangeToArray(wsData, "A", "B")
    arrSec2 = LoadRangeToArray(wsData, "C", "D")

    ' 3. Build a Dictionary from Security 1 for instant date lookups
    Set dictSec1 = BuildDictionary(arrSec1)
    If dictSec1.count = 0 Then
        MsgBox "No valid data found for Security 1.", vbCritical
        Exit Sub
    End If

    ' 4. Process common dates and perform all calculations in-memory
    outArr = ProcessAndCalculate(dictSec1, arrSec2, outRows)

    ' 5. Exit if no overlap was found
    If outRows = 0 Then
        MsgBox "No common dates found between the two securities.", vbExclamation
        Exit Sub
    End If

    ' 6. Dump the final calculated array onto a new spreadsheet
    Call OutputResults(outArr, outRows)

End Sub


' ==========================================
' HELPER FUNCTIONS
' ==========================================

' Dynamically finds the last row and loads two columns into a 2D memory array
Private Function LoadRangeToArray(ws As Worksheet, col1 As String, col2 As String) As Variant
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, col1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2 ' Failsafe for empty columns

    LoadRangeToArray = ws.Range(ws.Cells(2, col1), ws.Cells(lastRow, col2)).Value
End Function

' Converts a 2D array (Date/Price) into a Dictionary for lightning-fast matching
Private Function BuildDictionary(arr As Variant) As Object
    Dim dict As Object
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")

    If Not IsEmpty(arr) Then
        For i = 1 To UBound(arr)
            ' Ensure we are looking at a valid date and a valid number
            If IsDate(arr(i, 1)) And IsNumeric(arr(i, 2)) Then
                dict(arr(i, 1)) = arr(i, 2)
            End If
        Next i
    End If

    Set BuildDictionary = dict
End Function

' Loops through Security 2, checks the Dictionary for matches, and calculates the benchmark
Private Function ProcessAndCalculate(dict As Object, arr2 As Variant, ByRef outRows As Long) As Variant
    Dim tempArr() As Variant
    Dim i As Long
    Dim dateVal As Variant, price1 As Double, price2 As Double
    Dim base1 As Double, base2 As Double
    Dim isBaseSet As Boolean

    ' Max possible size is the length of array 2, spanning 6 columns
    ReDim tempArr(1 To UBound(arr2), 1 To 4)
    outRows = 0
    isBaseSet = False

    For i = 1 To UBound(arr2)
        dateVal = arr2(i, 1)

        If dict.Exists(dateVal) And IsNumeric(arr2(i, 2)) Then
            outRows = outRows + 1
            price1 = dict(dateVal)
            price2 = arr2(i, 2)

            ' Assign Raw Data
            tempArr(outRows, 1) = dateVal
            'tempArr(outRows, 2) = price1
            'tempArr(outRows, 3) = price2

            ' Set the Base 100 starting prices on the very first matched date
            If Not isBaseSet Then
                base1 = price1
                base2 = price2
                isBaseSet = True
            End If

            ' Calculate Normalized Prices and Benchmark
            tempArr(outRows, 2) = (price1 / base1) * 100
            tempArr(outRows, 3) = (price2 / base2) * 100
            tempArr(outRows, 4) = (tempArr(outRows, 2) + tempArr(outRows, 3)) / 2
        End If
    Next i

    ' Return the populated array
    ProcessAndCalculate = tempArr
End Function

' Generates the final worksheet, drops the array, and applies formatting
Private Sub OutputResults(outArr As Variant, outRows As Long)
    Dim wsOut As Worksheet

    Set wsOut = ThisWorkbook.Sheets("BenchData")

    ' Write Headers
    wsOut.Range("E1:H1").Value = Array("Common Date", "Norm Price 1 (Base 100)", "Norm Price 2 (Base 100)", "Equally Weighted Benchmark")

    ' Bulk-drop the array data onto the sheet
    wsOut.Range("E2").Resize(outRows, 4).Value = outArr

    ' Apply Formatting
    wsOut.Columns("E:E").NumberFormat = "dd/mm/yyyy"
    wsOut.Columns("F:H").NumberFormat = "0.00"
    wsOut.Columns.AutoFit

End Sub
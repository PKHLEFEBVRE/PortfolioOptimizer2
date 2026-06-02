Attribute VB_Name = "DataUtils"
Option Explicit

Function updateAllPriceHistoryFromInfin()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")

    ws.Cells.ClearContents

    Dim equityPositions As Dictionary
    Set equityPositions = GetEquityPositionsFromFund()

    Dim pos As Position, p As Variant, i As Integer

    i = 1
    ws.Activate

    ' Fetch benchmark assets on the same sheet FIRST (Columns 1-4)
    Dim symbols As Variant, names As Variant
    symbols = Array("", "SPX", "", "RTY")
    names = Array("", "S&P 500 Index", "", "Russel 2000 Index")

    Dim k As Integer
    For k = 1 To 3 Step 2
        ws.Cells(1, i) = symbols(k)
        ws.Cells(1, i + 1) = names(k)
        ws.Cells(2, i).Formula2 = "=INFIN.GETSECURITYHISTORY(" & ws.Cells(1, i).Address & ")"
        i = i + 2
    Next k

    ' Fetch fund assets AFTER the benchmark
    For Each p In equityPositions
        Set pos = equityPositions(p)
        ws.Cells(1, i) = pos.posInstrument.Symbol
        ws.Cells(1, i + 1) = pos.posInstrument.Name
        ws.Cells(2, i).Formula2 = "=INFIN.GETSECURITYHISTORY(" & ws.Cells(1, i).Address & ")"
        i = i + 2
    Next p

    ' Deprecated logic: used to write portfolio max constraints to the old dashboard.
    ' Now handled natively inside OptimizerCls/Positions Dashboard.
    ' ThisWorkbook.Sheets("Positions Dashboard").Cells(5, 2).Value = CInt(4 * equityPositions.Count / 5)

End Function

Sub AlignSecurityDataRefactored()
    Dim wsInput As Worksheet
    Dim lastCol As Long
    Dim startDate As Date
    Dim masterDates() As Date
    Dim dictBenchPrices As Object
    Dim dictPrices As Object

    Application.ScreenUpdating = False

    ' Retrieve parameters from the new dashboard
    On Error Resume Next
    startDate = ThisWorkbook.Sheets("Positions Dashboard").Cells(4, 2).Value
    On Error GoTo 0
    If startDate = 0 Then startDate = Date - 365 * 3 ' 3 years default
    Set wsInput = ThisWorkbook.Sheets("Data")

    ' 1. Extract all asset prices
    lastCol = wsInput.Cells(1, wsInput.Columns.Count).End(xlToLeft).Column
    Set dictPrices = ExtractAssetData(wsInput, lastCol)

    Dim assetNames() As Variant
    assetNames = dictPrices.Keys()
    If UBound(assetNames) < 1 Then
        MsgBox "Not enough benchmark assets found.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim b1 As String, b2 As String
    b1 = assetNames(0)
    b2 = assetNames(1)

    ' 2. Determine overlapping common dates for the benchmark
    If Not DetermineMasterDates(dictPrices(b1), dictPrices(b2), startDate, masterDates) Then
        MsgBox "No overlapping benchmark dates found on or after the start date.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 3. Build the 50/50 benchmark curve
    Set dictBenchPrices = BuildBenchmarkFromAssets(dictPrices(b1), dictPrices(b2), masterDates)

    ' 4. Backfill any missing historical fund data using Downside Beta against the benchmark
    Dim betaString As String
    betaString = BackfillMissingAssetHistory(dictPrices, masterDates, dictBenchPrices, startDate, b1, b2)

    ' 5. Output the cleaned matrix
    WriteFundAlignedDataWithArray wsInput, masterDates, dictPrices

    RunUpdateMatrices

    On Error Resume Next
    ThisWorkbook.Sheets("Positions Dashboard").Activate
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "Data aligned and backfilled successfully!" & vbCr & "From " & masterDates(1) & vbCr & "To " & masterDates(UBound(masterDates)) & betaString, vbInformation
End Sub

' -------------------------------------------------------------------------
' MODERATE REFACTOR HELPERS
' -------------------------------------------------------------------------

Private Function DetermineMasterDates(dictB1 As Object, dictB2 As Object, startDate As Date, ByRef outDates() As Date) As Boolean
    Dim dKey As Variant
    Dim allDates As Object
    Set allDates = CreateObject("Scripting.Dictionary")

    For Each dKey In dictB1.Keys
        If dictB2.Exists(dKey) Then
            If CDate(dKey) >= startDate Then
                If Not allDates.Exists(CDate(dKey)) Then
                    allDates.Add CDate(dKey), 1
                End If
            End If
        End If
    Next dKey

    If allDates.Count = 0 Then
        DetermineMasterDates = False
        Exit Function
    End If

    Dim arrD() As Date
    ReDim arrD(1 To allDates.Count)
    Dim count As Long
    count = 0
    For Each dKey In allDates.Keys
        count = count + 1
        arrD(count) = CDate(dKey)
    Next dKey

    SortDatesAscending arrD

    ReDim outDates(1 To UBound(arrD))
    For count = 1 To UBound(arrD)
        outDates(count) = arrD(count)
    Next count

    DetermineMasterDates = True
End Function

Private Function BuildBenchmarkFromAssets(dictB1 As Object, dictB2 As Object, masterDates() As Date) As Object
    Dim dictBench As Object
    Set dictBench = CreateObject("Scripting.Dictionary")

    Dim base1 As Double, base2 As Double
    Dim isBaseSet As Boolean
    isBaseSet = False
    Dim i As Long

    For i = 1 To UBound(masterDates)
        If Not isBaseSet Then
            base1 = dictB1(masterDates(i))
            base2 = dictB2(masterDates(i))
            isBaseSet = True
        End If

        Dim val1 As Double, val2 As Double
        val1 = (dictB1(masterDates(i)) / base1) * 100
        val2 = (dictB2(masterDates(i)) / base2) * 100

        dictBench.Add masterDates(i), (val1 + val2) / 2
    Next i

    Set BuildBenchmarkFromAssets = dictBench
End Function

Private Function BackfillMissingAssetHistory(dictPrices As Object, masterDates() As Date, dictBenchPrices As Object, startDate As Date, b1 As String, b2 As String) As String
    Dim assetKey As Variant
    Dim secDict As Object
    Dim firstAssetDate As Date
    Dim assetBeta As Double
    Dim assetBetas As Object
    Dim dKey As Variant

    Set assetBetas = CreateObject("Scripting.Dictionary")

    For Each assetKey In dictPrices.Keys
        If assetKey <> b1 And assetKey <> b2 Then
            Set secDict = dictPrices(assetKey)

            firstAssetDate = #12/31/9999#
            For Each dKey In secDict.Keys
                If CDate(dKey) < firstAssetDate Then
                    firstAssetDate = CDate(dKey)
                End If
            Next dKey

            If firstAssetDate > startDate Then
                assetBeta = CalculateLogBeta(masterDates, dictBenchPrices, secDict, firstAssetDate)
                assetBetas.Add assetKey, assetBeta
                BackfillAssetPrices secDict, masterDates, dictBenchPrices, startDate, firstAssetDate, assetBeta
            End If
        End If
    Next assetKey

    Dim betaString As String
    betaString = ""
    If assetBetas.Count > 0 Then
        betaString = vbCr & vbCr & "Backfilling Beta used :" & vbCr
        For Each assetKey In assetBetas
            betaString = betaString & assetKey & " : " & Round(assetBetas(assetKey), 2) & vbCr
        Next assetKey
    End If

    BackfillMissingAssetHistory = betaString
End Function

Public Function GetEquityPositionsFromFund() As Dictionary

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
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Positions Dashboard")
    On Error GoTo 0

    If Not ws Is Nothing Then ws.Activate

    Dim ptfs As Dictionary
    Set ptfs = AddInData.GetPortfolios()

    Dim ptf As Portfolio, p As Variant

    Dim targetID As String
    If Not ws Is Nothing Then
        targetID = CStr(ws.Cells(5, 2).Value)
    Else
        targetID = "DEFAULT_ID"
    End If

    For Each p In ptfs
        If ptfs(p).IdentifierVIA = targetID Then
            Set ptf = ptfs(p)
            ptf.dt = Date
            Exit For
        End If
    Next p

    ptf.GetPositions

    Set GetPortfolio = ptf

End Function

Private Function ExtractAssetData(ws As Worksheet, lastCol As Long) As Object
    Dim dictPrices As Object, secDict As Object
    Dim i As Long, j As Long, lastRow As Long
    Dim dateVal As Variant, priceVal As Variant
    Dim assetName As String
    Dim srcArray As Variant

    Set dictPrices = CreateObject("Scripting.Dictionary")

    For i = 1 To lastCol Step 2
        lastRow = ws.Cells(ws.Rows.Count, i).End(xlUp).Row

        If lastRow >= 2 Then
            assetName = ws.Cells(1, i + 1).Value
            Set secDict = CreateObject("Scripting.Dictionary")
            srcArray = ws.Range(ws.Cells(2, i), ws.Cells(lastRow, i + 1)).Value

            For j = 1 To UBound(srcArray, 1)
                dateVal = srcArray(j, 1)
                priceVal = srcArray(j, 2)

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

Private Function CalculateLogBeta(masterDates() As Date, dictBenchPrices As Object, dictAssetPrices As Object, firstAssetDate As Date) As Double
    Dim i As Long, count As Long
    Dim pAsset1 As Double, pAsset0 As Double
    Dim pBench1 As Double, pBench0 As Double
    Dim d1 As Date, d0 As Date

    Dim retAsset As Double, retBench As Double
    Dim sumAsset As Double, sumBench As Double

    Dim arrRetA() As Double, arrRetB() As Double
    ReDim arrRetA(1 To UBound(masterDates))
    ReDim arrRetB(1 To UBound(masterDates))

    count = 0
    sumAsset = 0
    sumBench = 0

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
            CalculateLogBeta = 1
        End If
    Else
        CalculateLogBeta = 1
    End If
End Function

Private Sub BackfillAssetPrices(ByRef secDict As Object, masterDates() As Date, dictBenchPrices As Object, startDate As Date, firstAssetDate As Date, assetBeta As Double)
    Dim i As Long, startIndex As Long
    Dim d1 As Date, d0 As Date
    Dim pBench1 As Double, pBench0 As Double
    Dim pAsset1 As Double, pAsset0 As Double
    Dim rBench As Double, rProxy As Double

    For i = 1 To UBound(masterDates)
        If masterDates(i) = firstAssetDate Then
            startIndex = i
            Exit For
        End If
    Next i

    If startIndex <= 1 Then Exit Sub

    For i = startIndex To 2 Step -1
        d1 = masterDates(i)
        d0 = masterDates(i - 1)

        If d0 < startDate Then Exit For

        If dictBenchPrices.Exists(d1) And dictBenchPrices.Exists(d0) And secDict.Exists(d1) Then
            pBench1 = dictBenchPrices(d1)
            pBench0 = dictBenchPrices(d0)
            pAsset1 = secDict(d1)

            If pBench1 > 0 And pBench0 > 0 Then
                rBench = Application.WorksheetFunction.Ln(pBench1 / pBench0)
                rProxy = assetBeta * rBench
                pAsset0 = pAsset1 / Exp(rProxy)

                If Not secDict.Exists(d0) Then
                    secDict.Add d0, pAsset0
                End If
            End If
        End If
    Next i
End Sub

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

    ReDim validDates(1 To UBound(masterDates))
    validDateCount = 0

    For r = 1 To UBound(masterDates)
        dateKey = masterDates(r)
        allExist = True

        For c = 0 To UBound(assetNames)
            If Not dictPrices(assetNames(c)).Exists(dateKey) Then
                allExist = False
                Exit For
            End If
        Next c

        If allExist Then
            validDateCount = validDateCount + 1
            validDates(validDateCount) = dateKey
        End If
    Next r

    If validDateCount = 0 Then
        MsgBox "After alignment, no dates exist where ALL assets have a price.", vbExclamation
        Exit Sub
    End If

    ReDim outArray(1 To validDateCount + 1, 1 To dictPrices.Count + 1)

    outArray(1, 1) = "Aligned Date"
    For c = 0 To UBound(assetNames)
        outArray(1, c + 2) = assetNames(c)
    Next c

    For r = 1 To validDateCount
        dateKey = validDates(r)
        outArray(r + 1, 1) = dateKey

        For c = 0 To UBound(assetNames)
            outArray(r + 1, c + 2) = dictPrices(assetNames(c))(dateKey)
        Next c
    Next r

    Set wsOutput = ThisWorkbook.Sheets("PriceHistory")
    wsOutput.Cells.ClearContents

    wsOutput.Cells(1, 1).Resize(validDateCount + 1, dictPrices.Count + 1).Value = outArray

    wsOutput.Columns(1).NumberFormat = "dd/mm/yyyy"
    wsOutput.Columns.AutoFit
End Sub

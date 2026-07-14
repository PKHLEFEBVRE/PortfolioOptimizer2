Function CreateCombinedAssetDictionary(sourceRange1 As Range, sourceRange2 As Range) As Object
    InitializeGlobals
    Dim mainDict As Object
    Dim innerDict As Object
    Dim dataArray As Variant
    Dim r As Long, c As Long
    Dim primaryKey As String, secondaryKey As String
    Dim i As Integer
    Dim rangesArray() As Variant

    Set mainDict = CreateObject("Scripting.Dictionary")

    ReDim rangesArray(1 To 2)
    rangesArray(1) = sourceRange1.Value
    rangesArray(2) = sourceRange2.Value

    Dim start As Integer

    For i = 1 To 2
        dataArray = rangesArray(i)
        If IsArray(dataArray) Then
            If UBound(dataArray, 1) >= 2 And UBound(dataArray, 2) >= 2 Then
                For r = 2 To UBound(dataArray, 1)
                    If Not IsEmpty(dataArray(r, 1)) Then
                        primaryKey = CStr(dataArray(r, 1))
                        If Not mainDict.Exists(primaryKey) Then
                            Set innerDict = CreateObject("Scripting.Dictionary")
                            mainDict.Add primaryKey, innerDict
                        Else
                            Set innerDict = mainDict(primaryKey)
                        End If
                        If i = 2 Then start = 3 Else start = 2
                        For c = start To UBound(dataArray, 2)
                            secondaryKey = CStr(dataArray(1, c))
                            innerDict(secondaryKey) = dataArray(r, c)
                        Next c
                    End If
                Next r
            End If
        End If
    Next i

    Set CreateCombinedAssetDictionary = mainDict
End Function



Sub SaveResultsToTop(dictAssets As Object, targetSheetName As String)
    InitializeGlobals
    Dim ws As Worksheet
    Dim numRows As Long, numCols As Long
    Dim innerDict As Object
    Dim keysArray As Variant, firstKey As Variant
    Dim r As Long, c As Long
    Dim assetKey As Variant, attrKey As Variant

    ' 1. Set the target sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Target sheet '" & targetSheetName & "' not found.", vbCritical
        Exit Sub
    End If

    numRows = dictAssets.count
    If numRows = 0 Then Exit Sub ' Nothing to save

    ' 2. Determine columns (Inner dict count + 1 for Primary Key + 1 for Date)
    keysArray = dictAssets.Keys
    firstKey = keysArray(0)
    Set innerDict = dictAssets.Item(firstKey)
    numCols = innerDict.count + 3

    ' 3. Insert new rows at row 2, pushing existing data down
    ws.Rows("2:" & CStr(numRows + 1)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' 4. Create a 2D array to hold the data
    Dim dataArray() As Variant
    ReDim dataArray(1 To numRows, 1 To numCols)

    ' 5. Populate the array and write headers
    r = 1
    For Each assetKey In dictAssets.Keys
        Set innerDict = dictAssets.Item(assetKey)

        ' Write the Primary Key and Date into the first two columns
        dataArray(r, 1) = assetKey
        dataArray(r, 2) = Date ' Use Now instead of Date if you want the exact time too

        ' Write the static headers in Row 1
        If r = 1 Then
            ws.Cells(1, 1).Value = "Asset_ID"
            ws.Cells(1, 2).Value = "Date_Saved"
        End If

        ' Start populating inner dictionary data from the 3rd column
        c = 3
        For Each attrKey In innerDict.Keys
            ' Write the dynamic headers to Row 1
            If r = 1 Then
                ws.Cells(1, c).Value = attrKey
            End If

            ' Populate the 2D array with inner dict values
            dataArray(r, c) = innerDict.Item(attrKey)
            c = c + 1
        Next attrKey

        dataArray(r, c) = innerDict("DIFF target") / innerDict("DIFF")
        r = r + 1

    Next assetKey

    ' 6. Dump the array into the newly inserted rows
    ws.Range(ws.Cells(2, 1), ws.Cells(numRows + 1, numCols)).Value = dataArray

    'MsgBox "Historical data saved successfully!", vbInformation
End Sub
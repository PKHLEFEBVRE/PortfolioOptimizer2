Sub UpdateConvictions()
    InitializeGlobals

    Dim convictions As Dictionary
    Set convictions = getConvictions

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")

    Dim i As Integer, lastRow As Integer, startRow As Integer, colStart As Integer
    startRow = ws.Range("InputTableStart").Row + 1
    lastRow = ws.Range("InputTableStart").End(xlDown).Row
    colStart = ws.Range("InputTableStart").Column

    For i = ws.Range("InputTableStart").Row + 1 To lastRow
        If ws.Cells(i, 2).Value <> "" Then
            ws.Cells(i, colStart + 2).Value = convictions(CStr(ws.Cells(i, 2).Value))("conviction")
            ws.Cells(i, colStart + 3).Value = convictions(CStr(ws.Cells(i, 2).Value))("TP")
            ws.Cells(i, colStart + 4).Value = convictions(CStr(ws.Cells(i, 2).Value))("TP proba")
            ws.Cells(i, colStart + 7).Value = Date + 365
            ws.Cells(i, colStart + 8).Value = 1 / (4 * (lastRow - startRow + 1))
            ws.Cells(i, colStart + 9).Value = 2 / (lastRow - startRow + 1)
        Else
            Exit For
        End If
    Next i

End Sub

Public Function getConvictions() As Dictionary
    InitializeGlobals

    Dim convictions As Dictionary, singleConviction As Dictionary
    Set convictions = New Dictionary

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Convictions")

    Dim i As Integer, lastRow As Integer
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    i = 2

    For i = 2 To lastRow
        Set singleConviction = New Dictionary
        singleConviction.Add "conviction", ws.Cells(i, 4).Value
        singleConviction.Add "TP", ws.Cells(i, 5).Value
        singleConviction.Add "TP proba", ws.Cells(i, 6).Value
        singleConviction.Add "SP", ws.Cells(i, 7).Value
        singleConviction.Add "SP proba", ws.Cells(i, 8).Value
        convictions.Add ws.Cells(i, 1).Value, singleConviction
    Next i

    Set getConvictions = convictions

End Function
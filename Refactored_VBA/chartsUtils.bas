Option Explicit
Option Base 1




Sub UpdateDashboardCharts(nAssets As Integer, nStrats As Integer, lastDate As Date)
    InitializeGlobals
    Dim wsDash As Worksheet, wsChart As Worksheet
    Set wsDash = Sheets(config.DashSheet)
    Set wsChart = Sheets(config.ChartSheet)

    Dim lastRow As Long
    lastRow = wsChart.Cells(wsChart.Rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    Dim xRange As Range
    Set xRange = wsChart.Range(wsChart.Cells(2, 1), wsChart.Cells(lastRow, 1))

    Dim chtObj As ChartObject
    For Each chtObj In wsDash.ChartObjects
        Dim title As String
        On Error Resume Next
        title = chtObj.Chart.ChartTitle.text
        On Error GoTo 0

        Dim startCol As Integer, count As Integer
        Dim isPerf As Boolean
        isPerf = False

        If title = "Assets Cumulative Returns" Then
            startCol = 2
            count = nAssets
            isPerf = True
        ElseIf title = "Strategies Cumulative Returns" Then
            startCol = nAssets + 2
            count = nStrats
            isPerf = True
        ElseIf title = "Strategies Drawdowns" Then
            startCol = nAssets + 2 + nStrats
            count = nStrats
            isPerf = False
        Else
            startCol = 0
        End If

        If startCol > 0 Then
            UpdateSingleChartSeries chtObj.Chart, xRange, wsChart, startCol, count

            Dim totalMonths As Long
            ' Calcul de la durée en mois entre le début et la fin


            With chtObj.Chart.Axes(xlCategory)
            totalMonths = DateDiff("m", .MinimumScale, .MaximumScale)
                .CategoryType = xlTimeScale
                .TickLabels.NumberFormat = "mmm yy"
                .MaximumScale = lastDate
                ' Ajustement dynamique de l'unité majeure
                Select Case totalMonths
                    Case Is <= 12
                        .MajorUnitScale = xlMonths
                        .MajorUnit = 1        ' Un label par mois
                    Case 13 To 24
                        .MajorUnitScale = xlMonths
                        .MajorUnit = 3        ' Un label tous les trimestres
                    Case 25 To 48
                        .MajorUnitScale = xlMonths
                        .MajorUnit = 6        ' Un label tous les semestres
                    Case Else
                        .MajorUnitScale = xlYears
                        .MajorUnit = 1        ' Un label par an
                End Select
            End With

            If isPerf Then
                Dim dRange As Range
                Set dRange = wsChart.Range(wsChart.Cells(2, startCol), wsChart.Cells(lastRow, startCol + count - 1))
                Dim minVal As Double, maxVal As Double
                minVal = Application.WorksheetFunction.Min(dRange)
                maxVal = Application.WorksheetFunction.Max(dRange)
                Dim rangeVal As Double
                rangeVal = maxVal - minVal
                If rangeVal = 0 Then rangeVal = 10
                Dim axisMin As Double, axisMax As Double
                axisMin = minVal - (rangeVal * 0.05)
                axisMax = maxVal + (rangeVal * 0.05)
                With chtObj.Chart.Axes(xlValue)
                    .MinimumScale = axisMin
                    .MaximumScale = axisMax
                    .TickLabels.NumberFormat = "0"
                End With
            Else
                Dim ddRange As Range
                Set ddRange = wsChart.Range(wsChart.Cells(2, startCol), wsChart.Cells(lastRow, startCol + count - 1))
                Dim minDD As Double
                minDD = Application.WorksheetFunction.Min(ddRange)
                With chtObj.Chart.Axes(xlValue)
                    .MaximumScale = 0
                    .MinimumScale = minDD * 1.1
                    .TickLabels.NumberFormat = "0%"
                End With
            End If
        End If
    Next chtObj
End Sub


Sub UpdateSingleChartSeries(cht As Chart, xRange As Range, wsData As Worksheet, startCol As Integer, count As Integer)
    InitializeGlobals
    Dim s As Series
    Do While cht.SeriesCollection.count > 0
        cht.SeriesCollection(1).Delete
    Loop
    Dim i As Integer
    For i = 0 To count - 1
        Dim colIdx As Integer
        colIdx = startCol + i
        Set s = cht.SeriesCollection.NewSeries
        s.XValues = xRange
        s.Values = wsData.Range(wsData.Cells(2, colIdx), wsData.Cells(xRange.Rows.count + 1, colIdx))
        s.Name = wsData.Cells(1, colIdx).Value
        s.Format.Line.Weight = 1.25
    Next i
    With cht.Legend
        .Position = xlLegendPositionTop
        .IncludeInLayout = True
        .Top = cht.ChartTitle.Top + cht.ChartTitle.Height + 5
    End With
End Sub
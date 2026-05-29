Attribute VB_Name = "portfolioUtils"
Option Explicit
Option Base 1

Private Const DASH_SHEET As String = "Dashboard"
Private Const WEIGHTS_SHEET As String = "SimulatedWeights"
Private Const RISK_FREE_ROW As Integer = 4




Sub OutputMetricsToRow(rowOffset As Integer, TypeStr As String, mRet, mVol, mSharpe, mMDD, mLen, mVaR, w, RangeName As String, WriteWeights As Boolean)
    Dim activeSht As Worksheet
    Set activeSht = ThisWorkbook.ActiveSheet

    Dim ws As Worksheet
    Set ws = Sheets(DASH_SHEET)
    Dim startRow As Long, startCol As Long
    startRow = ws.Range(RangeName).Row + 1
    startCol = ws.Range(RangeName).Column

    Dim r As Integer
    r = startRow + rowOffset
    If RangeName = "AssetMetricsStart" Then
        startCol = startCol - 1
    Else
        ws.Cells(r, startCol).Value = TypeStr
    End If

    ws.Cells(r, startCol + 1).Value = mRet
    ws.Cells(r, startCol + 2).Value = mVol
    ws.Cells(r, startCol + 3).Value = mSharpe
    ws.Cells(r, startCol + 4).Value = mMDD
    ws.Cells(r, startCol + 5).Value = mLen
    ws.Cells(r, startCol + 6).Value = mVaR

    If WriteWeights Then

        Dim equityPositions As Dictionary
        Set equityPositions = DataUtils.GetEquityPositionsFromFund

        Dim equitySum As Double, i As Integer
        equitySum = 0

        For i = 1 To equityPositions.count
            equitySum = equitySum + equityPositions(equityPositions.Keys(i - 1)).Weight
        Next i

        startRow = ws.Range("StrategyWeightsStart").Row
        startCol = ws.Range("StrategyWeightsStart").Column + 1 + rowOffset

        ws.Cells(startRow, startCol).Value = TypeStr

        Dim nCols As Integer
        nCols = UBound(w, 2)

        For i = 1 To nCols
            ws.Cells(startRow + i, startCol).Value = w(1, i)
            ws.Cells(startRow + i, startCol + 7).Value = w(1, i) * equitySum
        Next i
    End If

    activeSht.Activate
End Sub
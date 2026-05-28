Option Explicit
Option Base 1


Sub WriteToEngine(covMat() As Double, expReturns() As Double)
    InitializeGlobals
    Dim wsEng As Worksheet
    Dim nAssets As Long
    Dim i As Long, j As Long
    Set wsEng = Sheets(config.EngineSheet)
    nAssets = UBound(expReturns)
    wsEng.Cells.Clear
    wsEng.Range("A1").Value = "Expected Returns"
    wsEng.Range(wsEng.Cells(2, 1), wsEng.Cells(2, nAssets)).Value = expReturns
    wsEng.Range(wsEng.Cells(2, 1), wsEng.Cells(2, nAssets)).Name = "ExpReturns"
    wsEng.Range("A5").Value = "Covariance Matrix"
    wsEng.Range(wsEng.Cells(6, 1), wsEng.Cells(6 + nAssets - 1, nAssets)).Value = covMat
    wsEng.Range(wsEng.Cells(6, 1), wsEng.Cells(6 + nAssets - 1, nAssets)).Name = "CovMatrix"
    Dim solverRow As Long
    solverRow = 6 + nAssets + 5
    Call SetupOptimizationFormulas(wsEng, nAssets, solverRow)
End Sub

' ---------------------------------------------------------
' 1. SETUP: Added Log Calculation Rows for the "Log Trick"
' ---------------------------------------------------------
Sub SetupOptimizationFormulas(wsEng As Worksheet, nTotal As Long, startRow As Long)
    InitializeGlobals
    wsEng.Cells(startRow - 1, 1).Value = "Weights (Solver)"
    Dim i As Long
    ' Initialize with Equal Weights (Safe starting point for Log)
    For i = 1 To nTotal
        wsEng.Cells(startRow, i).Value = 1 / nTotal
    Next i
    wsEng.Range(wsEng.Cells(startRow, 1), wsEng.Cells(startRow, nTotal)).Name = "OptWeights"

    Dim rf As Double
    rf = Sheets(config.DashSheet).Range("D" & config.RiskFreeRow).Value

    Dim fRow As Long
    fRow = startRow + 2

    ' Standard Portfolio Calcs
    wsEng.Cells(fRow, 1).Value = "Port Return"
    wsEng.Cells(fRow, 2).Formula = "=SUMPRODUCT(OptWeights, ExpReturns)"
    wsEng.Cells(fRow, 2).Name = "PortRet"

    wsEng.Cells(fRow + 1, 1).Value = "Port Variance"
    wsEng.Cells(fRow + 1, 2).FormulaArray = "=MMULT(OptWeights, MMULT(CovMatrix, TRANSPOSE(OptWeights)))"
    wsEng.Cells(fRow + 1, 2).Name = "PortVar"

    wsEng.Cells(fRow + 2, 1).Value = "Port Vol"
    wsEng.Cells(fRow + 2, 2).Formula = "=SQRT(PortVar)"
    wsEng.Cells(fRow + 2, 2).Name = "PortVol"

    wsEng.Cells(fRow + 3, 1).Value = "Sharpe"
    wsEng.Cells(fRow + 3, 2).Formula = "=(PortRet - " & rf & ")/PortVol"
    wsEng.Cells(fRow + 3, 2).Name = "PortSharpe"

    wsEng.Cells(fRow + 4, 1).Value = "Kelly"
    wsEng.Cells(fRow + 4, 2).Formula = "=PortRet - (PortVar/2)"
    wsEng.Cells(fRow + 4, 2).Name = "PortKelly"

    wsEng.Cells(fRow + 5, 1).Value = "Sum Weights"
    wsEng.Cells(fRow + 5, 2).Formula = "=SUM(OptWeights)"
    wsEng.Cells(fRow + 5, 2).Name = "SumWeights"

    wsEng.Cells(fRow + 6, 1).Value = "Portfolio HHI"
    wsEng.Cells(fRow + 6, 2).Formula = "=SUMPRODUCT(OptWeights, OptWeights)"
    wsEng.Cells(fRow + 6, 2).Name = "PortHHI"

    wsEng.Cells(fRow + 7, 1).Value = "Max HHI Limit"
    wsEng.Cells(fRow + 7, 2).Formula = 1 / Sheets(config.DashSheet).Range("G3").Value
    wsEng.Cells(fRow + 7, 2).Name = "MaxHHI"

    ' --- ERC UNCSTRD Log-Barrier Setup ---
    Dim logRow As Long
    logRow = fRow + 9 ' Shifted down by 2 to accommodate the HHI rows
    wsEng.Cells(logRow, 1).Value = "Log Weights"
    Dim wRange As Range
    Set wRange = wsEng.Range("OptWeights")
    For i = 1 To nTotal
        wsEng.Cells(logRow, 1 + i).Formula = "=LN(MAX(" & wRange.Cells(1, i).Address & ", 0.00001))"
    Next i
    wsEng.Cells(logRow + 1, 1).Value = "Sum Logs"
    wsEng.Cells(logRow + 1, 2).Formula = "=SUM(" & wsEng.Range(wsEng.Cells(logRow, 2), wsEng.Cells(logRow, 1 + nTotal)).Address & ")"
    wsEng.Cells(logRow + 1, 2).Name = "SumLogWeights"

    ' Visualization of Risk Contributions (Optional, for checking result)
    Dim rcRow As Long
    rcRow = logRow + 3
    wsEng.Cells(rcRow, 1).Value = "Risk Contribs Check"
    wsEng.Range(wsEng.Cells(rcRow, 2), wsEng.Cells(rcRow, 1 + nTotal)).FormulaArray = _
        "=OptWeights * TRANSPOSE(MMULT(CovMatrix, TRANSPOSE(OptWeights)))"

End Sub
' ---------------------------------------------------------
' 2. SOLVER: Implemented High Precision & Log Logic
' ---------------------------------------------------------
Sub RunSolver(Mode As String, minWeights() As Double, maxWeights() As Double, Optional ShowMsg As Boolean = True)
    InitializeGlobals
    Dim wsEng As Worksheet
    Set wsEng = Sheets(config.EngineSheet)
    Dim nAssets As Integer
    nAssets = UBound(maxWeights)

    SolverReset
    SolverOptions Precision:=0.00000001, Convergence:=0.00000001, Derivatives:=2, IntTolerance:=0, AssumeNonNeg:=True, StepThru:=False

    Dim optRange As Range
    Set optRange = wsEng.Range("OptWeights")
    Dim i As Integer

    Dim eqW As Double
    eqW = 1 / nAssets
    For i = 1 To nAssets
        wsEng.Range("OptWeights").Cells(1, i).Value = eqW
    Next i

    If Mode = "ERC UNCSTRD" Then
        ' --- STRATÉGIE DUAL CONVEXE (La plus robuste) ---
        ' On cherche à Maximiser la Diversification (Sum Logs) pour un budget de risque donné.
        SolverOk SetCell:="SumLogWeights", MaxMinVal:=1, ValueOf:=0, ByChange:="OptWeights", Engine:=1
        SolverAdd CellRef:="PortVar", Relation:=1, FormulaText:="1"
        SolverAdd CellRef:="OptWeights", Relation:=3, FormulaText:="0.001"

    Else
        ' --- STRATÉGIES CLASSIQUES (Sharpe, Variance, Kelly) ---
        SolverAdd CellRef:="SumWeights", Relation:=2, FormulaText:="1"
        For i = 1 To nAssets
            SolverAdd CellRef:=optRange.Cells(1, i).Address, Relation:=1, FormulaText:=CStr(maxWeights(i))
            SolverAdd CellRef:=optRange.Cells(1, i).Address, Relation:=3, FormulaText:=CStr(minWeights(i))
        Next i
        Select Case Mode
            Case "SHARPE":
                SolverOk SetCell:="PortSharpe", MaxMinVal:=1, ValueOf:=0, ByChange:="OptWeights", Engine:=1
                SolverAdd CellRef:="PortHHI", Relation:=1, FormulaText:="MaxHHI"
            Case "MIN VAR":
                SolverOk SetCell:="PortVar", MaxMinVal:=2, ValueOf:=0, ByChange:="OptWeights", Engine:=1
            Case "KELLY":
                SolverOk SetCell:="PortKelly", MaxMinVal:=1, ValueOf:=0, ByChange:="OptWeights", Engine:=1
        End Select
    End If

    ' Exécution
    On Error Resume Next
    Dim res As Integer
    res = SolverSolve(True)
    On Error GoTo 0

    ' --- NORMALISATION FINALE POUR ERC UNCSTRD ---
    If Mode = "ERC UNCSTRD" And res <= 2 Then
        Dim currentSum As Double
        currentSum = Application.WorksheetFunction.sum(optRange)

        If currentSum <> 0 Then
            Dim finalWeights() As Variant
            ReDim finalWeights(1 To nAssets)
            For i = 1 To nAssets
                finalWeights(i) = optRange.Cells(1, i).Value / currentSum
            Next i
            optRange.Value = finalWeights
        End If
    End If

    ' Gestion d'erreur
    If res > 2 Then
        For i = 1 To nAssets
            wsEng.Range("OptWeights").Cells(1, i).Value = 1 / nAssets
        Next i
        MsgBox Mode & " - Échec de convergence. Retour aux poids égaux.", vbExclamation
    ElseIf ShowMsg Then
        MsgBox "Optimisation terminée (" & Mode & ")", vbInformation
    End If

End Sub
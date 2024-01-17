Attribute VB_Name = "TSP"
Option Explicit

Function WipeTableData(TargetRange As Range)
    With Application
        .StatusBar = "Clearing Table..."
    End With
    ' remove data precedent arrows
    TargetRange.ShowPrecedents Remove:=False

    ' clear text on screen
    TargetRange.ClearContents

End Function


Function ResetData()

    Data.Range("B18").value = 1
    Data.Range("B19").value = 2
    Data.Range("B20").value = 3
    Data.Range("B21").value = 4

End Function

Function PopulateMap(TargetRange As Range, DisplacementX As Byte, DisplacementY As Byte)
    With Application
        .StatusBar = "Populating Map..."
    End With
    
    
    Dim InitialRange As Range
    Dim WPTargetRange As Range
    Dim PopulateTargetRange As Range
    Dim PrevPopulateTargetRange As Range
    Dim TaskValue As Variant
    Dim Task As Range
    Dim PrevTask As Range
    
    Set InitialRange = TargetRange.Resize(1, 1)
    Set WPTargetRange = Main.Range("E15")
    
    Dim i As Byte
    For i = 1 To 5
        Set PrevPopulateTargetRange = PopulateTargetRange
        Set PopulateTargetRange = InitialRange.Offset(Range(WPTargetRange).Column - 1, Range(WPTargetRange).Row - 1)
        Set PrevTask = WPTargetRange
        Set Task = WPTargetRange.Offset(1, 0)
        
        
        ' populate the current cell with formula, its the first in order, give it a standard text
        If Not PrevTask Is Nothing Then
            TaskValue = Task.value
            If Not PrevPopulateTargetRange Is Nothing Then
                PopulateTargetRange.Formula = "=CONCAT(""" & Task & """, ""-""," & "RIGHT(" & PrevPopulateTargetRange.Address & ", LEN(" & PrevPopulateTargetRange.Address & ") - FIND(""-"", " & PrevPopulateTargetRange.Address & ") - 1), " & i & ")"
                PopulateTargetRange.ShowPrecedents
            Else
                PopulateTargetRange.Formula = "START-1"
            End If
        End If
        
        ' prep for next loop
        Set WPTargetRange = WPTargetRange.Offset(0, 1)
    Next i
End Function

Function PopulateTable(DisplacementX As Byte, DisplacementY As Byte)
    With Application
        .StatusBar = "Populating Table..."
    End With
    
    Dim FromCell As Range
    Set FromCell = Main.Range("O13")
    Dim ToCellX As Range
    Dim ToCellY As Range
    Set ToCellX = Data.Range("D10")
    Set ToCellY = Data.Range("E10")
    
    Dim ToCellName As Range
    Set ToCellName = Data.Range("G10")
    
    Dim i As Byte
    For i = 1 To 5
        ToCellName.value = FromCell.value
        ToCellX.value = Range(FromCell.value).Column - 1
        ToCellY.value = Range(FromCell.value).Row - 1
        
        Set FromCell = FromCell.Offset(1, 0)
        Set ToCellX = ToCellX.Offset(1, 0)
        Set ToCellY = ToCellY.Offset(1, 0)
        Set ToCellName = ToCellName.Offset(1, 0)
    Next i


End Function

Sub RunSolverEvolutionary()
    With Application
        .StatusBar = "Solving..."
    End With

    ' excel solver w evolutionary method
    Worksheets("Data").Activate
    SolverReset
    SolverOk SetCell:=Range("Data!$D$24"), MaxMinVal:=2, ByChange:=Range("Data!$B$18:$B$21"), Engine:=3, EngineDesc:="Evolutionary"
    SolverOptions AssumeNonNeg:=True
    SolverAdd CellRef:=Range("Data!$B$18:$B$21").Address, Relation:=6
    SolverSolve True
    Worksheets("Main").Activate
End Sub



Sub TravellingSalesmanProblem()

    Dim TableRange As Range
    Set TableRange = Main.Range("B2:M8")
    
    Dim TableDisplacementX As Byte
    Dim TableDisplacementY As Byte
    TableDisplacementX = 2
    TableDisplacementY = 2
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
    
    Call WipeTableData(TableRange)
    Call ResetData
    DoEvents
    
    Call PopulateTable(TableDisplacementX, TableDisplacementY)
    Call RunSolverEvolutionary
    DoEvents
    Call PopulateMap(TableRange, TableDisplacementX, TableDisplacementY)
    DoEvents
    

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = ""
        .Calculation = xlCalculationAutomatic
    End With

End Sub


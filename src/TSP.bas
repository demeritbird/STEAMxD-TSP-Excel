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


Function PopulateTable(TargetRange As Range, DisplacementX As Byte, DisplacementY As Byte)
    With Application
        .StatusBar = "Populating Table..."
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
    Call PopulateTable(TableRange, TableDisplacementX, TableDisplacementY)

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = ""
        .Calculation = xlCalculationAutomatic
    End With

End Sub



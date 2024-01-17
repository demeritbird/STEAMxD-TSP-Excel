Attribute VB_Name = "TSP"
Option Explicit

Function WipeTableData(TargetRange As Range)

' remove data precedent arrows
    TargetRange.ShowPrecedents Remove:=False

    ' clear text on screen
    TargetRange.ClearContents

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
        .StatusBar = "Clearing Table..."
        .Calculation = xlCalculationManual
    End With
    
    Call WipeTableData(TableRange)

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = ""
        .Calculation = xlCalculationAutomatic
    End With

End Sub



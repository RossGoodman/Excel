Private Sub Worksheet_Change(ByVal Target As Range)
 
If Application.Range("TrackChangesOn")(1).Value = "Yes" Then
    If Intersect(Target, Range("HeaderRows, TrackingColumns")) Is Nothing Then
        For Each aCell In Target
            'MsgBox "You just changed " & aCell.Row & " " & Application.UserName
            'MsgBox "cells " & Range("TrackingColumns").Column & ":" & aCell.Row
    
            Cells(aCell.Row, Range("UpdateDate").Column).Value = Date
            Cells(aCell.Row, Range("UpdateBy").Column).Value = Application.UserName
        Next
    End If
End If
 
End Sub

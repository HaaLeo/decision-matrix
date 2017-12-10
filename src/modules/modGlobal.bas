Attribute VB_Name = "modGlobal"
Option Explicit

'variables
'=========

'letzte Zeile vor der aktuellen Eingabe
Public gOldLastRow As Long

'functions
'=========

Public Function LastRowInColumnA( _
    ByVal nMyWorksheet As Worksheet) _
    As Long
    
    LastRowInColumnA = nMyWorksheet.Cells(Rows.Count, 1).End(xlUp).Row
End Function

Public Sub RangeToArray( _
    ByRef nArray() As Variant, _
    ByVal nRange As range)
    
    Erase nArray
    ReDim nArray(1 To nRange.Count, 1 To 1)
        
    If nRange.Count = 1 Then
        nArray(1, 1) = nRange.Value
    Else
        nArray = range(Cells(nRange.Row, 1), Cells(nRange.Row + nRange.Count / nRange.Columns.Count - 1, 1)).Value
    End If
End Sub

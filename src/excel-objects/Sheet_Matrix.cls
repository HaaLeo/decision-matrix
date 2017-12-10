VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change( _
    ByVal Target As range)
    
    Dim currentRow As Long
    Dim loopKriteria As Long
    Dim lastRow As Long
    Dim sheetShaper As clsMatrixHandler
    Dim matrixSheet As Worksheet
    Dim currentEditedKriteria() As Variant
    Dim originTargetRow As Long
    
    If Target.column = 1 Then
    
        Set sheetShaper = New clsMatrixHandler
        Set matrixSheet = Sheet_Matrix
        Call RangeToArray(currentEditedKriteria, Target)
        currentRow = Target.Row
        originTargetRow = Target.Row
        
        Application.EnableEvents = False
        matrixSheet.Unprotect
         
        'Spalte A ausgenommen A1 Sperrung aufheben
        matrixSheet.Columns(1).Locked = False
        matrixSheet.Cells(1, 1).Locked = True
        
        '"Design"
        For loopKriteria = 1 To Target.Count / Target.Columns.Count
            If Not currentEditedKriteria(loopKriteria, 1) = "" Then
                'Führende und nachfolgende Leerzeichen "absägen" und enter löschen
                'nUserInput = Trim(nMatrixSheet.Cells(nCurrentKriteriaRow, 1))
                currentEditedKriteria(loopKriteria, 1) = Replace(currentEditedKriteria(loopKriteria, 1), Chr(13), "")
                
                'Für Lesbarkeit bei Eingabehilfe
                'geschütztes Leerzeichen einfügen
                currentEditedKriteria(loopKriteria, 1) = Replace(currentEditedKriteria(loopKriteria, 1), " ", Chr(160))
                
                matrixSheet.Cells(loopKriteria + originTargetRow - 1, 1) = currentEditedKriteria(loopKriteria, 1)
            End If
        Next loopKriteria
        
        'Target kann mehrdimensional sein
        For loopKriteria = 1 To Target.Count / Target.Columns.Count
            
            Call ShapeMatrix(CStr(currentEditedKriteria(loopKriteria, 1)), currentRow)
            currentRow = currentRow + 1
        Next loopKriteria
        
        'wurde Inhalt gelöscht?
        lastRow = LastRowInColumnA(matrixSheet)
        If gOldLastRow > lastRow Then
            Call sheetShaper.DeleteEmptyCells(matrixSheet, lastRow)
        End If
        
        'Weil sich beim löschen Zeilenanzahl nocheinmal ändern kann
        gOldLastRow = LastRowInColumnA(matrixSheet)
        
        matrixSheet.Protect
        Application.EnableEvents = True
        
        'Speicher freigeben
        Erase currentEditedKriteria
    
    End If
   
    
End Sub

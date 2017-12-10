Attribute VB_Name = "modShapeMatrix"
Option Explicit

Public Sub ShapeMatrix( _
    ByVal nCurrentKriteria As String, _
    ByVal nCurrentRow As Long)
    
    Dim matrixSheet As Worksheet
    Dim matrixShaper As clsMatrixHandler
    Dim lastRow As Long
    Dim isDuplicate As Boolean
    
    isDuplicate = False
    
    Set matrixSheet = Sheet_Matrix
    Set matrixShaper = New clsMatrixHandler
    
    lastRow = LastRowInColumnA(matrixSheet)
    
    'wurde nichts gelöscht?
    If gOldLastRow <= lastRow Then
        If Not nCurrentKriteria = "" Then
             Call matrixShaper.ValidateUserInput(matrixSheet, nCurrentKriteria, nCurrentRow, lastRow, isDuplicate)
        End If
        
        Call matrixShaper.UpdateDataValidation(matrixSheet, nCurrentRow)

        'wurde etwas hinzugefügt?
        'wenn lastRow < 2 würd der erste Eintrag eine Validation erhalten
        If gOldLastRow < lastRow And lastRow > 2 And Not isDuplicate Then
            Call matrixShaper.ShapeBorderOfFilledCells(matrixSheet, lastRow)
            Call matrixShaper.AddFormula(matrixSheet, lastRow)
        End If
    End If

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvaluationHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ReadMatrix( _
    ByVal nMatrixSheet As Worksheet, _
    ByRef nAllCriteria() As clsCriterium, _
    ByRef nError As Boolean)
    
    Dim loopRow As Integer
    Dim lastRow As Integer
    Dim currentCriterium As clsCriterium
    
    
    
    lastRow = LastRowInColumnA(nMatrixSheet)
    ReDim nAllCriteria(1 To lastRow - 1)
    
    For loopRow = 2 To lastRow
        If Not nMatrixSheet.Cells(loopRow, 1) = "" Then
            Set currentCriterium = New clsCriterium
            
            currentCriterium.Priority = WorksheetFunction.Sum(range(nMatrixSheet.Cells(loopRow, 2), nMatrixSheet.Cells(loopRow, lastRow)))
            currentCriterium.Name = nMatrixSheet.Cells(loopRow, 1)
    
            Set nAllCriteria(loopRow - 1) = currentCriterium
        'Error: Matrix ist nicht komplett ausgefüllt
        ElseIf nMatrixSheet.Cells(loopRow, 1) = "" Then
            nError = True
            MsgBox "Die Zelle A" & loopRow & " darf nicht leer sein!" & vbLf & vbLf & "Das Makro wird kontrolliert beendet!", , "Matrix unvollständig!"
            nMatrixSheet.Cells(loopRow, 1).Select
            Exit Sub
        Else
            nError = True
            MsgBox "Ein unbekannter Fehler ist aufgetreten!" & vbLf & "Bitte öffnen Sie die Datei neu oder wenden Sie sich an den Entwickler!", , "Unbekannter Fehler!"
            Exit Sub
        End If
    Next loopRow
End Sub

Public Sub BubbleSort( _
    ByRef nAllCriteria() As clsCriterium)

    Dim loopOuterCriterium As Integer
    Dim loopInnerCriterium As Integer
    Dim tempCriterium As clsCriterium
  
    For loopOuterCriterium = UBound(nAllCriteria) To 1 Step -1
        For loopInnerCriterium = 1 To loopOuterCriterium - 1
            If nAllCriteria(loopInnerCriterium).Priority < nAllCriteria(loopInnerCriterium + 1).Priority Then
                Set tempCriterium = nAllCriteria(loopInnerCriterium)
                Set nAllCriteria(loopInnerCriterium) = nAllCriteria(loopInnerCriterium + 1)
                Set nAllCriteria(loopInnerCriterium + 1) = tempCriterium
            End If
        Next loopInnerCriterium
    Next loopOuterCriterium

End Sub

Public Sub ClearEvaluationSheet( _
    ByVal EvaluationSheet As Worksheet)
    
    Dim lastRow As Long
    
    lastRow = LastRowInColumnA(EvaluationSheet)
    
    EvaluationSheet.range(EvaluationSheet.Cells(1, 1), EvaluationSheet.Cells(lastRow, 3)).Delete xlShiftUp
End Sub

Public Sub WriteAllCriteria( _
    ByVal nEvaluationSheet As Worksheet, _
    ByRef nAllCriteria() As clsCriterium)
    
    Dim loopCriterium As Integer
    
    With nEvaluationSheet
    
        'Header schreiben
        .Cells(1, 1) = "Rang"
        .Cells(1, 2) = "Kriterien"
        .Cells(1, 3) = "Priorität"
        
        'Rahmen
        .Cells(1, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Cells(1, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        .range(nEvaluationSheet.Cells(1, 1), nEvaluationSheet.Cells(1, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
       
        'Ausrichten und färben
        .range(nEvaluationSheet.Cells(1, 1), nEvaluationSheet.Cells(1, 3)).HorizontalAlignment = xlCenter
        .range(nEvaluationSheet.Cells(1, 1), nEvaluationSheet.Cells(1, 3)).VerticalAlignment = xlCenter
        .range(nEvaluationSheet.Cells(1, 1), nEvaluationSheet.Cells(1, 3)).Interior.ColorIndex = 46
        
        For loopCriterium = 1 To UBound(nAllCriteria)
            
            'Schreiben
            .Cells(loopCriterium + 1, 1) = loopCriterium
            .Cells(loopCriterium + 1, 2) = nAllCriteria(loopCriterium).Name
            .Cells(loopCriterium + 1, 3) = nAllCriteria(loopCriterium).Priority
            
            'Rahmen
            .Cells(loopCriterium + 1, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Cells(loopCriterium + 1, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
            
            'Zentrieren
            .range(nEvaluationSheet.Cells(loopCriterium + 1, 1), nEvaluationSheet.Cells(loopCriterium + 1, 3)).HorizontalAlignment = xlCenter
            .range(nEvaluationSheet.Cells(loopCriterium + 1, 1), nEvaluationSheet.Cells(loopCriterium + 1, 3)).VerticalAlignment = xlCenter
            
            'Zellen einfärben
            If loopCriterium Mod 2 = 0 Then
                .range(nEvaluationSheet.Cells(loopCriterium + 1, 1), nEvaluationSheet.Cells(loopCriterium + 1, 3)).Interior.ColorIndex = 45
            Else
                .range(nEvaluationSheet.Cells(loopCriterium + 1, 1), nEvaluationSheet.Cells(loopCriterium + 1, 3)).Interior.ColorIndex = 44
            End If
            
        Next loopCriterium
        
    End With
End Sub

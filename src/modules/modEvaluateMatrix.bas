Attribute VB_Name = "modEvaluateMatrix"
Option Explicit

Public Sub EvaluateMatrix()
    Dim matrixWorksheet As Worksheet
    Dim EvaluationSheet As Worksheet
    
    Dim evaluationJudge As clsEvaluationHandler
    Dim error As Boolean
    Dim allCriteria() As clsCriterium
    
    Set matrixWorksheet = Sheet_Matrix
    Set EvaluationSheet = Sheet_Evaluation
    
    Set evaluationJudge = New clsEvaluationHandler
    
    
    error = False
    Application.EnableEvents = False
    
    Call evaluationJudge.ReadMatrix(matrixWorksheet, allCriteria, error)
    
    If Not error Then
        Call evaluationJudge.BubbleSort(allCriteria)
        
        EvaluationSheet.Unprotect
        Call evaluationJudge.ClearEvaluationSheet(EvaluationSheet)
        Call evaluationJudge.WriteAllCriteria(EvaluationSheet, allCriteria)
        EvaluationSheet.Protect
        
        EvaluationSheet.Select
        MsgBox "Matrix wurde erfolgreich ausgewertet!", , "Sucess!"
    End If
    Erase allCriteria
    Application.EnableEvents = True
End Sub


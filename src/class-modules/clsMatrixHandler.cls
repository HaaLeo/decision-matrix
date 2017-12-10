VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatrixHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub DeleteEmptyCells( _
    ByVal nWorksheet As Worksheet, _
    ByVal nLastRow As Long)

    'Ein Eintrag soll immer stehen bleiben
    If nLastRow > 1 Then
        'unteren Teil löschen
        nWorksheet.range(nWorksheet.Cells(nLastRow + 1, 1), nWorksheet.Cells(gOldLastRow, gOldLastRow)).Delete xlShiftUp
        
        'rechten Teil löschen
        nWorksheet.range(Cells(1, nLastRow + 1), Cells(nLastRow, gOldLastRow)).Delete xlShiftUp

    ElseIf gOldLastRow > 2 Then
        'unteren Teil löschen
        nWorksheet.range(Cells(3, 1), Cells(gOldLastRow, gOldLastRow)).Delete xlShiftUp
        'rechten Teil löschen
        nWorksheet.range(Cells(1, 3), Cells(gOldLastRow, gOldLastRow)).Delete xlShiftUp
        
        'Starttext mit geschützten Leerzeichen um bei UserInputValidation erkannt zu werden
        nWorksheet.Cells(2, 1) = "[In" & Chr(160) & "Spalte" & Chr(160) & "A" & Chr(160) & "Kriterien" & Chr(160) & "eintragen]"
        
    ElseIf gOldLastRow = 2 Then
        With nWorksheet
            .Cells(2, 1) = "[In" & Chr(160) & "Spalte" & Chr(160) & "A" & Chr(160) & "Kriterien" & Chr(160) & "eintragen]"
            
            'Hack to shape cells if duplicate is detected in first rows
            .Cells(2, 1).Interior.ColorIndex = 46
            .Cells(2, 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
    End If

End Sub

Public Sub ShapeBorderOfFilledCells( _
    ByVal nMatrixSheet As Worksheet, _
    ByVal nLastRow As Long)
    
    Dim loopRow As Long
    Dim lastColumn As Long
    
    'Weil Matrix symmetrisch ist
    lastColumn = gOldLastRow
    
    'Alle Rahmenlinien setzen
    nMatrixSheet.range(Cells(1, 1), Cells(nLastRow, nLastRow)).Borders.LineStyle = xlContinuous
    
    'Matrixinhalte zentrieren
    With nMatrixSheet.range(Cells(2, 2), Cells(nLastRow, nLastRow))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    'erste Zeile um  90 Grad drehen, färben, Spaltenbreite setzen und Text ausrichten
    With nMatrixSheet.range(Cells(1, 2), Cells(1, nLastRow))
        .Orientation = 90
        .Interior.ColorIndex = 46
        .ColumnWidth = 2.29
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    
    'Erste Spalte färben und Text ausrichten
    With nMatrixSheet.range(Cells(2, 1), Cells(nLastRow, 1))
        .Orientation = 0
        .Interior.ColorIndex = 46
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .RowHeight = 15.75
    End With

    For loopRow = gOldLastRow + 1 To nLastRow
        With nMatrixSheet
            '"diagonale" Zellen grau einfärben
            .Cells(loopRow, loopRow).Interior.ColorIndex = 16
            
            'Farbe den Zeilen und Spalten zuweisen
            If loopRow > 2 And loopRow Mod 2 = 0 Then
                .range(Cells(loopRow, 2), Cells(loopRow, lastColumn)).Interior.ColorIndex = 45
                .range(Cells(2, loopRow), Cells(lastColumn, loopRow)).Interior.ColorIndex = 45
            ElseIf loopRow > 2 Then
                .range(Cells(loopRow, 2), Cells(loopRow, lastColumn)).Interior.ColorIndex = 44
                .range(Cells(2, loopRow), Cells(lastColumn, loopRow)).Interior.ColorIndex = 44
            End If
        End With
        lastColumn = lastColumn + 1
    Next loopRow
        
End Sub

Public Sub AddFormula( _
    ByVal nMatrixSheet As Worksheet, _
    ByVal nLastRow As Long)
    
    Dim loopRow As Long
    Dim formula As String
    
    For loopRow = gOldLastRow + 1 To nLastRow
        'Zeilen mit 0 initialisieren
        nMatrixSheet.range(Cells(loopRow, 2), Cells(loopRow, loopRow - 1)) = 0
        
        'Formeln in senkrechte Zellen einfügen
        '"komplizierte" Formel um mit "runterziehen" arbeiten zu könnnen und Leistung zu sparen
        '=WENN(INDEX($B$10:$I$10;ZEILE(A1))=1;0;1)
        '"=IF(INDEX(R10C2:R10C9,ROW(R[-1]C[-10]))=1,0,1)"
        formula = "=IF(INDEX(R" + CStr(loopRow) + "C2:R" + CStr(loopRow) + "C" + CStr(loopRow - 1) + ",ROW(R[-1]C[-" + CStr(loopRow - 1) + "]))=1,0,1)"

        'Formel eintragen und "runterziehen"
        nMatrixSheet.Cells(2, loopRow).FormulaR1C1 = formula
        
        'Erst ab 3ter Spalte werden Formeln benötigt
        If loopRow > 3 Then
            nMatrixSheet.range(Cells(2, loopRow), Cells(loopRow - 1, loopRow)).FillDown
        End If
        'Erste Zeile mit Formeln befüllen (z.B. =A12)
        nMatrixSheet.Cells(1, loopRow).FormulaR1C1 = "=R" + CStr(loopRow) + "C[-" + CStr(loopRow - 1) + "]"
        
    Next loopRow
    
End Sub

Public Sub UpdateDataValidation( _
    ByVal nMatrixSheet As Worksheet, _
    ByVal nCurrentKriteriumRow As Long)
    
    Dim loopRow As Long
    Dim loopColumn As Long
    Dim lastRow As Long
    Dim endRowIndex As Long
    Dim startRowIndex As Long
    Dim currentKriterium As String
    
    lastRow = LastRowInColumnA(nMatrixSheet)
    currentKriterium = CStr(nMatrixSheet.Cells(nCurrentKriteriumRow, 1).Value)
    
    'Aktualisiere Spalte nur wenn nichts erweitert oder gelöscht wurde
    If gOldLastRow = lastRow Then
        With nMatrixSheet.range(Cells(nCurrentKriteriumRow + 1, nCurrentKriteriumRow), Cells(lastRow, nCurrentKriteriumRow)).validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0,1"
            .IgnoreBlank = False
            .InCellDropdown = False
            .InputTitle = "Eingabehinweis"
            .ErrorTitle = "Ungültige Eingabe!"
            .ErrorMessage = "Es werden nur die Werte ""0"" und ""1"" akzeptiert"
            .ShowInput = True
            .ShowError = True
        End With
        For loopRow = nCurrentKriteriumRow + 1 To lastRow
                nMatrixSheet.Cells(loopRow, nCurrentKriteriumRow).validation.InputMessage = "Ist """ & nMatrixSheet.Cells(loopRow, 1) & """ wichtiger als """ & currentKriterium & """?" & vbLf & _
                    "[ja] = 1" & vbLf & "[nein] = 0"
        Next loopRow
        
    End If
    
    'erstes Element muss nicht geupdated werden
    If nCurrentKriteriumRow > 2 Then
        'aktualisiere Validation der Zeile
        With nMatrixSheet.range(Cells(nCurrentKriteriumRow, 2), Cells(nCurrentKriteriumRow, nCurrentKriteriumRow - 1)).validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0,1"
            .IgnoreBlank = False
            .InCellDropdown = False
            .InputTitle = "Eingabehinweis"
            .ErrorTitle = "Ungültige Eingabe!"
            .ErrorMessage = "Es werden nur die Werte ""0"" und ""1"" akzeptiert"
            .ShowInput = True
            .ShowError = True
        End With
        For loopColumn = 2 To nCurrentKriteriumRow - 1
            nMatrixSheet.Cells(nCurrentKriteriumRow, loopColumn).validation.InputMessage = "Ist """ & currentKriterium & """ wichtiger als """ & nMatrixSheet.Cells(loopColumn, 1).Text & """?" & vbLf & _
                "[ja] = 1" & vbLf & "[nein] = 0"
        Next loopColumn
        nMatrixSheet.range(Cells(nCurrentKriteriumRow, 2), Cells(nCurrentKriteriumRow, nCurrentKriteriumRow - 1)).Locked = False
    End If
End Sub

Public Sub ValidateUserInput( _
    ByVal nMatrixSheet As Worksheet, _
    ByRef nUserInput As String, _
    ByVal nCurrentKriteriaRow As Long, _
    ByRef nLastRow As Long, _
    ByRef nIsDuplicate As Boolean)
    
    Dim result As range
    Dim completeMatrix As range
    Dim FoundCell As range
    Dim LastCell As range
    Dim FirstAddr As String
    Dim duplicateCells As String
    Dim test As Long
    
    Set completeMatrix = nMatrixSheet.range(Cells(2, 1), Cells(nLastRow, 1))
   'Suchen
    With completeMatrix
        Set LastCell = .Cells(.Cells.Count)
    End With
    
    Set FoundCell = completeMatrix.Find(what:=nUserInput, after:=LastCell, lookat:=xlWhole)
    
    If Not FoundCell Is Nothing Then
        FirstAddr = FoundCell.Address
        Set result = FoundCell
        duplicateCells = "A" & FoundCell.Row
    End If
    
    Do Until FoundCell Is Nothing
        Set FoundCell = completeMatrix.FindNext(after:=FoundCell)
        Set result = Union(result, FoundCell)
        If FoundCell.Address = FirstAddr Then
            Exit Do
        Else
            duplicateCells = duplicateCells & ", A" & FoundCell.Row
            
        End If
    Loop

    'Ergebnis auswerten
    If result Is Nothing Then
        'nichts gefunden
    ElseIf result.Count < 2 Then
        'nur sich selbst gefunden
    Else
        nIsDuplicate = True
        If Not InStr(duplicateCells, ",") = 0 Then
            duplicateCells = Left(duplicateCells, InStrRev(duplicateCells, ",") - 1) & Replace(duplicateCells, ",", " und", InStrRev(duplicateCells, ","), 1)
        End If
        
        result.Select
        Cells(nCurrentKriteriaRow, 1).Activate
        MsgBox """" & nUserInput & """ steht bereits in den Zellen " & duplicateCells & "." & vbLf & "Ein Kriterium darf nur genau EINMAL in der Matrix enthalten sein!" & vbLf & "Der Inhalt dieser Zellen wird gelöscht", , "Duplikat!"
        
        result = ""
        'Wenn Duplikat außerhalb der Matrix ist muss Format angepasst werden
        If nCurrentKriteriaRow > gOldLastRow Then
            With result
                .Interior.ColorIndex = xlNone
                .Borders.LineStyle = xlNone
            End With
        End If
        nLastRow = LastRowInColumnA(nMatrixSheet)
    End If

End Sub


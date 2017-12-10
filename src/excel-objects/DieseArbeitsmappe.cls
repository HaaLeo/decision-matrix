VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Initialisierung der globalen Variabeln
Private Sub Workbook_Open()
    Sheet_Matrix.Activate
    gOldLastRow = LastRowInColumnA(Sheet_Matrix)
End Sub

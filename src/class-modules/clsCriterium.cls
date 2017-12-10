VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCriterium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' defines the priority of the kriterium
Private mPriority As Long

' the name of the Kriterium
Private mName As String

' Properties
' ==========

Public Property Get Priority() As Long
    Priority = mPriority
End Property

Public Property Let Priority(ByVal nPriority As Long)
    mPriority = nPriority
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal nName As String)
    mName = nName
End Property





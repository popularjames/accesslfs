Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Public Enum ExpTypeVal
    Base = 0
    Leaf = 1
End Enum

Public ExpTable As String 'Name of table being scripted
Public ExpTyp As ExpTypeVal
Public ExpScript As String
Public ExpLevel As Integer
Public ExpName As String
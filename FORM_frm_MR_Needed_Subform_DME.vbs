Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const maxDME1454 As Long = 2048

Public Property Get maxSelected() As Long
    maxSelected = maxDME1454
End Property

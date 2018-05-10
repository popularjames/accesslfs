Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const maxSelectedHH As Long = 65536

Public Property Get maxSelected() As Long
    maxSelected = maxSelectedHH
End Property

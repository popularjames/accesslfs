Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const maxPharm As Long = 131072

Public Property Get maxSelected() As Long
    maxSelected = maxPharm
End Property

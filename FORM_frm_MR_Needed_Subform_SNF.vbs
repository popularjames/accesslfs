Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const maxSNF As Long = 8192

Public Property Get maxSelected() As Long
    maxSelected = maxSNF
End Property

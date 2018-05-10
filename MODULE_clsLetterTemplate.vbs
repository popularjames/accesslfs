Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mvarLetterType As String
Private mvarTemplateLoc As String
Private mlngContractId As Long


Public Property Let TemplateLoc(ByVal vData As String)
    mvarTemplateLoc = vData
End Property

Public Property Get TemplateLoc() As String
    TemplateLoc = mvarTemplateLoc
End Property

Public Property Let LetterType(ByVal vData As String)
    mvarLetterType = UCase(vData)
End Property

Public Property Get LetterType() As String
    LetterType = UCase(mvarLetterType)
End Property

Public Property Get ContractId() As Long
    ContractId = mlngContractId
End Property
Public Property Let ContractId(lngContractId As Long)
    mlngContractId = lngContractId
End Property

Public Sub test()
    MsgBox "test"
End Sub
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private csErrMsg As String
Private csErrDetails As String
Private clErrNum As Long
Private csErrProc As String
Private cbFatal As Boolean
Private csInstnaceId As String
Private clBatchId As Long
Private csAuditor As String
Private csLetterType As String
Private csCnlyClaimNums As String



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get ErrorMessage() As String
    ErrorMessage = csErrMsg
End Property
Public Property Let ErrorMessage(sErrMsg As String)
    csErrMsg = sErrMsg
End Property

Public Property Get ErrorDetails() As String
    ErrorDetails = csErrDetails
End Property
Public Property Let ErrorDetails(sErrDetails As String)
    csErrDetails = sErrDetails
End Property


Public Property Get ErrorProc() As String
    ErrorProc = csErrProc
End Property
Public Property Let ErrorProc(sErrProc As String)
    csErrProc = sErrProc
End Property


Public Property Get ErrorNum() As Long
    ErrorNum = clErrNum
End Property
Public Property Let ErrorNum(lErrNum As Long)
    clErrNum = lErrNum
End Property



Public Property Get FatalError() As Boolean
    FatalError = cbFatal
End Property
Public Property Let FatalError(bFatal As Boolean)
    cbFatal = bFatal
End Property



Public Property Get InstanceID() As String
    InstanceID = csInstnaceId
End Property
Public Property Let InstanceID(sErrProc As String)
    csInstnaceId = sErrProc
End Property

Public Property Get CnlyClaimNums() As String
    CnlyClaimNums = csCnlyClaimNums
End Property
Public Property Let CnlyClaimNums(sCnlyClaimNums As String)
    csCnlyClaimNums = sCnlyClaimNums
End Property

Public Property Get BatchID() As Long
    BatchID = clBatchId
End Property
Public Property Let BatchID(lBatchId As Long)
    clBatchId = lBatchId
End Property


Public Property Get LetterType() As String
    LetterType = csLetterType
End Property
Public Property Let LetterType(sLetterType As String)
    csLetterType = sLetterType
End Property


Public Property Get Auditor() As String
    Auditor = csAuditor
End Property
Public Property Let Auditor(sAuditor As String)
    csAuditor = sAuditor
End Property
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private csConceptId As String
Private ciPayerNameId As Integer
Private csPayerName As String
Private csSelectedLetterType As String
Private cbCanceled As Boolean


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

'' This property is generic for some generic functions like KDShowFormAndWait:

Public Property Get SelectedId() As String
    SelectedId = SelectedLetterType
End Property
Public Property Let SelectedId(sSelectedLetterType As String)
    SelectedLetterType = sSelectedLetterType

    Me.TimerInterval = 250
    
End Property



Public Property Get SelectedLetterType() As String
    SelectedLetterType = csSelectedLetterType
End Property
Public Property Let SelectedLetterType(sSelectedLetterType As String)
    csSelectedLetterType = sSelectedLetterType
End Property

Public Property Get Canceled() As Boolean
    Canceled = cbCanceled
End Property
Public Property Let Canceled(bUserCanceled As Boolean)
    cbCanceled = bUserCanceled
End Property


Public Property Get ConceptID() As String
    ConceptID = csConceptId
End Property
Public Property Let ConceptID(sConceptId As String)
    csConceptId = sConceptId
    Me.lblConceptId.Caption = sConceptId
End Property


Public Property Get PayerName() As String
    If csPayerName = "" Then
        csPayerName = GetPayerNameFromID(Me.PayerNameId)
    End If
    PayerName = csPayerName
End Property
Public Property Let PayerName(sPayerName As String)
    csPayerName = sPayerName
End Property


Public Property Get PayerNameId() As Integer
    PayerNameId = ciPayerNameId
End Property
Public Property Let PayerNameId(iPayerNameId As Integer)
    ciPayerNameId = iPayerNameId
    If iPayerNameId > 1000 Then
        Me.lblPayerName.Caption = " and Payer: " & Me.PayerName
    Else
        Me.lblPayerName.Caption = " "
    End If
End Property




Private Sub CmdCancel_Click()
On Error GoTo Block_Exit
Dim strProcName As String
    strProcName = ClassName & ".cmdCancel_Click"
    
    Me.SelectedLetterType = ""
    Me.Canceled = True
    
    Me.visible = False
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdUse_Click()
On Error GoTo Block_Exit
Dim strProcName As String

    strProcName = ClassName & ".cmdUse_Click"
Stop
    ' if nothing is selected tell them!
       
    If Me.lstbAdrLetterTypes.ListIndex < 0 Then
        LogMessage strProcName, "USER ERROR", "Please select the letter to use first!", , True
        GoTo Block_Exit
    End If
    
    Me.SelectedLetterType = Nz(Me.lstbAdrLetterTypes.Column(0, Me.lstbAdrLetterTypes.ListIndex + 1), "")
    If Me.SelectedLetterType = "" Then
        LogMessage strProcName, "ERROR", "For some reason we can't get the letter type you selected - cancel and try again please?!", , , csConceptId
        GoTo Block_Exit
    End If
    
    Me.visible = False
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdViewSample_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sDocPath As String
Dim oWordApp As Word.Application
Dim oWordDoc As Word.Document
Dim sTempLoc As String

    strProcName = ClassName & ".cmdViewSample_Click"
       
    If Me.lstbAdrLetterTypes.ListIndex < 0 Then
        LogMessage strProcName, "USER ERROR", "Please select the letter to preview first!", , True
        GoTo Block_Exit
    End If
    
    sDocPath = Nz(Me.lstbAdrLetterTypes.Column(2, Me.lstbAdrLetterTypes.ListIndex + 1), "")
    
    If sDocPath = "" Then
        LogMessage strProcName, "ERROR", "Could not determine sample document path for this concept", , True, Me.ConceptID
        GoTo Block_Exit
    End If
    

    sTempLoc = GetUniqueFilename(, , FileExtension(sDocPath))
    If CopyFile(sDocPath, sTempLoc, False) = False Then
        LogMessage strProcName, "ERROR", "Could not copy the sample document to your temp folder", sTempLoc, True, Me.ConceptID
        GoTo Block_Exit
    End If
    
    Set oWordApp = New Word.Application
        ' open it read only!!!!
    Set oWordDoc = oWordApp.Documents.Open(sTempLoc, , True)
    oWordApp.visible = True
    
    oWordApp.Activate
    Call AppActivate(oWordApp.ActiveDocument.Name, False)
    Call SendKeys("^a", False)
    Sleep 500
    Call SendKeys("%{F9}", False)
    oWordApp.Activate
    oWordApp.ShowMe
    
Block_Exit:
    '' Don't need to quit because we set it as visible = True
    Set oWordApp = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub Form_Timer()
Dim iIndex As Integer
Dim sLtrType As String

    Me.TimerInterval = 0
    sLtrType = Nz(Me.SelectedLetterType, "")
    If sLtrType <> "" Then
        ' select the type in our data
        For iIndex = 0 To Me.lstbAdrLetterTypes.ListCount
            If Me.lstbAdrLetterTypes.Column(0, iIndex) = sLtrType Then
                'Me.lstbAdrLetterTypes.Column(0, iIndex).Selected = True
                Me.lstbAdrLetterTypes.ListIndex = iIndex - 1
                Exit For
            End If
        Next
    End If

    
End Sub

Private Sub lstbAdrLetterTypes_Click()
'Debug.Print "Click event"
'
'    Call ItemSelected
End Sub

Private Sub lstbAdrLetterTypes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print "Mouse up"
'    Call ItemSelected
End Sub

Public Sub ItemSelected()
On Error GoTo Block_Err
'Dim strProcName As String
'Dim oRs As DAO.Recordset
'Dim sDocPath As String
'
'    strProcName = ClassName & ".ItemSelected"
'
'
''Use listboxControl.Column(intColumn,intRow). Both Column and Row are zero-based.
'    If Me.lstbAdrLetterTypes.ListIndex > 0 Then
'        Debug.Print Me.lstbAdrLetterTypes.Column(0, Me.lstbAdrLetterTypes.ListIndex)
'        sDocPath = Nz(Me.lstbAdrLetterTypes.Column(3, Me.lstbAdrLetterTypes.ListIndex), "")
'    Stop
'
'    Else
'        Debug.Print "nothing selected"
'        Stop
'        GoTo Block_Exit
'    End If
'    If Me.lstbAdrLetterTypes.ItemsSelected.Count = 0 Then
'        Me.oOLE.SourceDoc = ""
'        Stop
'        GoTo Block_Exit
'    End If
'Debug.Print Me.lstbAdrLetterTypes.Column(0, Me.lstbAdrLetterTypes.ListIndex)
'    sDocPath = Me.lstbAdrLetterTypes.Column(3, Me.lstbAdrLetterTypes.ListIndex)
'
'
''
''    Set oRs = Me.lstbAdrLetterTypes.Recordset
''    sDocPath = Nz(oRs("SampleDocPath").Value, "")
''    If sDocPath = "" Then
''        Stop
''    End If
'
'    Me.oOLE.SourceDoc = sDocPath
'    Me.oOLE.Action = 1
'
    
    
Block_Exit:
    Exit Sub
Block_Err:
'    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

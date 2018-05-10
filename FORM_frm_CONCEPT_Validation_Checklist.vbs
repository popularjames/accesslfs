Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


''' Last Modified: 07/19/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 07/19/2012 - added payer column
'''  - 04/19/2012 - Created...
'''
''' AUTHOR
'''  =====================================
''' Kevin Dearing
'''
'''
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################

Private csConceptId As String
Private coConcept As clsConcept


Private cbValidationFailed As Boolean

Private Const lBackColorGood As Long = 12189625
Private Const lBackColorBad As Long = 2960895


Private coRs As ADODB.RecordSet


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

    ' #############################################
Public Property Get ConceptID() As String
    ConceptID = csConceptId
End Property
Public Property Let ConceptID(sConceptId As String)
    csConceptId = sConceptId
    Me.txtConceptID = sConceptId
End Property

    ' #############################################
Public Property Get ValidationFailed() As Boolean
    ValidationFailed = cbValidationFailed
End Property
Public Property Let ValidationFailed(bValidationFailed As Boolean)
    cbValidationFailed = bValidationFailed
    Select Case bValidationFailed
    Case True
        Me.lblColorIndicator.BackColor = lBackColorBad
    Case False
        Me.lblColorIndicator.BackColor = lBackColorGood
    End Select
    Me.Repaint
End Property



    ' #############################################
Public Property Get ValidationReport() As ADODB.RecordSet
    Set ValidationReport = coRs
End Property
Public Property Let ValidationReport(oRs As ADODB.RecordSet)
    Set coRs = oRs
End Property



Public Sub ShowReport(oRs As ADODB.RecordSet)
    ValidationReport = oRs
    Call ValidateAndDisplay
End Sub


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub cmdRefresh_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRegEx As RegExp

    strProcName = ClassName & ".cmdRefresh_Click"
    
    Set oRegEx = New RegExp
    oRegEx.Pattern = "^CM\_C\d{4}$"
    oRegEx.IgnoreCase = True
    
    If oRegEx.test(CStr("" & Me.txtConceptID)) = False Then
        MsgBox "Invalid Concept ID! Please check and try again", vbCritical, "Error"
        GoTo Block_Exit
    End If
    
    Me.ConceptID = Me.txtConceptID
    
    Call ValidateAndDisplay

Block_Exit:
    Set oRegEx = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub


Private Sub ValidateAndDisplay()
On Error GoTo Block_Err
Dim strProcName As String
Dim sValidateRpt As String
Dim bReValidate As Boolean

    strProcName = ClassName & ".ValidateAndDisplay"
    DoCmd.Hourglass True
    DoCmd.Echo True, "Validating..."
    
    
'    If coConcept Is Nothing Or csConceptId <> Me.txtConceptID Then
        Set coConcept = New clsConcept
        If coConcept.LoadFromId(csConceptId) = False Then
            LogMessage strProcName, "ERROR", "Concept not found!: " & csConceptId, csConceptId, True
            ValidationFailed = True
            GoTo Block_Exit
        End If
        
        If coRs Is Nothing Then
            bReValidate = True
            GoTo ValidateNow    ' have to do this to skip the error for the next check.
        End If
        If coRs.recordCount = 0 Then
            bReValidate = True
        End If
ValidateNow:
        If bReValidate = True Then
            ValidationFailed = Not coConcept.ValidateForSubmission(coRs)
        Else
            coRs.MoveFirst
            While Not coRs.EOF
                If coRs("Success").Value = False Then
                    ValidationFailed = True
                End If
                coRs.MoveNext
            Wend
        End If
        
        
        Call ParseRpt
        
'    End If
    If ValidationFailed = True Then
        MsgBox "Concept is NOT ready to submit", vbCritical, "Validation Failed"
    Else
        MsgBox "This concept is ready to submit", vbOKOnly, "Validation Success!"
    End If
        

Block_Exit:
    DoCmd.Hourglass False
    DoCmd.Echo True, "Ready..."

    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub Form_Load()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".Form_Load"
    
'    Call cmdRefresh_Click

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub ParseRpt()
On Error GoTo Block_Err
Dim strProcName As String
'Dim sarItems() As String
Dim iIdx As Integer
'    Dim oLView As ListView
'    Dim oLItem As ListItem

Dim oLView As Object
Dim oLItem As Object


    strProcName = ClassName & ".ParseRpt"
    
    Set oLView = Me.lvwCheckList
    oLView.ListItems.Clear
    
    If coRs Is Nothing Then GoTo Block_Exit
    If coRs.recordCount < 1 Then GoTo Block_Exit
    coRs.MoveFirst
    While Not coRs.EOF
        Set oLItem = oLView.ListItems.Add(coRs("RowId").Value, , coRs("Item Checked").Value)
'        oLItem.Selected = IIf(coRS("Success").Value = 0, False, True)
        oLItem.Checked = IIf(coRs("Success").Value = 0, False, True)
        oLItem.SubItems(1) = coRs("PayerName").Value
        oLItem.SubItems(2) = coRs("Notes").Value
        coRs.MoveNext
    Wend
    
    

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub

Private Sub Form_Resize()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".Form_Resize"
    
    Me.lblColorIndicator.Width = Me.InsideWidth
    Me.lblColorIndicator.Height = Me.InsideHeight
    Me.lvwCheckList.Height = Me.InsideHeight - Me.lvwCheckList.top - Me.lvwCheckList.left   ' a little bit of the lable peeking out
    Me.lvwCheckList.Width = Me.InsideWidth - (Me.lvwCheckList.left * 2)   ' a little space around it
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub lvwCheckList_ColumnClick(ByVal ColumnHeader As Object)
'    Dim oLView As ListView
'    Dim oLItem As ListItem

Dim oLView As Object
Dim oLItem As Object



    
    Set oLView = Me.lvwCheckList
    oLView.Sorted = False
    oLView.SortKey = ColumnHeader.index - 1
    oLView.SortOrder = IIf(oLView.SortOrder = lvwDescending, lvwAscending, lvwDescending)
    oLView.Sorted = True

End Sub



Private Sub txtConceptID_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call cmdRefresh_Click
    End If
End Sub

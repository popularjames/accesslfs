Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'''  Represents a CMS Concept, basically a "hook" into the
'''     _CLAIMS.dbo.CONCEPT_Hdr table
'''  With validation and various other methods..
'''
'''  TODO:
'''  =====================================
'''  - Maybe make actual properties instead of using GetField()
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 07/19/2012 - added payername
'''  - 03/14/2012 - Created class
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

Private coRs As ADODB.RecordSet

Private ciRows As Integer
Private csConceptId As String


    ''' ##############################################################################
Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


    ''' ##############################################################################
Public Property Get ConceptID() As String
    ConceptID = csConceptId
End Property
Public Property Let ConceptID(sConceptId As String)
    csConceptId = sConceptId
End Property
        '' Just an alias for ease of use!
    Public Property Get ID() As String
        ID = ConceptID
    End Property
    Public Property Let ID(sNewId As String)
        ConceptID = sNewId
    End Property



    ''' ##############################################################################
Public Property Get RowCount() As Integer
    RowCount = ciRows
End Property

    ''' ##############################################################################
Public Property Get GetRecordset() As ADODB.RecordSet
    Set GetRecordset = coRs
End Property



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function AddNote(bSuccess As Boolean, sItemDesc As String, sPayerName As String, sNotes As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
    
    strProcName = ClassName & ".AddNote"
    
    ciRows = ciRows + 1
    
    coRs.AddNew
    coRs("RowId") = ciRows
    coRs("Success") = IIf(bSuccess, -1, 0)
    coRs("Item Checked") = sItemDesc
    coRs("PayerName") = sPayerName
    coRs("Notes") = sNotes
    coRs.Update
    
    AddNote = True

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    ciRows = ciRows - 1
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub AddFieldsToRS()
On Error GoTo Block_Err
Dim strProcName As String
    
    strProcName = ClassName & ".AddFieldsToRS"
    
    Set coRs = New ADODB.RecordSet
    With coRs
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        Set .ActiveConnection = Nothing

        .Fields.Append "RowId", adInteger, 1
        .Fields.Append "Success", adInteger, 1
        .Fields.Append "Item Checked", adLongVarWChar, 1
        .Fields.Append "PayerName", adLongVarWChar, 1
        .Fields.Append "Notes", adLongVarWChar, 1
        .Open
    End With
    

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



'########################################################################################################
'########################################################################################################
'########################################################################################################
'########################################################################################################
'
'       Class Init / Term
'
'########################################################################################################
'########################################################################################################
'########################################################################################################
'########################################################################################################


Private Sub Class_Initialize()
    Call AddFieldsToRS
End Sub


Private Sub Class_Terminate()
   Set coRs = Nothing
End Sub
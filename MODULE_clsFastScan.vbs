Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Public Event FastScanError(ErrMsg As String, ErrNum As Long, ErrSource As String)

Private mrsCoverSheet As ADODB.RecordSet
Private mrsNotes As ADODB.RecordSet

Private mAppID As String
Private mCoverSheetNum As String
Private mbCoverSheetExists As Boolean
Private mNoteID As Long
Private mbLockedForEdit As Boolean
Private mstrLockedUser As String
Private mdLockedDt As Date
Private mstrCurrentUser As String

Public Property Let CoverSheetNum(ByVal vData As String)
    mCoverSheetNum = vData
End Property
Public Property Get CoverSheetNum() As String
    If Nz(mCoverSheetNum, "") = "" Then mCoverSheetNum = "NA"
    CoverSheetNum = mCoverSheetNum
End Property
Public Property Get rsCoverSheet() As ADODB.RecordSet
    Set rsCoverSheet = mrsCoverSheet
End Property
Public Property Get rsNotes() As ADODB.RecordSet
    
    '8/17/2012
    '************DPR Moved the refresh of these objects out of the initial load.  Only setting them when needed
    
    If mrsNotes Is Nothing Then
        MyAdo.sqlString = " SELECT * from FastScanMaint.v_CA_NOTE_Detail where NoteID = " & mNoteID
        Set mrsNotes = MyAdo.OpenRecordSet
        Set rsNotes = mrsNotes
    Else
        Set rsNotes = mrsNotes
    End If
    


End Property

Public Property Get LockedForEdit() As Boolean
     LockedForEdit = mbLockedForEdit
End Property

Public Property Get LockedUser() As String
     LockedUser = mstrLockedUser
End Property

Public Property Get LockedDate() As Date
     LockedDate = mdLockedDt
End Property

Public Property Get CoverSheetExists() As Boolean
    CoverSheetExists = mbCoverSheetExists
End Property

Public Function LoadCoverSheet(strCoverSheetNum As String) As Boolean
'Damon 7/15/08
'Loaads a claim based on what is passed to the method
    On Error GoTo ErrHandler
        
    Dim strErrSource As String
    strErrSource = "clsFastScan_LoadClaim"
    
   
    mbLockedForEdit = False
    mbCoverSheetExists = False
    
    'set the objects claim number to the passed in claim
    Me.CoverSheetNum = strCoverSheetNum
    
    MyAdo.sqlString = " SELECT * from Scanning_FastScan_Log WHERE CoverSheetNum = '" & strCoverSheetNum & "' and AccountID = " & gintAccountID
    'open the audit claims header and disconnect
    Set mrsCoverSheet = MyAdo.OpenRecordSet()
    
    If Not (mrsCoverSheet.BOF And mrsCoverSheet.EOF) Then
        
                'get the noteid associated with the claim
        mNoteID = Nz(mrsCoverSheet("NoteID"), -1)
        
        LoadCoverSheet = True
        mbCoverSheetExists = True
    Else
        mbCoverSheetExists = False
        LoadCoverSheet = False
    End If
    
Exit_Function:
    Exit Function
    
ErrHandler:
    RaiseEvent FastScanError(Err.Description, Err.Number, strErrSource)
    Resume Exit_Function
End Function

Public Function UnLoadCoverSheet() As Boolean
'Damon 7/15/08
'Loaads a claim based on what is passed to the method
    On Error GoTo ErrHandler
        
    Dim strErrSource As String
    strErrSource = "clsFastScan_UnLoadClaim"
    
   
    mbLockedForEdit = False
    mbCoverSheetExists = False
    
    'set the objects claim number to the passed in claim
    Me.CoverSheetNum = "NA"
    
    MyAdo.sqlString = " SELECT [CoverSheetNum] = 'NA', [ReceivedDt] = '', [ReceivedMeth] = '', [TrackingNum] = '', [ProviderFolder] = '', [ImageName] ='', [ScannedDt] = '', [ProcStatusCd] = '', [ProcStatusLastUpDt] = '', [ProcStatusLastUserID] = '', [NoMatchReasonCd] = '', [ADR2DBarCode] = '' "
    'open the audit claims header and disconnect
    Set mrsCoverSheet = MyAdo.OpenRecordSet()
    
    mbCoverSheetExists = False
    UnLoadCoverSheet = True
    
Exit_Function:
    Exit Function
    
ErrHandler:
    RaiseEvent FastScanError(Err.Description, Err.Number, strErrSource)
    Resume Exit_Function
End Function


Public Function ValidateCoverSheet() As Boolean
    ValidateCoverSheet = True
End Function

Public Function SaveCoverSheet() As Boolean
    
'8/17/2012
'DPR change the save events to only save things that are loaded.  Do not call savs for objects that are not loaded
    
    Dim bResult As Boolean
    Dim strErrSource As String
    Dim iResult As Integer
    Dim strErrMsg As String

    On Error GoTo ErrHandler
    
    
    strErrSource = "clsFastScan_SaveCoverSheet"
    
    myCode_ADO.BeginTrans
    bResult = False

    ' check claim data to make sure it is OK to save
    If Me.ValidateCoverSheet() = False Then
        bResult = False
        Err.Raise 65000, strErrSource, "CoverSheet could not be validated.  Record not saved."
    Else
        bResult = True
    End If
    
    If bResult Then
        ' save the notes
        If Not mrsNotes Is Nothing Then
            bResult = SaveData_Notes
            If bResult = False Then
                Err.Raise 65000, strErrSource, "Error saving claim notes.  Record not saved"
            End If
        Else
            bResult = True
        End If
    End If
    
 
    
    If bResult Then
        ' save claim header
        mrsCoverSheet("ProcStatusLastUpDt") = Now
        mrsCoverSheet("ProcStatusLastUserID") = mstrCurrentUser
        bResult = myCode_ADO.Update(mrsCoverSheet, "FastScan.usp_Fastscan_Log_Update")
        
        If bResult = False Then
            Err.Raise 65000, strErrSource, "Error saving claim header data.  Record not saved"
        End If
        
        'We just saved the recordset, so let's move to the beginning so we can access its values
        If mrsCoverSheet.EOF Then
            mrsCoverSheet.MoveFirst
        End If
    End If
    
      
    myCode_ADO.CommitTrans
    SaveCoverSheet = bResult
    
Exit_Sub:
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    RaiseEvent FastScanError(Err.Description, Err.Number, strErrSource)
    myCode_ADO.RollbackTrans
    SaveCoverSheet = bResult
    GoTo Exit_Sub
End Function



Private Function SaveData_Notes() As Boolean
    Dim bResult As Boolean
    
    On Error GoTo ErrHandler
    
    If Not (mrsNotes.BOF = True And mrsNotes.EOF = True) Then
        'If the noteID is -1 then we need to create a new ID
        If mNoteID = -1 Then
            'This is a public function that gets a unique ID based on the app being passed to the method
            mNoteID = GetAppKey("NOTE")
            'Set the recordset of the header to contain the new note ID
            Me.UpdateField "NoteID", mNoteID
            'Apply this new noteID to all of the records in the note recordset
            If Not (mrsNotes.BOF = True And mrsNotes.EOF = True) Then
                mrsNotes.MoveFirst
                While Not mrsNotes.EOF
                    mrsNotes.Update
                    mrsNotes("NoteID") = mNoteID
                    mrsNotes.MoveNext
                Wend
            End If
        End If
        'Pass the recordset back to SQL synching the results
        bResult = myCode_ADO.Update(mrsNotes, "usp_NOTE_Detail_Apply")
        
        If bResult = False Then
            Err.Raise 65000, "", "Error updating claim note"
        End If
   Else
        bResult = True
    End If
    
    SaveData_Notes = bResult

Exit_Sub:
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    SaveData_Notes = False
    GoTo Exit_Sub
End Function


Public Function UpdateField(FieldName As String, FieldValue As Variant)
    Dim strErrSource As String
    
    On Error GoTo ErrHandler
    strErrSource = "clsFastScan_UpdateField"
    mrsCoverSheet(FieldName).Value = FieldValue
    mrsCoverSheet.UpdateBatch adAffectAllChapters 'must do this to avoid recordset not syncing when displaying on forms
Exit Function

ErrHandler:
    RaiseEvent FastScanError(Err.Description, Err.Number, strErrSource)
End Function

Private Sub Class_Initialize()
    Set MyAdo = New clsADO
    Set myCode_ADO = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
    myCode_ADO.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
    
    mAppID = "FastScan"
    mstrCurrentUser = Identity.UserName
    
End Sub
Private Sub Class_Terminate()
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set mrsCoverSheet = Nothing
    Set mrsNotes = Nothing
End Sub


Public Function LockCoverSheet() As Boolean
           
    On Error GoTo ErrHandler
           
    Dim iResult As Long
    Dim LocCmd As New ADODB.Command
    Dim strErrMsg As String
        
    Dim strErrSource As String
    strErrSource = "clsFastScan_LockClaim"
    
    'set the objects claim number to the passed in claim
    Me.CoverSheetNum = Me.CoverSheetNum
    
    ' lock claim for edit
    myCode_ADO.sqlString = "FastScan.usp_FastScan_Lock"
    myCode_ADO.SQLTextType = StoredProc

    Set LocCmd = New ADODB.Command
    LocCmd.ActiveConnection = myCode_ADO.CurrentConnection
    LocCmd.CommandText = myCode_ADO.sqlString
    LocCmd.commandType = adCmdStoredProc
    LocCmd.Parameters.Refresh
            
    LocCmd.Parameters("@pCoverSheetNum") = Me.CoverSheetNum
    LocCmd.Parameters("@pLockUser") = Identity.UserName()
    LocCmd.Parameters("@pLockTime") = Now()
   

    iResult = myCode_ADO.Execute(LocCmd.Parameters)
    iResult = LocCmd.Parameters("@RETURN_VALUE")
    strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
    
    If iResult <> 0 Then
        MsgBox strErrMsg
        LockCoverSheet = False
        GoTo ErrHandler
    End If
    
    LockCoverSheet = True
Exit Function

ErrHandler:
    LockCoverSheet = False
    RaiseEvent FastScanError(Err.Description, Err.Number, strErrSource)
    
End Function
Public Function UnlockCoverSheet(Optional MsgSuppress As Boolean = False) As Boolean
           
    On Error GoTo ErrHandler
           
    Dim iResult As Long
    Dim LocCmd As New ADODB.Command
    Dim strErrMsg As String
        
    Dim strErrSource As String
    strErrSource = "clsFastScan_UnLockClaim"
    
    'set the objects claim number to the passed in claim
    Me.CoverSheetNum = Me.CoverSheetNum
    
    'lock claim for edit
    myCode_ADO.Connect
    myCode_ADO.sqlString = "FastScan.usp_FastScan_UnLock"
    myCode_ADO.SQLTextType = StoredProc

    Set LocCmd = New ADODB.Command
    LocCmd.ActiveConnection = myCode_ADO.CurrentConnection
    LocCmd.CommandText = myCode_ADO.sqlString
    LocCmd.commandType = adCmdStoredProc
    LocCmd.Parameters.Refresh
            
    LocCmd.Parameters("@pCoverSheetNum") = Me.CoverSheetNum
    LocCmd.Parameters("@pLockUserID") = Identity.UserName()
    
    
    iResult = myCode_ADO.Execute(LocCmd.Parameters)
    iResult = LocCmd.Parameters("@RETURN_VALUE")
    strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
    
    Select Case iResult
        Case Is <> 0
            If MsgSuppress = False Then
                ' this is to suppress message when use exit claim screen
                MsgBox "CoverSheet is locked by " & strErrMsg
            End If
            UnlockCoverSheet = False
        Case Else
            UnlockCoverSheet = True
    End Select
Exit Function

ErrHandler:
    UnlockCoverSheet = False
    RaiseEvent FastScanError(Err.Description, Err.Number, strErrSource)
    
End Function

Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "frm_FastScan_Main : ADO Error"
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "frm_FastScan_Main : ADO Error"
End Sub
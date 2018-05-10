Option Compare Database
Option Explicit




''' Last Modified: 05/20/2015
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  This is the main "automator"
'''     That will do all of the work!
'''
'''  TODO:
'''  =====================================
'''  - Maybe make actual properties instead of using GetField()
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 05/20/2015 - KD: tweaking existing functionality and adding some objects for the upcoming Letter Rules feature.
'''  - 04/22/2015 - KD: added reprint functionality to the processor (just set it to status = 'R' and Held = 0 to have it be reprocessed
'''     without trying to update the status or requiring the claim to be in the correct status / queue..
'''  - 10/9/2014 - KD: Added AddSecPagesCodeNoObjects() as an option to setting up using
'''         clsLetterInstance and clsLetterInstanceDct
'''  - 05/20/2014 - Created class
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



'' 8/15/2014: KD Need to make sure this is imported to the Master Claim Admin

Private Const ClassName As String = "mod_BOLD_LetterAutomation"
'Public Const csSAMPLEFILE_PATH As String = "\\ccaintranet.com\dfs-cms-ds\Data\CMS\AnalystFolders\KevinD\Letter_Automation\_sample_request.txt"
Public Const csSAMPLEFILE_PATH As String = "C:\_sample_request.txt"
Public Const cs_Sample_Message As String = "Please pull the letter addressed to:" & vbCrLf & vbCrLf & "Dr.John a.Doe" & vbCrLf & "1234 Some Street" & vbCrLf & "SomeTown, PA 19090" & vbCrLf & vbCrLf & "This is a: VRRL letter and should be the 23rd letter in the batch" & vbCrLf & "(starting at page number: 62)" & vbCrLf & vbCrLf & "[DYNAMICTEXT]" & vbCrLf & vbCrLf & "When you do please contact:" & vbCrLf & "Kevin Dearing at Extension 1145"

Public goSettings As clsSettings

Private Const cs_USER_TEMPLATE_PATH_ROOT As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_PROCESSING\"
Public Const cs_TEMP_RULE_TABLE_NAME As String = "tmp_BOLD_LETTER_Rules"

Public Const cs_MAIN_SERVER_NAME As String = "DS-FLD-009"
Public Const cs_MAIN_DB_NAME As String = "CMS_AUDITORS_CODE"  ' -- Should be the code db if there is one..


Public goBOLD_Processor As clsBOLD_Processor
Private clSelectedAccountId As Long

Private gstrCodeConnString As String
Private gstrDataConnString As String
Public cdctFoldersToCleanUp As Scripting.Dictionary


Public Enum MailLevels
    Batch = 1
    Instance = 2
    ProviderLevel = 3
    Claim = 4
End Enum


Public Property Get CurrentBatchId() As Long
Stop
'Dim oFrm As Form_frm_AutoRunProcessor
'Dim oProcessor As clsProcessor
'
'    If IsLoaded("frm_AutoRunProcessor") = False Then
'        CurrentBatchId = 0
'        GoTo Block_Exit
'    End If
'
'    Set oFrm = Forms("frm_AutoRunProcessor")
'    Set oProcessor = oFrm.ProcessorRef
'    If oProcessor Is Nothing Then
'        CurrentBatchId = 0
'    Else
'        CurrentBatchId = oProcessor.CurrentBatchId
'    End If
    

Block_Exit:
'    Set oFrm = Nothing
'    Set oProcessor = Nothing
End Property


Public Property Get CurrentJobId() As Long
Stop
'Dim oFrm As Form_frm_AutoRunProcessor
'Dim oProcessor As clsProcessor
'
'    If IsLoaded("frm_AutoRunProcessor") = False Then
'        CurrentJobId = 0
'        GoTo Block_Exit
'    End If
'
'    Set oFrm = Forms("frm_AutoRunProcessor")
'    Set oProcessor = oFrm.ProcessorRef
'    If oProcessor Is Nothing Then
'        CurrentJobId = 0
'    Else
'        CurrentJobId = oProcessor.CurrentJobId
'    End If
'
'Block_Exit:
'    Set oFrm = Nothing
'    Set oProcessor = Nothing
End Property
'Public goProcessor As clsProcessor
 

Public Property Get CodeConnString() As String
    If gstrCodeConnString = "" Then
        gstrCodeConnString = GetSetting("SPROC_CONN_STRING")
    End If
    CodeConnString = gstrCodeConnString
End Property
Public Property Let CodeConnString(strCodeConnString As String)
    gstrCodeConnString = strCodeConnString
End Property


Public Property Get DataConnString() As String
    If gstrDataConnString = "" Then
        gstrDataConnString = GetSetting("DATA_CONN_STRING")
    End If
    DataConnString = gstrDataConnString
End Property
Public Property Let DataConnString(strDataConnString As String)
    gstrDataConnString = strDataConnString
End Property

 
Public Property Get CurrentProcessor() As clsBOLD_Processor
    If goBOLD_Processor Is Nothing Then Set goBOLD_Processor = New clsBOLD_Processor
    Set CurrentProcessor = goBOLD_Processor
End Property


Public Function StartProcessor() As Boolean
    CurrentProcessor.StartProcessing
End Function


Public Function RunProcessor() As Boolean
    Set goBOLD_Processor = New clsBOLD_Processor
    
    goBOLD_Processor.StartProcessing

    LogMessage "RunProcessor", , "Got here"
    
End Function

Public Property Get GlobalSelectedAccountId() As Long
    If clSelectedAccountId = 0 Then clSelectedAccountId = 1
    
    GlobalSelectedAccountId = clSelectedAccountId
End Property
Public Property Let GlobalSelectedAccountId(lSelectedAccountId As Long)
    clSelectedAccountId = clSelectedAccountId
End Property


Public Property Get GetSetting(strSettingName As String) As Variant
Static dtLastUsed As Date

    If goSettings Is Nothing Then Set goSettings = New clsSettings
    
    ' when should we refresh the settings object?
    If DateDiff("n", dtLastUsed, Now()) >= 5 Then
        goSettings.Refresh
        dtLastUsed = Now()
    End If
    
    GetSetting = goSettings.GetSetting(strSettingName)
    
End Property



'Public Function ReleaseLetterTypes(sLtrTypeInClause As String) As Boolean
Public Function ReleaseLetterTypes(sLetterType As String, sLetterReqDt As String, sHeld As String, sManualOverRide As String, _
     sQueueStatusDt As String, sQueueDt As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
'Dim sSql As String


    strProcName = ClassName & ".ReleaseLetterTypes"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_ReleaseQueueLetters_Straight"
'        .sqlString = "usp_LETTER_Automation_ReleaseQueueLetters"
        
        .Parameters.Refresh
        .Parameters("@pAccountId") = GlobalSelectedAccountId
        .Parameters("@pLetterType") = sLetterType
        .Parameters("@pLetterReqDt") = sLetterReqDt
        .Parameters("@pHeld") = sHeld
        .Parameters("@pManualOverRide") = sManualOverRide
        .Parameters("@pStatusDt") = sQueueStatusDt
        .Parameters("@pQueueDt") = sQueueDt
'        .Parameters("@pIDList") = sLtrTypeInClause
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
    End With
    
    ReleaseLetterTypes = True


Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



''' Recordset should have at least LetterType and TemplateLoc for the ones to process
''' strtemplatepath is returned as the temp directory used
'''
''' The dictionary contains the clsLetterTemplate copy
''' Note that the dictionary template is returned populated with:
'''  Key = lettertype, value = clsLetterTemplate object
'''
'''
Public Function CopyTemplatesToTempWorkFldr(oLetterTypeRS As ADODB.RecordSet, strTemplatePath As String, Optional dctTemplatesDict As Scripting.Dictionary) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim objLetterInfo As clsLetterTemplate
Dim strLocalTemplate As String
Dim strSQL As String
Dim strChkFile As String
Dim strErrMsg As String
Dim iFolderChkLoop As Integer

    strProcName = ClassName & ".CopyTemplatesToTempWorkFldr"

    If dctTemplatesDict Is Nothing Then Set dctTemplatesDict = New Scripting.Dictionary

    ' create template directory
    iFolderChkLoop = 0

    strTemplatePath = QualifyFldrPath(GetUserTempDirectory()) & "LETTERTEMPLATE"
    
    AddFolderToCleanUp (strTemplatePath)
    
    If CreateFolders(strTemplatePath) = False Then
        LogMessage strProcName, "ERROR", "Could not create user temp folder!", strTemplatePath, True
        GoTo Block_Exit
    End If
            
    If FolderExist(strTemplatePath) = False Then
        strErrMsg = "ERROR: can not create folder " & strTemplatePath
        GoTo Block_Err
    End If

    ' copy templates to local directory. Skip if template already there
    Do While Not oLetterTypeRS.EOF
        With oLetterTypeRS
            strLocalTemplate = strTemplatePath & "\" & GetFileName(!TemplateLoc)

            If FileExists(oLetterTypeRS("TemplateLoc").Value) = True Then
                If CopyFile(oLetterTypeRS("TemplateLoc").Value, strLocalTemplate, False, strErrMsg) = False Then
                    Stop
                End If
            
            Else
                strErrMsg = "Error: source template " & oLetterTypeRS("TemplateLoc").Value & " not found"
            
            End If
                    
            Set objLetterInfo = New clsLetterTemplate
            objLetterInfo.LetterType = Trim(!LetterType)
            objLetterInfo.TemplateLoc = strLocalTemplate
            
            If dctTemplatesDict.Exists(oLetterTypeRS("LetterType").Value) = True Then
                Set dctTemplatesDict.Item(oLetterTypeRS("LetterType").Value) = objLetterInfo
            Else
                dctTemplatesDict.Add oLetterTypeRS("LetterType").Value, objLetterInfo
            End If

            .MoveNext
        End With
    Loop
    
    CopyTemplatesToTempWorkFldr = True
    
Block_Exit:

    Exit Function
Block_Err:
    If strErrMsg <> "" Then
        LogMessage strProcName, "ERROR", strErrMsg
    Else
        ReportError Err, strProcName
    End If
    
    GoTo Block_Exit
End Function


Private Sub DeleteTemplates()
'    Dim Person As New ClsIdentity
    Dim strTemplatePath As String
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oFile As Scripting.file

    ' delete template directory
    strTemplatePath = cs_USER_TEMPLATE_PATH_ROOT & GetUserName & "\LETTERTEMPLATE"
    
    'JS Change 20130305 no more delete the folder, now it will only delete the contents
    'DeleteFolder (strTemplatePath)
    
'    On Error Resume Next
    
    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder(strTemplatePath)
    For Each oFile In oFldr.Files
        Dim sFilePath As String
        sFilePath = oFile.Path
        Set oFile = Nothing
        If DeleteFile(sFilePath, False) = False Then
            LogMessage ClassName & ".DeleteTemplates", , "Tried to delete a file - not there, probably fine!", strTemplatePath
        End If
    Next
   

    Set oFile = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing

End Sub


Public Function HoldLetterTypes(sLtrTypeWhereClause As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sSql As String
'Stop

    strProcName = ClassName & ".HoldLetterTypes"
    sSql = "UPDATE Q SET Held = 1, ReleaseDateTime = NULL, ReleaseUser = '" & GetUserName() & "', Status = 'Q', StatusDt = GetDate() " & _
        " FROM LETTER_Print_Queue Q INNER Join ADMin.Admin_Account_Config a ON Q.AccountId = A.AccountID " & _
        " WHERE " & sLtrTypeWhereClause & " AND Error = 0 AND Held = 0"

    
    Set oAdo = New clsADO
    With oAdo

        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = sSql
        .Execute
        
    End With
    
    HoldLetterTypes = True
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Public Function MultipleValuesToInClauseSQL(sFieldName As String, vArray() As String, Optional bNumbers As Boolean = False) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sRet As String
Dim sDelim As String

    strProcName = ClassName & ".MultipleValuesToInClauseSQL"
    
    
    If bNumbers = False Then
        sDelim = "','"
    Else
        sDelim = ","
    End If
    
    sRet = Join(vArray, sDelim)
    If bNumbers = True Then
        sRet = " IN (" & sRet & ") "
    Else
        sRet = " IN ('" & sRet & "') "
    End If
    
    MultipleValuesToInClauseSQL = " " & sFieldName & sRet
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




Public Function MultipleColumnsAndRowsToXmlForJoin(sFieldName As String, vValuesArray() As String, Optional sDelimiter As String = ":", _
            Optional sRootName As String = "list", Optional sRowIdentifier As String = "row") As String
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sRet As String
'Dim sDelim As String
'Dim sXmlStart As String
'Dim sXmlEnd As String
'Dim sAryFields() As String
'Dim iValIdx As Integer
'
'Stop
'    strProcName = ClassName & ".MultipleColumnsAndRowsToXmlForJoin"
'    ' Ultimately we want to turn up with
'    ' something like this:
'    ' <sFieldName value="val1" />
'
'    sAryFields = Split(sFieldName, sDelimiter)
'
'    For iValIdx = 0 To UBound(vValuesArray)
'
'    Next
'
'
'
'
'    sXmlStart = "<" & LCase(sRootName) & " value="""
'    sXmlEnd = """ />"
'
'    sDelim = sXmlEnd & sXmlStart
'
'    sRet = Join(vArray, sDelim)
'
''    sRet = MakeXmlString(sRet, "list", sDelim)
'
'    sRet = sXmlStart & sRet & sXmlEnd
'
'    MultipleColumnsAndRowsToXmlForJoin = "<list>" & sRet & "</list>"
'
'
'Block_Exit:
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
End Function


Public Function te()
Dim vary(3) As String

    vary(0) = "123"
    vary(1) = "124"
    vary(2) = "125"
    vary(3) = "127"
    
    
    Debug.Print MultipleColumnsToXmlForJoin("Batchid", vary)
    
End Function


Public Function MultipleColumnsToXmlForJoin(sFieldName As String, vArray() As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sRet As String
Dim sDelim As String
Dim sXmlStart As String
Dim sXmlEnd As String
Stop
    strProcName = ClassName & ".MultipleColumnsToXmlForJoin"
    ' Ultimately we want to turn up with
    ' something like this:
    ' <sFieldName value="val1" />
    
    sXmlStart = "<" & LCase(sFieldName) & " value="""
    sXmlEnd = """ />"
    
    sDelim = sXmlEnd & sXmlStart
    
    sRet = Join(vArray, sDelim)
    
    
    sRet = sXmlStart & sRet & sXmlEnd
    
    MultipleColumnsToXmlForJoin = "<list>" & sRet & "</list>"
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function MultipleRowsAndColumnsToXmlForJoin(sRootName As String, vXmlRowArray() As String, Optional bForceToLowerCase As Boolean = True) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sWholeData As String

    strProcName = ClassName & ".MultipleRowsAndColumnsToXmlForJoin"
    ' Ultimately we want to turn up with
    ' something like this:
    ' <sFieldName value="val1" />
    
    sWholeData = Join(vXmlRowArray, vbCrLf)
    If bForceToLowerCase = True Then
        sWholeData = "<" & LCase(sRootName) & ">" & vbCrLf & sWholeData & "</" & LCase(sRootName) & ">"
    
    Else
    sWholeData = "<" & sRootName & ">" & vbCrLf & sWholeData & "</" & sRootName & ">"
    End If
    
    MultipleRowsAndColumnsToXmlForJoin = sWholeData
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

'Public Function MultipleColumnsToXmlForJoin(sFieldName As String, vArray() As String) As String
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sRet As String
'Dim sDelim As String
'Dim sXmlStart As String
'Dim sXmlEnd As String
'Stop
'    strProcName = ClassName & ".MultipleColumnsToXmlForJoin"
'    ' Ultimately we want to turn up with
'    ' something like this:
'    ' <sFieldName value="val1" />
'
'    sXmlStart = "<" & LCase(sFieldName) & " value="""
'    sXmlEnd = """ />"
'
'    sDelim = sXmlEnd & sXmlStart
'
'    sRet = Join(vArray, sDelim)
'
'    sRet = MakeXmlString(sRet, "list", sDelim)
'
'    sRet = sXmlStart & sRet & sXmlEnd
'
'    MultipleColumnsToXmlForJoin = "<list>" & sRet & "</list>"
'
'
'Block_Exit:
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function



Public Function MultipleValuesToXml(sFieldName As String, vArray() As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sRet As String
Dim sDelim As String
Dim sXmlStart As String
Dim sXmlEnd As String

    strProcName = ClassName & ".MultipleValuesToInClauseSQL"
    ' Ultimately we want to turn up with
    ' something like this:
    ' <sFieldName value="val1" />
    
    sXmlStart = "<" & LCase(sFieldName) & " value="""
    sXmlEnd = """ />"
    
    sDelim = sXmlEnd & sXmlStart
    
    sRet = Join(vArray, sDelim)
    
    
    sRet = sXmlStart & sRet & sXmlEnd
    
    MultipleValuesToXml = "<list>" & sRet & "</list>"
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function GetTotalLettersForMaxQueueID(lSelectedId As Long) As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim lRet As Long
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetTotalLettersForMaxQueueId"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString

        .SQLTextType = StoredProc
        
        .sqlString = "usp_LETTER_Automation_GetTtlLettersForMaxQID"
        .Parameters.Refresh
        .Parameters("@pAccountId") = lSelectedId
        Set oRs = .ExecuteRS
    End With
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Private Function IsCommandBar(sNameToCheck As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCmdBar As CommandBar

    strProcName = ClassName & ".IsCommandBar"
    
    For Each oCmdBar In CommandBars
        If LCase(sNameToCheck) = LCase(oCmdBar.Name) Then
            IsCommandBar = True
            Exit For
        End If
    Next
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function FixAddress() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFrm As Form_frm_BOLD_Ops_Dashboard
Dim iPages As Integer
Dim oLV As Object
Dim oLI As ListItem
Dim sInClause As String

Dim sMsg As String


    strProcName = ClassName & ".FixAddress"
    
    
    sMsg = "The plan is for this feature to fix the address in the QUEUE tables, but will send the address to a confurable stored procedure in order to fix the address in the main system"
    
    MsgBox sMsg, vbInformation, "Please note:"
    
    
'    Stop
'
'    If IsOpen("frm_BOLD_Dashboard", acForm) = False Then
'            Stop
'    End If
'
'    Set oFrm = Application.Forms("frm_BOLD_Dashboard")
'    ' first, which page is selected
'    Debug.Print oFrm.tabDisplay.Value
'
'    Select Case oFrm.tabDisplay.Value
'    Case 0  ' queue tab
'        Set oLV = oFrm.lvQueue
'    Case 1  ' Generate page
'        Set oLV = oFrm.lvQueue
'    Case 2  ' output page
'        Set oLV = oFrm.lvQueue
'    End Select
'
'    Set oLI = oFrm.RightClickedListItem
'
'
'    Debug.Print oLI.Text
'
'    sInClause = " LetterType IN ('" & oLI.Text & "') "
'    Call HoldLetterTypes(sInClause)
'
'    Call oFrm.RefreshData

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

' Note: This requires a reference to Microsoft Office Object Library
Public Sub SetUpContextMenu(Optional ByVal sSubMenuName As String = "Generate")
On Error GoTo Block_Err
Dim strProcName As String
Dim combo As CommandBarControl

    strProcName = ClassName & ".SetUpContextMenu"
    
    sSubMenuName = UCase(sSubMenuName)
    
    
    If IsCommandBar("BOLD_RightClickmnu") = True Then
        CommandBars("BOLD_RightClickmnu").Delete
    End If


    ' Make this menu a popup menu
    With CommandBars.Add(Name:="BOLD_RightClickmnu", position:=msoBarPopup)
        Select Case sSubMenuName
'        Case "GENERATE", "QUEUE"
            ' Provide the user the ability to input text using the msoControlEdit type
        Case "QUEUE"
            ' Provide the user the ability to input text using the msoControlEdit type
            Set combo = .Controls.Add(Type:=msoControlEdit)
            combo.Caption = "Request Mail Room to pull because:" ' Add a label the user will see
            combo.OnAction = "=MailRoomPullBatch()" ' Add the name of a function to call
'            combo.Parameter = combo.Value
            
            
            ' Provide the user the ability to click a menu option to execute a function
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Show Claims" ' Add a label the user will see
            combo.OnAction = "=ShowAssociatedClaims()" ' Add the name of a function to call
            
            combo.BeginGroup = True ' Add a line to separate above group
            combo.Caption = "Reprint" ' Add label the user will see
            combo.OnAction = "=Reprint()" ' Add the name of a function to call
            
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Manual Override"
            combo.OnAction = "=ManualOverride()" ' Add the name of a function to call
            
            
        Case "OUTPUT"
            ' Provide the user the ability to input text using the msoControlEdit type
            Set combo = .Controls.Add(Type:=msoControlEdit)
            combo.Caption = "Request Mail Room to pull because:" ' Add a label the user will see
            combo.OnAction = "=MailRoomPullBatch()" ' Add the name of a function to call
'            combo.Parameter = combo.Value
            
            
            ' Provide the user the ability to click a menu option to execute a function
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Open Claim" ' Add a label the user will see
            combo.OnAction = "=OpenErrorClaim()" ' Add the name of a function to call
            
            combo.BeginGroup = True ' Add a line to separate above group
            combo.Caption = "Reprint" ' Add label the user will see
            combo.OnAction = "=Reprint()" ' Add the name of a function to call

            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Manual Override"
            combo.OnAction = "=ManualOverride()" ' Add the name of a function to call
            
            
            
        Case "OUTPUTERRORS"
            Set combo = .Controls.Add(Type:=msoControlEdit)
            combo.Caption = "Request Mail Room to pull because:" ' Add a label the user will see
            combo.OnAction = "=MailRoomPullBatch()" ' Add the name of a function to call
'            combo.Parameter = combo.Value
            
            
            ' Provide the user the ability to click a menu option to execute a function
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.BeginGroup = True ' Add a line to separate above group
            combo.Caption = "Re-Combine this batch and send to Mail room" ' Add label the user will see
            combo.OnAction = "=Reprint()" ' Add the name of a function to call

            ' Provide the user the ability to click a menu option to execute a function
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.BeginGroup = True ' Add a line to separate above group
            combo.Caption = "Not an error" ' Add label the user will see
            combo.OnAction = "=NotAnOutputError()" ' Add the name of a function to call
            
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Manual Override"
            combo.OnAction = "=ManualOverride()" ' Add the name of a function to call
            
            

        Case "GENERATEERRORS"
        
            ' Provide the user the ability to input text using the msoControlEdit type
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Open Claim" ' Add a label the user will see
            combo.OnAction = "=OpenErrorClaim()" ' Add the name of a function to call
            
            ' Provide the user the ability to click a menu option to execute a function
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.BeginGroup = True ' Add a line to separate above group
            combo.Caption = "Re-Generate this batch"  ' Add label the user will see
            combo.OnAction = "=Reprint()" ' Add the name of a function to call

            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Manual Override"
            combo.OnAction = "=ManualOverride()" ' Add the name of a function to call
            
            
        Case "QUEUEERRORS"
            ' Provide the user the ability to input text using the msoControlEdit type
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Open Claim" ' Add a label the user will see
            combo.OnAction = "=OpenErrorClaim()" ' Add the name of a function to call

            ' Provide the user the ability to input text using the msoControlEdit type
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Fix Address" ' Add a label the user will see
            combo.OnAction = "=FixAddress()" ' Add the name of a function to call
            
            ' Provide the user the ability to click a menu option to execute a function
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.BeginGroup = True ' Add a line to separate above group
            combo.Caption = "Change the letter's date (Reprint)" ' Add label the user will see
            combo.OnAction = "=Reprint()" ' Add the name of a function to call

            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Manual Override"
            combo.OnAction = "=ManualOverride()" ' Add the name of a function to call
            
        Case "MAILQUEUE", "MAIL QUEUE"
            ' Provide the user the ability to input text using the msoControlEdit type
            Set combo = .Controls.Add(Type:=msoControlEdit)
            combo.Caption = "Request Mail Room to pull because:" ' Add a label the user will see
            combo.OnAction = "=MailRoomPullBatch()" ' Add the name of a function to call
'            combo.Parameter = combo.Value
            
            
            ' Provide the user the ability to click a menu option to execute a function
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Show Claims" ' Add a label the user will see
            combo.OnAction = "=ShowAssociatedClaims()" ' Add the name of a function to call
            
            combo.BeginGroup = True ' Add a line to separate above group
            combo.Caption = "Reprint" ' Add label the user will see
            combo.OnAction = "=Reprint()" ' Add the name of a function to call
            
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Manual Override"
            combo.OnAction = "=ManualOverride()" ' Add the name of a function to call
            
            
        Case Else
            ' Provide the user the ability to input text using the msoControlEdit type
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Place back on Hold" ' Add a label the user will see
            combo.OnAction = "=PlaceItemsOnHold()" ' Add the name of a function to call
            
            
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.Caption = "Manual Override"
            combo.OnAction = "=ManualOverride()" ' Add the name of a function to call
            
            
            ' Provide the user the ability to click a menu option to execute a function
            Set combo = .Controls.Add(Type:=msoControlButton)
            combo.BeginGroup = True ' Add a line to separate above group
            combo.Caption = "Reprint" ' Add label the user will see
            combo.OnAction = "=Reprint()" ' Add the name of a function to call

        End Select
        
        
'        ' Provide the user the ability to click a menu option to execute a function
'        Set combo = .Controls.Add(Type:=msoControlButton)
'        combo.Caption = "Delete Record" ' Add a label the user will see
'
'        combo.OnAction = "=DeleteRecordFunction()" ' Add the name of the function to call"
'        combo.SetFocus
'        combo.Visible = True
    End With

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Function ManualOverride() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFrm As Form_frm_BOLD_Ops_Dashboard
Dim iPages As Integer
'Dim oLV As ListView
Dim oLV As Object
Dim oLI As ListItem
Dim sInClause As String
Dim sLVName As String

    strProcName = ClassName & ".ManualOverride"
    
   
    If IsOpen("frm_BOLD_Ops_Dashboard", acForm) = False Then
            Stop
    End If
    
    Set oFrm = Application.Forms("frm_BOLD_Ops_Dashboard")
    ' first, which page is selected
    Debug.Print oFrm.tabDisplay.Value
    
    Select Case oFrm.tabDisplay.Value
    Case 0  ' queue tab
        Set oLV = oFrm.lvQueue
        sLVName = "Queue"
    Case 1  ' Generate page
        Set oLV = oFrm.lvQueue
        sLVName = "Generate"
    Case 2  ' output page
        Set oLV = oFrm.lvQueue
        sLVName = "Output"
    End Select
    
    Set oLI = oFrm.RightClickedListItem
    
    
    Debug.Print oFrm.QueueColumns.GetDetails(sLVName, "LetterType")
    
    Debug.Print oFrm.QueueColumns.GetDetails(sLVName, "Held")
    
Stop
    
    Debug.Print oLI.Text
    Debug.Print oLI.ListSubItems(1)
    
    sInClause = " LetterType = '" & oLI.Text & "' AND LetterReqDt = '" & oLI.SubItems(oFrm.QueueColumns.GetDetails(sLVName, "LetterReqDt")) & "' AND Held = '" & oLI.SubItems(oFrm.QueueColumns.GetDetails(sLVName, "Held")) & "'"
    Call SetManualOverride(sInClause)
    
    Call oFrm.RefreshData
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function SetManualOverride(sInClause As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".SetManualOverride"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_SetManualOverride"
        .Parameters.Refresh
        .Parameters("@InClause") = sInClause
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
    End With
    
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function MailRoomPullBatch(Optional sValue As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Dim oTxt As Scripting.TextStream
Dim sOutFile As String
Dim sMsg As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sAlertFolderPath As String

    Dim oFrm As Form_frm_BOLD_Ops_Dashboard
'Dim oPage As Page
Dim iPages As Integer
'Dim oLV As ListView
Dim oLV As Object
'Dim oLI As ListItem
Dim oLI As Object
Dim sInClause As String
Dim iaryFields() As Integer

    strProcName = ClassName & ".MailRoomPullBatch"

    If IsOpen("frm_BOLD_Ops_Dashboard", acForm) = False Then
            Stop
    End If
    
    Set oFrm = Application.Forms("frm_BOLD_Ops_Dashboard")
    ' first, which page is selected
    Debug.Print oFrm.tabDisplay.Value
    
    
    Select Case oFrm.tabDisplay.Value
    Case 0  ' queue tab
        Set oLV = oFrm.lvQueue
        ReDim iaryFields(3)
        iaryFields(0) = 0   ' 0 is the text value (letter type)
        iaryFields(1) = 2   ' AccountName
        iaryFields(2) = 3   ' LetterReqDt
        iaryFields(3) = 11   ' ManualOverride
        
    Case 1  ' Generate page
        Set oLV = oFrm.lvGenerate
        ReDim iaryFields(3)
        iaryFields(0) = 0  ' 0 is the text value (LetterType)
        iaryFields(1) = 1  ' accountname
        iaryFields(2) = 2  ' LetterDt
        iaryFields(3) = 3  ' BatchId
    Case 2  ' output page
        Set oLV = oFrm.lvOutput
'        ReDim iaryFields(3)
'        iaryFields(0) = 0  ' 0 is the text value (LetterType)
'        iaryFields(1) = 2  ' accountname
'        iaryFields(2) = 3  ' LetterDt
'        iaryFields(3) = 4  ' BatchId
    End Select
    
    Set oLI = oFrm.RightClickedListItem
    


    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_GetMailroomPullDetails"
        .Parameters.Refresh

        .Parameters("@pAccountid") = gintAccountID
        .Parameters("@pInstanceId") = oLI.SubItems(1)
'        .Parameters("@pCnlyClaimNum") = sValue
        Set oRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg").Value, oLI.SubItems(1)
            GoTo Block_Exit
        End If
    End With
    
    sAlertFolderPath = GetSetting("ALERT_FOLDER_PATH")
    
    sAlertFolderPath = QualifyFldrPath("") & Format(Now, "yyyymmdd") & "\"
    
    
    sValue = CommandBars("BOLD_RightClickmnu").Controls("Request Mail Room to pull because:").Text
    sOutFile = sAlertFolderPath & Format(Now(), "yyyymmddhhnn") & ".txt"
    

    Set oFso = New Scripting.FileSystemObject
    Set oTxt = oFso.CreateTextFile(csSAMPLEFILE_PATH, True, True)
    oTxt.Write Replace(cs_Sample_Message, "[DYNAMICTEXT]", sValue)
    oTxt.Close
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Set oTxt = Nothing
    Set oFso = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Err
End Function

Public Function PlaceItemsOnHold() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFrm As Form_frm_BOLD_Ops_Dashboard
'Dim oPage As Page
Dim iPages As Integer
'Dim oLV As ListView
Dim oLV As Object
Dim oLI As ListItem
Dim sInClause As String
Dim iaryFields() As Integer

    strProcName = ClassName & ".PlaceItemsOnHold"
'    Stop
    
    If IsOpen("frm_BOLD_Ops_Dashboard", acForm) = False Then
            Stop
    End If
    
    Set oFrm = Application.Forms("frm_BOLD_Ops_Dashboard")
    ' first, which page is selected
    Debug.Print oFrm.tabDisplay.Value
    
    
    Select Case oFrm.tabDisplay.Value
    Case 0  ' queue tab
        Set oLV = oFrm.lvQueue
        ReDim iaryFields(3)
        iaryFields(0) = 0   ' 0 is the text value (letter type)
        iaryFields(1) = 2   ' AccountName
        iaryFields(2) = 3   ' LetterReqDt
        iaryFields(3) = 11   ' ManualOverride
        
    Case 1  ' Generate page
        Set oLV = oFrm.lvGenerate
        ReDim iaryFields(3)
        iaryFields(0) = 0  ' 0 is the text value (LetterType)
        iaryFields(1) = 1  ' accountname
        iaryFields(2) = 2  ' LetterDt
        iaryFields(3) = 3  ' BatchId
    Case 2  ' output page
        Set oLV = oFrm.lvOutput
'        ReDim iaryFields(3)
'        iaryFields(0) = 0  ' 0 is the text value (LetterType)
'        iaryFields(1) = 2  ' accountname
'        iaryFields(2) = 3  ' LetterDt
'        iaryFields(3) = 4  ' BatchId
    End Select
    
    Set oLI = oFrm.RightClickedListItem
    
    '' KD NOTE TO SELF: I need to have this a bit generic - I need to pass multiplevalues
    '' since the LI.Text is not unique on many of the List views..
    '' Queue:
        '' LetterType
        '' AccountName
        '' LetterReqDt
        '' ManualOverride
    '' Generate:
        '' LetterType
        '' AccountName
        '' LetterDate
        '' BatchId if there is one
    '' Output:
        '' BatchId
        '' BatchType
    
    '' So, my query should be something like:
    
    '' update PQ SET Held = 1 WHERE
    ''  LetterType = '' AND AccountNmae = '' And LetterDate = '' and Batchid = ''
    '' so maybe I can use a dictionary..
    '' key = fieldname
    '' value = required value
Dim iIdx As Integer
Dim sName As String
Dim sVal As String
Dim iSubItemId As Integer


    For iIdx = 0 To UBound(iaryFields)
        iSubItemId = iaryFields(iIdx)
        sName = oLV.ColumnHeaders.Item(iSubItemId + 1)

        If iSubItemId = 0 Then
            sVal = oLI.Text
        Else
            
            sVal = oLI.SubItems(iSubItemId)
        End If
        If sVal = "" Then
            sName = "ISNULL(" & sName & ", '')"
            
            
        End If
'        Stop
        sInClause = sInClause & sName & " = '" & sVal & "' AND "
    Next
    sInClause = left(sInClause, Len(sInClause) - 4)
Debug.Print sInClause
'Stop
    
'    Debug.Print oLI.Text
    
    'sInClause = " LetterType IN ('" & oLI.Text & "') "
    Call HoldLetterTypes(sInClause)
    
    Call oFrm.RefreshData

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function Reprint() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFrm As Form_frm_BOLD_Ops_Dashboard
Dim oMFrm As Form_frm_BOLD_Mail_Dashboard
Dim iPages As Integer
'Dim oLV As ListView
Dim oLV As Object
Dim oLI As ListItem
Dim sInClause As String
Dim iaryFields() As Integer
Dim oColPos As clsLVColumnPositions

    strProcName = ClassName & ".Reprint"

    
    If IsOpen("frm_BOLD_Ops_Dashboard", acForm) = False Then
        If IsOpen("frm_BOLD_Mail_Dashboard", acForm) = True Then
            Set oMFrm = Application.Forms("frm_BOLD_Mail_Dashboard")
            oMFrm.RefreshData
            Set oLV = oMFrm.lvQErrorDetails
            Set oColPos = oMFrm.QueueColumns
            
        Else
            Stop
        End If
        Set oFrm = New Form_frm_BOLD_Ops_Dashboard
        oFrm.RefreshData
        
        
        
            Stop
    End If
    
    Set oFrm = Application.Forms("frm_BOLD_Ops_Dashboard")
    ' first, which page is selected
    Debug.Print oFrm.tabDisplay.Value
    
    Set oColPos = oFrm.QueueColumns
    
    Select Case oFrm.tabDisplay.Value
    Case 2  ' queue tab
        Set oLV = oFrm.lvQueue
        
        ReDim iaryFields(3)
        iaryFields(0) = oColPos.GetDetails("QUEUE", "LETTERTYPE")
'        iaryFields(0) = 0   ' 0 is the text value (letter type)
        
        iaryFields(1) = oColPos.GetDetails("QUEUE", "ClientName")
'        iaryFields(1) = 2   ' AccountName
        
        
        
        iaryFields(1) = oColPos.GetDetails("QUEUE", "LetterReqDt")
'        iaryFields(2) = 3   ' LetterReqDt
        
        
        iaryFields(1) = oColPos.GetDetails("QUEUE", "ManualOverride")
'        iaryFields(3) = 11   ' ManualOverride
        
        
    Case 3 ' Generate page
        Set oLV = oFrm.lvGenerate
        
        ReDim iaryFields(3)
        iaryFields(0) = oColPos.GetDetails("GENERATE", "LetterType")
        iaryFields(0) = 0  ' 0 is the text value (LetterType)
        
        iaryFields(1) = oColPos.GetDetails("GENERATE", "ClientName")
        iaryFields(1) = 1  ' accountname
        
        iaryFields(2) = oColPos.GetDetails("GENERATE", "LetterReqDt")
        iaryFields(2) = 2  ' LetterReqDt
        
        iaryFields(3) = oColPos.GetDetails("GENERATE", "BatchId")
        iaryFields(3) = 3  ' BatchId
    Case 4  ' output page
Stop
        Set oLV = oFrm.lvOutput
        
        Set oColPos = New clsLVColumnPositions
           
        Set oColPos = oFrm.QueueColumns
        
'        Call oColPos.SetDetails("Output", oLV)
    
        ReDim iaryFields(4)
        iaryFields(0) = oColPos.GetDetails("OUTPUT", "LetterType")
'        iaryFields(0) = 0  ' 0 is the text value (LetterType)
        
        iaryFields(1) = oColPos.GetDetails("OUTPUT", "ClientName")
'        iaryFields(1) = 1  ' accountname
        
        iaryFields(2) = oColPos.GetDetails("OUTPUT", "LetterReqDt")
'        iaryFields(2) = 2  ' LetterReqDt
        
        iaryFields(3) = oColPos.GetDetails("OUTPUT", "BatchId")
        iaryFields(4) = oColPos.GetDetails("OUTPUT", "InstanceID")
Stop
'        iaryFields(3) = 3  ' BatchId
    Case 0  ' Not loaded Errors
    
    
    Case 1  ' Error Queue
    
    End Select
    
    Set oLI = oFrm.RightClickedListItem
    
  '' update PQ SET Held = 1 WHERE
    ''  LetterType = '' AND AccountNmae = '' And LetterDate = '' and Batchid = ''
    '' so maybe I can use a dictionary..
    '' key = fieldname
    '' value = required value
Dim iIdx As Integer
Dim sName As String
Dim sVal As String
Dim iSubItemId As Integer
Dim sXml As String
Dim sAttributes As String
Dim sInstanceId As String
Dim sCurDt As String


    For iIdx = 0 To UBound(iaryFields)
        iSubItemId = iaryFields(iIdx)
        sName = oLV.ColumnHeaders.Item(iSubItemId + 1)

        If iSubItemId = 0 Then
            sVal = oLI.Text
        Else
            
            sVal = oLI.SubItems(iSubItemId)
        End If
        If sVal = "" Then
            sName = "ISNULL(" & sName & ", '')"
            
            
        End If
'        Stop


        If UCase(sName) = "LETTERREQDT" Then
            sCurDt = Format(sVal, "m/d/yyyy")
        End If
        

        If UCase(sName) = "INSTANCEID" Then
            sInstanceId = sVal
        Else
            sAttributes = sAttributes & " " & sName & "=""" & sVal & """"
        End If

        sInClause = sInClause & sName & " = '" & sVal & "' AND "
        
        
'        sInstanceId = ""
    Next
    
    sXml = sXml & "<list><instance instanceid=""" & sInstanceId & """ " & sAttributes & " /></list>"
    
    
    sInClause = left(sInClause, Len(sInClause) - 4)
Debug.Print sInClause
Stop
    
'    Debug.Print oLI.Text
    
    'sInClause = " LetterType IN ('" & oLI.Text & "') "
    Call ReprintMatchingInstances(sXml, sCurDt)
    
    Call oFrm.RefreshData

    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function ReprintMatchingInstances(sConstraintsXml As String, sCurDt As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim dtNewDate As Date
Dim sDate As String

    strProcName = ClassName & ".ReprintMatchingInstances"
        
    sDate = InputBox("What date should we use", "New Date?", sCurDt)
    
    If sDate = "" Then
        ' canceled
        Stop
        GoTo Block_Exit
    End If
    
    If IsDate(sDate) = True Then
        dtNewDate = CDate(sDate)
    Else
        Stop
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_Reprint_Letters"
        .Parameters.Refresh
        .Parameters("@pConstraintsIn") = sConstraintsXml
        .Parameters("@pLtrDate") = dtNewDate
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was a problem executing " & .sqlString, .Parameters("@pErrMsg").Value, True
            Stop
        End If
    End With
    
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function GetAccountsNConfigs(Optional bReprint As Boolean) As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO

    strProcName = ClassName & ".GetAccountsNConfigs"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_GetAccountConfigsWithLetters"
        .Parameters.Refresh
        
        .Parameters("@pReprint") = IIf(bReprint, 1, 0)
        
        Set oRs = .ExecuteRS
        If .GotData = False Or Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "Problem obtaining accounts and configurations!", Nz(.Parameters("@pErrMsg").Value, "Nothing to print")
            Stop
        End If
    End With
    Set GetAccountsNConfigs = oRs
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




''' Recordset is
''' dictionary is
''' temp folder
'''
'''
'''
Public Function PerformIndividualMailMerges(oWordApp As Word.Application, oRs As ADODB.RecordSet, dtcTemplates As Scripting.Dictionary, Optional sTempFolder As String, _
        Optional lRowsAffected As Long = -1, Optional bSampleOnly As Boolean = False, Optional sPreviewOutPath As String, Optional bReprint As Boolean = False) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim oLtrTmplt As clsLetterTemplate
Dim oTemplateDoc As Word.Document
Dim oMergeDoc As Word.Document
Dim sTemplateFldr As String
Dim sLocalTemplatePath As String
Dim strODCFile As String
Dim strOutputFldr As String
Dim sOutFullPath As String
Dim strDocumentSaveBasePath As String
Dim oLetterInst As clsLetterInstance
Dim oLetters As clsLetterInstanceDct
Dim dFileSize As Double
Dim dSleepTime As Double
Dim sErrMsg  As String
Dim sLastTemplate As String
Dim lTtl As Long
Dim lCurrent As Long
Dim lTtlLtrsWithoutErrors As Long
Dim lPctDone As Long
Dim oBAdo As clsADO
Dim oAcctsRS As ADODB.RecordSet
Dim lThisAcct As Long
Dim sMailMergeSproc As String

    '' This function really needs to be broken down into a couple separate functions..
    '' KD COMEBACK: it also belongs in the clsProcessor but am keeping it here for now until we kill the legacy process
    strProcName = ClassName & ".PerformIndividualMailMerges"
    
    ' Ok, so, get the list of accounts so we can loop over them
    ' we'll do that by querying the queue, not the acount table
    ' because there may not be any letters to process for a given account
    Set oAcctsRS = GetAccountsNConfigs(bReprint)
    If oAcctsRS Is Nothing Then
        LogMessage strProcName, "ERROR", "Nothing to do?"
        Stop
    End If
    lTtl = oRs.recordCount
    
    While Not oAcctsRS.EOF
        lThisAcct = oAcctsRS("AccountId").Value
        strODCFile = oAcctsRS("ODC_ConnectionFile").Value
        strDocumentSaveBasePath = oAcctsRS("LetterOutputLocation").Value
        
        If sTempFolder = "" Then
            sTempFolder = GetUserTempDirectory()
        End If

        sTemplateFldr = sTempFolder
    '    Stop
        ' kev, need to get
        strDocumentSaveBasePath = GetSetting("LetterOutputLocation")
'        strODCFile = GetSetting("ODC_ConnectionFile")
        ' Let's filter the recordset by the accountid
        
        If bSampleOnly = True Then
            strDocumentSaveBasePath = sTempFolder
            sMailMergeSproc = "usp_LETTER_Automation_MailMergeSource_ManualOverrides"
        Else
'            If bReprint = True Then
'                sMailMergeSproc = "usp_LETTER_Automation_MailMergeSourceForReprint"
'            Else
            sMailMergeSproc = "usp_LETTER_Automation_MailMergeSource"
'            End If
        End If
        
        oRs.filter = "AccountId = " & CStr(lThisAcct)
        If oRs.BOF And oRs.EOF Then
            GoTo NxtAcct
        End If
        oRs.MoveFirst
        
'        lTtl = oRs.RecordCount
        If oRs.recordCount = 0 Then GoTo NxtAcct
        
        ' The plan here is to:
        ' For each letter (instanceid)
        '' RS has: InstanceId, BatchId, LetterType, AccountId, Status
        While Not oRs.EOF
            ' - check that all cnlyclaimnum's are in good status still
            '    - if not make it an error and mark the queue accordingly
            lCurrent = lCurrent + 1
            
            lPctDone = CDbl((lCurrent / lTtl) * 100)
            
            Set oLetterInst = New clsLetterInstance
            With oLetterInst
                .InstanceId = oRs("InstanceId").Value
                .ProvNum = oRs("CnlyProvId").Value
                .LetterBatchId = oRs("BatchId").Value
                .LetterType = UCase(oRs("LetterType").Value)
                .LetterCreateDt = oRs("LetterReqDt").Value
                .LetterQueueStatus = oRs("Status").Value
                .AccountID = lThisAcct
            End With
            
            If bSampleOnly = False Then
                Call UpdateProcessorStatus("Mail Merge", goBOLD_Processor.ThisQueueRunId, lPctDone, oRs("BatchId").Value, oRs("InstanceId").Value, False)
            End If
            
            If dtcTemplates.Exists(oLetterInst.LetterType) = False Then
                Stop
            End If
            Set oLtrTmplt = dtcTemplates.Item(oLetterInst.LetterType)
            
            sLocalTemplatePath = oLtrTmplt.TemplateLoc

'' KD: ONly for real deal - not for samples
'             If bSampleOnly = False Then
            If bReprint = False Then
                If InsureAllStatusesAreGood(oRs("InstanceId").Value, lThisAcct) = False Then
                    LogMessage strProcName, "ERROR", "Some claims status' were not appropriate for this letter type. These have been marked as status 'E'", oRs("InstanceId").Value
                    Stop
                End If
            End If
            
            If oWordApp Is Nothing Then
                Set oWordApp = New Word.Application
                oWordApp.visible = False
            End If
    
            If sLastTemplate <> sLocalTemplatePath Then
                ' close the last one if we have it open:
                If oWordApp.Documents.Count > 0 Then
                    For Each oTemplateDoc In oWordApp.Documents
                        oTemplateDoc.Close False
                    Next
                End If
                Set oTemplateDoc = oWordApp.Documents.Add(sLocalTemplatePath, , False)
            Else
                If oWordApp.Documents.Count > 0 Then
                    Set oTemplateDoc = oWordApp.ActiveDocument
                Else
                    Set oTemplateDoc = oWordApp.Documents.Add(sLocalTemplatePath, , False)
                End If
                sLastTemplate = sLocalTemplatePath
            End If
            
            ' - get the mailmerge recordset (well, details we need for the sproc)
            ' Set data source for mail merge.  Data will be from new Temp Table
            oTemplateDoc.MailMerge.OpenDataSource Name:=strODCFile, SqlStatement:="exec " & sMailMergeSproc & " '" & oLetterInst.InstanceId & "'"
            
            ' - open the word template from the dictionary.
            ' - Do the mail merge
            oTemplateDoc.MailMerge.MainDocumentType = 3 'wdDirectory
            oTemplateDoc.MailMerge.Destination = 0 'wdSendToNewDocument
            oTemplateDoc.MailMerge.Execute Pause:=False
            If left(oWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
                oWordApp.visible = True
                'MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
    '            Call ErrorCallStack_Add(clBatchId, "Error encountered with mail merge.", strProcName, "Instance ID: " & strInstanceId, , , strInstanceId, oLtrTmplt.LetterType)
Stop

                oWordApp.ActiveDocument.Activate
                GoTo Block_Exit
            End If
            
            '' Ok, now deal with Barcode issues:
            Call AddSecPagesCode(oWordApp.ActiveDocument, oLetterInst)
            
            ' - save it where it needs to be
            Set oMergeDoc = oWordApp.Documents(oWordApp.ActiveDocument.Name)
            
            If bSampleOnly = True Then
                Call ADDWATERMARK(oWordApp, oMergeDoc, "")
            End If

            strOutputFldr = QualifyFldrPath(strDocumentSaveBasePath) & QualifyFldrPath(oLetterInst.ProvNum)
            Call CreateFolder(strOutputFldr)
    
            If Not FolderExists(strOutputFldr) Then
                sErrMsg = "Provider folder for letter was not created for instance: " & oLetterInst.InstanceId & vbNewLine + vbNewLine & "Process will continue for the instances left."
                ' kd comeback, need to mark this instance as an error
                Stop
                GoTo Block_Err
            End If
            
            'Added to rename reprints...

            If bSampleOnly = True Then
                sOutFullPath = strOutputFldr & "" & oLetterInst.LetterType & "-" & oLetterInst.InstanceId & ".doc"
            Else
                If oLetterInst.LetterQueueStatus = "RR" Then
                'If pstrInstanceStatus = "R" Then
                    sOutFullPath = strOutputFldr & "" & oLetterInst.LetterType & "-Reprint-" & oLetterInst.InstanceId & ".doc"
                Else
                    sOutFullPath = strOutputFldr & "" & oLetterInst.LetterType & "-" & oLetterInst.InstanceId & ".doc"
                End If
            End If
            
            oMergeDoc.spellingchecked = True
            SleepEvents 1
            
            DoEvents
            DoEvents
            
            If UnlinkWordFields(oWordApp, oMergeDoc) = False Then
                LogMessage strProcName, "LETTER ERROR", "Failed to unlink word fields for some reason!", oLetterInst.InstanceId
                Stop
            End If
            oMergeDoc.SaveAs sOutFullPath
            sPreviewOutPath = sOutFullPath
            
                '' Collect static details
            If bSampleOnly = False Then
                With oLetterInst
        '            If .LetterBatchId = 0 Then
        '                .LetterBatchId = Me.MostRecentBatchId
        '            End If
            
                        '' KD: Idea: the below takes a long time for Word to determine how many pages
                        '' so we should probably move this to some other process
                        '' that runs virtually all the time.  We would look at a QUEUE table
                        '' and get the number of pages in those documents and then save them in the
                        '' LETTER_Static_Detail table
                        '' but this will have to happen #1: Quickly before the user gets the chance
                        '' to print the letters
                        '' and #2 quietly (without locking the documents and such)
                        
                        ''' Temporarilly commenting this out until I can come up with a better way..
            
                    DoEvents
                    .PageCount = oMergeDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
                    
                        ' Work some magic here to deal with MS Word's strangeness..
                        ' We have to sleep for just the right amount of time to get an accurate number of pages
                        ' here..
                    If .PageCount = 1 Then
                        dFileSize = FileLen(oMergeDoc.Path)
                        dSleepTime = (0.25 + (dFileSize * 0.0000001)) * 1000
                        
                        If dSleepTime > 2000 Then dSleepTime = 1500
                        DoEvents
    
                        oMergeDoc.Repaginate
                        DoEvents
                        Sleep dSleepTime
            '            Call SleepEvents(CLng(dSleepTime / 1000)) ' looks like we've been getting away with a max of 1.5 seconds
                                ' so I went back to sleep instead of taking a full 2 seconds each time
            
                        DoEvents
    
                        .PageCount = oMergeDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
                    End If
                    .LetterPath = sOutFullPath
                    .LetterNumInBatch = lCurrent
                    .SaveStaticDetails (True)
                End With
            End If
            
    
            oMergeDoc.Save
            If bSampleOnly = True Then
                oMergeDoc.Application.selection.HomeKey Unit:=wdStory
            Else
                oMergeDoc.Close
            End If
            
                    ' - Start a transaction
                    ' - Update the claim's:
                    '   - Reference table entry
                    '   - Advance claim status according to the process_logic table
                    ' - Commit the transaction
            
    
            '' Since we are still testing, the stored proc isn't being called to update the status (of the claim as well as the queue status..)
            '' therefore, for testing I'm going to do that here too:
            lTtlLtrsWithoutErrors = lTtlLtrsWithoutErrors + 1

           
            If bSampleOnly = False Then
                If bReprint = False Then
                    If UpdateDbWithLetterDetails(oLetterInst, sErrMsg) = False Then
                        LogMessage strProcName, "ERROR", "Could not advance some claims status' for instance id: " & oLetterInst.InstanceId, oLetterInst.InstanceId
                        Stop
                        GoTo NxtOne
                    End If
                End If
                
                If bReprint = True Then
                    ' need to update the paths if they changed..
                    If UpdateDbWithLetterDetailsForReprint(oLetterInst, sErrMsg) = False Then
                        LogMessage strProcName, "ERROR", "Could not update the db with letter details for Reprint instance id: " & oLetterInst.InstanceId, oLetterInst.InstanceId
                        Stop
                        GoTo NxtOne
                    End If
                End If
                
                Set oBAdo = New clsADO
                With oBAdo
                    .ConnectionString = DataConnString
                    .SQLTextType = sqltext
                    If bReprint = False Then
                        .sqlString = "update LETTER_Print_Queue SET Status = 'G' WHERE Status = 'QR' AND Error = 0 AND InstanceId = '" & oLetterInst.InstanceId & "' " ' AND AccountId = " & CStr(lThisAcct)
                    Else
                        ' this is actually already done in the above function..
                        .sqlString = "update LETTER_Print_Queue SET Status = 'P' WHERE Status = 'RR' AND Error = 0 AND InstanceId = '" & oLetterInst.InstanceId & "' " ' AND AccountId = " & CStr(lThisAcct)
                    End If
                    .Execute
                End With
            End If
            If bSampleOnly = True Then
                lRowsAffected = lTtlLtrsWithoutErrors
                PerformIndividualMailMerges = True
                GoTo Block_Exit
            End If
NxtOne:
            oRs.MoveNext

        Wend
NxtAcct:
        oAcctsRS.MoveNext
        
    Wend
    PerformIndividualMailMerges = True
    lRowsAffected = lTtlLtrsWithoutErrors
    
Block_Exit:
    ' Clean up any word objects..
    If Not oWordApp Is Nothing Then
        If oWordApp.Documents.Count > 0 Then
            For Each oTemplateDoc In oWordApp.Documents
                If bSampleOnly = True Then
                    If oTemplateDoc.Name <> oMergeDoc.Name Then
                        oTemplateDoc.Close False
                    End If
                Else
                    oTemplateDoc.Close False
                End If
            Next
        End If
'        If bSampleOnly = False Then
'        oWordApp.Quit
'        Set oWordApp = Nothing
'        End If
    End If
    Set oBAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


''' Recordset is
''' dictionary is
''' temp folder
'''
'''
'''
Public Function PerformIndividualMailMerges_LEGACY(oRs As ADODB.RecordSet, dtcTemplates As Scripting.Dictionary, Optional sTempFolder As String, Optional lRowsAffected As Long = -1) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim oLtrTmplt As clsLetterTemplate
Dim oWordApp As Word.Application
Dim oTemplateDoc As Word.Document
Dim oMergeDoc As Word.Document
Dim sTemplateFldr As String
Dim sLocalTemplatePath As String
Dim strODCFile As String
Dim strOutputFldr As String
Dim sOutFullPath As String
Dim strDocumentSaveBasePath As String
Dim oLetterInst As clsLetterInstance
Dim oLetters As clsLetterInstanceDct
Dim dFileSize As Double
Dim dSleepTime As Double
Dim sErrMsg  As String
Dim sLastTemplate As String
Dim lTtl As Long
Dim lCurrent As Long
Dim lTtlLtrsWithoutErrors As Long
Dim lPctDone As Long
Dim oBAdo As clsADO
Dim oAcctsRS As ADODB.RecordSet
Dim lThisAcct As Long


    '' This function really needs to be broken down into a couple separate functions..
    strProcName = ClassName & ".PerformIndividualMailMerges_LEGACY"
    
    ' Ok, so, get the list of accounts so we can loop over them
    ' we'll do that by querying the queue, not the acount table
    ' because there may not be any letters to process for a given account
    Set oAcctsRS = GetAccountsNConfigs()
    If oAcctsRS Is Nothing Then
        LogMessage strProcName, "ERROR", "Nothing to do?"
        Stop
    End If
    lTtl = oRs.recordCount
    
    While Not oAcctsRS.EOF
        lThisAcct = oAcctsRS("AccountId").Value
        strODCFile = oAcctsRS("ODC_ConnectionFile").Value
        strDocumentSaveBasePath = oAcctsRS("LetterOutputLocation").Value
        
        If sTempFolder = "" Then
            sTempFolder = GetUserTempDirectory()
        End If

        sTemplateFldr = sTempFolder
    '    Stop
        ' kev, need to get
        strDocumentSaveBasePath = GetSetting("LetterOutputLocation")
'        strODCFile = GetSetting("ODC_ConnectionFile")
        ' Let's filter the recordset by the accountid
        
        oRs.filter = "AccountId = " & CStr(lThisAcct)
        If oRs.BOF And oRs.EOF Then
            GoTo NxtAcct
        End If
        oRs.MoveFirst
        
'        lTtl = oRs.RecordCount
        If oRs.recordCount = 0 Then GoTo NxtAcct
        
        ' The plan here is to:
        ' For each letter (instanceid)
        '' RS has: InstanceId, BatchId, LetterType, AccountId, Status
        While Not oRs.EOF
            ' - check that all cnlyclaimnum's are in good status still
            '    - if not make it an error and mark the queue accordingly
            lCurrent = lCurrent + 1
            
            lPctDone = CDbl((lCurrent / lTtl) * 100)
            
    
            Set oLetterInst = New clsLetterInstance
            With oLetterInst
                .InstanceId = oRs("InstanceId").Value
                .ProvNum = oRs("CnlyProvId").Value
                .LetterBatchId = oRs("BatchId").Value
                .LetterType = UCase(oRs("LetterType").Value)
                .LetterCreateDt = oRs("LetterReqDt").Value
                .LetterQueueStatus = oRs("Status").Value
                .AccountID = lThisAcct
            End With
            
            
'' KD: ONly for real deal - not for samples
            Call UpdateProcessorStatus("Mail Merge", goBOLD_Processor.ThisQueueRunId, lPctDone, oRs("BatchId").Value, oRs("InstanceId").Value, False)
            
            
            If dtcTemplates.Exists(oLetterInst.LetterType) = False Then
                Stop
            End If
            Set oLtrTmplt = dtcTemplates.Item(oLetterInst.LetterType)
            
            
            sLocalTemplatePath = oLtrTmplt.TemplateLoc

'' KD: ONly for real deal - not for samples
            Call InsureAllStatusesAreGood(oRs("InstanceId").Value, lThisAcct)
            
            If oWordApp Is Nothing Then
                Set oWordApp = New Word.Application
                oWordApp.visible = False
            End If
    
            If sLastTemplate <> sLocalTemplatePath Then
                ' close the last one if we have it open:
                If oWordApp.Documents.Count > 0 Then
                    For Each oTemplateDoc In oWordApp.Documents
                        oTemplateDoc.Close False
                    Next
                End If
                Set oTemplateDoc = oWordApp.Documents.Add(sLocalTemplatePath, , False) 'tried didn't effect change
            Else
                If oWordApp.Documents.Count > 0 Then
                    Set oTemplateDoc = oWordApp.ActiveDocument
                Else
                    Set oTemplateDoc = oWordApp.Documents.Add(sLocalTemplatePath, , False) 'tried didn't effect change
                End If
                sLastTemplate = sLocalTemplatePath
            End If
            
            
            ' - get the mailmerge recordset (well, details we need for the sproc)
            ' sproc = usp_LETTER_Automation_MailMergeSource @pInstanceId (don't need accountid because instanceids are unique across accounts
            
            ' Set data source for mail merge.  Data will be from new Temp Table
            oTemplateDoc.MailMerge.OpenDataSource Name:=strODCFile, SqlStatement:="exec usp_LETTER_Automation_MailMergeSource '" & oLetterInst.InstanceId & "'"
            
            ' - open the word template from the dictionary.
            ' - Do the mail merge
            oTemplateDoc.MailMerge.MainDocumentType = 3 'wdDirectory
            oTemplateDoc.MailMerge.Destination = 0 'wdSendToNewDocument
            oTemplateDoc.MailMerge.Execute Pause:=False
            If left(oWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
                oWordApp.visible = True
                'MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
    '            Call ErrorCallStack_Add(clBatchId, "Error encountered with mail merge.", strProcName, "Instance ID: " & strInstanceId, , , strInstanceId, oLtrTmplt.LetterType)
    '            bMergeError = True
                oWordApp.ActiveDocument.Activate
                GoTo Block_Exit
            End If
            
            
            
            '' Ok, now deal with Barcode issues:
            Call AddSecPagesCode(oWordApp.ActiveDocument, oLetterInst)
            
            ' - save it where it needs to be
            Set oMergeDoc = oWordApp.Documents(oWordApp.ActiveDocument.Name)
            strOutputFldr = QualifyFldrPath(strDocumentSaveBasePath) & QualifyFldrPath(oLetterInst.ProvNum)
            Call CreateFolder(strOutputFldr)
    
            If Not FolderExists(strOutputFldr) Then
                sErrMsg = "Provider folder for letter was not created for instance: " & oLetterInst.InstanceId & vbNewLine + vbNewLine & "Process will continue for the instances left."
                ' kd comeback, need to mark this instance as an error
                GoTo Block_Err
            End If
            
            'Added to rename reprints...
            If oLetterInst.LetterQueueStatus = "R" Then
            'If pstrInstanceStatus = "R" Then
                sOutFullPath = strOutputFldr & "" & oLetterInst.LetterType & "-Reprint-" & oLetterInst.InstanceId & ".doc"
            Else
                sOutFullPath = strOutputFldr & "" & oLetterInst.LetterType & "-" & oLetterInst.InstanceId & ".doc"
            End If
            
            oMergeDoc.spellingchecked = True
            SleepEvents 1
            
            DoEvents
            DoEvents
            
            If UnlinkWordFields(oWordApp, oMergeDoc) = False Then
                LogMessage strProcName, "LETTER ERROR", "Failed to unlink word fields for some reason!", oLetterInst.InstanceId
            End If
            oMergeDoc.SaveAs sOutFullPath
            
            
                '' Collect static details
            With oLetterInst
    '            If .LetterBatchId = 0 Then
    '                .LetterBatchId = Me.MostRecentBatchId
    '            End If
        
                    '' KD: Idea: the below takes a long time for Word to determine how many pages
                    '' so we should probably move this to some other process
                    '' that runs virtually all the time.  We would look at a QUEUE table
                    '' and get the number of pages in those documents and then save them in the
                    '' LETTER_Static_Detail table
                    '' but this will have to happen #1: Quickly before the user gets the chance
                    '' to print the letters
                    '' and #2 quietly (without locking the documents and such)
                    
                    ''' Temporarilly commenting this out until I can come up with a better way..
        
                DoEvents
                .PageCount = oMergeDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
                
                    ' Work some magic here to deal with MS Word's strangeness..
                    ' We have to sleep for just the right amount of time to get an accurate number of pages
                    ' here..
                If .PageCount = 1 Then
                    dFileSize = FileLen(oMergeDoc.Path)
                    dSleepTime = (0.25 + (dFileSize * 0.0000001)) * 1000
                    
                    If dSleepTime > 2000 Then dSleepTime = 1500
                    DoEvents

                    oMergeDoc.Repaginate
                    DoEvents
                    Sleep dSleepTime
        '            Call SleepEvents(CLng(dSleepTime / 1000)) ' looks like we've been getting away with a max of 1.5 seconds
                            ' so I went back to sleep instead of taking a full 2 seconds each time
        
                    DoEvents

                    .PageCount = oMergeDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
                End If
        '        .PageCount = 0  ' just doing this for now
                .LetterPath = sOutFullPath
                .LetterNumInBatch = lCurrent
                .SaveStaticDetails (True)
            End With
                    
    
            oMergeDoc.Save
            oMergeDoc.Close
            
                    ' - Start a transaction
                    ' - Update the claim's:
                    '   - Reference table entry
                    '   - Advance claim status according to the process_logic table
                    ' - Commit the transaction
    '        If UpdateDbWithLetterDetails(oLetterInst, sErrMsg) = False Then
    '            Stop
    '        End If
            
    
                '' Since we are still testing, the stored proc isn't being called to update the status (of the claim as well as the queue status..)
                '' therefore, for testing I'm going to do that here too:
                lTtlLtrsWithoutErrors = lTtlLtrsWithoutErrors + 1
    
                Set oBAdo = New clsADO
                With oBAdo
                    .ConnectionString = DataConnString
                    .SQLTextType = sqltext
                    .sqlString = "update LETTER_Print_Queue SET Status = 'G' WHERE Status = 'QR' AND Error = 0 AND InstanceId = '" & oLetterInst.InstanceId & "' " ' AND AccountId = " & CStr(lThisAcct)
                    .Execute
                End With
            
            oRs.MoveNext
        Wend
NxtAcct:
        oAcctsRS.MoveNext
        
    Wend
    PerformIndividualMailMerges_LEGACY = True
    lRowsAffected = lTtlLtrsWithoutErrors
    
Block_Exit:
    ' Clean up any word objects..
    If Not oWordApp Is Nothing Then
        If oWordApp.Documents.Count > 0 Then
            For Each oTemplateDoc In oWordApp.Documents
                oTemplateDoc.Close False
            Next
        End If
        oWordApp.Quit
        Set oWordApp = Nothing
    End If
    Set oBAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function UpdateDbWithLetterDetails(oLetterInst As clsLetterInstance, Optional sErrMsg As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim oCmd As ADODB.Command

    strProcName = ClassName & ".UpdateDbWithLetterDetails"
    
    ' I should really make this a global object (or at least modular level) so I only have to open it once
    ' but let's call this my Scotty factor - when I need to speed it up, I'll do this.. :)
    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = CodeConnString
        .CursorLocation = adUseClient
        .Open
            ' We are going to begin a transaction here at the connection level
            ' even though all of our work will be in a sproc that should manage the transactions itself
        .BeginTrans
    End With
    
    If oLetterInst.LetterType = "RTRWD" Then
        LogMessage strProcName, "DEBUGGING", oLetterInst.LetterType & " Letter type about to roll forward in status...", oLetterInst.InstanceId
    End If
    
    

    Set oCmd = New ADODB.Command
    With oCmd
        Set .ActiveConnection = oCn
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_UpdateDbAfterLetterGeneration"
        .CommandTimeout = 300
        .Parameters.Refresh

        .Parameters("@pInstanceId") = oLetterInst.InstanceId
        .Parameters("@pLetterBatchId") = oLetterInst.BatchID
        .Parameters("@pLetterFullPath") = oLetterInst.LetterPath
        .Parameters("@pLetterType") = oLetterInst.LetterType
        .Parameters("@pAccountId") = oLetterInst.AccountID
'Stop    ' Kev, make sure this is right, and besides.. Make sure that we can run in parallel without screwing things up!
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
Stop
            sErrMsg = .Parameters("@pErrMsg").Value
            LogMessage strProcName, "ERROR", "Problem moving claim status' forward", oLetterInst.InstanceId & " : " & oLetterInst.LetterType
            oCn.RollbackTrans
            GoTo Block_Exit
        End If
    
    End With
    
    
        
    If Not oCn Is Nothing Then
        UpdateDbWithLetterDetails = True
 
        If oLetterInst.LetterType = "RTRWD" Then
            LogMessage strProcName, "DEBUGGING", oLetterInst.LetterType & " Letter type about to commitTrans!", oLetterInst.InstanceId
        End If
        
        oCn.CommitTrans
    End If
    
    UpdateDbWithLetterDetails = True
    
Block_Exit:
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If
    Exit Function
Block_Err:
    ReportError Err, strProcName
    sErrMsg = Err.Description
    If Not oCn Is Nothing Then
        oCn.RollbackTrans
    End If
    GoTo Block_Exit
End Function


Public Function UpdateDbWithLetterDetailsForReprint(oLetterInst As clsLetterInstance, Optional sErrMsg As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim oCmd As ADODB.Command

    strProcName = ClassName & ".UpdateDbWithLetterDetailsForReprint"
    
    ' I should really make this a global object (or at least modular level) so I only have to open it once
    ' but let's call this my Scotty factor - when I need to speed it up, I'll do this.. :)
    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = CodeConnString
        .CursorLocation = adUseClient
        .Open
            ' We are going to begin a transaction here at the connection level
            ' even though all of our work will be in a sproc that should manage the transactions itself
        .BeginTrans
    End With
    
    
    Set oCmd = New ADODB.Command
    With oCmd
        Set .ActiveConnection = oCn
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_UpdateDbAfterLetterGenerationForReprint"
        .Parameters.Refresh

        .Parameters("@pInstanceId") = oLetterInst.InstanceId
        .Parameters("@pLetterBatchId") = oLetterInst.BatchID
        .Parameters("@pLetterFullPath") = oLetterInst.LetterPath
        .Parameters("@pLetterType") = oLetterInst.LetterType
        .Parameters("@pAccountId") = oLetterInst.AccountID
'Stop    ' Kev, make sure this is right, and besides.. Make sure that we can run in parallel without screwing things up!
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            sErrMsg = .Parameters("@pErrMsg").Value
            LogMessage strProcName, "ERROR", "Problem moving claim status' forward", oLetterInst.InstanceId & " : " & oLetterInst.LetterType
            oCn.RollbackTrans
            GoTo Block_Exit
        End If
    
    End With
    
    
        
    If Not oCn Is Nothing Then
        UpdateDbWithLetterDetailsForReprint = True
        oCn.CommitTrans
    End If
    
'    UpdateDbWithLetterDetailsForReprint = True
    
Block_Exit:
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If
    Exit Function
Block_Err:
    ReportError Err, strProcName
    sErrMsg = Err.Description
    If Not oCn Is Nothing Then
        oCn.RollbackTrans
    End If
    GoTo Block_Exit
End Function

Public Function InsureAllStatusesAreGood(strInstanceID As String, lAccountId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim lClaimsLeft As Long

    strProcName = ClassName & ".InsureAllStatusesAreGood"
    ' KD: COMEBACK: if we don't get the expected ClaimsLeft then we need to log that somewhere - so a claim is accounted for
    '   we also will need the instance id so if they want to fix the claim's status and regenerate the letter they can
    '   easily.
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_CheckClaimStatus"
        .Parameters.Refresh
        .Parameters("@pInstanceId") = strInstanceID
        .Parameters("@pAccountId") = lAccountId
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was an error when checking claim statuses for instance " & strInstanceID, .Parameters("@pErrMsg").Value
'            GoTo Block_Exit
            Stop
        End If
        lClaimsLeft = Nz(.Parameters("@pClaimsLeft").Value, 0)
    End With
    
    If lClaimsLeft < 1 Then
        LogMessage strProcName, , "No valid claims left for this instanceid: " & strInstanceID, CStr(lAccountId)
        InsureAllStatusesAreGood = False
    Else
        InsureAllStatusesAreGood = True
    End If
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Function PrintLetterInstance(oLetterInst As clsLetterInstance, pstrTemplateName As String, _
            pstrOutputFileName As String, pstrOutputBasePath As String, pstrProvNum As String, _
            pstrODCFile As String, pstrLetterType As String, _
            Optional iPageCount As Integer) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oCn As ADODB.Connection
Dim oCmd As ADODB.Command
Dim strSQLcmd As String
Dim bMergeError As Boolean
Dim strOutputPath As String
Dim strChkFile As String
Dim strErrMsg As String
Dim iRtnCd As Integer
Dim varItem As Variant
Dim iAnswer As Integer
Dim iCnt As Integer
Dim i As Integer
Dim dThisOne As Date
Dim dtSt As Date
Dim dFileSize As Double
Dim dSleepTime As Double
Dim objLetterInfo As clsLetterTemplate
Dim objWordApp As Word.Application, _
    objWordDoc As Word.Document, _
    objWordMergedDoc As Word.Document
Dim oProp As ADODB.Property, iPropCnt As Integer

      Debug.Print "Outfile path: " & pstrOutputBasePath
      
    strProcName = ClassName & ".PrintLetterInstance"
    
    Set oAdo = New clsADO
    oAdo.ConnectionString = CodeConnString
    
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = CodeConnString
    oCn.CursorLocation = adUseClient
    oCn.Open
    
    '' check to make sure that the transaction is supported
    
'    For iPropCnt = 0 To oCn.Properties.Count
'        Set oProp = oCn.Properties(iPropCnt)
'    Next
'    For Each oProp In oCn.Properties
'        Debug.Print oProp.Name
''        Stop
'    Next
'    Set oProp = Nothing
'   Stop
'   For Each oProp In oAdo.CurrentConnection.Properties
'        Debug.Assert oProp.Name <> "Transaction DDL"
'   Next
'
'   Stop
   
    Set objLetterInfo = New clsLetterTemplate
    strErrMsg = ""

    Set objWordApp = New Word.Application
    objWordApp.visible = False
    
    ' check if template exists
    strChkFile = Dir(pstrTemplateName)
    If strChkFile = "" Then
        strErrMsg = "Missing letter template." & vbCrLf & "Template name = " & objLetterInfo.TemplateLoc & ""
        GoTo Block_Err
    End If

    ' open template
    Set objWordDoc = objWordApp.Documents.Add(pstrTemplateName, , False)
       
    ' load letter info
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.commandType = adCmdStoredProc
    
    '' KD Added this for now to deal with the QR 2D barcodes..
    '' we'll change this and modify the real usp when we decide to
    ''  "go live" with it.
Dim sMergeSproc As String

    If UCase(oLetterInst.LetterType) = "VADRA_QR" Then
        oCmd.CommandText = "usp_LETTER_Get_Info_load_KD"
        sMergeSproc = "usp_LETTER_Get_Info_KD"
    Else
        oCmd.CommandText = "usp_LETTER_Get_Info_load"
        sMergeSproc = "usp_LETTER_Get_Info"
    End If
    
    oCmd.Parameters.Refresh
    oCmd.Parameters("@InstanceID") = oLetterInst.InstanceId
    oCmd.Execute
    
    strErrMsg = Trim(oCmd.Parameters("@ErrMsg").Value) & ""
    If strErrMsg <> "" Then
        GoTo Block_Err
    End If
    
    dtSt = Now  ' For tracking how much time this one takes
    
    ' Set data source for mail merge.  Data will be from new Temp Table
    objWordDoc.MailMerge.OpenDataSource Name:=pstrODCFile, _
                        SqlStatement:="exec " & sMergeSproc & " '" & oLetterInst.InstanceId & "'"
                    


    ' Perform mail merge.
    objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
    objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
    objWordDoc.MailMerge.Execute Pause:=False
    If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
        objWordApp.visible = True
        '''MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
        bMergeError = True
        objWordApp.ActiveDocument.Activate
        strErrMsg = "Error encountered with mail merge."
        GoTo Block_Err
    End If

    
    ' 20130219 KD: Add the Sec Pages field in the footer for barcodes
Stop
    Call AddSecPagesCode(objWordApp.ActiveDocument, oLetterInst)
    
    ' Save the output doc
    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
    strOutputPath = pstrOutputBasePath & "\" & pstrProvNum & "\"
    Call CreateFolders(strOutputPath)

    If Not FolderExists(strOutputPath) Then
        strErrMsg = "Provider folder for letter was not created for instance: " & oLetterInst.InstanceId & vbNewLine + vbNewLine & "Process will continue for the instances left."
        GoTo Block_Err
    End If
    
    If oLetterInst.LetterQueueStatus = "R" Then
    'If pstrInstanceStatus = "R" Then
        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-Reprint-" & oLetterInst.InstanceId & ".doc"
    Else
        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-" & oLetterInst.InstanceId & ".doc"
    End If
    
    objWordMergedDoc.spellingchecked = True
    objWordMergedDoc.Repaginate
    DoEvents
    DoEvents
    
    ' 20130219 KD: Add the Sec Pages field in the footer for barcodes
'Stop
'    Call AddSecPagesCode(objWordApp.ActiveDocument)
    
    If UnlinkWordFields(objWordApp, objWordMergedDoc) = False Then
        LogMessage strProcName, "LETTER ERROR", "Failed to unlink word fields for some reason!", oLetterInst.InstanceId
    End If

    
    On Error Resume Next
    '' Note to Data Services user:
    '' sometimes this fails - there seems to be some sort of
    '' delay in the filesystem - perhaps it's anti-virus doing it's scan
    '' either way, if you get a run-time error on the next line of code, just press f5
    '' to continue execution - it should work the 2nd time.
    '' the code I've put in will hopefully take care of it if you don't have break on all errors turned
    '' on..
    objWordMergedDoc.SaveAs pstrOutputFileName
    If Err.Number <> 0 Then
        SleepEvents 1
        objWordMergedDoc.SaveAs pstrOutputFileName
        Err.Clear
    End If
    On Error GoTo Block_Err

    
    With oLetterInst
        If .LetterBatchId = 0 Then
'            .LetterBatchId = MostRecentBatchId
        End If

            '' KD: Idea: the below takes a long time for Word to determine how many pages
            '' so we should probably move this to some other process
            '' that runs virtually all the time.  We would look at a QUEUE table
            '' and get the number of pages in those documents and then save them in the
            '' LETTER_Static_Detail table
            '' but this will have to happen #1: Quickly before the user gets the chance
            '' to print the letters
            '' and #2 quietly (without locking the documents and such)
            
            ''' Temporarilly commenting this out until I can come up with a better way..

        DoEvents
        .PageCount = objWordMergedDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
        
        If .PageCount = 1 Then
            dFileSize = FileLen(objWordMergedDoc.Path)
            dSleepTime = (0.25 + (dFileSize * 0.0000001)) * 1000
            
            If dSleepTime > 2000 Then dSleepTime = 1500
            DoEvents

'If gbVerboseLogging = True Then LogMessage strProcName, "LETTER", "Sleeping for " & CStr(dSleepTime) & " milliseconds"
            objWordMergedDoc.Repaginate
            DoEvents
            Sleep dSleepTime
'            Call SleepEvents(CLng(dSleepTime / 1000))

            DoEvents
            .PageCount = objWordMergedDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
'If gbVerboseLogging = True Then LogMessage strProcName, "LETTER", "Page count: " & CStr(.PageCount)
        End If
'        .PageCount = 0  ' just doing this for now
        .LetterPath = pstrOutputFileName

Stop
        .SaveStaticDetails (True)
    End With

    
   
'If gbVerboseLogging = True Then Debug.Print ProcessTookHowLong(dtSt)
    
    If Not objWordDoc Is Nothing Then '07/01/2013
        objWordDoc.Close wdDoNotSaveChanges
    End If
    
    If Not objWordMergedDoc Is Nothing Then '07/01/2013
        objWordMergedDoc.Close wdDoNotSaveChanges
    End If
    
    Set objWordDoc = Nothing
    Set objWordMergedDoc = Nothing
    
    If Not objWordApp Is Nothing Then
        On Error Resume Next
        objWordApp.Quit wdDoNotSaveChanges
        On Error GoTo Block_Err
    End If
    Set objWordApp = Nothing

    DoEvents
    DoEvents

    If Not FileExists(pstrOutputFileName) Then
        strErrMsg = "PrintLetterInstance: Letter was not created for instance " & oLetterInst.InstanceId & vbNewLine + vbNewLine & "Process will continue for the instances left."
        LogMessage strProcName, "ERROR", "Generated letter does not exist where it should", pstrOutputFileName
        
        GoTo Block_Err
    End If

    
    ' KD: 5/8/2014: So until today, this, clear tmp table stuff
    ' was BEFORE the claims status got updated, but one of the criteria for the sproc is:
    ' where t2.Status not in ('R','W')
    ' Which means that it's not getting cleared..
    
    ' clear letter info
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = "usp_LETTER_Get_Info_tmp_clear"
    oCmd.Parameters.Refresh
    'oCmd.Parameters("@pInstanceID") = pstrInstanceID
    oCmd.Execute
    
                                
    ' start letter transaction
'    oAdo.BeginTrans
    oCn.BeginTrans
    
    ' update LETTER status
    Set oCmd = New ADODB.Command
    'oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.ActiveConnection = oCn
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = "usp_LETTER_Update_Status"
    oCmd.Parameters.Refresh
    oCmd.Parameters("@InstanceID").Value = oLetterInst.InstanceId
    oCmd.Parameters("@LetterName").Value = pstrOutputFileName
    oCmd.Parameters("@pNextStatus").Value = "G" ' for Generated, not yet printed..
    oCmd.Execute
            
    strErrMsg = Trim(oCmd.Parameters("@ErrMsg").Value) & ""
    If strErrMsg <> "" Then
'        oAdo.RollbackTrans
        LogMessage strProcName, "ERROR", oCmd.CommandText & " ERROR: " & strErrMsg, oLetterInst.InstanceId
        oCn.RollbackTrans
        GoTo Block_Err
    End If
                            

                            
                            
    ' update claim status & move to next queue
    ' note, only does this where the letter status is W
    Set oCmd = New ADODB.Command
    'oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.ActiveConnection = oCn
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = "usp_LETTER_AuditClaims_Update"
    oCmd.Parameters.Refresh
    oCmd.Parameters("@pInstanceID").Value = oLetterInst.InstanceId
    oCmd.Parameters("@pInstanceStatus").Value = oLetterInst.LetterQueueStatus
    oCmd.Execute
            
    strErrMsg = Trim(Nz(oCmd.Parameters("@pErrMsg").Value, ""))
    If strErrMsg <> "" Then
'        oAdo.RollbackTrans
        LogMessage strProcName, "ERROR", oCmd.CommandText & " ERROR: " & strErrMsg, oLetterInst.InstanceId

        oCn.RollbackTrans
        GoTo Block_Err
    End If
                                
                                
    ' commit letter transaction
'    oAdo.CommitTrans
    oCn.CommitTrans
    PrintLetterInstance = True
    
    



Block_Exit:

'    Call SetDefaultPrinterToAcrobat("", sOrigPrinter)

    ' Release references.
    If Not objWordDoc Is Nothing Then '07/01/2013
        objWordDoc.Close wdDoNotSaveChanges
    End If
    
    If Not objWordMergedDoc Is Nothing Then '07/01/2013
        objWordMergedDoc.Close wdDoNotSaveChanges
    End If
    
    Set objWordDoc = Nothing
    Set objWordMergedDoc = Nothing
    
    If Not objWordApp Is Nothing Then
        On Error Resume Next
        objWordApp.Quit wdDoNotSaveChanges
        On Error GoTo 0
    End If
    Set objWordApp = Nothing
    
    Set oCmd = Nothing
    Set oAdo = Nothing
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If
    Exit Function
Block_Err:

    If strErrMsg <> "" Then
'        MsgBox strErrMsg, vbCritical
        LogMessage strProcName, "USAGE DETAIL", strErrMsg
Stop
'Call ErrorCallStack_Add(clBatchId, strErrMsg, strProcName)
    Else
'        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
        ReportError Err, strProcName
Stop
'        Call ErrorCallStack_Add(clBatchId, Err.Description, strProcName)
    End If
    PrintLetterInstance = False
    
    'Call DeleteFile(pstrOutputFileName, False)
    
    GoTo Block_Exit
End Function


Public Function GetUserTempDirectory() As String
On Error GoTo Block_Err
Dim strProcName As String
Dim strTempPath As String
Dim strErrMsg As String

    strProcName = ClassName & ".GetUserTempDirectory"


'    strTempPath = cs_USER_TEMPLATE_PATH_ROOT & GetUserName & "\LETTERTempDir\"
    
    strTempPath = QualifyFldrPath(GetSystemTempFolder()) & "LETTERTempDir\"
    
    
    ' Make sure it's empty (if it already exists)
    Call DeleteFullFolder(strTempPath, False)
        
    
    If CreateFolders(strTempPath) = False Then
        LogMessage strProcName, "ERROR", "Could not create user temp folder!", strTempPath, True
        GoTo Block_Exit
    End If
            
    If FolderExist(strTempPath) = False Then
        strErrMsg = "ERROR: can not create folder " & strTempPath
        GoTo Block_Err
    End If
    
    
Block_Exit:
    GetUserTempDirectory = strTempPath
    Exit Function
Block_Err:
    If Err.Number <> 0 Then
        ReportError Err, strProcName
    Else
        LogMessage strProcName, "ERROR", strErrMsg
    End If
    GoTo Block_Exit
End Function

Function GetUserOutDirectory() As String
On Error GoTo Block_Err
Dim strProcName As String
Dim strTempPath As String
Dim strErrMsg As String

    strProcName = ClassName & ".GetUserOutDirectory"


'    strTempPath = cs_USER_TEMPLATE_PATH_ROOT & GetUserName & "\LETTERTempDir\"
    strTempPath = GetSetting("MAILROOM_LETTER_PATH")
    
    ' Make sure it's empty (if it already exists)
    Call DeleteFullFolder(strTempPath)
        
    
    If CreateFolders(strTempPath) = False Then
        LogMessage strProcName, "ERROR", "Could not create user temp folder!", strTempPath, True
        GoTo Block_Exit
    End If
            
    If FolderExist(strTempPath) = False Then
        strErrMsg = "ERROR: can not create folder " & strTempPath
        GoTo Block_Err
    End If
    
    
Block_Exit:
    GetUserOutDirectory = strTempPath
    Exit Function
Block_Err:
    If Err.Number <> 0 Then
        ReportError Err, strProcName
    Else
        LogMessage strProcName, "ERROR", strErrMsg
    End If
    GoTo Block_Exit
End Function



Public Sub SamplePseudoWordMailMergeCode()
'
'    ' Set data source for mail merge.  Data will be from new Temp Table
'    objWordDoc.MailMerge.OpenDataSource Name:=pstrODCFile, _
'                        SQLStatement:="exec usp_LETTER_Get_Info '" & oLetterInst.InstanceID & "'"
'
'
'    ' Perform mail merge.
'    objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
'    objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
'    objWordDoc.MailMerge.Execute Pause:=False
'    If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
'        objWordApp.visible = True
'        MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
'        bMergeError = True
'        objWordApp.ActiveDocument.Activate
'        strErrMsg = "Error encountered with mail merge."
'        GoTo Block_Err
'    End If
'
'    ' 20130219 KD: Add the Sec Pages field in the footer for barcodes
'    Call AddSecPagesCode(objWordApp.ActiveDocument)
'
'
'    ' Save the output doc
'    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
'    strOutputPath = pstrOutputBasePath & "\" & pstrProvNum & "\"
'    Call CreateFolders(strOutputPath)
'
'    If Not FolderExists(strOutputPath) Then
'        strErrMsg = "Provider folder for letter was not created for instance: " & oLetterInst.InstanceID & vbNewLine + vbNewLine & "Process will continue for the instances left."
'        GoTo Block_Err
'    End If
'
'    If oLetterInst.LetterQueueStatus = "R" Then
'    'If pstrInstanceStatus = "R" Then
'        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-Reprint-" & oLetterInst.InstanceID & ".doc"
'    Else
'        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-" & oLetterInst.InstanceID & ".doc"
'    End If
'
'    objWordMergedDoc.spellingchecked = True
'    objWordMergedDoc.Repaginate
'
'    If UnlinkWordFields(objWordApp, objWordMergedDoc) = False Then
'        LogMessage strProcName, "ERROR", "There was an error unlinking the fields. Check that the fields are correct!", pstrOutputFileName, True
'    End If
'
'    objWordMergedDoc.SaveAs pstrOutputFileName
'    SleepEvents 1
'
'    With oLetterInst
'        If .LetterBatchId = 0 Then
'            .LetterBatchId = Me.MostRecentBatchId
'        End If
'        If objWordMergedDoc.BuiltInDocumentProperties(14) = 1 Then
'            Stop
'        End If
'        .PageCount = objWordMergedDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
'        .LetterPath = pstrOutputFileName
'        .SaveStaticDetails
'    End With
'
'    objWordMergedDoc.Close
'
'    Set objWordMergedDoc = Nothing
'
'    If Not FileExists(pstrOutputFileName) Then
'        strErrMsg = "PrintLetterInstance: Letter was not created for instance " & oLetterInst.InstanceID & vbNewLine + vbNewLine & "Process will continue for the instances left."
'        LogMessage strProcName, "ERROR", "Generated letter does not exist where it should", pstrOutputFileName
'
'        GoTo Block_Err
'    End If
'
'
'Block_Exit:
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''
''''' Required functions
'''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'' This function will look for a bookmark named 'SecPages' and will replace that with
'' the Sec Pages field
Public Function AddSecPagesCode(objWordDoc As Object, oLetterInst As clsLetterInstance) As Boolean
'Private Function AddSecPagesCode(objWordDoc As Word.Document) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim objWordField As FormField
Dim sInstanceId As String

'Dim objWordSection As Object
Dim objWordSection As Object
Dim iFoot As Integer
Dim oFooter As Object
Dim oField As Object
Dim oRange As Object
Dim sBookmarkName As String
Dim saryBkmarks(4) As String
Dim iBkmark As Integer
Dim iCurrentSetting As Integer
Dim oTFrame As Word.TextFrame
Dim dtStart As Date
Dim lTtlPageCnt As Long
Dim iDuplexPageCnt As Integer
Dim lPageCount As Long

    strProcName = ClassName & ".AddSecPagesCode"
    dtStart = Now()
    LogMessage strProcName, "EFFICIENCY TESTING", "Starting"
'Stop
    sInstanceId = oLetterInst.InstanceId
'    lTtlPageCnt = oLetterInst.PageCount
'    iDuplexPageCnt = CInt(lTtlPageCnt / 2)
    
    saryBkmarks(0) = "SecPages"
    saryBkmarks(1) = "SecPages2"
    saryBkmarks(2) = "SecPagesDuplex"
    saryBkmarks(3) = "ProvNum"
    saryBkmarks(4) = "ProvNum2"
    
    
    'oletterinst.ProvNum
    
    DoEvents
    DoEvents
    DoEvents
    SleepEvents 1
    
    ' I hate to do this but I don't have time to figure out a better
    ' way to determine which is going to work...
    iCurrentSetting = Application.GetOption("Error Trapping")
    Application.SetOption "Error Trapping", 2
    
'Debug.Assert objWordDoc.Sections.Count < 6

    
    For iBkmark = 0 To UBound(saryBkmarks)
        sBookmarkName = saryBkmarks(iBkmark)
    
        If IsBookMark(objWordDoc, sBookmarkName) = False Then
            ' nothing else to do
            GoTo NextBkmark
        End If
        
        For Each objWordSection In objWordDoc.Sections
            For iFoot = 1 To objWordSection.Footers.Count
                    ' shape range should take care of the text box..
                    '' Note: may want to do the Headers here too
                Set oFooter = objWordSection.Footers.Item(iFoot)
                    ' Looks like 2007 + use a different means..
                If oFooter.Shapes.Count > 0 Then
                
                    If IsBookMark(objWordDoc, sBookmarkName) = False Then
                        ' nothing else to do
                        GoTo NextBkmark
                    End If
                    
                    On Error Resume Next
                        ' I hate to do this but I don't have time to figure out a better
                        ' way to determine which is going to work...


                    If oFooter.Shapes(1).TextFrame Is Nothing Then
                        Err.Clear
                        GoTo OldVersion
                    End If
                    
                    Set oTFrame = oFooter.Shapes(1).TextFrame

                    On Error Resume Next
                    If oTFrame.TextRange Is Nothing Then
                        If Err.Number <> 0 Then
                            Err.Clear
                            GoTo OldVersion
                        End If
                    End If
                    If Err.Number <> 0 Then
                        Err.Clear
                        GoTo OldVersion
                    End If


                    If oTFrame.TextRange.Fields Is Nothing Then
                        If Err.Number <> 0 Then
                            Err.Clear
                            GoTo OldVersion
                        End If
                    End If
                    

                    On Error GoTo Block_Err

                    objWordDoc.Repaginate

                    lPageCount = objWordDoc.BuiltInDocumentProperties(14)
    

                    Set oRange = objWordDoc.Bookmarks(sBookmarkName).Range
                    
'                    Set oField = oFooter.Shapes(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
                    Set oField = oFooter.Shapes(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldEmpty, "\#""00""", True)
'                    Stop
                    If InStr(1, sBookmarkName, "duplex", vbTextCompare) > 0 Then
                        oField.Result.Text = Format(lPageCount / 2, "00")
                    Else
                        If left(sBookmarkName, 7) = "ProvNum" Then
                            oField.Result.Text = oLetterInst.ProvNum
                        Else
                            oField.Result.Text = Format(lPageCount, "00")
                        End If
                    End If
                    
                    
                    ' now lets remove the bookmark all together - I've been finding examples where this is replaced several times
                    ' resulting in barcodes like 01030303
                    ' when it should be 0103
'                    oField.Unlink
                    objWordDoc.Bookmarks(sBookmarkName).Delete
                    
                ElseIf oFooter.Range.ShapeRange.Count > 0 Then
OldVersion:
                    If IsBookMark(objWordDoc, sBookmarkName) = False Then
                        ' nothing else to do
                        GoTo NextBkmark
                    End If
'                    oField.Unlink
                    
                    objWordDoc.Repaginate

                    lPageCount = objWordDoc.BuiltInDocumentProperties(14)
                        
                    
                    Set oRange = objWordDoc.Bookmarks(sBookmarkName).Range
                    ''Set oField = oFooter.Range.ShapeRange.Item(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
                    Set oField = oFooter.Range.ShapeRange.Item(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldEmpty, "\#""00""", True)
                    If InStr(1, sBookmarkName, "duplex", vbTextCompare) > 0 Then
                        oField.Result.Text = Format(lPageCount / 2, "00")
                    Else
                        If left(sBookmarkName, 7) = "ProvNum" Then
                            oField.Result.Text = oLetterInst.ProvNum
                        Else
                            oField.Result.Text = Format(lPageCount, "00")
                        End If
                    End If

                    objWordDoc.Bookmarks(sBookmarkName).Delete

                    
                End If
            Next
        Next
NextBkmark:
    Next
    
    AddSecPagesCode = True  ' In this case true = no error.. :)
    
Block_Exit:
    ' Make sure we set this back the way it was..
    Application.SetOption "Error Trapping", iCurrentSetting
    Set oField = Nothing
    Set oRange = Nothing
    Set oFooter = Nothing
    Set objWordSection = Nothing
    LogMessage strProcName, "EFFICIENCY TESTING", "Ending", CStr("" & ProcessTookHowLong(dtStart))
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



'' This function will look for a bookmark named 'SecPages' and will replace that with
'' the Sec Pages field
Public Function AddSecPagesCodeNoObjects(objWordDoc As Object, sInstanceId As String) As Boolean
'Private Function AddSecPagesCode(objWordDoc As Word.Document) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim objWordField As FormField

'Dim objWordSection As Object
Dim objWordSection As Object
Dim iFoot As Integer
Dim oFooter As Object
Dim oField As Object
Dim oRange As Object
Dim sBookmarkName As String
Dim saryBkmarks(2) As String
Dim iBkmark As Integer
Dim iCurrentSetting As Integer
Dim oTFrame As Word.TextFrame
Dim dtStart As Date
Dim lTtlPageCnt As Long
Dim iDuplexPageCnt As Integer
Dim lPageCount As Long

    strProcName = ClassName & ".AddSecPagesCodeNoObjects"
    dtStart = Now()
    LogMessage strProcName, "EFFICIENCY TESTING", "Starting"
'Stop

    
    saryBkmarks(0) = "SecPages"
    saryBkmarks(1) = "SecPages2"
    saryBkmarks(2) = "SecPagesDuplex"
    
    DoEvents
    DoEvents
    DoEvents
    SleepEvents 1
    
    ' I hate to do this but I don't have time to figure out a better
    ' way to determine which is going to work...
    iCurrentSetting = Application.GetOption("Error Trapping")
    Application.SetOption "Error Trapping", 2
    
'Debug.Assert objWordDoc.Sections.Count < 6

    
    For iBkmark = 0 To UBound(saryBkmarks)
        sBookmarkName = saryBkmarks(iBkmark)
    
        If IsBookMark(objWordDoc, sBookmarkName) = False Then
            ' nothing else to do
            GoTo NextBkmark
        End If
        
        For Each objWordSection In objWordDoc.Sections
            For iFoot = 1 To objWordSection.Footers.Count
                    ' shape range should take care of the text box..
                    '' Note: may want to do the Headers here too
                Set oFooter = objWordSection.Footers.Item(iFoot)
                    ' Looks like 2007 + use a different means..
                If oFooter.Shapes.Count > 0 Then
                
                    If IsBookMark(objWordDoc, sBookmarkName) = False Then
                        ' nothing else to do
                        GoTo NextBkmark
                    End If
                    
                    On Error Resume Next
                        ' I hate to do this but I don't have time to figure out a better
                        ' way to determine which is going to work...


                    If oFooter.Shapes(1).TextFrame Is Nothing Then
                        Err.Clear
                        GoTo OldVersion
                    End If
                    
                    Set oTFrame = oFooter.Shapes(1).TextFrame

                    On Error Resume Next
                    If oTFrame.TextRange Is Nothing Then
                        If Err.Number <> 0 Then
                            Err.Clear
                            GoTo OldVersion
                        End If
                    End If
                    If Err.Number <> 0 Then
                        Err.Clear
                        GoTo OldVersion
                    End If


                    If oTFrame.TextRange.Fields Is Nothing Then
                        If Err.Number <> 0 Then
                            Err.Clear
                            GoTo OldVersion
                        End If
                    End If
                    

                    On Error GoTo Block_Err

                    objWordDoc.Repaginate

                    lPageCount = objWordDoc.BuiltInDocumentProperties(14)
    

                    Set oRange = objWordDoc.Bookmarks(sBookmarkName).Range
                    
'                    Set oField = oFooter.Shapes(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
                    Set oField = oFooter.Shapes(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldEmpty, "\#""00""", True)
'                    Stop
                    If InStr(1, sBookmarkName, "duplex", vbTextCompare) > 0 Then
                        oField.Result.Text = Format(lPageCount / 2, "00")
                    Else
                        oField.Result.Text = Format(lPageCount, "00")
                    End If
                    
                    
                    ' now lets remove the bookmark all together - I've been finding examples where this is replaced several times
                    ' resulting in barcodes like 01030303
                    ' when it should be 0103
'                    oField.Unlink
                    objWordDoc.Bookmarks(sBookmarkName).Delete
                    
                ElseIf oFooter.Range.ShapeRange.Count > 0 Then
OldVersion:
                    If IsBookMark(objWordDoc, sBookmarkName) = False Then
                        ' nothing else to do
                        GoTo NextBkmark
                    End If
'                    oField.Unlink
                    
                    objWordDoc.Repaginate

                    lPageCount = objWordDoc.BuiltInDocumentProperties(14)
                        
                    
                    Set oRange = objWordDoc.Bookmarks(sBookmarkName).Range
                    ''Set oField = oFooter.Range.ShapeRange.Item(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
                    Set oField = oFooter.Range.ShapeRange.Item(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldEmpty, "\#""00""", True)
                    If InStr(1, sBookmarkName, "duplex", vbTextCompare) > 0 Then
                        oField.Result.Text = Format(lPageCount / 2, "00")
                    Else
                        oField.Result.Text = Format(lPageCount, "00")
                    End If

                    objWordDoc.Bookmarks(sBookmarkName).Delete

                    
                End If
            Next
        Next
NextBkmark:
    Next
    
    AddSecPagesCodeNoObjects = True  ' In this case true = no error.. :)
    
Block_Exit:
    ' Make sure we set this back the way it was..
    Application.SetOption "Error Trapping", iCurrentSetting
    Set oField = Nothing
    Set oRange = Nothing
    Set oFooter = Nothing
    Set objWordSection = Nothing
    LogMessage strProcName, "EFFICIENCY TESTING", "Ending", CStr("" & ProcessTookHowLong(dtStart))
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


'' This function will look for a bookmark named 'SecPages' and will replace that with
'' the Sec Pages field
Public Function AddInstanceIdQRCode(objWordDoc As Object, sInstanceId As String, sQRCodePath As String) As Boolean
'Private Function AddSecPagesCode(objWordDoc As Word.Document) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim objWordField As FormField

'Dim objWordSection As Object
Dim objWordSection As Object
Dim iFoot As Integer
Dim oFooter As Object
Dim oField As Object
Dim oRange As Object
Dim sBookmarkName As String
Dim saryBkmarks(1) As String
Dim iBkmark As Integer
Dim iCurrentSetting As Integer
Dim oTFrame As Word.TextFrame
Dim dtStart As Date
'Dim saryBkmarks(1) As String
'Dim iBkmark As Integer


    strProcName = ClassName & ".AddInstanceIdQRCode"
    
    saryBkmarks(0) = "InstanceIdBarcode"
    saryBkmarks(1) = "InstanceIdBarcode2"
    
    
    
    
    dtStart = Now()
    LogMessage strProcName, "EFFICIENCY TESTING", "Starting"
    
    DoEvents
    DoEvents
    DoEvents
    SleepEvents 1
    
    ' I hate to do this but I don't have time to figure out a better
    ' way to determine which is going to work...
    iCurrentSetting = Application.GetOption("Error Trapping")
    Application.SetOption "Error Trapping", 2
    
    sBookmarkName = "InstanceIdBarcode"
    For iBkmark = 0 To UBound(saryBkmarks)
        sBookmarkName = saryBkmarks(iBkmark)
        
        If IsBookMark(objWordDoc, sBookmarkName) = True Then
          For Each objWordSection In objWordDoc.Sections
                For iFoot = 1 To objWordSection.Footers.Count
                        ' shape range should take care of the text box..
                        '' Note: may want to do the Headers here too
                    Set oFooter = objWordSection.Footers.Item(iFoot)
                        ' Looks like 2007 + use a different means..
                    If oFooter.Shapes.Count > 0 Then
                    
                        If IsBookMark(objWordDoc, sBookmarkName) = False Then
                            ' nothing else to do
                            Exit For
                        End If
                        
                        On Error Resume Next
                            ' I hate to do this but I don't have time to figure out a better
                            ' way to determine which is going to work...
    
    
                        If oFooter.Shapes(1).TextFrame Is Nothing Then
                            Err.Clear
                            GoTo OldVersion2
                        End If
                        
                        Set oTFrame = oFooter.Shapes(1).TextFrame
    
                        On Error Resume Next
                        If oTFrame.TextRange Is Nothing Then
                            If Err.Number <> 0 Then
                                Err.Clear
                                GoTo OldVersion2
                            End If
                        End If
                        If Err.Number <> 0 Then
                            Err.Clear
                            GoTo OldVersion2
                        End If
    
    
    
                        If oTFrame.TextRange.Fields Is Nothing Then
                            If Err.Number <> 0 Then
                                Err.Clear
                                GoTo OldVersion2
                            End If
                        End If
                        
                        On Error GoTo Block_Err
                        
                        Set oRange = objWordDoc.Bookmarks(sBookmarkName).Range
                        
                        ''Set oField = oFooter.Shapes(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
                        
    '                    Set oField = oFooter.Shapes(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldEmpty, "", True)
    
                        Set oField = oFooter.Shapes(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldIncludePicture, sQRCodePath, True)
                        oField.Unlink
                        
    '                    oField.Result.Text = "*" & sInstanceID & "*"
                        ' now lets remove the bookmark all together - I've been finding examples where this is replaced several times
                        ' resulting in barcodes like 01030303
                        ' when it should be 0103
    '                    oField.Unlink
                        If IsBookMark(objWordDoc, sBookmarkName) = True Then
                            objWordDoc.Bookmarks(sBookmarkName).Delete
                        End If
                        
                    ElseIf oFooter.Range.ShapeRange.Count > 0 Then
OldVersion2:
                        If IsBookMark(objWordDoc, sBookmarkName) = False Then
                            ' nothing else to do
                            Exit For
                        End If
    '                    oField.Unlink
                        
                        Set oRange = objWordDoc.Bookmarks(sBookmarkName).Range
    '                    Set oField = oFooter.Range.ShapeRange.Item(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
    
    '                    Set oField = oFooter.Range.ShapeRange.Item(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldEmpty, "", True)
                        Set oField = oFooter.Range.ShapeRange.Item(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldIncludePicture, sQRCodePath, True)
    '                    oField.Result.Text = "*" & sInstanceId & "*"
                        oField.Unlink
                        
                        
                        If IsBookMark(objWordDoc, sBookmarkName) = True Then
                            objWordDoc.Bookmarks(sBookmarkName).Delete
                        End If
                        
                    End If
                    
                Next
                
            Next
        End If
NextBkmark:
    Next
    AddInstanceIdQRCode = True  ' In this case true = no error.. :)
    
Block_Exit:
    ' Make sure we set this back the way it was..
    Application.SetOption "Error Trapping", iCurrentSetting
    Set oField = Nothing
    Set oRange = Nothing
    Set oFooter = Nothing
    Set objWordSection = Nothing
    LogMessage strProcName, "EFFICIENCY TESTING", "Ending", CStr("" & ProcessTookHowLong(dtStart))
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




Public Function UnlinkWordFields(oWordApp As Word.Application, oDoc As Word.Document, Optional sLetterType As String = "") As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim objWordField As Word.Field
Dim objWordSection As Word.Section
Dim i As Integer
Dim dtStart As Date
Dim lSections As Long

    strProcName = ClassName & ".UnlinkWordFields"
    dtStart = Now()
    ' 20130219 KD: Make sure that the section pages start at 1
      
    oDoc.Repaginate
    SleepEvents 1
    DoEvents
    DoEvents
    DoEvents
    
'''    With oDoc
'''        lSections = .Sections.Count
'''        LogMessage strProcName, "EFFICIENCY TESTING", "Starting for " & sLetterType, "Sections: " & CStr(lSections)
'''        For i = 1 To .Sections.Count
'''            .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
'''            .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
'''            .Repaginate
'''            .Sections(i).Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
'''            .Sections(i).Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
'''            .Repaginate
'''
'''            '' how about shapes?
'''            Dim oShape As Word.Shape
'''            'For Each oShape In .Sections(i).Footers(wdHeaderFooterPrimary).Shapes
'''             '   oShape.TextFrame.TextRange.Fields.Unlink
'''
'''            'Next
'''
''''            For Each oShape In .Sections(i).headers(wdHeaderFooterPrimary).Shapes
''''                oShape.TextFrame.TextRange.Fields.Unlink
''''            Next
'''
'''        Next i
'''        '.Fields.Unlink
'''    End With
      
    oDoc.Activate
      
        '' Hardcoded (shame) for QR barcodes: need to make this data driven at some point..
'    If sLetterType <> "VADRA_QR" And sLetterType <> "VADRA" Then
'    ' by the way, this breaks the ADR footer's Page X of Y (even though it doesn't break the
''        For Each objWordSection In oWordApp.ActiveDocument.Sections
''            For i = 1 To objWordSection.Footers.Count
''                For Each objWordField In objWordSection.Footers.Item(i).Range.Fields
''                    Debug.Print objWordField.Code
''                    objWordField.Update
''                    objWordField.Unlink
''                Next
''
''            Next
''        Next
'Stop
'
'    Else
'    If left(sLetterType, 4) = "VADR" Then
'Stop
        '' this should be unlinking the bar codes
        For Each objWordSection In oWordApp.ActiveDocument.Sections
            For Each objWordField In objWordSection.Range.Fields
                objWordField.Update
                  objWordField.Unlink
            Next
        Next
'    End If
        
    Dim oRng As Word.Range, hLink As Word.Hyperlink
'
'    With oDoc
'        ' Loop through Story Ranges and update.
'        ' Note that this may trigger interactive fields (eg ASK and FILLIN).
'        For Each oRng In .StoryRanges
'            Do
'                oRng.Fields.Unlink
'                For Each hLink In oRng.Hyperlinks
'                    hLink.Delete
'                Next
'                Set oRng = oRng.NextStoryRange
'            Loop Until oRng Is Nothing
'        Next
'    End With
    
   
        
      UnlinkWordFields = True
Block_Exit:
    LogMessage strProcName, "EFFICIENCY TESTING", "Ending for " & sLetterType, CStr("" & ProcessTookHowLong(dtStart))
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function CopyIndividualLtrsToTempFldr(oRs As ADODB.RecordSet, ByVal sPathFieldName As String, ByRef sTempFldr As String, Optional lLtrCount As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRsCln As ADODB.RecordSet
Dim sErrMsg As String
Dim SFileName As String
Dim lErrCnt As Long

    strProcName = ClassName & ".CopyIndividualLtrsToTempFldr"
    lLtrCount = 0
    
    If Not oRs Is Nothing Then
        If isField(oRs, sPathFieldName) = False Then
            Stop
        End If
    End If
    

'    If sTempFldr = "" Then
        sTempFldr = GetUserTempDirectory()
        If sTempFldr = "" Then
            LogMessage strProcName, "ERROR", "There was an error creating a work folder!", , True
            GoTo Block_Exit
        End If
'    End If
    

    ' start out optimistically:
    CopyIndividualLtrsToTempFldr = True
    
    Set oRsCln = oRs.Clone
    
    oRsCln.MoveFirst
    While Not oRsCln.EOF
        SFileName = GetFileName(oRsCln(sPathFieldName).Value)

            ' if it already exists then remove it and copy it again to be sure we are using a good copy
        If FileExists(sTempFldr & SFileName) = True Then
            ' what do we do now? I guess it's just a copy so we should probably delete, then copy it again
            ' assuming that the source is still there
            If FileExists(oRsCln(sPathFieldName).Value) = False Then
                LogMessage strProcName, "ERROR", "The source file to load for InstanceID: " & oRs("InstanceId").Value & " is missing", oRsCln(sPathFieldName).Value
'                Call ErrorCallStack_Add(0, "Source file to load for Instance id: " & oRS("InstanceId").Value & " is missing", strProcName, oRsCln(sPathFieldName).Value, , , oRS("InstanceID").Value, oRS("LetterType").Value)
                lErrCnt = lErrCnt + 1
                CopyIndividualLtrsToTempFldr = False
                GoTo CopyNext
            Else
                If DeleteFile(sTempFldr & SFileName, False) = False Then
                    LogMessage strProcName, "WARNING", "The file being copied already exists and appears to be locked open", sTempFldr & SFileName
'                    Call ErrorCallStack_Add(0, "The file being copied already exists and appears to be locked open.", strProcName, sTempFldr & sFileName, , , oRS("InstanceID").Value, oRS("LetterType").Value)
                    lErrCnt = lErrCnt + 1
                    CopyIndividualLtrsToTempFldr = False
                    GoTo CopyNext
                End If
            End If
        End If
        If CopyFile(oRsCln(sPathFieldName).Value, sTempFldr, False, sErrMsg) = False Then
            Stop
            GoTo CopyNext
        End If
        
        lLtrCount = lLtrCount + 1
CopyNext:
        oRsCln.MoveNext
    Wend
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function GetLastLoadQueueRunTime(Optional lLastQueueRunId As Long) As Date
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".GetLastLoadQueueRunTime"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_LastRunStartTime"
        .Parameters.Refresh
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
        GetLastLoadQueueRunTime = Nz(.Parameters("@pLastRunStartTime").Value, CDate("1/1/1900"))
        lLastQueueRunId = Nz(.Parameters("@pLastQueueRunId").Value, 0)

    End With
    
        
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function GetProcessorState(Optional lThisQueueRunId As Long) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".GetProcessorState"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_GetProcessorState"
        .Parameters.Refresh
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
        GetProcessorState = .Parameters("@pProcessorState").Value
        lThisQueueRunId = .Parameters("@pLastQueueRunId")
    End With
    
        
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function UpdateProcessorStatus(sCurrentPhase As String, lQueueRunId As Long, Optional lPctDone As Long = 0, Optional lBatchId As Long = 0, Optional sInstanceId As String, _
            Optional bFinished As Boolean = False, Optional bStarting As Boolean = False, Optional lRecordsAffected As Long = -1) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Static oAdo As clsADO

    strProcName = ClassName & ".UpdateProcessorStatus"
    
    If oAdo Is Nothing Then
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = CodeConnString
            .SQLTextType = StoredProc
            .sqlString = "usp_LETTER_Automation_SetProcessStatus"
        End With
    End If
    
    With oAdo
        .Parameters.Refresh
        .Parameters("@pCurrentPhase") = sCurrentPhase
        .Parameters("@pQueueRunId") = lQueueRunId
        If lPctDone > 0 Then .Parameters("@pPercentDone") = lPctDone
        If lBatchId > 0 Then .Parameters("@pCurrentBatchId") = lBatchId
        If sInstanceId <> "" Then .Parameters("@pCurrentInstanceId") = sInstanceId
        
        .Parameters("@pStarting") = IIf(bStarting, 1, 0)
        .Parameters("@pFinished") = IIf(bFinished, 1, 0)
        If lRecordsAffected > -1 Then
            .Parameters("@pRowsAffected") = lRecordsAffected
        End If
        .Execute
    End With
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



' For the life of me, I don't know who keeps changing my code to late bound
' I mean, it's not difficult to go from one version of word to another and the
' benefits of early bound outweigh any kind of other issues (unless you have
' mixed users of course - we don't here in CMS!!!!)
Public Function IsBookMark(objWordDoc As Object, sBookmarkName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oBkmk As Object

    strProcName = ClassName & ".IsBookMark"
    
    For Each oBkmk In objWordDoc.Bookmarks
        If UCase(oBkmk.Name) = UCase(sBookmarkName) Then
            IsBookMark = True
            GoTo Block_Exit
        End If
    Next
    
Block_Exit:
    Set oBkmk = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function





Public Function InsertWordDocAtEndOfCurrentDoc(oWordDoc As Word.Document, sFileToInsert As String, Optional lTotalPagesAfterMerge As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oWordApp As Word.Application
Dim lTtlPages As Long
'Dim lCurPage As Long

'Const iMethodToTry As Integer = 1

    strProcName = ClassName & ".InsertWordDocAtEndOfCurrentDoc"
    If FileExists(sFileToInsert) = False Then
        LogMessage strProcName, "ERROR", "File to append does not seem to exist where specified!", sFileToInsert
        GoTo Block_Exit
    End If
    
    ' to insert a file we need to get a selection where we want the document to start..
    ' so, loop through the pages
'''    LogMessage strProcName, "EFFICIENCY TESTING", "Starting to insert word doc", "Method: " & CStr(iMethodToTry)
    
    
    Set oWordApp = oWordDoc.Application
''''    Select Case iMethodToTry
''''    Case 1
''''        oWordDoc.Select
''''        lTtlPages = oWordApp.selection.Information(4)   'wdNumberOfPagesInDocument)
''''
''''        oWordApp.selection.Goto 1, 2, lTtlPages
''''        oWordApp.selection.WholeStory
''''        oWordApp.selection.EndKey Unit:=wdStory
''''    Case 2
        oWordDoc.Activate
        lTtlPages = oWordDoc.BuiltInDocumentProperties(14)  ' wdPropertyPages
        
        oWordApp.selection.GoTo 1, 2, lTtlPages
'        oWordApp.selection.WholeStory
        oWordApp.selection.EndKey Unit:=wdStory

''''    End Select
    '' below may be quicker than selecting..
    '' ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)

   
    
    oWordApp.selection.InsertBreak (2) '   wdSectionBreakNextPage    'insert a page break befor the file so next insert does not Bunch together. doing it before so you don't get an empty last page
    
    oWordApp.selection.InsertFile (sFileToInsert)
    
    lTotalPagesAfterMerge = oWordDoc.BuiltInDocumentProperties(14)  ' wdPropertyPages
''''    LogMessage strProcName, "EFFICIENCY TESTING", "Finished inserting word doc", "Method: " & CStr(iMethodToTry) & "," & CStr(lTtlPages)


    InsertWordDocAtEndOfCurrentDoc = True
    
Block_Exit:
    Set oWordApp = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


'Public Function InsertWordDocAtStartOfCurrentDoc(oWordDoc As Word.Document, sFileToInsert As String, Optional lTotalPagesAfterMerge As Long) As Boolean
' total pages after merge isn't set up because that takes time to calculate and that would be inefficient
Public Function InsertWordDocAtStartOfCurrentDoc(oWordDoc As Word.Document, sFileToInsert As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oWordApp As Word.Application
Dim lTtlPages As Long
Dim oWordInsertDoc As Word.Document
'Const iMethodToTry As Integer = 1
Dim objWordField As Word.Field
Dim objWordSection As Word.Section
Dim oShape As Word.Shape
Dim i As Integer

' KD: ok, so to deal with a bug, design flaw, whatever
' we need to first open the file to be inserted, and make sure it starts and finishes with a Continuous section break.

    strProcName = ClassName & ".InsertWordDocAtStartOfCurrentDoc"
    If FileExists(sFileToInsert) = False Then
        LogMessage strProcName, "ERROR", "File to append does not seem to exist where specified!", sFileToInsert
        GoTo Block_Exit
    End If
    
    
    Set oWordApp = oWordDoc.Application

    oWordDoc.Activate

    oWordApp.selection.GoTo 1, 1, 1

    oWordApp.selection.HomeKey Unit:=wdStory
    
    oWordApp.selection.InsertFile (sFileToInsert)
    oWordApp.selection.InsertBreak (2) '   wdSectionBreakNextPage    'insert a page break befor the file so next insert does not Bunch together. doing it before so you don't get an empty last page
    
    
    '2014:04:29:JS: Why do this for each section every time? change to do it for the last one inserted, which would be always section 1
    With oWordDoc
        .Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
        .Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
        '.Repaginate
        .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
        .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
    End With
    
'    lTotalPagesAfterMerge = oWordDoc.BuiltInDocumentProperties(14)  ' wdPropertyPages
''''    LogMessage strProcName, "EFFICIENCY TESTING", "Finished inserting word doc", "Method: " & CStr(iMethodToTry) & "," & CStr(lTtlPages)

    InsertWordDocAtStartOfCurrentDoc = True
    
Block_Exit:
    Set oWordApp = Nothing  ' Do not .Quit here as it'll kill our combined document
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function GetGenerationDetails42Day(oForm As Form_frm_BOLD_Ops_Dashboard) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".GetGenerationDetails42Day"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_GetGenerationDetails42Day"
        .Parameters.Refresh
        .Execute
    
        oForm.txtTtlClaimsGenToday = Nz(.Parameters("@pTtlClaimsGenToday").Value, 0)
        oForm.txtTtlLtrsGenToday = Nz(.Parameters("@pTtlLtrsGenToday").Value, 0)
    
    End With
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




Public Function GetOutputTodayDetails42Day(oForm As Form_frm_BOLD_Ops_Dashboard) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".GetOutputTodayDetails42Day"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_GetOutputTodayDetails42Day"
        .Parameters.Refresh
        .Execute
    
        oForm.txtTtlClaimsCombined2Day = Nz(.Parameters("@pTtlClaimsCombined2Day").Value, 0)
        oForm.txtTtlLtrsCombined2Day = Nz(.Parameters("@pTtlLtrsCombined2Day").Value, 0)
        oForm.txtTtlBatchesCombined2Day = Nz(.Parameters("@pTtlBatchesCombined2Day").Value, 0)
        oForm.txtTtlPagesCombined2Day = Nz(.Parameters("@pTtlPagesCombined2Day").Value, 0)
    
    End With
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function PrepareXmlForMailRoom() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oDetailsRs As ADODB.RecordSet
Dim sOutFolder As String
Dim sXmlFileName As String
Dim oFso As Scripting.FileSystemObject
Dim oTxt As Scripting.TextStream
Dim sXmlSummary As String
Dim sXmlDetails As String


    strProcName = ClassName & ".PrepareXmlForMailRoom"

    Set oFso = New Scripting.FileSystemObject

    sOutFolder = GetSetting("MAILROOM_DATA_PATH")
    If sOutFolder = "" Then
        Stop
    End If
    
    sOutFolder = QualifyFldrPath(sOutFolder)
    
    sOutFolder = sOutFolder & QualifyFldrPath(Format(Now(), "yyyy-mm-dd"))
    
    ' Make sure it's empty:
    Call DeleteFullFolder(sOutFolder)
    
    
    CreateFolders (sOutFolder)
    

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_GetXMLOutputDetails"
        .Parameters.Refresh
        .Parameters("@pCombineQueueRunId") = 35 ''  Me.ThisQueueRunId
        Set oDetailsRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
        sXmlSummary = .Parameters("@pXmlSummaryOut").Value
    End With
    
    
    
'    sXmlFileName = Format(Me.ThisQueueRunId, "0###") & "_Summary.xml"
    sXmlFileName = Format(35, "0###") & "_Summary.xml"
    
    Set oTxt = oFso.CreateTextFile(sOutFolder & sXmlFileName, True, False)
    oTxt.Write sXmlSummary
    oTxt.Close

    sXmlFileName = Replace(sXmlFileName, "Summary", "Details")
    
    Set oTxt = oFso.CreateTextFile(sOutFolder & sXmlFileName, True, True)
    oTxt.Write oDetailsRs(0).Value
    oTxt.Close

    
    
    
Block_Exit:
    If Not oDetailsRs Is Nothing Then
        If oDetailsRs.State = adStateOpen Then oDetailsRs.Close
        Set oDetailsRs = Nothing
    End If
    Set oTxt = Nothing
    Set oAdo = Nothing
    Set oFso = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function CleanupFolders() As Boolean
On Error GoTo Block_Exit
Dim strProcName As String
Dim vKey As Variant

    strProcName = ClassName & ".CleanupFolders"
    
    If cdctFoldersToCleanUp Is Nothing Then
        CleanupFolders = True
        GoTo Block_Exit
    End If
    
    For Each vKey In cdctFoldersToCleanUp
        If FolderExists(CStr(vKey)) = True Then
            DeleteFullFolder (CStr(vKey))
        End If
        cdctFoldersToCleanUp.Remove vKey
    Next
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
 
End Function



Public Function AddFolderToCleanUp(ByVal strPath As String) As Long
    If cdctFoldersToCleanUp Is Nothing Then Set cdctFoldersToCleanUp = New Scripting.Dictionary
    If cdctFoldersToCleanUp.Exists(strPath) = False Then
        cdctFoldersToCleanUp.Add strPath, 1
    End If
    AddFolderToCleanUp = cdctFoldersToCleanUp.Count
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Sub SendErrorEmail(sErrorMessage As String, Optional lJobId As Long = 0, _
    Optional lBatchId As Long = 0)
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".SendNotification"
    

    Set oAdo = New clsADO
    With oAdo
Stop
        '''.ConnectionString = GetConnectString("CONVERT_Jobs")
'        .ConnectionString = GetConverterConnectionString()
        .SQLTextType = StoredProc
        .sqlString = "usp_CONVERTER_SendErrorNotification"
        .Parameters.Refresh
        .Parameters("@pErrMsg") = sErrorMessage
        .Parameters("@pJobId") = lJobId
        .Parameters("@pBatchId") = lBatchId
        .Execute
        
    End With
    

Block_Exit:
    Exit Sub
Block_Err:
'    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function LoadMailRoomData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oSubFldr As Scripting.Folder
Dim oFile As Scripting.file
Dim sStartFldr As String
Dim sPrintFldr As String
Dim lBatchCount As Long
Dim lLetterCount As Long
Dim oCmd As ADODB.Command
Dim oCn As ADODB.Connection

'Stop
'LoadMailRoomData = True
'GoTo Block_Exit

    strProcName = ClassName & ".LoadMailRoomData"
    
    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = CodeConnString
        .Open
    End With
    
    Set oCmd = New ADODB.Command
    With oCmd
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_MAILOPS_CheckDataLoad"
        .ActiveConnection = oCn
        .Parameters.Refresh
    End With
    
    ' normally we'll just do today
    sStartFldr = GetSetting("MAILROOM_DATA_PATH")
    sPrintFldr = GetSetting("MAILROOM_LETTER_PATH")
    
    CreateFolders (sStartFldr)
    CreateFolders (sPrintFldr)
    
    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder(sStartFldr)
    
    If oFldr Is Nothing Then
        Stop
    End If
Dim sDtFolder As String

    sDtFolder = Format(DateAdd("d", -6, Now()), "yyyy-mm-dd")
    
    For Each oSubFldr In oFldr.SubFolders
        If IsDate(oSubFldr.Name) = False Then
            GoTo NxtFldr
        End If
        If CDate(oSubFldr.Name) < CDate(sDtFolder) Then
            GoTo NxtFldr
        End If
        For Each oFile In oSubFldr.Files
            ' If we have loaded this one already then skip it
            oCmd.Parameters("@pFilePath") = oFile.Path
            oCmd.Execute
            If oCmd.Parameters("@RETURN_VALUE").Value > 0 Then
'                Stop
                GoTo NextFile
            Else
'                Stop
            End If
            Debug.Print "File: " & oFile.Name & " = " & oFile.Size
            If LoadMailRoomDataXML(oFile.Path, lBatchCount, lLetterCount) = False Then
                Stop
                LogMessage strProcName, "ERROR", "Bad file - going to next"
            End If
NextFile:
        Next
NxtFldr:
    Next
    
    Debug.Print "Loaded: " & CStr(lBatchCount) & " Batches"
    Debug.Print "Loaded: " & CStr(lLetterCount) & " letters"
    
    LoadMailRoomData = True
Block_Exit:
    Set oCmd = Nothing
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If
    Set oFile = Nothing
    Set oSubFldr = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''= "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\Scan Ops\Automated Letters\DataForDashboard\2014-07-11\0853_Details_20140711_115102.xml"
Public Function LoadMailRoomDataXML(sXDocPath As String, Optional lBatchsLoaded As Long, Optional lLettersLoaded As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oXDoc As MSXML2.DOMDocument30
Dim xRoot As MSXML2.IXMLDOMElement
Dim oNode As MSXML2.IXMLDOMElement
Dim oInstNode As MSXML2.IXMLDOMElement
Dim oXNodeList As MSXML2.IXMLDOMNodeList
Dim oXAtt As MSXML2.IXMLDOMAttribute
Dim oXInstanceNodeList As MSXML2.IXMLDOMNodeList

Dim iCurNodeIdx As Integer
Dim sCombinedFolder As String
Dim oCn As ADODB.Connection
Dim oBatchCmd As ADODB.Command
Dim oInstanceCmd As ADODB.Command
Dim iInstanceNode As Integer
Dim lMailBatchId As Long
Dim lLetterBatchId As Long
Dim lCombinedDocNum As Long
Dim oFinalCmd As ADODB.Command
Dim dctOpsBatchIds As Scripting.Dictionary
Dim vKey As Variant


    strProcName = ClassName & ".LoadMailRoomDataXml"
    Set dctOpsBatchIds = New Scripting.Dictionary
    
    'Set oXDoc = New MSXML2.FreeThreadedDOMDocument
    Set oXDoc = New MSXML2.DOMDocument
    oXDoc.async = False
'    oXDoc.SetProperty "SelectionLanguage", "XPath"
    If oXDoc.Load(sXDocPath) = False Then
        Debug.Print "Bad file: " & sXDocPath
        Stop
        Call DeleteFile(sXDocPath, False)
        GoTo Block_Exit
    End If
    
    ' Set up our cmd to insert
    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = CodeConnString
        .Open
    End With
    
    Set oBatchCmd = New ADODB.Command
    With oBatchCmd
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_MAILOPS_BatchInsert"
        .ActiveConnection = oCn
        .Parameters.Refresh
    End With

    Set oInstanceCmd = New ADODB.Command
    With oInstanceCmd
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_MAILOPS_InstanceInsert"
        .ActiveConnection = oCn
        .Parameters.Refresh
    End With
    
    
    
    Set oXNodeList = oXDoc.getElementsByTagName("CombinedDocNum")
    
    For iCurNodeIdx = 0 To oXNodeList.Length - 1
        Set oNode = oXNodeList.Item(iCurNodeIdx)
        
        oBatchCmd.Parameters("@pAuditNum") = IIf(gintAuditId = 0, 4490, gintAuditId)
        oBatchCmd.Parameters("@pAccountId") = gintAccountID
        oBatchCmd.Parameters("@pBatchDataXmlFilePath") = sXDocPath
'        obatchcmd.Parameters("@pCombinedDocNum") =
        
        lCombinedDocNum = oNode.Attributes(0).NodeValue
        
        For Each oXAtt In oNode.Attributes
            Debug.Print oXAtt.Name & " = " & oXAtt.Value
            If oXAtt.Name = "CombinedFilePath" Then
                sCombinedFolder = QualifyFldrPath(oXAtt.Value)   '   old: ParentFolderPath(oXAtt.Value)
                oBatchCmd.Parameters("@pCombinedFilePath") = sCombinedFolder
            ElseIf oXAtt.Name = "LetterBatchId" Then
                If dctOpsBatchIds.Exists(oXAtt.Value) = False Then
                    dctOpsBatchIds.Add oXAtt.Value, 1
                End If
                If isParameter(oBatchCmd, "@pOps" & oXAtt.Name) Then
                    oBatchCmd.Parameters("@pOps" & oXAtt.Name) = oXAtt.Value
                Else
                    oBatchCmd.Parameters("@p" & oXAtt.Name) = oXAtt.Value
                End If
           
            ElseIf isParameter(oBatchCmd, "@p" & oXAtt.Name) = True Then
                oBatchCmd.Parameters("@p" & oXAtt.Name) = oXAtt.Value
            ElseIf isParameter(oBatchCmd, "@pOps" & oXAtt.Name) Then
                oBatchCmd.Parameters("@pOps" & oXAtt.Name) = oXAtt.Value
            End If

        Next

        
        oBatchCmd.Execute
        If Nz(oBatchCmd.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", oBatchCmd.Parameters("@pErrMsg").Value
            GoTo NextBatch
        End If
        lMailBatchId = Nz(oBatchCmd.Parameters("@pNewId").Value, 0)

        '' Get all of the instances
        Set oXInstanceNodeList = oNode.selectNodes("//Instances/LetterInstance")
        
        ' KD: ok, this is getting ALL of the LetterInstances, not just for the current CombinedDocNum
        ' So, see if (with JQUERY) I can get where CombinedDocNum = our number
        ' but now that I read my code a little further, I see I took care of that
        ' in a different way with VBA...
        For iInstanceNode = 0 To oXInstanceNodeList.Length - 1
            Set oInstNode = oXInstanceNodeList.Item(iInstanceNode)
            oInstanceCmd.Parameters("@pMailBatchId") = lMailBatchId
            
            'If oInstNode.Attributes("CombinedDocNum") = lCombinedDocNum Then
            If oInstNode.Attributes(0).NodeValue = lCombinedDocNum Then
                For Each oXAtt In oInstNode.Attributes
                    Debug.Print oXAtt.Name & " = " & oXAtt.Value
                    If oXAtt.Name = "CombinedFilePath" Then
                        sCombinedFolder = ParentFolderPath(oXAtt.Value)
                        oInstanceCmd.Parameters("@pPrintPath") = oXAtt.Value
                    ElseIf isParameter(oInstanceCmd, "@p" & oXAtt.Name) = True Then
    
                        oInstanceCmd.Parameters("@p" & oXAtt.Name) = oXAtt.Value
                    ElseIf isParameter(oInstanceCmd, "@pOps" & oXAtt.Name) Then
                        oInstanceCmd.Parameters("@pOps" & oXAtt.Name) = oXAtt.Value
                    End If
                Next
                oInstanceCmd.Execute
                If Nz(oInstanceCmd.Parameters("@pErrMsg").Value, "") <> "" Then
                    LogMessage strProcName, "ERROR", oInstanceCmd.Parameters("@pErrMsg").Value
                    GoTo NextInstance
                End If
                lLettersLoaded = lLettersLoaded + 1
            Else
'                Stop
            End If

NextInstance:
        Next
        
        
        Set oFinalCmd = New ADODB.Command
        With oFinalCmd
            .commandType = adCmdStoredProc
            .CommandText = "usp_LETTER_Automation_MAILOPS_FinalizeInsert"
            .ActiveConnection = oCn
            .Parameters.Refresh
            .Parameters("@pMailBatchId") = lMailBatchId
'            .Parameters("@pOpsLetterBatchid") = ""
            .Parameters("@pCombinedFilePath") = sCombinedFolder
            .Parameters("@pAccountId") = gintAccountID      ''
            
            .Execute
            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                LogMessage strProcName, "ERROR", .Parameters("@pErrMsg").Value, "In: " & .CommandText
                Stop
                GoTo NextBatch
            End If
        End With

        
        lBatchsLoaded = lBatchsLoaded + 1
NextBatch:
    Next
    
    Set oBatchCmd = New ADODB.Command
    With oBatchCmd
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_MAILOPS_FixBatchTypes"
        .ActiveConnection = oCn
        .Parameters.Refresh
    End With
    
    For Each vKey In dctOpsBatchIds.Keys
        If IsNumeric(vKey) = True Then
            With oBatchCmd
                .Parameters("@pOpsBatchId") = vKey
                .Execute
                If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                    LogMessage strProcName, "ERROR", .Parameters("@pErrMsg").Value
                    GoTo Block_Exit
                End If
            End With
        End If
    Next
    
        '' Now we can update the Batch table with the lettercount
    ' (just do where they are null)

    
    ' usp_LETTER_Automation_MAILOPS_FinalizeInsert
    
    
    
    '' Now we can update the Batch table with the lettercount
    ' (just do where they are null)
    Set oBatchCmd = New ADODB.Command
    With oBatchCmd
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_MAILOPS_BatchLtrCountUpdate"
        .ActiveConnection = oCn
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg").Value
            GoTo Block_Exit
        End If
    End With
    
    LoadMailRoomDataXML = True
Block_Exit:
    Set oXAtt = Nothing
    Set oNode = Nothing
    Set oXNodeList = Nothing
    Set xRoot = Nothing
    Set oNode = Nothing
    Set oXDoc = Nothing
    
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


'' Ok, we are going to have to build in a delay
'' in case we try to send too many jobs to the print queue


'' Ok, we are going to have to build in a delay
'' in case we try to send too many jobs to the print queue

Public Function CreateCoverPageAndPrint(lAccountId As Long, lMailBatchId As Long, sBatchType As String, sFolderPath As String, lTrackRowId As Long, _
    sSelectedPrinter As String, Optional oFrmStatus As Form_ScrStatus) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sTemplatePath As String
Dim objWordApp As Word.Application
Dim objWordDoc As Word.Document
Dim objWordMergedDoc As Word.Document
Dim strODCFile As String
Dim db As DAO.Database
Dim rsLetterConfig As DAO.RecordSet
Dim bMergeError As Boolean
Dim sTempFilename As String
Dim sSql As String
Dim sFilePath As String
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oFile As Scripting.file
Dim lModulous As Long
Dim lSleepSecs As Long
Dim lCurCount As Long
Dim oCmdStart As ADODB.Command
Dim oCmdEnd As ADODB.Command
Dim oCn As ADODB.Connection
'Dim lThisAcct As Long



    strProcName = ClassName & ".CreateCoverPageAndPrint"
            ' This is going to be done with a mail merge (ugh, I know!!!)

    sTemplatePath = GetSetting("MAILROOM_COVERSHEET_TEMPLATE_PATH")
            'sTemplatePath = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\Scan Ops\Automated Letters\kd_cover_change_test.doc"
    If FileExists(sTemplatePath) = False Then
        LogMessage strProcName, "ERROR", "Could not find the cover sheet template where specified", sTemplatePath, True
        Stop
        GoTo Block_Exit
    End If

'    lThisAcct = GetAccountFromOpsBatchId(lMailBatchId)

    ' this was for CMS
'    If gintAccountID = 0 Then gintAccountID = 1
    
    lModulous = CLng("0" & GetSetting("PRINT MODULOUS"))
    lSleepSecs = CLng("0" & GetSetting("PRINT SLEEP SECONDS"))
    
    
    ' Just in case:
    If lModulous < 1 Then
        lModulous = 4
    End If
    If lSleepSecs < 1 Then
        lSleepSecs = 2
    End If
    
    
    
    ' Set the based path for saving merge doc
'    Set db = CurrentDb
'        'TL add account ID logic
'    Set rsLetterConfig = db.OpenRecordSet("SELECT * FROM LETTER_Config WHERE AccountID = " & CStr(lAccountId))
'
'
'    strODCFile = rsLetterConfig("ODC_ConnectionFile").value
    
    strODCFile = GetSetting("ODC_ConnectionFile")
    
''    rsLetterConfig.Close
'    Set rsLetterConfig = Nothing
'    Set db = Nothing
    
    
    Set objWordApp = New Word.Application
    objWordApp.visible = False
    Set objWordDoc = objWordApp.Documents.Add(sTemplatePath, , False)
    
    sSql = "SELECT * FROM (SELECT * FROM v_LETTER_Automation_MAILOPS_CoverLetter WHERE BatchId = " & CStr(lMailBatchId) & " AND RowId = " & CStr(lTrackRowId) & " AND BatchType = '" & sBatchType & "') A "
    objWordDoc.MailMerge.OpenDataSource Name:=strODCFile, SqlStatement:=sSql, Connection:=GetConnectString("v_code_database")
        

    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = GetConnectString("v_code_database")
        .CursorLocation = adUseNone
        .Open
    End With

    Set oCmdStart = New ADODB.Command
    With oCmdStart
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_MAILOPS_StartPrintingInstance"
        .ActiveConnection = oCn
        .Parameters.Refresh
    End With
    
    Set oCmdEnd = New ADODB.Command
    With oCmdEnd
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_MAILOPS_EndPrintingInstance"
        .ActiveConnection = oCn
        .Parameters.Refresh
    End With


    objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
    objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
    objWordDoc.MailMerge.Execute Pause:=False
    If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
        objWordApp.visible = True
        MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
        bMergeError = True
        objWordApp.ActiveDocument.Activate
        Stop
        GoTo Block_Exit
    End If
    ''------------------- here is where we convert to pdf instead of word ----------------''
    ' Save the output doc
    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)

    '' Now I need to append this to the front of the other file
    ' but that means I need to save this one and close it
    ' hang on now.. I just need to print it out, I don't need to append it because then the '
    ' formatting gets all tweaked.
    
'    sTempFilename = GetUniqueFilename(, , "doc")
'    objWordMergedDoc.SaveAs2 sTempFilename
'
    objWordApp.ActivePrinter = sSelectedPrinter
'Stop
'    objWordApp.PrintOut
    objWordMergedDoc.PrintOut Background:=False
    
    
    objWordMergedDoc.Close False
    
    ' Get a recordset
    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_MAILOPS_PrintJob"
        .Parameters.Refresh
        .Parameters("@pMailBatchId") = lMailBatchId
        .Parameters("@pBatchType") = sBatchType
        Set oRs = .ExecuteRS
        
    End With
    
    If Not oFrmStatus Is Nothing Then
        oFrmStatus.ProgMax = oRs.recordCount
        oFrmStatus.StatusMessage "Printing batch " & CStr(lMailBatchId) & " on " & sSelectedPrinter
'        oFrmStatus.StatusCaption.Caption = sSelectedPrinter
    End If
    
    While Not oRs.EOF
        lCurCount = lCurCount + 1

        sFilePath = oRs("PathForPrinting").Value
        
        If Not oFrmStatus Is Nothing Then
            oFrmStatus.ProgVal = lCurCount
'            If oFrmStatus.EvalStatus(2) = True Then
'                strProgressMsg = "Cancel has been selected." ' at " & i & " / " & fmrStatus.ProgMax
'                oFrmStatus.StatusMessage strProgressMsg
'                DoEvents
'                strErrMsg = strProgressMsg
'                GoTo Block_Exit
'            End If
        End If
        
        ' usp_LETTER_Automation_MAILOPS_StartPrintingInstance
        oCmdStart.Parameters.Refresh
        oCmdStart.Parameters("@pMailInstanceId") = oRs("MailInstanceId").Value
        oCmdStart.Parameters("@pPrinterName") = sSelectedPrinter
        oCmdStart.Execute
        If Nz(oCmdStart.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "Problem with " & oCmdStart.CommandText, oCmdStart.Parameters("@pErrMsg").Value
            Stop
        End If
        
        Set objWordMergedDoc = Nothing
        Set objWordMergedDoc = objWordApp.Documents.Open(sFilePath)

        
    'Stop
        objWordApp.ActivePrinter = sSelectedPrinter
    '    objWordApp.PrintOut
        
        objWordMergedDoc.PrintOut Background:=False
        
        While objWordMergedDoc.Application.BackgroundPrintingStatus > 0
            SleepEvents 1
        Wend
        
        If lModulous > 0 Then
            If lCurCount Mod lModulous = 0 Then
                LogMessage strProcName, "EFFICIENCY TESTING", "Sleeping before printing more letters..."
                SleepEvents lSleepSecs
            End If
        End If
        objWordMergedDoc.Close SaveChanges:=False
        
        Set objWordMergedDoc = Nothing
        
        ' usp_LETTER_Automation_MAILOPS_EndPrintingInstance
        oCmdEnd.Parameters.Refresh
        oCmdEnd.Parameters("@pMailInstanceId") = oRs("MailInstanceId").Value
        oCmdEnd.Execute
        If Nz(oCmdEnd.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            LogMessage strProcName, "ERROR", "Problem with " & oCmdEnd.CommandText, oCmdEnd.Parameters("@pErrMsg").Value
        End If

        If Not oFrmStatus Is Nothing Then
            oFrmStatus.ProgVal = lCurCount
'            If oFrmStatus.EvalStatus(2) = True Then
'                strProgressMsg = "Cancel has been selected." ' at " & i & " / " & fmrStatus.ProgMax
'                oFrmStatus.StatusMessage strProgressMsg
'                DoEvents
'                strErrMsg = strProgressMsg
'                GoTo Block_Exit
'            End If
        End If

        
        oRs.MoveNext
    Wend
    
    '' Ok, don't do this - just print both out separatly to the same printer

    
' Stop
    CreateCoverPageAndPrint = True
    
Block_Exit:
    Set oCmdStart = Nothing
    Set oCmdEnd = Nothing
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If

    If Not objWordApp Is Nothing Then
        If objWordApp.Documents.Count > 0 Then
            For Each objWordDoc In objWordApp.Documents
                objWordDoc.Close SaveChanges:=False
            Next
        End If
        objWordApp.Quit
        Set objWordApp = Nothing
    End If
    Exit Function
    
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




'' Ok, we are going to have to build in a delay
'' in case we try to send too many jobs to the print queue

Public Function CreateCoverPageAndPrint_LEGACY(lMailBatchId As Long, sBatchType As String, sFolderPath As String, lTrackRowId As Long, sSelectedPrinter As String, Optional oFrmStatus As Form_ScrStatus) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sTemplatePath As String
Dim objWordApp As Word.Application
Dim objWordDoc As Word.Document
Dim objWordMergedDoc As Word.Document
Dim strODCFile As String
Dim db As DAO.Database
Dim rsLetterConfig As DAO.RecordSet
Dim bMergeError As Boolean
Dim sTempFilename As String
Dim sSql As String
Dim sFilePath As String
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oFile As Scripting.file
Dim lModulous As Long
Dim lSleepSecs As Long
Dim lCurCount As Long
Dim oCmdStart As ADODB.Command
Dim oCmdEnd As ADODB.Command
Dim oCn As ADODB.Connection


    strProcName = ClassName & ".CreateCoverPageAndPrint"
            ' This is going to be done with a mail merge (ugh, I know!!!)

    sTemplatePath = GetSetting("MAILROOM_COVERSHEET_TEMPLATE_PATH")
            'sTemplatePath = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\Scan Ops\Automated Letters\kd_cover_change_test.doc"
    If FileExists(sTemplatePath) = False Then
        LogMessage strProcName, "ERROR", "Could not find the cover sheet template where specified", sTemplatePath, True
        GoTo Block_Exit
    End If

    If gintAccountID = 0 Then gintAccountID = 1
    
'    If oSetting Is Nothing Then Set oSetting = New clsSettings
    lModulous = CLng("0" & GetSetting("PRINT MODULOUS"))
    lSleepSecs = CLng("0" & GetSetting("PRINT SLEEP SECONDS"))
    
    
    ' Just in case:
    If lModulous < 1 Then
        lModulous = 4
    End If
    If lSleepSecs < 1 Then
        lSleepSecs = 2
    End If
    
    ' Set the based path for saving merge doc
    Set db = CurrentDb
        'TL add account ID logic
    Set rsLetterConfig = db.OpenRecordSet("select * from LETTER_Config WHERE (AccountID = " & CStr(gintAccountID) & " or " & CStr(gintAccountID) & " = 0)")

    strODCFile = rsLetterConfig("ODC_ConnectionFile").Value
    
    rsLetterConfig.Close
    Set rsLetterConfig = Nothing
    Set db = Nothing
    
    
    Set objWordApp = New Word.Application
    objWordApp.visible = False
    Set objWordDoc = objWordApp.Documents.Add(sTemplatePath, , False)
    
    sSql = "SELECT * FROM (SELECT * FROM v_LETTER_Automation_MAILOPS_CoverLetter WHERE BatchId = " & CStr(lMailBatchId) & " AND RowId = " & CStr(lTrackRowId) & " AND BatchType = '" & sBatchType & "') A "
    objWordDoc.MailMerge.OpenDataSource Name:=strODCFile, SqlStatement:=sSql, Connection:=GetConnectString("v_code_database")
        

    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = GetConnectString("v_code_database")
        .CursorLocation = adUseNone
        .Open
    End With

    Set oCmdStart = New ADODB.Command
    With oCmdStart
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_MAILOPS_StartPrintingInstance"
        .ActiveConnection = oCn
        .Parameters.Refresh
    End With
    
    Set oCmdEnd = New ADODB.Command
    With oCmdEnd
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_MAILOPS_EndPrintingInstance"
        .ActiveConnection = oCn
        .Parameters.Refresh
    End With


    objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
    objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
    objWordDoc.MailMerge.Execute Pause:=False
    If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
        objWordApp.visible = True
        MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
        bMergeError = True
        objWordApp.ActiveDocument.Activate
        Stop
        GoTo Block_Exit
    End If
    ''------------------- here is where we convert to pdf instead of word ----------------''
    ' Save the output doc
    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)

    '' Now I need to append this to the front of the other file
    ' but that means I need to save this one and close it
    ' hang on now.. I just need to print it out, I don't need to append it because then the '
    ' formatting gets all tweaked.
    
'    sTempFilename = GetUniqueFilename(, , "doc")
'    objWordMergedDoc.SaveAs2 sTempFilename
'
    objWordApp.ActivePrinter = sSelectedPrinter
'Stop
'    objWordApp.PrintOut
    objWordMergedDoc.PrintOut Background:=False
    
    
    objWordMergedDoc.Close False
    
    ' Get a recordset
    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_MAILOPS_PrintJob"
        .Parameters.Refresh
        .Parameters("@pMailBatchId") = lMailBatchId
        .Parameters("@pBatchType") = sBatchType
        Set oRs = .ExecuteRS
    End With
    
    If Not oFrmStatus Is Nothing Then
        oFrmStatus.ProgMax = oRs.recordCount
        oFrmStatus.StatusMessage "Printing batch " & CStr(lMailBatchId) & " on " & sSelectedPrinter
'        oFrmStatus.StatusCaption.Caption = sSelectedPrinter
    End If
    
    While Not oRs.EOF
        lCurCount = lCurCount + 1

        sFilePath = oRs("PathForPrinting").Value
        
        If Not oFrmStatus Is Nothing Then
            oFrmStatus.ProgVal = lCurCount
'            If oFrmStatus.EvalStatus(2) = True Then
'                strProgressMsg = "Cancel has been selected." ' at " & i & " / " & fmrStatus.ProgMax
'                oFrmStatus.StatusMessage strProgressMsg
'                DoEvents
'                strErrMsg = strProgressMsg
'                GoTo Block_Exit
'            End If
        End If
        
        ' usp_LETTER_Automation_MAILOPS_StartPrintingInstance
        oCmdStart.Parameters.Refresh
        oCmdStart.Parameters("@pMailInstanceId") = oRs("MailInstanceId").Value
        oCmdStart.Parameters("@pPrinterName") = sSelectedPrinter
        oCmdStart.Execute
        If Nz(oCmdStart.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
        
        Set objWordMergedDoc = Nothing
        Set objWordMergedDoc = objWordApp.Documents.Open(sFilePath)

        
    'Stop
        objWordApp.ActivePrinter = sSelectedPrinter
    '    objWordApp.PrintOut
        
        objWordMergedDoc.PrintOut Background:=False
        
        While objWordMergedDoc.Application.BackgroundPrintingStatus > 0
            SleepEvents 1
        Wend
        
        If lModulous > 0 Then
            If lCurCount Mod lModulous = 0 Then
                LogMessage strProcName, "EFFICIENCY TESTING", "Sleeping before printing more letters..."
                SleepEvents lSleepSecs
            End If
        End If
        objWordMergedDoc.Close SaveChanges:=False
        
        Set objWordMergedDoc = Nothing
        
        ' usp_LETTER_Automation_MAILOPS_EndPrintingInstance
        oCmdEnd.Parameters.Refresh
        oCmdEnd.Parameters("@pMailInstanceId") = oRs("MailInstanceId").Value
        oCmdEnd.Execute
        If Nz(oCmdEnd.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If

        If Not oFrmStatus Is Nothing Then
            oFrmStatus.ProgVal = lCurCount
'            If oFrmStatus.EvalStatus(2) = True Then
'                strProgressMsg = "Cancel has been selected." ' at " & i & " / " & fmrStatus.ProgMax
'                oFrmStatus.StatusMessage strProgressMsg
'                DoEvents
'                strErrMsg = strProgressMsg
'                GoTo Block_Exit
'            End If
        End If

        
        oRs.MoveNext
    Wend
    
    '' Ok, don't do this - just print both out separatly to the same printer

    
' Stop
    CreateCoverPageAndPrint_LEGACY = True
    
Block_Exit:
    Set oCmdStart = Nothing
    Set oCmdEnd = Nothing
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If

    If Not objWordApp Is Nothing Then
        If objWordApp.Documents.Count > 0 Then
            For Each objWordDoc In objWordApp.Documents
                objWordDoc.Close SaveChanges:=False
            Next
        End If
        objWordApp.Quit
        Set objWordApp = Nothing
    End If
    Exit Function
    
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function





'' Ok, we are going to have to build in a delay
'' in case we try to send too many jobs to the print queue
Public Function CreatecoverPageAndPrint_4Legacy(sLetterType As String, sLtrReqDt As String, sFldr As String, sSelectedPrinter As String, oForm As Form_frm_LETTER_Legacy_Print_Tool, _
        lNumPrinted As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sTemplatePath As String
Dim objWordApp As Word.Application
Dim objWordDoc As Word.Document
Dim objWordMergedDoc As Word.Document
Dim strODCFile As String
Dim db As DAO.Database
Dim rsLetterConfig As DAO.RecordSet
Dim bMergeError As Boolean
Dim sTempFilename As String
Dim sSql As String
Dim sFilePath As String
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oFile As Scripting.file
Dim lModulous As Long
Dim lSleepSecs As Long
Dim lCurCount As Long
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sInstanceId As String
Dim lFileCnt As Long
Dim dctFilesToMove As Scripting.Dictionary
Dim vKey As Variant
Dim sDoneFldr As String
Dim lProcessedCnt As Long
Dim sDocNum As String
Dim lDocNum As Long
Dim lIdToUpdate As Long

    strProcName = ClassName & ".CreateCoverPageAndPrint_4Legacy"
    ' This is going to be done with a mail merge (ugh, I know!!!)

    sTemplatePath = GetSetting("MAILROOM_COVERSHEET_LEGACY_TEMPLATE_PATH")
    'sTemplatePath = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\Scan Ops\Automated Letters\kd_cover_change_test.doc"
    If FileExists(sTemplatePath) = False Then
        LogMessage strProcName, "ERROR", "Could not find the cover sheet template where specified", sTemplatePath, True
        GoTo Block_Exit
    End If

    If gintAccountID = 0 Then gintAccountID = 1
    
'    If oSetting Is Nothing Then Set oSetting = New clsSettings
    lModulous = CLng("0" & GetSetting("PRINT MODULOUS"))
    lSleepSecs = CLng("0" & GetSetting("PRINT SLEEP SECONDS"))
'    lMiliSleepSecs = lMiliSleepSecs * 1000
    
    sDoneFldr = QualifyFldrPath(sFldr) & "\done"
    If CreateFolders(sDoneFldr) = False Then
        Stop
        GoTo Block_Exit
    End If
    
    
    ' Set the based path for saving merge doc
    Set db = CurrentDb
    'TL add account ID logic
    Set rsLetterConfig = db.OpenRecordSet("select * from LETTER_Config WHERE (AccountID = " & CStr(gintAccountID) & " or " & CStr(gintAccountID) & " = 0)")

    strODCFile = rsLetterConfig("ODC_ConnectionFile").Value
    
    rsLetterConfig.Close
    Set rsLetterConfig = Nothing
    Set db = Nothing
    
        ' First create the coversheet..
    Set objWordApp = New Word.Application
    objWordApp.visible = False
    Set objWordDoc = objWordApp.Documents.Add(sTemplatePath, , False)
    
    sSql = "SELECT * FROM (SELECT * FROM v_LETTER_Automation_MAILOPS_CoverLetter_Legacy WHERE LetterType = '" & sLetterType & "' AND FileName = '" & QualifyFldrPath(sFldr) & "') A "
    objWordDoc.MailMerge.OpenDataSource Name:=strODCFile, SqlStatement:=sSql, Connection:=GetConnectString("v_code_database")

    objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
    objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
    objWordDoc.MailMerge.Execute Pause:=False
    If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
        objWordApp.visible = True
        MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
        bMergeError = True
        objWordApp.ActiveDocument.Activate
        Stop
        GoTo Block_Exit
    End If
    ''------------------- here is where we convert to pdf instead of word ----------------''
    ' Save the output doc
    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)

    objWordApp.ActivePrinter = sSelectedPrinter
    objWordMergedDoc.PrintOut Background:=False
    objWordMergedDoc.Close False
    
    ' we don't have a recordset in the legacy tool -
    ' we loop over a folder
    
    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder(sFldr)
     
    Set dctFilesToMove = New Scripting.Dictionary
    oForm.MaxToProcess = oFldr.Files.Count
    
    For Each oFile In oFldr.Files

        lProcessedCnt = lProcessedCnt + 1
        DoEvents
        If Not oForm Is Nothing Then
            oForm.CurrentLetterNum = lProcessedCnt

            DoEvents
            If oForm.StopNow = True Then
                GoTo Block_Exit
            End If
            oForm.UpdateStatus
            While oForm.Paused = True
                SleepEvents 2
'                Stop
            Wend
            
        End If
     
        sDocNum = left(oFile.Name, 4)
        lDocNum = CLng("0" & sDocNum)

        sInstanceId = Replace(oFile.Path, sFldr & "\", "", , , vbTextCompare)
        sInstanceId = Replace(sInstanceId, sDocNum & "_", "", , , vbTextCompare)

        lCurCount = lCurCount + 1
        If lModulous > 0 Then
            If lCurCount Mod lModulous = 0 Then
                LogMessage strProcName, "EFFICIENCY TESTING", "Sleeping before printing more letters..."
                SleepEvents lSleepSecs
            End If
        End If
        
        sFilePath = oFile.Path
        
        Call logPrintJob(lDocNum, sInstanceId, sSelectedPrinter, sLetterType, sFldr, "S", lIdToUpdate)
        
        Set objWordMergedDoc = Nothing
        Set objWordMergedDoc = objWordApp.Documents.Open(sFilePath)
    
        objWordApp.ActivePrinter = sSelectedPrinter
        
        objWordMergedDoc.PrintOut Background:=False
        objWordMergedDoc.Close SaveChanges:=False
        
        Set objWordMergedDoc = Nothing
        
        dctFilesToMove.Add sFilePath, 1
        
        Call logPrintJob(lDocNum, sInstanceId, sSelectedPrinter, sLetterType, sFldr, "F", lIdToUpdate)
        
        lIdToUpdate = 0 ' Reset...
    Next

    lNumPrinted = lCurCount

    sDoneFldr = QualifyFldrPath(sFldr) & "done\"
    Set oFile = Nothing
    Set oFldr = Nothing
    
    For Each vKey In dctFilesToMove
        oFso.MoveFile CStr(vKey), sDoneFldr
    Next
    
    CreatecoverPageAndPrint_4Legacy = True
    
Block_Exit:
    If Not objWordApp Is Nothing Then
        If objWordApp.Documents.Count > 0 Then
            For Each objWordDoc In objWordApp.Documents
                objWordDoc.Close SaveChanges:=False
            Next
        End If
        objWordApp.Quit
        Set objWordApp = Nothing
    End If
    Exit Function
    
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function logPrintJob(lDocNum As Long, sInstanceId As String, sPrinter As String, sLetterType As String, sFullPath As String, ByVal sStartOrEnd, ByRef lIdForUpdate As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Static oAdo As clsADO

    strProcName = ClassName & ".logPrintJob"
    
    If Not oAdo Is Nothing Then
        If oAdo.CurrentConnection.State <> adStateOpen Then
            Stop
            Set oAdo = Nothing
        End If
    End If

    If oAdo Is Nothing Then
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("v_Code_Database")
            .SQLTextType = StoredProc
            .sqlString = "usp_Letter_LogPrintJob_Legacy"
            .Parameters.Refresh
        End With
    End If
    
    
    
    With oAdo
        .Parameters.Refresh
        .Parameters("@pInstanceId") = sInstanceId
        
        .Parameters("@pDocNum") = lDocNum
        .Parameters("@pPrinterName") = sPrinter
        .Parameters("@pLetterType") = sLetterType
        
        .Parameters("@pFolderPath") = sFullPath
        
        .Parameters("@pStartOrEnd") = sStartOrEnd
        If sStartOrEnd <> "S" Then
            .Parameters("@pAutoIdForUpdate") = lIdForUpdate
        End If
        
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
        If sStartOrEnd = "S" Then
            lIdForUpdate = .Parameters("@pAutoId").Value
        End If
    End With
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

'' Returns shortname of audit
Public Function GetAuditNameFromId(lAuditId As Long, Optional sFullName As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim oRs As ADODB.RecordSet
Dim oCmd As ADODB.Command

    strProcName = ClassName & ".GetAuditNameFromId"
    
    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" & _
                "Data Source=data.sql.ccaintranet.com;Initial Catalog=ARMS;"
        .CursorLocation = adUseNone
        .Open
    End With
    
    Set oCmd = New ADODB.Command
    With oCmd
        .commandType = adCmdText
        .CommandText = "SELECT AuditId, AuditDesc, ShortName, RptShortName FROM ARMSUser.vRequestsCompanyAudits WHERE AuditId = " & CStr(lAuditId)
        .ActiveConnection = oCn
        Set oRs = .Execute
    End With
    
    sFullName = oRs("AuditDesc").Value
    GetAuditNameFromId = Nz(oRs("ShortName").Value, oRs("RptShortName").Value)
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oCmd = Nothing
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If
    Exit Function
    
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function ClaimAsManualOverrides(oLetterToGenerate As clsLetterInstance) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".ClaimAsManualOverrides"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_ClaimAsManualOverride"
        .Parameters.Refresh
        .Parameters("@pAccountId") = gintAccountID
        .Parameters("@pDynamicInstanceId") = oLetterToGenerate.InstanceId
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was a problem converting this to manual override: '" & oLetterToGenerate.InstanceId & "'", .Parameters("@pErrMsg").Value
            GoTo Block_Exit
        End If
    End With
    ClaimAsManualOverrides = True

Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function MoveLettersToPrintFolder_for_Print_Room() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Dim lDocCount As Long
Dim sFromPath As String
Dim sToPath As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim sNewFileName As String
Dim sCurFileName As String



    strProcName = ClassName & ".MoveLettersToPrintFolder_for_Print_Room"
    
    sToPath = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\Scan Ops\Automated Letters\ToPrint\2015-12-30\1272\"
    
'    sSql = "SELECT DISTINCT X.LetterName , SDD.PageCount  From LETTER_Work_Queue Q INNER Join LETTER_Xref X ON Q.InstanceID = X.InstanceID " & _
'        " INNER Join ( SELECT SD.InstanceId, MIN(PageCount) As PageCount     From     LETTER_Static_Details Sd     GROUP BY Sd.InstanceId " & _
'        " ) AS SDD ON Q.InstanceID = SDD.InstanceId WHERE Q.RowCreateDt > '10/16/2014' AND Q.LetterType = 'VADRA' AND Q.Status = 'P' ORDER BY SDD.PageCount DESC "
            
            
    sSql = "SELECT Distinct X.LetterName, SDD.PageCount  FROM LETTER_Work_Queue Q INNER Join " & _
        " LETTER_Xref X ON Q.InstanceID = X.InstanceID INNER Join (     SELECT SD.InstanceId, MIN(PageCount) As PageCount     From     LETTER_Static_Details Sd     GROUP BY Sd.InstanceId " & _
        " ) AS SDD ON Q.InstanceID = SDD.InstanceId WHERE Q.RowCreateDt > '10/31/2014' AND Q.LetterType = 'VADRA' ORDER BY SDD.PageCount "
        
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
    End With
    
    While Not oRs.EOF
        lDocCount = lDocCount + 1
'        sNewFileName = "VADRA_" & Format(lDocCount, "0000") & ".doc"
        sNewFileName = QualifyFldrPath(ParentFolderPath(oRs("LetterName").Value))
        
        sNewFileName = Replace(oRs("LetterName").Value, sNewFileName, "")
        sNewFileName = Format(lDocCount, "0000") & "_" & sNewFileName
'Stop
        
        If FileExists(oRs("LetterName").Value) = False Then
            Stop
        End If
        If CopyFile(oRs("LetterName").Value, sToPath & sNewFileName, False) = False Then
            Stop
        End If
        oRs.MoveNext
    Wend
    
    MoveLettersToPrintFolder_for_Print_Room = True
    
Block_Exit:
    Stop
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function ADDWATERMARK(objWordApp As Word.Application, objWordDoc As Word.Document, ByRef strErrMsg As String) As Boolean
'Private Function ADDWATERMARK(objWordApp As Word.Application, objWordDoc As Word.Document, ByRef strErrMsg As String) As Boolean
On Error GoTo Error_Encountered

    objWordDoc.Select
    'take our open word document. add a watermark since we are previewing.  This will deter the auditors from printing this.  (ALTHOUGH they could remove it manually)
    objWordApp.ActiveDocument.Sections(1).Range.Select
    objWordApp.ActiveWindow.ActivePane.View.seekview = 9 'wdSeekCurrentPageHeader
    'objWordApp.activewindow.activepane.View.seekview = 9 'wdseekcurrentPageHeader
    
    objWordApp.selection.HeaderFooter.Shapes.AddTextEffect(vbNull, _
        "CONNOLLY - INTERNAL", "Times New Roman", 1, False, False, 0, 0).Select
    objWordApp.selection.ShapeRange.Name = "PowerPlusWaterMarkObject1"
    objWordApp.selection.ShapeRange.TextEffect.NormalizedHeight = False
    objWordApp.selection.ShapeRange.Line.visible = False
    objWordApp.selection.ShapeRange.Fill.visible = True
    objWordApp.selection.ShapeRange.Fill.Solid
    objWordApp.selection.ShapeRange.Fill.ForeColor.RGB = RGB(153, 153, 153)
    objWordApp.selection.ShapeRange.Fill.Transparency = 0
    objWordApp.selection.ShapeRange.Rotation = 315
    objWordApp.selection.ShapeRange.LockAspectRatio = True
    objWordApp.selection.ShapeRange.top = -999995  'wdShapeCenter
    objWordApp.selection.ShapeRange.left = -999995 'wdShapeCenter
    objWordApp.selection.ShapeRange.Height = objWordApp.inchestopoints(2) '1.69)
    objWordApp.selection.ShapeRange.Width = objWordApp.inchestopoints(6.77) '500 '487.45 'InchesToPoints(6.77)
    objWordApp.selection.ShapeRange.WrapFormat.AllowOverlap = True
    objWordApp.selection.ShapeRange.WrapFormat.Side = 3 'wdWrapNone
    objWordApp.selection.ShapeRange.WrapFormat.Type = 3
    objWordApp.selection.ShapeRange.RelativeHorizontalPosition = 0 ' wdRelativeVerticalPositionMargin
    objWordApp.selection.ShapeRange.RelativeVerticalPosition = 0 '  wdRelativeVerticalPositionMargin
    objWordApp.ActiveWindow.ActivePane.View.seekview = 0 'wdSeekMainDocument
    
    ADDWATERMARK = True
    Exit Function
                    
Error_Encountered:
    strErrMsg = Nz(strErrMsg, "") & "Error adding watermark.  " & Err.Source & " " & Err.Number & " " & Err.Description
    ADDWATERMARK = False
    
End Function

Public Function ShowAssociatedClaims() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFrm As Form_frm_GENERAL_Datasheet_ADO
Dim oDBFrm As Form_frm_BOLD_Ops_Dashboard
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sInstanceId As String
Dim oLI As ListItem
Dim oLV As ListView


    strProcName = ClassName & ".ShowAssociatedClaims"
    ' going to open a form (general datasheet I suppose)
    ' which should allow us to get the claims associated with this instance id
    

    If IsOpen("frm_BOLD_Ops_Dashboard", acForm) = False Then
            Stop
    End If

    Set oDBFrm = Application.Forms("frm_BOLD_Ops_Dashboard")
    ' first, which page is selected
    Debug.Print oDBFrm.tabDisplay.Value

    Select Case oDBFrm.tabDisplay.Value
    Case 0  ' queue tab
        Set oLV = oDBFrm.lvQueue
    Case 1  ' Generate page
        Set oLV = oDBFrm.lvQueue
    Case 2  ' output page
        Set oLV = oDBFrm.lvQueue
    End Select

    Set oLI = oDBFrm.RightClickedListItem


    Debug.Print oLI.Text
    
    Set oFrm = New Form_frm_GENERAL_Datasheet_ADO
    
    With oFrm
        .AllowFormView = True
        .ViewsAllowed = 0
        .DefaultView = 1
        .MinMaxButtons = 3
        .BorderStyle = 2
    End With
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = ""
        .Parameters.Refresh
        .Parameters("@pAccountId") = gintAccountID
        .Parameters("@pInstanceId") = sInstanceId
    End With
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function CombineDocsFromRs(oRs As ADODB.RecordSet, sFilePathFieldName As String, sOutFilePathAndName As String, _
            Optional bOpenAfterFinished As Boolean = False, _
            Optional sErrMsg As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Dim oWordApp As Word.Application
Dim oCombinedDoc As Word.Document
Dim oWordDoc As Word.Document
Dim bFirst As Boolean
Dim sThisFilePath As String
Dim lDocCount As Long
Dim lErrCnt As Long


    strProcName = ClassName & ".CombineDocsFromRs"
    
    If oRs Is Nothing Then
        LogMessage strProcName, "ERROR", "Invalid recordset passed to function", , True
        GoTo Block_Exit
    End If
    If oRs.EOF And oRs.BOF Then
        LogMessage strProcName, "ERROR", "Empty recordset passed to function", , True
        GoTo Block_Exit
    End If
    If isField(oRs, sFilePathFieldName) = False Then
        LogMessage strProcName, "ERROR", "File Path Fieldname not found in recordset", sFilePathFieldName, True
        GoTo Block_Exit
    End If
    
    Set oFso = New Scripting.FileSystemObject
    
    If oFso.FileExists(sOutFilePathAndName) = True Then
        LogMessage strProcName, "ERROR", "File to create already exists!", sOutFilePathAndName, True
        GoTo Block_Exit
    End If
    
    Set oWordApp = New Word.Application
    
    bFirst = True
    While Not oRs.EOF
        
        sThisFilePath = oRs(sFilePathFieldName)
        If oFso.FileExists(sThisFilePath) = False Then
            LogMessage strProcName, "ERROR", "The file to combine does not exist where specified", sThisFilePath
            sErrMsg = sErrMsg & sThisFilePath & " not found" & vbCrLf
            GoTo NextFile
        End If
        
        If bFirst = True Then
            If Not oCombinedDoc Is Nothing Then
                Stop
            End If
            
                ' Open the doc...
            Set oCombinedDoc = oWordApp.Documents.Open(sThisFilePath, False, True, False, , , , , , , , False)

            lDocCount = 1
            
                ' Probably don't need to do this but it's only on the first one so not hurting much
            If UnlinkWordFields(oWordApp, oCombinedDoc) = False Then
                LogMessage strProcName, "ERROR", "There was a problem unlinking Word Fields - bar codes may be incorrect?!?!"
                    '                Call ErrorCallStack_Add(0, "There was a problem unlinking Word Fields - bar codes may be incorrect?!?!", strProcName, , , , oRS("InstanceID").Value, oRS("LetterType").Value)
            End If
            
                '2014:04:29:JS Addded this here because InsertWordDocAtStartOfCurrentDoc only does it for the inserted documents now.
            With oCombinedDoc
                .Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                .Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
                .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
                .Repaginate
            End With

            oCombinedDoc.SaveAs2 (sOutFilePathAndName)
            bFirst = False
        Else    ' not the first - this one needs to be added to the combined doc
            If InsertWordDocAtStartOfCurrentDoc(oCombinedDoc, sThisFilePath) = False Then

'                        Call ErrorCallStack_Add(0, "There was a problem adding a letter to the beginning of the merged document", strProcName, sFileName, , , oRS("InstanceID").Value, oRS("LetterType").Value)
                lErrCnt = lErrCnt + 1
                sErrMsg = sErrMsg & sThisFilePath & " was not able to be added to the combined document!" & vbCrLf
                LogMessage strProcName, "ERROR", "Unable to add a letter to the combined doc!", sThisFilePath
            Else
                ' Save every 50 or so letters just in case (saving each time will take FOREVER!)
                If lDocCount Mod 50 = 0 Then
                    oCombinedDoc.Save
                End If
            End If
    

        End If
    
NextFile:
        oRs.MoveNext
    Wend
    oCombinedDoc.Save
    
    If bOpenAfterFinished = True Then
    
    Else
        oCombinedDoc.Close False    ' already saved it
        oWordApp.Quit
    End If
    
    CombineDocsFromRs = True
    
Block_Exit:
    Set oCombinedDoc = Nothing
    Set oWordApp = Nothing
    Set oFso = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




Public Function GenerateAndDisplayPullMessage() As Boolean
On Error GoTo Block_Err
Dim strProcName As String

Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GenerateAndDisplayPullMessage"

    Set oRs = GetPullRS
    If oRs.recordCount > 0 Then
        LogMessage strProcName, "LETTER(S) TO PULL", "There are letters to pull, displaying the report!"
        DoCmd.OpenReport "rpt_BOLD_MailOps_Pull_Request", acViewReport, , , acDialog
        
    End If

Block_Exit:

    Exit Function
Block_Err:
    ReportError Err, strProcName
    GenerateAndDisplayPullMessage = False
    GoTo Block_Exit
End Function


Public Function GetPullRS() As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetPullRS"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_BOLD_LETTER_Automation_MAILOPS_PullNotice"
        Set oRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            LogMessage strProcName, "ERROR", "An error occurred in " & .sqlString, .Parameters("@pErrMsg").Value
            GoTo Block_Exit
        End If
    End With

    Set GetPullRS = oRs

Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Set GetPullRS = Nothing
    GoTo Block_Exit
End Function



Public Function MakePassthroughQryFromSproc(strSprocNameNParams As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oDb As DAO.Database
Dim oQDef As DAO.QueryDef

    strProcName = ClassName & ".MakePassthroughQryFromSproc"
    Debug.Print Now() & " " & strProcName & "." & cs_MAIN_SERVER_NAME
    
        ' Create a passthrough query from our sPassthroughSql
    Set oQDef = New DAO.QueryDef
    Set oDb = CurrentDb()
    
    MakePassthroughQryFromSproc = "ptqry_Temp_Notify_" & Format(Now(), "hhnnss")
    
    If IsQuery(MakePassthroughQryFromSproc) = True Then
        oDb.QueryDefs.Delete (MakePassthroughQryFromSproc)
    End If
    
    With oQDef
        .Name = MakePassthroughQryFromSproc
            '' KD COMEBACK: use the constants here will ya?
        .Connect = GetDSNLessODBCConnString(cs_MAIN_SERVER_NAME, cs_MAIN_DB_NAME)
        .SQL = strSprocNameNParams
    End With
    oDb.QueryDefs.Append oQDef
    oDb.QueryDefs.Refresh
'    DoEvents
'
'        '' Now execute a select into from our above passthrough:
'    If bCreateDestTbl = True Or IsTable(sDestTable) = False Then
'        oDb.Execute ("SELECT * INTO " & sDestTable & " FROM " & sPTQName)
'    Else
'            '' Note: we need to make sure that the SQL matches the dest table
'            '' KD COME back.. for now I'll leave it, but I need to
'            '' make sure that we have the correct column counts - I'm not going
'            '' to deal with the data types.. :P
'        sFieldList = BuildFieldList(GetLikeFieldNames(sPTQName, sDestTable))
'
'        oDb.Execute ("INSERT INTO " & sDestTable & " (" & sFieldList & ") SELECT " & sFieldList & " FROM " & sPTQName)
'    End If
'
'    If IsQuery(sPTQName) = True Then
'            ' clean up our query
'        oDb.QueryDefs.Delete (sPTQName)
'        oDb.QueryDefs.Refresh
''        DoEvents
'    End If
'
'    MakeTableFromPassthroughQry = 1

    MakePassthroughQryFromSproc = MakePassthroughQryFromSproc

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    MakePassthroughQryFromSproc = ""
    GoTo Block_Exit
End Function



Public Function MailRoomMarkBatchAsPulled() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oFrm As Form_frm_BOLD_Mail_Dashboard
Dim oLI As ListItem
Dim sMsg As String
Dim iPullLevel As MailLevels
Dim lPullId As Long
Dim lOpsBatchId As Long
Dim sBatchType As String
Dim sPlural As String

    strProcName = ClassName & ".MailRoomMarkBatchAsPulled"
    
  ' Now, we need to get the item that was clicked - which depends on what list view was clicked..
    ' In this case it's going to be the MAILOPS dashboard - should not be the Mail Room Dashboard
    If IsOpen("frm_BOLD_Mail_Dashboard", acForm) = False Then
        LogMessage strProcName, "ERROR", "Got into this routine without the Mail room Dashboard being opened!"
        LogMessage strProcName, "ERROR", "This function is not enabled in this window yet!", , True
            Stop
        GoTo Block_Exit
    End If
    
    Set oFrm = Application.Forms("frm_BOLD_Mail_Dashboard")

    '' Get the item selected..
    Set oLI = oFrm.RightClickedListItem
    ' What level is this request?
    
    iPullLevel = Nz("0" & oFrm.QueueColumns.GetLiValue(oLI, oLI.Tag, "PullLevel"), 0)
    lPullId = Nz(oFrm.QueueColumns.GetLiValue(oLI, oLI.Tag, "PullId"), 0)
    lOpsBatchId = Nz(oFrm.QueueColumns.GetLiValue(oLI, oLI.Tag, "MailBatchId"), 0)  ' Confusing aren't these names? Not enough time spent designing I suppose
    sBatchType = Nz(oFrm.QueueColumns.GetLiValue(oLI, oLI.Tag, "BatchType"), "Regular Batch")
    
    
    If iPullLevel = 0 Or lPullId = 0 Then
        LogMessage strProcName, "ERROR", "Problem with Pull Level and or ID"
        Stop
        GoTo Block_Exit
    End If
    
    '' How many are requested as being pulled?
    Set oRs = GetPullRS()
        
    If oRs.recordCount = 0 Then
NotThere:
        sMsg = "There do not appear to be any letters (or batches) that are requested to be pulled at this time." & vbCrLf & _
            "Perhaps someone else marked them as pulled!"
            
        MsgBox sMsg, vbOKOnly, "No requests to mark as fulfilled"
        GoTo Block_Exit

    End If
    
    oRs.filter = "OpsLetterBatchId = " & CStr(lOpsBatchId) & " AND BatchType = '" & sBatchType & "'"
        ' still have some?
    If oRs.recordCount = 0 Then
        GoTo NotThere
    End If


    If iPullLevel = Batch Then
        sPlural = IIf(oRs.recordCount > 1, "s", "")
        sMsg = "Please confirm that you have successfully pulled, " & CStr(oRs.recordCount) & " letter" & sPlural
        If MsgBox(sMsg, vbOKCancel, "CONFIRM " & CStr(oRs.recordCount) & "Letter" & sPlural) = vbCancel Then
            LogMessage strProcName, , "User canceled at confirmation"
            Stop
            GoTo Block_Exit
        End If
    End If
    
    
    ' All ready to go:
    ' if it's a batch, then send the batchid
    ' if it's an instance, send the autoid
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_BOLD_Letter_Automation_MAILOPS_MarkAsPulled"
        .Parameters.Refresh
        .Parameters("@pPullId") = IIf(iPullLevel = Batch, lOpsBatchId, lPullId)
        .Parameters("@pPullLevel") = iPullLevel
        
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            LogMessage strProcName, "ERROR", "Something went wrong with the process!", .Parameters("@pErrMsg").Value & " (" & CStr(lPullId) & ", " & CStr(iPullLevel) & ")", True
        End If
    End With


Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oFrm = Nothing
    Set oLI = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    MailRoomMarkBatchAsPulled = False
    GoTo Block_Exit
End Function


Public Function MailRoomMarkInstanceAsPulled() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oFrm As Form_frm_BOLD_Mail_Dashboard
Dim oLI As ListItem
Dim sMsg As String
Dim iPullLevel As MailLevels
Dim lPullId As Long
Dim lOpsBatchId As Long
Dim sBatchType As String
Dim sPlural As String

    strProcName = ClassName & ".MailRoomMarkInstanceAsPulled"
    
  ' Now, we need to get the item that was clicked - which depends on what list view was clicked..
    ' In this case it's going to be the MAILOPS dashboard - should not be the Mail Room Dashboard
    If IsOpen("frm_BOLD_Mail_Dashboard", acForm) = False Then
        LogMessage strProcName, "ERROR", "Got into this routine without the Mail room Dashboard being opened!"
        LogMessage strProcName, "ERROR", "This function is not enabled in this window yet!", , True
            Stop
        GoTo Block_Exit
    End If

Stop ' I really didn't do anything with this code yet - it's a direct copy of the batch one which should be able to be reused instead of this.. Just trying
' to separate for now
Stop


   
    Set oFrm = Application.Forms("frm_BOLD_Mail_Dashboard")

    '' Get the item selected..
    Set oLI = oFrm.RightClickedListItem
    ' What level is this request?
    
    iPullLevel = Nz(oFrm.QueueColumns.GetLiValue(oLI, oLI.Tag, "PullLevel"), 0)
    lPullId = Nz(oFrm.QueueColumns.GetLiValue(oLI, oLI.Tag, "PullId"), 0)
    lOpsBatchId = Nz(oFrm.QueueColumns.GetLiValue(oLI, oLI.Tag, "MailBatchId"), 0)  ' Confusing aren't these names? Not enough time spent designing I suppose
    sBatchType = Nz(oFrm.QueueColumns.GetLiValue(oLI, oLI.Tag, "BatchType"), "Regular Batch")
    
    
    If iPullLevel = 0 Or lPullId = 0 Then
        LogMessage strProcName, "ERROR", "Problem with Pull Level and or ID"
        Stop
        GoTo Block_Exit
    End If
    
    '' How many are requested as being pulled?
    'Call GenerateAndDisplayPullMessage
    Set oRs = GetPullRS
    If oRs.recordCount = 0 Then
NotThere:
        sMsg = "There do not appear to be any letters (or batches) that are requested to be pulled at this time." & vbCrLf & _
            "Perhaps someone else marked them as pulled!"
            
        MsgBox sMsg, vbOKOnly, "No requests to mark as fulfilled"
        GoTo Block_Exit

    End If
    
    oRs.filter = "OpsLetterBatchId = " & CStr(lOpsBatchId) & " AND BatchType = '" & sBatchType & "'"
        ' still have some?
    If oRs.recordCount = 0 Then
        GoTo NotThere
    End If


    If iPullLevel = Batch Then
        sPlural = IIf(oRs.recordCount > 1, "s", "")
        sMsg = "Please confirm that you have successfully pulled, " & CStr(oRs.recordCount) & " letter" & sPlural
        If MsgBox(sMsg, vbOKCancel, "CONFIRM " & CStr(oRs.recordCount) & "Letter" & sPlural) = vbCancel Then
            LogMessage strProcName, , "User canceled at confirmation"
            GoTo Block_Exit
        End If
    End If
    
    
    ' All ready to go:
    ' if it's a batch, then send the batchid
    ' if it's an instance, send the autoid
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_BOLD_Letter_Automation_MAILOPS_MarkAsPulled"
        .Parameters.Refresh
        .Parameters("@pPullId") = IIf(iPullLevel = Batch, lOpsBatchId, lPullId)
        .Parameters("@pPullLevel") = iPullLevel
        
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            LogMessage strProcName, "ERROR", "Something went wrong with the process!", .Parameters("@pErrMsg").Value & " (" & CStr(lPullId) & ", " & CStr(iPullLevel) & ")", True
        End If
    End With


Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oFrm = Nothing
    Set oLI = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    MailRoomMarkInstanceAsPulled = False
    GoTo Block_Exit
End Function
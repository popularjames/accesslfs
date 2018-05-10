Option Compare Database
Option Explicit


''' Last Modified: 12/13/2012
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
'''  - 12/13/2012 - Allow users to be able to manually replace a NIRF (archive the old..)
'''  - 09/12/2012 - Tweaked a number of processes concerning submission

'''  - 08/28/2012 - fixed document attaching
'''  - 06/27/2012 - added bunch of stuff for new concept payer types
'''  - 06/18/2012 - Added IsConceptPayerType
'''  = 06/16/2012 - Added GetRelatedPayers(sConceptId) as collection
'''  - 06/14/2012 - Fixed DateSubmitted
'''  - 06/04/2012 - They wanted to archive the NIRFs instead of deleting them
'''  - 06/01/2012 - Concept SQL button - add some more colors!
'''  - 05/30/2012 - changed NIRF creation to NOT update the date. Submit button is the only
'''         thing that will
'''  - 05/29/2012 - added optional wait for it param to AddAttachedDocToConversionQueue
'''             and fixed a couple other things (making sure the NIRF report was closed before trying
'''             to create it, etc..)
'''  - 05/23/2012 - Added code to get the text of a stored proc for the given concept
'''  - 05/08/2012 - tweaked various things around validation, which buttons are enabled
'''     and which aren't, and file conversion queue stuff
'''  - 04/09/2012 - added PromptUserForTaggedClaimsExceptions and a couple other functions
'''  - 03/27/2012 - Added a bunch more functions and a couple constants
'''  - 03/23/2012 - Added MakeWhereListFromTableList and some supporting functions
'''     - Added SynchTaggedClaims
'''  - 03/12/2012 - Added various methods and functions which will soon be used
'''  - 03/06/2012 - Created...
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

Private Const ClassName As String = "mod_Concept_Specific"

    '' The following enum is basically just the CMS_AUDITORS_ERAC.dbo.CnlyEracActions table
Public Enum enuConceptActions
    PkgIDReqDBPopulate = 1  ' = "Package Request DB Populate - populated the database with info for webcall"
    PkgIDReqBatchExe = 2    ' = "Request Client Issue Id aka Package ID batch script executing"
    PkgIdReqReceipt = 3     ' = "Client Issue ID (aka Package ID) received"
    PkgIDReqTimeout = 4     ' = "Client Issue ID (aka Package ID) Time out while waiting for ID"
    
    '' CMS Notifications
    CmsNotComplete = 5
    CmsNotDetermination = 6
    CmsNotIncomplete = 7
    CmsNotMailing = 8
    CmsNotNewRequest = 9
    CmsNotNewReview = 10
    CmsNotRequestModification = 11
    CmsNotTermination = 12
    
    '' Other stuff.
    SubmitPkgDBPopulate = 13  '= "Set status for submission"
    WebCallSubmitPkg01 = 14
    WebCallSubmitPkg02 = 15
    WebCallSubmitPkg03 = 16
    Submitted = 17
    ValidatedConcept = 18
    PromptUserForClaimNum = 19
    getclientissueid = 20
    NirfCreated = 21
    AttachedDocsCreateFunct = 22
    NotificationNotLoaded = 23
    
    DocumentsConverted      ' = ""
    
    DBPreppedForSubmission  '= "DB prepped for submission"
    CheckFilesExist         ' = "Verified Document Existance"
    Batch1                  ' = "Batch 1: Submit Package"
    Batch2                  '= "Batch 2: Submit Claims"
    Batch3                  '= "Batch 3: Complete submission"
    NotifyRequestor         '= "Notify concept owner about submission"
End Enum


Private Const ciCHART_FILE_ID As Integer = 12

Private Const ciCONCEPT_SUBMIT_EMAIL_NOTIFICIATION_ID As Integer = 2


Public gdctTableIndexes As Scripting.Dictionary
Private mblnSetup As Boolean

Public gdctAuditFieldNames As Scripting.Dictionary


Public Const csCONCEPT_SUBMISSION_WORK_FLDR As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\AUDITCLM_ATTACHMENT\E-RAC\"
Public Const csCONCEPT_SUBMISSION_SAVE_FLDR As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\AUDITCLM_ATTACHMENT\ConceptID\"
    
Private dctOpenForms As Scripting.Dictionary

Public giHdrFormSelectedPage As Integer


Private Const WINZIP_PATH As String = "\\Ccaintranet.com\dfs-cms-fld\Audits\CMS\APPEALS\WINZIP\wzzip.exe "



Private Sub start()
    DoCmd.Hourglass True
    DoCmd.SetWarnings False
    'DoCmd.Echo True, ""
End Sub
Private Sub Done()
    DoCmd.Hourglass False
    DoCmd.SetWarnings True
    DoCmd.Echo True, "Ready..."
    DoEvents
End Sub




Public Sub InsertDetailForPayer(oForm As Form, oPayerDtlRS As ADODB.RecordSet, lngPayerNameId As Long, rsConceptHdr As ADODB.RecordSet, Optional bUpdateHeaderToo As Boolean, _
    Optional sClientIssueNumToUse As String)
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim bResult As Boolean
Dim oCtl As Control


    strProcName = ClassName & ".InsertDetailForPayer"

    If oPayerDtlRS.recordCount < 1 Then
        oPayerDtlRS.AddNew
    Else
        LogMessage strProcName, "ERROR", "It looks like the form may have gotten out of synch. Please close the form and reopen to continue", , True, Nz(oForm.FormConceptID, "")
        GoTo Block_Exit
    End If
   
        'Loop through the controls setting their control source to the recordset
    For Each oCtl In oForm.Controls
        If oCtl.Tag <> "" Then

            If InStr(1, oCtl.Tag, ".", vbTextCompare) > 0 Then
                Select Case UCase(left(oCtl.Tag, InStr(1, oCtl.Tag, ".", vbTextCompare) - 1))
                Case "CONCEPT_HDR"

                Case "CONCEPT_PAYER_DTL"

                    If isField(oPayerDtlRS, oCtl.Name) = True Then
                        oPayerDtlRS.Fields(oCtl.Name).Value = oForm.Controls(oCtl.Name).Value
                    Else
                        LogMessage strProcName, , "Notta field " & oCtl.Name, , , Nz(oForm.FormConceptID, "")
                    End If
                Case "BOTH"
                    If isField(oPayerDtlRS, oCtl.Name) = True Then
                        oPayerDtlRS.Fields(oCtl.Name).Value = oForm.Controls(oCtl.Name).Value
                    Else
                        LogMessage strProcName, , "Notta field " & oCtl.Name, , , Nz(oForm.FormConceptID, "")
                    End If
                    If bUpdateHeaderToo = True Then
                        If rsConceptHdr Is Nothing Then
                            LogMessage strProcName, , "rs is nothing", , , Nz(oForm.FormConceptID, "") ' should never get here
                        ElseIf rsConceptHdr.recordCount < 1 Then
                            LogMessage strProcName, , "Rs count < 1", , , Nz(oForm.FormConceptID, "")
                        End If
    '
                        If isField(rsConceptHdr, oCtl.Name) = True Then
                            rsConceptHdr.Fields(oCtl.Name).Value = oForm.Controls(oCtl.Name).Value
                        End If
                    End If
                End Select
            End If
        End If
    Next
    
    oPayerDtlRS("PayerNameId") = lngPayerNameId
    oPayerDtlRS("ConceptID") = oForm.ConceptID

    If sClientIssueNumToUse <> "" Then
        oPayerDtlRS("ClientIssueNum") = sClientIssueNumToUse
    End If


    oPayerDtlRS.Update
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_CODE_DATABASE")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_PAYER_Dtl_Apply"
        bResult = .Update(oPayerDtlRS, "usp_CONCEPT_PAYER_Dtl_Apply")
        If bResult = False Then
            LogMessage strProcName, "ERROR", "There was an error applying the Payer detail", Nz(oForm.ConceptID, "")
        End If
    End With



Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Nz(oForm.ConceptID, "")
    Err.Clear
    GoTo Block_Exit
End Sub




Public Sub NullOutConceptHdrDueToPayerDtl(oForm As Form, rsConceptHdr As ADODB.RecordSet)
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim bResult As Boolean
Dim oCtl As Control


    strProcName = ClassName & ".NullOutConceptHdrDueToPayerDtl"
    
    If rsConceptHdr Is Nothing Then
        GoTo Block_Exit
    End If
    
    If rsConceptHdr.recordCount <> 1 Then
        GoTo Block_Exit
    End If
    
    rsConceptHdr.MoveFirst
  
        'Loop through the controls setting their control source to the recordset
    For Each oCtl In oForm.Controls
        If oCtl.Tag <> "" Then

            If InStr(1, oCtl.Tag, ".", vbTextCompare) > 0 Then
                Select Case UCase(left(oCtl.Tag, InStr(1, oCtl.Tag, ".", vbTextCompare) - 1))
                Case "CONCEPT_HDR"
    '                    Stop    ' do nothing, leave it be

                Case "CONCEPT_PAYER_DTL"

                    If isField(rsConceptHdr, oCtl.Name) = True Then
                        rsConceptHdr.Fields(oCtl.Name).Value = Null
                    Else
                        LogMessage strProcName, , "Notta field " & oCtl.Name, , , Nz(oForm.FormConceptID, "")
                    End If
                Case "BOTH"
                    If isField(rsConceptHdr, oCtl.Name) = True Then
                        rsConceptHdr.Fields(oCtl.Name).Value = Null
                    Else
                        LogMessage strProcName, , "Notta field " & oCtl.Name, , , Nz(oForm.FormConceptID, "")
                    End If
                End Select
            End If
        End If
    Next
    

    rsConceptHdr.Update
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_CODE_DATABASE")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_Hdr_Update_WithNull"
        bResult = .Update(rsConceptHdr, "usp_CONCEPT_Hdr_Update_WithNull")
        If bResult = False Then
            LogMessage strProcName, "ERROR", "There was an error nulling out the concept header fiels", , , Nz(oForm.ConceptID, "")
        End If
    End With



Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Nz(oForm.ConceptID, "")
    Err.Clear
    GoTo Block_Exit
End Sub

Public Sub ShowFormAndWaitForHwndRemoval(frm As Form)
Dim bFormClosed   As Boolean
Dim strFormName As String
Dim lFrmHwnd As Long


    If dctOpenForms Is Nothing Then
        Set dctOpenForms = New Scripting.Dictionary
    End If


    strFormName = frm.Name
    lFrmHwnd = frm.hwnd
    If dctOpenForms.Exists(CStr("" & lFrmHwnd)) = False Then
        dctOpenForms.Add CStr("" & lFrmHwnd), strFormName
    End If
    frm.visible = True
     
    Do
        'Is it still Open?
        If dctOpenForms.Exists(CStr("" & lFrmHwnd)) = True Then
            DoEvents
            Wait 1
        Else
            bFormClosed = True
        End If
    
    Loop Until bFormClosed
End Sub


Public Sub UnloadForm(oForm As Form)
Dim lFrmHwnd As Long

    lFrmHwnd = oForm.hwnd
    If dctOpenForms.Exists(CStr("" & lFrmHwnd)) = True Then
        dctOpenForms.Remove CStr("" & lFrmHwnd)
    End If
    oForm.visible = False
     
End Sub


Public Function PromptUserForPayers(sPromptText As String, oConcept As clsConcept, Optional sLimitToThesePayerIds As String, _
    Optional saryPayerIdsStr As String, Optional saryPayerNamesArray As Variant, Optional bShowOldPayers As Boolean) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim ofPayers As Form_frm_Prompt_For_Payers
Dim saryPayerIdsArray() As String

    strProcName = ClassName & ".IsConceptPayerType"
    

        ' This needs to do the following:
        ' Prompt the user for the payers that pertain to this concept
    Set ofPayers = New Form_frm_Prompt_For_Payers
    With ofPayers
        .PromptText = sPromptText
        .AllowedPayers = sLimitToThesePayerIds
        .ShowPastPayers = bShowOldPayers
        ShowFormAndWaitForHwndRemoval ofPayers
    End With
    
    
    
    If IsOpen("frm_Prompt_For_Payers") = True Then
'        Set ofPayers = Forms("frm_Prompt_For_Payers")
'        Stop
    Else
        ''  canceled..
        LogMessage strProcName, , "User canceled", , , oConcept.ConceptID
        GoTo Block_Exit
    End If
    
    saryPayerNamesArray = ofPayers.SelPayerNamesArray
    saryPayerIdsStr = ofPayers.SelPayerNameIdsString
    saryPayerIdsArray = Split(ofPayers.SelPayerNameIdsString, ",")
    
    
        ' just a quick double check here..
    If UBound(saryPayerNamesArray) < 0 Then
        PromptUserForPayers = False
        GoTo Block_Exit
    End If
    If UBound(saryPayerIdsArray) < 0 Then
        PromptUserForPayers = False
        GoTo Block_Exit
    End If
    
    DoCmd.Close acForm, "frm_Prompt_For_Payers", acSaveNo
    
    PromptUserForPayers = True
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    PromptUserForPayers = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function ConceptConvertConceptDtlCodes(oConcept As clsConcept, saryAllowedPayerIds As Variant) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim iPayerID As Integer
Dim oAdo As clsADO

    strProcName = ClassName & ".ConceptConvertConceptDtlCodes"

    LogMessage strProcName, , "Starting to convert for " & oConcept.ConceptID, "Params: " & Join(saryAllowedPayerIds, ","), , oConcept.ConceptID
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_CVT_Codes"
        .Parameters.Refresh
        .Parameters("@pConceptID") = oConcept.ConceptID
        .Parameters("@pPayerIDs") = Join(saryAllowedPayerIds, ",")
        Call .Execute
        If Nz(.Parameters("@pErrMsg"), "") <> "" Then
            LogMessage strProcName, "ERROR", "There was a problem converting the codes", .Parameters("@pErrMsg"), True, oConcept.ConceptID
            GoTo Block_Exit
        End If
    End With
    
    ConceptConvertConceptDtlCodes = True
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    ConceptConvertConceptDtlCodes = False
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function ConceptConvertConceptStates(oConcept As clsConcept, saryAllowedPayerIds As Variant) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim iPayerID As Integer
Dim oAdo As clsADO

    strProcName = ClassName & ".ConceptConvertConceptStates"
    
    ' We can only do this if we only have 1 payer.. and even then, I'm not sure
    If IsArray(saryAllowedPayerIds) = True Then
        If UBound(saryAllowedPayerIds) = 0 Then
            Set oAdo = New clsADO
            With oAdo
                .ConnectionString = GetConnectString("V_Data_Database")
                .SQLTextType = sqltext
                .sqlString = "UPDATE CONCEPT_Dtl_State SET PayerNameId = " & saryAllowedPayerIds(0) & " WHERE " & _
                        " ConceptId = '" & oConcept.ConceptID & "' AND PayerNameID IS NULL "
                If .Execute < 0 Then
                    GoTo Block_Exit
                End If
            End With
            
            ConceptConvertConceptStates = True
        Else
            LogMessage strProcName, "USER NOTIFICATION", "The States procedure has not yet run for this concept. As a result, state details will not be 'per payer' yet.", , True, oConcept.ConceptID
        End If
    End If
    
    GoTo Block_Exit
    

    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    ConceptConvertConceptStates = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function ConceptConvertTaggedClaims(oConcept As clsConcept) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim iPayerID As Integer
Dim oAdo As clsADO

    strProcName = ClassName & ".ConceptConvertTaggedClaims"
    
        ' we are only going to delete everything, that way it'll be reimported


    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = sqltext
        .sqlString = "DELETE FROM CnlyTaggedClaimsByConcept WHERE " & _
                " ConceptId = '" & oConcept.ConceptID & "'"
        
        Call .Execute

    End With
    
    ConceptConvertTaggedClaims = True
    
    GoTo Block_Exit
    
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    ConceptConvertTaggedClaims = False
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function ConceptConvertTracking(oConcept As clsConcept, saryAllowedPayerIds As Variant) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim iPayerID As Integer
Dim oAdo As clsADO

    strProcName = ClassName & ".ConceptConvertTracking"
    
        ' We can only do this if we only have 1 payer.. and even then, I'm not sure
    If IsArray(saryAllowedPayerIds) = True Then
        If UBound(saryAllowedPayerIds) = 0 Then
            Set oAdo = New clsADO
            With oAdo
                .ConnectionString = GetConnectString("v_Data_Database")
                .SQLTextType = sqltext
                .sqlString = "UPDATE CONCEPT_Tracking SET PayerNameId = " & saryAllowedPayerIds(0) & " WHERE " & _
                        " ConceptId = '" & oConcept.ConceptID & "' AND PayerNameID IS NULL "
                If .Execute < 0 Then
                    GoTo Block_Exit
                End If
            End With
            
            ConceptConvertTracking = True
        Else
'            Stop    ' what do we do if there are more than 1 payer?
            ' does nothing work? I think it will as these will be concidered Concept Tracknig (not, Concept Payer tracking)
        End If
    End If
    
    GoTo Block_Exit
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    ConceptConvertTracking = False
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'Public Function PromptForPayerOnConceptReferences(oConcept As clsConcept, saryAllowedPayerIds() As String) As Boolean
Public Function PromptForPayerOnConceptReferences(oConcept As clsConcept, sAllowedPayerIds As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAtchDoc As clsConceptDoc
Dim oForm As Form_frm_Prompt_For_Payers
Dim saryDocPayerIds() As String
Dim sDocPayerIdStr As String
Dim iPayerID As Integer
Dim sPromptText As String
Dim oCReqDocType As clsConceptReqDocType
Dim saryAllowedPayerIds() As String
Dim dctRowNPayers As Scripting.Dictionary
Dim vRowId As Variant
Dim lNewRowId As Long

    strProcName = ClassName & ".PromptForPayerOnConceptReferences"
    
    saryAllowedPayerIds = Split(sAllowedPayerIds, ",")
    Set dctRowNPayers = New Scripting.Dictionary

    For Each oAtchDoc In oConcept.AttachedDocuments
        Set oCReqDocType = New clsConceptReqDocType
        If oCReqDocType.LoadFromDocName(oAtchDoc.DocTypeName) = False Then
            Stop    ' this may be an old document type like a screen grab or what have you
                    ' from CnlyDocTypes table (not the new, ConceptDocTypes table)
                    ' Which by the way, is the main reason we are using the DocTypeName (which is the
                    ' RefSubType in the _CLAIMS.dbo.CONCEPT_References table
            GoTo SkipThisDoc
        End If
    
        If oCReqDocType.IsPayerDoc = True Then
                '   if there is only 1 payer for the whole concept then there's no need to prompt
            If UBound(saryAllowedPayerIds) > 0 Then
                
                sPromptText = "Please select the payers for this " & oAtchDoc.DocTypeName & " attached document: " & vbCrLf & _
                        oAtchDoc.FileName
                    
                If PromptUserForPayers(sPromptText, oConcept, Join(saryAllowedPayerIds, ","), sDocPayerIdStr) = False Then
                            ' user cancel or error or something!
                    LogMessage strProcName, "ERROR", "There seems to have been an error while prompting the user to pick the payer for this document!", _
                            oAtchDoc.FileName, , oConcept.ConceptID
                    GoTo SkipThisDoc
                End If
                
            Else
                ' Only 1 payer for this concept so we are using that to assign to all docs.
                ReDim saryDocPayerIds(0)
                saryDocPayerIds(0) = CInt(saryAllowedPayerIds(0))
            End If

            saryDocPayerIds = Split(sDocPayerIdStr, ",")

            If UBound(saryDocPayerIds) < 0 Then
                ' no payers.. what's up with that?  Could be a 'Concept' Specific document
                ' which is really for phase 2 as phase 1 we are making copies of the document for
                ' each payer..
                If ConvertConceptReference(oConcept.ConceptID, oAtchDoc.RowID) = False Then
                    LogMessage strProcName, "ERROR", "There was a problem converting the reference for Row: " & oAtchDoc.RowID, , , oConcept.ConceptID
                End If
            
            End If
            
            For iPayerID = 0 To UBound(saryDocPayerIds)
                If iPayerID = 0 Then
                    ' no insert into Concept_References, only into the _ERAC database
                    Call InsertPayerDocRecord(CLng(saryDocPayerIds(iPayerID)), oConcept.ConceptID, oAtchDoc.RowID)
                    If dctRowNPayers.Exists(CStr(oAtchDoc.RowID)) = True Then
                        ' why? shouldn't have it more than one?
                    Else
                        dctRowNPayers.Add CStr(oAtchDoc.RowID), saryDocPayerIds(iPayerID)
                    End If
                Else
                    ' need to insert into Concept_References a copy of it
                    If CopyAttachTypeFromRowId(CLng(saryDocPayerIds(iPayerID)), oConcept.ConceptID, oAtchDoc.RowID, lNewRowId) = False Then
                        LogMessage strProcName, "ERROR", "There was a problem with creating a copy of this document for a different payer!", CStr(oAtchDoc.Id), , oConcept.ConceptID
                    End If
                    If dctRowNPayers.Exists(CStr(lNewRowId)) = True Then
    '                        Stop    ' why? shouldn't have it more than once
                    Else
                        dctRowNPayers.Add CStr(lNewRowId), saryDocPayerIds(iPayerID)
                    End If
                End If
            Next
            
            ' Don't forget, we need to create a subfolder for the payername
            ' and move this document there..
        ' do we or is that done somewhere else? yes, we do that somewhere else.. But, we should do it here because
        ' that's what's playing havvok with the database fields.. well part of it
        ' so we are going to do it down below..
            
        Else    ' concept document
            If ConvertConceptReference(oConcept.ConceptID, oAtchDoc.RowID) = False Then
                LogMessage strProcName, "ERROR", "There was a problem converting the concept document for Row: " & oAtchDoc.RowID, , , oConcept.ConceptID
            End If
            
        End If
SkipThisDoc:
    Next


    ''
    '' Now that we're done inserting and such,
    '   we need to move the file into a subdirectory named with the payername
'    ' then update the CONCEPT_References where RowId = ..


Dim sDelFilePath As String
Dim dctDelFiles As Scripting.Dictionary
Dim vKey As Variant

    Set dctDelFiles = New Scripting.Dictionary
    
    For Each vRowId In dctRowNPayers.Keys
        If MovePayerReferenceToPayerSubFldr(CLng(vRowId), dctRowNPayers.Item(vRowId), sDelFilePath) = False Then
            LogMessage strProcName, "ERROR", "There was a problem moving the reference to a payer sub folder", "Row: " & CStr(vRowId), , oConcept.ConceptID
        End If
        If dctDelFiles.Exists(sDelFilePath) = False Then
            dctDelFiles.Add sDelFilePath, 0
        End If
    Next


    ' ANd finally, we can delete the files:
    For Each vKey In dctDelFiles.Keys
        If DeleteFile(CStr(vKey), False) = False Then
            LogMessage strProcName, "ERROR", "Could not delete orig file after copying to payer subfolder", CStr(vKey), , oConcept.ConceptID
        End If
    Next
    
    PromptForPayerOnConceptReferences = True

Block_Exit:
    Set oAtchDoc = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''  Caches boolean values of concepts - true if they are new ones that have payer specific
'''   data, false if they are the old ones (old as of 6/19/2012)
'''
Public Function IsConceptPayerType(Optional sConceptId As String, Optional oConcept As clsConcept, Optional bForceRefresh As Boolean = False) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim dctConcepts As Scripting.Dictionary
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim bReturnVal As Boolean

    strProcName = ClassName & ".IsConceptPayerType"
    
    
    If oConcept Is Nothing And sConceptId = "" Then
        LogMessage strProcName, "ERROR", "No concept passed to function!"
        ' debug! this shouldn't get here kev!  Why?!?!?!?!?  FIX!!!!
        IsConceptPayerType = False
        GoTo Block_Exit
    End If
    
    If sConceptId = "" Then sConceptId = oConcept.ConceptID
    
    If dctConcepts Is Nothing Or bForceRefresh = True Then
        Set dctConcepts = New Scripting.Dictionary
    End If


    If dctConcepts.Exists(sConceptId) = False Then
        sSql = "SELECT * FROM CONCEPT_Hdr AS H INNER JOIN CONCEPT_PAYER_Dtl AS PD ON H.ConceptId = PD.ConceptID " & _
            " WHERE PD.PayerNameId <> 1000 " & _
            " AND H.ConceptID = '" & sConceptId & "'"   '' Note: 1000 = All
            
            
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("V_DATA_DATABASE")
            .SQLTextType = sqltext
            .sqlString = sSql
            Set oRs = .ExecuteRS
            If .GotData = False Then
                bReturnVal = False
            End If
            If oRs.recordCount > 0 Then bReturnVal = True
        End With
        
        IsConceptPayerType = bReturnVal
        dctConcepts.Add sConceptId, bReturnVal
    Else
        IsConceptPayerType = dctConcepts.Item(sConceptId)
    End If

    
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function PrepConceptSubmitEmail(oConcept As clsConcept, lngPayerNameId As Long, _
        Optional sResubmitFldr As String, Optional sOutMsg As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oNotification As clsNotification
Dim sMsgBody As String
Dim sSubject As String
Dim sAttachPaths As String
Dim sClaimDocNote As String


    strProcName = ClassName & ".PrepConceptSubmitEmail"
    Set oNotification = New clsNotification
    If oNotification.LoadFromId(ciCONCEPT_SUBMIT_EMAIL_NOTIFICIATION_ID) = False Then
        LogMessage strProcName, "ERROR", "Could not load notification object", CStr(ciCONCEPT_SUBMIT_EMAIL_NOTIFICIATION_ID), , oConcept.ConceptID

        LogActionToHistory oConcept.ConceptID, NotificationNotLoaded, "Failure to load notification id", , , CStr(ciCONCEPT_SUBMIT_EMAIL_NOTIFICIATION_ID)
    
        GoTo Block_Exit
    End If
        
    '' One thing to note.. If there are no claim level documents (dtllvl) then we don't want to say anything about that..
    If oConcept.TaggedClaims.Count > 0 Then
        sClaimDocNote = vbCrLf & vbCrLf & "The claim documents for this concept will arrive on a CD via regular mail in a few days." & vbCrLf & vbCrLf
    End If
            
        
    sMsgBody = oConcept.ParseStringForDetails(oNotification.EmailMsg, lngPayerNameId)
    sSubject = oConcept.ParseStringForDetails(oNotification.EmailSubject, lngPayerNameId)
    sMsgBody = Replace(sMsgBody, "[Claim_Level_Document_Note]", sClaimDocNote, , , vbTextCompare)
    
        '' Build my comma separated list of attachments
'    sAttachPaths = oConcept.SubmitDocPaths(, sResubmitFldr)
        ' We are now only sending the NIRF...
Dim oAtch As clsConceptDoc
        
    For Each oAtch In oConcept.AttachedDocuments
        If oAtch.PayerNameId = lngPayerNameId Then
            If oAtch.CnlyAttachType = "ERAC_NIRF" Then
                sAttachPaths = QualifyFldrPath(oAtch.FolderPath) & oAtch.FileName
                
                sAttachPaths = oAtch.ConvertedFilePath
                
'                sAttachPaths = oAtch.ImageLink
'
'                sAttachPaths = Mid(sAttachPaths, InStr(1, sAttachPaths, "#") + 1)
                
                If FileExists(sAttachPaths) = False Then
                    LogMessage strProcName, "ERROR", "Could not retrieve"
                End If
                Exit For
            End If
        End If
    Next
    
    If sAttachPaths = "" Then
        Stop
    End If
    
    If SendOutlookEmail(oNotification.EmailTo, sMsgBody, sSubject, sAttachPaths) = False Then
        LogActionToHistory oConcept.ConceptID, NotifyRequestor, "Failure to create outlook email", , , oConcept.ConceptID
        GoTo Block_Exit
    End If

    LogActionToHistory oConcept.ConceptID, lngPayerNameId, NotifyRequestor, "Success creating outlook email", , , oConcept.ConceptID
            
    PrepConceptSubmitEmail = True
    
Block_Exit:
    DoCmd.Hourglass False
    DoCmd.Echo True, "Ready..."
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    PrepConceptSubmitEmail = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function SubmitConcept(oConcept As clsConcept, lPayerNameId As Long, Optional sOutMsg As String, Optional bPassedValidation As Boolean = False) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sClientIssueId As String
Dim oAtchdDoc As clsConceptDoc
Dim oDocType As clsConceptReqDocType
Dim sOutFolder As String
Dim sOutFileName As String
Dim sNewPath As String
Dim sAttachFilePath As String
Dim lBatchId As Long
Dim oReqDoc As clsConceptReqDocType
Dim sCreatedFilePath As String
Dim bTimedOutWaiting As Boolean
Dim oPayerDtl As clsConceptPayerDtl
Dim sThisPayerName As String
Dim bPrompt4OtherPayers As Boolean


    strProcName = ClassName & ".SubmitConcept"
    
    MsgBox "Submitting is not quite yet available. Please notify IT that this concept is ready to submit"
    GoTo Block_Exit
    
    For Each oPayerDtl In oConcept.ConceptPayers
    ' KD finish this up: need to enforce the prompt once thing...
            '' If we don't have the expected amount of tagged claims, then prompt the user
            '' to see if this is an exception, if so, save how many we are supposed to submit
        If PromptUserForTaggedClaimsException(oConcept, oPayerDtl.PayerNameId, oConcept.TaggedClaims.Count, oConcept.RequiredClaimsNum, bPrompt4OtherPayers) = False Then
            sOutMsg = "Discrepancy with the number of tagged claims for concept: " & oConcept.ConceptID
            LogMessage strProcName, "ERROR", sOutMsg, , True, oConcept.ConceptID
            
            LogActionToHistory oConcept.ConceptID, lPayerNameId, PromptUserForClaimNum, "Failure ", , , sOutMsg
        
            GoTo Block_Exit
        End If
    
        LogActionToHistory oConcept.ConceptID, lPayerNameId, PromptUserForClaimNum, "Finished checking Claim nums / Prompting", , , sOutMsg
        

        If oPayerDtl.ClientIssueId <> "" Then
    
            LogActionToHistory oConcept.ConceptID, getclientissueid, "Already have ID", , , oPayerDtl.ClientIssueId
        
            GoTo AlreadyHaveClientIssueId
        End If
    
        sClientIssueId = IssueClientIssueNum(oConcept, oPayerDtl.PayerNameId, sOutMsg)
   
AlreadyHaveClientIssueId:
        DoCmd.Hourglass True
        DoCmd.Echo True, "Copying and converting required documents as needed"
        
                ' Go through all of the required documents and if they require creation, create them! :D
        For Each oReqDoc In oConcept.RequirementRuleObj.RequiredDocs
                ' I don't have to, but I'm going to skip the CreatePackageNirf and do that next
            If oReqDoc.CreateFunctionName <> "" And oReqDoc.CreateFunctionName <> "CreatePackageNirf" Then
                Application.Run oReqDoc.CreateFunctionName, oConcept.ConceptID, CInt(oPayerDtl.PayerNameId)
                Sleep 500   ' just in case
                    ' NOTE: the create function should copy the files to the work folder
    
                LogActionToHistory oConcept.ConceptID, AttachedDocsCreateFunct, "Finished", , , oReqDoc.CreateFunctionName
        
            End If
        Next
        
    
    Next
    
        '' get the client issue id (if we don't already have it)
    DoCmd.Hourglass True
    DoCmd.Echo True, "Getting Client Issue ID"

    
    Stop
    ' shouldn't get here anymore
    

 
        '' Create the NIRF (and attach it) - ok this is going to be called below
    '' Update the submit date (assuming it's not already set)
    If oConcept.DateSubmitted = CDate("1/1/1900") Then
        oConcept.DateSubmitted = Now()
        oConcept.SaveNow
    End If
        ' if the nirf is already there, then prompt to see if they should make a new one
    If oConcept.NIRF_Exists = True Then
        If MsgBox("There is an existing NIRF. Do you want to replace it with a new one?", vbYesNo, "Replace Existing NIRF?") = vbYes Then
            If CreatePackageNirf(oConcept.ConceptID, 1000, False, False) = False Then
                sOutMsg = "There was a problem creating the NIRF for concept: " & oConcept.ConceptID
                LogMessage strProcName, "ERROR", sOutMsg, , True, oConcept.ConceptID
        
                LogActionToHistory oConcept.ConceptID, NirfCreated, "Failure creating nirf", , , CLng(oPayerDtl.PayerNameId)
            
                GoTo Block_Exit
            End If
        End If
    Else
        If CreatePackageNirf(oConcept.ConceptID, CLng(oPayerDtl.PayerNameId), False, False) = False Then
            sOutMsg = "There was a problem creating the NIRF for concept: " & oConcept.ConceptID
            LogMessage strProcName, "ERROR", sOutMsg, , True, oConcept.ConceptID
    
            LogActionToHistory oConcept.ConceptID, NirfCreated, "Failure creating nirf", , , oConcept.ClientIssueId(oPayerDtl.PayerNameId)
        
            GoTo Block_Exit
        End If
    End If
    Sleep 1500  ' sleep for the NIRF to be converted     (even though it should have been..
    
    LogActionToHistory oConcept.ConceptID, NirfCreated, "Success", , , oConcept.ClientIssueId(oPayerDtl.PayerNameId)
    
        '' Make sure all documents are converted and ready to go
    For Each oAtchdDoc In oConcept.AttachedDocuments
        sAttachFilePath = oAtchdDoc.ConvertedFilePath
        Set oDocType = oAtchdDoc.GetEracReqDocType
'''Debug.Assert oAtchdDoc.DocTypeName <> "ERAC_ScreenShot"
        
        sNewPath = oDocType.ParseFileName(oConcept.ConceptID, oConcept.ClientIssueId(oPayerDtl.PayerNameId), oAtchdDoc.Icn, sAttachFilePath)
        
        sOutFolder = ParentFolderPath(sNewPath)
Stop ' kd: didn't do this yet.
        If sOutFolder = "" Or sOutFolder = "\" Then
            If oAtchdDoc.GetEracReqDocType.IsPayerDoc Then
                sOutFolder = oConcept.ConceptWorkFolder & "_BURN\"
            Else
                sOutFolder = oConcept.ConceptWorkFolder
            End If
            
        End If
        
        sOutFileName = Replace(sNewPath, sOutFolder, "")
        
        If oDocType.CreateFunctionName <> "" And oDocType.CreateFunctionName <> "CreatePackageNirf" Then
            Application.Run oDocType.CreateFunctionName, oConcept.ConceptID
            Sleep 500   ' just in case
        End If
        
        If FileExists(sAttachFilePath) = False Then
            DoCmd.Echo True, "Converting files, please check your email for notification when it's done"
            sOutMsg = "Some files were not converted on attachment. They are being converted to the proper type now. Please check your email for a notice that " & _
                " they have completed being processed "
                
            If lBatchId < 1 Then
                    ' create a batch of individual jobs in that batch:
                If AddConverterQueueBatch(oAtchdDoc.FolderPath, False, oDocType.SendAsFileType, sOutFolder, , 1, True, False, , , , lBatchId) = False Then
                    LogMessage strProcName, "ERROR", "There was a problem adding a batch conversion, you may have to convert some files manually", , True, oConcept.ConceptID

                    LogActionToHistory oConcept.ConceptID, DocumentsConverted, "Failure adding batch", , , oConcept.ClientIssueId(oPayerDtl.PayerNameId)
    
                    GoTo Block_Exit
                End If
            End If
            
            If mod_ConverterQueueAPI.AddConverterQueueJob(oAtchdDoc.FolderPath & oAtchdDoc.FileName, oDocType.SendAsFileType, sOutFolder, sOutFileName, False, True, False, 0, lBatchId) = False Then
                LogMessage strProcName, "ERROR", "There was a problem adding a convert job, you may have to convert this file manually: " & oDocType.DocName, , , oConcept.ConceptID

                LogActionToHistory oConcept.ConceptID, DocumentsConverted, "Failure adding job", , , oConcept.ClientIssueId(oPayerDtl.PayerNameId)
    
                GoTo NextAttachedFile
            End If
            
        End If
NextAttachedFile:
    Next

    If lBatchId > 0 Then
    
            ' Close the batch
        If CloseBatch(lBatchId) = False Then
            LogMessage strProcName, "ERROR", "Problem closing the batch: " & CStr(lBatchId), CStr(lBatchId), , oConcept.ConceptID
            
            LogActionToHistory oConcept.ConceptID, DocumentsConverted, "Failure closing batch", , , oConcept.ClientIssueId(oPayerDtl.PayerNameId)
            
        Else
                ' Wait for files to be converted..
            If WaitForBatchOrJobFinish(lBatchId, , sOutMsg, bTimedOutWaiting) = False Then
                LogMessage strProcName, "ERROR", "There was a problem converting the files to proper name / format", "Timed out: " & CStr(bTimedOutWaiting), True, oConcept.ConceptID
                
                LogActionToHistory oConcept.ConceptID, DocumentsConverted, "Failure waiting for conversion", , , CStr(bTimedOutWaiting)
                
                GoTo Block_Exit
            End If
        End If
    End If

        '' Now, create the email from Ken's template (ok, from my template that Ken approved)
    Call PrepConceptSubmitEmail(oConcept, lPayerNameId)

        ' And finally, mark it as being submitted
    Call EracSetConceptAsSubmitted(oConcept)

    SubmitConcept = True

Block_Exit:
    Set oAdo = Nothing
    DoCmd.Hourglass False
    DoCmd.Echo True, "Ready..."
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    SubmitConcept = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function CreateResubmitFolder_n_Package(oConcept As clsConcept, Optional sResubmitFolderPath As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAtchDoc As clsConceptDoc
Dim oRequireDocType As clsConceptReqDocType
Dim sOrigFilePath As String
Dim sNewFileName As String
Dim sConceptOutFldr As String
Dim sOutFolder As String
Dim iResubmitCnt As Integer
Dim sIcn As String
Dim oClaim As clsEracClaim

    strProcName = ClassName & ".CreateResubmitFolder_n_Package"

    iResubmitCnt = 1
    sConceptOutFldr = oConcept.ConceptWorkFolder
    sOutFolder = sConceptOutFldr & "Resubmit_" & Format(iResubmitCnt, "00#") & "\"
    
    While FolderExists(sOutFolder)
        iResubmitCnt = iResubmitCnt + 1
        sOutFolder = sConceptOutFldr & "Resubmit_" & Format(iResubmitCnt, "00#") & "\"
    Wend

    sResubmitFolderPath = sOutFolder
    Call CreateFolders(sResubmitFolderPath)


    For Each oAtchDoc In oConcept.AttachedDocuments
        
        sOrigFilePath = oAtchDoc.FolderPath & oAtchDoc.FileName
        Set oRequireDocType = oAtchDoc.GetEracReqDocType
        
        If oAtchDoc.eRacTaggedClaimId <> 0 Then
            Set oClaim = GetClaimDetailsFromEracTaggedClaimId(oAtchDoc.eRacTaggedClaimId)
            sNewFileName = oRequireDocType.ParseFileName(oConcept.ConceptID, oConcept.ClientIssueId(0), Nz(oClaim.Icn, ""))
            
        Else
            If oAtchDoc.CnlyAttachType = "ERAC_ScreenShot" Then
                sIcn = left(oAtchDoc.FileName, InStr(1, oAtchDoc.FileName, ".") - 1)
                sNewFileName = oRequireDocType.ParseFileName(oConcept.ConceptID, oConcept.ClientIssueId(0), sIcn)
            Else
                sNewFileName = oRequireDocType.ParseFileName(oConcept.ConceptID, oConcept.ClientIssueId(0), "")
            End If
        End If
    
        
        Call mod_ConverterQueueAPI.AddConverterQueueJob(sOrigFilePath, oAtchDoc.GetEracReqDocType.SendAsFileType, sOutFolder, sNewFileName, False, False, False)
        
    Next
    CreateResubmitFolder_n_Package = True

Block_Exit:
    DoCmd.Hourglass False
    DoCmd.Echo True, "Ready..."
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    CreateResubmitFolder_n_Package = False
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function IssueClientIssueNum(oConcept As clsConcept, Optional lPayerNameId As Long, Optional sOutMsg As String = "") As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sClientIssueId As String

    strProcName = ClassName & ".IssueClientIssueNum"
    
    If lPayerNameId = 0 Or lPayerNameId = 1000 Then
        LogMessage strProcName, , "No payer name id!", , , oConcept.ConceptID
        GoTo Block_Exit
    End If
 

    oConcept.RefreshPayerCollection
 
        ' Get the ClientIssueId
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        
        .sqlString = "usp_CMS_Get_New_ClientIssueNum"
        .Parameters.Refresh
        .Parameters("@pConceptID") = oConcept.ConceptID
        .Parameters("@pPayerNameId") = lPayerNameId
        .Execute
        If CStr("" & .Parameters("@pErrMsg").Value) <> "" Then
            sOutMsg = "There was a problem generating the Client Issue Id for concept: " & oConcept.ConceptID & " " & CStr("" & .Parameters("@pErrMsg").Value)
            LogMessage strProcName, "ERROR", sOutMsg, oConcept.ConceptID, True, oConcept.ConceptID

            LogActionToHistory oConcept.ConceptID, lPayerNameId, getclientissueid, "Failure getting client issue id", , , sOutMsg
    
            GoTo Block_Exit
        End If
        sClientIssueId = Nz(.Parameters("@pClientIssueId").Value, "")
    End With

    If sClientIssueId = "" Then
        sOutMsg = "There was a problem generating the Client Issue Id for concept: " & oConcept.ConceptID
        LogMessage strProcName, "ERROR", sOutMsg, oConcept.ConceptID, True, oConcept.ConceptID

        LogActionToHistory oConcept.ConceptID, lPayerNameId, getclientissueid, "Failure getting client issue id", , , ""
        
        GoTo Block_Exit
    End If
    
    Call oConcept.SetClientIssueNum(lPayerNameId, sClientIssueId)

    IssueClientIssueNum = sClientIssueId
Block_Exit:
    Set oAdo = Nothing
    DoCmd.Hourglass False
    DoCmd.Echo True, "Ready..."
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function IssueClientIssueNumToHdr(oConcept As clsConcept, Optional sOutMsg As String = "") As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sClientIssueId As String

    strProcName = ClassName & ".IssueClientIssueNumToHdr"

 
        ' Get the ClientIssueId
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        
        .sqlString = "usp_CMS_Get_New_ClientIssueNum"
        .Parameters.Refresh
        .Parameters("@pConceptID") = oConcept.ConceptID

        .Execute
        If CStr("" & .Parameters("@pErrMsg").Value) <> "" Then
            sOutMsg = "There was a problem generating the Client Issue Id for concept: " & oConcept.ConceptID & " " & CStr("" & .Parameters("@pErrMsg").Value)
            LogMessage strProcName, "ERROR", sOutMsg, oConcept.ConceptID, True, oConcept.ConceptID

            LogActionToHistory oConcept.ConceptID, 0, getclientissueid, "Failure getting client issue id", , , sOutMsg
    
            GoTo Block_Exit
        End If
        sClientIssueId = Nz(.Parameters("@pClientIssueId").Value, "")
    End With

    If sClientIssueId = "" Then
        sOutMsg = "There was a problem generating the Client Issue Id for concept: " & oConcept.ConceptID
        LogMessage strProcName, "ERROR", sOutMsg, oConcept.ConceptID, True, oConcept.ConceptID

        LogActionToHistory oConcept.ConceptID, 0, getclientissueid, "Failure getting client issue id", , , ""
        
        GoTo Block_Exit
    End If
    
    Call oConcept.SetClientIssueNum(0, sClientIssueId)

    IssueClientIssueNumToHdr = sClientIssueId
Block_Exit:
    Set oAdo = Nothing
    DoCmd.Hourglass False
    DoCmd.Echo True, "Ready..."
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function InsertPayerDocRecord(intPayerId As Integer, sConceptId As String, lRowId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sSql As String
Dim sPayerN As String


    strProcName = ClassName & ".InsertPayerDocRecord"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_ConceptReferences_Insert_PayerDtl"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pRowId") = lRowId
        .Parameters("@pPayerNameID") = intPayerId
        Call .Execute
        If .Parameters("@pErrMsg").Value <> "" Then
            LogMessage strProcName, "ERROR", "An error occurred trying to insert records with " & .sqlString, .Parameters("@pErrMsg").Value, True, sConceptId
            GoTo Block_Exit
        End If
        
    End With
    
    
    InsertPayerDocRecord = True
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    InsertPayerDocRecord = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function MovePayerReferenceToPayerSubFldr(lRowId As Long, lngPayerNameId As Long, Optional sOrigFilePathForDelete As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim sPayerName As String
Dim sOldPath As String
Dim SFileName As String
Dim sNewPath As String
Dim sImgLnk As String


    strProcName = ClassName & ".MovePayerReferenceToPayerSubFldr"

    sPayerName = GetPayerNameFromID(lngPayerNameId)
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_DATA_Database")
        .SQLTextType = sqltext
        .sqlString = "SELECT * FROM Concept_References WHERE RowID = " & CStr(lRowId)
        
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "An error occurred trying to get the reference detail so we could move it to payer subfolder", , True
            GoTo Block_Exit
        End If
        
    End With
    
    ' Ok, got the details, now,
    sOldPath = QualifyFldrPath(Nz(oRs("RefPath").Value, ""))
    SFileName = Nz(oRs("RefFileName").Value, "")
    sImgLnk = Nz(oRs("RefLink").Value, "")
    
    
    sNewPath = QualifyFldrPath(sOldPath & sPayerName)
    sImgLnk = Replace(sImgLnk, sOldPath, sNewPath, 1, 1, vbTextCompare)

        ' if we are able to move it, then we need to update the database (but only if we can move it)
    Call CreateFolders(sNewPath)
    If CopyFile(sOldPath & SFileName, sNewPath & SFileName, False) = False Then
        LogMessage strProcName, "ERROR", "There was a problem copying the file to a payer sub directory", sOldPath & SFileName & " to " & sNewPath & SFileName
        GoTo Block_Exit
    End If
    sOrigFilePathForDelete = sOldPath & SFileName

    
        ' Update the db..
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_MoveRefToPayerSFldr"
        .Parameters.Refresh
        .Parameters("@pRowId") = lRowId
        .Parameters("@pNewPath") = sNewPath
        .Parameters("@pNewLink") = sImgLnk
    
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was a problem updating the db with the payer sub folder name", "RowID: " & CStr(lRowId) & " To: " & sNewPath
            GoTo Block_Exit
        End If
    End With
    
    MovePayerReferenceToPayerSubFldr = True
    
Block_Exit:
    Set oAdo = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    MovePayerReferenceToPayerSubFldr = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function ConvertConceptReference(sConceptId As String, lRowId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sSql As String


    strProcName = ClassName & ".ConvertConceptReference"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_ConceptReferences_Insert"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pRowId") = lRowId

        Call .Execute
        If .Parameters("@pErrMsg").Value <> "" Then
            LogMessage strProcName, "ERROR", "An error occurred trying to insert records with " & .sqlString, .Parameters("@pErrMsg").Value, True, sConceptId
            GoTo Block_Exit
        End If
        
    End With
    
    ConvertConceptReference = True
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    ConvertConceptReference = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' This creates a COPY of the rowid in _CLAIMS Concept_References and then inserts into the ERAC tables
'''
Public Function CopyAttachTypeFromRowId(lPayerNameId As Long, sConceptId As String, lRowId As Long, _
            Optional lNewRowId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sSql As String

    strProcName = ClassName & ".CopyAttachTypeFromRowId"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_ConceptReferences_CopyForOtherPayer"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pRowId") = lRowId
        .Parameters("@pPayerNameid") = lPayerNameId

        Call .Execute
        If .Parameters("@pErrMsg").Value <> "" Then
            LogMessage strProcName, "ERROR", "An error occurred trying to insert records with " & .sqlString, .Parameters("@pErrMsg").Value, True, sConceptId
            GoTo Block_Exit
        End If
        lNewRowId = .Parameters("@pNewRowId").Value
    End With
    

    CopyAttachTypeFromRowId = True
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    CopyAttachTypeFromRowId = False
    GoTo Block_Exit
End Function





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function NewConceptBasedOnExisting(sConceptId As String, Optional sOutMsg As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".ValidateConcept"

        
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_Concept_Duplicate"
        .Parameters.Refresh
        .Parameters("@pOldConceptID") = sConceptId
        .Execute
        If CStr("" & .Parameters("@pErrMsg").Value) <> "" Then
            LogMessage strProcName, "ERROR", "There was a problem duplicating this concept!", , True, sConceptId
            GoTo Block_Exit
        End If
        NewConceptBasedOnExisting = CStr("" & .Parameters("@pNewConceptID").Value)
    End With


Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    NewConceptBasedOnExisting = ""
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' Returns True if the concept is ready for submission
''' false otherwise..
''' sReport will contain a vbCrlf delimited record of what's wrong with it
'''
Public Function ValidateConcept(sConceptId As String, Optional sReport As String) As Boolean
    Stop    ' why are we still calling this?
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oConcept As clsConcept
'Dim dSubmitDate As Date
'Dim sOutMessage As String
'
'    strProcName = ClassName & ".ValidateConcept"
'' does this ever get called? I don't think so
'    If sConceptId = "" Then
'        sReport = "Concept ID not set"
'        GoTo Block_Exit
'    End If
'
'    Set oConcept = New clsConcept
'    If oConcept.LoadFromID(sConceptId) = False Then
'        sReport = "Could not load Concept ID Object"
'        GoTo Block_Exit
'    End If
'
'    '' Does it have the ClientIssueID
'    If oConcept.ClientIssueId = "" Then
'        sReport = sReport & "Concept does not have a client issue id yet" & vbCrLf
'    End If
'
'    '' Does it have all of the required fields in CONCEPT_hdr?
'    If oConcept.HasRequiredFields(sOutMessage) = False Then
'        sReport = sReport & "Missing required fields: " & vbCrLf & sOutMessage & vbCrLf
'        sOutMessage = ""
'    End If
'
'    '' Does it have all of the required documents attached?
'    sOutMessage = oConcept.GetMissingRequiredDocsMessage()
'    If sOutMessage <> "" Then
'        sReport = sReport & "Missing required documents: " & sOutMessage
'    End If
'    sOutMessage = ""
'
'    '' Has it been submitted yet?
'    If WasConceptSubmitted(oConcept.ConceptID) = True Then
'        sReport = sReport & "Concept was already submitted" & vbCrLf
'    End If
'    sOutMessage = ""
'
'    dSubmitDate = oConcept.AlreadySubmitted(sOutMessage)
'    If dSubmitDate <> CDate("1/1/1900") Then
'        sReport = sReport & "Concept was submitted on " & dSubmitDate & vbCrLf & sOutMessage & vbCrLf
'        '' Don't need to continue now do we?
'    End If
'    sOutMessage = ""
'
'    '' Is the notification status 3 (or 7)
'    If oConcept.IsStatusOkForEracSubmission(sOutMessage) = False Then
'        sReport = sReport & sOutMessage & vbCrLf
'    End If


'Block_Exit:
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    Err.Clear
'    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function WasConceptSubmitted(sConceptId As String, Optional iPayerNameId As Integer) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim dAlreadySubmitted As Date

    strProcName = ClassName & ".WasConceptSubmitted"
    If sConceptId = "" Then GoTo Block_Exit

    dAlreadySubmitted = CDate("1/1/1900")
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_EracConceptCurrentStatus"
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pPayerNameId") = iPayerNameId
        Set oRs = .ExecuteRS()

        If .GotData = False Then
            WasConceptSubmitted = False
            GoTo Block_Exit
        End If
    
    End With
    
    dAlreadySubmitted = Nz(oRs("SubmitDate"), CDate("1/1/1900"))
    WasConceptSubmitted = IIf(dAlreadySubmitted = CDate("1/1/1900"), False, True)


Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Returns true if some were added OR removed..
''' sActionReport will contain details about what happened if supplied
Public Function SynchTaggedClaims(sConceptId As String, Optional ByRef sActionReport As String, Optional sPayerNameId As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sSql As String
Dim colRemoveClaims As Collection
Dim oCmd As ADODB.Command
Dim oRs As ADODB.RecordSet
Dim oParams As ADODB.Parameters
Dim oParam As ADODB.Parameter
Dim bTransStarted As Boolean
Dim iTranCount As Integer
Dim iRemoveCount As Integer
Dim iAddCount As Integer
Dim iDocRemoveCount As Integer
Dim vCnlyClaimNum As Variant
Dim sRemovedFiles As String
    ' NOTE: We are not going to do this at the payer level, just do it all at the concept level..
        ' they have the cmbPayer filter box to see what they want to see
    strProcName = ClassName & ".SynchTaggedClaims"
    If sConceptId = "" Then
        SynchTaggedClaims = False
        GoTo Block_Exit
    End If
    
    Set colRemoveClaims = New Collection
    
    '' Make sure that this concept hasn't been submitted first
        '' KDCOMEBACK: DO THIS PART!
    
    start   '' set hourglass, etc..
    LogMessage strProcName, , "Starting 'Synch' operation for concept: " & sConceptId, , , sConceptId
    
        ' Now, check to see if any have been untagged
    DoCmd.Echo True, "Checking for untagged claims"
        ' because if so, we'll have to remove any documents that were linked to
        ' those tagged claims
        ' SP: usp_EracTaggedClaimsRemoved
    Set oAdo = New clsADO
    Set oCmd = New ADODB.Command
    
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_EracRemoveUnTaggedClaims"
        
        oCmd.ActiveConnection = .CurrentConnection
        oCmd.commandType = adCmdStoredProc
        oCmd.CommandText = .sqlString
        oCmd.Parameters.Refresh
        
        oCmd.Parameters("@pConceptId") = sConceptId
'        oCmd.Parameters("@pPayerNameId") = sPayerNameId
        
'        bTransStarted = True
'        iTranCount = iTranCount + 1
'        .BeginTrans
    
        LogMessage strProcName, , "Executing " & .sqlString, , , sConceptId

        Set oRs = .ExecuteRS(oCmd.Parameters)
        
        If Not oRs Is Nothing Then
            If oRs.recordCount > 0 Then
                sActionReport = "The following ICN's have been un-tagged, therefore removed from eRAC submission system" & vbCrLf
                
                While Not oRs.EOF
                    colRemoveClaims.Add Nz(oRs("ICN").Value, ""), Nz(oRs("CnlyClaimNum").Value, "")
                    sActionReport = sActionReport & " - " & Nz(oRs("ICN").Value, "") & vbCrLf
                    oRs.MoveNext
                Wend
            End If
        End If
        iRemoveCount = colRemoveClaims.Count
        LogMessage strProcName, , CStr(iRemoveCount) & " claims have been un-tagged", , , sConceptId

    End With
    
    DoCmd.Echo True, "Adding newly tagged claims"

    ' Then, just add any that aren't already here
    With oAdo
        .SQLTextType = StoredProc
        .sqlString = "usp_EracGetTaggedClaimsNotInErac"
        Set oCmd = New ADODB.Command
        
        oCmd.ActiveConnection = .CurrentConnection
        oCmd.commandType = adCmdStoredProc
        oCmd.CommandText = .sqlString

        oCmd.Parameters.Refresh
        
        oCmd.Parameters("@pConceptId") = sConceptId
        
'        bTransStarted = True
'        iTranCount = iTranCount + 1
'        .BeginTrans
        
        LogMessage strProcName, , "Executing " & .sqlString, , , sConceptId
        Set oRs = .ExecuteRS(oCmd.Parameters)

        If Not oRs Is Nothing Then
            If oRs.recordCount > 0 Then
                sActionReport = sActionReport & vbCrLf & "The following newly tagged ICN's have been added: " & vbCrLf
            
                While Not oRs.EOF
                    sActionReport = sActionReport & " - " & Nz(oRs("ICN").Value, "") & vbCrLf
                    oRs.MoveNext
                Wend
            End If
        End If
        iAddCount = oRs.recordCount
        LogMessage strProcName, , CStr(iAddCount) & " claims have been newly tagged", , , sConceptId
    End With
    

        '' Now, if we removed any, then we need to remove the documents attached to those (if any)
    If colRemoveClaims.Count > 0 Then
        DoCmd.Echo True, "Checking for documents linked to un-tagged claims and removing them"

        LogMessage strProcName, , "Checking for attached document for un-tagged claims", , , sConceptId
        
        For Each vCnlyClaimNum In colRemoveClaims
            iDocRemoveCount = iDocRemoveCount + RemoveAttachedDocument(sConceptId, CStr("" & vCnlyClaimNum), sConceptId, sRemovedFiles)
        Next
        LogMessage strProcName, , CStr(iDocRemoveCount) & " files have been removed due to un-tagged claims", sRemovedFiles, , sConceptId
        
        If Trim(sRemovedFiles) <> "" Then
            sActionReport = sActionReport & vbCrLf & "The following files associated with removed tagged claims have been deleted: " & vbCrLf & sRemovedFiles
        End If
    End If

        '' And finally, commit any transactions
'    If bTransStarted = True Then
'        While iTranCount > 0
'            oAdo.CommitTrans
'            iTranCount = iTranCount - 1
'        Wend
'    End If
    LogMessage strProcName, , "Finished", , , sConceptId
    
Block_Exit:
    Done
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    sActionReport = sActionReport & vbCrLf & vbCrLf & "ERROR: " & Err.Description & vbCrLf
    
        '' Rollback any transactions we started.
'    If bTransStarted = True Then
'        While iTranCount > 0
'            oAdo.RollbackTrans
'            iTranCount = iTranCount - 1
'        Wend
'    End If
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetFieldListFromTableList(strTableList As String) As Scripting.Dictionary
On Error GoTo Block_Err
Dim strProcName As String
Dim dctFields As Scripting.Dictionary
Dim oTDef As DAO.TableDef
Dim oDb As DAO.Database
Dim sTableList() As String
Dim iIdx As Integer
Dim sCurTableName As String
Dim sTextual As String
Dim oFld As DAO.Field
Dim sFieldName As String


    strProcName = ClassName & ".GetFieldListFromTableList"
    Set dctFields = New Scripting.Dictionary
    
    Set oDb = CurrentDb()
    sTableList = Split(strTableList, ",")
    
    
    For iIdx = 0 To UBound(sTableList)
            ' remove brackets if found
        sCurTableName = Replace(sTableList(iIdx), "[", "")
        sCurTableName = Replace(sCurTableName, "]", "")
        
            ' if it's a table, get it:
        If IsTable(sCurTableName) = False Then
            GoTo NextTable
        End If
        
        Set oTDef = oDb.TableDefs(sCurTableName)
        
        For Each oFld In oTDef.Fields
        
            '' Take brackets off before I put them on.. :D
            sFieldName = Replace(oFld.Name, "[", "")
            sFieldName = Replace(sFieldName, "]", "")
            
            sFieldName = "[" & sFieldName & "]"
            
            '' Ok, actually, we need the table name too:
            sFieldName = "[" & sCurTableName & "]." & sFieldName
            
            ' Determine if we should treat it as a textual field:
            Select Case oFld.Type
            Case dbChar, dbDate, dbGUID, dbMemo, dbText, dbTime, dbTimeStamp
                sTextual = "'"
            Case dbBigInt, dbBoolean, dbCurrency, dbDecimal, dbDouble, dbFloat, dbInteger, dbLong, dbNumeric, dbSingle
                sTextual = " "
            Case Else
                sTextual = ""   ' skip it
            End Select
            
            If sTextual <> "" Then
                If dctFields.Exists(sFieldName) Then
                    '' hmm.. is it different?
                    '' What should we do?
                    '' nothing for now..
'                    Stop ' should never happen cause we are including the table name too..
                Else
                    dctFields.Add sFieldName, sTextual
                End If
            End If

        Next
        
NextTable:
    Next
    


Block_Exit:
    Set oFld = Nothing
    Set oTDef = Nothing
    Set oDb = Nothing
    Set GetFieldListFromTableList = dctFields
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    Set GetFieldListFromTableList = Nothing
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function MakeWhereListFromTableList(strTableList As String, sValToFind As String, _
            Optional bNumericSearch As Boolean = False, Optional bForAdo As Boolean = True) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim dctFields As Scripting.Dictionary
Dim iIdx As Integer
Dim sCurTableName As String
Dim sTextual As String
Dim sFieldName As String
Dim vDctKey As Variant
Dim sRet As String
Dim sWCardChar As String

    strProcName = ClassName & ".MakeWhereListFromTableList"
    If strTableList = "" Then GoTo Block_Exit
    
    Set dctFields = GetFieldListFromTableList(strTableList)
    
    If dctFields.Count < 1 Then
        GoTo Block_Exit
    End If
    
    If bForAdo = True Then
        sWCardChar = "%"
    Else
        sWCardChar = "*"
    End If
    
    sRet = "("
    
    '' Each one should look like: [tablename].[fieldname] LIKE '*VALUE*' OR
    
    For Each vDctKey In dctFields
        If bNumericSearch = True Then
            If dctFields.Item(vDctKey) = " " Then
                sRet = sRet & CStr("" & vDctKey) & " = " & sValToFind & " OR "
            End If
        Else
            ' So, it's NOT a numeric search so, we can do something like
            ' fieldname like '*34*'
            '' If we start getting too many false positives (likely) then uncomment the conditional
'            If dctFields.Item(vDctKey) <> " " Then
                sRet = sRet & CStr("" & vDctKey) & " LIKE '" & sWCardChar & _
                        sValToFind & sWCardChar & "' OR "
'            End If
        End If
    Next
    
    '' Remove the final OR
    sRet = left(sRet, Len(sRet) - 3)
    
    sRet = sRet & ")"
    
    
Block_Exit:
    MakeWhereListFromTableList = sRet
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    MakeWhereListFromTableList = ""
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function RemoveAttachedDocument(sConceptId As String, Optional sCnlyClaimNum As String = "", Optional sIcn As String = "", Optional ByRef sMessage As String = "") As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oCmd As ADODB.Command
Dim iSequence As Integer
Dim iRowID As Integer
Dim sFilePath As String
Dim SFileName As String
Dim iDeleteCount As Integer

    strProcName = ClassName & ".RemoveAttachedDocument"

    ' First get a recordset of all of the attached documents with this Concept (we find the Claim Num later)
    ' Note: needs to be sorted by RefSequence
    
    If sCnlyClaimNum = "" And sIcn = "" Then
        LogMessage strProcName, "WARNING", "No Claim ID sent to function", , , sConceptId
        GoTo Block_Exit
    End If
    
    If sCnlyClaimNum = "" Then sCnlyClaimNum = Chr(0)
    If sIcn = "" Then sIcn = Chr(0)
    
    
    LogMessage strProcName, , "Getting Concept references for " & sConceptId, , , sConceptId, sCnlyClaimNum
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_AllAttachedDocsForConcept"
        
        Set oCmd = New ADODB.Command
        oCmd.ActiveConnection = oAdo.CurrentConnection
        oCmd.commandType = adCmdStoredProc
        oCmd.CommandText = .sqlString
        oCmd.Parameters.Refresh
        
        oCmd.Parameters("@pConceptId") = sConceptId
        LogMessage strProcName, , "Executing " & .sqlString, , , sConceptId, sCnlyClaimNum
        
        Set oRs = .ExecuteRS(oCmd.Parameters)
        If .GotData = False Then
            ' None found, nothing to do
            LogMessage strProcName, , "No references found", , , sConceptId, sCnlyClaimNum
            GoTo Block_Exit
        End If
        
    End With
    
    
    ' Then, if we find our item
    Do While Not oRs.EOF
        LogMessage strProcName, , "Looking for ICN: " & sIcn & " or CnlyClaimNum: " & sCnlyClaimNum, , , sConceptId, sCnlyClaimNum
        
        If CStr("" & oRs("CnlyClaimNum").Value) = sCnlyClaimNum _
                    Or CStr("" & oRs("ICN").Value) = sIcn Then
            
            LogMessage strProcName, , "Found an attachment: RowId: " & CStr(oRs("RowId").Value), , , sConceptId, sCnlyClaimNum
            
            ' Save some data: RefSequence, refLink (filepath), RefFileName
            iSequence = Nz(oRs("RefSequence").Value, -1)
            sFilePath = Nz(oRs("RefLink").Value, "")
            SFileName = Nz(oRs("RefFileName").Value, "")
            iRowID = Nz(oRs("RowId").Value, -1)
            sIcn = Nz(oRs("ICN").Value, "")
            
            If iRowID < 1 Then GoTo Block_Exit
            
            ' Delete the record from _CLAIMS.dbo.CONCEPT_References
            If DeleteRowFromReferenceTbl(sConceptId, iSequence, iRowID) = False Then
                LogMessage strProcName, "WARNING", "Unable to delete row from Reference table: " & CStr(iRowID), , , sConceptId, sCnlyClaimNum
                GoTo Block_Exit
            End If
            
            ' Now, if we have the document, delete it:
            If FileExists(sFilePath) = True Then
                LogMessage strProcName, , "Found existing linked file", sFilePath, , sConceptId, sCnlyClaimNum
                
                If DeleteFile(sFilePath, False) = True Then
                    iDeleteCount = iDeleteCount + 1
                    sMessage = sMessage & vbCrLf & " - " & sFilePath
                Else
                    sMessage = sMessage & vbCrLf & " - a file was found, but we were unable to delete it: " & vbCrLf & "    " & sFilePath
                    LogMessage strProcName, "WARNING", "Delete File returned false", sFilePath, , sConceptId, sCnlyClaimNum
                End If
            End If
            
            
            Exit Do
        End If
        oRs.MoveNext
    Loop


Block_Exit:
    Set oCmd = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    RemoveAttachedDocument = iDeleteCount
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId, sCnlyClaimNum
    Err.Clear
    RemoveAttachedDocument = 0
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function DeleteRowFromReferenceTbl(sConceptId As String, intSequence As Integer, intRowID As Integer) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim iResult As Integer
Dim oCmd As ADODB.Command
    
    strProcName = ClassName & ".DeleteRowFromReferenceTbl"

    Set oAdo = New clsADO
    oAdo.SQLTextType = StoredProc
    oAdo.ConnectionString = GetConnectString("v_CODE_Database")
    oAdo.sqlString = "usp_CONCEPT_References_Delete"
    
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = "usp_CONCEPT_References_Delete"
    oCmd.Parameters.Refresh
    
    oCmd.Parameters("@pRowID") = intRowID
    iResult = oAdo.Execute(oCmd.Parameters)
    
        '' At this point we should update the rest of the sequences
    Set oAdo = New clsADO
    oAdo.ConnectionString = GetConnectString("CONCEPT_References")
    oAdo.SQLTextType = StoredProc
    oAdo.sqlString = "usp_UpdateSequenceInConceptReferencesTbl"
    
    Set oCmd = New ADODB.Command
    
    oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = oAdo.sqlString
    oCmd.Parameters.Refresh
    oCmd.Parameters("@pConceptId") = sConceptId
    oCmd.Parameters("@pSequenceNum") = intSequence
    
    If oAdo.Execute() < 0 Then
        ' trouble
        DeleteRowFromReferenceTbl = False
        GoTo Block_Exit
    End If
    DeleteRowFromReferenceTbl = True

Block_Exit:
    Set oCmd = Nothing
    Set oAdo = Nothing
    Exit Function

Block_Err:
    'Return false so the whole save event can rollback.
    DeleteRowFromReferenceTbl = False
    GoTo Block_Exit
End Function





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' This simply queries the AuditClm_Dtl (and _Hdr) table for how many claims
''' it finds for that concept id.  After the ERAC update this should point to the
''' _ERAC table for the tagged claims
'''
Public Function HowManyClaimsForGivenConcept(sConceptId As String) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim sDtlSql As String
Dim sHdrSql As String
Dim cAdo As clsADO
Dim rs As ADODB.RecordSet
Dim sConnString As String
Dim iClaimsFound As Integer


    strProcName = ClassName & ".HowManyClaimsForGivenConcept"
    
    If sConceptId = "" Then GoTo Block_Exit
    
    sConnString = GetConnectString("AUDITCLM_Dtl")
    
    ' Detail claims
    sDtlSql = "select * from CMS_AUDITORS_CLAIMS.dbo.AUDITCLM_Dtl d WHERE D.Adj_ConceptID = '" & sConceptId & "'"
    
    ' header claims
    sHdrSql = "select * from CMS_AUDITORS_CLAIMS.dbo.AuditClm_Hdr d WHERE D.Adj_ConceptID = '" & sConceptId & "'"
    
    Set cAdo = New clsADO
    cAdo.ConnectionString = sConnString
    cAdo.Connect
    
    Set rs = cAdo.OpenRecordSet(sDtlSql)
    iClaimsFound = iClaimsFound + rs.recordCount
    
    Set rs = cAdo.OpenRecordSet(sHdrSql)
    iClaimsFound = iClaimsFound + rs.recordCount
    

Block_Exit:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    Set cAdo = Nothing
    HowManyClaimsForGivenConcept = iClaimsFound
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    iClaimsFound = -1   ' indicate error
    GoTo Block_Exit     ' I don't like resume's because sometimes we get the ' Resume without error
End Function






''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' This simply queries the AuditClm_Dtl (and _Hdr) table for how many claims
''' it finds for that concept id.  After the ERAC update this should point to the
''' _ERAC table for the tagged claims
'''
Public Function GetTaggedClaimsForConcept(sConceptId As String) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim sDtlSql As String
Dim sHdrSql As String
Dim cAdo As clsADO
Dim rs As ADODB.RecordSet
Dim sConnString As String
Dim iClaimsFound As Integer


    strProcName = ClassName & ".HowManyClaimsForGivenConcept"
    
    If sConceptId = "" Then GoTo Block_Exit
    
    sConnString = GetConnectString("AUDITCLM_Dtl")
    
    ' Detail claims
    sDtlSql = "select * from CMS_AUDITORS_CLAIMS.dbo.AUDITCLM_Dtl d WHERE D.Adj_ConceptID = '" & sConceptId & "'"
    
    ' header claims
    sHdrSql = "select * from CMS_AUDITORS_CLAIMS.dbo.AuditClm_Hdr d WHERE D.Adj_ConceptID = '" & sConceptId & "'"
    
    Set cAdo = New clsADO
    cAdo.ConnectionString = sConnString
    cAdo.Connect
    
    Set rs = cAdo.OpenRecordSet(sDtlSql)
    iClaimsFound = iClaimsFound + rs.recordCount
    
    Set rs = cAdo.OpenRecordSet(sHdrSql)
    iClaimsFound = iClaimsFound + rs.recordCount
    

Block_Exit:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    Set cAdo = Nothing
    GetTaggedClaimsForConcept = iClaimsFound
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    iClaimsFound = -1   ' indicate error
    GoTo Block_Exit
End Function






''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' Desc: Use this function to log an entry to the Concept History table
'''  Returns the ID from the table (unless updating)
'''  To insert: specify ConceptId, action type, and optionally action result, and or notes
'''  To update with action result: specify: ConceptId, ActionResult and IdToUpdate (
'''     and optionally notes which will be appended - NOT replaced!)
'''
Public Function LogActionToHistory(sConceptId As String, lPayerNameId As Long, lActionType As enuConceptActions, _
    Optional sActionResult As String = "", Optional sNotes As String = "", _
    Optional lIdToUpdate As Long = 0, Optional sPackageId As String = "", Optional sNotificationId As String = "", _
    Optional sConceptOwner As String = "") As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim sUser As String
Dim oCn As ADODB.Connection
Dim rs As ADODB.RecordSet
Dim sConnString As String
Dim oCmd As ADODB.Command
Dim sMsg As String
Dim iIndx As Integer

    strProcName = ClassName & ".LogActionToHistory"
    
    sUser = Replace(Identity.UserName(), "CCA-AUDIT\", "")
    
    ' Always need concept id..
    If sConceptId = "" Then GoTo Block_Exit

    Set oCn = New ADODB.Connection
    Set oCmd = New ADODB.Command
    
    oCn.ConnectionString = GetConnectString("ConceptDocTypes")
    oCn.Open
    
    '' We'll first do the raw sql then we'll create a usp for it..
    

    With oCmd
        .ActiveConnection = oCn
        .commandType = adCmdStoredProc
        .CommandText = "usp_LogConceptHistory"
        .Parameters.Refresh
        
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pPayerNameId") = lPayerNameId
        .Parameters("@pEracActionId") = lActionType
        
        If sActionResult <> "" Then
            .Parameters("@pActionResult") = IIf(sActionResult = "", vbNull, sActionResult)
        End If
        
        If lIdToUpdate > 0 Then
            .Parameters("@pIdToUpdate") = lIdToUpdate
        End If
        If sPackageId <> "" Then
            .Parameters("@pPackageId") = left(sPackageId, 12)
        End If
        
        If sNotificationId <> "" Then
            .Parameters("@pNotificationId") = sNotificationId
        End If
        
        If sNotes <> "" Then
            .Parameters("@pNotes") = sNotes
        End If
            
            '' .Parameters("@pNewId") This is an output param..
    End With

    Set rs = oCmd.Execute
    
        '' Any errors?
    If oCn.Errors.Count > 0 Then
        For iIndx = 0 To oCn.Errors.Count - 1
            sMsg = sMsg & oCn.Errors(iIndx) & vbCrLf
        Next
        LogMessage strProcName, "ERROR", sMsg, sConceptId, , sConceptId
        GoTo Block_Exit
    End If
    
    If Nz(oCmd.Parameters("@pNewId"), -1) < 0 Then
        sMsg = "There was an error with the query: " & oCmd.CommandText
        LogMessage strProcName, "ERROR", sMsg, oCmd.Parameters("@pErrMsg"), True, sConceptId
        GoTo Block_Exit
    Else
        LogActionToHistory = oCmd.Parameters("@pNewId")
    End If

Block_Exit:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    Set oCmd = Nothing
    Set oCn = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    LogActionToHistory = -1
    GoTo Block_Exit
End Function

    

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' This simply queries the ConceptPackageId table to retrieve the username, assuming
''' that is who "owns" it
'''
Public Function GetConceptOwner(sConceptId As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sReturn As String


    strProcName = ClassName & ".GetConceptOwner"

    sSql = "select [UserId] FROM ConceptPackageId WHERE ConceptId = '" & sConceptId & "'"
    
    Set oAdo = New clsADO
    With oAdo
'        .ConnectionString = GetConnectString("ConceptPackageId")
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = sqltext
        .sqlString = sSql
        .Connect
        Set oRs = .ExecuteRS()
    End With
    
    If oRs.recordCount > 0 Then
        sReturn = oRs("UserId").Value
    End If
    
    If sReturn = "" Then sReturn = "{Unkown}"

Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    GetConceptOwner = sReturn
    Exit Function
    
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Public Function IsEditableField(strTableName As String, strFieldName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".IsEditableField"

    If gdctTableIndexes Is Nothing Then
        If GetTableIndicies = False Then
'            LogMessage "Cannot tell if field is editable or not: " & strTableName & "." & strFieldName, strProcName, "ERROR", strTableName & "." & strFieldName
            IsEditableField = False
            GoTo Block_Exit
        End If
    End If

    IsEditableField = IIf(gdctTableIndexes.Exists(UCase(strTableName) & "." & UCase(strFieldName)), False, True)

    If IsEditableField = False Then
'        LogMessage "Field is NOT editable: " & strTableName & "." & strFieldName, strProcName, "DEBUG"
        Debug.Print ""
    End If

Block_Exit:
    Exit Function
    
Block_Err:
    IsEditableField = False
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Private Function GetTableIndicies() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oTDef As DAO.TableDef
Dim oTIndex As DAO.index
Dim oFld As DAO.Field

    strProcName = ClassName & ".GetTableIndicies"
    
    If mblnSetup = True Then
        GetTableIndicies = True
        GoTo Block_Exit
    End If
    

    Set gdctTableIndexes = New Scripting.Dictionary
    
    Set oDb = CurrentDb()
    For Each oTDef In oDb.TableDefs
        If left(oTDef.Name, 4) <> "MSys" Then
            For Each oTIndex In oTDef.Indexes
                If oTIndex.Unique = True Then
                    For Each oFld In oTIndex.Fields
                        If gdctTableIndexes.Exists(UCase(oTDef.Name) & "." & UCase(oFld.Name)) = False Then
                            gdctTableIndexes.Add UCase(oTDef.Name) & "." & UCase(oFld.Name), True
                        End If
                    Next
                End If
            Next
        End If
    Next

Block_Exit:
    Set oTIndex = Nothing
    Set oTDef = Nothing
    Set oDb = Nothing
    Exit Function

Block_Err:
    ReportError Err, strProcName
    Resume Next
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' This function performs (or orchastrates) all validation  on a particular concept
''' to see if it's ready to submit to CMS
'''
Public Function ConceptReadyForSubmission(sConceptId As String, Optional sMissingDetails As String = "") As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oEracDoc As clsConceptReqDocType
Dim oEracReqRule As clsEracRequirementRule
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sCurrentSql As String
Dim sConceptOwner As String
Dim sCnlyReviewType As String
Dim sCnlyDataType As String
Dim iReviewTypeId As Integer
Dim iDataTypeId As Integer
Dim sClientIssueNum As String
Dim sNotificationId As String
Dim sMsg As String

    strProcName = ClassName & ".ConceptReadyForSubmission"

    Set oAdo = New clsADO
    
    '' Let's get the details: reviewtype, data type, owner...
    
'    sCurrentSql = "SELECT CH.Auditor ConceptOwner, CH.ClientIssueNum, CH.ReviewType, CH.DataType FROM " & _
'        " CMS_AUDITORS_CLAIMS.dbo.Concept_Hdr CH WHERE Ch.ConceptID = '" & sConceptId & "'"
'
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = sqltext
        .sqlString = sCurrentSql
    End With
'
'    Set oRs = oAdo.ExecuteRS
'
'
'    '' Did we get anything?
'    If oRs Is Nothing Then
'        ConceptReadyForSubmission = False
'        sMsg = "Concept not found in Concept Header, please verify and try again"
'        GoTo Block_Exit
'    End If
'
'    If oRs.RecordCount < 1 Then
'        ConceptReadyForSubmission = False
'        sMsg = "Concept not found in Concept Header, please verify and try again"
'        GoTo Block_Exit
'    End If
'
'
'    If Not oRs.EOF Then
'        sConceptOwner = Nz(oRs("ConceptOwner").Value, "")
'        sCnlyReviewType = Nz(oRs("ReviewType").Value, "")
'        sCnlyDataType = Nz(oRs("DataType").Value, "")
'        ' We don't expect more than 1, so no need to movenext
'    End If
    
    If GetConceptHeaderDetails(sConceptId, sCnlyReviewType, sCnlyDataType, sConceptOwner) = False Then
        'Stop
        ' KD COMEBACK: What's this?
    End If
    
    '' Now get our requirement rules..
    iReviewTypeId = TranslateCnlyReviewTypeToCMS(sCnlyReviewType)
    If sCnlyDataType = "DME" Then
        sCurrentSql = "SELECT * FROM v_ConceptRequirements WHERE ReviewTypeID = " & CStr(iReviewTypeId) & " AND DataTypeCode = 'DME' "
    Else
        sCurrentSql = "SELECT * FROM v_ConceptRequirements WHERE ReviewTypeID = " & CStr(iReviewTypeId) & " AND ISNULL(DataTypeCode,'') = '' "
    End If
    
    oAdo.sqlString = sCurrentSql
    Set oRs = oAdo.ExecuteRS
    
    While Not oRs.EOF
        Set oEracReqRule = New clsEracRequirementRule
        If oEracReqRule.LoadFromRS(oRs) = False Then
            ' KD COMEBACK: What's this?
        End If
    
        ''
    
        oRs.MoveNext
    Wend
    
    


Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    ConceptReadyForSubmission = False
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' This function performs (or orchastrates) all validation  on a particular concept
''' to see if it's ready to submit to CMS
'''
Public Function GetConceptHeaderDetails(sConceptId As String, Optional ByRef sReviewTypeCode As String = "", _
    Optional ByRef sDataTypeCode As String = "", Optional ByRef sConceptOwner As String = "", _
    Optional ByRef sClientIssueNum As String = "", Optional iCmsReviewTypeId As Integer = 0) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet


    strProcName = ClassName & ".GetConceptReviewType"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = sqltext
        .sqlString = "SELECT CH.Auditor ConceptOwner, CH.ClientIssueNum, CH.ReviewType, CH.DataType FROM " & _
            " CMS_AUDITORS_CLAIMS.dbo.Concept_Hdr CH WHERE Ch.ConceptID = '" & sConceptId & "'"
    End With

   
    Set oRs = oAdo.ExecuteRS
    
    
    '' Did we get anything?
    If oRs Is Nothing Then
        LogMessage strProcName, "WARNING", "Concept not found in Concept Header, please verify and try again", , , sConceptId
        GoTo Block_Exit
    End If
    
    If oRs.recordCount < 1 Then
        LogMessage strProcName, "WARNING", "Concept not found in Concept Header, please verify and try again", , , sConceptId
        GoTo Block_Exit
    End If
    
    If Not oRs.EOF Then
        sConceptOwner = Nz(oRs("ConceptOwner").Value, "")
        sReviewTypeCode = Nz(oRs("ReviewType").Value, "")
        sDataTypeCode = Nz(oRs("DataType").Value, "")
        sClientIssueNum = Nz(oRs("ClientIssueNum").Value, "")
        ' We don't expect more than 1, so no need to movenext
    End If

    iCmsReviewTypeId = TranslateCnlyReviewTypeToCMS(sReviewTypeCode)

    GetConceptHeaderDetails = True

Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GetConceptHeaderDetails = False
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' This simply queries the _ERAC.dbo.XrefReviewType table to convert
''' Connolly's review type code to CMS' numeric value
'''
Public Function TranslateCnlyReviewTypeToCMS(sCnlyReviewTypeCode As String) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

'' TEMPORARY!
If sCnlyReviewTypeCode = "A" Then
    TranslateCnlyReviewTypeToCMS = 0
Else
    TranslateCnlyReviewTypeToCMS = 1
End If
GoTo Block_Exit


    strProcName = ClassName & ".TranslateCnlyReviewTypeToCMS"
    TranslateCnlyReviewTypeToCMS = -1

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = sqltext
        .sqlString = "SELECT ReviewTypeId FROM CMS_AUDITORS_ERAC.dbo.XrefReviewType WHERE CnlyReviewType = '" & Trim(sCnlyReviewTypeCode) & "'"
        Set oRs = .ExecuteRS()
    End With

    If oRs Is Nothing Then GoTo Block_Exit
    If oRs.recordCount < 1 Then GoTo Block_Exit
    
    TranslateCnlyReviewTypeToCMS = oRs("ReviewTypeId")

Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
Public Function EracReviewTypeFromCnlyCode(sCnlyReviewType As String) As Integer
On Error GoTo Block_Err
Static dctReviewTypeIds As Scripting.Dictionary
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".EracReviewTypeFromCnlyCode"

    If dctReviewTypeIds Is Nothing Then
        '' Cache the info as it doesn't change hardly ever
        '' (i.e. we populate our static dictionary which hangs around until recompile time
        ''  or until the DB is opened again)
        Set dctReviewTypeIds = New Scripting.Dictionary
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("ConceptDocTypes")
            .SQLTextType = sqltext
            .sqlString = "SELECT ReviewTypeID, CnlyReviewTypeCode FROM XrefReviewType " ' WHERE ReviewTypeID = " & Me.ReviewTypeId
            Set oRs = .ExecuteRS
            If .GotData = False Then
                ' error, so return nothing
                EracReviewTypeFromCnlyCode = -1
                GoTo Block_Exit
            End If
        End With
        
        While Not oRs.EOF
            dctReviewTypeIds.Add CStr("" & oRs("CnlyReviewTypeCode").Value), oRs("ReviewTypeID").Value
            oRs.MoveNext
        Wend
        
    End If
    
    If dctReviewTypeIds.Exists(sCnlyReviewType) Then
        EracReviewTypeFromCnlyCode = dctReviewTypeIds.Item(sCnlyReviewType)
    Else
        EracReviewTypeFromCnlyCode = -1
    End If
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If

    Set oAdo = Nothing
    Exit Function
Block_Err:
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
Public Function EracDataTypeDesc(sCnlyDataType As String) As String
On Error GoTo Block_Err
Static dctDataTypeIds As Scripting.Dictionary
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".EracDataTypeDesc"

    If dctDataTypeIds Is Nothing Then
        '' Cache the info as it doesn't change hardly ever
        '' (i.e. we populate our static dictionary which hangs around until recompile time
        ''  or until the DB is opened again)
        Set dctDataTypeIds = New Scripting.Dictionary
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("XREF_ReviewType")
            .SQLTextType = sqltext
            .sqlString = "SELECT DataType, DataTypeDesc FROM XREF_DataType WHERE DataType = '" & sCnlyDataType & "'"
            Set oRs = .ExecuteRS
            If .GotData = False Then
                ' error, so return nothing
                EracDataTypeDesc = ""
                GoTo Block_Exit
            End If
        End With
        
        While Not oRs.EOF
            dctDataTypeIds.Add CStr("" & oRs("DataType").Value), oRs("DataTypeDesc").Value
            oRs.MoveNext
        Wend
        
    End If
    
    If dctDataTypeIds.Exists(sCnlyDataType) Then
        EracDataTypeDesc = CStr("" & dctDataTypeIds.Item(sCnlyDataType))
    Else
        EracDataTypeDesc = ""
    End If
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
Public Function EracReviewTypeNameFromCnlyCode(sCnlyReviewType As String) As String
On Error GoTo Block_Err
Static dctReviewTypeIds As Scripting.Dictionary
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".EracReviewTypeNameFromCnlyCode"

    If dctReviewTypeIds Is Nothing Then
        '' Cache the info as it doesn't change hardly ever
        '' (i.e. we populate our static dictionary which hangs around until recompile time
        ''  or until the DB is opened again)
        Set dctReviewTypeIds = New Scripting.Dictionary
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("ConceptDocTypes")
            .SQLTextType = sqltext
            .sqlString = "SELECT ReviewTypeID, ReviewTypeName, CnlyReviewTypeCode FROM XrefReviewType " ' WHERE ReviewTypeID = " & Me.ReviewTypeId
            Set oRs = .ExecuteRS
            If .GotData = False Then
                ' error, so return nothing
                GoTo Block_Exit
            End If
        End With
        
        While Not oRs.EOF
            dctReviewTypeIds.Add CStr("" & oRs("CnlyReviewTypeCode").Value), oRs("ReviewTypeName").Value
            oRs.MoveNext
        Wend
        
    End If
    
    If dctReviewTypeIds.Exists(sCnlyReviewType) Then
        EracReviewTypeNameFromCnlyCode = dctReviewTypeIds.Item(sCnlyReviewType)
    Else
        EracReviewTypeNameFromCnlyCode = -1
    End If
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
Public Function EracCnlyReviewTypeFromEracId(iEracReviewTypeId As Integer) As String
On Error GoTo Block_Err
Static dctReviewTypeCodes As Scripting.Dictionary
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".EracCnlyReviewTypeFromEracId"

    If dctReviewTypeCodes Is Nothing Then
        '' Cache the info as it doesn't change hardly ever
        '' (i.e. we populate our static dictionary which hangs around until recompile time
        ''  or until the DB is opened again)
        Set dctReviewTypeCodes = New Scripting.Dictionary
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("ConceptDocTypes")
            .SQLTextType = sqltext
            .sqlString = "SELECT ReviewTypeID, CnlyReviewTypeCode FROM XrefReviewType " ' WHERE ReviewTypeID = " & Me.ReviewTypeId
            Set oRs = .ExecuteRS
            If .GotData = False Then
                ' error, so return nothing
                EracCnlyReviewTypeFromEracId = ""
                GoTo Block_Exit
            End If
        End With
        
        While Not oRs.EOF
            dctReviewTypeCodes.Add CStr("" & oRs("ReviewTypeID").Value), CStr("" & oRs("CnlyReviewTypeCode").Value)
            oRs.MoveNext
        Wend
        
    End If
    
    If dctReviewTypeCodes.Exists(CStr("" & iEracReviewTypeId)) Then
        EracCnlyReviewTypeFromEracId = dctReviewTypeCodes.Item(CStr("" & iEracReviewTypeId))
    Else
        EracCnlyReviewTypeFromEracId = "- unknown -"
    End If
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
Public Function GetClaimDetailsFromEracTaggedClaimId(lEracTaggedClaimId As Long) As clsEracClaim
On Error GoTo Block_Err
Dim strProcName As String
Dim oConceptClaim As clsEracClaim

    strProcName = ClassName & ".GetClaimDetailsFromEracTaggedClaimId"

    Set oConceptClaim = New clsEracClaim
    If oConceptClaim.LoadFromTaggedClaimId(lEracTaggedClaimId) = False Then
        LogMessage strProcName, "ERROR", "There was a problem loading the concept claim details", CStr(lEracTaggedClaimId)
        GoTo Block_Exit
    End If

    
Block_Exit:
    Set GetClaimDetailsFromEracTaggedClaimId = oConceptClaim
    Exit Function
Block_Err:
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Returns true if we imported some
''' false if not
'''
Public Function GetTaggedClaimsRS(sConceptId As String, Optional iPayerNameId As Integer = 0) As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oCmd As ADODB.Command
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetTaggedClaimsRS"

    ' see if we have any for this concept..
    If sConceptId = "" Then GoTo Block_Exit
    
    Set oAdo = New clsADO
    Set oCmd = New ADODB.Command
    
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_EracTaggedClaimsForConcept"
        oCmd.ActiveConnection = oAdo.CurrentConnection
        
        oCmd.commandType = adCmdStoredProc
        oCmd.CommandText = .sqlString
        oCmd.Parameters.Refresh
        oCmd.Parameters("@pConceptId") = sConceptId
        Set oRs = .ExecuteRS(oCmd.Parameters)
        If .GotData = False Then
            If oCmd.ActiveConnection.Errors.Count > 0 Then
                LogMessage strProcName, "ERROR", oCmd.ActiveConnection.Errors(0).Description & " " & oCmd.ActiveConnection.Errors(0).Source, , , sConceptId
            End If
            
            GoTo Block_Exit
        End If
    End With
    
    Set GetTaggedClaimsRS = oRs
    
Block_Exit:
    Set oCmd = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    Set GetTaggedClaimsRS = Nothing
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Performs the initial import of tagged claims for this concept - NOT a refresh
''' Returns true if we imported some
''' false if not
'''
Public Function GetTaggedClaimsIfNone(sConceptId As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oCmd As ADODB.Command
Dim oRs As ADODB.RecordSet


    strProcName = ClassName & ".GetTaggedClaimsIfNone"

    ' see if we have any for this concept..
    If sConceptId = "" Then GoTo Block_Exit
    
    Set oAdo = New clsADO
    Set oCmd = New ADODB.Command
    
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_EracTaggedClaimsForConcept"
        
        oCmd.ActiveConnection = oAdo.CurrentConnection
        oCmd.commandType = adCmdStoredProc
        oCmd.CommandText = .sqlString
        oCmd.Parameters.Refresh
        oCmd.Parameters("@pConceptId") = sConceptId
        Set oRs = .ExecuteRS(oCmd.Parameters)
        If .GotData = True Then
            ' Since we got some then we don't want to do anything
            GoTo Block_Exit
        End If
    End With
    
    ' We don't have any in our system so import them

  
    With oAdo
        .SQLTextType = StoredProc
        .sqlString = "usp_EracGetTaggedClaimsNotInErac"
        
        oCmd.ActiveConnection = oAdo.CurrentConnection
        oCmd.commandType = adCmdStoredProc
        oCmd.CommandText = .sqlString
        oCmd.Parameters.Refresh
        oCmd.Parameters("@pConceptId") = sConceptId
        Set oRs = .ExecuteRS(oCmd.Parameters)
        If .GotData = True Then
            GetTaggedClaimsIfNone = True
            GoTo Block_Exit
        End If
    End With

    GetTaggedClaimsIfNone = False
    
Block_Exit:
    Set oCmd = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GetTaggedClaimsIfNone = False
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Creates the GetMedicalRecords files in the concept's send folder
'''
Public Function GetMedicalRecords(sConceptId As String, Optional bDontUpdateSubmitDate As Boolean = False, _
        Optional bOpenExplorer As Boolean = False, Optional bViewPdf As Boolean = True) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oConcept As clsConcept
Dim oAdo As clsADO
Dim oReqDoc As clsConceptReqDocType
Dim sConceptWorkFldr As String
Dim lBatchId As Long
Dim oRs As ADODB.RecordSet
Dim sIcn As String
Dim sToFilename  As String
Dim lJobId As Long
Dim sOutMsg As String
Dim bTimedOutWaiting    As Boolean


    strProcName = ClassName & ".GetMedicalRecords"
    
    Stop ' - not currently being used, but may come back
    
    DoCmd.Hourglass True
    DoCmd.Echo True, "Getting medical records..."
    
    Set oConcept = New clsConcept
    If oConcept.LoadFromId(sConceptId) = False Then
        LogMessage strProcName, "ERROR", "Could not load concept object from id", sConceptId, , sConceptId
        GoTo Block_Exit
    End If
    
    Set oReqDoc = New clsConceptReqDocType
    If oReqDoc.LoadFromId(12) = False Then
'        Stop
    End If
        
        ' Get the files, copy them to the WORK directory
    sConceptWorkFldr = QualifyFldrPath(oConcept.ConceptWorkFolder) & "_BURN\"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_GetMedicalRecordsForConcept"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
'        .Parameters("") = ""
        Set oRs = .ExecuteRS
        If .GotData = False Then
'            Stop
        End If
    End With
    
    
        ' Add a batch
    If AddConverterQueueBatch("", False, oReqDoc.SendAsFileType, sConceptWorkFldr, , False, True, False, , , , lBatchId) = False Then
        LogMessage strProcName, "ERROR", "Could not add a batch of jobs these to the converter queue", sConceptWorkFldr, , sConceptId
        GoTo Block_Exit
    End If

        ' Add each one to the conversion queue
        ' output should be the sconceptworkfldr path
        ' input is the original file location
    
    While Not oRs.EOF
        If FileExists(CStr("" & oRs("ImagePath").Value)) Then
        
            sIcn = CStr("" & oRs("ICN").Value)
            sToFilename = oReqDoc.ParseFileName(sConceptId, oConcept.ClientIssueId(0), sIcn, oRs("ImagePath").Value) & _
                "." & LCase(oReqDoc.SendAsFileType)
        
            If AddConverterQueueJob(oRs("ImagePath").Value, oReqDoc.SendAsFileType, sConceptWorkFldr, sToFilename, False, True, False, 0, lBatchId, lJobId) = False Then
                LogMessage strProcName, "ERROR", "Could not add a job these to the converter queue", sConceptWorkFldr, , sConceptId
                GoTo NextOneThanks
            End If
        End If
NextOneThanks:
        oRs.MoveNext
    Wend
    
        ' Close the batch
    If CloseBatch(lBatchId) = False Then
        LogMessage strProcName, "ERROR", "Problem closing the batch for ID: " & CStr(lBatchId), , , sConceptId
        GoTo Block_Exit
    End If

    ' Wait for files to be converted..
    If WaitForBatchOrJobFinish(lBatchId, , sOutMsg, bTimedOutWaiting) = False Then
        LogMessage strProcName, "ERROR", "There was a problem converting the files to proper name / format", "Timed out: " & CStr(bTimedOutWaiting), True, sConceptId
        GoTo Block_Exit
    End If
    
    ' Done??
    
  
    GetMedicalRecords = True
    
Block_Exit:
    Set oConcept = Nothing
    Set oReqDoc = Nothing
    Set oAdo = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    DoCmd.Hourglass False
    DoCmd.Echo True, "Ready..."

    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GetMedicalRecords = False
    GoTo Block_Exit
End Function





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Creates the NIRF
'''
Public Function CreatePackageNirf(sConceptId As String, lPayerID As Long, _
        Optional bOpenExplorer As Boolean = False, Optional bViewPdf As Boolean = True, _
        Optional sPayerName As String, Optional bCreateWorkCopy As Boolean = False, Optional sResubFolder As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oConcept As clsConcept
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim sClientIssueNum As String
Dim sDestFileName As String
Dim sFullDestPath As String
Dim sFullDestFldr As String
Dim strNewIssueReportName As String
Dim strReportCondition As String
Dim oNirfFile As clsConceptReqDocType
Dim oPayer As clsConceptPayerDtl
Dim bUpdateSubmitDate As Boolean
Dim sSubDt As String

    strProcName = ClassName & ".CreatePackageNirf"
    
    DoCmd.Hourglass True
    DoCmd.Echo True, "Creating NIRF..."


    If lPayerID <> 0 And lPayerID <> 1000 Then
        If sPayerName = "" Then
            sPayerName = GetPayerNameFromID(lPayerID)
        End If
        Set oPayer = New clsConceptPayerDtl
        If oPayer.LoadFromConceptNPayer(sConceptId, lPayerID) = False Then
            LogMessage strProcName, "ERROR", "Problem getting the payer object!", CStr(lPayerID) & " concept: " & sConceptId, , sConceptId
        End If
        
        sSubDt = oPayer.DateSubmitted

        If sResubFolder <> "" Then
            bUpdateSubmitDate = False
        Else
            If sSubDt > "" And sSubDt <> "1/1/1900" Then
                If MsgBox("There is already a submit date of '" & sSubDt & "' for " & oPayer.PayerName & vbCrLf & vbCrLf & "Do you want to keep that date (yes) or use today for the submit date (no)" _
                                , vbYesNo, "Use existing submit date as the submit date?") = vbYes Then
                        bUpdateSubmitDate = False
                    Else
                        bUpdateSubmitDate = True
                    End If
            Else
                bUpdateSubmitDate = True
            End If
        End If


        If bUpdateSubmitDate = True Then
            oPayer.DateSubmitted = Now()
            oPayer.SaveNow
        End If
    End If


    Set oNirfFile = New clsConceptReqDocType
    If oNirfFile.LoadFromDocName("ERAC_NIRF") = False Then
        LogMessage strProcName, "ERROR", "Problem getting the NIRF document type", , , sConceptId
        GoTo Block_Exit
    End If

    strNewIssueReportName = "rpt_CONCEPT_New_Issue"
    strReportCondition = "ConceptID = '" & sConceptId & "' And PayerNameId = " & CStr(lPayerID)
    
    Set oConcept = New clsConcept
    If oConcept.LoadFromId(sConceptId) = False Then
        LogMessage strProcName, "ERROR", "Could not load concept object from id", sConceptId, , sConceptId
        GoTo Block_Exit
    End If
    
        '' If there's a NIRF there already, remove that record..
    If sResubFolder = "" Then
        RemoveSpecificDocForThisConcept sConceptId, lPayerID, "ERAC_NIRF"
    End If
    
    Set oFso = New Scripting.FileSystemObject

    If oFso.FolderExists(oConcept.ConceptFolder) = False Then
        Set oFldr = oFso.CreateFolder(oConcept.ConceptFolder)
    Else
        Set oFldr = oFso.GetFolder(oConcept.ConceptFolder)
    End If

    If oFldr Is Nothing Then
        CreatePackageNirf = False
        LogMessage strProcName, "ERROR", "Could not get the concept folder", oConcept.ConceptFolder, , sConceptId
        GoTo Block_Exit
    End If


    If oPayer Is Nothing Then
        If oConcept.ClientIssueId(0) <> "" Then
            sDestFileName = oNirfFile.ParseFileName(sConceptId, oConcept.ClientIssueId(0), , , "")
        Else
Stop
        End If
    Else
        sDestFileName = oNirfFile.ParseFileName(sConceptId, oPayer.ClientIssueId, , , oPayer.PayerName, oPayer)
    End If

    If sResubFolder <> "" Then

        If InStr(1, sDestFileName, "_RESUBMIT_", vbTextCompare) < 1 Then
            sDestFileName = sDestFileName & "_RESUBMIT_" & Format(Now, "yyyymmdd")
        End If
    End If


    If sPayerName <> "" Then
    
        sFullDestPath = QualifyFldrPath(oFldr.Path) & sPayerName & "\" & sDestFileName & ".pdf"
        Call CreateFolders(sFullDestPath)
    Else
        sFullDestPath = QualifyFldrPath(oFldr.Path) & sDestFileName & ".pdf"
        Call CreateFolders(sFullDestPath)
    End If



        '' Make sure that report isn't open already
    If IsOpen(strNewIssueReportName, acReport) = True Then
        LogMessage strProcName, , "Closing NIRF report so we can create it..", strReportCondition, , sConceptId
        DoCmd.Close acReport, strNewIssueReportName, acSaveNo
    End If
    

'        ' Print concept report as PDF
    CreatePackageNirf = ConvertReportToPDF(strNewIssueReportName, strReportCondition, , sFullDestPath, False, bViewPdf)
'        CreatePackageNirf = RunReportAsPDF(strNewIssueReportName, strReportCondition, sFullDestPath)
    ' copy it to the 'work' folder
    If CreatePackageNirf = False Then
        LogMessage strProcName, "ERROR", "Could not create the NIRF for some reason", , , sConceptId
        GoTo Block_Exit
    End If
    
    If bCreateWorkCopy = True Then
        CreateFolders oConcept.ConceptWorkFolder
        If CopyFile(sFullDestPath, oConcept.ConceptWorkFolder & sPayerName, False) = False Then
            LogMessage strProcName, "ERROR", "Could not copy the file to the submit folder", oConcept.ConceptWorkFolder & sPayerName, , sConceptId
        End If
    End If
    
    sFullDestFldr = QualifyFldrPath(Replace(sFullDestPath, sDestFileName & ".pdf", ""))
    
        ' make sure it's "attached" to the concept
    If AddAttachedDocToDb(oConcept, sFullDestPath, sFullDestFldr, sDestFileName, "ERAC_NIRF", 0, sPayerName, oPayer) = False Then '' 1 is the NIRF type, 0 is for not a tagged claim doc
        LogMessage strProcName, "ERROR", "Could not attach the NIRF to the concept for some reason", , , sConceptId
    End If
    
    '        CreatePackageNirf = RunReportAsPDF(strNewIssueReportName, strReportCondition, sFullDestPath)
    LogMessage strProcName, , "Package NIRF created: " & CStr(CreatePackageNirf), sConceptId, , sConceptId
    
    Sleep 4000

    If bOpenExplorer = True And FileExists(sFullDestPath) = True Then
        If InStr(1, sFullDestPath, " ", vbTextCompare) > 0 And left(sFullDestPath, 1) <> """" Then
            sFullDestPath = """" & sFullDestPath & """"
        End If
        Shell "explorer.exe " & ParentFolderPath(sFullDestPath), vbNormalFocus
    End If
        

Block_Exit:
    Set oConcept = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing
    DoCmd.Hourglass False
    DoCmd.Echo True, "Ready..."

    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    CreatePackageNirf = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Creates the NIRF
'''
Public Function CreatePackageNirfOldStyle(sConceptId As String, Optional bOpenExplorer As Boolean = False, Optional bViewPdf As Boolean = True, _
        Optional sPayerName As String, Optional bCreateWorkCopy As Boolean = False, Optional sResubFolder As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oConcept As clsConcept
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim sClientIssueNum As String
Dim sDestFileName As String
Dim sFullDestPath As String
Dim sFullDestFldr As String
Dim strNewIssueReportName As String
Dim strReportCondition As String
Dim oNirfFile As clsConceptReqDocType
Dim oPayer As clsConceptPayerDtl
Dim bUpdateSubmitDate As Boolean
Dim sSubDt As String

    strProcName = ClassName & ".CreatePackageNirfOldStyle"
    
    DoCmd.Hourglass True
    DoCmd.Echo True, "Creating NIRF..."


'    If lPayerID <> 0 And lPayerID <> 1000 Then
'        If sPayerName = "" Then
'            sPayerName = GetPayerNameFromID(lPayerID)
'        End If
'        Set oPayer = New clsConceptPayerDtl
'        If oPayer.LoadFromConceptNPayer(sConceptId, lPayerID) = False Then
'            LogMessage strProcName, "ERROR", "Problem getting the payer object!", CStr(lPayerID) & " concept: " & sConceptId, , sConceptId
'        End If
        
        'sSubDt = oPayer.DateSubmitted
        sSubDt = oConcept.DateSubmitted
        
        If sResubFolder <> "" Then
            bUpdateSubmitDate = False
        Else
            If sSubDt > "" And sSubDt <> "1/1/1900" Then
                If MsgBox("There is already a submit date of '" & sSubDt & "' for " & oPayer.PayerName & vbCrLf & vbCrLf & "Do you want to keep that date (yes) or use today for the submit date (no)" _
                                , vbYesNo, "Use existing submit date as the submit date?") = vbYes Then
                        bUpdateSubmitDate = False
                    Else
                        bUpdateSubmitDate = True
                    End If
            Else
                bUpdateSubmitDate = True
            End If
        End If


'        If bUpdateSubmitDate = True Then
'            oPayer.DateSubmitted = Now()
'            oPayer.SaveNow
'        End If
'    End If


    Set oNirfFile = New clsConceptReqDocType
    If oNirfFile.LoadFromDocName("ERAC_NIRF") = False Then
        LogMessage strProcName, "ERROR", "Problem getting the NIRF document type", , , sConceptId
        GoTo Block_Exit
    End If

    strNewIssueReportName = "rpt_CONCEPT_New_Issue"
    strReportCondition = "ConceptID = '" & sConceptId & "'"
    
    Set oConcept = New clsConcept
    If oConcept.LoadFromId(sConceptId) = False Then
        LogMessage strProcName, "ERROR", "Could not load concept object from id", sConceptId, , sConceptId
        GoTo Block_Exit
    End If
    
        '' If there's a NIRF there already, remove that record..
    If sResubFolder = "" Then
'        RemoveSpecificDocForThisConcept sConceptId, lPayerID, "ERAC_NIRF"
Stop
' kev, you need to either "fix" or create a new version of the below as sending 0 isn't going to work...
        RemoveSpecificDocForThisConcept sConceptId, 0, "ERAC_NIRF"
    End If
    
    Set oFso = New Scripting.FileSystemObject

    If oFso.FolderExists(oConcept.ConceptFolder) = False Then
        Set oFldr = oFso.CreateFolder(oConcept.ConceptFolder)
    Else
        Set oFldr = oFso.GetFolder(oConcept.ConceptFolder)
    End If

    If oFldr Is Nothing Then
        CreatePackageNirfOldStyle = False
        LogMessage strProcName, "ERROR", "Could not get the concept folder", oConcept.ConceptFolder, , sConceptId
        GoTo Block_Exit
    End If


    If oConcept.ClientIssueId(0) <> "" Then
        sDestFileName = oNirfFile.ParseFileName(sConceptId, oConcept.ClientIssueId(0), , , "")
    Else
Stop
    End If


    If sResubFolder <> "" Then

        If InStr(1, sDestFileName, "_RESUBMIT_", vbTextCompare) < 1 Then
            sDestFileName = sDestFileName & "_RESUBMIT_" & Format(Now, "yyyymmdd")
        End If
    End If


 
    sFullDestPath = QualifyFldrPath(oFldr.Path) & sDestFileName & ".pdf"
    Call CreateFolders(sFullDestPath)




        '' Make sure that report isn't open already
    If IsOpen(strNewIssueReportName, acReport) = True Then
        LogMessage strProcName, , "Closing NIRF report so we can create it..", strReportCondition, , sConceptId
        DoCmd.Close acReport, strNewIssueReportName, acSaveNo
    End If
    

'        ' Print concept report as PDF
    CreatePackageNirfOldStyle = ConvertReportToPDF(strNewIssueReportName, strReportCondition, , sFullDestPath, False, bViewPdf)
'        CreatePackageNirf = RunReportAsPDF(strNewIssueReportName, strReportCondition, sFullDestPath)
    ' copy it to the 'work' folder
    If CreatePackageNirfOldStyle = False Then
        LogMessage strProcName, "ERROR", "Could not create the NIRF for some reason", , , sConceptId
        GoTo Block_Exit
    End If
    
    If bCreateWorkCopy = True Then
        CreateFolders oConcept.ConceptWorkFolder
        If CopyFile(sFullDestPath, oConcept.ConceptWorkFolder & sPayerName, False) = False Then
            LogMessage strProcName, "ERROR", "Could not copy the file to the submit folder", oConcept.ConceptWorkFolder & sPayerName, , sConceptId
        End If
    End If
    
    sFullDestFldr = QualifyFldrPath(Replace(sFullDestPath, sDestFileName & ".pdf", ""))
    
        ' make sure it's "attached" to the concept
    If AddAttachedDocToDb(oConcept, sFullDestPath, sFullDestFldr, sDestFileName, "ERAC_NIRF", 0, sPayerName, oPayer) = False Then '' 1 is the NIRF type, 0 is for not a tagged claim doc
        LogMessage strProcName, "ERROR", "Could not attach the NIRF to the concept for some reason", , , sConceptId
    End If
    
    '        CreatePackageNirf = RunReportAsPDF(strNewIssueReportName, strReportCondition, sFullDestPath)
    LogMessage strProcName, , "Package NIRF created: " & CStr(CreatePackageNirfOldStyle), sConceptId, , sConceptId
    
    Sleep 4000

    If bOpenExplorer = True And FileExists(sFullDestPath) = True Then
        If InStr(1, sFullDestPath, " ", vbTextCompare) > 0 And left(sFullDestPath, 1) <> """" Then
            sFullDestPath = """" & sFullDestPath & """"
        End If
        Shell "explorer.exe " & ParentFolderPath(sFullDestPath), vbNormalFocus
    End If
        

Block_Exit:
    Set oConcept = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing
    DoCmd.Hourglass False
    DoCmd.Echo True, "Ready..."

    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    CreatePackageNirfOldStyle = False
    GoTo Block_Exit
End Function




'' 20120912 KD: Replaced by RemoveSpecificDocForThisConcept
'''''' ##############################################################################
'''''' ##############################################################################
'''''' ##############################################################################
''''''
'''''' if there's an existing NIRF, this will:
'''''' - rename the file
'''''' - delete it from the table (but not the file itself,)
''''''
'''Public Function RemoveNirfForThisConcept(sConceptId As String, lPayerNameID As Long) As Boolean
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''Dim oAdo As clsADO
'''Dim oRs As ADODB.Recordset
'''Dim oFso As Scripting.FileSystemObject
'''Dim SFileName As String
'''Dim sFolderPath As String
'''Dim sNewName As String
'''Dim iFileCnt As Integer
'''Dim lRowId As Long
'''Dim sRowIds As String
'''Dim saryRows() As String
'''Dim iIdx As Integer
'''
'''    strProcName = ClassName & ".RemoveNirfForThisConcept"
'''    Set oFso = New Scripting.FileSystemObject
'''
'''
'''    Set oAdo = New clsADO
'''    With oAdo
'''        .ConnectionString = GetConnectString("CONCEPT_Hdr")
'''        .SQLTextType = sqltext
'''        .SQLstring = "SELECT R.* FROM CONCEPT_References R INNER JOIN CMS_AUDITORS_ERAC.dbo.Concept_Submission_References CR ON R.RowID = CR.ReferenceRowId " & _
'''            " LEFT JOIN CMS_AUDITORS_ERAC.dbo.Concept_Submission_ReferencePayers CRP ON CR.ConceptReferenceId = CRP.ConceptReferenceId " & _
'''            "  WHERE R.ConceptId = '" & sConceptId & "' AND R.RefSubType = 'ERAC_NIRF' AND " & _
'''            " CRP.PayerNameId = " & CStr(lPayerNameID)
'''        Set oRs = .ExecuteRS
'''        If .GotData = True Then
'''            iFileCnt = 1
'''            While Not oRs.EOF
'''                SFileName = oRs("RefFileName").Value
'''                sFolderPath = Replace(oRs("RefLink").Value, SFileName, "")
'''                sFolderPath = QualifyFldrPath(sFolderPath)
'''                sNewName = Replace(SFileName, ".pdf", "_" & Format(iFileCnt, "0###") & ".pdf")
'''                While oFso.FileExists(sFolderPath & sNewName)
'''                iFileCnt = iFileCnt + 1
'''                    sNewName = Replace(SFileName, ".pdf", "_" & Format(iFileCnt, "0###") & ".pdf")
'''                Wend
'''                If RenameFile(sFolderPath & SFileName, sFolderPath & sNewName) = False Then
'''                    LogMessage strProcName, "ERROR", "Could not rename the NIRF file", "From:" & sFolderPath & SFileName & " to " & sNewName
'''                End If
'''                sRowIds = sRowIds & CStr(oRs("RowID").Value) & ","
'''                oRs.MoveNext
'''            Wend
'''        Else
''''            Stop    ' nothing to delete
'''            GoTo Block_Exit
'''        End If
'''    End With
'''
'''
'''    If sRowIds = "" Then
''''        Stop    ' nothing to delete
'''        GoTo Block_Exit
'''    End If
'''
'''    sRowIds = left(sRowIds, Len(sRowIds) - 1)
''''Stop
'''
'''
'''
'''    saryRows = Split(sRowIds, ",")
'''
'''    For iIdx = 0 To UBound(saryRows)
'''        lRowId = saryRows(iIdx)
'''
'''            '' Now we can delete it..
'''        With oAdo
'''            .ConnectionString = GetConnectString("v_CODE_DATABASE")
'''            .SQLTextType = StoredProc
'''            .SQLstring = "usp_CONCEPT_References_Delete"
'''            .Parameters.Refresh
'''            .Parameters("@pRowID") = lRowId
'''            .Execute
'''            If .Parameters("@pErrMsg").Value <> "" Then
'''                LogMessage strProcName, "ERROR", "There was a problem removing the document from the database. RowID: " & CStr(lRowId)
'''
'''            End If
'''        End With
'''    Next
'''
'''    RemoveNirfForThisConcept = True
'''
'''Block_Exit:
'''    Set oFso = Nothing
'''    Set oRs = Nothing
'''    Set oAdo = Nothing
'''    Exit Function
'''Block_Err:
'''    ReportError Err, strProcName
'''    Err.Clear
'''    RemoveNirfForThisConcept = False
'''    GoTo Block_Exit
'''End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' if there's an existing NIRF, this will:
''' - rename the file
''' - delete it from the table (but not the file itself,)
'''
Public Function RemoveSpecificDocForThisConcept(sConceptId As String, lPayerNameId As Long, Optional sDocTypeToRemove As String = "ERAC_NIRF") As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oFso As Scripting.FileSystemObject
Dim SFileName As String
Dim sFolderPath As String
Dim sNewName As String
Dim iFileCnt As Integer
Dim lRowId As Long
Dim sRowIds As String
Dim saryRows() As String
Dim iIdx As Integer
Dim sExt As String


    strProcName = ClassName & ".RemoveSpecificDocForThisConcept"
    Set oFso = New Scripting.FileSystemObject

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_Data_Database")
        .SQLTextType = sqltext
        
        ' Make sure we've passed the correct RefSubType:
        .sqlString = "SELECT TOP 1 RefSubType FROM CONCEPT_References WHERE RefSubType = '" & sDocTypeToRemove & "'"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            RemoveSpecificDocForThisConcept = False
            LogMessage strProcName, "ERROR", "Incorrect RefSubType specified to remove", sDocTypeToRemove & " for " & sConceptId & " PayerNameID: " & CStr(lPayerNameId), , sConceptId
            GoTo Block_Exit
        End If
        
        
        .sqlString = "SELECT R.* FROM CONCEPT_References R INNER JOIN CMS_AUDITORS_ERAC.dbo.Concept_Submission_References CR ON R.RowID = CR.ReferenceRowId " & _
            " LEFT JOIN CMS_AUDITORS_ERAC.dbo.Concept_Submission_ReferencePayers CRP ON CR.ConceptReferenceId = CRP.ConceptReferenceId " & _
            "  WHERE R.ConceptId = '" & sConceptId & "' AND R.RefSubType = '" & sDocTypeToRemove & "' AND " & _
            " CRP.PayerNameId = " & CStr(lPayerNameId)
        Set oRs = .ExecuteRS
        If .GotData = True Then
            iFileCnt = 1
            While Not oRs.EOF
                SFileName = oRs("RefFileName").Value
                sExt = oFso.GetExtensionName(oRs("RefFileName").Value)
                sFolderPath = Replace(oRs("RefLink").Value, SFileName, "")
                sFolderPath = QualifyFldrPath(sFolderPath)
                sNewName = Replace(SFileName, ".pdf", "_" & Format(iFileCnt, "0###") & ".pdf")
                While oFso.FileExists(sFolderPath & sNewName)
                    iFileCnt = iFileCnt + 1
                    sNewName = Replace(SFileName, "." & sExt, "_" & Format(iFileCnt, "0###") & "." & sExt)
                Wend
                If RenameFile(sFolderPath & SFileName, sFolderPath & sNewName) = False Then
                    LogMessage strProcName, "ERROR", "Could not rename the '" & sDocTypeToRemove & "' file", "From:" & sFolderPath & SFileName & " to " & sNewName, , sConceptId
                End If
                sRowIds = sRowIds & CStr(oRs("RowID").Value) & ","
                oRs.MoveNext
            Wend
        Else
'            Stop    ' nothing to delete
            GoTo Block_Exit
        End If
    End With
    
    
    If sRowIds = "" Then
'        Stop    ' nothing to delete
        GoTo Block_Exit
    End If
    
    sRowIds = left(sRowIds, Len(sRowIds) - 1)
'Stop
    


    saryRows = Split(sRowIds, ",")
    
    For iIdx = 0 To UBound(saryRows)
        lRowId = saryRows(iIdx)
 
            '' Now we can delete it..
        With oAdo
            .ConnectionString = GetConnectString("v_CODE_DATABASE")
            .SQLTextType = StoredProc
            .sqlString = "usp_CONCEPT_References_Delete"
            .Parameters.Refresh
            .Parameters("@pRowID") = lRowId
            .Execute
            If .Parameters("@pErrMsg").Value <> "" Then
                LogMessage strProcName, "ERROR", "There was a problem removing the document from the database. RowID: " & CStr(lRowId), , , sConceptId
                RemoveSpecificDocForThisConcept = False
            End If
        End With
    Next
   
    RemoveSpecificDocForThisConcept = True

Block_Exit:
    Set oFso = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    RemoveSpecificDocForThisConcept = False
    GoTo Block_Exit
End Function




'' 20120416 KD Added function:


Public Function AddAttachedDocToDb(oConcept As clsConcept, strFullFilePath As String, _
    strFolderPath As String, strFileName As String, strAttachmentTypeName As String, _
    intTaggedClaimId As Integer, sPayerName As String, Optional oPayer As clsConceptPayerDtl, Optional bWaitForJobToFinish As Boolean = False) As Boolean
Dim strProcName As String
On Error GoTo Block_Err
Dim myCode_ADO As clsADO
'Dim colPrms As ADODB.Parameters
'Dim prm As ADODB.Parameter
'Dim LocCmd As New ADODB.Command
Dim iResult As Integer
Dim lRowId As Long
Dim strErrMsg As String
'Dim cmd As ADODB.Command
Dim lJobId As Integer
Dim iPayerNID As Integer
Dim sExtension As String
Dim oFso As Scripting.FileSystemObject

    strProcName = ClassName & ".AddAttachedDocToDb"

    If Not oPayer Is Nothing Then
        iPayerNID = oPayer.PayerNameId
    End If
    
    
    Set oFso = New Scripting.FileSystemObject
    sExtension = oFso.GetExtensionName(strFullFilePath)
    
    If InStr(1, strFileName, sExtension, vbTextCompare) < 1 Then
        strFileName = strFileName & "." & sExtension
    End If


    Set myCode_ADO = New clsADO
    
    With myCode_ADO
        .SQLTextType = StoredProc
        .ConnectionString = GetConnectString("v_CODE_Database")
        .sqlString = "usp_CONCEPT_References_Insert_NEW_PayerDtl"
        
        .Parameters.Refresh
        
        .Parameters("@pConceptID") = oConcept.ConceptID
        .Parameters("@pPayerNameID") = iPayerNID
        .Parameters("@pCreateDt") = Now
        .Parameters("@pRefType") = "DOC"
            
        .Parameters("@pRefSubType") = strAttachmentTypeName
    
        .Parameters("@pRefLink") = strFullFilePath
        
        .Parameters("@pRefPath") = strFolderPath
        .Parameters("@pRefFileName") = strFileName
        .Parameters("@pRefSequence") = 0 ' the SProc will take care of this..
        .Parameters("@pRefURL") = ""
        
        .Parameters("@pRefDesc") = ""
        .Parameters("@pRefOnReport") = ""
        .Parameters("@pURLOnReport") = ""
        .Parameters("@pEracTaggedClaimId") = intTaggedClaimId
        
        LogMessage strProcName, , "Executing " & myCode_ADO.sqlString, , , oConcept.ConceptID
        
        iResult = myCode_ADO.Execute
        strErrMsg = Nz(.Parameters("@pErrMsg").Value, "")
    End With
    

        'Make sure there are no errors
    If Not Nz(strErrMsg, "") = "" Then
        AddAttachedDocToDb = False
            'Err.Raise 65000, "SaveData", "Error updating Hours - " & strErrMsg
        LogMessage strProcName, "WARNING", "Error in proc: " & strErrMsg, , True, oConcept.ConceptID
            'MsgBox "SaveData", "Error Logging Image - " & strErrMsg
    Else
        AddAttachedDocToDb = True
        
        lRowId = Nz(myCode_ADO.Parameters("@pRowId").Value, 0)
        
            '' Now, add it to the Converter Queue
            '' but only if it's NOT the correct type: KD COMEBACK!
        If AddAttachedDocToConversionQueue(oConcept.ConceptID, lRowId, strFolderPath, strFileName, sPayerName, bWaitForJobToFinish) = 0 Then
            LogMessage strProcName, "WARNING", "Didn't add to converter queue!", , , oConcept.ConceptID
        End If
        Sleep 1000
        
    End If
    
    AddAttachedDocToDb = True

Block_Exit:
    Set oFso = Nothing
    Set myCode_ADO = Nothing
    Exit Function

Block_Err:
    'Rollback anything we did up until this point
    'Return false so the whole save event can rollback.
    AddAttachedDocToDb = False
    ReportError Err, strProcName, , , oConcept.ConceptID
    GoTo Block_Exit
End Function




'''
'''''' ##############################################################################
'''''' ##############################################################################
'''''' ##############################################################################
''''''
'''''' returns the folder path if at least 1 was found and copied there.
''''''
'''Public Function GetClaimChartsLEGACY(sConceptId As String) As String
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''Dim oConcept As clsConcept
'''Dim oFso As Scripting.FileSystemObject
'''Dim oFldr As Scripting.Folder
'''Dim oAdo As clsADO
'''Dim oRs As ADODB.Recordset
'''Dim oCmd As ADODB.Command
'''Dim oChartFile As clsConceptReqDocType
'''Dim sNewName As String
'''Dim iCopyCount As Integer
'''Dim sOrigExtension As String
'''
'''    strProcName = ClassName & ".GetClaimChartsLEGACY"
'''
'''    If sConceptId = "" Then GoTo Block_Exit
'''
'''    Set oConcept = New clsConcept
'''    If oConcept.LoadFromID(sConceptId) = False Then
'''        LogMessage strProcName, "ERROR", "Could not load Concept object", sConceptId
'''        GoTo Block_Exit
'''    End If
'''
'''        '' Get our required doc type for reference
'''    Set oChartFile = New clsConceptReqDocType
'''    If oChartFile.LoadFromID(ciCHART_FILE_ID) = False Then
'''        LogMessage strProcName, "ERROR", "Could not load Med Chart object", "Id: " & CStr(ciCHART_FILE_ID)
'''        GoTo Block_Exit
'''    End If
'''
'''        ' Get our folder
'''    Set oFso = New Scripting.FileSystemObject
'''    If oFso.FolderExists(oConcept.ConceptWorkFolder) = False Then
'''        Set oFldr = oFso.CreateFolder(oConcept.ConceptWorkFolder)
'''    Else
'''        Set oFldr = oFso.GetFolder(oConcept.ConceptWorkFolder)
'''    End If
'''
'''
'''        '' Now get the record
'''    Set oAdo = New clsADO
'''    With oAdo
'''        .ConnectionString = GetConnectString("ConceptDocTypes")
'''        .SQLTextType = StoredProc
'''        .SQLstring = "usp_ChartFilePathsForConcept"
'''
'''        Set oCmd = New ADODB.Command
'''        oCmd.ActiveConnection = oAdo.CurrentConnection
'''        oCmd.CommandType = adCmdStoredProc
'''        oCmd.CommandText = .SQLstring
'''
'''        oCmd.Parameters.Refresh
'''
'''        oCmd.Parameters("@pConceptId") = sConceptId
'''
'''        Set oRs = .ExecuteRS(oCmd.Parameters)
'''
'''        If .GotData = False Then
'''                '' nothing to do
'''            GoTo Block_Exit
'''        End If
'''    End With
'''
'''
'''    If Not oRs Is Nothing Then
'''        While Not oRs.EOF
'''                ' If the source isn't there. make a note and move to the next one
'''            If oFso.FileExists(Nz(oRs("ImagePath").Value, "")) = False Then
'''                LogMessage strProcName, "WARNING", "Chart does not exist where specified", Nz(oRs("ImagePath").Value, "")
'''                GoTo NextChart
'''            End If
'''                '' Grab the original extension so we can: a) put it back on the new filename and b) see if we need to add it
'''                '' to the convert file queue
'''            sOrigExtension = oFso.GetExtensionName(Nz(oRs("ImagePath").Value, ""))
'''
'''                '' New name:
'''            sNewName = oChartFile.ParseFileName(sConceptId, oConcept.ClientIssueId, Nz(oRs("ICN").Value, ""))
'''            If sNewName <> "" Then
'''                sNewName = QualifyFldrPath(oFldr.Path) & sNewName
'''            Else
'''                sNewName = QualifyFldrPath(oFldr.Path)
'''            End If
'''
'''            sNewName = sNewName & "." & LCase(sOrigExtension)
'''
'''            '' Does it need to be converted to whatever?
'''            If LCase(sOrigExtension) <> LCase(oChartFile.SendAsFileType) Then
'''                '' KD COMEBACK Add it to the Convert Queue !
'''            End If
'''
'''            If CopyFile(Nz(oRs("ImagePath").Value, ""), sNewName, False) = False Then
'''                LogMessage strProcName, "WARNING", "Could not copy file", Nz(oRs("ImagePath").Value, "") & " to " & oFldr.Path
'''                GoTo NextChart
'''            End If
'''            iCopyCount = iCopyCount + 1
'''NextChart:
'''            oRs.MoveNext
'''        Wend
'''    End If
'''
'''    If iCopyCount > 0 Then
'''        GetClaimChartsLEGACY = oFldr.Path
'''    End If
'''
'''Block_Exit:
'''    Set oConcept = Nothing
'''    Set oCmd = Nothing
'''    If Not oRs Is Nothing Then
'''        If oRs.State = adStateOpen Then oRs.Close
'''        Set oRs = Nothing
'''    End If
'''    Set oFso = Nothing
'''    Set oFldr = Nothing
'''    Set oAdo = Nothing
'''    Exit Function
'''Block_Err:
'''    ReportError Err, strProcName
'''    Err.Clear
'''    GetClaimChartsLEGACY = ""
'''    GoTo Block_Exit
'''End Function






''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' returns the folder path if at least 1 was found and copied there.
'''
Public Function GetClaimCharts(sConceptId As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oConcept As clsConcept
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oChartFile As clsConceptReqDocType
Dim sNewName As String
Dim iCopyCount As Integer
Dim sOrigExtension As String

    strProcName = ClassName & ".GetClaimCharts"

    If sConceptId = "" Then GoTo Block_Exit
    
    Set oConcept = New clsConcept
    If oConcept.LoadFromId(sConceptId) = False Then
        LogMessage strProcName, "ERROR", "Could not load Concept object", sConceptId, , sConceptId
        GoTo Block_Exit
    End If

        '' Get our required doc type for reference
    Set oChartFile = New clsConceptReqDocType
    If oChartFile.LoadFromId(ciCHART_FILE_ID) = False Then
        LogMessage strProcName, "ERROR", "Could not load Med Chart object", "Id: " & CStr(ciCHART_FILE_ID), , sConceptId
        GoTo Block_Exit
    End If

        ' Get our folder
    Set oFso = New Scripting.FileSystemObject
    If oFso.FolderExists(oConcept.ConceptWorkFolder) = False Then
        Set oFldr = oFso.CreateFolder(oConcept.ConceptWorkFolder)
    Else
        Set oFldr = oFso.GetFolder(oConcept.ConceptWorkFolder)
    End If


        '' Now get the record
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_ChartFilePathsForConcept"
        
        .Parameters("@pConceptId") = sConceptId
        
        Set oRs = .ExecuteRS()
        
        If .GotData = False Then
                '' nothing to do
            GoTo Block_Exit
        End If
    End With
    
    
    If Not oRs Is Nothing Then
        While Not oRs.EOF
                ' If the source isn't there. make a note and move to the next one
            If oFso.FileExists(Nz(oRs("ImagePath").Value, "")) = False Then
                LogMessage strProcName, "WARNING", "Chart does not exist where specified", Nz(oRs("ImagePath").Value, ""), , sConceptId
                GoTo NextChart
            End If
                '' Grab the original extension so we can: a) put it back on the new filename and b) see if we need to add it
                '' to the convert file queue
            sOrigExtension = oFso.GetExtensionName(Nz(oRs("ImagePath").Value, ""))
            
                '' New name:
            sNewName = oChartFile.ParseFileName(sConceptId, oConcept.ClientIssueId(0), Nz(oRs("ICN").Value, ""))
            If sNewName <> "" Then
                sNewName = QualifyFldrPath(oFldr.Path) & sNewName
            Else
                sNewName = QualifyFldrPath(oFldr.Path)
            End If
            
            sNewName = sNewName & "." & LCase(sOrigExtension)
            
            '' Does it need to be converted to whatever?
            If LCase(sOrigExtension) <> LCase(oChartFile.SendAsFileType) Then
                '' KD COMEBACK Add it to the Convert Queue !
            End If
            
            If CopyFile(Nz(oRs("ImagePath").Value, ""), sNewName, False) = False Then
                LogMessage strProcName, "WARNING", "Could not copy file", Nz(oRs("ImagePath").Value, "") & " to " & oFldr.Path, , sConceptId
                GoTo NextChart
            End If
            iCopyCount = iCopyCount + 1
NextChart:
            oRs.MoveNext
        Wend
    End If

    If iCopyCount > 0 Then
        GetClaimCharts = oFldr.Path
    End If

Block_Exit:
    Set oConcept = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oFso = Nothing
    Set oFldr = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GetClaimCharts = ""
    GoTo Block_Exit
End Function





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' If the concept doesn't have exactly the number of tagged claims that it is required
''' to have for submission then the user will be prompted to set the _Exceptions table
''' FYI: bPrompt4OtherPayers, if this is true then we'll prompt if we should apply that to all payers
''' for this concept, if user finally says no, then that's the number we use for ALL payers (even if
''' we've already prompted for one and they gave a different answer!!!)
Public Function PromptUserForTaggedClaimsException(oConcept As clsConcept, lPayerNameId As Long, iTaggedClaims As Integer, _
            iExpectedClaims As Integer, bPrompt4OtherPayers As Boolean) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sMsg As String
Dim sAnswer As String
Dim sThisPayerName As String

    strProcName = ClassName & ".PromptUserForTaggedClaimsException"
    If iExpectedClaims = 0 Then
        PromptUserForTaggedClaimsException = True
        GoTo Block_Exit
    End If
    
    sThisPayerName = GetPayerNameFromID(lPayerNameId)
    
    sMsg = "The concept: '" & oConcept.ConceptID & "' for payer: '" & sThisPayerName & _
                "' is expected to have " & CStr(iExpectedClaims) & " tagged claims but "
    If iExpectedClaims = 0 Then
        sMsg = sMsg & " does not have any."
    ElseIf iTaggedClaims > iExpectedClaims Then
        sMsg = sMsg & " has " & CStr(iTaggedClaims) & " which is more than expected!"
    ElseIf iExpectedClaims > iTaggedClaims Then
        sMsg = sMsg & "only has " & CStr(iTaggedClaims)
    ElseIf iExpectedClaims = iTaggedClaims Then
        sMsg = ""
        PromptUserForTaggedClaimsException = True
        GoTo Block_Exit
    End If

    sMsg = sMsg & vbCrLf & vbCrLf & "Please enter the number of claims that will be submitted with this concept."
    
    sAnswer = InputBox(sMsg, "How many claims will be submitted with this concept (" & oConcept.ConceptID & ")" & _
            "For payer: " & sThisPayerName, CStr(iTaggedClaims))
    
    If sAnswer = "" Or IsNumeric(sAnswer) = False Then
            '' We can assume they canceled..
        LogMessage strProcName, "USER ACTION", "User canceled", , , oConcept.ConceptID
        GoTo Block_Exit
    End If

        ' Now, does the user want to apply that to all payers for this concept?
        ' of course we don't need to prompt if there's only 1 payer for the concept..
    If oConcept.ConceptPayers.Count = 1 Then
        bPrompt4OtherPayers = False ' no need to prompt again - we only have 1!!!
    Else
        If MsgBox("Should all payers for this concept ('" & oConcept.ConceptID & "') have " & sAnswer & " claims?" & vbCrLf & _
                    " (Saying no means that you will be prompted for each unless you have the correct amount of tagged " & _
                    " claims for that payer)" _
                    , vbYesNo, "All payers should submit " & sAnswer & " sample claims?") = vbNo Then
            bPrompt4OtherPayers = True  ' they said no, so we'll prompt them again next time.
        Else
            bPrompt4OtherPayers = False ' they chose yes, so we don't want to prompt them again
        End If
            
    End If

        ' Now we need to update the _Exceptions table (or insert) but that'll be done with the Stored proc..
    If oConcept.SetRequiredClaimsNum(CInt(sAnswer), sMsg, lPayerNameId, bPrompt4OtherPayers) = False Then
        LogMessage strProcName, "ERROR", sMsg, oConcept.ConceptID, , oConcept.ConceptID
        GoTo Block_Exit
    End If

    PromptUserForTaggedClaimsException = True
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    PromptUserForTaggedClaimsException = False
    GoTo Block_Exit
End Function





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Public Function EracSetConceptAsSubmitted(oConcept As clsConcept) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".GetAttachTypeFromRowId"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_EracSetConceptAsSubmitted"
        .Parameters.Refresh
        .Parameters("@pConceptId") = oConcept.ConceptID
        .Parameters("@pSubmitUser") = Identity.UserName()
        .Execute
        If .Parameters("@pErrMsg") <> "" Then
            LogMessage strProcName, "ERROR", "Could not mark the concept as submitted for some reason", , , oConcept.ConceptID
            GoTo Block_Exit
        End If
    End With

    EracSetConceptAsSubmitted = True
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    EracSetConceptAsSubmitted = False
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Public Function AddAttachedDocToConversionQueue(sConceptId As String, lRowIdOfInsertedDoc As Long, _
    sAttachedPath As String, strFileName As String, Optional sPayerName As String = "", Optional bWaitForConversion As Boolean = True) As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttachType As clsConceptReqDocType
Dim sOutFileName As String
Dim sToType As String
Dim sOutFolder As String
Dim lJobId As Long
Dim sClientIssueNum As String
Dim sIcn As String
Dim bTimedOutWaiting As Boolean
Dim lPayerNameId As Long

    strProcName = ClassName & ".GetAttachTypeFromRowId"
    
    Set oAttachType = New clsConceptReqDocType
    Set oAttachType = GetAttachTypeFromRowId(lRowIdOfInsertedDoc, sConceptId, lPayerNameId, sIcn)
    
'    sOutFileName = Trim(oAttachType.ParseFileName(sConceptId, sClientIssueNum, sICN, , sPayerName))
    If InStr(1, strFileName, ".") > 1 Then
        sOutFileName = left(strFileName, InStr(1, strFileName, ".", vbTextCompare) - 1)
    Else
        sOutFileName = strFileName
    End If
    
    sToType = oAttachType.SendAsFileType
    
    sOutFolder = csCONCEPT_SUBMISSION_WORK_FLDR & sConceptId & "\"
    
    If sPayerName <> "" Then
        sOutFolder = sOutFolder & UCase(sPayerName) & "\"
    End If
''Stop ' kd: didn't do this yet.
'''    If oAttachType.IsPayerDoc Then
'''        LogMessage strProcName, "BURN NOTICE", "Need to burn this to DVD"
'''        sOutFolder = sOutFolder & "_BURN\"
'''    End If
    
    If AddConverterQueueJob(QualifyFldrPath(sAttachedPath) & strFileName, sToType, sOutFolder, sOutFileName, False, True, False, , , lJobId) = False Then
        LogMessage strProcName, "ERROR", "Could not add a job to the converter queue for this one..", sAttachedPath & " to: " & sToType, , sConceptId
    Else
        LogMessage strProcName, , "Created jobid: " & CStr(lJobId), , , sConceptId
    End If
    
    If bWaitForConversion = True And lJobId > 0 Then
            ' Wait for files to be converted..
        If WaitForBatchOrJobFinish(, lJobId, , bTimedOutWaiting) = False Then
            LogMessage strProcName, "ERROR", "There was a problem while waiting for the conversion job to finish", "Timed out: " & CStr(bTimedOutWaiting), True, sConceptId
            
            LogActionToHistory sConceptId, DocumentsConverted, "Failure waiting for conversion", , , CStr(bTimedOutWaiting)
            
            GoTo Block_Exit
        End If
    End If
    
    AddAttachedDocToConversionQueue = lJobId
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    AddAttachedDocToConversionQueue = 0
    GoTo Block_Exit
End Function





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Public Function AddPayerToEstablishedConceptFromPayerCopy(sConceptId As String, iPayerNameId2Add As Integer, iPayerNameIdToCopy As Integer, _
            Optional sClientIssueReturned As String, Optional sErrMsg As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".AddPayerToEstablishedConceptFromPayerCopy"
   
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_CODE_DATABASE")
        .SQLTextType = StoredProc
        
        
        If iPayerNameIdToCopy > 0 Then
            .sqlString = "usp_ConMgmt_AddPayerToExistConcept"
        Else
            .sqlString = "usp_ConMgmt_AddPayerToExistConceptWitoutCopyFrom"
        End If
        
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pPayerNameIdToAdd") = iPayerNameId2Add
        If iPayerNameIdToCopy > 0 Then
            .Parameters("@pPayerNameIdToCopyFrom") = iPayerNameIdToCopy
        End If
        .Execute
        If .Parameters("@pErrMsg").Value <> "" Then
            sErrMsg = .Parameters("@pErrMsg").Value
            LogMessage strProcName, "ERROR", "There was an error while trying to add a payer to the concept:" & sErrMsg, "ConceptID: " & sConceptId & " PayerToAdd: " & _
                        CStr(iPayerNameId2Add) & " Payer to copy from: " & CStr(iPayerNameIdToCopy), , sConceptId
        End If
        sClientIssueReturned = .Parameters("@pClientIssueNum").Value
    End With
   
    AddPayerToEstablishedConceptFromPayerCopy = True
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    AddPayerToEstablishedConceptFromPayerCopy = False
    GoTo Block_Exit
End Function





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Public Function DeleteConvertedFile(lRowId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim sSql As String
Dim oAttachType As clsConceptReqDocType
Dim sAttachedPath As String
Dim SFileName As String
Dim sPathToDelete As String
Dim sClientIssueNum As String
Dim sIcn As String
Dim sOutFolder As String
Dim sOutFileName As String
Dim sToType As String
Dim sFromType As String
Dim sConceptId As String
Dim iEracTaggedClaimId As Integer
Dim lPayerNameId As Long
Dim oPayer As clsConceptPayerDtl

    strProcName = ClassName & ".DeleteConvertedFile"

        '' Get the details from the database..
    sSql = "SELECT R.ConceptId, R.RefSubType, R.eRacTaggedClaimId, PayerNameId, R.RefFileName, DocTypeId FROM " & _
            " v_CONCEPT_References R WHERE R.RowId = " & CStr(lRowId)

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_Database")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "WARNING", "Could not find record of attached document with row id: " & CStr(lRowId)
            GoTo Block_Exit
        End If
    End With
    SFileName = CStr("" & oRs("RefFileName").Value)
    iEracTaggedClaimId = Nz(oRs("EracTaggedClaimId").Value, 0)
    sConceptId = Nz(oRs("ConceptID").Value, "")
    
    lPayerNameId = Nz(oRs("PayerNameId").Value, 999)

    If lPayerNameId <> 999 Then
        Set oPayer = New clsConceptPayerDtl
        If oPayer.LoadFromConceptNPayer(sConceptId, lPayerNameId) = False Then
            Stop    ' hammer time!
        End If
    End If
        
        '' Is there a package copy of the file, if so, grab it..
    Set oAttachType = New clsConceptReqDocType
    Set oAttachType = GetAttachTypeFromRowId(lRowId, sConceptId, lPayerNameId, sIcn)
'Stop ' kd: didn't do this yet.
'''    If oAttachType.IsPayerDoc = False Then
'''        ' get the ICN
'''        Set oAdo = New clsADO
'''        With oAdo
'''            .ConnectionString = GetConnectString("ConceptDocTypes")
'''            .SQLTextType = sqltext
'''            .SQLstring = "SELECT * FROM CnlyTaggedClaimsByConcept WHERE eRacTaggedClaimId = " & iEracTaggedClaimId
'''            Set oRs = .ExecuteRS
'''            If .GotData = False Then
'''                sICN = ""
'''                LogMessage strProcName, , "Could not find erac tagged claim id: " & CStr(iEracTaggedClaimId)
'''            Else
'''                sICN = CStr("" & oRs("ICN").Value)
'''            End If
'''        End With
'''    End If
    If oPayer Is Nothing Then
        sOutFileName = oAttachType.ParseFileName(sConceptId, sClientIssueNum, sIcn, SFileName)
    Else
        sOutFileName = oAttachType.ParseFileName(sConceptId, sClientIssueNum, sIcn, SFileName, oPayer.PayerName, oPayer)
    End If
    
    If sOutFileName = "" Then
        sOutFileName = CStr("" & oRs("RefFileName").Value)
    End If
    sToType = oAttachType.SendAsFileType

    

    sOutFolder = csCONCEPT_SUBMISSION_WORK_FLDR & sConceptId & "\"
'Stop ' kd: didn't do this yet.
    If oAttachType.IsPayerDoc And Not oPayer Is Nothing Then
        sOutFolder = sOutFolder & oPayer.PayerName & "\"
    End If

    sPathToDelete = sOutFolder & sOutFileName & "." & sToType

    LogMessage strProcName, , "Converted file path to delete: " & sPathToDelete, sPathToDelete, , sConceptId
    If FileExists(sPathToDelete) = False Then
        LogMessage strProcName, , "File to delete not found!", sPathToDelete, , sConceptId
    Else
        DeleteConvertedFile = DeleteFile(sPathToDelete, False)
        If DeleteConvertedFile = False Then
            LogMessage strProcName, "ERROR", "File could not be deleted!", sPathToDelete, , sConceptId
        Else
            LogMessage strProcName, , "Converted file was deleted!", sPathToDelete, , sConceptId
        End If

    End If
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    DeleteConvertedFile = False
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Public Function GetConvertedFileLocation(oAttachType As clsConceptReqDocType) As String
On Error GoTo Block_Err
Dim strProcName As String


    strProcName = ClassName & ".GetConvertedFileLocation"
        ' this would have to see if it's a claim level doc
        ' if so, get the ICN and CNLY Claim Num
        ' also would need the concept id
        
Stop    ' what calls this?

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GetConvertedFileLocation = ""
    GoTo Block_Exit
End Function





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Public Function GetAttachTypeFromRowId(lRowId As Long, Optional sConceptId As String, _
    Optional lPayerNumId As Long, Optional sIcn As String) As clsConceptReqDocType
On Error GoTo Block_Err
Dim strProcName As String
Dim oReturn As clsConceptReqDocType
Dim oAdo As clsADO
Dim sSql As String
Dim oRs As ADODB.RecordSet
Dim lDocType As Long
Dim iTaggedClaimId As Integer
Dim oClaim As clsEracClaim

    strProcName = ClassName & ".GetAttachTypeFromRowId"
    
    sSql = "SELECT R.ConceptId, R.RefSubType, R.eRacTaggedClaimId, R.DocTypeId, R.PayerNameID, R.ConceptReferenceID FROM v_CONCEPT_References R " & _
            " WHERE R.RowId = " & CStr(lRowId)
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_Database")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Could not load concept reference with rowid: " & CStr(lRowId), , , sConceptId
            GoTo Block_Exit
        End If
    End With

    sConceptId = Nz(oRs("ConceptId").Value, "")
    lDocType = Nz(oRs("DocTypeId").Value, 2)
    iTaggedClaimId = Nz(oRs("eRacTaggedClaimId").Value, 0)
    lPayerNumId = Nz(oRs("PayerNameID").Value, 999)

    
    Set oReturn = New clsConceptReqDocType
    
    If oReturn.LoadFromId(lDocType) = False Then
        LogMessage strProcName, "ERROR", "Problem determining what type of file this is..", CStr(lDocType), , sConceptId
        GoTo Block_Exit
    End If
    
    
    Set GetAttachTypeFromRowId = oReturn

Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Function





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Public Function GetConceptStoredProcSQL(sConceptId As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oReturn As clsConceptReqDocType
Dim oAdo As clsADO
Dim sConceptSql As String
Dim oRs As ADODB.RecordSet
Dim oFso As Scripting.FileSystemObject
Dim oTxt As Scripting.TextStream
Dim sTempPath As String
Dim sCmd As String
Dim sFileText As String
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim lTxtFile As Long
    
    strProcName = ClassName & ".GetConceptStoredProcSQL"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_Concept_GetSQL"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        Set oRs = .ExecuteRS
        If .Parameters("@pErrMsg") <> "" Then
            LogMessage strProcName, "ERROR", "Could not find the SQL for this concept.", "Error was: " & CStr(.Parameters("@pErrMsg")), True, sConceptId
            GoTo Block_Exit
        End If
    End With

    sTempPath = QualifyFldrPath(Environ("TEMP")) & "concept_sql_" & sConceptId & ".doc"
    
    While Not oRs.EOF
        '' Add coloring??
        sFileText = sFileText & CStr("" & oRs("Text").Value) & vbCrLf
        oRs.MoveNext
    Wend


    GetConceptStoredProcSQL = AddSqlColoringToSql(sFileText)



    lTxtFile = FreeFile
    
    On Error Resume Next    ' in case someone doesn't have permissions
    If FileExists(sTempPath) = True Then
        Call DeleteFile(sTempPath, False)
    End If
    Open sTempPath For Append Access Write Lock Write As lTxtFile
    Print #lTxtFile, GetConceptStoredProcSQL
    Close #lTxtFile
    On Error GoTo 0

    Set oFso = New Scripting.FileSystemObject
    
    Set oWord = New Word.Application
    
    Set oDoc = oWord.Documents.Open(sTempPath)
        
        
    If oWord.ActiveWindow.View.SplitSpecial = 0 Then    ' wdPaneNone Then
        oWord.ActiveWindow.ActivePane.View.Type = 3 ' wdPrintView
    Else
        oWord.ActiveWindow.View.Type = 3    'wdPrintView
    End If
        
    oWord.visible = True
    ActivateApplicationWindow , oDoc.Name & " - " & oWord.Caption
    oWord.Activate
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Set oDoc = Nothing
    Set oFso = Nothing
    Set oTxt = Nothing
    Set oWord = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function AddSqlColoringToSql(sInSql As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oRegEx As RegExp
Dim saryLines() As String
Dim iIndex As Integer
Dim bCommented As Boolean
Dim sThisLine As String
Dim sOut As String
Const sGreenStart As String = "<font color=""#00aa00"">"
Const sColorEnd As String = "</font>"
Const sBlueStart As String = "<font color=""#0000aa"">"
Const sSqlCmdRegEx As String = "(SELECT|FROM|WHERE|IN|EXISTS|GROUP BY|ORDER BY|NOT|BETWEEN|AND|OR|ANY|ALL|CASE|WHEN|ELSE|IF|BEGIN|SET|END)"

    strProcName = ClassName & ".AddSqlColoringToSql"
    
    sOut = "<html><body>"
    
    Set oRegEx = New RegExp
    
    oRegEx.IgnoreCase = True
    oRegEx.Global = True
    oRegEx.MultiLine = True
    
    
        '' Multi line comments:
    oRegEx.Pattern = "(\/\*[.\r\n\s\S\b\w]+\*\/)"
'    If oRegEx.test(sInSql) = False Then
'        Stop    ' fix the regex
'    End If

    sInSql = oRegEx.Replace(sInSql, sGreenStart & "$1" & sColorEnd)


        '' Single line comments:
    oRegEx.Pattern = "(--[^\r\n]+?[\r\n])"


    sInSql = oRegEx.Replace(sInSql, sGreenStart & "$1" & sColorEnd)
    
        '' Many extra blank lines:
'    oRegEx.Pattern = "[\r\n]{8,}"
'    sInSql = oRegEx.Replace(sInSql, vbCrLf)
        '' SQL Keywords:
            '' SELECT FROM WHERE IN AS CREATE PROCEDURE DECLARE EXEC OR AND
'            ' except that we don't want to do this when it's IN a comment..
'
'    oRegEx.Pattern = "(?<!" & QuoteMeta(sGreenStart) & ")" & sSqlCmdRegEx & "(?!" & QuoteMeta(sColorEnd) & ")"
'
'    oRegEx.Pattern = "(?<!\/\*+)" & sSqlCmdRegEx & "(?!\*+\/)"
'
'
'    If oRegEx.test(sInSql) = True Then
'        Stop
'    Else
'        Stop
'    End If
'    sInSql = oRegEx.Replace(sInSql, sBlueStart & "$1" & sColorEnd)
    sInSql = Replace(sInSql, vbCrLf, "<br />" & vbCrLf)
    
    
    sOut = sOut & sInSql
    

Block_Exit:
    AddSqlColoringToSql = sOut & "</body></html>"
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetRelatedPayerNameIDs(sConceptId As String) As Collection
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oRetCol As Collection

    strProcName = ClassName & ".GetRelatedPayerNameIDs"

    Set oRetCol = New Collection
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_DATA_DATABASE")
        .SQLTextType = sqltext
        .sqlString = "SELECT DISTINCT PayerNameId FROM CONCEPT_PAYER_Dtl WHERE ConceptID = '" & sConceptId & "'"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Could not retrieve concept payers - perhaps this is an old concept?", , , sConceptId
            GoTo Block_Exit
        End If
    End With
    
    While Not oRs.EOF
        oRetCol.Add oRs("PayerNameId").Value
        oRs.MoveNext
    Wend


Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Set GetRelatedPayerNameIDs = oRetCol
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetRelatedPayerNameIDsForFilter(sConceptId As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sRet As String

    strProcName = ClassName & ".GetRelatedPayerNameIDsForFilter"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_DATA_DATABASE")
        .SQLTextType = sqltext
        .sqlString = "SELECT DISTINCT PayerNameId FROM CONCEPT_PAYER_Dtl WHERE ConceptID = '" & sConceptId & "'"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Could not retrieve concept payers - perhaps this is an old concept?", , , sConceptId
            GoTo Block_Exit
        End If
    End With
    
    While Not oRs.EOF
        sRet = sRet & CStr("" & oRs("PayerNameID").Value) & ","
        oRs.MoveNext
    Wend
    sRet = left(sRet, Len(sRet) - 1)


Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    GetRelatedPayerNameIDsForFilter = sRet
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetRelatedPayerNames(sConceptId As String) As Collection
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oRetCol As Collection

    strProcName = ClassName & ".GetRelatedPayerNames"

    Set oRetCol = New Collection
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_DATA_DATABASE")
        .SQLTextType = sqltext
        .sqlString = "SELECT DISTINCT X.PayerName FROM CONCEPT_PAYER_Dtl CPD INNER JOIN XREF_PAYERNAMES X ON CPD.PayerNameId = X.PayerNameId WHERE ConceptID = '" & sConceptId & "'"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Could not retrieve concept payers", , , sConceptId
            GoTo Block_Exit
        End If
    End With
    
    While Not oRs.EOF
        oRetCol.Add oRs("PayerName").Value
        oRs.MoveNext
    Wend


Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Set GetRelatedPayerNames = oRetCol
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetRelatedPayerNamesStr(sConceptId As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oRetCol As Collection
Dim sRet As String

    strProcName = ClassName & ".GetRelatedPayerNamesStr"

    Set oRetCol = New Collection
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_DATA_DATABASE")
        .SQLTextType = sqltext
        .sqlString = "SELECT DISTINCT X.PayerName FROM CONCEPT_PAYER_Dtl CPD INNER JOIN XREF_PAYERNAMES X ON CPD.PayerNameId = X.PayerNameId WHERE ConceptID = '" & sConceptId & "'"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Could not retrieve concept payers", , , sConceptId
            GoTo Block_Exit
        End If
    End With
    
    While Not oRs.EOF
        sRet = sRet & oRs("PayerName").Value & ", "
        oRs.MoveNext
    Wend
    If Len(sRet) > 2 Then
        sRet = left(sRet, Len(sRet) - 2)    ' remove final , + space
    End If


Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    GetRelatedPayerNamesStr = sRet
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetPayerNameFromID(lngPayerNameId As Long) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sRet As String
Static dctPayers As Scripting.Dictionary


    strProcName = ClassName & ".GetPayerNameFromID"

    If dctPayers Is Nothing Then
        Set dctPayers = New Scripting.Dictionary
        
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("V_DATA_DATABASE")
            .SQLTextType = sqltext
            .sqlString = "SELECT DISTINCT X.PayerName, X.PayerNameId FROM XREF_PAYERNAMES X "
            Set oRs = .ExecuteRS
            If .GotData = False Then
                LogMessage strProcName, "ERROR", "Could not retrieve payers names"
                GoTo Block_Exit
            End If
        End With
        
        While Not oRs.EOF
            If dctPayers.Exists(oRs("PayerNameId").Value) Then
                Stop    ' should not have this!
                dctPayers.Item(oRs("PayerNameID").Value) = oRs("PayerName").Value
            Else
                dctPayers.Add oRs("PayerNameID").Value, oRs("PayerName").Value
            End If
            
            oRs.MoveNext
        Wend
        
        
    End If
    
    If dctPayers.Exists(lngPayerNameId) = False Then
        sRet = ""
    Else
        sRet = dctPayers.Item(lngPayerNameId)
    End If


Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    GetPayerNameFromID = sRet
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    sRet = ""
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetPayerNameIDFromName(ByVal strPayerName As String) As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim lRet As Long
Static dctPayers As Scripting.Dictionary


    strProcName = ClassName & ".GetPayerNameIDFromName"
    strPayerName = UCase(strPayerName)
    
    If dctPayers Is Nothing Then
        Set dctPayers = New Scripting.Dictionary
        
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("V_DATA_DATABASE")
            .SQLTextType = sqltext
            .sqlString = "SELECT DISTINCT X.PayerName, X.PayerNameId FROM XREF_PAYERNAMES X "
            Set oRs = .ExecuteRS
            If .GotData = False Then
                LogMessage strProcName, "ERROR", "Could not retrieve payers names"
                GoTo Block_Exit
            End If
        End With
        
        While Not oRs.EOF
            If dctPayers.Exists(oRs("PayerNameId").Value) Then
                Stop    ' should not have this!
                dctPayers.Item(UCase(oRs("PayerName").Value)) = oRs("PayerNameID").Value
            Else
                dctPayers.Add UCase(oRs("PayerName").Value), oRs("PayerNameID").Value
            End If
            
            oRs.MoveNext
        Wend
        
        
    End If
    
    If dctPayers.Exists(strPayerName) = False Then
        lRet = 0
    Else
        lRet = dctPayers.Item(strPayerName)
    End If


Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    GetPayerNameIDFromName = lRet
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    lRet = 0
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' [usp_CONCEPT_MarkQAd]
Public Function WasPackageCreated(sConceptId As String, lngPayerNameId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

Dim sRet As String
Static dctPayers As Scripting.Dictionary


    strProcName = ClassName & ".WasPackageCreated"

    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_DATABASE")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_Was_Pkg_Created"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pPayerNameId") = lngPayerNameId
        .Execute
        If CStr(Nz(.Parameters("@pErrMsg").Value, "")) <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg").Value, , , sConceptId
            GoTo Block_Exit
        End If
        WasPackageCreated = IIf(.Parameters("@pPkgWasCreated").Value = 0, False, True)
    End With


Block_Exit:

    Set oAdo = Nothing

    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    WasPackageCreated = False
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetStatusTextFromNum(sStatusNum As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sRet As String
Static dctStatus As Scripting.Dictionary


    strProcName = ClassName & ".GetStatusTextFromNum"

    If dctStatus Is Nothing Then
        Set dctStatus = New Scripting.Dictionary
        
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("V_DATA_DATABASE")
            .SQLTextType = sqltext
            .sqlString = "SELECT DISTINCT X.ConceptStatus, X.StatusDescription FROM CONCEPT_XREF_Status X "
            Set oRs = .ExecuteRS
            If .GotData = False Then
                LogMessage strProcName, "ERROR", "Could not retrieve payers names"
                GoTo Block_Exit
            End If
        End With
        
        While Not oRs.EOF
            If dctStatus.Exists(oRs("ConceptStatus").Value) Then
                Stop    ' should not have this!
                dctStatus.Item(oRs("ConceptStatus").Value) = oRs("StatusDescription").Value
            Else
                dctStatus.Add oRs("ConceptStatus").Value, oRs("StatusDescription").Value
            End If
            
            oRs.MoveNext
        Wend
        
        
    End If
    
    If dctStatus.Exists(sStatusNum) = False Then
        sRet = ""
    Else
        sRet = dctStatus.Item(sStatusNum)
    End If


Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    GetStatusTextFromNum = sRet
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    sRet = ""
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Creates an excel file for a payer and
Public Function CreateSampleClaimsDoc(oConcept As clsConcept, lPayerNameId As Long, Optional sOutMsg As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim iRequiredClaims As Integer
Dim oSClaimsDoc As clsConceptReqDocType
Dim oPayer As clsConceptPayerDtl
Dim colPayers As Collection

    strProcName = ClassName & ".CreateSampleClaimsDoc"
    Call StartMethod
    
    
    LogMessage strProcName, , "Starting to create the doc(s)", , , oConcept.ConceptID
    DoCmd.Hourglass True

    Set colPayers = New Collection
    If lPayerNameId = 1000 Then
        Set colPayers = oConcept.ConceptPayers
    Else
        Set oPayer = New clsConceptPayerDtl
        If oPayer.LoadFromConceptNPayer(oConcept.ConceptID, lPayerNameId) = False Then
            LogMessage strProcName, "ERROR", "Could not load the payer object for payer: " & CStr(lPayerNameId), , , oConcept.ConceptID
            GoTo Block_Exit
        End If
        colPayers.Add oPayer
    End If


    For Each oPayer In colPayers
    
        ' Get our doc type so we can parse the name
        Set oSClaimsDoc = New clsConceptReqDocType
        If oSClaimsDoc.LoadFromDocName("Payer_Claims") = False Then
            LogMessage strProcName, "ERROR", "Could not load the payer doc type object " & oPayer.PayerName, , , oConcept.ConceptID
            GoTo Block_Exit
        End If
        
    
        ' Does this concept / payer need claims - perhaps it's an exception
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("ConceptDocTypes")
            .SQLTextType = StoredProc
            .sqlString = "usp_EracNumOfClaimsRequired"
            .Parameters.Refresh
            .Parameters("@pConceptId") = oConcept.ConceptID
            .Parameters("@pPayerNameID") = oPayer.PayerNameId
            .Parameters("@pRequirementId") = oConcept.RequirementRuleObj.Id
            Set oRs = .ExecuteRS
            If .GotData = False Then
                LogMessage strProcName, , "No tagged claims found for this concept and payer", , , oConcept.ConceptID
                iRequiredClaims = 0
                'GoTo Block_Exit
                GoTo NxtPayer
            End If
        
           If Not oRs.EOF Then
               If Nz(oRs("ExceptionClaims"), -1) > -1 Then
                       ' This is an exception..
                   sOutMsg = "This is an exception, it normally requires " & CStr(oRs("NumClaimsPerConcept").Value) & _
                       " claims, but is set (with Connolly) to have " & CStr(oRs("ExceptionClaims").Value)
                   iRequiredClaims = oRs("ExceptionClaims").Value
               Else
                   iRequiredClaims = oRs("NumClaimsPerConcept").Value
               End If
                   ' Should only be 1 row, no need to movenext
           End If
            
        End With
        
        If iRequiredClaims = 0 Then
            LogMessage strProcName, , "This concept / Payer do not require any tagged claims", , , oConcept.ConceptID
            'GoTo Block_Exit
            GoTo NxtPayer
        End If
    
            ' get the tagged claims for this payer
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("v_Data_Database")
            .SQLTextType = sqltext
            
'            .sqlString = "SELECT V.ICN, VS.State FROM v_CONCEPT_TaggedClaims V INNER JOIN V_CONCEPT_ValidationSummary VS on V.CnlyClaimNum = VS.CnlyClaimNum AND V.ConceptID = VS.Adj_ConceptID " & _
'                " WHERE V.ConceptID = '" & oConcept.ConceptID & "' AND V.PayerNameID = " & CStr("" & oPayer.PayerNameID)


            .sqlString = "SELECT V.ICN, V.SampleState FROM CMS_AUDITORS_CLAIMS.dbo.CONCEPT_NIRF_Financials_Samples V " & _
                " WHERE V.ConceptID = '" & oConcept.ConceptID & "' AND V.PayerNameID = " & CStr("" & oPayer.PayerNameId)



            Set oRs = .ExecuteRS
            If .GotData = False Then
                LogMessage strProcName, "WARNING", "No tagged claims found for this concept and payer", , , oConcept.ConceptID
                GoTo NxtPayer
            End If
        End With
      
        
        
    Dim oExcel As Excel.Application
    Dim oWb As Excel.Workbook
    Dim oWs As Excel.Worksheet
    Dim oRange As Excel.Range
        
        ' dump to an excel spreadsheet

        
        If oExcel Is Nothing Then
           Set oExcel = New Excel.Application
        End If
        oExcel.visible = False
        Set oWb = oExcel.Workbooks.Add
        Set oWs = oWb.Sheets(1)
        
        oWs.Cells(1, 1) = "Sample Claim ICNs"
        oWs.Cells(1, 2) = "Country Code"
        
        oWs.Cells(1, 1).Select
        oExcel.selection.Font.Bold = True
        
        oWs.Cells(1, 2).Select
        oExcel.selection.Font.Bold = True
        
        
        Set oRange = oWs.Cells(2, 1)
        oRange.CopyFromRecordset oRs

        oWs.Cells.Select
        oWs.Cells.EntireColumn.AutoFit

        ' save in concept package folder
    Dim sOutFldr As String
    Dim sOutFile As String
    
        sOutFldr = QualifyFldrPath(oConcept.ConceptFolder) & oPayer.PayerName & "\"
        sOutFile = oSClaimsDoc.ParseFileName(oConcept.ConceptID, oPayer.ClientIssueId, , , oPayer.PayerName, oPayer)

        Call CreateFolders(sOutFldr)
        If FileExists(sOutFldr & sOutFile & ".xls") Then
            RemoveSpecificDocForThisConcept oConcept.ConceptID, oPayer.PayerNameId, "PAYER_CLAIMS"
'            Call DeleteFile(sOutFldr & sOutFile & ".xls", False)
            ' if it's in the database then remove that too...
            Call DeleteFile(sOutFldr & sOutFile & ".xls", False)

        End If
        oWb.SaveAs FileName:=sOutFldr & sOutFile & ".xls"
        
        oWb.Close
        Set oWb = Nothing

            ' make sure it's "attached" to the concept
        If AddAttachedDocToDb(oConcept, sOutFldr & sOutFile & ".xls", sOutFldr, sOutFile & ".xls", "Payer_Claims", 0, oPayer.PayerName, oPayer) = False Then
            LogMessage strProcName, "ERROR", "Could not attach the Sample Claims doc to the concept for some reason", , , oConcept.ConceptID
            Stop
        End If
        
        '' Add it to the conversion queue
        Sleep 1500
'''            If AddConverterQueueJob(sOutFldr & sOutFile & ".xls", "PDF", oConcept.ConceptWorkFolder & oPayer.PayerName & "\") = False Then
'''                Stop
'''            End If
        CreateSampleClaimsDoc = sOutFldr & sOutFile & ".xls"

NxtPayer:
    Next

Block_Exit:
    Call FinishMethod

    If Not oWs Is Nothing Then Set oWs = Nothing
    If Not oWb Is Nothing Then
        oWb.Close SaveChanges:=False
        Set oWb = Nothing
    End If
    If Not oExcel Is Nothing Then
        oExcel.Quit
        Set oExcel = Nothing
    End If
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    
    DoCmd.Hourglass False
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function ZipConceptSubmitPackage(oConcept As clsConcept, lPayerNameId As Long, Optional bResubmitting As Boolean) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAtchDoc As clsConceptDoc
Dim oRequireDocType As clsConceptReqDocType
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim lConverterJobId As Long
Dim sOutFolder As String
Dim colZipFiles As Collection
Dim sThisPayerName As String
Dim sOrigFilePath As String
Dim sIcn As String
Dim oClaim As clsEracClaim
Dim sNewFileName As String
Dim oPayer As clsConceptPayerDtl
Dim sThisClientIssueId As String
Dim bSampleClaimDocFound As Boolean

    strProcName = ClassName & ".ZipConceptSubmitPackage"


        '' What we need to do here is to mimic the stuff I did for the Resubmit old concept stuff..
            '' Bottom line, go through all of the docs,
            '' make sure they are properly converted
    
    sThisPayerName = GetPayerNameFromID(CInt(lPayerNameId))
    
    sOutFolder = QualifyFldrPath(oConcept.ConceptWorkFolder) & QualifyFldrPath(sThisPayerName)
    
    If bResubmitting = True Then
        sOutFolder = sOutFolder & "resub_" & Format(Now, "yyyymmdd") & "\"
    End If
    
    Call CreateFolders(sOutFolder)
    
    '' Let's delete everything from there first:
'    Call DeleteFullFolder(sOutFolder)
    
    '' Create the NIRF anew and make sure it is added to the package:
    If oConcept.NIRF_Exists(lPayerNameId) = False Or bResubmitting = True Then
        Call CreatePackageNirf(oConcept.ConceptID, lPayerNameId, False, False, sThisPayerName, True, IIf(bResubmitting, sOutFolder, ""))
    Else
        '' Do they want to create a new one?
        If MsgBox("There is already a NIRF for " & sThisPayerName & ". Do you want to use the existing one (Yes) or create a new one (No)?)", vbYesNo, "Use existing NIRF?") = vbNo Then
            Call CreatePackageNirf(oConcept.ConceptID, lPayerNameId, False, False, sThisPayerName, True, IIf(bResubmitting, sOutFolder, ""))
        End If
    End If
        '' we don't need to add the newly created nirf to the conversion queue because it's
        '' now an attached document, BUT! we DO need to refresh the AttachedDocuments
        '' collection in the Concept Object
    Call oConcept.RefreshObject
    
    Set colZipFiles = New Collection
    
    For Each oAtchDoc In oConcept.AttachedDocuments
        If oAtchDoc.PayerNameId <> 0 And oAtchDoc.PayerNameId <> lPayerNameId Then

            GoTo SkipAttachment
        End If
Dim coFlds As Collection
        Set coFlds = oAtchDoc.Fields
        If oAtchDoc.GetField("SendToPayers") <> True Then
'        If coFlds.Item("SendToPayers") <> 1 Then
'            Stop
            GoTo SkipAttachment
        End If
        
        If oAtchDoc.GetEracReqDocType.DocTypeId = 9 Then
            bSampleClaimDocFound = True
        End If
        
        Set oPayer = New clsConceptPayerDtl
        If oPayer.LoadFromConceptNPayer(oConcept.ConceptID, lPayerNameId) = False Then
            LogMessage strProcName, "ERROR", "Could not load the payer detail object!", , , oConcept.ConceptID
            GoTo SkipAttachment
        End If
        
        sThisClientIssueId = oPayer.ClientIssueId
        If sThisClientIssueId = "" Then
            sThisClientIssueId = oConcept.ClientIssueId(0)
            If sThisClientIssueId = "" Then
                LogMessage strProcName, "ERROR", "No client issue id found!", "Payer: " & oPayer.PayerName, , oConcept.ConceptID
                GoTo SkipAttachment
            End If
        End If
        
        If InStr(1, oAtchDoc.FolderPath, oPayer.PayerName, vbTextCompare) < 1 Then
            If oAtchDoc.IsPayerDoc = False Then
                sOrigFilePath = QualifyFldrPath(oAtchDoc.FolderPath) & oAtchDoc.FileName
            Else
                sOrigFilePath = QualifyFldrPath(oAtchDoc.FolderPath) & oPayer.PayerName & "\" & oAtchDoc.FileName
            End If
        Else
            sOrigFilePath = QualifyFldrPath(oAtchDoc.FolderPath) & oAtchDoc.FileName
        End If
        
        If InStr(1, oAtchDoc.FileName, ".", vbTextCompare) < 1 Then
            sOrigFilePath = sOrigFilePath & "." & oAtchDoc.GetEracReqDocType.SendAsFileType
        End If
        Set oRequireDocType = oAtchDoc.GetEracReqDocType
        
        If oAtchDoc.eRacTaggedClaimId <> 0 Then
            Set oClaim = GetClaimDetailsFromEracTaggedClaimId(oAtchDoc.eRacTaggedClaimId)
            sNewFileName = oRequireDocType.ParseFileName(oConcept.ConceptID, sThisClientIssueId, Nz(oClaim.Icn, ""))
            
        Else
        
            '' KD: Ok, so since the attach process should be doing the naming, we should not have to do it again
            ' only thing we need to do is to remove the file extension
            If InStr(1, oAtchDoc.FileName, ".", vbTextCompare) > 0 Then
                sNewFileName = left(oAtchDoc.FileName, InStr(1, oAtchDoc.FileName, ".", vbTextCompare) - 1)
            Else
                sNewFileName = oAtchDoc.FileName
            End If

        End If
    
        Call MarkPackageAsCreated(oConcept.ConceptID, oPayer.PayerNameId, sOutFolder)

        Call AddConverterQueueJob(sOrigFilePath, oAtchDoc.GetEracReqDocType.SendAsFileType, sOutFolder, sNewFileName, False, True, False, , , lConverterJobId)


        colZipFiles.Add sOutFolder & sNewFileName & "." & oAtchDoc.GetEracReqDocType.SendAsFileType
SkipAttachment:
    Next

    ' Make sure that there was a sample claims doc, and if not, create one:
'Dim oClaimDoc As clsConceptReqDocType
    bSampleClaimDocFound = False    ' we decided to recreate this for each click of the Create Package button
    If bSampleClaimDocFound = False Then
'        Set oClaimDoc = New clsConceptReqDocType
''        If oClaimDoc.LoadFromID(9) = False Then
''            Stop
''        End If
        
        sOutFolder = CreateSampleClaimsDoc(oConcept, CLng(oPayer.PayerNameId))
        ' But now we want the converted path:
        sOutFolder = Replace(sOutFolder, oConcept.ConceptFolder, oConcept.ConceptWorkFolder)
        sOutFolder = Replace(sOutFolder, ".xlsx", ".pdf")
        sOutFolder = Replace(sOutFolder, ".xls", ".pdf")
        
'        Application.Run oClaimDoc.CreateFunctionName, oConcept, CLng(oPayer.PayerNameId)
        
        colZipFiles.Add sOutFolder
        '' Need to attach it, or, it should have been attached in the above?

        Sleep 2500
    End If



        '' Now that we know they are properly converted, zip them up (put the zip cmd in a table then call a sproc)
        '' 20120801: KD ok, so... We're just going to generate a batch file to zip them properly
    ' Need to get the package name (zip file name)
    ' We are going to call it PayerName_[NDM_Packagename].zip or, PayerName_.zip if no packagename is set up
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_DataBase")
        .SQLTextType = sqltext
        .sqlString = "SELECT PayerNameId, DataType, NDM_Address, EmailToAddresses, SubmitThroughEmailOnly FROM CONCEPT_Payer_Send_Dtl WHERE " & _
                " PayerNameID = " & CStr(lPayerNameId) & " AND (DataType IS NULL OR DataType = '" & oConcept.GetField("DataType") & "') "
'Stop    ' this table changed, and I don't do this here anyway..
GoTo Skipto
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Could not get the zip details for " & sThisPayerName, , , oConcept.ConceptID
        End If
    End With

Skipto:
        
Dim sZipFileDestination As String
Dim oFso As Scripting.FileSystemObject
Dim oBatFile As Scripting.TextStream
Dim vFileToZip As Variant
Dim sZipCmd As String
Dim sSafePayerName As String
Dim sZipName As String
Dim sNovitasFldr
Dim oFldr As Scripting.Folder
Dim oFile As Scripting.file
Dim sZipSource As String

    sSafePayerName = Replace(oPayer.PayerName, " ", "")
    Set oFso = New Scripting.FileSystemObject
    
    If UCase(oPayer.PayerName) = "NOVITAS" Then
        sZipFileDestination = Replace(sOutFolder, "Novitas\", "") & "NEW_ISSUE_VALIDATION-" & Format(Now(), "yyyymmddhhnnss") & ".ZIP"
        
        
        sZipSource = ParentFolderPath(sOutFolder) & "NEW_ISSUE_VALIDATION\NEW_ISSUE_VALIDATION"
        
        Call CreateFolders(sZipSource)
        Set oFldr = oFso.GetFolder(ParentFolderPath(sOutFolder))
        For Each oFile In oFldr.Files
            Call CopyFile(oFile.Path, sZipSource, False)
        Next
        
        
        '' ok, novitas needs some extra stuff too - we need to create a subfolder named NEW_ISSU_VALIDATION
        ' which should also contain a NEW_ISSUE_VALIDATION subfolder in it (2 because of winzip's issues)
        ' ulitimate goal is to have a zip file named sZipFileDestination which contains
        ' a single folder: NEW_ISSU_VALIDATION
        
        
        Set oBatFile = oFso.CreateTextFile(left(ParentFolderPath(sOutFolder), Len(ParentFolderPath(sOutFolder)) - 1) & ".bat", True, False)
            '        oBatFile.WriteLine WINZIP_PATH & " -a -p -r -m """ & sZipFileDestination & """ """ & sOutFolder & "NEW_ISSUE_VALIDATION\*"""
        oBatFile.WriteLine WINZIP_PATH & " -a -p -r """ & sZipFileDestination & """ """ & ParentFolderPath(sOutFolder) & "NEW_ISSUE_VALIDATION\*"""
        oBatFile.Close
        
    Else
        If bResubmitting = True Then
            sZipFileDestination = sOutFolder & "CONCEPT_" & sSafePayerName & ".zip"
        Else
            sZipFileDestination = oConcept.ConceptWorkFolder & "CONCEPT_" & sSafePayerName & ".zip"
        End If
    
    End If
    
    
    If UCase(oPayer.PayerName) <> "NOVITAS" Then
        Set oBatFile = oFso.CreateTextFile(Replace(sZipFileDestination, ".zip", ".bat"), True, False)
        
        For Each vFileToZip In colZipFiles
            ' Generate a batch script for zipping..
            sZipCmd = WINZIP_PATH & " """ & sZipFileDestination & """ """ & CStr(vFileToZip) & """"
            oBatFile.WriteLine sZipCmd
        Next
        oBatFile.Close
    End If
    
    Sleep 1000
    '    Call Shell(Replace(sZipFileDestination, ".zip", ".bat"), vbHide)

    ZipConceptSubmitPackage = sZipFileDestination


Block_Exit:
    Set oFso = Nothing
    Set oBatFile = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing

    Exit Function
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    ZipConceptSubmitPackage = ""
    GoTo Block_Exit
End Function


Public Function MarkConceptAsSubmitted(sConceptId As String, lPayerNameId As Long, Optional bResubmitting As Boolean) As Boolean
On Error GoTo Block_Err
Dim oAdo As clsADO
Dim strProcName As String

    strProcName = ClassName & ".MarkConceptAsSubmitted"
    
      '' log the details to the DB:
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_MarkSubmitted"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pPayerNameID") = lPayerNameId
        .Parameters("@pResubmitting") = IIf(bResubmitting, 1, 0)
'Stop ' for testing.. don't update the database
        .Execute
        If Nz(.Parameters("@pErrMsg"), "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg"), , True, sConceptId
            GoTo Block_Exit
        End If
    End With
    

Block_Exit:
    Set oAdo = Nothing

    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Function



Public Function MarkConceptAsQAd(sConceptId As String, lPayerNameId As Long, Optional bResubmitting As Boolean) As Boolean
On Error GoTo Block_Err
Dim oAdo As clsADO
Dim strProcName As String

    strProcName = ClassName & ".MarkConceptAsQAd"
    
      '' log the details to the DB:
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_MarkQAd"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pPayerNameID") = lPayerNameId
'Stop ' for testing.. don't update the database
        .Execute
        If Nz(.Parameters("@pErrMsg"), "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg"), , True, sConceptId
            GoTo Block_Exit
        End If
    End With
    
    MarkConceptAsQAd = True
Block_Exit:
    Set oAdo = Nothing

    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Function





Public Function MarkConceptAsSentToCms(sConceptId As String, lPayerNameId As Long) As Boolean
On Error GoTo Block_Err
Dim oAdo As clsADO
Dim strProcName As String

    strProcName = ClassName & ".MarkConceptAsSentToCms"
    
      '' log the details to the DB:
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_Mark_Nirf_Sent_to_CMS"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pPayerNameID") = lPayerNameId

        .Execute
        If Nz(.Parameters("@pErrMsg"), "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg"), , True, sConceptId
            GoTo Block_Exit
        End If
    End With
    
    MarkConceptAsSentToCms = True
Block_Exit:
    Set oAdo = Nothing

    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Function



Public Function MarkPackageAsCreated(sConceptId As String, lPayerNameId As Long, sPackageFolder As String) As Boolean
On Error GoTo Block_Err
Dim oAdo As clsADO
Dim strProcName As String

    strProcName = ClassName & ".MarkPackageAsCreated"
    
      '' log the details to the DB:
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_MarkPkgCreated"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pPayerNameID") = lPayerNameId
        .Parameters("@pPackageFldr") = sPackageFolder
'Stop ' for testing.. don't update the database
        .Execute
        If Nz(.Parameters("@pErrMsg"), "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg"), , True, sConceptId
            GoTo Block_Exit
        End If
    End With
    
    MarkPackageAsCreated = True
Block_Exit:
    Set oAdo = Nothing

    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' saves the validation status
Public Sub SaveValidationHist(sConceptId As String, lngPayerNameId As Long, bFailed As Boolean, sFailMsg As String)
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".SaveValidationHist"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_ValidationHist"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pPayerNameID") = lngPayerNameId
        .Parameters("@pPassed") = IIf(bFailed, 0, 1)
        .Parameters("@pFailMessage") = sFailMsg
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg").Value, "Concept: " & sConceptId & ":" & CStr(lngPayerNameId), , sConceptId
            
            GoTo Block_Exit
        End If
    End With


Block_Exit:
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Sub




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Gets the validation status
Public Function GetValidationHist(sConceptId As String, lngPayerNameId As Long, Optional dtDateLastPassed As Date, Optional sFailMsg As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim bFailed As Boolean

    

    strProcName = ClassName & ".GetValidationHist"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_GetValidationHist"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pPayerNameID") = lngPayerNameId
        
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg").Value, "Concept: " & sConceptId & ":" & CStr(lngPayerNameId), , sConceptId
            GoTo Block_Exit
        End If
        bFailed = IIf(.Parameters("@pPassed").Value = 1, False, True)
        sFailMsg = Nz(.Parameters("@pFailMessage"), "")
        dtDateLastPassed = Nz(.Parameters("@pLastDt").Value, CDate("1/1/1900"))
        GetValidationHist = Not bFailed
    End With


Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Gets the validation status
Public Function GetConceptSettingVal(sName As String) As Variant
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Static oSetting As clsSettings

    If oSetting Is Nothing Then
        Set oSetting = New clsSettings
        oSetting.CurrentTableName = "CONCEPT_Settings"
        oSetting.ConnString = GetConnectString("CnlyDocTypes")
        
    End If

    

'
'    strProcName = ClassName & ".GetConceptSettingVal"
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("ConceptDocTypes")
'        .SQLTextType = StoredProc
'        .SQLstring = "usp_CONCEPT_GetValidationHist"
'        .Parameters.Refresh
'        .Parameters("@pConceptId") = sConceptId
'        .Parameters("@pPayerNameID") = lngPayerNameId
'
'        .Execute
'        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
'            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg").Value, "Concept: " & sConceptId & ":" & CStr(lngPayerNameId)
'            GoTo Block_Exit
'        End If
'
'        dtDateLastPassed = Nz(.Parameters("@pLastDt").Value, CDate("1/1/1900"))
'        GetValidationHist = bFailed
'    End With


Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function


'
'
'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Sub SaveSelectedPayerNameId(lPayerNameId As Long)
'On Error GoTo Block_Err
'Dim strProcName As String
'
'
'    strProcName = ClassName & ".SaveSelectedPayerNameId"
'
'
'
'Block_Exit:
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    Err.Clear
'    GoTo Block_Exit
'End Sub
'










Public Function ManuallyFix_Attachments(Optional sConceptId As String = "CM_C1609") As Boolean
Dim oRs As ADODB.RecordSet
Dim oCn As ADODB.Connection
Dim oCmd As ADODB.Command

Dim sSql As String
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oPayFldr As Scripting.Folder
Dim oFile As Scripting.file
Dim sFileExt As String

    sSql = "SELECT RowiD, RefLink, RefFileName  FROM CONCEPT_References WHERE RefFileName = ? AND ConceptId = '" & sConceptId & "'"

    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = GetConnectString("v_DATA_DATABASE")
'        .ConnectionString = "Provider=SQLNCLI10.1;Integrated Security=SSPI;Persist Security Info=False;User ID="""";Initial Catalog=CMS_AUDITORS_CLAIMS;Data Source=DS-FLD-009;"
        .Open
    End With
    
    Set oRs = New ADODB.RecordSet
    With oRs
'        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .ActiveConnection = oCn
    End With
    
    Set oCmd = New ADODB.Command
    With oCmd
        .ActiveConnection = oCn
        .CommandText = sSql
        .commandType = adCmdText
        
    End With
    
    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder("\\ccaintranet.com\dfs-cms-fld\Audits\CMS\AUDITCLM_ATTACHMENT\ConceptID\" & sConceptId & "\")
    
    For Each oPayFldr In oFldr.SubFolders
        For Each oFile In oPayFldr.Files
            sFileExt = oFso.GetExtensionName(oFile.Path)
            
            oCmd.Parameters(0) = Replace(oFile.Name, "." & sFileExt, "", , , vbTextCompare)
            
    
            Set oRs = oCmd.Execute
            If oRs Is Nothing Then
                Stop
            Else
                If oRs.EOF And oRs.BOF Then
                
                Else
'                    Stop
                    Debug.Print oRs("RowID").Value
                    sSql = "UPDATE CONCEPT_References SET RefLink = RefLink + '." & sFileExt & "', RefFileName = RefFileName + '." & sFileExt & "' WHERE RowId = " & CStr(oRs("RowId").Value)
                    
                    oCn.Execute sSql
                End If
            End If
            
            
        Next
    Next

End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'Public Function TestEracSomething(Optional sConceptId As String = "CM_C1370") As Boolean
'Dim oEracDoc As clsConceptReqDocType
'Dim sClientIssueNum As String
'Dim sICN As String
'
'    Set oEracDoc = GetAttachTypeFromRowId(6964, sConceptId, sClientIssueNum, sICN)
'
' Debug.Print "To Type: " & oEracDoc.SendAsFileType
' Stop
'
'End Function


Public Function ManuallyPrintListOfCmsOnlyNIRFs()
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oConcept As clsConcept
Dim strNewIssueReportName As String
Dim strReportCondition As String
Dim sFullDestPath As String
Const sOutFldr As String = "Y:\Data\CMS\AnalystFolders\KevinD\_Concept_Mgmt\Denise\NIRFs\"
Dim sClientIssueNum As String
Dim sConceptId As String
Dim strProcName As String


    strProcName = ClassName & ".ManuallyPrintListOfCmsOnlyNIRFs"
    
    strNewIssueReportName = "rpt_CONCEPT_New_Issue_CMS_ONLY"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=" & CurrentCMSServer() & ";Initial Catalog=CMS_AUDITORS_WORKSPACE;"
        .SQLTextType = sqltext
        .sqlString = "SELECT ICN_Maybe as ConceptID FROM KD_Tmp "
        Set oRs = .ExecuteRS
    End With
    
    While Not oRs.EOF
        ' Get our concept object
        
        sConceptId = oRs("ConceptId").Value
        
        Set oConcept = New clsConcept
        If oConcept.LoadFromId(sConceptId) = False Then
            Stop
        End If
        
        sClientIssueNum = oConcept.ClientIssueId(0)

            '' Make sure that report isn't open already
        If IsOpen(strNewIssueReportName, acReport) = True Then
            LogMessage strProcName, , "Closing NIRF report so we can create it..", strReportCondition, , sConceptId
            DoCmd.Close acReport, strNewIssueReportName, acSaveNo
        End If
        
If sClientIssueNum = "" Then
Stop
End If
        strReportCondition = "ConceptID = '" & sConceptId & "'"
        sFullDestPath = sOutFldr & sConceptId & "_" & sClientIssueNum & "_CONCEPT_New_Issue_CMS_Only.pdf"
'        ' Print concept report as PDF
        If ConvertReportToPDF(strNewIssueReportName, strReportCondition, , sFullDestPath, False, False) = False Then
            Stop
            'LogMessage strProcName, "ERROR", "Could not create the NIRF for some reason", , , sConceptId
            'GoTo Block_Exit
        End If
        
        oRs.MoveNext
    Wend
    
End Function




Public Function NIRFsForKenAdhoc()
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oConcept As clsConcept
Dim strNewIssueReportName As String
Dim strReportCondition As String
Dim sFullDestPath As String
Const sOutFldr As String = "Y:\Data\CMS\AnalystFolders\KevinD\_Concept_Mgmt\Ken\NIRFs\"
Dim sClientIssueNum As String
Dim sConceptId As String
Dim strProcName As String
Dim sFinalOutFldr As String
Dim sPayerName As String
Dim lPayerNameId As Long


    strProcName = ClassName & ".NIRFsForKenAdhoc"
    
    strNewIssueReportName = "rpt_CONCEPT_New_Issue_Ken_AdHoc"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=" & CurrentCMSServer() & ";Initial Catalog=CMS_AUDITORS_WORKSPACE;"
        .SQLTextType = sqltext
        .sqlString = "SELECT T.*, P.PayerName FROM KD_TMP_ClaimNum_n_Auditor T INNER JOIN CMS_AUDITORS_CLAIMS.dbo.XREF_Payernames P on T.PayerNameId = P.PayerNameId"
        Set oRs = .ExecuteRS
    End With
    
    While Not oRs.EOF
        ' Get our concept object
        sPayerName = oRs("PayerName").Value
        lPayerNameId = CLng(Nz(oRs("PayerNameId").Value, 999))
        
        
        sConceptId = oRs("ConceptId").Value
        
        Set oConcept = New clsConcept
        If oConcept.LoadFromId(sConceptId) = False Then
            Stop
        End If
        
        sClientIssueNum = oConcept.ClientIssueId(0)

            '' Make sure that report isn't open already
        If IsOpen(strNewIssueReportName, acReport) = True Then
            LogMessage strProcName, , "Closing NIRF report so we can create it..", strReportCondition, , sConceptId
            DoCmd.Close acReport, strNewIssueReportName, acSaveNo
        End If
        
If sClientIssueNum = "" Then
    sClientIssueNum = oConcept.ClientIssueId(lPayerNameId)
    strReportCondition = "ConceptID = '" & sConceptId & "' AND PayerNameID = " & CStr(lPayerNameId)

Else
        strReportCondition = "ConceptID = '" & sConceptId & "'"

End If
'        strReportCondition = "ConceptID = '" & sConceptId & "' AND PayerNameId = " & CStr(lPayerNameId)
        sFullDestPath = sOutFldr & sPayerName & "\" & sConceptId & "_" & sClientIssueNum & "_CONCEPT_New_Issue_CMS_Only.pdf"
        CreateFolders sFullDestPath
'        ' Print concept report as PDF
        If ConvertReportToPDF(strNewIssueReportName, strReportCondition, , sFullDestPath, False, False) = False Then
            Stop
            'LogMessage strProcName, "ERROR", "Could not create the NIRF for some reason", , , sConceptId
            'GoTo Block_Exit
        End If
        
        oRs.MoveNext
    Wend
    
End Function


Public Function IsMgrOrDS() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sUserProfile As String

    strProcName = ClassName & ".IsMgrOrDs"
    
 '   sUserProfile = Identity.UserName()
    sUserProfile = GetUserProfile()
    Select Case UCase(sUserProfile)
    Case "CM_ADMIN", "CM_AUDIT_MANAGERS"
        IsMgrOrDS = True
    Case Else
        IsMgrOrDS = False
    End Select
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
Option Compare Database
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Const ClassName As String = "cnlyStandardFunctions"

'* This Module contains general functions copied from "Standard" Connolly tools (powerbars, etc.)


Public Function LinkTableExists(strTableName As String) As Boolean
'Confirm the existance of the SQL View "ActiveAudit"
'If there is not yet a link to it, prompt the user to do so before proceeding.
    Dim strTest As String

    On Error Resume Next
    'Attempt to access the connect string for the linked table "ActiveAudit"
    strTest = CurrentDb.TableDefs(strTableName).Connect

    If Err.Number = 3265 Then    'It is not linked.
        LinkTableExists = False
    Else
        LinkTableExists = True
    End If

End Function


Public Function doesControlExist(frm As Form, CtrlName As String) As Boolean


    Dim ctl As String
    On Error Resume Next
    ctl = frm.Controls(CtrlName)


    If Err <> 0 Then
        doesControlExist = False
    Else
        doesControlExist = True
    End If


End Function

'VS 11/17/15 Change ErrorCode visibility rules - check box is visible if there is no recovery reason chosen at the line level - RVC effort.
Public Function ClaimOfRightReviewTypeDoesNotHaveLineLevelReason(sCnlyClaimNum As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim AvailableReviewType As Boolean

    strProcName = ClassName & ".ClaimOfRightReviewTypeDoesNotHaveLineLevelReason"
    
    If sCnlyClaimNum = "" Then
        LogMessage strProcName, "ERROR", "No connolly claim number to lookup!!!"
        GoTo Block_Exit
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "SELECT 1 FROM AUDITCLM_Hdr hdr where Adj_ReviewType in ('C', 'CV', 'CVDRG', 'CVMU', 'PRP', 'S', 'SV') and cnlyclaimnum = '" & sCnlyClaimNum & "'"
        '"and ClmStatus in ('320', '320.2', '322') and cnlyclaimnum = '" & sCnlyClaimNum & "'"
          
        Set oRs = .ExecuteRS
          
          If .GotData = True Then
            ClaimOfRightReviewTypeDoesNotHaveLineLevelReason = True
          Else
            ClaimOfRightReviewTypeDoesNotHaveLineLevelReason = False
          End If
            
        .sqlString = "SELECT 1 FROM AUDITCLM_Dtl dtl where RecoveryReason is not null and RecoveryReason <> '' and cnlyclaimnum = '" & sCnlyClaimNum & "'"
          
        Set oRs = .ExecuteRS
        
            If .GotData = True Then
                ClaimOfRightReviewTypeDoesNotHaveLineLevelReason = False
            End If
        
    End With
    
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    
    GoTo Block_Exit
End Function


Public Function IsThisClaimTherapyCongress(sCnlyClaimNum As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".IsThisClaimTherapyCongress"
    
    If sCnlyClaimNum = "" Then
        LogMessage strProcName, "ERROR", "No connolly claim number to lookup!!!"
        GoTo Block_Exit
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "SELECT 1 FROM AUDITCLM_Hdr CH INNER JOIN CONCEPT_Hdr H ON CH.Adj_ConceptID = H.ConceptID WHERE H.ConceptGroup = 'Therapy (Congress)' AND CH.CnlyClaimNum = '" & sCnlyClaimNum & "'"
        Set oRs = .ExecuteRS
        If .GotData = True Then
            
            IsThisClaimTherapyCongress = True

        Else

            IsThisClaimTherapyCongress = False
        End If
        
    End With
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    
    GoTo Block_Exit
End Function
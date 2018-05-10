Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private carsTables As Variant
Private cdctTableSources As Scripting.Dictionary
Private csConceptId As String


'' Last Modified: 20130426: Added No Payer stuff..

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


''' #########################################################################################
''' #########################################################################################
''' #########################################################################################
'''
'''         SET UP HERE:
'''
''' #########################################################################################
''' #########################################################################################
''' #########################################################################################
Private Function TableArray() As Variant
Dim vArr(2) As String

    vArr(0) = "v_CONCEPT_NewIssueProposal_NEW_Dtl_State_MANUAL"
    vArr(1) = "v_CONCEPT_NewIssueProposal_NEW_Dtl_State_Value_MANUAL"
    vArr(2) = "v_CONCEPT_NewIssueProposal_NEW_Manual"

    Set cdctTableSources = New Scripting.Dictionary
    
    
    ''' Note for this hokey solution:
    '' table alias C is where we want to get the ConceptId AND payernameid from
'    '' table alias P is where we want to get the payernameid from
    
    cdctTableSources.Add vArr(0), "SELECT DISTINCT C.ConceptState, C.ConceptID, X.StateName, " & _
        " C.PayerNameId FROM (CMS_AUDITORS_CLAIMS.dbo.CONCEPT_Dtl_State C INNER JOIN CMS_AUDITORS_CLAIMS.dbo.CONCEPT_Hdr H ON C.ConceptID = H.ConceptID) " & _
        " INNER JOIN CMS_AUDITORS_CLAIMS.dbo.CONCEPT_XREF_State X ON C.ConceptState = X.StateID WHERE "
        
    cdctTableSources.Add vArr(1), "SELECT DISTINCT S.State As ConceptState, R.ConceptID, S.StateDesc As StateName, C.EstAvailableClaims As ClaimCount, " & _
        " C.EstAvailableDollars As ClaimValue, NULL as ClaimCountSample, Null as ClaimValueState, R.DataType as Reference, C.PayerNameID " & _
        " FROM CMS_AUDITORS_REPORTS.dbo.RPT_R0043C AS R  INNER JOIN (CMS_AUDITORS_REPORTS.dbo.RPT_R0043G AS C LEFT JOIN CMS_AUDITORS_CLAIMS.dbo.XREF_State AS S " & _
        " ON C.ProvStCd = S.State) ON R.ConceptID = C.ConceptID " & _
        " WHERE C.EstAvailableClaims <> 0 AND C.EstAvailableDollars <> 0 "



    cdctTableSources.Add vArr(2), "SELECT C.* FROM CMS_AUDITORS_CODE.dbo.v_CONCEPT_NewIssueProposal_NEW C WHERE "
    

    TableArray = vArr
    
End Function
''' #########################################################################################
''' #########################################################################################
''' #########################################################################################
'''
'''         END SET UP
'''
''' #########################################################################################
''' #########################################################################################
''' #########################################################################################


Private Sub cmdOpenReport_Click()

    DoCmd.OpenReport "rpt_CONCEPT_New_Issue_Manual", acViewPreview, , "ConceptID = '" & csConceptId & "'"
    
End Sub

Private Sub cmdOpenTables_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim iIdx As Integer

    strProcName = ClassName & ".cmdOpenTables_Click"
    carsTables = TableArray
    
    For iIdx = 0 To UBound(carsTables)
        DoCmd.OpenTable carsTables(iIdx), acViewNormal, acEdit
    Next
    
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdSaveCurrentData_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim iIdx As Integer
Dim oDb As DAO.Database ' DAO Because we are only using local tables not linked tables
Dim oConcept As clsConcept
Dim sChosenPayerNIds As String
Dim sConceptId As String
Dim sConceptPayerIds As String



    strProcName = ClassName & ".cmdSaveCurrentData_Click"
    carsTables = TableArray
    
    Set oDb = CurrentDb
    
    For iIdx = 0 To UBound(carsTables)
        oDb.Execute "DELETE FROM " & carsTables(iIdx)
    Next
    
    sConceptId = InputBox("Please enter the concept ID in full form: CM_C####", "WHat concept?")
    If sConceptId = "" Then
        GoTo Block_Exit
    End If
    csConceptId = sConceptId
    
    Set oConcept = New clsConcept
    If oConcept.LoadFromId(sConceptId) = False Then
        Stop
    End If
    
    sConceptPayerIds = oConcept.ConceptPayerIDString

    If Nz(Me.ckNoPayers, False) = False Then
        '' Now, prompt for the Concept and Payers:
        If PromptUserForPayers("Pick the payer for this concept", oConcept, sConceptPayerIds, sChosenPayerNIds, , True) = False Then
            Stop
        End If
    End If
    
    ' Now, we have to populate our tables.
    Call PopulateLocalTables(carsTables, sConceptId, sChosenPayerNIds)
    
    
    
    
Block_Exit:
    Set oDb = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub PopulateLocalTables(vTableArray As Variant, sConceptId As String, sChosenPayerIDs As String)
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim iIdx As Integer
Dim sCurTable As String
Dim sSql As String
Dim oDb As DAO.Database
Dim oFld As ADODB.Field
Dim oLocalRS As DAO.RecordSet


    strProcName = ClassName & ".PopulateLocalTables"
    Set oDb = CurrentDb
    
    
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
    End With
    
    For iIdx = 0 To UBound(vTableArray)
        sCurTable = CStr(vTableArray(iIdx))
        
        If IsTable(sCurTable) = False Then
            Stop
        End If
        
        Set oLocalRS = oDb.OpenRecordSet(sCurTable)
        
        
        sSql = cdctTableSources.Item(sCurTable)
        
        '' append the where clause stuff..
        If sChosenPayerIDs = "" Then
            sSql = sSql & "AND C.ConceptId = '" & sConceptId & "' "
        Else
            sSql = sSql & "AND C.ConceptId = '" & sConceptId & "' AND C.PayerNameID IN (" & sChosenPayerIDs & ") "
        End If
        
        sSql = Replace(sSql, "WHERE AND ", "WHERE ")
        sSql = Replace(sSql, "WHERE  AND ", "WHERE ")
        
        oAdo.sqlString = sSql

        Set oRs = oAdo.ExecuteRS
        
        If oAdo.GotData = False Then
            GoTo NextTable
        End If
        
        While Not oRs.EOF
            With oLocalRS
                .AddNew
                For Each oFld In oRs.Fields
                    oLocalRS(oFld.Name) = oRs(oFld.Name).Value
                Next
                .Update
            End With
            

            oRs.MoveNext
        Wend
NextTable:
    Next
    
    
Block_Exit:
    Set oDb = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

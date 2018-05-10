Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim mstrSelHeight As Integer
Dim mstrSelWidth As Integer







Private Sub ICN_DblClick(Cancel As Integer)
    Navigate "frm_FastScan_SearchResults", "AUDITCLM", "DblClick", Me.CnlyClaimNum
End Sub

Private Sub ProcessInd_AfterUpdate()
    

  '  Stop

End Sub

Private Sub ProcessInd_BeforeUpdate(Cancel As Integer)

 'Stop

End Sub

Public Sub ProcessInd_Click()



    If Not (Me.RecordSet.EOF Or Me.RecordSet.BOF) And Not (Me.RecordSet.recordCount = 0) Then
        Me.Parent.Form.cmdSearch.SetFocus
    
        Dim sqlUpdate As String
        
        Dim CurrentClaim As String
        
        Dim ResultTypeTohandle As String
        
        If Me.Parent.Form.TogShowBarCodeResults.Value = -1 Then
            ResultTypeTohandle = "B"
        Else
            ResultTypeTohandle = "M"
        End If
       
        CurrentClaim = Me.CnlyClaimNum
        If Me.ProcessInd Then
            
            sqlUpdate = "UPDATE v_CA_SCANNING_FastScan_Search SET ProcessInd = False WHERE AccountID = " & gintAccountID & " and (CnlyClaimNum <> '" & Me.CnlyClaimNum & "' or RelatedClaimNum is not null) and UserID = '" & Me.UserID & "' and SessionID = " & Me.SessionID & " and ProcessInd = True and ResultType in ('R','M','B')"
            CurrentDb.Execute (sqlUpdate)
            
            'sqlUpdate = "UPDATE v_CA_SCANNING_FastScan_Search SET ProcessInd = False WHERE (CnlyClaimNum <> '" & Me.CnlyClaimNum & "' or RelatedClaimNum is not null) and UserID = '" & Me.UserID & "' and SessionID = " & Me.SessionID & " and ProcessInd = True and ResultType = '" & "M" & "'"
            'CurrentDb.Execute (sqlUpdate)
    
            sqlUpdate = "UPDATE v_CA_SCANNING_FastScan_Search SET ProcessInd = True WHERE AccountID = " & gintAccountID & " and (CnlyClaimNum = '" & Me.CnlyClaimNum & "') and UserID = '" & Me.UserID & "' and SessionID = " & Me.SessionID & " and ProcessInd = False and ResultType = '" & ResultTypeTohandle & "'"
            CurrentDb.Execute (sqlUpdate)
        Else
            sqlUpdate = "UPDATE v_CA_SCANNING_FastScan_Search SET ProcessInd = False WHERE AccountID = " & gintAccountID & " and UserID = '" & Me.UserID & "' and SessionID = " & Me.SessionID & "  and ResultType = '" & ResultTypeTohandle & "' and ProcessInd = true "
            CurrentDb.Execute (sqlUpdate)
        End If
        Me.Requery
        Me.RecordSet.FindFirst "CnlyClaimNum = " + "'" + CurrentClaim + "'"
       
       
        If Me.Parent.Form.TogRelated.Value = -1 Then

            If Me.ProcessInd Then
                
                Me.Parent.subfrm_Results_Other.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Search where AccountID = " & gintAccountID & " and CoverSheetNum = '" & Me.CoverSheetNum & "' and CnlyClaimNum = '" & Me.CnlyClaimNum & "' and UserID = '" & Me.UserID & "' and SessionID = " & Me.SessionID & " and ResultType = 'R'"
                Me.Parent.subfrm_Results_Other.Form.Requery
            
            Else
            
                Me.Parent.subfrm_Results_Other.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Search where 1=2"
                Me.Parent.subfrm_Results_Other.Form.Requery
            
            End If
            
            If Me.RecordSet.EOF And Me.RecordSet.BOF Then
                Me.Parent.cmdRelatedSelectAll.Enabled = False
                Me.Parent.cmdRelatedUnselectAll.Enabled = False
            Else
                Me.Parent.cmdRelatedSelectAll.Enabled = True
                Me.Parent.cmdRelatedUnselectAll.Enabled = True
            End If
            
        End If
    End If
    
    Call Me.Parent.DecisionSwitch
    
End Sub

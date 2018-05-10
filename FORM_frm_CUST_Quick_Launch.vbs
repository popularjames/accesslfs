Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub getSearchResults()

If Me.txtSearchFor = "" Then
Exit Sub
End If

    Select Case Me.optFindID.Value
        Case 1
            If Not IsNumeric(Me.txtSearchFor) Then Exit Sub
            FilterSelection (7)
        Case 2
            If Not IsNumeric(Me.txtSearchFor) Then Exit Sub
            FilterSelection (8)
        Case 3
            FilterSelection (9)
    End Select
    
    Me.optUserFilter = ""

End Sub


Private Sub frmSelection(intOPT As Integer)

Dim strUser As String
Dim isAdmin As String

strUser = Identity.UserName
isAdmin = DLookup("[isAdmin]", "CUST_Security_User", "[UserID] ='" & strUser & "'")

Select Case intOPT
    Case 1
            Select Case isAdmin
                Case "Y"
                    Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_EventSub"
                    Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "EventID"
                    Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
                Case "N"
                    Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_EventSub"
                    '2015-03-03 TK: change CreatedByUserID to AssignedTo
                    'Me.v_CUST_Serv_UserAssignment.Form.filter = "CreatedByUserID = '" & strUser & "'"
                    Me.v_CUST_Serv_UserAssignment.Form.filter = "AssignedTo = '" & strUser & "'"
                    Me.v_CUST_Serv_UserAssignment.Form.FilterOn = True
                    Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "EventID"
                    Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
            End Select
    Case 2
            Select Case isAdmin
                Case "Y"
                    Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_UserAssignmentSub"
                    Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "ActionID"
                    Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
                Case "N"
                    Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_UserAssignmentSub"
                    Me.v_CUST_Serv_UserAssignment.Form.filter = "AssignedToUserID = '" & strUser & "'"
                    Me.v_CUST_Serv_UserAssignment.Form.FilterOn = True
                    Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "ActionID"
                    Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
            End Select

    Case 3
            Select Case isAdmin
                Case "Y"
                    Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_EventClaimSub"
                    Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "ICN"
                    Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
                Case "N"
                    Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_EventClaimSub"
                    Me.v_CUST_Serv_UserAssignment.Form.filter = "LastUpdateUser = '" & strUser & "'"
                    Me.v_CUST_Serv_UserAssignment.Form.FilterOn = True
                    Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "ICN"
                    Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
            End Select


End Select
    
Me.v_CUST_Serv_UserAssignment.Form.Requery

End Sub

Private Sub FilterSelection(strFilter As String)

Dim strUser As String
Dim strSQL As String
Dim isAdmin As String
Dim strFilterEventSub As String
Dim strFilterAssgnSub As String
Dim strFilterClaimSub As String

On Error GoTo ErrHandler

strUser = Identity.UserName

isAdmin = DLookup("[isAdmin]", "CUST_Security_User", "[UserID] ='" & strUser & "'")


Select Case isAdmin
    Case "Y" 'User is an Admin
            strFilterEventSub = "EventID =" & Me.txtSearchFor & ""
            strFilterAssgnSub = "ActionID =" & Me.txtSearchFor & ""
            strFilterClaimSub = "ICN ='" & Me.txtSearchFor & "'"
            
    Case "N" 'User is not an Admin
            strFilterEventSub = "EventID =" & Me.txtSearchFor & " and CreatedByUserID = '" & strUser & "'"
            strFilterAssgnSub = "ActionID =" & Me.txtSearchFor & " and AssignedToUserID = '" & strUser & "'"
            strFilterClaimSub = "ICN ='" & Me.txtSearchFor & "' and LastUpdateUser = '" & strUser & "'"
End Select


Select Case Me.optFindID
        Case 1 'EventID
            Select Case strFilter 'allow users to see only there events
                Case 6
                        Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_EventSub"
                        Me.v_CUST_Serv_UserAssignment.Form.filter = "CreatedByUserID = '" & strUser & "'"
                        Me.v_CUST_Serv_UserAssignment.Form.FilterOn = True
                        Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "EventID"
                        Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
            
            Case 7
                        Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_EventSub"
                        Me.v_CUST_Serv_UserAssignment.Form.filter = strFilterEventSub
                        Me.v_CUST_Serv_UserAssignment.Form.FilterOn = True
                        Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "EventID"
                        Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True

                        
            End Select
            
        Case 2 'ActionID
           Select Case strFilter 'allow users to see only there events
                Case 2
                        Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_UserAssignmentSub"
                        Me.v_CUST_Serv_UserAssignment.Form.filter = "AssignedToUserID = '" & strUser & "'"
                        Me.v_CUST_Serv_UserAssignment.Form.FilterOn = True
                        Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "ActionID"
                        Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
                Case 3
                        Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_UserAssignmentSub"
                        Me.v_CUST_Serv_UserAssignment.Form.filter = "ReAssignedToUserID = '" & strUser & "'"
                        Me.v_CUST_Serv_UserAssignment.Form.FilterOn = True
                        Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "ActionID"
                        Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
                Case 4
                        Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_UserAssignmentSub"
                        Me.v_CUST_Serv_UserAssignment.Form.filter = "CCUserID = '" & strUser & "'"
                        Me.v_CUST_Serv_UserAssignment.Form.FilterOn = True
                        Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "ActionID"
                        Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
                   
                Case 5
                        Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_UserAssignmentSub"
                        Me.v_CUST_Serv_UserAssignment.Form.filter = "LastUpdateUser = '" & strUser & "'"
                        Me.v_CUST_Serv_UserAssignment.Form.FilterOn = True
                        Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "ActionID"
                        Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
                 
                Case 8
                        Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_UserAssignmentSub"
                        Me.v_CUST_Serv_UserAssignment.Form.filter = strFilterAssgnSub
                        Me.v_CUST_Serv_UserAssignment.Form.FilterOn = True
                        Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "ActionID"
                        Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
            End Select
        Case 3 ' Claim Number
            Select Case strFilter
                Case 9
                        Me.v_CUST_Serv_UserAssignment.SourceObject = "frm_Cust_EventClaimSub"
                        Me.v_CUST_Serv_UserAssignment.Form.filter = strFilterClaimSub
                        Me.v_CUST_Serv_UserAssignment.Form.FilterOn = True
                        Me.v_CUST_Serv_UserAssignment.Form.OrderBy = "ICN"
                        Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
            End Select
End Select

Me.v_CUST_Serv_UserAssignment.Form.OrderByOn = True
Me.v_CUST_Serv_UserAssignment.Form.Requery
Exit Sub

ErrHandler:
   If Err.Number = 2448 Then
    MsgBox "Please check your search criteria and try again", vbOKOnly + vbCritical, "Wrong Search Criteria"
   Else
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error Encountered"
   End If
End Sub

Private Sub cmdRefresh_Click()

Dim isAdmin As String
Dim strType As String

    isAdmin = DLookup("[isAdmin]", "CUST_Security_User", "[UserID] ='" & Identity.UserName & "'")
    
    strType = ""
    
    If isAdmin = "Y" Then
        strType = "(Admin)"
    End If
    
    Me.lblAppTitle.Caption = "Hello " & StrConv(left(Identity.UserName, InStr(1, Identity.UserName, ".") - 1), 3) & "     " & strType
    
    frmSelection (Me.optFindID.Value)
    Me.txtSearchFor = ""

End Sub

Private Sub cmdSearch_Click()
    getSearchResults
End Sub

Private Sub Command405_Click()
    DoCmd.OpenForm "frm_CTS_Hdr_Create"
End Sub

Private Sub Form_Load()
    Dim isAdmin As String
    Dim strType As String
    
    isAdmin = DLookup("[isAdmin]", "CUST_Security_User", "[UserID] ='" & Identity.UserName & "'")
    
    strType = ""
    
    If isAdmin = "Y" Then
        strType = "(Admin)"
    End If
    
    Me.optFindID.Value = 1
    Me.optUserFilter.Value = 6
    frmSelection (1)
    
    Me.lblAppTitle.Caption = "Hello " & StrConv(left(Identity.UserName, InStr(1, Identity.UserName, ".") - 1), 3) & "     " & strType


End Sub

Private Sub optFindID_AfterUpdate()
Dim intOptFind As Integer

intOptFind = Me.optFindID.Value

frmSelection (intOptFind)
Select Case intOptFind
    Case 1
        Me.optUserFilter.Value = 6
    Case 2
        Me.optUserFilter.Value = 2
    Case 3
        Me.optUserFilter.Value = 1
End Select

End Sub

Private Sub optUserFilter_AfterUpdate()

FilterSelection (Me.optUserFilter.Value)

End Sub

Private Sub txtSearchFor_AfterUpdate()
getSearchResults
End Sub

Option Compare Database
Option Explicit
'Damon - 06/03/08
'This function controls navigation of forms in the claim system
'Requires entry in the GENERAL_Navigate table
'strParentName - What called the navigation.  This is used to get the routing info from the GENERAL_Navigate table
'strAppID - What type of navigation.  This is needed because one parent form can have multiple navigation points
'strAction - What initiated the navigation (e.g. DblClick)
'strSearchParameter - A Value to pass to the form that we are navigating to
Public Sub Navigate(strParentName As String, _
                    strAppID As String, _
                    strAction As String, _
                    strSearchParameter As String)
                    
    Dim rst As DAO.RecordSet
    Dim lngTabID As Long
    Dim strFunction As String
    Dim strFunctionResult As String
    Dim strSQL As String
    Dim arrParameters() As String
    
    
    strSQL = " select * from GENERAL_Navigate where ParentForm  = '" & strParentName & "' and SearchType  = '" & strAppID & "' and ActionName = '" & strAction & "' "
    
    Set rst = CurrentDb.OpenRecordSet(strSQL)
        
    If Not rst.EOF Then
        'Hard coding the navigation points for now, as I am not sure how to make this dynamic
        'For now, it works, will revisit later.
        If Nz(rst!DestinationForm, "") = "frm_AUDITCLM_Main" Then
            NewMain strSearchParameter, "Main Claim"
        End If
        
        If Nz(rst!DestinationForm, "") = "frm_Prov_Hdr" Then
            NewProvider strSearchParameter, strSearchParameter
        End If
        
        If Nz(rst!DestinationForm, "") = "External" Then
            Application.FollowHyperlink strSearchParameter
        End If
        
        If Nz(rst!DestinationForm, "") = "frm_CONCEPT_Hdr" Then
            NewConcept strSearchParameter, strSearchParameter
        End If

        If Nz(rst!DestinationForm, "") = "frm_COLL_CNLY_Adj_Main" Then
            NewManualAdjustment strSearchParameter, strSearchParameter
        End If
        
    End If
End Sub
Public Function GetNavigateTabSQL(lngTabID As Long, frm As Form, _
                                  ByRef strFormValue As String, _
                                  ByRef strSQLCharacter As String, _
                                  ByRef strSQLValue As String, _
                                  ByRef strFormName As String)
    Dim strSQL As String
    Dim rst As DAO.RecordSet
    On Error GoTo ErrHandler
    Dim strOrderBy As String
  
    strSQL = " select * from GENERAL_Tabs where RowID = " & lngTabID
    
    Set rst = CurrentDb.OpenRecordSet(strSQL, dbOpenSnapshot, dbSeeChanges)
    
    If Not rst.EOF Then
        strFormName = Nz(rst!FormName, "")
        strFormValue = Nz(rst!FormValue, "")
        strSQLCharacter = Nz(rst!SQLCharacter, "")
        strSQLValue = Nz(rst!SQLValue, "")
        '' 20120910 KD
        If Nz(rst!RowSource, "") <> "" Then
            strSQL = Nz(rst!RowSource, "") & " WHERE " & strSQLValue & " = " & strSQLCharacter & Nz(frm.Controls(strFormValue), -1) & strSQLCharacter
            
            'Alex C - add Order By clause to form SQL source
            strOrderBy = Nz(rst!OrderBy, "")
            If Len(strOrderBy) > 0 Then
                strSQL = strSQL + " order by " & strOrderBy
            End If
    
            GetNavigateTabSQL = strSQL
        
        End If
    Else
        GetNavigateTabSQL = ""
    End If

Exit Function
ErrHandler:
        GetNavigateTabSQL = ""
End Function


Public Sub DisplayAuditClmMainScreen(CnlyClaimNum As String)
    Dim frm_AUDITCLM_Main As Form_frm_AUDITCLM_Main
    Dim iWindowHandle As Long
    Dim f As clsWindowHandles
    Dim bFound As Boolean
    
   
    bFound = False
    
    If CnlyClaimNum & "" <> "" Then
        For Each f In ColWindows
            iWindowHandle = f.WindowHandle
            SetForegroundWindow iWindowHandle
            If f.WindowName = "ClaimMain" & CnlyClaimNum Then
                SetForegroundWindow iWindowHandle
                
                If screen.ActiveForm.hwnd = iWindowHandle Then
                    bFound = True
                    Exit For
                End If
            End If
        Next
    
        If Not bFound Then
            Set f = New clsWindowHandles
            Set frm_AUDITCLM_Main = New Form_frm_AUDITCLM_Main
            f.WindowHandle = frm_AUDITCLM_Main.hwnd
            f.WindowName = "ClaimMain" & CnlyClaimNum
            ColWindows.Add f, f.WindowHandle & ""
            ColObjectInstances.Add Item:=frm_AUDITCLM_Main, Key:=frm_AUDITCLM_Main.hwnd & " "
            
            frm_AUDITCLM_Main.Caption = "CMS: ClaimNum : " & CnlyClaimNum
    
            frm_AUDITCLM_Main.visible = True
            frm_AUDITCLM_Main.CnlyClaimNum = CnlyClaimNum
            frm_AUDITCLM_Main.LoadData
        End If
    End If
    
    Set f = Nothing

End Sub
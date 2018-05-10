Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "GeneralTabs"

'=============================================
' ID:          Form_frm_ADMIN_General_Tabs
' Author:      Barbara Dyroff
' Create Date: 2009-10-09
' Description:
'   Maintain the GENERAL_TABS.  This is called
'   by the Claim Admin - Admin Maint - General Tabs function.
'   Form included in Form_frm_ADMIN_General_Tabs_Main
'
' Modification History:
'
' =============================================

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Cancel = CheckForNull(Me.Controls, True)
End Sub


Private Sub Form_Current()
    If IsSubForm(Me) Then
        Me.Parent.RowID = Me.RowID
        If Me.NewRecord Then
            Me.Parent.GeneralTabsDetails.Enabled = False
            Me.Parent.lblGeneralTabsDetail.Caption = "General Tabs Details"
        Else
            Me.Parent.lblGeneralTabsDetail.Caption = Chr(34) & Me.RowID & " - " & Me.TabName & Chr(34) & " General Tabs Details"
            Me.Parent.GeneralTabsDetails.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "General Tabs Maintenance"
    
    Call Account_Check(Me)
    
    If IsSubForm(Me) = False Then
        iAppPermission = UserAccess_Check(Me)
        If iAppPermission = 0 Then Exit Sub
    End If

End Sub

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "GeneralTabsDtl"

'=============================================
' ID:          Form_frm_ADMIN_General_Tabs_Detail
' Author:      Barbara Dyroff
' Create Date: 2009-10-09
' Description:
'   Maintain the GENERAL_TABS_Linked_Profile_IDs.  This is called
'   by the Claim Admin - Admin Maint - General Tabs function.
'   Form included in Form_frm_ADMIN_General_Tabs_Main.
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


Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "General Tab Linked Profile Details"
    
    Call Account_Check(Me)
    
    If IsSubForm(Me) Then
        Me.RowID.ColumnHidden = True
    Else
        Me.RowID.ColumnHidden = False

        RecordSource = "SELECT * FROM General_Tabs_Linked_ProfileIDs "
    
        iAppPermission = UserAccess_Check(Me)
        If iAppPermission = 0 Then Exit Sub
    End If
End Sub

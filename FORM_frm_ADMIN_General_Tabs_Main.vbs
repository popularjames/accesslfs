Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_ADMIN_General_Tabs_Main
' Author:      Barbara Dyroff
' Create Date: 2009-10-09
' Description:
'   Maintain the GENERAL_TABS and GENERAL_Tabs_Linked_ProfileIDs tables.  This is called
'   by the Claim Admin - Admin Maint - General Tabs function.
'   Checks permission before providing this Admin Maint function for the user.
'   Includes Form_frm_ADMIN_General_Tabs and Form_frm_ADMIN_General_Tabs_Detail
'
' Modification History:
'
' =============================================

Const CstrFrmAppID As String = "GeneralTabsMain"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Private Sub Form_Load()
    Dim iAppPermission As Integer
        
    Me.Caption = "General Tabs Maintenance"
        
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    lblGeneralTabs.Caption = gstrAcctDesc & " General Tabs"
    
End Sub

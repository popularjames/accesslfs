Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "UserProfileMain"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub Form_Load()
    
    Me.Caption = "User Profile Maintenance"
    
    Call Account_Check(Me)
    
    lblProfiles.Caption = gstrAcctDesc & " User Profiles"
    
    Call UserAccess_Check(Me)           ' must be last statement on form load
End Sub

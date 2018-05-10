Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "ProfileMain"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Private Sub Form_Load()
    Dim iAppPermission As Integer
        
    Me.Caption = "Profile Maintenance"
        
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    lblProfiles.Caption = gstrAcctDesc & " Profiles"
    
End Sub

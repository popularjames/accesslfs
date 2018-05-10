Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    Me.AllowAdditions = False
    Me.AllowDeletions = False
    Me.AllowEdits = True
    
    Me.ImageType.RowSource = "select ImageType, ImageTypeDisplay, ImageTypeDesc from v_SCANNING_XREF_ImageType_Account where Active = 'Y' and DataEntryVisible = 'Y' and AccountID = " & str(gintAccountID)
    
End Sub

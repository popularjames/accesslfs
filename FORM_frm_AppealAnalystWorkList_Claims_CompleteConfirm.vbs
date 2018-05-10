Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'******************************************************************************
'Created by Kathleen  C Flanagan Friday 8/8/2014
'Description:  Will provide a display for the user of the updates made when the 'Complete' button is chosen on the AppealAnalystDocumentsWorkList form
'Will mimic the Claim Note that is written for the claim.
'******************************************************************************

Sub cmdClose_Click()

    Me.txtUserFindings.Value = ""
    Me.visible = False
    
Exit_cmdCancel_Click:
    Exit Sub
    
Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub


Sub Form_Open(Cancel As Integer)


strCommentType = Me.OpenArgs
Me.txtUserFindings = ""

On Error GoTo Err_handler
 
    Me.txtUserFindings.visible = True
    Me.txtUserFindings = ""

Exit_Sub:
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub
 
End Sub

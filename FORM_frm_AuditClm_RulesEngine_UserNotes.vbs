Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'******************************************************************************
'Modified by Kathleen  C Flanagan Thursday 2/28/2013
'Description:  For the new functionality to updatea QA score - Upon Update QA, form
'will open for user comments for reason for change.
'******************************************************************************

Sub cmdClose_Click()

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
     Me.txtUserFindings = "Test Display"
    'Me.lblQAComment.Caption = "Reviewer Comments for Claims being waived from QA Review"
    'Me.lblUserID_Label.visible = False
    'Me.cboUserID.visible = False



Exit_Sub:
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub
 
End Sub

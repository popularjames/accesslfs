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

Sub CmdCancel_Click()

    Me.txtCommentCmd = "QACommentCancel"
    Me.txtQAComment = ""
    
    Me.visible = False
    
Exit_cmdCancel_Click:
    Exit Sub
    
Err_cmdCancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdCancel_Click
    
End Sub

Sub cmdSubmit_Click()

Me.txtCommentCmd = "QACommentSubmit"

Me.visible = False

Exit_cmdSubmit_Click:
    Exit Sub
    
Err_cmdSubmit_Click:
    MsgBox Err.Description
    Resume Exit_cmdSubmit_Click

End Sub

Sub Form_Open(Cancel As Integer)
Dim strCommentType

strCommentType = Me.OpenArgs
Me.txtQAComment = ""

On Error GoTo Err_handler
 
If strCommentType = "Waive" Then
    Me.txtQAComment.visible = True
    Me.lblQAComment.Caption = "Reviewer Comments for Claims being waived from QA Review"
    Me.lblUserID_Label.visible = False
    Me.cboUserID.visible = False
ElseIf strCommentType = "DRGReassign" Then
    Me.txtQAComment.visible = True
    Me.lblQAComment.Caption = "Reviewer Comments for Claim re-assignment to DRG"
'BEGIN 2/28/2013 KCF - Formatting for the Update QA functionality
ElseIf strCommentType = "UpdateQA" Then
    Me.txtQAComment.visible = True
    Me.lblQAComment.Caption = "Reviewer Comments for QA Update for Claim"
    Me.lblUserID_Label.visible = False
    Me.cboUserID.visible = False
'END 2/28/2013 KCF - Formatting for the Update QA functionality
End If

Exit_Sub:
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub
 
End Sub

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrAction As String

Public Event ReAssignQueue(AssignedTo As String, Comment As String, Action As String)

Public Property Let Action(data As String)
    mstrAction = data
    Me.Caption = data
    If UCase(mstrAction) = "REPLY" Then
        Me.txtComment.SetFocus
        Me.cboAssignedTo.visible = False
    Else
        Me.cboAssignedTo.visible = True
        Me.cboAssignedTo.SetFocus
    End If
End Property

Private Sub CmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdOk_Click()
    Dim strComment As String
    Dim strAssignedTo As String
    strComment = Me.txtComment & ""
    strAssignedTo = Me.cboAssignedTo & ""
    
    If strComment = "" And UCase(mstrAction) = "REPLY" Then
        MsgBox "Comment can not be blank when you are replying"
        Me.txtComment.SetFocus
        Exit Sub
    End If
    
    If strAssignedTo = "" And UCase(mstrAction) <> "REPLY" Then
        MsgBox "Please select a user first", vbInformation
        Me.cboAssignedTo.SetFocus
        Exit Sub
    End If
    
    RaiseEvent ReAssignQueue(strAssignedTo, strComment, mstrAction)
    DoCmd.Close acForm, Me.Name
End Sub


Private Sub Form_Close()
    RemoveObjectInstance Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Re-assignment"
    
    Call Account_Check(Me)
    'Me.cboAssignedTo.RowSource = "select * from ADMIN_User where UserID <> '" & Identity.Username & "'"
End Sub

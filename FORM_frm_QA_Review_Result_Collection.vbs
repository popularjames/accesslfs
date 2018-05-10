Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mbFormDirty As Boolean

Public Event process(AmtCorrect As String, DRGCorrect As String, RationaleCorrect As String, QAStatus As String, ReviewComment As String)
Public Event Cancel()


Private Sub AmtCorrect_AfterUpdate()
    mbFormDirty = True
End Sub

Private Sub CmdCancel_Click()
    RaiseEvent Cancel
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdOk_Click()
    Dim strChkStatus As String
    
    strChkStatus = UCase(Me.QAStatus & "")
    If mbFormDirty Then
        If InStr(1, "P,F", strChkStatus) = 0 Then
            MsgBox "Please indicate whether the review is passed or failed", vbCritical
            Exit Sub
        End If
    
        If strChkStatus = "F" And Me.ReviewComment & "" = "" Then
            MsgBox "You mark the review as 'Failed' but did not provide a comment.  Please add a comment for the auditor as to why it fails review. Thanks", vbCritical
            Exit Sub
        End If
    
        RaiseEvent process(Me.AmtCorrect, Me.DRGCorrect, Me.RationaleCorrect, Me.QAStatus, Me.ReviewComment)
            
        DoCmd.Close acForm, Me.Name
    Else
        MsgBox "You have not done anything yet"
    End If
End Sub


Private Sub DRGCorrect_AfterUpdate()
    mbFormDirty = True
End Sub

Private Sub Form_Close()
    RemoveWindow Me
End Sub

Private Sub Form_Load()
    Me.AmtCorrect = ""
    Me.DRGCorrect = ""
    Me.RationaleCorrect = ""
    Me.QAStatus = ""
    Me.ReviewComment = ""
    
End Sub

Private Sub QAStatus_AfterUpdate()
    mbFormDirty = True
End Sub

Private Sub QAStatus_Enter()
    If Me.AmtCorrect = "N" Or Me.DRGCorrect = "N" Or Me.RationaleCorrect = "N" Then
        Me.QAStatus.RowSource = "F;Fail"
    Else
        Me.QAStatus.RowSource = "P;Pass;F;Fail"
    End If
End Sub


Private Sub RationaleCorrect_AfterUpdate()
    mbFormDirty = True
End Sub

Private Sub ReviewComment_AfterUpdate()
    If Me.ReviewComment & "" <> "" Then mbFormDirty = True
End Sub

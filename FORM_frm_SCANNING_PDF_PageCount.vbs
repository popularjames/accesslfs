Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mstrCalledFrom As String

Const CstrFrmAppID As String = "ImageVal"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdValidate_Click()
    Dim strErrMsg As String
    
    Dim strErrSource As String
    strErrSource = "cmdValdation"
    
    On Error GoTo Err_handler
    
    ' reset display info
    lstFiles.RowSource = ""
    lstFiles.Requery
    
    ' get images
    Call GetImagePageCounts(Me)
        
    MsgBox "Update completed"
    
Exit_Sub:
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox "Error in module " & strErrSource & vbCrLf & vbCrLf & strErrMsg
    GoTo Exit_Sub
End Sub

Private Sub Form_Close()
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    Dim strScannedStation As String
    
    'Me.Caption = "PDF Page Count"
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    strScannedStation = GetPCName()
    If UCase(left(strScannedStation, 3)) <> "TS-" Then
        MsgBox "This function is only available on TS-DC-DEV session", vbInformation
        cmdValidate.Enabled = False
    End If
        
    
    lstFiles.RowSource = ""
    lstFiles.Requery
    
    If Me.OpenArgs() & "" <> "" Then
        mstrCalledFrom = Me.OpenArgs()
    End If
End Sub


Private Sub lstFiles_DblClick(Cancel As Integer)
    Dim strFileName As String
    Dim strImagePath As String
    
    strFileName = lstFiles.Column(1)
    strImagePath = Mid(strFileName, 1, Len(strFileName) - InStr(1, StrReverse(strFileName), "\"))
    Shell "explorer.exe " & strImagePath, vbNormalFocus
    Shell "explorer.exe " & strFileName, vbNormalFocus

End Sub

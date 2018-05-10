Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mbContinue As Boolean
Dim miError As Long
Dim miFileCnt As Long

Dim mstrCalledFrom As String

Const CstrFrmAppID As String = "ImageVal"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdValidate_Click()
    Dim db As Database
    Dim fso As New FileSystemObject
    Dim f As file
    
    'Dim myado As clsADO
    Dim rs As DAO.RecordSet
    
    Dim strFileName As String
    
    Dim strErrMsg As String
    Dim strErrSource As String
    Dim strOutcome As String
    
    strErrSource = "cmdValdation"
    
    Dim bResult As Boolean
    Dim iResult As Long
    
    On Error GoTo Err_handler
    
    'Set myado = New clsADO
    'myado.ConnectionString = GetConnectString("v_DATA_Database")
   
    
    ' reset display info
    lstFiles.RowSource = ""
    lstFiles.Requery
    miError = 0
    miFileCnt = 0
    
    ' get images
    Set db = CurrentDb
    Set rs = db.OpenRecordSet("SCANNING_Image_Retransfer")
    'Set rs = myado.OpenRecordSet("select * from SCANNING_Image_Log where ValidationDt is null or ValidationDt = '1/1/1900'")
    
   
    If (rs.BOF = True And rs.EOF = True) Then
        MsgBox "Nothing to do"
    Else
        rs.MoveFirst
        With rs
            While Not .EOF
                strFileName = !LocalPath
                If fso.FileExists(!LocalPath) Then
                    Set f = fso.GetFile(!LocalPath)
                    f.Attributes = f.Attributes Or Archive
                    strOutcome = "OK"
                ElseIf fso.FileExists(!DailyScan) Then
                    Set f = fso.GetFile(!DailyScan)
                    f.Attributes = f.Attributes Or Archive
                    strOutcome = "OK"
                Else
                    strOutcome = "NOFILE"
                End If
                    
                miFileCnt = miFileCnt + 1
                If miFileCnt > 30 Then
                    lstFiles.RemoveItem (0)
                End If
                lstFiles.AddItem strOutcome & ";" & strFileName
                DoEvents
                DoEvents
                DoEvents
                DoEvents
                DoEvents
                .MoveNext
            Wend
        End With
    End If
    MsgBox "Update completed"
    
Exit_Sub:
    Set rs = Nothing
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox "Error in module " & strErrSource & vbCrLf & vbCrLf & strErrMsg
    GoTo Exit_Sub
End Sub

Private Sub cmsStop_Click()
    mbContinue = False
End Sub


Private Sub Form_Close()
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    Dim strScannedStation As String
    
    Me.Caption = "Set Archive Bit"
    
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

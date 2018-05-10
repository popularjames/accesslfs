Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Dim mstrLocalHoldPath As String
Dim mstrLocalPath As String
Dim mstrCalledFrom As String

Const CstrFrmAppID As String = "PageCountEntry"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Form_Close()
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
End Sub

Private Sub Form_Load()
    Dim iLocalAccountID As Integer
    lblImageName.Caption = ""
    Label5.Caption = ""
    
    If Me.OpenArgs() & "" <> "" Then
        mstrCalledFrom = Me.OpenArgs()
    End If
    
    mstrLocalHoldPath = "" & DLookup("LocalHoldPath", "SCANNING_Config", "AccountID = " & gintAccountID)
    If mstrLocalHoldPath = "" Then
        MsgBox "There is no local path setup for SCANNING_Config.  Please see IT."
    Else
        If Right(mstrLocalHoldPath, 1) <> "\" Then mstrLocalHoldPath = mstrLocalHoldPath & "\"
    End If
    
    mstrLocalPath = "" & DLookup("LocalPath", "SCANNING_Config", "AccountID = " & gintAccountID)
    If mstrLocalPath = "" Then
        MsgBox "There is no local path setup for SCANNING_Config.  Please see IT."
    Else
        If Right(mstrLocalPath, 1) <> "\" Then mstrLocalPath = mstrLocalPath & "\"
    End If
    
End Sub

Private Sub Text3_AfterUpdate()
    Dim rsImageValidation As New ADODB.RecordSet
    Dim strSQL As String
      
    Dim fso As FileSystemObject
    Dim strImgPath As String
    Dim strPDFFile As String
    Dim strTIFFile As String
    Dim bFileExists As Boolean
    Dim strImageFile As String
    Dim iPageCount As Integer

    
    On Error Resume Next
    
    ' Clear any previous results
    lblImageName.Caption = ""
    Label5.Caption = ""

    Set MyAdo = New clsADO
    strSQL = " SELECT * from SCANNING_Image_Log_TMP where ImageName = '" & Mid(Me.Text3, 6) & "'"
    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
    MyAdo.sqlString = strSQL
    
    'open the audit claims header and disconnect

    lblImageName.Caption = Mid(Me.Text3, 6)

    Set rsImageValidation = MyAdo.OpenRecordSet()
    
    With rsImageValidation
        ' Check if record return any result
        If .EOF And .BOF Then
            Me.Label5.Caption = "This file is not on the daily scan log."
            Me.Label5.ForeColor = vbRed
            Beep
            GoTo Exit_Sub
        End If

'' TK: 8/5/2011 Remove file checking
'        Set fso = CreateObject("scripting.filesystemobject")
'
'        ' Check if file is (PDF or TIFF) and exists
'        strImageFile = mstrLocalHoldPath & !CnlyProvID & "\" & !ImageName & ".tif"
'        bFileExists = fso.FileExists(strImageFile)
'
'        If bFileExists = False Then
'            strImageFile = mstrLocalHoldPath & !CnlyProvID & "\" & !ImageName & ".pdf"
'            bFileExists = fso.FileExists(strImageFile)
'        End If
'
'        If bFileExists = False Then
'            strImageFile = mstrLocalPath & !CnlyProvID & "\" & !ImageName & ".tif"
'            bFileExists = fso.FileExists(strImageFile)
'        End If
'
'        If bFileExists = False Then
'            strImageFile = mstrLocalPath & !CnlyProvID & "\" & !ImageName & ".pdf"
'            bFileExists = fso.FileExists(strImageFile)
'        End If
'
'        ' display error message if file not exists
'        If Not bFileExists Then
'            MsgBox "There is an entry in daily scan. However, the image does NOT EXIST." & vbCrLf & _
'            "No file: " & Space(5) & "Image " & strImageFile, vbCritical
'            GoTo Exit_Sub
'        End If

        ' Display curent pagecount before user's input
        Me.Label5.Caption = "PageCount: " & !PageCnt
        Me.Label5.ForeColor = vbBlue
        
        ' Prompt user for page count
         iPageCount = 0
         Do
            iPageCount = InputBox(Prompt:="Enter page count for this record: ")
            If Nz(iPageCount, "") = "" Then
                Exit Do
            End If
            
            If (IsNumeric(iPageCount)) Then
                If val(iPageCount) > 0 Then
                    !PageCnt = iPageCount
                    MyAdo.BatchUpdate rsImageValidation
                    Me.Label5.Caption = "PageCount: " & !PageCnt
                    Me.Label5.ForeColor = vbBlue
                    Beep
                    Exit Do
                Else
                    MsgBox "Please enter a valid page count"
                End If
            Else
                MsgBox "Please enter a valid page count"
            End If
         Loop
    End With

    
Exit_Sub:
    Set fso = Nothing
    Set MyAdo = Nothing
    Set rs = Nothing
    Exit Sub

End Sub



Private Sub Text3_Enter()
    Me.Text3 = ""
End Sub

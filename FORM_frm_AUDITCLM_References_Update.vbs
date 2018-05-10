Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrCurrErrorCode As String
Private mstrCurrImageType As String
Private miCurrPageCount As Integer

Public Event UpdateReferences(ErrorCode As String, NewImageType As String, NewPageCount As Integer, Comment As String)

Const CstrFrmAppID As String = "AuditClmRef"


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Let CurrErrorCode(data As String)
     mstrCurrErrorCode = data
End Property


Property Let CurrImageType(data As String)
     mstrCurrImageType = data
End Property


Property Let CurrPageCount(data As Integer)
     miCurrPageCount = data
End Property


Public Sub RefreshScreen()
    Dim strError As String
    On Error GoTo ErrHandler
    
    If mstrCurrErrorCode & "" <> "" Then
        Me.ErrorCode.RowSource = "select * from SCANNING_Image_Error_Code where ErrorCd <> '" & mstrCurrErrorCode & "'"
    Else
        Me.ErrorCode.RowSource = "select * from SCANNING_Image_Error_Code"
    End If
    
    If mstrCurrImageType & "" <> "" Then
        Me.NewImageType.RowSource = "select ImageType, ImageTypeDisplay, ImageTypeDesc from v_SCANNING_XREF_ImageType_Account where Active = 'Y' and DataEntryVisible = 'Y' AND AccountID = " & str(gintAccountID) & " and ImageType <> '" & mstrCurrImageType & "'"
    Else
        Me.NewImageType.RowSource = "select ImageType, ImageTypeDisplay, ImageTypeDesc from v_SCANNING_XREF_ImageType_Account where Active = 'Y' and DataEntryVisible = 'Y' AND AccountID = " & str(gintAccountID)
    End If
    
    
    Select Case UCase(Me.ErrorCode)
        Case "PAGECNT"
            Me.lblNewImageType.visible = False
            Me.NewImageType.visible = False
            Me.NewImageType = ""
            
            Me.lblNewPageCount.visible = True
            Me.NewPageCount.visible = True
        Case "IMAGETYPE"
            Me.lblNewImageType.visible = True
            Me.NewImageType.visible = True
            
            Me.lblNewPageCount.visible = False
            Me.NewPageCount.visible = False
            Me.NewPageCount = ""
        Case Else
            Me.lblNewImageType.visible = False
            Me.NewImageType.visible = False
            Me.NewImageType = ""
            
            Me.lblNewPageCount.visible = False
            Me.NewPageCount.visible = False
            Me.NewPageCount = ""
    End Select

exitHere:
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub


Private Sub CmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdOk_Click()
    Dim strErrMsg As String
    
    
    If Me.Comment & "" = "" Then
        MsgBox "Error: Comment can not be blank.", vbInformation
        Exit Sub
    End If
    
    If Me.ErrorCode = "PAGECNT" Then
        strErrMsg = ""
        If IsNumeric(Me.NewPageCount & "") Then
            Select Case val(Me.NewPageCount & "")
                Case Is = miCurrPageCount
                    strErrMsg = "Error: New page count is the same as current page count"
                Case Is < 1
                    strErrMsg = "Error: New page count can not be less than 1"
            End Select
        Else
            strErrMsg = "Error: Please enter a valid page count"
        End If
        
        If strErrMsg <> "" Then
            MsgBox strErrMsg, vbInformation
            Exit Sub
        End If
    End If
    
    RaiseEvent UpdateReferences(Me.ErrorCode, Me.NewImageType, val(Me.NewPageCount), Me.Comment)
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub ErrorCode_AfterUpdate()
    Call RefreshScreen
End Sub


Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub


Private Sub Form_Load()
    
    Dim iAppPermission As Integer
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    Me.Caption = "AuditClm_References Update"
    
    Call RefreshScreen
End Sub

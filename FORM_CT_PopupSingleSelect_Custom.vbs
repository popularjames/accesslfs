Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mvTitle As String
Private MvListTitle As String
Private MvStartupWidth As Single
Private MvCmboDefault As String
Private MvCmboTitle As String
Private bRequired As Boolean

Private MvChkCaption1 As String
Private MvChkCaption2 As String
Private MvChkCaption3 As String
Private MvChkCaption4 As String
Private MvChkCaption5 As String
Private MvChkCaption6 As String

Public Results As Integer
Public MvSelections As Collection
Public CmboResults As Variant

Private Sub Form_Load()
    Me.FormHeader.visible = False
End Sub

Public Property Let CmboTitle(data As String)
    MvCmboTitle = data
    Me.LblCmbo.Caption = MvCmboTitle
    Me.LblCmbo.visible = True
    Me.Cmbo.visible = True
    Me.FormHeader.visible = True
End Property
Public Property Get CmboTitle() As String
    CmboTitle = MvCmboTitle
End Property
Public Property Let CmboDefault(data As String)
    MvCmboDefault = data
    Me.Cmbo = MvCmboDefault
    Cmbo_AfterUpdate
End Property
Public Property Get CmboDefault() As String
    CmboDefault = MvCmboDefault
End Property

Public Property Let ChkCaption1(data As String)
    MvChkCaption1 = data
    Me.LblChk1.Caption = MvChkCaption1
    Me.LblChk1.visible = True
    Me.Check1 = -1
    Me.Check1.visible = True
End Property
Public Property Get ChkCaption1() As String
    ChkCaption1 = MvChkCaption1
End Property
Public Property Let ChkCaption2(data As String)
    MvChkCaption2 = data
    Me.LblChk2.Caption = MvChkCaption2
    Me.LblChk2.visible = True
    Me.Check2 = -1
    Me.Check2.visible = True
End Property
Public Property Get ChkCaption2() As String
    ChkCaption2 = MvChkCaption2
End Property
Public Property Let ChkCaption3(data As String)
    MvChkCaption3 = data
    Me.LblChk3.Caption = MvChkCaption3
    Me.LblChk3.visible = True
    Me.Check3 = -1
    Me.Check3.visible = True
End Property
Public Property Get ChkCaption3() As String
    ChkCaption3 = MvChkCaption3
End Property
Public Property Let ChkCaption4(data As String)
    MvChkCaption4 = data
    Me.LblChk4.Caption = MvChkCaption4
    Me.LblChk4.visible = True
    Me.Check4 = -1
    Me.Check4.visible = True
End Property
Public Property Get ChkCaption4() As String
    ChkCaption4 = MvChkCaption4
End Property
Public Property Let ChkCaption5(data As String)
    MvChkCaption5 = data
    Me.LblChk5.Caption = MvChkCaption5
    Me.LblChk5.visible = True
    Me.Check5 = -1
    Me.Check5.visible = True
End Property
Public Property Get ChkCaption5() As String
    ChkCaption5 = MvChkCaption5
End Property
Public Property Let ChkCaption6(data As String)
    MvChkCaption6 = data
    Me.LblChk6.Caption = MvChkCaption6
    Me.LblChk6.visible = True
    Me.Check6 = -1
    Me.Check6.visible = True
End Property
Public Property Get ChkCaption6() As String
    ChkCaption6 = MvChkCaption6
End Property

Public Property Let StartupWidth(data As Single)
On Error GoTo ErrorHappened
    Dim StSplit() As String, SgWidth As Single, X As Integer
    ' AUTO CALCULATE THE WIDTH BASED ON COLUMN WIDTHS
    If data = -1 Then
        StSplit = Split(Me.Lst.ColumnWidths, ";")
        For X = 0 To UBound(StSplit)
            SgWidth = SgWidth + CSng(StSplit(X))
        Next X
        SgWidth = SgWidth + (0.5 * 1440)
        If SgWidth < Me.InsideWidth Then
            data = Me.InsideWidth
        Else
            data = SgWidth
        End If
    End If
    MvStartupWidth = data
    'Me.Width = MvStartupWidth
    Me.InsideWidth = MvStartupWidth
    
ExitNow:
    On Error Resume Next
    Exit Property
ErrorHappened:
    MsgBox "Error Setting Startup For Width." & vbCrLf & vbCrLf & Err.Description, vbCritical, CodeContextObject.Name & ".StartupWidth"
    Resume ExitNow
End Property

Public Property Get Required() As Boolean
   Required = bRequired
End Property
Public Property Let Required(ByVal bReq As Boolean)
    bRequired = bReq
End Property
Public Property Get Selections() As Collection
   Set Selections = MvSelections
End Property
Public Property Let Title(data As String)
    mvTitle = data
    Me.Caption = mvTitle
End Property
Public Property Get Title() As String
    Title = mvTitle
End Property

Public Property Let ListTitle(data As String)
    MvListTitle = data
    Me.LblLst.Caption = MvListTitle
End Property
Public Property Get ListTitle() As String
    ListTitle = MvListTitle
End Property

Public Sub Cancel()
    Set MvSelections = Nothing
    Results = vbCancel
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Public Sub Ok()
    FillSelections
    Results = vbOK
    Me.visible = False
End Sub

Private Sub CmdCancel_Click()
    Cancel
End Sub

Private Sub cmdOk_Click()
    Ok
End Sub

Private Sub FillSelections()
On Error GoTo ErrorHappened
    Dim MyCol As New Collection
    Dim myAry() As String
    Dim MyItem As Variant, col As Integer

    'SET THE NUMBER OF COLUMNS
    ReDim myAry(Lst.ColumnCount) As String
    
    For Each MyItem In Lst.ItemsSelected
        For col = 0 To Lst.ColumnCount - 1
            myAry(col) = Lst.Column(col, CLng(MyItem))
        Next col
        MyCol.Add myAry, CStr(MyItem)
    Next MyItem
    Set MvSelections = MyCol
ExitNow:
    On Error Resume Next
    Set MyItem = Nothing
    Set MyCol = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, CodeContextObject.Name & ".FillSelections()"
    Resume ExitNow
End Sub

' make it so that OK is only enabled once they select something in the list
Private Sub Lst_AfterUpdate()
    If bRequired Then Me.CmdOK.Enabled = True
End Sub

Private Sub Lst_DblClick(Cancel As Integer)
    Ok
End Sub

Private Sub Cmbo_AfterUpdate()
    CmboResults = Me.Cmbo.ItemData(Me.Cmbo.ListIndex)
End Sub

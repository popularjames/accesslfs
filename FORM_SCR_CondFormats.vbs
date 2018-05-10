Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private MvScreenID As Long
Private mvFormId As Byte
Private MvScr As Form_SCR_MainScreens
Private Type CnlyFormat
    ID As Long
    FormatName As String
    FieldName As String
    BackColor As Long
    Expression1 As String
    Expression2 As String
    FontBold As Boolean
    FontItalic As Boolean
    FontUnderline As Boolean
    ForeColor As Long
    Operator As Byte
    Type As Byte
End Type
Private Const acAllFieldsExpression = 254
Private Const acRowHasFocus = 255

Public Property Get ActiveFormatCount() As Long
    ActiveFormatCount = LstApply.ListCount
End Property
Public Property Get SQLList() As String
    Dim X As Long, StIDs As String
    For X = 0 To LstApply.ListCount - 1
        If X <> LstApply.ListCount - 1 Then
            StIDs = StIDs & LstApply.Column(0, X) & ","
        Else
            StIDs = StIDs & LstApply.Column(0, X)
        End If
    Next X
    SQLList = StIDs
End Property
Private Sub CmdColor_Click()
TxtSample.BackColor = ChooseColor(TxtSample.BackColor, Me.hwnd)
SaveFormat
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHappened
Dim stName As String, lngId As Long

If Me.Dirty = True Then Me.Dirty = False
If Me.Lst.ListIndex = -1 Then
    MsgBox "You must select a Conditional Format to delete.", vbInformation, "Delete Conditional Format"
    GoTo ExitNow
End If

stName = Me.Lst.Column(1, Me.Lst.ListIndex)
lngId = Me.Lst.Column(0, Me.Lst.ListIndex)
If MsgBox("Delete Conditional Format '" & stName & "'?", vbQuestion + vbDefaultButton2 + vbOKCancel, "Confirm Delete") = vbOK Then
    CurrentDb.Execute ("Delete * From SCR_ScreensCondFormats Where CondFormatID = " & CStr(lngId))
    Me.Lst.Requery
    DBEngine.Idle dbRefreshCache + dbForceOSFlush
    Me.Lst.Requery
    Me.Requery
    
End If



ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
    Resume
End Sub

Private Sub CmdFont_Click()
Dim cls As New CT_ClsFont

With cls
    .PropertiesFromControl TxtSample
    .ShowEffects = True
    .ShowSize = False 'NO FONT NAME CHANGE ALLOWED IN DATASHEET
    If .DialogFont = True Then
        With TxtSample
              .ForeColor = cls.color
              '.FontSize = Cls.Height 'NO FONT NAME CHANGE ALLOWED IN DATASHEET
              .FontWeight = cls.Weight
              .FontItalic = cls.Italic
              .FontUnderline = cls.UnderLine
              'Ctl.FontName = .Name - NO FONT NAME CHANGE ALLOWED IN DATASHEET
        End With
        SaveFormat
    End If
End With

Set cls = Nothing
End Sub


Private Sub CmdFormatAdd_Click()
If Lst.ListIndex <> -1 Then
    ApplyFormatAdd Lst.Column(0, Lst.ListIndex), Lst.Column(1, Lst.ListIndex)
End If
End Sub

Private Sub CmdFormatApply_Click()
If Me.Dirty = True Then Me.Dirty = False
ApplyFormats
End Sub

Private Sub CmdFormatDelete_Click()
    ApplyFormatClear
End Sub

Private Sub CmdFormatRemove_Click()
If LstApply.ListIndex <> -1 Then
    ApplyFormatRemove LstApply.Column(0, LstApply.ListIndex)
End If
End Sub

Private Sub cmdNew_Click()
On Error GoTo ErrorHappened
Dim db As DAO.Database
Dim stName As String
Dim SQL As String
Dim NewID As Long



stName = InputBox("Enter New Name For Conditional Format", "New Format")
If "" & stName <> "" Then
    SQL = "Insert Into SCR_ScreensCondFormats(ScreenID, FormatName, FieldName, Computer,UserName) "
    SQL = SQL & "Values(" & MvScreenID & ", " & Chr(34) & stName & Chr(34) & ",'NA', "
    SQL = SQL & Chr(34) & Identity.Computer & Chr(34) & ", " & Chr(34) & Identity.UserName & Chr(34) & ")"
    Set db = CurrentDb
    db.Execute SQL, dbFailOnError
    NewID = DLookup("CondFormatID", "SCR_ScreensCondFormats", "ScreenID =" & MvScreenID & " and FormatName = " & Chr(34) & stName & Chr(34))
    Me.Lst.Requery
    If Me.RecordSet.BOF And Me.RecordSet.EOF Then
        Me.Undo
    End If
    Me.Requery
    Lst.Value = NewID
    Lst_AfterUpdate
End If

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
    Resume
End Sub

Private Sub cmdRefresh_Click()
Lst.Requery
End Sub

Private Sub CmdRename_Click()
On Error GoTo ErrorHappened
Dim stName As String, lngId As Long
Dim StNewName As String

If Me.Lst.ListIndex = -1 Then
    MsgBox "You must select a format to rename.", vbInformation, "Rename format"
    GoTo ExitNow
End If

stName = "" & Me.Lst.Column(1, Me.Lst.ListIndex)
lngId = Me.Lst.Value

StNewName = InputBox("Please enter a new name for '" & stName & "'", "Rename format", stName)
If "" & StNewName = "" Or "" & StNewName = stName Then
    GoTo ExitNow
End If

CurrentDb.Execute ("Update SCR_ScreensCondFormats Set FormatName = " & Chr(34) & StNewName & Chr(34) & " Where CondFormatID = " & CStr(lngId))
Me.Lst.Requery
Me.Requery
Me.Lst.Value = lngId
Lst_AfterUpdate

ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
    Resume
End Sub

Private Sub SaveFormat()
With TxtSample
    Me!ForeColor = .ForeColor
    Me!BackColor = .BackColor
    'Me!FontSize = .FontSize
    If .FontWeight = 700 Then 'BOLD
        Me!FontBold = True
    Else
        Me!FontBold = False
    End If
    Me!FontItalic = .FontItalic
    Me!FontUnderline = .FontUnderline
    Lst.SetFocus
End With
End Sub





Private Sub CmdZoom1_Click()
ZoomText Me.Expression1, "Expression"
End Sub

Private Sub CmdZoom2_Click()
ZoomText Me.Expression2, "Between"
End Sub

Private Sub Form_Current()
On Error GoTo ErrorHappened
Dim BlEnabled As Boolean

If "" & Me.RecordSource = "" Then
    Exit Sub
End If

If Nz(Me!CondFormatID, 0) = 0 Then
    BlEnabled = False
ElseIf (Me.RecordsetClone.EOF Or Me.RecordsetClone.BOF) Then
    BlEnabled = False
Else
    BlEnabled = True
End If

If BlEnabled = False Then
    Me.CmdColor.Enabled = BlEnabled
    Me.CmdFont.Enabled = BlEnabled
    Me.Operator.Enabled = BlEnabled
    Me.FieldName.Enabled = BlEnabled
    Me.Expression1.Enabled = BlEnabled
    Me.Expression2.Enabled = BlEnabled
    Me.Type.Enabled = BlEnabled
Else
    Type_AfterUpdate
End If

With TxtSample
    .ForeColor = Nz(Me!ForeColor, 0)
    .BackColor = Nz(Me!BackColor, RGB(255, 255, 255))
    'Me!FontSize = .FontSize
    If Nz(Me!FontBold, False) = True Then 'BOLD
        .FontWeight = 700
    Else
        .FontWeight = 400
    End If
    .FontItalic = Nz(Me!FontItalic, False)
   .FontUnderline = Nz(Me!FontUnderline, False)
End With

ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error ON Current Record", vbCritical, "Conditional Formats"
    Resume ExitNow
    Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set MvScr = Nothing
End Sub


Private Sub Lst_AfterUpdate()
On Error GoTo ErrorHappened
Dim rst As DAO.RecordSet

If Lst.ListIndex = -1 Then
    GoTo ExitNow
End If

Set rst = Me.RecordsetClone

With rst
    .MoveFirst
    .FindFirst "CondFormatID = " & Lst.Column(0, Lst.ListIndex)
    If .NoMatch = False Then
        Me.Bookmark = .Bookmark
    End If
End With
ExitNow:
    On Error Resume Next
    Set rst = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error Selecting Record", vbCritical, "Conditional Formats"
    Resume ExitNow
    Resume
End Sub


Public Sub InitData(ScreenID As Long, FormID As Byte)
On Error GoTo LoadError


If Me.Parent Is Nothing Then
    MsgBox "This form is not meant to run independatly of a screen!", vbCritical
    DoCmd.Close acForm, Me.Name, acSaveNo
    GoTo ExitNow
Else
    MvScreenID = ScreenID
    mvFormId = FormID
    Lst.RowSource = "SELECT CondFormatID, FormatName FROM SCR_ScreensCondFormats Where ScreenID = " & MvScreenID & " ORDER BY FormatName "
    Set MvScr = Scr(mvFormId)
    Me.FieldName.RowSource = MvScr.Config.PrimaryRecordSource
    Me.RecordSource = "SELECT * FROM SCR_ScreensCondFormats Where ScreenID = " & MvScreenID
    Me.Requery
End If

ExitNow:
    Exit Sub
LoadError:
    MsgBox Err.Description, vbCritical, "Error Loading Conditional Formats"
    'DoCmd.Close acForm, Me.Name, acSaveNo
    Resume ExitNow
    Resume
End Sub
Public Sub ApplyFormatAdd(FormatID As Long, FormatName As String)
Dim X As Long
For X = 0 To LstApply.ListCount - 1
    If LstApply.Column(0, X) = FormatID Then
        Exit Sub
    End If
Next X
LstApply.AddItem (FormatID & ";" & Chr(34) & FormatName & Chr(34))
End Sub
Public Sub ApplyFormatRemove(FormatID As Long)
Dim X As Long
LstApply.Value = Null
For X = 0 To LstApply.ListCount - 1
    If LstApply.Column(0, X) = FormatID Then
        LstApply.RemoveItem (X)
        Exit For
    End If
Next X
LstApply.Requery
DoEvents
End Sub
Public Sub ApplyFormatClear()
Dim X As Long

If LstApply.ListCount > 0 Then
    For X = LstApply.ListCount - 1 To 0 Step -1
        LstApply.RemoveItem (X)
    Next X
End If
End Sub
Private Sub Lst_DblClick(Cancel As Integer)
If Lst.ListIndex <> -1 Then
    ApplyFormatAdd Lst.Column(0, Lst.ListIndex), Lst.Column(1, Lst.ListIndex)
End If
End Sub

Private Sub LstApply_DblClick(Cancel As Integer)
If LstApply.ListIndex <> -1 Then
    LstApply.RemoveItem LstApply.ListIndex
    LstApply.Requery
    DoEvents
End If
End Sub
Public Sub ApplyFormats()
Dim X As Long, StFormatIDs As String
Dim SQL As String, db As DAO.Database, rst As DAO.RecordSet
Dim TxtFld As TextBox, LocFormat As CnlyFormat ' FmtCdt As FormatCondition
DoCmd.Hourglass True

If Me.Dirty = True Then ' Save Changes first
    Me.Dirty = False
End If

'FIRST REMOVE ALL OF THE EXISTING FORMATS
MvScr.GridForm.FormatsClear

'NOW FIGURE OUT WHAT NEEDS TO BE APPLIED -SQL STRING
With LstApply
    If .ListCount > 0 Then
        StFormatIDs = Me.SQLList
    Else
        GoTo ExitNow
    End If
End With

'GET THE VALUES FROM THE DATABASE
SQL = "Select * From SCR_ScreensCondFormats Where CondFormatID in (" & StFormatIDs & ")"
DBEngine.Idle dbRefreshCache + dbForceOSFlush
DoEvents
Set db = CurrentDb
Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)
With rst
    Do Until .EOF
        LocFormat.ID = .Fields("CondFormatID")
        LocFormat.FormatName = .Fields("FormatName")
        LocFormat.FieldName = "" & .Fields("FieldName")
        LocFormat.Type = .Fields("Type")
        LocFormat.ForeColor = Nz(.Fields("ForeColor"), 0)
        LocFormat.FontBold = Nz(.Fields("FontBold"), False)
        LocFormat.BackColor = Nz(.Fields("BackColor"), RGB(255, 255, 255))
        LocFormat.FontItalic = Nz(.Fields("FontItalic"), False)
        LocFormat.FontUnderline = Nz(.Fields("FontUnderline"), False)
        LocFormat.Operator = Nz(.Fields("Operator"), 0) ' acbetween
        LocFormat.Expression1 = "" & .Fields("Expression1")
        LocFormat.Expression2 = "" & .Fields("Expression2")

        
        
        'FIND THE FIELD TO ADD THE FORMAT TO
        With MvScr.GridForm
             For X = 1 To 255
                Set TxtFld = .Controls("Field" & CStr(X))
                With TxtFld
                    If LocFormat.Type = acAllFieldsExpression Then
                        If "" & TxtFld.ControlSource = "" Then 'ONLY GO UNTIL THERE IS NO CONTROL SOURCE
                            Exit For
                        End If
                        ApplyFormatsField TxtFld, LocFormat
                    ElseIf LocFormat.Type = acRowHasFocus Then
                        If "" & TxtFld.ControlSource = "" Then 'ONLY GO UNTIL THERE IS NO CONTROL SOURCE
                            Exit For
                        End If
                        ApplyFormatsField TxtFld, LocFormat
                   ElseIf UCase("" & TxtFld.ControlSource) = UCase(LocFormat.FieldName) Then
                        ApplyFormatsField TxtFld, LocFormat
                        Exit For
                    ElseIf "" & TxtFld.ControlSource = "" Then 'ONLY GO UNTIL THERE IS NO CONTROL SOURCE
                        Exit For
                    End If
                End With
             Next X
        End With
        
NextRecord:
        Set TxtFld = Nothing
        .MoveNext
    Loop
    .Close
End With

ExitNow:
    On Error Resume Next
    'MvScr.GridForm.Recalc
    DoCmd.Hourglass False
    Set rst = Nothing
    Set db = Nothing
    Set TxtFld = Nothing

    Exit Sub
End Sub

Private Function ApplyFormatsField(TxtFld As Access.TextBox, Format As CnlyFormat) As String
On Error GoTo ErrorHappened
Dim FmtCdt As FormatCondition
If Not TxtFld Is Nothing Then
    Select Case Format.Type
    Case acExpression, acAllFieldsExpression
        Set FmtCdt = TxtFld.FormatConditions.Add(acExpression, , Format.Expression1)
    Case acFieldHasFocus
        Set FmtCdt = TxtFld.FormatConditions.Add(acFieldHasFocus)
    Case acRowHasFocus
        Set FmtCdt = TxtFld.FormatConditions.Add(acFieldHasFocus)
    Case acFieldValue
        Set FmtCdt = TxtFld.FormatConditions.Add(acFieldValue, Format.Operator, Format.Expression1, Format.Expression2)
    End Select

    FmtCdt.ForeColor = Format.ForeColor
    FmtCdt.FontBold = Format.FontBold
    FmtCdt.BackColor = Format.BackColor
    FmtCdt.FontItalic = Format.FontItalic
    FmtCdt.FontUnderline = Format.FontUnderline

    ApplyFormatsField = ""
Else
    ApplyFormatsField = "Null Field Reference Passed in to 'ApplyFormatsField'"
End If
ExitNow:
    On Error Resume Next
    Set FmtCdt = Nothing
    Exit Function
ErrorHappened:
    ApplyFormatsField = "ApplyFormatsField: " & Err.Description
    Resume ExitNow
    Resume
End Function
Private Sub ZoomText(ctl As TextBox, Optional ByVal Title As String = "")
Dim FrmZoom As Form_CT_Text
Dim Txt As String

Txt = "" & ctl

Set FrmZoom = New Form_CT_Text
FrmZoom.Move ctl.left
With FrmZoom
    .Text = Txt
    .Title = Title
    .visible = True
    Do Until .Results <= 0
        DoEvents
    Loop
    If .Results = True Then
        ctl = "" & .Text
    End If
End With

End Sub

Private Sub Operator_AfterUpdate()
    Type_AfterUpdate
End Sub

Private Sub Type_AfterUpdate()
Me.Type.Enabled = True
Me.Lst.SetFocus 'Move Focus to control fields

Me.CmdColor.Enabled = True
Me.CmdFont.Enabled = True
    
Select Case Me.Type
Case acExpression
    Me.FieldName.Enabled = True
    Me.CmdZoom1.Enabled = True
    Me.Expression1.Enabled = True
    Me.Operator.Enabled = False
    Me.Expression2.Enabled = False
    Me.CmdZoom2.Enabled = False
Case acFieldHasFocus
    Me.FieldName.Enabled = True
    Me.Operator.Enabled = False
    Me.Expression1.Enabled = False
    Me.Expression2.Enabled = False
    Me.CmdZoom1.Enabled = False
    Me.CmdZoom2.Enabled = False
Case acRowHasFocus
    Me.FieldName.Enabled = False
    Me.Operator.Enabled = False
    Me.Expression1.Enabled = False
    Me.Expression2.Enabled = False
    Me.CmdZoom1.Enabled = False
    Me.CmdZoom2.Enabled = False
Case acFieldValue
    Me.FieldName.Enabled = True
    Me.Expression1.Enabled = True
    Me.CmdZoom1.Enabled = True
    Me.Operator.Enabled = True
    Me.Expression1.Enabled = True
    If Me.Operator = acNotBetween Or Me.Operator = acBetween Then
        Me.Expression2.Enabled = True
        Me.CmdZoom2.Enabled = True
    Else
        Me.Expression2.Enabled = False
        Me.CmdZoom2.Enabled = False
    End If
Case acAllFieldsExpression
    Me.FieldName.Enabled = False
    Me.CmdZoom1.Enabled = True
    Me.Expression1.Enabled = True
    Me.Operator.Enabled = False
    Me.Expression2.Enabled = False
    Me.CmdZoom2.Enabled = False
End Select
End Sub

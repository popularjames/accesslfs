Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'SA 03/22/2012 - CR2609 Moved and resized various screen objects

Private MvScreenID As Long
Private WithEvents MvGridMain As Form_CT_SubGenericDataSheet
Attribute MvGridMain.VB_VarHelpID = -1


Private Sub CmdCopyExisting_Click()
On Error GoTo ErrorHappened
Dim frm As New Form_CT_PopupSelect
Dim SQL As String

SQL = "SELECT Max(CSFF.FieldFormatID) AS FieldFormatID, CS.ScreenName, CSFF.RecordSource, Sum(1) AS FieldCT "
SQL = SQL & "FROM SCR_Screens as CS "
SQL = SQL & "INNER JOIN SCR_ScreensFieldFormats as CSFF ON "
SQL = SQL & "CS.ScreenID = CSFF.ScreenID "
If "" & Me!Src <> "" Then
    SQL = SQL & "WHERE CS.ScreenID <> " & Me.ScreenIDCurrent & " or "
    SQL = SQL & "CSFF.RecordSource <> " & Chr(34) & Me!Src & Chr(34) & " "
End If
SQL = SQL & "GROUP BY CS.ScreenName, CSFF.RecordSource "


With frm
    With .Lst
        .RowSource = SQL
        .ColumnHeads = True
        .BoundColumn = 1
        .ColumnCount = 4
        .ColumnWidths = "0;1.5" & Chr(34) & ";1.5" & Chr(34) & ";0.5" & Chr(34)

        .Requery
    End With
    .Title = "Copy Formats"
    .ListTitle = "Select Recordsource(s)"
    .StartupWidth = -1  'AUTO SIZE TO COLUMN WIDTHS '4 * 1440
    .visible = True

    Do While .Results = vbApplicationModal
        DoEvents
    Loop

    If .Results = vbOK Then
        Dim X As Long, MyCol As Collection
        SQL = ""
        Set MyCol = .Selections
        DoCmd.Hourglass True
        If Not MyCol Is Nothing Then
            For X = 1 To MyCol.Count
                SQL = "INSERT INTO SCR_ScreensFieldFormats (ScreenID, RecordSource, FieldName, Alias, FieldWidth, Align, Format, Decimals ) "
                SQL = SQL & "SELECT " & Me.ScreenIDCurrent & " as ScreenID, " & Chr(34) & Me!Src & Chr(34) & " as RecordSource, T1.FieldName, T1.Alias, T1.FieldWidth, T1.Align, T1.Format, T1.Decimals "
                SQL = SQL & "FROM SCR_ScreensFieldFormats AS T1 INNER JOIN SCR_ScreensFieldFormats AS T2 ON (T1.RecordSource = T2.RecordSource) AND (T1.ScreenID = T2.ScreenID) "
                SQL = SQL & "WHERE T2.FieldFormatID=" & MyCol.Item(X)(0) & " AND "
                SQL = SQL & "NOT T1.FieldName in (Select FieldName From SCR_ScreensFieldFormats Where ScreenID = " & Me.ScreenIDCurrent & " AND RecordSource = " & Chr(34) & Me!Src & Chr(34) & ") "
                CurrentDb.Execute SQL, dbFailOnError
NextITEM:
            Next X
        End If
    End If
End With
ExitNow:
    On Error Resume Next
    Me.SfFormatFields.Requery
    DoCmd.Hourglass False
    Set MyCol = Nothing
    Set frm = Nothing
    Exit Sub

ErrorHappened:
    Select Case Err.Number
    Case 3022 ' Duplicate Index
        If MsgBox("Failed Adding Field Formats From  (" & MyCol.Item(X)(1) & " --> " & MyCol.Item(X)(3) & ") because a field already exists." & vbCrLf & vbCrLf & "Would you like to continue adding field formats?", vbQuestion + vbDefaultButton1 + vbYesNo, "Add Failed --> " & CodeContextObject.Name) = vbYes Then
            Resume NextITEM
        Else
            Resume ExitNow
        End If
    Case Else
        MsgBox Err.Description, vbCritical, "Error Adding Field Formats --> " & CodeContextObject.Name
        Resume ExitNow
    End Select
    Resume
    

End Sub


Private Sub cmdRefresh_Click()
    Me.Lst.Requery
End Sub

Private Sub cmdLoad_Click()
On Error GoTo ErrorHappened
Dim SQL As String
Dim stTop As String
Dim stWhere As String

DoCmd.Hourglass True

If Me.txtTop > 0 Then
    stTop = "Top " & Me.txtTop
End If
If Nz(Me.txtWhere, "") <> "" Then
    stWhere = "Where " & Me.txtWhere
End If
SQL = "Select " & stTop & " * From " & Me.Lst.Value & " " & stWhere
Me.SubFormSizeFields.Form.RecordSource = SQL

ExitNow:
    On Error Resume Next
    DoCmd.Hourglass False
    Exit Sub
    
ErrorHappened:
    If Err.Number = 3146 Then
        MsgBox "Timeout Error Returning Rows." & vbCrLf & vbCrLf & "Try Entering a Filter in the Where Clause."
    Else
        MsgBox Err.Description
    End If
    Resume ExitNow
    Resume
End Sub

Private Sub cmdSave_Click()
On Error GoTo CmdSaveError

Dim RecSrc As String
Dim db As DAO.Database
Dim rst As DAO.RecordSet
Dim SQL As String
Dim X As Long
Dim TxtFld As Access.TextBox

DoCmd.Hourglass (True)

Set db = CurrentDb

RecSrc = Me.Lst.Value
Set MvGridMain = Me.SubFormSizeFields.Form

With MvGridMain
    For X = 1 To .FldCT
        Set TxtFld = .Controls("Field" & CStr(X))
        SQL = "SELECT *"
        SQL = SQL & " FROM SCR_ScreensFieldFormats"
        SQL = SQL & " WHERE ScreenId = " & Me.Parent!ScreenID
        SQL = SQL & " AND RecordSource = " & Chr(34) & RecSrc & Chr(34)
        SQL = SQL & " AND FieldName = " & Chr(34) & TxtFld.ControlSource & Chr(34)
        
        
        Set rst = db.OpenRecordSet(SQL, dbOpenDynaset)
        
        If Not rst.EOF Then
            rst.Edit
            rst!FieldWidth = IIf(TxtFld.ColumnHidden = True, 0, Round(TxtFld.ColumnWidth / 1440, 2))
            rst.Update
        Else
            rst.AddNew
            rst!ScreenID = Me.Parent!ScreenID
            rst!RecordSource = RecSrc
            rst!FieldName = TxtFld.ControlSource
            rst!FieldWidth = IIf(TxtFld.ColumnHidden = True, 0, Round(TxtFld.ColumnWidth / 1440, 2))
            rst.Update
        End If
    Next X
End With

Me.SfFormatFields.Form.Requery

CmdSaveExit:
    DoCmd.Hourglass False
    On Error Resume Next
    rst.Close
    db.Close
    Exit Sub
    
CmdSaveError:
    MsgBox Err.Description
    Resume CmdSaveExit
    Resume
End Sub

Private Sub Form_Current()
On Error Resume Next

'GET THE CURRENT SCREEN ID PROPERTIES
If Nz(Me!ScreenID, 0) = 0 Then
    If Me.Parent.Name = "SCRCFG_Screens" Then
        Me.ScreenIDCurrent = Nz(Me.Parent!ScreenID, 0)
    Else
        Me.ScreenIDCurrent = 0
    End If
Else
   Me.ScreenIDCurrent = Nz(Me!ScreenID, 0)
End If
'Change The Tabbed SubformFields List Sources
ChangeRecordSource
End Sub

Public Property Let ScreenIDCurrent(data As Long)
Dim SQL As String

    If MvScreenID <> data Then
        MvScreenID = data
        SQL = "SELECT ScreenID,PrimaryRecordSource as Src FROM SCR_Screens "
        SQL = SQL & "Where ScreenID = " & MvScreenID & " "
        SQL = SQL & "UNION "
        SQL = SQL & "Select ScreenID, RecordSource From SCR_ScreensTabs "
        SQL = SQL & "Where ScreenID = " & MvScreenID & " "
        SQL = SQL & "ORDER BY Src "
        ' added the secondary, and tertiary record source choices - PrimaryListBoxRecordSource, SecondaryListBoxRecordSource, TertiaryListBoxRecordSource
        SQL = SQL & " UNION "
        SQL = SQL & " SELECT ScreenId, PrimaryListBoxRecordSource as Scr FROM SCR_Screens "
        SQL = SQL & " WHERE ScreenId = " & MvScreenID & " AND PrimaryListBoxMulti = -1 "
        SQL = SQL & " UNION "
        SQL = SQL & " SELECT ScreenId, SecondaryListBoxRecordSource as Scr FROM SCR_Screens "
        SQL = SQL & " WHERE ScreenId = " & MvScreenID & " AND (SecondaryListBoxUse = -1 AND SecondaryListBoxMulti = -1)"
        SQL = SQL & " UNION "
        SQL = SQL & " SELECT ScreenId, TertiaryListBoxRecordSource as Scr FROM SCR_Screens "
        SQL = SQL & " WHERE ScreenId = " & MvScreenID & " AND (TertiaryListBoxUse = -1 AND TertiaryListBoxMulti = -1) "
        Lst.RowSource = SQL
    End If
End Property

Public Property Get ScreenIDCurrent() As Long
 ScreenIDCurrent = MvScreenID
End Property

Private Sub Lst_AfterUpdate()
Dim rst As DAO.RecordSet
Dim db As DAO.Database
Dim TbDef As TableDef
Dim QryDef As QueryDef
Dim RecType As Byte
Dim TmpSub As String

Set rst = Me.RecordsetClone
With rst
    .MoveFirst
    .FindFirst "[Src] =" & Chr(34) & Me.Lst.Value & Chr(34)
    If .NoMatch = False Then
        Me.Bookmark = .Bookmark
    End If
End With

'************************************************************************
'CODE FOR RESIZING FIELDS - ADD BY LINO
TmpSub = Me.SubFormSizeFields.SourceObject
Me.SubFormSizeFields.SourceObject = ""

Set db = CurrentDb
RecType = 99
db.TableDefs.Refresh
For Each TbDef In db.TableDefs
    If TbDef.Name = Me.Lst.Value Then
        RecType = 0
        Exit For
    End If
Next TbDef

db.QueryDefs.Refresh
For Each QryDef In db.QueryDefs
    If QryDef.Name = Me.Lst.Value Then
        RecType = 1
        Exit For
    End If
Next QryDef

    Me.SubFormSizeFields.SourceObject = TmpSub
If RecType <> 99 Then
    Me.SubFormSizeFields.Form.InitData Me.Lst.Value, RecType
End If
'************************************************************************

End Sub


Sub ChangeRecordSource()
On Error GoTo ChangeRecordSourceError


With Me.SfFormatFields
    .Form!FieldName.RowSource = Me!Src
End With

ChangeRecordSourceExit:
    On Error Resume Next
    Exit Sub
    
ChangeRecordSourceError:
    Select Case Err.Number
    Case 13 'Type MisMatch
        Resume ChangeRecordSourceExit
    Case Else
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error getting field list!", vbInformation
        Resume ChangeRecordSourceExit
        Resume
    End Select
End Sub

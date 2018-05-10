Option Compare Database
Option Explicit


Private Const ToolbarName As String = "Screens"
Private Const ToolbarPopupDC As String = "Data Center"
'
'
'Public Function CommandBarOpenScreen(DropDownMenuItemID As Integer)
'On Error GoTo OpenFormError
'Dim combo As CommandBarComboBox
'
'Set combo = CommandBars(ToolbarName).FindControl(tag:=DropDownMenuItemID, recursive:=True)
'
'If "" & combo.text <> "" Then
'    DoCmd.Hourglass True
'    Dim ClsScr As ClsScreenData
'    Set ClsScr = New ClsScreenData
'    With ClsScr
'        .CreateScreen combo.text
'        If .NewScreen = True Then
'            .GetConfig
'            RunEvent "Screen Load", .ScreenForm.ScreenID, .ScreenForm.FormID
'        End If
'    End With
'Else
'    MsgBox "You must select a screen to open first!"
'End If
'
'
'
'OpenFormExit:
'    On Error Resume Next
'    Set ClsScr = Nothing
'    DoCmd.Hourglass False
'    Set combo = Nothing
'    Exit Function
'
'OpenFormError:
'    MsgBox err.Description
'    Resume OpenFormExit
'
'    Resume
'End Function

Public Sub CommandBarMakeApp(ByVal oBar As CommandBar, ByVal AppID As Long, ByVal ParentID As Long)
    On Error GoTo ErrorHappened
    Dim db As DAO.Database, rs As DAO.RecordSet
    Dim SQL As String
    Dim oPopup As Office.CommandBarPopup, oButton As Office.CommandBarButton, oCbo As Office.CommandBarComboBox
    Dim oParent As Object
    If oBar Is Nothing Then
        MsgBox "No Command Bar Specified", vbCritical, "CnlyCommandBars.CommandBarMakeApp"
        GoTo ExitNow
    End If

    Set db = CurrentDb

    SQL = "SELECT CnlyAppsMenus.* "
    SQL = SQL & "From CnlyAppsMenus "
    If ParentID = 0 Then
        SQL = SQL & "Where ParentID is null "
        Set oParent = oBar
    Else
        SQL = SQL & "Where ParentID = " & ParentID & " "
        Set oParent = oBar.FindControl(Tag:=ParentID, recursive:=True)
    End If
    SQL = SQL & "And AppID = " & AppID & " "
    SQL = SQL & "ORDER BY CnlyAppsMenus.Ordinal, CnlyAppsMenus.Caption;"

    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot)

    Do While Not rs.EOF And Not rs.BOF
        CommandBarMakeAppItem oParent, rs
        'Debug.Print RS.Fields("Caption")
        CommandBarMakeApp oBar, AppID, rs.Fields("MenuItemID")
        rs.MoveNext
    Loop

    rs.Close

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Set rs = Nothing
    Exit Sub


ErrorHappened:
    MsgBox Err.Description, vbCritical, "CnlyCommandBars.CommandBarMakeApp"
    Resume ExitNow
    Resume

End Sub
Public Sub CommandBarMakeAppItem(ByRef oParent As Object, ByRef rs As DAO.RecordSet)
    On Error GoTo ErrorHappened
    Dim oPopup As Office.CommandBarPopup, oButton As Office.CommandBarButton, oCbo As Office.CommandBarComboBox


    If oParent Is Nothing Then
        MsgBox "No Parent Object Specified", vbCritical, "CnlyCommandBars.CommandBarMakeAppItem"
        GoTo ExitNow
    End If

    Select Case rs.Fields("ControlStyleID")
    Case msoControlPopup
        Set oPopup = oParent.Controls.Add(msoControlPopup, , , , False)
        With oPopup
            .Tag = rs.Fields("MenuItemID")
            .BeginGroup = True
            .Caption = rs.Fields("Caption")
            .ToolTipText = "" & rs.Fields("TooltipText")
            .DescriptionText = "" & rs.Fields("TooltipText")
            .visible = True
            .HelpContextId = "" & rs.Fields("HelpContextId")
            .helpFile = "" & rs.Fields("HelpFile")
        End With
    Case msoControlDropdown
        Set oCbo = oParent.Controls.Add(msoControlDropdown, , , , False)
        With oCbo
            .Tag = rs.Fields("MenuItemID")
            .Caption = rs.Fields("Caption")
            .ToolTipText = "" & rs.Fields("TooltipText")
            .DescriptionText = "" & rs.Fields("TooltipText")
            .visible = True
            .style = msoComboNormal
            .HelpContextId = "" & rs.Fields("HelpContextId")
            .helpFile = "" & rs.Fields("HelpFile")
            .Width = Nz(rs.Fields("Width"), 150)
            .DropDownLines = Nz(rs.Fields("CboDropDownLines"), 8)
            .DropDownWidth = Nz(rs.Fields("CboDropDownWidth"), 300)
            .ListIndex = Nz(rs.Fields("CboListIndex"), 0)
            FillDropDown oCbo, "" & rs.Fields("CboFillSQL"), Nz(rs.Fields("CboFillFieldOrdinal"), 0)
        End With
    Case msoControlButton
        Set oButton = oParent.Controls.Add(msoControlButton, , , , False)
        With oButton
            .Tag = rs.Fields("MenuItemID")
            .Caption = rs.Fields("Caption")
            .FaceId = rs.Fields("FaceId")
            .OnAction = "" & rs.Fields("OnAction")
            .ToolTipText = "" & rs.Fields("TooltipText")
            .style = rs.Fields("ButtonStyleID")
            .Width = rs.Fields("Width")
            .DescriptionText = "" & rs.Fields("TooltipText")
            .visible = True
            .HelpContextId = "" & rs.Fields("HelpContextId")
            .helpFile = "" & rs.Fields("HelpFile")
        End With

    Case Else
        MsgBox "What are you trying to pull?"
    End Select


ExitNow:
    On Error Resume Next
    Set oPopup = Nothing
    Set oButton = Nothing
    Set oCbo = Nothing
    Exit Sub


ErrorHappened:
    MsgBox Err.Description, vbCritical, "CnlyCommandBars.CommandBarMakeAppItem"
    Resume ExitNow
    Resume

End Sub
Private Function GetCommandBar(Name As String) As CommandBar
    On Error GoTo ErrorHappened
    Dim oBar As CommandBar

    For Each oBar In Application.CommandBars
        If UCase(oBar.Name) = UCase(Name) Then
            Set GetCommandBar = oBar
            GoTo ExitNow
        End If
    Next oBar


ExitNow:
    On Error Resume Next
    Set oBar = Nothing
    Exit Function

ErrorHappened:
    MsgBox Err.Description, vbCritical, "CnlyCommandBars.GetCommandBar"
    Resume ExitNow
    Resume
End Function

'Public Function OpenScreen(ScreenName As String) As Long
'    Dim ClsScr As ClsScreenData
'    Set ClsScr = New ClsScreenData
'    With ClsScr
'        .CreateScreen ScreenName
'        If .NewScreen = True Then
'            .GetConfig
'        End If
'        OpenScreen = .ScreenForm.FormID
'    End With
'
'    Set ClsScr = Nothing
'End Function
Public Function CCACommandBarOpenFormGeneric(FormName As String, Optional OpenArgs As String)
    On Error GoTo OpenFormError

    If "" & OpenArgs <> "" Then
        DoCmd.OpenForm FormName, , , , , , OpenArgs
    Else
        Debug.Print Now()
        DoCmd.OpenForm FormName
        Debug.Print Now()

    End If



OpenFormExit:
    Exit Function

OpenFormError:
    MsgBox Err.Description
    Resume OpenFormExit
End Function
Public Sub CCACommandBarDelete()
    On Error GoTo ErrorHappened
    Dim MyBar As CommandBar

    Set MyBar = Application.CommandBars.Item(ToolbarName)
    MyBar.Delete

ExitNow:
    On Error Resume Next
    Set MyBar = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "CommandBarDelete"
    Resume ExitNow
End Sub



Public Function CCACommandBarMake()
    On Error GoTo CCACommandBarMakeError
    Dim ScrDb As DAO.Database
    Dim ScrRst As DAO.RecordSet
    Dim MyBar As Office.CommandBar
    Dim CmndB As Office.CommandBarButton
    Dim combo As Office.CommandBarComboBox
    Dim ComboFrm As Office.CommandBarPopup
    Dim oPopupButton As Office.CommandBarButton
    Dim SQL As String

    'Clear The Old one Out
    If Not GetCommandBar(ToolbarName) Is Nothing Then
        CCACommandBarDelete
    End If

    Set ScrDb = CurrentDb
    'Set ScrRst = ScrDb.OpenRecordset("Select Screenname from CnlyScreens where included = true Order By Screenname", dbOpenSnapshot, dbReadOnly)
    Set MyBar = CommandBars.Add(Name:=ToolbarName, position:=msoBarTop, Temporary:=False, MenuBar:=False)


    With MyBar
        .visible = True
    End With



    '' **** DECIPHER SECTION **** '
    'CommandBarMakeApp MyBar, 1, 0
    '
    ' **** RETREIVER SECTION **** '
    'If Nz(DLookup("Value", "CnlyScreenOptions", "OptionName = 'UseRetreiver'"), True) = True Then
    '    CommandBarMakeApp MyBar, 3, 19
    'End If
    ''
    '' **** DUP TOOL SECTION **** '
    'If Nz(DLookup("Value", "CnlyScreenOptions", "OptionName = 'UseDupTool'"), True) = True Then
    '    CommandBarMakeApp MyBar, 4, 26
    'End If
    '
    '' **** CLAIMSPLUS SECTION **** '
    'If Nz(DLookup("Value", "CnlyScreenOptions", "OptionName = 'UseClaimsPlus'"), True) = True Then
    '    CommandBarMakeApp MyBar, 2, 0
    'End If
    '
    '' **** StatementTool SECTION **** '
    'If Nz(DLookup("Value", "CnlyScreenOptions", "OptionName = 'UseStatementTool'"), True) = True Then
    '    CommandBarMakeApp MyBar, 5, 0
    '    CreateStmtMenus
    'End If


    '** QUICK LAUNCH GROUP **
    Set ScrRst = ScrDb.OpenRecordSet("Select ListName,FormName, IconID, OpenArgs from CnlyScreensQuickLaunch Order By SortOrder, ListName", dbOpenSnapshot, dbReadOnly)
    If Not (ScrRst.EOF And ScrRst.BOF) Then
        Set ComboFrm = MyBar.Controls.Add(msoControlPopup, , , , False)
        With ComboFrm
            .Caption = "Quick Launch"
            .BeginGroup = True
            .ToolTipText = "Select the Form To Launch"


            Do Until ScrRst.EOF
                Set oPopupButton = .CommandBar.Controls.Add(Office.MsoControlType.msoControlButton)
                ' Change the face ID and caption for the button.
                oPopupButton.FaceId = ScrRst!IconID
                oPopupButton.Caption = ScrRst!ListName
                If "" & ScrRst!OpenArgs <> "" Then
                    oPopupButton.OnAction = "=CCACommandBarOpenFormGeneric(" & Chr(34) & ScrRst!FormName & Chr(34) & "," & Chr(34) & ScrRst!OpenArgs & Chr(34) & ")"
                Else
                    oPopupButton.OnAction = "=CCACommandBarOpenFormGeneric('" & ScrRst!FormName & "')"
                End If
                ScrRst.MoveNext
            Loop
            '.HelpFile = CCAHelpFile
            .helpFile = Identity.CCAHelp
            .HelpContextId = 3
        End With
    End If

    CommandBarPopupVisibleDataCenter


CCACommandBarMakeExit:
    On Error Resume Next
    ScrRst.Close
    Set ScrRst = Nothing
    Set ScrDb = Nothing
    Set MyBar = Nothing
    Set CmndB = Nothing
    Set combo = Nothing
    Exit Function

CCACommandBarMakeError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Making Screens Command Bar!", vbCritical, "FATAL ERROR"
    Resume CCACommandBarMakeExit
    Resume

End Function

Public Sub CommandBarPopupVisibleDataCenter()
    Dim visible As Boolean

    'sComputer = Identity.Computer
    'Select Case True
    'Case Left(sComputer, 5) = "DCWS-"
    '    Visible = True
    'Case Left(sComputer, 6) = "TS-DC-"
    '    Visible = True
    'Case Left(sComputer, 3) = "DC-"
    '    Visible = True
    'Case Else
    '    Visible = False
    'End Select

    '  visible = IsDcUser

    'Toggle the Data Center Menu Visibility
    ' CommandBarPopupVisible ToolbarPopupDC, visible
    'Toggle The Hide and Unhide Database Windows Visibility
    'Application.CommandBars.Item("Window").Controls.Item("Hide").visible = visible
    ' Application.CommandBars.Item("Window").Controls.Item(Application.CommandBars.Item("Window").Controls.Item("Hide").Index + 1).visible = visible
End Sub

Private Sub CommandBarPopupVisible(Name As String, Value As Boolean)
    On Error Resume Next
    Application.CommandBars.Item(ToolbarName).Controls.Item(Name).visible = Value
End Sub
Public Function CCACommandBarMakeIcons()
    On Error GoTo CCACommandBarMakeError
    Dim MyBar As CommandBar, CmndB As CommandBarButton, combo As CommandBarComboBox
    Dim ComboFrm As Office.CommandBarPopup, oPopupButton As CommandBarButton
    Dim Sb2 As Office.CommandBarPopup



    For Each MyBar In Application.CommandBars
        If MyBar.Name = "Icons For Testing" Then
            MyBar.Delete
            Exit For
        End If
    Next MyBar

    Set MyBar = Nothing

    Set MyBar = CommandBars.Add(Name:="Icons For Testing", position:=msoBarTop, Temporary:=True, MenuBar:=False)


    MyBar.visible = True

    Dim X As Long, Sb As Office.CommandBarPopup, Y As Long
    Set ComboFrm = MyBar.Controls.Add(msoControlPopup, , , , False)
    With ComboFrm
        .Caption = "Image Chooser"
        .ToolTipText = "Select the Form To Launch"
        Set Sb = ComboFrm.Controls.Add(msoControlPopup, , , , False)
        Sb.Caption = "0001 - 0100"
        Set Sb2 = Sb.Controls.Add(msoControlPopup, , , , False)
        Sb2.Caption = Right("0000" & (X + 1), 4) & " - " & Right("0000" & (X + 18), 4)
        For X = 1 To 4399

            Set oPopupButton = Sb2.CommandBar.Controls.Add(Office.MsoControlType.msoControlButton)
            ' Change the face ID and caption for the button.
            oPopupButton.FaceId = X
            oPopupButton.Caption = X

            If X Mod 100 = 0 Then
                Set Sb = ComboFrm.Controls.Add(msoControlPopup, , , , False)
                Sb.Caption = Right("0000" & (X + 1), 4) & " - " & Right("0000" & (X + 100), 4)
                DoEvents
            End If
            If X Mod 20 = 0 Then
                Set Sb2 = Sb.Controls.Add(msoControlPopup, , , , False)
                Sb2.Caption = Right("0000" & (X + 1), 4) & " - " & Right("0000" & (X + 20), 4)
            End If
        Next X
        '.HelpFile = CCAHelpFile
        .helpFile = Identity.CCAHelp
        .HelpContextId = 3
    End With


CCACommandBarMakeExit:
    On Error Resume Next
    Set MyBar = Nothing
    Set CmndB = Nothing
    Set combo = Nothing
    Exit Function

CCACommandBarMakeError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Making Screens Command Bar!", vbCritical, "FATAL ERROR"
    Resume CCACommandBarMakeExit
    Resume

End Function
'
'Public Function CreateDesktopShortcut()
'On Error GoTo ErrorHappened
'Dim fso ' as Scripting.FileSystemObject
'Dim WshShell ' as WScript.Shell
'Dim StFolder As String, StIcon As String, StFileName As String, StShortcutName As String
'Dim ObjShortCut ' as WScript.WshShortcut
'
'Set fso = CreateObject("Scripting.FileSystemObject")
'Set WshShell = CreateObject("WScript.Shell")
'StFolder = WshShell.SpecialFolders("Desktop")
'StFileName = CurrentDb.Name
'StIcon = GetIconName()
'StShortcutName = StFolder & "\" & Identity.ClientName & " - Decipher.lnk"
'If fso.FileExists(StShortcutName) = True Then
'    fso.DeleteFile StShortcutName
'End If
'
'Set ObjShortCut = WshShell.CreateShortcut(StShortcutName)
'With ObjShortCut
'    .TargetPath = StFileName
'    .IconLocation = StIcon
'    .WorkingDirectory = fso.GetFile(StFileName).ParentFolder.Path
'    .Description = "Connolly Decipher"
'    .WindowStyle = 3 'MAXIMIZED
'    .save
'End With
'ExitNow:
'    On Error Resume Next
'    Set ObjShortCut = Nothing
'    Set WshShell = Nothing
'    Set fso = Nothing
'    Exit Function
'ErrorHappened:
'    MsgBox err.Description, vbInformation, "Error Creating Desktop Shortcut"
'    Resume ExitNow
'    Resume
'
'End Function

Public Function DatabaseVisible(Value As String) As Boolean
    On Error Resume Next

    If Value = "True" Then
        DoCmd.RunCommand acCmdWindowUnhide
    Else
        DoCmd.RunCommand acCmdWindowHide
    End If

End Function

Private Function FillDropDown(oCbo As Office.CommandBarComboBox, SQL As String, FieldOrdinal As Long)
    On Error GoTo ErrorHappened
    Dim db As DAO.Database, rs As DAO.RecordSet

    'Set db = CurrentDb
    'Set Rs = db.OpenRecordSet(Sql, dbOpenSnapshot)
    '
    'Do While Not Rs.EOF And Not Rs.BOF
    '    oCbo.addItem Rs.Fields(FieldOrdinal)
    '    Rs.MoveNext
    'Loop
    'Rs.Close
ExitNow:
    On Error Resume Next
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrorHappened:
    MsgBox Err.Description, vbInformation, "Error in FillDropDown"
    Resume ExitNow
    Resume
End Function
'
'Public Function SaveAllScreens()
'On Error GoTo ErrorHappened
'Dim i As Integer
'
'For i = 1 To 20
'    If Not Scr(i) Is Nothing Then
'        Scr(i).CmdScreenSave_Click
'    End If
'Next i
'
'AllDone:
'Exit Function
'
'ErrorHappened:
'MsgBox err.Description
'Resume AllDone
'
'End Function
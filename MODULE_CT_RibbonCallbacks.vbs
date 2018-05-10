Option Compare Database
Option Explicit

Public decipherRibbon As IRibbonUI                  'variable to reference the ribbon object
Public ribbonIsDCUser As Boolean
Private Const QI As String = """"                   'used for quotations in sql calls
Private Const DecipherPrefix As String = "SCR_"     'prefix for the Decipher Screens
Private Const ClaimsPlusFrPrefix As String = "CP_"  'prefix for the ClaimsPlus Framework

'It manages all the required calls to run Decipher add-ins
Private addInManager As New CT_ClsCnlyAddinSupport

Public Sub LoadRibbonCallback(ribbon As IRibbonUI)
    Set decipherRibbon = ribbon
    ' call the normal dcuser routine and set ribbonIsDcUser
    ribbonIsDCUser = isDcUser
End Sub

Public Sub getImages(imageId As String, ByRef Image)
' Images Call back
  Set Image = getIconFromTable(imageId)

End Sub

' Routines found in the CnlyRibbonStdVisibleCallbacks table used to by the ribbon to determine tab,
' group, and control visibility
Public Sub OpenRibbonCreator(Control As IRibbonControl)
    On Error GoTo OpenRibbonCreatorError:
  
    addInManager.OpenRibbonDesigner
   
   'JL 1/17/2012 - Added telemetry
    Telemetry.RecordOpen "Form", "Ribbon Designer"
OpenRibbonCreatorExit:
     Exit Sub
OpenRibbonCreatorError:
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Description
    Resume OpenRibbonCreatorExit
End Sub

Public Sub BuildRibbonBar(Control As IRibbonControl)
    On Error GoTo BuildRibbonBarError:
   
    addInManager.BuildRibbonBar
   
BuildRibbonBarExit:
     Exit Sub
BuildRibbonBarError:
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Description
    Resume BuildRibbonBarExit
End Sub

Public Sub RibbonAlwaysShow(Control As IRibbonControl, ByRef visible As Variant)
    visible = True
End Sub
' Standard visible callbacks
Public Sub RibbonDCUser(Control As IRibbonControl, ByRef visible As Variant)
    ' this routine returns the value of RibbonisDCUser.  RibbonIsDCUser is set by the ribbon load on callback routine
    visible = ribbonIsDCUser
End Sub

Public Sub RibbonClaimsPlusUser(Control As IRibbonControl, ByRef visible As Variant)
    visible = IsClaimsPlusEnabled
End Sub

Public Sub RibbonDecipherUser(Control As IRibbonControl, ByRef visible As Variant)
    visible = IsDecipherEnabled
End Sub

Public Sub RibbonDecipherHideHelp(Control As IRibbonControl, ByRef visible As Variant)
    visible = IsDecipherHelpShowEnabled
End Sub

Public Sub RibbonDonotShow(Control As IRibbonControl, ByRef visible As Variant)
    visible = False
End Sub

Public Sub RibbonShowExample(Control As IRibbonControl, ByRef visible As Variant)
    visible = False
End Sub

' Standard action callbacks
Public Sub LoadForm(Control As IRibbonControl)
On Error GoTo LoadFormError
    Dim itemForm As String
    Dim itemOpenArgs As String
    Dim position As Integer
    position = InStr(Control.Tag, ";")
    If position > 0 Then
        itemForm = Mid(Control.Tag, 1, position - 1)
        If position + 1 <= Len(Control.Tag) Then
            itemOpenArgs = Mid(Control.Tag, position + 1)
        End If
    Else
        itemForm = Control.Tag
    End If
    
    If "" & itemOpenArgs <> "" Then
        DoCmd.OpenForm itemForm, , , , , , itemOpenArgs
    Else
        
        DoCmd.OpenForm itemForm
    End If

    'JL 1/17/2012 - Added telemetry
    If itemForm = "AIV_AdImageViewer" Then
        Telemetry.RecordOpen "Form", "Add Image Viewer"
    End If
LoadFormExit:
    Exit Sub
    
LoadFormError:
    If itemForm <> "frm_ADMIN_User_Hours" Then  ' we are forcing it to cancel
        MsgBox Err.Description
    End If
    Resume LoadFormExit
End Sub

Public Sub RunReport(Control As IRibbonControl)
'JL 11/15/2012 - Send report to printer
On Error GoTo RunReportError
    Dim itemReport As String
    Dim itemOpenArgs As String
    Dim position As Integer
    position = InStr(Control.Tag, ";")
    If position > 0 Then
        itemReport = Mid(Control.Tag, 1, position - 1)
        If position + 1 <= Len(Control.Tag) Then
            itemOpenArgs = Mid(Control.Tag, position + 1)
        End If
    Else
        itemReport = Control.Tag
    End If
    
    If "" & itemOpenArgs <> "" Then
        DoCmd.OpenReport itemReport, acViewNormal, , , , itemOpenArgs
    Else
        DoCmd.OpenReport itemReport, acViewNormal
    End If

RunReportExit:
    Exit Sub
    
RunReportError:
    MsgBox Err.Description, vbCritical, "Print Report Error"
    Resume RunReportExit
End Sub

Public Sub PreviewReport(Control As IRibbonControl)
'JL 11/15/2012 - New callback to open the report in a preview window
On Error GoTo PreviewReportError
    Dim itemReport As String
    Dim itemOpenArgs As String
    Dim position As Integer
    position = InStr(Control.Tag, ";")
    If position > 0 Then
        itemReport = Mid(Control.Tag, 1, position - 1)
        If position + 1 <= Len(Control.Tag) Then
            itemOpenArgs = Mid(Control.Tag, position + 1)
        End If
    Else
        itemReport = Control.Tag
    End If
    
    If "" & itemOpenArgs <> "" Then
        DoCmd.OpenReport itemReport, acViewPreview, , , , itemOpenArgs
    Else
        DoCmd.OpenReport itemReport, acViewPreview
    End If

PreviewReportExit:
    Exit Sub
    
PreviewReportError:
    MsgBox Err.Description, vbCritical, "Preview Report Error"
    Resume PreviewReportExit
End Sub

Public Sub OpenTable(Control As IRibbonControl)
On Error GoTo OpenTableError
    Dim itemTable As String
    Dim position As Integer
    position = InStr(Control.Tag, ";")
    If position > 0 Then
        itemTable = Mid(Control.Tag, 1, position - 1)
    Else
        itemTable = Control.Tag
    End If
    
    If "" & itemTable <> "" Then
        DoCmd.OpenTable itemTable
    End If
 
OpenTableExit:
    Exit Sub
    
OpenTableError:
    MsgBox Err.Description
    Resume OpenTableExit
End Sub
' Specific implementation of ribbon bar buttons

Public Sub UnhideDatabaseWindow(Control As IRibbonControl, pressed As Boolean)
    DoCmd.SelectObject acTable, , True
    If Not pressed Then
        DoCmd.RunCommand acCmdWindowHide
    End If
End Sub

Public Sub CreateDesktopShortcut(ByVal Control As IRibbonControl)
On Error GoTo ErrorHappened
    Dim fso ' as Scripting.FileSystemObject
    Dim WshShell ' as WScript.Shell
    Dim StFolder As String, StIcon As String, StFileName As String, StShortcutName As String
    Dim ObjShortCut ' as WScript.WshShortcut
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set WshShell = CreateObject("WScript.Shell")
    StFolder = WshShell.SpecialFolders("Desktop")
    StFileName = CurrentDb.Name
    StIcon = GetIconName()
    StShortcutName = StFolder & "\" & Identity.ClientName & " - Decipher.lnk"
    If fso.FileExists(StShortcutName) = True Then
        fso.DeleteFile StShortcutName
    End If
    
    Set ObjShortCut = WshShell.CreateShortcut(StShortcutName)
    With ObjShortCut
        .TargetPath = StFileName
        .IconLocation = StIcon
        .WorkingDirectory = fso.GetFile(StFileName).ParentFolder.Path
        .Description = "Connolly Decipher"
        .WindowStyle = 3 'MAXIMIZED
        .Save
    End With
ExitNow:
        On Error Resume Next
        Set ObjShortCut = Nothing
        Set WshShell = Nothing
        Set fso = Nothing
        Exit Sub

ErrorHappened:
        MsgBox Err.Description, vbInformation, "Error Creating Desktop Shortcut"
        Resume ExitNow
        Resume

End Sub

Public Sub RibbonSetApplicationTitle(ByVal Control As IRibbonControl)
    SetApplicationTitle
End Sub

Public Sub ShowDecipherHelp(ByVal Control As IRibbonControl)
    ShowHelpFile "Decipher2.chm"
    'JL 1/17/2012 - Added telemetry
    Telemetry.RecordOpen "Help", "Decipher Help"
End Sub

Public Sub ShowClaimsPlusHelp(ByVal Control As IRibbonControl)
    ShowHelpFile ("ClaimsPlusQuickStart.chm")
End Sub

Private Sub ShowHelpFile(ByVal FileName As String)
    Dim helpPath As String

    helpPath = GetSystemPath() & "\Help\" & FileName
    If Dir(helpPath) <> "" Then
        ShellExe (helpPath)
    Else
        MsgBox "Cannot locate Help File: " & helpPath
    End If

End Sub

' -- Private
' private functions/sub used the ribbon
Private Function getIconFromTable(strAppName As String) As Picture
    On Error GoTo ErrorRoutine

' routine to return an icon from the CnlyIcon Table
    Dim LSize As Long
    Dim arrBin() As Byte
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim sqlString As String
    
    Set db = CurrentDb
    Set rs = Nothing
    sqlString = "SELECT Image" & _
        " FROM CT_Icons" & _
        " WHERE AppName = " & QI & strAppName & QI
        
    Set rs = db.OpenRecordSet(sqlString, dbReadOnly)
    If Not rs.EOF Or Not rs.BOF Then
        LSize = rs!Image.FieldSize - 1
        ReDim arrBin(LSize)
        arrBin = rs!Image
        Set getIconFromTable = ArrayToPicture(arrBin)
    End If
 
Done:
    Erase arrBin
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function
    
ErrorRoutine:
' need to fix this to return a std picture when not found
    Resume Done
End Function

Private Function IsDecipherEnabled() As Boolean
On Error Resume Next
    Dim Result As Boolean
    Result = False
    Result = DCount("[RibbonPrefix]", "CT_InstalledApps", "RibbonPrefix=" & QI & DecipherPrefix & QI)
    IsDecipherEnabled = Result
End Function

Private Function IsClaimsPlusEnabled() As Boolean
On Error Resume Next
    Dim Result As Boolean
    Result = False
    Result = DCount("[RibbonPrefix]", "CT_InstalledApps", "RibbonPrefix=" & QI & ClaimsPlusFrPrefix & QI)
    IsClaimsPlusEnabled = Result
End Function

Private Function IsDecipherHelpShowEnabled() As Boolean
On Error Resume Next
    Dim Result As Boolean
    Result = False
    Result = DLookup("[Value]", "CT_Options", "OptionName=" & QI & "UseDecipherHelp" & QI)
    IsDecipherHelpShowEnabled = Result
End Function

' Example Actions
' call back for a standard ribbon control
'Public Sub CloseActionExample(control As IRibbonControl, ByRef cancelDefault)
'    MsgBox ("Close Action Callback")
'
'End Sub
'Public Sub ComboBoxActionExample(control As IRibbonControl, strtext As String)
'    MsgBox ("combo item selected = " & strtext)
'End Sub
'Public Sub ToggleButtonActionExample(control As IRibbonControl, pressed As Boolean)
'    If pressed Then
'        MsgBox ("ToggleButton -- pressed")
'    Else
'        MsgBox ("ToggleButton -- not pressed")
'    End If
'End Sub
'Public Sub CheckBoxActionExample(control As IRibbonControl, pressed As Boolean)
'    If pressed Then
'        MsgBox ("CheckBox-- checked")
'    Else
'        MsgBox ("CheckBox-- unchecked")
'    End If
'End Sub
'Public Sub EditBoxActionExample(control As IRibbonControl, strtext As String)
'    MsgBox "Edit Box - value: " & strtext
'End Sub
'
'Public Sub DropDownActionExample(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
'    MsgBox ("drop down item selected id: " & selectedId & " index: " & selectedIndex)
'End Sub
'
'Public Sub DialogLauncherActionExample(ByVal control As IRibbonControl)
'    MsgBox ("DialogLauncherActionExample")
'End Sub
'
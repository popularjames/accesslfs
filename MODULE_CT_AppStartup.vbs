Option Compare Database
Option Explicit

Private genUtils As New CT_ClsGeneralUtilities
Public Function SetAppStartup()
        
    'If the ribbon bar is greater than 60 then it is expanded.
    If Application.CommandBars("Ribbon").Height > 60 Then
        CommandBars.ExecuteMso "MinimizeRibbon"
    End If
    
    SetApplicationTitle
    
    'Create or update the Icon as necessary
    WriteIcon

    SetStartupProperty "AppIcon", dbText, GetIconName
    Application.RefreshTitleBar
            
    ' HC 5/2010 -- removed for 2010
    'CommandBarPopupVisibleDataCenter ' SHOW THE DATA CENTER MENU IF ON A DC MACHINE

    RunAppStartFunctions 'Run list of functions in CT_AppStartupSeq table.
    
    ' JL 3/2011 - load any available add-ins
    LoadAccessAddins
End Function

' DLC 11/2012 - This needs to be accessible from AppManager following a core template upgrade
Public Sub LoadAccessAddins()
    'Manages all the required calls to run Decipher add-ins
    Dim addInManager As New CT_ClsCnlyAddinSupport
    
    addInManager.LoadAddins isDcUser
End Sub

Public Sub ReadIcon(FileName As String)
On Error GoTo ErrorHappened
Dim ClIcon As New CT_ClsIcon
    With ClIcon
        .FileName = FileName
        .LoadFile CnlyAppName
    End With

ExitNow:
    On Error Resume Next
    Set ClIcon = Nothing
    Exit Sub

ErrorHappened:
    MsgBox Err.Description, vbCritical, "Load Icon File"
    Resume ExitNow
    Resume

End Sub


Private Sub WriteIcon()
On Error GoTo ErrorHappened
Dim ClIcon As New CT_ClsIcon
    With ClIcon
        .FileName = GetIconName
        .SaveFile CnlyAppName
    End With

ExitNow:
    On Error Resume Next
    Set ClIcon = Nothing
    Exit Sub

ErrorHappened:
    MsgBox Err.Description, vbCritical, "Save Icon File"
    Resume ExitNow
    Resume

End Sub
Public Function GetIconName() As String
On Error GoTo ErrorHappened

Dim oShell ' WScript.Shell
Dim oFso ' as Scripting.FileSystemObject
Dim FileName As String
    
    Set oShell = CreateObject("WScript.Shell")
    
    
    FileName = oShell.SpecialFolders("AppData") & "\Connolly"
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    If oFso.FolderExists(FileName) = False Then
        Call oFso.CreateFolder(FileName)
    End If
    
    FileName = FileName & "\" & CnlyAppName & ".ico"
    


'FileName = CurrentDb.Name
'FileName = Replace(FileName, ".mdb", ".ico")

    GetIconName = FileName
ExitNow:
    On Error Resume Next
    Set oShell = Nothing
    Set oFso = Nothing
    Exit Function
ErrorHappened:
    MsgBox Err.Description, vbCritical, "CnlyAppStartup.GetIconName"
    Resume ExitNow
    Resume
    
End Function

Private Function IconExists() As Boolean
On Error GoTo ErrorHappened
Dim fso 'Scripting.FileSystemObject

    Set fso = CreateObject("Scripting.FileSystemObject")
    IconExists = fso.FileExists(GetIconName())

ExitNow:
    On Error Resume Next
    Set fso = Nothing
    Exit Function

ErrorHappened:
    MsgBox Err.Description, vbCritical, "Load Icon File"
    Resume ExitNow
    Resume
End Function


Public Function SetApplicationTitle()
On Error GoTo ErrorHappened
Dim tmpStr As String
Dim visible As Boolean
Dim SAppPath As String
Dim LinkedLocation As String
    SAppPath = "     " & CurrentProject.FullName
    If SavedLocationGet() <> "" Then
        LinkedLocation = " - Linked to: " & SavedLocationGet
    Else
        LinkedLocation = ""
    End If
    tmpStr = Identity.ClientName
    tmpStr = IIf(tmpStr = "", CnlyAppName, tmpStr & " - " & CnlyAppName)
    visible = isDcUser
    'DLC 05/20/2010 Do not attempt to set the title if the database is readonly
    If Not genUtils.ApplicationIsReadOnly Then
        If SetStartupProperty("AppTitle", dbText, tmpStr & LinkedLocation & IIf(visible, SAppPath, "")) Then
            Application.RefreshTitleBar
        Else
            MsgBox "ERROR: Could not set Application Title"
        End If
    End If
ExitNow:
    On Error Resume Next
    Exit Function
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Set Application Title"
    Resume ExitNow
End Function



Public Function SetStartupProperty(prpName As String, prpType As Variant, prpValue As Variant) As Integer

Dim db As Database, prp As Property
Const ERROR_PROPNOTFOUND = 3270
Set db = CurrentDb()
' Set the startup property value.
On Error GoTo Err_SetStartupProperty
db.Properties(prpName) = prpValue
SetStartupProperty = True
         
         
Bye_SetStartupProperty:
      Exit Function
Err_SetStartupProperty:
         Select Case Err
         ' If the property does not exist, create it and try again.
         Case ERROR_PROPNOTFOUND
            Set prp = db.CreateProperty(prpName, prpType, prpValue)
            db.Properties.Append prp
            Resume
            Case Else
            SetStartupProperty = False
            Resume Bye_SetStartupProperty
         End Select
End Function
Public Sub RunAppStartFunctions()
On Error GoTo ErrorHappened
Dim db As DAO.Database
Dim rst As DAO.RecordSet
Dim SQL As String
Dim FunctionName As String
    
    SQL = "Select * From  CT_AppStartupSeq"
    SQL = SQL & " Order By Seq, Function"
    Set db = CurrentDb
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot, dbForwardOnly)
    With rst
        If .EOF And .BOF Then
            .Close
            GoTo ExitNow
        End If
        Do Until .EOF
            FunctionName = .Fields("Function")
            Application.Run FunctionName
            FunctionName = ""
            .MoveNext
        Loop
    End With

ExitNow:
    On Error Resume Next
    Set rst = Nothing
    Set db = Nothing
    Exit Sub
ErrorHappened:
    SQL = Err.Description & vbCrLf & vbCrLf
    SQL = SQL & "Startup Function: " & FunctionName & vbCrLf

    MsgBox SQL, vbCritical, "Error Running Startup Function"
    Resume ExitNow
End Sub
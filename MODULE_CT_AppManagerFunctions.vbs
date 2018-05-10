Option Compare Database
Option Explicit

Private AppManagerDbPath As String
Private OldAppManagerDbPath As String
Private AppSourceServer As String
Private AppSourceDB As String
Private Const AppManName As String = "AppSource"
Private addInManager As New CT_ClsCnlyAddinSupport

Private Function GetNetworkPath(ByVal DriveName As String) As String
    
    Dim objNtWork   As Object
    Dim objDrives   As Object
    Dim lngLoop     As Long
    
    Set objNtWork = CreateObject("WScript.Network")
    Set objDrives = objNtWork.enumnetworkdrives
    
    For lngLoop = 0 To objDrives.Count - 1 Step 2
        If UCase(objDrives.Item(lngLoop)) = UCase(DriveName) Then
            GetNetworkPath = objDrives.Item(lngLoop + 1)
            Exit For
        End If
    Next

End Function

Public Sub RibbonOpenAppManager(Control As IRibbonControl)
'Open app manager if user has exclusive access
On Error GoTo ErrorHappened
    If IsACCDE Then
        MsgBox "The App Manager cannot be launched from an ACCDE file", vbInformation + vbOKOnly, "App Manager"
        GoTo ExitNow
    End If
    If isDcUser Then
        If DCount("empCurrentlyLoggedIn", "CT_CurrentlyLoggedIn", "empCurrentlyLoggedIn<>'" & Replace(Identity.UserName, "'", "''") & "'") = 0 Then
            CheckUpdateManagerVersion
            AddAppManagerRef
            
            Telemetry.RecordOpen "Form", "APP_Manager"
        Else
            MsgBox "Multiple users have the database open. You must have it open exclusively to use the App Manager", vbInformation, "Exclusive Access Only"
        End If
    End If
ExitNow:
    Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

Public Sub RibbonOpenAppUploadManager(Control As IRibbonControl)
'Open app upload manager if user has exclusive access
On Error GoTo ErrorHappened
    If IsACCDE Then
        MsgBox "The App Upload Manager cannot be launched from an ACCDE file", vbInformation + vbOKOnly, "App Upload Manager"
        GoTo ExitNow
    End If
    If isDcUser Then
        If DCount("empCurrentlyLoggedIn", "CT_CurrentlyLoggedIn", "empCurrentlyLoggedIn<>'" & Replace(Identity.UserName, "'", "''") & "'") = 0 Then
            CheckUpdateManagerVersion
            AddAppManagerRef
            Run "OpenAppUploadManager"
            Telemetry.RecordOpen "Form", "APP_UploadManager"
        Else
            MsgBox "Multiple users have the database open. You must have it open exclusively to use the App Upload Manager", vbInformation, "Exclusive Access Only"
        End If
    End If
ExitNow:
    Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

Private Sub ReadConfig()
'Get database connection settings from table
On Error GoTo ErrorHappened
    Dim rs As DAO.RecordSet
    Dim db As DAO.Database
    
    If LenB(AppSourceServer) = 0 And LenB(AppSourceDB) = 0 Then
        Set db = CurrentDb
        Set rs = db.OpenRecordSet("SELECT Server, Database FROM CT_AppSource WHERE Active")
        If rs.recordCount > 0 Then
            AppSourceServer = Nz(rs!Server, vbNullString)
            AppSourceDB = Nz(rs!Database, vbNullString)
        End If
    End If
ExitNow:
On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

Private Sub CheckUpdateManagerVersion()
On Error GoTo ErrorHappened

    Dim StConnect As String
    Dim LocConn 'As ADODB.Connection
    Dim LocRst 'As ADODB.Recordset
    Dim LocErr 'As ADODB.Error
    Dim SQL As String
    Dim AppName As String
    Dim LocalVersion As String
    Dim LocalFileName As String
    Dim MasterProductID As Long
    Dim MasterVersionID As Long
    Dim MasterVersion As String
    Dim MasterFileID As String
    Dim MasterFileName As String
    Dim MasterFilePath As String
    Dim FilePath As String
    Dim NetworkPath As String
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim fso ' as Scripting.FileSystemObject
    Dim genUtils As New CT_ClsGeneralUtilities
    
    ReadConfig
    
    Set db = CurrentDb
    SQL = "SELECT LocalVersion,LocalFileName FROM CT_InstalledApps WHERE ProductName='" & AppManName & "'"
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot)
    If Not (rs.EOF Or rs.BOF) Then
        LocalVersion = Nz(rs!LocalVersion, 0)
        LocalFileName = Nz(rs!LocalFileName, "")
    End If
    
    'Set and open Connection and Command Objects
    DoCmd.Hourglass True
    StConnect = LINK_SRC_SQL & "Data Source=" & AppSourceServer & ";" & _
                "Initial Catalog=" & AppSourceDB & ";"
    Set LocConn = CreateObject("ADODB.Connection")
    With LocConn
        .ConnectionTimeout = 10
        .CommandTimeout = 150
        .ConnectionString = StConnect
        .Open
    End With
    
    Set LocRst = CreateObject("ADODB.Recordset")
    SQL = "SELECT ProductID, ProductName, ProductVersionID, VersionNumber, RepositoryPath, SourceFileName, FileID, Maturity, MaturityDesc " & _
            "FROM DecipherRW.vProductVersion " & _
            "WHERE Latest = 'Y' AND ProductName = '" & AppManName & "'"
    
    LocRst.Open SQL, LocConn
    If Not (LocRst.EOF Or LocRst.BOF) Then
        MasterProductID = LocRst!ProductID
        AppName = LocRst!productName
        MasterVersionID = LocRst!ProductVersionID
        MasterVersion = LocRst!VersionNumber
        MasterFileName = LocRst!SourceFileName
        MasterFileID = LocRst!FileID
        MasterFilePath = LocRst!RepositoryPath & "\"
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    FilePath = Identity.CurrentFolder & "\APPS\"
    NetworkPath = GetNetworkPath(left(FilePath, 2))
    If NetworkPath <> "" Then
        FilePath = Replace(FilePath, left(FilePath, 2), NetworkPath)
    End If
    If fso.FolderExists(FilePath) = False Then
        Call fso.CreateFolder(FilePath)
    End If
    AppManagerDbPath = FilePath & MasterFileName
    If "" & LocalFileName <> "" Then
        OldAppManagerDbPath = FilePath & LocalFileName
    End If
    
    'DLC 07/2012
    'Attempting to reinstall the app manager once it has been loaded will cause an error so don't
    If GetReferenceNumber(AppManName) > 0 Then
        'If the App Manager has been loaded and there is a new version available, inform the user.
        If MasterVersion <> LocalVersion Then
            MsgBox "There is a later version of App Manager available. To get the latest version you will need to close this database and reopen it.", vbInformation, "New Version Available"
        End If
    ElseIf MasterVersion <> LocalVersion Then
        'DLC 07/2012
        'If the App Manager has not been loaded and there is a new version available, prepare to install it
        SQL = "DELETE FROM CT_InstalledApps WHERE ProductID = " & MasterProductID
        db.Execute SQL
        SQL = "INSERT INTO CT_InstalledApps(ProductID, ProductVersionID, ProductName, LocalVersion, LocalFileName, DateInstalled, Maturity, MaturityDesc) " & _
            "VALUES(" & MasterProductID & ", " & MasterVersionID & ", '" & AppName & "', '" & MasterVersion & "', '" & MasterFileName & "', #" & Date & "#, " & LocRst!Maturity & ", '" & LocRst!MaturityDesc & "')"
        db.Execute SQL
        SQL = "UPDATE DecipherRW.vProductVersion SET InstallCount = InstallCount + 1 WHERE ProductVersionID = " & LocRst!ProductVersionID
        LocConn.Execute SQL
        OldAppManagerDbPath = AppManagerDbPath
        'Force new copy
        If Not CopyAppSourceFile(MasterFilePath & MasterFileID, True) Then
            MsgBox "There was a problem copying in the latest version of AppSource", vbCritical, "AppSource Update Error"
        End If
    Else
        'DLC 07/2012
        'If the version has not been increased, copy it only if it is missing
        CopyAppSourceFile MasterFilePath & MasterFileID
    End If

ExitNow:
On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    LocRst.Close
    Set LocRst = Nothing
    Set LocErr = Nothing
    Set LocConn = Nothing
    Set fso = Nothing
    DoCmd.Hourglass False
Exit Sub
ErrorHappened:
    MsgBox genUtils.CompleteDBExecuteError(Err.Description), vbCritical, "CheckUpdateManagerVersion"
    Resume ExitNow
End Sub

Public Function CopyAppSourceFile(ByVal SourceFile As String, Optional ByVal ForceCopy As Boolean = False) As Boolean
'Copy AppSource file into local directory to be used as a reference app
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim FileExists As Boolean
    
    'Check for existing
    If Len(Dir(AppManagerDbPath)) > 0 Then
        FileExists = True
    End If

    'Copy in or overwrite
    If Not FileExists Or ForceCopy Then
        FileCopy SourceFile, AppManagerDbPath
    End If
    
    'Confirm copy
    If Len(Dir(AppManagerDbPath)) > 0 Then
        Result = True
    Else
        Result = False
    End If
ExitNow:
On Error Resume Next
    CopyAppSourceFile = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

'Returns True if any of the MS Access references are missing (broken)
Public Function HasBrokenReference() As Boolean
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim vr As Access.Reference
    For Each vr In Access.References
        If vr.IsBroken Then
            Result = True
            Exit For
        End If
    Next
ExitNow:
    Set vr = Nothing
    HasBrokenReference = Result
    Exit Function
ErrorHappened:
    Resume ExitNow
End Function


Public Sub RemoveBrokenReferences()
On Error GoTo ErrorHappened
    Dim vr As Access.Reference
    For Each vr In Access.References
        If vr.IsBroken Then
            If MsgBox("Would you like to remove the broken reference to :" & vbLf + vr.Name, vbQuestion + vbYesNo, "Remove Broken Reference") = vbYes Then
                Access.References.Remove vr
                Call SysCmd(504, 16483)
            End If
        End If
    Next
ExitNow:
    Set vr = Nothing
    Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

Public Sub RemoveAppManagerRef()
On Error GoTo ErrorHappened
    Dim appManagerRef As Integer
    appManagerRef = GetReferenceNumber(AppManName)
    If appManagerRef > 0 Then
        On Error Resume Next
        Access.References.Remove Access.References(appManagerRef)
        On Error GoTo ErrorHappened
        DBEngine.Idle dbRefreshCache + dbForceOSFlush
        DoEvents
        On Error Resume Next
        Kill OldAppManagerDbPath
        On Error GoTo ErrorHappened
        ' Call a hidden SysCmd to automatically compile/save all modules.
        Call SysCmd(504, 16483)
    End If
ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "RemoveAppManagerRef"
    Resume ExitNow
End Sub

Public Sub AddAppManagerRef()
On Error GoTo ErrorHappened
    If GetReferenceNumber(AppManName) = 0 Then
        Access.References.AddFromFile AppManagerDbPath
    End If
    'Call a hidden SysCmd to automatically compile/save all modules.
    Call SysCmd(504, 16483)
ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

Private Function MatchReference(ByRef AccessRef As Access.Reference, ByVal ReferencePath As String) As Boolean
'Check to see if reference path matches current reference
On Error GoTo ErrorHappened
    Dim Result As Boolean
    
    If ReferencePath = AccessRef.FullPath Then
        Result = True
    Else
        Result = False
    End If

ExitNow:
    MatchReference = False
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

Public Function APP_RibbonBarInstall(InstallCmd As String) As Boolean
    APP_RibbonBarInstall = addInManager.RunRibbonUpdate(InstallCmd)
End Function

Public Sub APP_RibbonBarBuild()
    addInManager.BuildRibbonBar
End Sub

Public Sub APP_TelemetryAction(ByVal ActionName As String, ByVal ActionDetailsXml As String, Optional ByVal DecipherAppName As String = vbNullString)
    Telemetry.RecordAction ActionName, ActionDetailsXml, DecipherAppName
End Sub

Public Sub AppEmailNotification(ByVal MailTo As String, ByVal MailSubject As String, ByVal MailBody As String)
'Send email notification to app developer
On Error GoTo ErrorHappened
    Dim oMsg As New CT_ClsEmail
    
    If LenB(MailTo) > 0 And LenB(MailSubject) > 0 And LenB(MailBody) > 0 Then
        With oMsg
            .ErrorHandler = SuppressError
            .RecipientTo = MailTo
            .Subject = MailSubject
            .Body = MailBody
            .Send
        End With
    End If
    
ExitNow:
On Error Resume Next
    Set oMsg = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

'Returns the number of the specified reference, 0 if one was not found
Public Function GetReferenceNumber(ByVal refName As String) As Integer
    Dim i As Integer
    For i = Access.References.Count To 1 Step -1
        If refName = Access.References(i).Name Then
            Exit For
        End If
    Next
    GetReferenceNumber = i
End Function

Public Function IsProductInstalled(ByVal productName As String) As Boolean
    'Check to see if a product is installed
    If DCount("ProductID", "CT_InstalledApps", "ProductName='" & Replace(productName, "'", "''") & "'") > 0 Then
        IsProductInstalled = True
    Else
        IsProductInstalled = False
    End If
End Function

Public Function SortableVersion(ByVal Version As String) As String
'Convert version number into to sortable format (1.2.3001 -> 000100023001)
On Error GoTo ErrorHappened
    Dim i As Integer
    Dim Result As String
    Dim VersionArray() As String
    VersionArray() = Split(Version, ".")
    
    For i = 0 To UBound(VersionArray)
        Result = Result & Right("0000" & VersionArray(i), 4)
    Next i

ExitNow:
On Error Resume Next
    SortableVersion = Result
Exit Function
ErrorHappened:

    Resume ExitNow
End Function

Public Function IsACCDE()
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Result = (CurrentDb.Properties("MDE") = "T")
ExitNow:
    IsACCDE = Result
    Exit Function
ErrorHappened:
    Resume ExitNow
End Function
Option Compare Database
Option Explicit



Private Declare Function apiFindExecutable Lib "shell32.dll" _
                  Alias "FindExecutableA" (ByVal lpFile As String, _
                                           ByVal lpDirectory As String, _
                                           ByVal lpResult As String) As Long

Private Const ClassName As String = "mod_WinZip_Support"

Private Const cs_ZIP_FILE_SAMPLE_PATH As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\sample_file_for_finding_winzip_exe.zip"
Private Const cs_WINZIP_CMDLINE_APP_NAME As String = "WZZIP.EXE"
Private Const cs_WINZIP_UNZIP_CMDLINE_APP_NAME As String = "WZUNZIP.EXE"



Public Property Get WINZIP_PATH() As String
    WINZIP_PATH = getWinZipPath
End Property

Public Property Get WINZIP_CMD_PATH() As String
    WINZIP_CMD_PATH = getWinZipPath(, True)
End Property


Public Property Get WINZIP_UNZ_PATH() As String
    WINZIP_UNZ_PATH = getWinZipPath(, , True)
End Property



Public Function getWinZipPath(Optional sZipFilePath As String, Optional bReturnCmdLineZip As Boolean, Optional bReturnCmdLineUnZip As Boolean, Optional bCmdLineSafe As Boolean = True) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim SFileName As String
Dim sFolderName As String
Dim sExtension As String
Dim sReturn As String
Dim sWinZipGui As String
Dim sWinZipDir As String


    strProcName = ClassName & ".getWinZipPath"
    
    If sZipFilePath = "" Then sZipFilePath = cs_ZIP_FILE_SAMPLE_PATH
    
    Call PathInfoFromPath(sZipFilePath, SFileName, sFolderName, sExtension)
    
    sWinZipGui = Find_Exe_Name(SFileName & "." & sExtension, sFolderName)
    sWinZipGui = TrueTrim(sWinZipGui)
    
    Call PathInfoFromPath(sWinZipGui, , sWinZipDir)
    If bReturnCmdLineZip = True Then
        getWinZipPath = QualifyFldrPath(sWinZipDir) & cs_WINZIP_CMDLINE_APP_NAME
        
    ElseIf bReturnCmdLineUnZip = True Then
        getWinZipPath = QualifyFldrPath(sWinZipDir) & cs_WINZIP_UNZIP_CMDLINE_APP_NAME
    Else
        getWinZipPath = sWinZipGui
    End If

'    If bCmdLineSafe Or bReturnCmdLineUnZip = True Then
    If bCmdLineSafe = True Then
        If InStr(1, getWinZipPath, " ", vbTextCompare) > 1 Then
            getWinZipPath = """" & getWinZipPath & """"
        End If
    End If
    

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

'' The below are duplicated in Claims Admin but put here in Private context for easy transport to other tools
''
Private Function Find_Exe_Name(prmFile As String, prmDir As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim lReturn_Code As Long
Dim sReturn_Value As String
    
    strProcName = ClassName & ".Find_Exe_Name"
    
    sReturn_Value = Space(260)
    lReturn_Code = apiFindExecutable(prmFile, prmDir, sReturn_Value)
    
    If lReturn_Code > 32 Then
        Find_Exe_Name = sReturn_Value
    Else
        Find_Exe_Name = "Error: File Not Found"
    End If
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Find_Exe_Name = "Error: Could not determine"
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Given a path (UNC or mapped) will break the details up into sections
''' Note, strFileName will NOT have the . & Extension
'''

Private Function PathInfoFromPath(ByVal strFullPath As String, Optional ByRef strFileName As String, _
    Optional ByRef strParentFolder As String, Optional ByRef strFileExtension As String) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Dim oFile As Scripting.file
Dim oRegExp As RegExp
Dim oMatches As MatchCollection
Dim oMatch As Match
    
    strProcName = ClassName & ".PathInforFromPath"
    
    strFullPath = TrimNull(strFullPath)
    Set oFso = New Scripting.FileSystemObject
    If oFso.FileExists(strFullPath) = True Then
        Set oFile = oFso.GetFile(strFullPath)
    
        strParentFolder = QualifyFldrPath(oFile.ParentFolder)
        strFileExtension = oFso.GetExtensionName(strFullPath)
    
        strFileName = Replace(oFile.Name, "." & strFileExtension, "", 1, 1, vbTextCompare)
    Else
        ' File doesn't exist yet, so let's parse the string ourselves!
        Set oRegExp = New RegExp
        With oRegExp
            .Global = False
            .IgnoreCase = True
            .Pattern = "^(.*?\\*)([^\\]+)\\*$"
        End With
        
        Set oMatches = oRegExp.Execute(strFullPath)
        If oMatches.Count > 0 Then
            Set oMatch = oMatches.Item(0)
            strParentFolder = QualifyFldrPath(oMatch.SubMatches(0))
            strFileName = oMatch.SubMatches(1)
            
            oRegExp.Pattern = "^.+?(\.[^\\]+)$"
            strFileExtension = oRegExp.Replace(strFullPath, "$1")
            ' to match, we get rid of the period if found:
            strFileExtension = Replace(strFileExtension, ".", "")
            If strFileExtension = strFullPath Then strFileExtension = ""

        End If
        
    End If

    PathInfoFromPath = Len(strParentFolder)

Block_Exit:
    Set oFile = Nothing
    Set oFso = Nothing
    Set oRegExp = Nothing
    Set oMatches = Nothing
    Set oMatch = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    PathInfoFromPath = 0
End Function


Public Function TrueTrim(ByVal sIn As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oRegEx As RegExp

    strProcName = ClassName & ".TrueTrim"
    
    Set oRegEx = New RegExp
    With oRegEx
        .IgnoreCase = True
        .Global = False
        .MultiLine = False
        .Pattern = "^([ \t\s\r\n]+)"
    End With
    
    
    sIn = oRegEx.Replace(sIn, "")
    oRegEx.Pattern = "([ \t\s\r\n]+)$"
    sIn = oRegEx.Replace(sIn, "")
    
    
    sIn = Replace(sIn, Chr(0), "")
    
    
Block_Exit:
    TrueTrim = sIn
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Insures the path passed ends with a \
Private Function QualifyFldrPath(sPath As String) As String
    'add trailing slash if required
    If Right$(sPath, 1) <> "\" Then
        QualifyFldrPath = sPath & "\"
    Else
        QualifyFldrPath = sPath
    End If
End Function
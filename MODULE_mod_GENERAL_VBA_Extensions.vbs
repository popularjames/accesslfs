Option Compare Database
Option Explicit



'' Last Modified: 09/12/2012
''
'' ############################################################
'' ############################################################
'' ############################################################
''  DESCRIPTION:
''  =====================================
''   Some basic functions I use all the time and don't really
''   fit anywhere else.. These are basically things I wish
''   were included in VBA
''
''  TODO:
''  =====================================
''  - Lots, clean up, better commenting, etc..
''
''  HISTORY:
''  =====================================
''  - 09/12/2012  - KD: Added FinishMethod and StartMethod
''  - 04/24/2012  - KD: Added Create_Zip_File (which requires Microsoft.Shell controls and automation)
''  - 04/16/2012  - KD:  Added CurrentDBDir() and KeyExistsInCollection
''  - 03/12/2012  - KD: Added IsTable, removed redundant Computername and CurrentNetworkUser functions
''  - 02/15/2012  - ReCreated for Connolly
''
'' AUTHOR
''  =====================================
'' Kevin Dearing
''
''
''   Note: the stuff for the API calls (prompting for file primarilly)
''       is public because the old sub needs access to it and I didn't want
''       to leave it in the original module (modImport)
'' ############################################################
'' ############################################################
'' ############################################################
'' ############################################################


Private Const ClassName As String = "mod_GENERAL_VBA_Extensions"

Public Enum SortDictBy
    ByKey = 1
    ByItem = 2
End Enum


Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Long
Private Declare Function GetProfileStringA Lib "kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long



        ''' ##############################################################################
        '''
        ''' END APIs necessary for this module...
        '''
        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################
        ''' I normally separate things but I'm hesitant to add to where I think this should go (mod_GENERAL_API)
Public Sub LogMessage(strProcName As String, Optional strMessageType As String = "CODE TRAIL", _
    Optional strMessage As String = "", Optional strAdditionalInfo As String, Optional blnMsgBoxAlso As Boolean = False, _
    Optional sConceptId As String, Optional sCnlyClaimNum As String)
Dim lMsgFile As Long
Dim cUserName As String
Dim sLogPath As String
Dim bVerbose As Boolean

Const strDefaultErrorMsg As String = "" '    vbCrLf & vbCrLf & "If you see this message more than once and you don't expect to see it, please take a screen grab (ALT + Print Screen) and send it to support for the quickest turn around!"

    sLogPath = GetLogPath(bVerbose)
    
    cUserName = Identity.UserName()

    Debug.Print Time() & ">" & strProcName & "," & strMessageType & "," & strMessage
    
    If strAdditionalInfo <> "" Then
        strMessage = strMessage & "," & strAdditionalInfo
    End If
        
    If blnMsgBoxAlso = True Then
        Select Case UCase(strMessageType)
        Case "USER MESSAGE", "USR MESSAGE", "USR MSG", "USER MSG"
            MsgBox "NOTE: " & strMessage & strDefaultErrorMsg, vbOKOnly, strMessageType
        Case Else
            MsgBox UCase(strMessageType) & ": " & strMessage & strDefaultErrorMsg, vbOKOnly, strMessageType
        End Select
    End If
        
    
    If bVerbose = False And strMessageType = "CODE TRAIL" Then
        Exit Sub    ' don't log this to the file..
    End If

        ' The stored proc has the whole bVerbose logic in it so we can turn it on just by changing the value
    Call LogMessageDb(strProcName, strMessageType, strMessage, strAdditionalInfo, sConceptId, sCnlyClaimNum)
    
    
' 20130426 KD:
'   File is no longer relevant as it's stored on the C drive in the users profile space..
'   May want to save it somewhere else but for now the database is working well
'    lMsgFile = FreeFile
'
'    On Error Resume Next    ' in case someone doesn't have permissions
'
'    Open sLogPath For Append Access Write Lock Write As #lMsgFile
'    Print #lMsgFile, CStr(Now) & "," & strProcName & "," & strMessageType & "," & strMessage & "," & cUserName
'    Close #lMsgFile
'
'    On Error GoTo 0
    
End Sub

  
        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################

Public Sub LogMessageDb(strProcName As String, Optional strMessageType As String = "CODE TRAIL", _
    Optional strMessage As String = "", Optional strAdditionalInfo As String, Optional sConceptId As String, _
    Optional sCnlyClaimNum As String)
On Error GoTo Block_Err
Dim oAdo As clsADO

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CLAIM_ADMIN_LogMessage"
        .Parameters.Refresh
        .Parameters("@pProcName") = strProcName
        .Parameters("@pMessageType") = strMessageType
        .Parameters("@pMessage") = strMessage
        .Parameters("@pAdtnlDetails") = strAdditionalInfo
        If sConceptId <> "" Then
            .Parameters("@pConceptId") = sConceptId
        End If
        If sCnlyClaimNum <> "" Then
            .Parameters("@pCnlyClaimNum") = sCnlyClaimNum
        End If
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            ' uh, then what? nothing..
        End If
    End With
    
Block_Err:

End Sub

        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################

Public Function GetLogPath(Optional bVerboseSetting As Boolean) As String
Static oSettings As clsSettings
Static dRefreshTimer As Date
Dim sLogPath As String

    If oSettings Is Nothing Then Set oSettings = New clsSettings
    
    If DateDiff("n", dRefreshTimer, Now()) > 5 Then
        Set oSettings = New clsSettings
        dRefreshTimer = Now
    End If
    
    sLogPath = oSettings.GetSetting("LOG_FILE_PATH")
    If oSettings.GetSetting("VERBOSE_LOGGING") = "" Then
        oSettings.SetSetting "VERBOSE_LOGGING", "False"
    End If
    
    bVerboseSetting = IIf(oSettings.GetSetting("VERBOSE_LOGGING") = "", False, CBool(Nz(oSettings.GetSetting("VERBOSE_LOGGING"), "False")))
    
    If sLogPath = "" Then
        sLogPath = Replace(CurrentDb.Name, ".mdb", "", 1, 1, vbTextCompare)
        sLogPath = Replace(sLogPath, ".adodb", "", 1, 1, vbTextCompare)
        sLogPath = Replace(sLogPath, ".mde", "", 1, 1, vbTextCompare)
    End If
    sLogPath = sLogPath & "_LOG.txt"
    
    If FolderExist(ParentFolderPath(sLogPath)) = False Then CreateFolders (sLogPath)
    
    GetLogPath = sLogPath
End Function



        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################


Public Sub ReportError(oErr As ErrObject, strProcName As String, Optional strMessageType As String = "ERROR", _
        Optional strAddtnlMessage As String = "", Optional sConceptId As String, Optional sCnlyClaimNum As String)
Dim sMsg As String

    sMsg = oErr.Number & " : " & oErr.Description & IIf(strAddtnlMessage <> "", ":", "") & strAddtnlMessage & _
        " in " & strProcName
    
    LogMessage strProcName, strMessageType, sMsg, strAddtnlMessage, , sConceptId, sCnlyClaimNum

End Sub


        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################

Public Function ArrayContains(sFindString, varyList As Variant) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim iIndex As Integer

    strProcName = ClassName & ".ArrayContains"
    If IsArray(varyList) = False Then
        ArrayContains = False
        GoTo Block_Exit
    End If
    
    For iIndex = 0 To UBound(varyList)
        If sFindString = varyList(iIndex) Then
            ArrayContains = True
            GoTo Block_Exit
        End If
    Next
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    ArrayContains = False
    GoTo Block_Exit
End Function

            '' ##############################################################################
            '' ##############################################################################
            '' ##############################################################################



Public Function GetPrevFridayDate(ByVal dtStartDate As Date) As Date
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".GetPrevFridayDate"
    
    dtStartDate = DateAdd("d", -1, dtStartDate)
    
    While DatePart("w", dtStartDate) <> vbFriday
        dtStartDate = DateAdd("d", -1, dtStartDate)
    Wend
    
    GetPrevFridayDate = dtStartDate

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GetPrevFridayDate = CDate("1/1/1900")
    GoTo Block_Exit
End Function


            '' ##############################################################################
            '' ##############################################################################
            '' ##############################################################################

Public Function QuoteMeta(strInString As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oRegExp As RegExp

    strProcName = ClassName & ".QuoteMeta"
    
    Set oRegExp = New RegExp
    oRegExp.Global = True
    oRegExp.IgnoreCase = True
    
    oRegExp.Pattern = "([\<\>\\\|\/\{\}\[\]\.\*\?\+\!\@\#\$\^\&\(\)\""])"
    
    QuoteMeta = oRegExp.Replace(strInString, "\$1")
    
Block_Exit:
    Set oRegExp = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Resume Block_Exit
End Function


            '' ##############################################################################
            '' ##############################################################################
            '' ##############################################################################

Public Function RemoveMeta(strInString As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oRegExp As RegExp

    strProcName = ClassName & ".RemoveMeta"
    
    Set oRegExp = New RegExp
    oRegExp.Global = True
    oRegExp.IgnoreCase = True
    
    oRegExp.Pattern = "([\<\>\\\|\/\{\}\[\]\.\*\?\+\!\@\#\$\^\&\(\)\""\[\]\(\)])"
    
    RemoveMeta = oRegExp.Replace(strInString, "")
    
Block_Exit:
    Set oRegExp = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Resume Block_Exit
End Function

            '' ##############################################################################
            '' ##############################################################################
            '' ##############################################################################

Public Sub SleepEvents(lSeconds As Long)
Dim dtStart As Date

    If lSeconds > 999 Then
        lSeconds = lSeconds / 1000  ' someone must've thought it was Milliseconds
    End If

    dtStart = Now()
    While DateDiff("s", dtStart, Now()) <= lSeconds
        DoEvents
        DoEvents
        DoEvents
    Wend
    
End Sub

            '' ##############################################################################
            '' ##############################################################################
            '' ##############################################################################




Public Function ProcessTookHowLong(dtStartDate As Date, Optional dtEndDate As Date) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim cMinutes As String
Dim cSeconds As String
Dim lngSeconds As Long
Dim dtDummyDate As Date

    strProcName = ClassName & ".ProcessTookHowLong"
    
    If IsMissing(dtEndDate) = True Or dtEndDate = dtDummyDate Then
        dtEndDate = Now()
    End If
    
    lngSeconds = DateDiff("s", dtStartDate, dtEndDate)
    cMinutes = CStr(lngSeconds / 60)
    If CDbl(cMinutes) > 1 Then
        If InStr(1, cMinutes, ".") > 0 Then
            cMinutes = left(cMinutes, InStr(1, cMinutes, ".") - 1)
        End If
    Else
        cMinutes = "0"
    End If
    cSeconds = (lngSeconds - (CInt(cMinutes) * 60))
    If Len(cSeconds) < 2 Then cSeconds = "0" & cSeconds
    
    LogMessage strProcName, , "Process took " & cMinutes & ":" & cSeconds

    ProcessTookHowLong = cMinutes & ":" & cSeconds

Block_Exit:
    Exit Function
    
Block_Err:
    ReportError Err, strProcName
    Resume Block_Exit
End Function


            '' ##############################################################################
            '' ##############################################################################
            '' ##############################################################################

Public Function HowManyTimesFoundInString(strString As String, strFind As String, vbCompareMethod As vbCompareMethod) As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim aryNumFound() As String

    strProcName = ClassName & "."

    If InStr(1, strString, strFind, vbCompareMethod) = 0 Then
        HowManyTimesFoundInString = 0
    Else
        aryNumFound = Split(strString, strFind, -1, vbCompareMethod)
        HowManyTimesFoundInString = UBound(aryNumFound) ' Yes, array's are zero based (by default) but we are not counting
                                                        ' the sections, we are counting how many times the delimiter was found
    End If

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


            '' ##############################################################################
            '' ##############################################################################
            '' ##############################################################################
    ' This isn't really mine - just modified ever so slightly..
    ' anyway, will return true or false if the Image name (as seen in the task manager)
    ' is running or not..

Public Function IsImageNameRunning(strImageName As String) As Boolean
Dim oProc, oWMIServ, colProc
Dim strPC, strList
Dim strSpace

    strPC = "."
    
    Set oWMIServ = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strPC & "\root\cimv2")
    
    Set colProc = oWMIServ.ExecQuery("Select * from Win32_Process")
    
    strSpace = String(20, " ")
    
    For Each oProc In colProc
        If UCase(oProc.Name) = UCase(strImageName) Then
            IsImageNameRunning = True
            GoTo Block_Exit
        End If
    Next
Block_Exit:
    Exit Function
Block_Err:
    MsgBox "Error: " & Err.Description
    GoTo Block_Exit
End Function


        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################



' Returns true if the passed string is a valid table in this database
Public Function IsTable(ByVal strTableName As String) As Boolean
Dim objTblDef As TableDef
Dim objDB As Database

    Set objDB = CurrentDb()
    
    IsTable = False
    For Each objTblDef In objDB.TableDefs
        If UCase(strTableName) = UCase(objTblDef.Name) Then
            IsTable = True
            Exit Function
        End If
    Next
    Set objTblDef = Nothing
    Set objDB = Nothing
    
End Function



        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################



' Returns true if the passed string is a valid table in this database
Public Function IsQuery(ByVal strQueryName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim objQDef As DAO.QueryDef
Dim objDB As DAO.Database

    strProcName = ClassName & ".IsQuery"

    Set objDB = CurrentDb()
    
    For Each objQDef In objDB.QueryDefs
        If UCase(strQueryName) = UCase(objQDef.Name) Then
            IsQuery = True
            GoTo Block_Exit
        End If
    Next
    
Block_Exit:
    Set objQDef = Nothing
    Set objDB = Nothing

    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################



Public Function CurrentDBDir() As String
Dim strDBPath As String
Dim strDBFile As String
    strDBPath = CurrentDb.Name
    strDBFile = Dir(strDBPath, vbHidden)
    CurrentDBDir = left$(strDBPath, Len(strDBPath) - Len(strDBFile))
    If Right(CurrentDBDir, 1) = "\" Then CurrentDBDir = left(CurrentDBDir, Len(CurrentDBDir) - 1)
End Function




        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################


Private Function KeyExistsInCollection(ByVal strKey As String, ByVal cCollection As Collection) As Boolean
On Error GoTo Block_Err
Dim iItemCount As Integer
Dim strProcName As String

    strProcName = ClassName & ".KeyExistsInCollection"
    
    For iItemCount = 1 To cCollection.Count
        If UCase(strKey) = UCase(cCollection.Item(iItemCount)) Then
            KeyExistsInCollection = True
            GoTo Block_Exit
        End If
    Next

Block_Exit:
    Exit Function

Block_Err:
    ReportError Err, strProcName, strKey, False
    GoTo Block_Exit
End Function





        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################

Public Function Create_Zip_File(strFileToZip As String, strFolderToZipIn As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim srcfolderString As String
Dim dstfolderString As String
Dim abFileContents()    ' As Byte
Dim objShell As Shell32.Shell
Dim oSourceFldrItem As Shell32.FolderItem
Dim objFolderSrc As Shell32.Folder
Dim objFolderDst As Shell32.Folder
Dim objFolderItems As Shell32.FolderItems
Dim sOrigExtension As String
Dim lFileHdl As Long
Dim sSourceFolder As String
Dim sShortSourceFile As String
Dim oFso As Scripting.FileSystemObject
Dim iLoop As Integer
Dim lItemIndex As Long
Dim dtStarted As Date

    strProcName = ClassName & ".Create_Zip_File"
    If Len(strFileToZip) = 0 Or FileExists(strFileToZip) = False Then
        LogMessage strProcName, , "No file pasted or file does not exist where specified"
        GoTo Block_Exit
    End If
    Create_Zip_File = True

    sOrigExtension = Right(strFileToZip, 4)
    strFileToZip = left(strFileToZip, Len(strFileToZip) - 4)
    If Right(strFolderToZipIn, 1) <> "\" Then strFolderToZipIn = strFolderToZipIn & "\"


    Set oFso = New Scripting.FileSystemObject
    If oFso.FolderExists(strFolderToZipIn) = False Then
        LogMessage strProcName, , "Creating folders needed: " & strFolderToZipIn
        CreateFolders strFolderToZipIn
    End If

    sSourceFolder = oFso.GetParentFolderName(strFileToZip & sOrigExtension)
    sShortSourceFile = oFso.GetFileName(strFileToZip)

    If oFso.FileExists(strFolderToZipIn & sShortSourceFile & ".zip") = True Then
        oFso.DeleteFile strFolderToZipIn & sShortSourceFile & ".zip", True
    End If

        'create empty zip file
    abFileContents = Array(80, 75, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    LogMessage strProcName, , "Creating zip file"
    
    lFileHdl = FreeFile
    Open strFolderToZipIn & sShortSourceFile & ".zip" For Binary As lFileHdl
    For iLoop = 0 To UBound(abFileContents)
        Put #lFileHdl, iLoop + 1, CByte(abFileContents(iLoop))
    Next
    Close #lFileHdl

    Set objShell = New Shell32.Shell

    Set objFolderSrc = objShell.Namespace(sSourceFolder)
    Set objFolderDst = objShell.Namespace(strFolderToZipIn & sShortSourceFile & ".zip")

    For Each oSourceFldrItem In objFolderSrc.Items
        If oSourceFldrItem.Name = sShortSourceFile & sOrigExtension Then
            Set objFolderDst = objShell.Namespace(strFolderToZipIn & sShortSourceFile & ".zip")
            
            objFolderDst.CopyHere oSourceFldrItem, 20

            Sleep 3000      ' sleeping to make sure that the copy starts...
            ' we should be a little more careful actually:

            dtStarted = Now

            While FileExists(strFolderToZipIn & sShortSourceFile & ".zip") = False And DateDiff("s", dtStarted, Now()) < 120
                Debug.Print "File doesn't exist yet..."
                Sleep 1000
                DoEvents
            Wend

            dtStarted = Now
            While IsFileGrowing(strFolderToZipIn & sShortSourceFile & ".zip")
                LogMessage strProcName, , "File still growing"
                If DateDiff("n", dtStarted, Now()) > 10 Then
                    LogMessage strProcName, "ERROR", "File appears to still be growing... " & strFolderToZipIn & sShortSourceFile & ".zip"
                    Create_Zip_File = False
                    GoTo Block_Exit
                End If
                Sleep 1000
                DoEvents
            Wend

            GoTo Block_Exit
        End If
        lItemIndex = lItemIndex + 1
    Next

    Sleep 3000
    LogMessage strProcName, , "Finished zipping file"

Block_Exit:
    Set oFso = Nothing
    Set objFolderDst = Nothing
    Set objFolderItems = Nothing
    Set objFolderSrc = Nothing
    Set objShell = Nothing

    Exit Function

Block_Err:
    Create_Zip_File = False
    ReportError Err, strProcName
    Resume Block_Exit
End Function
        



        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################

        ' KD Comeback: Need to set a timeout for this..
Public Function IsFileGrowing(strFileToCheck As String) As Boolean
'On Error GoTo Funct_Err
On Error Resume Next
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Static dblLastFileSize As Double
Dim dblFileSize As Double

    strProcName = ClassName & ".IsFileGrowing"
    
    Set oFso = New Scripting.FileSystemObject
    
    If oFso.FileExists(strFileToCheck) = False Then ' Seems to return false if the file is locked..
        Sleep 1500
        IsFileGrowing = True
        GoTo Funct_Exit
    End If


    If dblLastFileSize = 0 Then
        Debug.Print "Last file size = 0"
        Sleep 1500
        dblLastFileSize = oFso.GetFile(strFileToCheck).Size
        DoEvents
    End If
    
    
    dblFileSize = oFso.GetFile(strFileToCheck).Size
    If dblLastFileSize = dblFileSize Then
        IsFileGrowing = False
        dblLastFileSize = 0
    Else
        IsFileGrowing = True
        dblLastFileSize = dblFileSize
    End If
    
Funct_Exit:
    Set oFso = Nothing
    Exit Function
Funct_Err:
    IsFileGrowing = True
    ReportError Err, strProcName, strFileToCheck
    Resume Funct_Exit
End Function





        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################

Public Sub FinishMethod()
    DoCmd.Hourglass False
End Sub

Public Sub StartMethod()
    DoCmd.Hourglass True
End Sub


        ''' ##############################################################################
        ''' ##############################################################################
        ''' ##############################################################################


' Call this to sort a dictionary (the one passed) by key or item
Public Function SortDictionary(ByRef objDict As Scripting.Dictionary, ByVal intSortBy As SortDictBy, Optional bDescendingOrder As Boolean = False) As Boolean
On Error GoTo Block_Exit
Dim strProcName As String
Dim strDict() As Variant
Dim objKey As Variant
Dim strKey, strItem
Dim iItemIdx As Integer, iInnerItemIdx As Integer, iItemCount As Integer
Const dictKey As Integer = SortDictBy.ByKey
Const dictItem As Integer = SortDictBy.ByItem

    strProcName = ClassName & ".SortDictionary"

    ' get the dictionary count
    iItemCount = objDict.Count
    
    ' we need more than one item to warrant sorting
    If iItemCount > 1 Then
        ' create an array to store dictionary information
        ReDim strDict(iItemCount, 2)
        iItemIdx = 0
        
        ' populate the string array
        For Each objKey In objDict
            strDict(iItemIdx, dictKey) = CStr(objKey)
            strDict(iItemIdx, dictItem) = CStr(objDict(objKey))
            iItemIdx = iItemIdx + 1
        Next
        
        ' perform a a shell sort of the string array
        For iItemIdx = 0 To (iItemCount - 2)
            For iInnerItemIdx = iItemIdx To (iItemCount - 1)
            If StrComp(strDict(iItemIdx, intSortBy), strDict(iInnerItemIdx, intSortBy), vbTextCompare) > 0 Then
                strKey = strDict(iItemIdx, dictKey)
                strItem = strDict(iItemIdx, dictItem)
                strDict(iItemIdx, dictKey) = strDict(iInnerItemIdx, dictKey)
                strDict(iItemIdx, dictItem) = strDict(iInnerItemIdx, dictItem)
                strDict(iInnerItemIdx, dictKey) = strKey
                strDict(iInnerItemIdx, dictItem) = strItem
            End If
            Next
        Next
        
            ' erase the contents of the dictionary object
        objDict.RemoveAll
        
            ' repopulate the dictionary with the sorted information
        If bDescendingOrder = True Then
            For iItemIdx = (iItemCount - 1) To 0 Step -1
                objDict.Add strDict(iItemIdx, dictKey), strDict(iItemIdx, dictItem)
            Next
        Else
            For iItemIdx = 0 To (iItemCount - 1)
                objDict.Add strDict(iItemIdx, dictKey), strDict(iItemIdx, dictItem)
            Next
        End If
    
    End If
    
    SortDictionary = True
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function SetDefaultPrinterToAcrobat(sOrigPrinter, Optional sSetPrinterTo As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oPrinter As Printer

    strProcName = ClassName & ".SetDefaultPrinterToAcrobat"
    
    If sSetPrinterTo = "" Then
        sSetPrinterTo = "Adobe PDF"
    End If
    
    ' First, get the default printer's name so we can return it (and eventually pass it back to this
    ' function to reset..
    
    
    Set oPrinter = Application.Printer
    sOrigPrinter = oPrinter.DeviceName
    ' HP_P3015_01_PCL6 on ssconsho01-001 (redirected 2)

    Set Application.Printer = Application.Printers(sSetPrinterTo)
    
    SetDefaultPrinterToAcrobat = True
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function SetDefaultPrinterToAcrobatAPI(sOrigPrinter, Optional sSetPrinterTo As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
'Dim oPrinter As Printer
Dim lRet As Long

    strProcName = ClassName & ".SetDefaultPrinterToAcrobatAPI"
    ' HP_P3015_01_PCL6 on ssconsho01-001 (redirected 7)
    If sSetPrinterTo = "" Then
        sSetPrinterTo = "Adobe PDF"
    End If
    
    ' First, get the default printer's name so we can return it (and eventually pass it back to this
    ' function to reset..
    
    
'    Set oPrinter = Application.Printer
    'sOrigPrinter = oPrinter.DeviceName
    sOrigPrinter = DefaultPrinterInfo()
    Debug.Print sOrigPrinter
    ' HP_P3015_01_PCL6 on ssconsho01-001 (redirected 2)

    lRet = SetDefaultPrinter(sSetPrinterTo)
    
    SetDefaultPrinterToAcrobatAPI = True
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function DefaultPrinterInfo() As String
Dim strLPT As String * 255
Dim Result As String
Dim ResultLength As Long
Dim Comma1 As Integer, Comma2 As Integer
Dim Driver As String
Dim Port As String
Dim sPrinter As String

    Call GetProfileStringA("Windows", "Device", "", strLPT, 254)
    
    Result = TrimNull(strLPT)
    ResultLength = Len(Result)

    Comma1 = InStr(1, Result, ",", 1)
    Comma2 = InStr(Comma1 + 1, Result, ",")

'   Gets printer's name
    sPrinter = left(Result, Comma1 - 1)
    DefaultPrinterInfo = sPrinter
'   Gets driver
    Driver = Mid(Result, Comma1 + 1, Comma2 - Comma1 - 1)

'   Gets last part of device line
    Port = Right(Result, ResultLength - Comma2)

    Debug.Print sPrinter
    Debug.Print Driver
    Debug.Print Port

End Function

Public Function KDShowFormAndWait(oFrm As Form) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim strFormName As String
Dim bFormClosed As Boolean

    strProcName = ClassName & ".KDShowFormAndWait"
    
    ColObjectInstances.Add oFrm, oFrm.hwnd & ""
'    oFrm.ConceptId = Me.txtConceptId.Value
'    oFrm.PayerNameID = Me.PayerNameID
'    oFrm.RefreshData
    
    strFormName = oFrm.Name
    oFrm.visible = True
    DoEvents
    Wait 1
    DoEvents
    Do
        'Is it still Open?
        If IsLoaded(strFormName) Then
            DoEvents
            Wait 1
        ElseIf oFrm.visible = False Then
            bFormClosed = True
        Else
            bFormClosed = True
        End If
        
        If Nz(oFrm.SelectedId, "") <> "" And oFrm.visible = False Then
            bFormClosed = True
        End If
        If oFrm.Canceled = True And oFrm.visible = False Then
            bFormClosed = True
        End If
       
    Loop Until bFormClosed = True
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function BrowseForFolderMSOFfice(Title As String, Optional InitialFolder As String = vbNullString, Optional InitialView As Office.MsoFileDialogView = _
            msoFileDialogViewList) As String
Dim V As Variant
Dim InitFolder As String
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".BrowseForFolderMSOFfice"
    
'    With Application.FileDialog(msoFileDialogFilePicker)
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = Title
        .InitialView = InitialView
        If Len(InitialFolder) > 0 Then
            If Dir(InitialFolder, vbDirectory) <> vbNullString Then
                InitFolder = InitialFolder
                If Right(InitFolder, 1) <> "\" Then
                    InitFolder = InitFolder & "\"
                End If
                .InitialFileName = InitFolder
            End If
        End If
        .show
            ' later we should change this to accept multiple selections
        If .SelectedItems.Count > 0 Then
            V = .SelectedItems(1)
        Else
            V = vbNullString
        End If
    End With
    BrowseForFolderMSOFfice = CStr(V)
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function GetFileCountFromDir(sFolder As String, Optional bIncludeSubDirs As Boolean = False, Optional sFileFilter As String, Optional bFilterForExtensionsOnly As Boolean = True) As Long
On Error GoTo Block_Exit
Dim strProcName As String
Dim lRet As Long
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oSubFldr As Scripting.Folder
Dim oFile As Scripting.file
Dim sExt As String
Dim saryFilters() As String
Dim sFltr As String
Dim iFltrIdx As Integer


    strProcName = ClassName & ".GetFileCountFromDir"
    Set oFso = New Scripting.FileSystemObject
    
    Set oFldr = oFso.GetFolder(sFolder)
    If bIncludeSubDirs = True Then
        For Each oSubFldr In oFldr.SubFolders
            lRet = lRet + GetFileCountFromDir(oSubFldr.Path, bIncludeSubDirs, sFileFilter)
        Next
    End If
    
    If sFileFilter = "" Then
        If oFldr.SubFolders.Count > 0 Then
            lRet = lRet + (oFldr.Files.Count - oFldr.SubFolders.Count)
        Else
            lRet = lRet + oFldr.Files.Count
        End If
        
    Else
    
        saryFilters = Split(sFileFilter, ";")
        
    
        For Each oFile In oFldr.Files
            For iFltrIdx = 0 To UBound(saryFilters)
                sFltr = saryFilters(iFltrIdx)
                sFltr = Replace(sFltr, "*", "")
                
                If bFilterForExtensionsOnly = True Then
                    sExt = oFso.GetExtensionName(oFile.Path)
                    If InStr(1, sExt, sFltr, vbTextCompare) > 0 Then
                        lRet = lRet + 1
                        GoTo NextFile
                    End If
                Else
                    If InStr(1, oFile.Name, sFltr, vbTextCompare) > 0 Then
                        lRet = lRet + 1
                        GoTo NextFile
                    End If
                End If
NextFltr:
            Next
NextFile:
        Next
    End If
    
    
Block_Exit:
    GetFileCountFromDir = lRet
    Set oSubFldr = Nothing
    Set oFile = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
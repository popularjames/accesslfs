''''Option Compare Database
''''Option Explicit
''''
''''Private Const strProcName As String = "mod_ClaimAdmin_Tools"
''''Private Const ClassName As String = "mod_ClaimAdmin_Tools"
''''
''''Private cdctLinkedTables As Scripting.Dictionary
''''
''''
''''
'''''' Last modified: 04/23/2015
'''''' 04/23/2015 KD: Added FileSpec code - to loop over the file specs saved with the database
'''''' 05/13/2013 KD: Adjusted the pop up message so we can deploy without prompting users
'''''' 05/08/2013 KD: Adjusted the pop up message to include the build numbers involved
''''    ''  with a later plan to adjust it so we can deploy immediately without prompting the
''''    ''  user, but the nightly deploy would do it...
'''''' 04/26/2013 KD: Added TurnOffDeveloperErrorHandling
''''
''''Public Const gs_STARTUP_TBL_NAME As String = "CT_AppStartupSeq"
''''
''''
''''' Use this when you updated the file number
''''Public Function UpdateLatestVersion(Optional lOldBuildId As Long, Optional lNewBuildId As Long, Optional lAVCID As Long) As Long
''''On Error GoTo Block_Err
''''Dim strProcName As String
''''Dim lFileVer As Long
''''Dim lTableVer As Long
''''Dim lDbVer As Long
''''Dim lLocalTblVer  As Long
''''Dim oDb As DAO.Database
''''Dim oTDef As DAO.TableDef
''''Dim oAdo As clsADO
''''Dim sMsg As String
''''    strProcName = ClassName & ".UpdateLatestVersion"
''''
''''    lFileVer = FileNameVersion()
''''    lDbVer = GetSQLServerVersionNum()
''''    lLocalTblVer = GetLocalVersionNum()
''''
''''    lOldBuildId = lDbVer
''''
''''    ' 20130426 KD: I'm going to ignore the file version since we are so close to using the Version Control
''''    If lDbVer > lLocalTblVer Then
''''        lOldBuildId = lDbVer
''''    ElseIf lLocalTblVer > lDbVer Then
''''        lOldBuildId = lLocalTblVer
''''    End If
''''
''''    lNewBuildId = GetMaxBuild() + 1
''''
''''    ' Now update the local table
''''    Call GetLocalVersionNum(lNewBuildId)
''''        '    Call GetLocalVersionNum(0)
''''
''''    ' and update SQL Server
''''    Call GetSQLServerVersionNum(lNewBuildId, , , lAVCID)
''''        '    Call GetSQLServerVersionNum(0)
''''
''''    If sMsg <> "" Then
''''        If Application.UserControl = True Then
''''            MsgBox sMsg, vbInformation + vbExclamation, "Don't forget!!!"
''''        End If
''''    End If
''''
''''
''''Block_Exit:
''''
''''    Exit Function
''''Block_Err:
''''    ReportError Err, strProcName
''''    GoTo Block_Exit
''''End Function
''''
''''' THis function gets the MAX build from the database.. Because if we don't
''''' do this then we risk assigning the same build id
''''' however, what this means is that we may skip numbers or maybe even go backwards.
''''' but, it's friday after 5 pm, I'm tired and can't think this through..
''''' so, whatever!!!
''''Public Function GetMaxBuild() As Long
''''On Error GoTo Block_Err
''''Dim strProcName As String
''''Dim oAdo As clsADO
''''Dim oRs As ADODB.RecordSet
''''
''''
''''    strProcName = ClassName & ".GetMaxBuild"
''''
''''    Set oAdo = New clsADO
''''    With oAdo
''''        '.ConnectionString = GetConnectString("v_Data_Database")
''''        .ConnectionString = GetConnectString("v_Workspace_Database")
''''        .SQLTextType = sqltext
''''        .sqlString = "SELECT MAX(Build) AS BuildID FROM AVC_Claim_Admin_Version"
''''        Set oRs = .ExecuteRS
''''        If .GotData = False Then
''''            GetMaxBuild = 0
''''        Else
''''            GetMaxBuild = oRs("BuildId").Value
''''        End If
''''    End With
''''
''''Block_Exit:
''''    If oRs.State = adStateOpen Then oRs.Close
''''    Set oRs = Nothing
''''    Set oAdo = Nothing
''''    Exit Function
''''Block_Err:
''''    ReportError Err, strProcName
''''    GoTo Block_Exit
''''End Function
''''
''''Public Sub CheckVersion()
''''On Error GoTo Block_Err
''''Dim strProcName As String
''''Dim lThisVersion As Long
''''Dim lDbVer As Long
''''Dim sMsg As String
''''Static dtFirstFound As Date
''''Static iTimesNewVersionFound As Integer
''''Const iMinutesToLag As Integer = 5
''''Dim bPromptUserToReload As Boolean
''''
''''    ' 20130412 KD: Changed this to allow 5 minutes (iMinutesToLag) before they are prompted..
''''
''''    strProcName = ClassName & ".CheckVersion"
''''
''''    lThisVersion = GetLocalVersionNum()
''''
''''    lDbVer = GetSQLServerVersionNum(, , bPromptUserToReload)
''''
''''    If lThisVersion < lDbVer Then
''''        iTimesNewVersionFound = iTimesNewVersionFound + 1
''''        Select Case iTimesNewVersionFound
''''        Case 0, 1
''''            dtFirstFound = Now()
''''            GoTo Block_Exit
''''        Case Else
''''            If DateDiff("n", dtFirstFound, Now()) < iMinutesToLag Then
''''                GoTo Block_Exit
''''            End If
''''        End Select
''''
''''        sMsg = "A newer version of Claim Admin has been deployed. (Build: " & CStr(lDbVer) & " - you have version " & CStr(lThisVersion) & ")" & _
''''            vbCrLf & "At your earliest convienance, please close this version of Claim Admin and relaunch Claim Admin from the _Launch shortcut!"
''''        LogMessage strProcName, , sMsg & " " & " Displaying: " & CStr(bPromptUserToReload)
''''        If bPromptUserToReload = True Then
''''            MsgBox sMsg, vbInformation + vbExclamation, "A newer version of Claim Admin is available!"
''''        End If
''''    End If
''''
''''Block_Exit:
''''    Exit Sub
''''Block_Err:
''''    ReportError Err, strProcName
''''    GoTo Block_Exit
''''End Sub
''''
''''
''''Public Function FileNameVersion() As Long
''''On Error GoTo Block_Err
''''Dim strProcName As String
''''Dim oRegEx As RegExp
''''Dim oMatchCol As MatchCollection
''''Dim oMatch As Match
''''Dim sThisFileName As String
''''Dim lFileVer As Long
''''
''''    strProcName = ClassName & ".FileNameVersion"
''''
''''    ' What does the filename have in it?
''''    sThisFileName = CurrentDb.Name
''''
''''    Set oRegEx = New RegExp
''''    With oRegEx
''''        .Global = False
''''        .IgnoreCase = True
''''        .MultiLine = False
''''        .Pattern = "\.(\d{3})\D+"
''''        Set oMatchCol = .Execute(sThisFileName)
''''
''''    End With
''''
''''    If oMatchCol.Count = 1 Then
''''        Set oMatch = oMatchCol.Item(0)
''''                    Debug.Print oMatch.SubMatches(0)
''''        If IsNumeric(oMatch.SubMatches(0)) = False Then
''''            lFileVer = 0    ' We'll have to use something else...
''''        Else
''''            lFileVer = CLng(oMatch.SubMatches(0))
''''        End If
''''    End If
''''
''''
''''Block_Exit:
''''    FileNameVersion = lFileVer
''''    Set oMatch = Nothing
''''    Set oMatchCol = Nothing
''''    Set oRegEx = Nothing
''''    Exit Function
''''Block_Err:
''''    ReportError Err, strProcName
''''    GoTo Block_Exit
''''End Function
''''
''''
''''Public Function GetLocalVersionNum(Optional lNewVerNum As Long) As Long
''''On Error GoTo Block_Err
''''Dim strProcName As String
''''Dim oDb As DAO.Database
''''Dim oTDef As DAO.TableDef
''''Dim sLocalTblVer As String
''''Dim lLocalTblVer As Long
''''Dim sDesc As String
''''Dim oRegEx As RegExp
''''
''''    strProcName = ClassName & ".GetLocalVersionNum"
''''
''''    ' Get what we have in the the About screen:
''''    Set oDb = CurrentDb()
''''    If IsTable(gs_STARTUP_TBL_NAME) = False Then
''''        LogMessage strProcName, "ERROR", "Our table is missing! Call Sherlock Holmes!", gs_STARTUP_TBL_NAME
''''        GoTo Block_Exit
''''    End If
''''    Set oTDef = oDb.TableDefs(gs_STARTUP_TBL_NAME)
''''    sDesc = oTDef.Properties("Description")
''''
''''    Set oRegEx = New RegExp
''''    With oRegEx
''''        .Global = False
''''        .IgnoreCase = True
''''        .MultiLine = False
''''        .Pattern = "^(.+?\.)(\d{3})(\s*?)$"
''''    End With
''''    ' 3.0.1101 CA: 01.060
''''
''''    sLocalTblVer = oRegEx.Replace(sDesc, "$2")
''''
''''    If IsNumeric(sLocalTblVer) = False Then
''''        lLocalTblVer = 0
''''    Else
''''        lLocalTblVer = CLng(sLocalTblVer)
''''    End If
''''
''''    If lNewVerNum > 0 Then
''''        If oRegEx.test(sDesc) = False Then
''''            LogMessage strProcName, "ERROR", "Hmm.. The Regex didn't find the pattern! Check the description on the '" & gs_STARTUP_TBL_NAME & "' maybe it changed?", , True
''''            GoTo Block_Exit
''''        Else
''''            sDesc = oRegEx.Replace(sDesc, "$1" & Format(lNewVerNum, "000"))
''''            oDb.TableDefs(gs_STARTUP_TBL_NAME).Properties("Description") = sDesc
''''            oDb.TableDefs.Refresh
''''        End If
''''
''''    End If
''''
''''
''''
''''Block_Exit:
''''    Set oTDef = Nothing
''''    Set oDb = Nothing
''''    Set oRegEx = Nothing
''''    GetLocalVersionNum = lLocalTblVer
''''    Exit Function
''''Block_Err:
''''    ReportError Err, strProcName
''''    GoTo Block_Exit
''''End Function
''''
''''
''''Public Function GetSQLServerVersionNum(Optional lNewVer As Long, Optional bDeployed As Boolean = False, _
''''    Optional bPromptUserToReload As Boolean = False, Optional lAVC_ID As Long) As Long
''''On Error GoTo Block_Err
''''Dim oAdo As clsADO
'''''Dim oRs As ADODB.Recordset
''''Dim lDbVer As Long
''''Dim strProcName As String
''''
''''    strProcName = ClassName & ".GetSQLServerVersionNum"
''''
''''    ' What does FLD-009 think the latest version is?
''''    Set oAdo = New clsADO
''''    With oAdo
''''        .ConnectionString = GetConnectString("v_CODE_Database")
''''        .SQLTextType = StoredProc
''''            '        .SqlString = "SELECT TOP 1 * FROM ADMIN_Claim_Admin_Version ORDER BY Build DESC"
''''        .sqlString = "usp_ADMIN_Claim_Admin_Version_Check"
''''        .Parameters.Refresh
''''
''''        .Execute
''''        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
''''            LogMessage strProcName, "ERROR", "Did not get the current version for some reason!!"
''''            GoTo Block_Exit
''''        End If
''''        lDbVer = .Parameters("@pServerVersion").Value
''''        bPromptUserToReload = CBool(.Parameters("@pPromptUserToReload").Value)
''''
''''    End With
''''
''''    If lNewVer > 0 Then
''''        lDbVer = lNewVer
''''        With oAdo
''''            .SQLTextType = StoredProc
''''            .sqlString = "usp_ADMIN_Claim_Admin_Version_Set"
''''            .Parameters.Refresh
''''            .Parameters("@pNewBuild") = lNewVer
''''            If bDeployed = True Then
''''                .Parameters("@pDeployDt") = Now() ' Now()
''''            End If
''''            If lAVC_ID <> 0 Then
''''                .Parameters("@pAVC_ID") = lAVC_ID
''''            End If
''''            .Execute
''''            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
''''                LogMessage strProcName, "ERROR", "An error occurred when trying to update the build of Claim Admin: " & .Parameters("@pErrMsg").Value
''''            End If
''''        End With
''''    End If
''''
''''Block_Exit:
''''
''''    Set oAdo = Nothing
''''    GetSQLServerVersionNum = lDbVer
''''    Exit Function
''''Block_Err:
''''    ReportError Err, strProcName
''''    GoTo Block_Exit
''''End Function
''''
''''
''''
'''''
''''''' Note: to use this you must change the ReadBinaryFile to Public - then change it back :)
'''''Public Function SaveIcon()
'''''Dim oIcon As CT_ClsIcon
'''''
'''''
'''''    Set oIcon = New CT_ClsIcon
'''''    oIcon.ReadBinaryFile "M:\2010_version\Claims_Admin.ico", "Claims Admin"
'''''End Function
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' This function is the main one I use for searching for keywords in a
''''''' MS Access project.. I pretty much just put the keyword in the strKeyword
''''''' then uncomment the calls for where I want to look
'''''''
''''Private Function ReportOnObjects() As String
''''Dim objTable As TableDef
''''Dim objQuery As QueryDef
''''Dim objSysTbl As DAO.RecordSet
''''Dim objSysFld As DAO.Field
''''Dim strKeyword As String
''''Dim strNameLike As String
''''Dim strSql As String
''''Dim strMessage As String
''''Dim intUsedInQueries As Integer
''''Dim intUsedInForms As Integer
''''Dim intUsedInMacros As Integer
''''Dim intUsedInModules As Integer
''''Dim intUsedInReports As Integer
''''Dim strAlias As String  ' for queries
''''Dim dctTblDetails As Scripting.Dictionary
''''
''''
''''    ' Right.. First look at the tables:
''''    ' How many are linked tables?
''''    ' get:  Access name, Foreign Name, Database date update
''''    ' Header for this section:
'''''    LogMessage "Access linked tables", strProcName
''''
'''''    strMessage = "Name,ObjectType,ForeignName,Database,DSN,DateUpdate,DateCreate,UsedInQueries,UsedInMacros,UsedInModules,UsedInForms,UsedInReports"
'''''    LogMessage strMessage
''''
''''
''''
''''    ' Here we'll call a sub to count how many references there are in:
''''    ' Queries
''''    ' Modules
''''    ' Macros
''''Dim sFoundList As String
''''
''''
''''    strKeyword = "New clsTable"
''''
''''
'''''    intUsedInQueries = FindQueriesUsingObject(strKeyword, False, False, sFoundList)
''''
'''''
'''''    Debug.Print
'''''    Debug.Print
'''''    Debug.Print "Queries:"
'''''    Debug.Print sFoundList
'''''
'''''Stop
'''''    Debug.Print FindLinkedTableByForeignName(strKeyword)
'''''
'''''    Set dctTblDetails = GetLinkedInfoFromObject(strKeyword)
''''
''''
'''''    Debug.Print FindLinkedTablesInDb(strKeyword)
'''''
'''''    sFoundList = ""
'''''    intUsedInReports = FindReportsUsingObject(strKeyword, True, False, sFoundList)
'''''
'''''    Debug.Print
'''''    Debug.Print
'''''    Debug.Print "Reports: " & CStr(intUsedInReports)
'''''    Debug.Print sFoundList
'''''
'''''Stop
''''    sFoundList = ""
''''
''''    intUsedInForms = FindFormsUsingObject(strKeyword, True, False, sFoundList)
''''
''''
''''    Debug.Print
''''    Debug.Print
''''    Debug.Print "Forms: " & CStr(intUsedInForms)
''''    Debug.Print sFoundList
''''
''''    Stop
''''    sFoundList = ""
''''
''''''    intUsedInMacros = FindMacrosUsingObject(strKeyword)
''''    intUsedInModules = FindModulesUsingObject(strKeyword, False, False, sFoundList)
''''
''''
''''    Debug.Print
''''    Debug.Print
''''    Debug.Print "Modules: " & CStr(intUsedInModules)
''''    Debug.Print sFoundList
''''
''''    Stop
''''    sFoundList = ""
''''
''''    Stop
'''''    intUsedInQueries = ResolveRecursiveQueries(strKeyword, "")
''''
''''
'''''    strMessage = Join(Array(strKeyword, "Table", dctTblDetails("ForeignName"), dctTblDetails("Database"), _
'''''        dctTblDetails("Connect"), dctTblDetails("DateUpdate"), dctTblDetails("DateCreate"), intUsedInQueries, _
'''''        intUsedInMacros, intUsedInModules, intUsedInForms, intUsedInReports), ",")
'''''    LogMessage strMessage, strprocname
''''
''''GoTo Block_Exit
''''
''''    If Not dctTblDetails Is Nothing Then
''''        Dim vKey As Variant
''''        Debug.Print strKeyword & " Details:"
''''
''''        For Each vKey In dctTblDetails.Keys
''''            Debug.Print " " & CStr(vKey) & " = " & dctTblDetails.Item(vKey)
''''        Next
''''    End If
''''
''''    Exit Function
''''
'''''
'''''    For Each objTable In CurrentDb().TableDefs
'''''
'''''        ' Here we'll call a sub to count how many references there are in:
'''''        ' Queries
'''''        ' Modules
'''''        ' Macros
'''''        If LCase(Left(objTable.Name, 4)) <> "msys" Then
'''''            intUsedInQueries = FindQuerysUsingObject(objTable.Name)
'''''
'''''            Set dctTblDetails = GetLinkedInfoFromObject(objTable.Name)
'''''
'''''            intUsedInReports = FindReportsUsingObject(objTable.Name)
'''''
'''''            intUsedInForms = FindFormsUsingObject(objTable.Name)
'''''            intUsedInMacros = FindMacrosUsingObject(objTable.Name)
'''''            intUsedInModules = FindModulesUsingObject(objTable.Name)
'''''
'''''
'''''            With objTable
'''''                strMessage = Join(Array(.Name, "Table", dctTblDetails("ForeignName"), dctTblDetails("Database"), _
'''''                    dctTblDetails("Connect"), dctTblDetails("DateUpdate"), dctTblDetails("DateCreate"), intUsedInQueries, _
'''''                    intUsedInMacros, intUsedInModules, intUsedInForms, intUsedInReports), ",")
'''''            End With
'''''            LogMessage strMessage, strProcName
'''''        End If
'''''    Next
'''''
'''''    ' next, queries
'''''    For Each objQuery In CurrentDb().QueryDefs
'''''
'''''        ' Here we'll call a sub to count how many references there are in:
'''''        ' Queries
'''''        ' Modules
'''''        ' Macros
'''''            intUsedInQueries = FindQuerysUsingObject(objQuery.Name)
'''''
'''''            Set dctTblDetails = GetLinkedInfoFromObject(objQuery.Name)
'''''
'''''            intUsedInReports = FindReportsUsingObject(objQuery.Name)
'''''
'''''            intUsedInForms = FindFormsUsingObject(objQuery.Name)
'''''            intUsedInMacros = FindMacrosUsingObject(objQuery.Name)
'''''            intUsedInModules = FindModulesUsingObject(objQuery.Name)
'''''
'''''
'''''            With objQuery
'''''                strMessage = Join(Array(.Name, "Query", dctTblDetails("ForeignName"), dctTblDetails("Database"), _
'''''                    dctTblDetails("Connect"), dctTblDetails("DateUpdate"), dctTblDetails("DateCreate"), intUsedInQueries, _
'''''                    intUsedInMacros, intUsedInModules, intUsedInForms, intUsedInReports), ",")
'''''            End With
'''''            LogMessage strMessage, strProcName
'''''    Next
'''''
'''''
''''''    On Error Resume Next
''''''    For Each objTable In CurrentDb().l
''''''        Debug.Print objTable.Name & " Type: " & objTable.SourceTableName
''''''    Next
'''''
''''''    For Each objQuery In CurrentDb().QueryDefs
''''''        Debug.Print objQuery.Name & " Type: " & objQuery.Type
''''''    Next
''''
''''Block_Exit:
''''    ReportOnObjects = ""
''''
''''End Function
''''
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Looks for linked tables where the connection string contains:
''''''' DATABASE= ?
'''''''
''''Private Function FindLinkedTablesInDb(strDatabaseName As String) As String
''''Dim oSysTbl As DAO.RecordSet
''''Dim sReturn As String
''''
''''    Set oSysTbl = CurrentDb().OpenRecordSet("SELECT * FROM mSysObjects WHERE TYPE = 4 AND Connect LIKE ""*DATABASE=" & strDatabaseName & "*""", dbOpenSnapshot, dbReadOnly)
''''
''''    While Not oSysTbl.EOF
''''        sReturn = sReturn & oSysTbl("[Name]") & vbCrLf
''''        oSysTbl.MoveNext
''''    Wend
''''
''''    FindLinkedTablesInDb = sReturn
''''    Set oSysTbl = Nothing
''''
''''End Function
''''
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Finds a linked table name when given the foreign table's name
'''''''
''''Private Function FindLinkedTableByForeignName(sForeignKeyword As String) As String
''''Dim objSysTbl As DAO.RecordSet
''''Dim strReturn As String
''''
''''    Set objSysTbl = CurrentDb().OpenRecordSet("SELECT * FROM mSysObjects WHERE TYPE = 4 AND ForeignName LIKE ""*" & sForeignKeyword & "*""", dbOpenSnapshot, dbReadOnly)
''''
''''    If objSysTbl.EOF Then
''''        strReturn = ""
''''    Else
''''        While Not objSysTbl.EOF
''''            strReturn = strReturn & CStr("" & objSysTbl("ForeignName")) & vbCrLf
''''            objSysTbl.MoveNext
''''        Wend
''''    End If
''''
''''    If Right(strReturn, 2) = vbCrLf Then
''''        strReturn = left(strReturn, Len(strReturn) - 2)
''''    End If
''''
''''    FindLinkedTableByForeignName = strReturn
''''
''''End Function
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Gets the details about the linked table in a dictionary
'''''''
''''Private Function GetLinkedInfoFromObject(strObjectName As String) As Scripting.Dictionary
''''Dim objSysTbl As DAO.RecordSet
''''Dim dctTemp As Scripting.Dictionary
''''
''''    Set dctTemp = New Scripting.Dictionary
''''
''''    Set objSysTbl = CurrentDb().OpenRecordSet("SELECT * FROM mSysObjects WHERE NAME = """ & strObjectName & """", dbOpenSnapshot, dbReadOnly)
''''
''''    While Not objSysTbl.EOF
'''''        If Len("" & objSysTbl("Database").Value) > 0 _
'''''            Or Len("" & objSysTbl("Connect").Value) > 0 Then
''''
''''            With objSysTbl
''''                dctTemp.Add "Database", CStr("" & !Database)
''''                dctTemp.Add "Connect", CStr("" & !Connect)
''''                dctTemp.Add "ForeignName", CStr("" & !ForeignName)
''''                dctTemp.Add "DateUpdate", CStr("" & !DateUpdate)
''''                dctTemp.Add "DateCreate", CStr("" & !DateCreate)
'''''                dctTemp.Add "DateUpdate", !Database
'''''                dctTemp.Add "DateUpdate", !Database
''''            End With
'''''        End If
''''        objSysTbl.MoveNext
''''    Wend
''''    Set GetLinkedInfoFromObject = dctTemp
''''
''''End Function
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Looks through forms using the passed keyword in the recordsource
'''''''
''''Private Function FindFormsUsingObject(ByVal strObjName As String, Optional bLookInControls As Boolean = False, _
''''    Optional bStopAtFirstFound As Boolean = True, Optional sFoundInList As String) As Integer
''''Dim objForm As Form
''''Dim objAO As Object
''''Dim strName As String
''''Dim strSql As String
''''Dim oCtl As Control
''''Dim oModule As Module
''''Dim strCode As String
''''Dim iSetting As Integer
''''
''''
''''    strObjName = LCase(strObjName)
''''    iSetting = Application.GetOption("Error Trapping")
''''    Application.SetOption "Error Trapping", 2
''''
''''    On Error Resume Next
''''
''''
''''    For Each objAO In Application.CurrentProject.AllForms
'''''    Debug.Print objAO.Name
''''
''''        DoCmd.OpenForm objAO.Name, acDesign, , , , acHidden
''''        Set objForm = Application.Forms(objAO.Name)
''''
''''        strSql = objForm.RecordSource
''''
''''        Set oModule = objForm.Module
''''        strCode = oModule.Lines(1, oModule.CountOfLines)
''''
''''        If InStr(1, strCode, strObjName, vbTextCompare) > 0 Then
''''            FindFormsUsingObject = FindFormsUsingObject + 1
''''            sFoundInList = sFoundInList & objForm.Name & " (CODE),"
''''            If bStopAtFirstFound = True Then GoTo Block_Exit
''''        End If
''''
''''        If InStr(1, strSql, strObjName, vbTextCompare) > 0 Then
''''            FindFormsUsingObject = FindFormsUsingObject + 1
''''            sFoundInList = sFoundInList & objForm.Name & ","
''''            If bStopAtFirstFound = True Then GoTo Block_Exit
''''        End If
''''
''''        If bLookInControls = True Then
''''            For Each oCtl In objForm.Controls
''''                If IsProperty(oCtl, "RowSource") = True Then
''''                    If InStr(1, oCtl.Properties("RowSource").Value, strObjName, vbTextCompare) > 0 Then
''''                        FindFormsUsingObject = FindFormsUsingObject + 1
''''                        sFoundInList = sFoundInList & objForm.Name & " (CONTROL SOURCE),"
''''                        If bStopAtFirstFound = True Then GoTo Block_Exit
''''                    End If
''''                End If
''''            Next
''''        End If
''''
''''
''''        DoCmd.Close acForm, objForm.Name, acSaveNo
'''''        Unload objForm
''''
''''    Next
''''
''''Block_Exit:
''''    Application.SetOption "Error Trapping", iSetting
''''    Set oCtl = Nothing
''''    Set objForm = Nothing
''''End Function
''''
''''
''''Public Function IsProperty(oControl As Control, ByVal sPropertyName As String) As Boolean
''''On Error GoTo Block_Err
''''Dim strProcName As String
''''Dim oProp As Property
''''
''''    strProcName = ClassName & ".IsProperty"
''''    sPropertyName = LCase(sPropertyName)
''''
''''    For Each oProp In oControl.Properties
''''        If LCase(oProp.Name) = sPropertyName Then
''''            IsProperty = True
''''            GoTo Block_Exit
''''        End If
''''    Next
''''
''''Block_Exit:
''''    Set oProp = Nothing
''''    Exit Function
''''Block_Err:
''''    ReportError Err, strProcName
''''    GoTo Block_Exit
''''End Function
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Debug.prints all properties of a given object
'''''''
''''Public Sub ListControlProps(ByRef objAO As AccessObject)
''''    Dim prp As Property
''''
''''    On Error GoTo props_err
''''
''''    For Each prp In objAO.Properties
''''        Debug.Print vbTab & prp.Name & " = " & prp.Value
''''    Next prp
''''
''''props_exit:
''''    Set prp = Nothing
''''Exit Sub
''''
''''props_err:
''''    If Err = 2187 Then
''''        Debug.Print vbTab & prp.Name & " = Only available at design time."
''''        Resume Next
''''    Else
''''        Debug.Print vbTab & prp.Name & " = Error Occurred: " & Err.Description
''''        Resume Next
''''    End If
''''End Sub
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Debug.prints all properties of a given object
'''''''
''''Public Sub ListControlProps2(ByRef objCtrl As Control)
''''    Dim prp As Property
''''
''''    On Error GoTo props_err
''''
''''    For Each prp In objCtrl.Properties
''''        Debug.Print vbTab & prp.Name & " = " & prp.Value
''''    Next prp
''''
''''props_exit:
''''    Set prp = Nothing
''''Exit Sub
''''
''''props_err:
''''    If Err = 2187 Then
''''        Debug.Print vbTab & prp.Name & " = Only available at design time."
''''        Resume Next
''''    Else
''''        Debug.Print vbTab & prp.Name & " = Error Occurred: " & Err.Description
''''        Resume Next
''''    End If
''''End Sub
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Finds macros using the keyword ... never mind - never finished
'''''''
''''Private Function FindMacrosUsingObject(ByVal strObjName As String, Optional bLookInControls As Boolean = False, _
''''    Optional bStopAtFirstFound As Boolean = True, Optional sFoundInList As String) As Integer
''''Dim objAO As Object
''''Dim strName As String
''''Dim strSql As String
''''
''''    sFoundInList = ""   ' make sure we zero it out first
''''
''''    On Error Resume Next
''''    For Each objAO In Application.CurrentProject.AllMacros
''''        DoCmd.OpenForm objAO.Name, acDesign, , , , acHidden
'''''        Set objForm = Application.Forms(objAO.Name)
''''
''''        ListControlProps objAO
''''
'''''        strSQL = objForm.RecordSource
''''
''''        If InStr(1, strSql, strObjName, vbTextCompare) > 0 Then
''''            FindMacrosUsingObject = FindMacrosUsingObject + 1
''''                sFoundInList = sFoundInList & objAO.Name & ","
''''                If bStopAtFirstFound = True Then GoTo Block_Exit
''''        End If
''''
'''''        Unload objForm
''''
''''    Next
''''
''''Block_Exit:
''''    Set objAO = Nothing
''''End Function
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Finds modules using the keyword in one or more of the lines of code
'''''''
''''Private Function FindModulesUsingObject(ByVal strObjName As String, Optional bLookInControls As Boolean = False, _
''''    Optional bStopAtFirstFound As Boolean = True, Optional sFoundInList As String) As Integer
''''Dim objModule As Module
''''Dim objOA As AccessObject
''''Dim strCode As String
''''Dim iSetting As Integer
''''
''''    sFoundInList = ""   ' make sure we zero it out first
''''    iSetting = Application.GetOption("Error Trapping")
''''    Application.SetOption "Error Trapping", 2
''''
''''    On Error Resume Next
''''    For Each objOA In Application.CurrentProject.AllModules
''''        DoCmd.OpenModule objOA.Name
''''        Set objModule = Application.Modules(objOA.Name)
''''        strCode = objModule.Lines(1, objModule.CountOfLines)
'''''        Unload objOA
''''
''''
''''        If InStr(1, strCode, strObjName, vbTextCompare) > 0 Then
''''            FindModulesUsingObject = FindModulesUsingObject + 1
'''''            LogMessage objModule.name & " contains: (" & strObjName & ")", strProcName
''''Debug.Print objModule.Name & " contains: (" & strObjName & ")", strProcName
''''            sFoundInList = sFoundInList & objModule.Name & ","
''''            If bStopAtFirstFound = True Then GoTo Block_Exit
''''        End If
''''        DoCmd.Close acModule, objOA.Name
''''        Set objOA = Nothing
''''    Next
''''
''''
''''
''''Block_Exit:
''''    Application.SetOption "Error Trapping", iSetting
''''    Set objOA = Nothing
''''End Function
''''
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Finds reports using the keyword passed in the SQL source for that report
'''''''
''''Private Function FindReportsUsingObject(ByVal strObjName As String, Optional bLookInControls As Boolean = False, _
''''    Optional bStopAtFirstFound As Boolean = True, Optional sFoundInList As String) As Integer
''''Dim objReport As Report
''''Dim objAccessObj As AccessObject
''''Dim strNameLike As String
''''Dim strSql As String
''''Dim iSetting As Integer
''''
''''    sFoundInList = ""   ' make sure we zero it out first
''''    iSetting = Application.GetOption("Error Trapping")
''''    Application.SetOption "Error Trapping", 2
''''
''''    On Error Resume Next
''''    For Each objAccessObj In Application.CurrentProject.AllReports
''''        If left(objAccessObj.Name, 6) <> "LEGACY" Then
''''            DoCmd.OpenReport objAccessObj.Name, acViewPreview
''''            Set objReport = Reports(objAccessObj.Name)
''''            strSql = objReport.RecordSource
''''
''''            If InStr(1, strSql, strObjName, vbTextCompare) > 0 Then
''''                FindReportsUsingObject = FindReportsUsingObject + 1
''''                sFoundInList = sFoundInList & objReport.Name & ","
''''                If bStopAtFirstFound = True Then GoTo Block_Exit
''''            End If
''''            DoCmd.Close acReport, objAccessObj.Name, acSaveNo
''''        End If
''''    Next
''''Block_Exit:
''''    Application.SetOption "Error Trapping", iSetting
''''
''''    Set objReport = Nothing
''''    Set objAccessObj = Nothing
''''
''''End Function
''''
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Finds queries where the keyword is found in the SQL
''''' Note to self:
''''' If a query uses several other queries that use this table, we won't count it (yet)
''''Private Function FindQuerysUsingObject_LEGACY(ByVal strObjName As String, Optional bLookInControls As Boolean = False, _
''''    Optional bStopAtFirstFound As Boolean = True, Optional sFoundInList As String) As Integer
'''''Dim objTable As TableDef
''''Dim objQuery As QueryDef
''''Dim strNameLike As String
''''Dim strSql As String
''''
''''
''''    sFoundInList = ""   ' make sure we zero it out first
''''    On Error Resume Next
''''    For Each objQuery In CurrentDb().QueryDefs
''''        strSql = objQuery.SQL
''''
''''        Debug.Print objQuery.Name
''''Debug.Assert LCase(objQuery.Name) <> "query2"
''''
''''        If InStr(1, strSql, strObjName, vbTextCompare) > 0 Then
''''            FindQuerysUsingObject_LEGACY = FindQuerysUsingObject_LEGACY + 1
''''            LogMessage objQuery.Name & " contains: (" & strObjName & ")", strProcName
''''                sFoundInList = sFoundInList & objQuery.Name & ","
''''                If bStopAtFirstFound = True Then GoTo Block_Exit
''''        End If
''''    Next
''''
''''Block_Exit:
''''    Set objQuery = Nothing
''''
''''End Function
''''
''''Private Function FindQueriesUsingObject(ByVal strObjName As String, Optional bLookInControls As Boolean = False, Optional bStopAtFirstFound As Boolean = False, Optional sFoundInList As String) As Integer
''''Dim oDb As CurrentData
''''Dim oQ As AccessObject
''''Dim oQDef As DAO.QueryDef
''''Dim strSql As String
''''
''''    Set oDb = Application.CurrentData
''''
''''    For Each oQ In oDb.AllQueries
'''''        Debug.Print oQ.name
''''        Set oQDef = CurrentDb.QueryDefs(oQ.Name)
'''''Debug.Assert oQ.name <> "Query2"
''''        strSql = oQDef.SQL
''''
''''        If InStr(1, strSql, strObjName, vbTextCompare) > 0 Then
''''            FindQueriesUsingObject = FindQueriesUsingObject + 1
'''''            LogMessage oQ.name & " contains: (" & strObjName & ")", strProcName
''''                sFoundInList = sFoundInList & oQ.Name & ","
''''                If bStopAtFirstFound = True Then GoTo Block_Exit
''''        End If
''''
'''''        Stop
''''    Next
''''
''''Block_Exit:
''''    Set oQ = Nothing
''''    Set oQDef = Nothing
''''    Set oDb = Nothing
''''
''''End Function
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Recursive function, looks at the tables and or saved queries used in all saved queries
''''''' for the given table name.. NOT FINISHED!
'''''''
''''Public Function ResolveRecursiveQueries(strEndTableNameSought As String, strQueryName As String) As Long
''''Dim objTable As TableDef
''''Dim objQuery As QueryDef
''''Dim objKD_PromptTable As DAO.RecordSet
''''Dim strKeyword As String
''''Dim strNameLike As String
''''Dim aqdfQueries() As String
''''Dim strSql As String
''''
''''    On Error Resume Next
''''
''''    Set objQuery = CurrentDb(strQueryName)
''''
''''    If objQuery Is Nothing Then
''''        Exit Function
''''    End If
''''
''''    strSql = objQuery.SQL
''''
''''    If InStr(1, strSql, strEndTableNameSought, vbTextCompare) > 0 Then
''''        ResolveRecursiveQueries = ResolveRecursiveQueries + 1
''''    Else
''''        aqdfQueries = GetTablesNQueriesUsedInQuery(strSql)
''''    End If
''''
''''
''''End Function
''''
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Attempts to extract any tables or saved queries from the passed SQL
'''''''
''''Public Function GetTablesNQueriesUsedInQuery(ByVal strSql As String)
''''Dim oRegEx As RegExp
''''Dim oMatches As VBScript_RegExp_55.MatchCollection
''''Dim oMatch As VBScript_RegExp_55.Match
''''
''''    Set oRegEx = New RegExp
''''    oRegEx.IgnoreCase = True
''''    oRegEx.Pattern = "FROM (.+)(ORDER|WHERE|HAVING)*"
''''
''''    Set oMatches = oRegEx.Execute(strSql)
''''    For Each oMatch In oMatches
''''        Debug.Print oMatch.Value
''''    Next
''''
''''    Set oRegEx = Nothing
''''
''''End Function
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''
''''
''''
'''''Public Sub LogMessage(strMessage As String)
'''''Dim lMsgFile As Long
'''''
'''''    lMsgFile = FreeFile
'''''
'''''Debug.Print "Log: " & strMessage
'''''
'''''    On Error Resume Next
'''''
'''''    Open CurrentDb.Name & "_LOG.txt" For Append Access Write Lock Write As #lMsgFile
'''''    Print #lMsgFile, strMessage
'''''    Close #lMsgFile
'''''
'''''End Sub
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''
''''
'''''
''''' Use this to store a binary file in the tbl_App_Dependencies table
'''''' Requires mod_Blobs
''''Public Function StoreDependency(strDependencyPath As String) As Boolean
''''On Error GoTo Funct_Err
''''Dim strProcName As String
''''Dim oDb As DAO.Database
''''Dim oRs As DAO.RecordSet
''''Dim sSql As String
''''Dim sDestPath As String
''''Dim sDepName As String
''''Dim oFso As Scripting.FileSystemObject
''''
''''
''''    strProcName = ClassName & ".StoreDependency"
''''
''''    Set oFso = New Scripting.FileSystemObject
''''    sDepName = oFso.GetFileName(strDependencyPath)
''''    Set oFso = Nothing
''''
''''    StoreDependency = True
''''''''    Debug.Assert 1 <> 1
''''
''''    Set oDb = CurrentDb()
''''    sSql = "SELECT * FROM tbl_App_Dependencies WHERE Active = True "
''''
''''    Set oRs = oDb.OpenRecordSet(sSql)
''''    oRs.AddNew
''''        oRs("DependencyName").Value = sDepName
''''        oRs("ModifyComputerName") = GetPCName
''''    oRs.Update
''''
''''    oRs.MoveLast
''''    ReadBLOB strDependencyPath, oRs, "DependencyOLE"
'''''
'''''
'''''
''''    If oRs.EOF And oRs.BOF Then
''''        LogMessage "No active dependencies found to extract", strProcName
''''        GoTo Funct_Exit
''''    End If
''''
''''    While Not oRs.EOF
''''        If IsNull(oRs("ExtractPath").Value) Or CStr("" & oRs("ExtractPath").Value) = "" Then
''''            sDestPath = CurrentDBDir() & "\" & oRs("DependencyName").Value
''''        Else
''''            sDestPath = oRs("ExtractPath").Value
''''        End If
''''
'''''        sDestPath = FixPath(sDestPath, MarketDate)
''''
''''        WriteBLOB oRs, "DependencyOLE", sDestPath
''''
''''        oRs.MoveNext
''''    Wend
'''''
'''''
''''Funct_Exit:
''''    Set oRs = Nothing
''''    Set oDb = Nothing
''''
''''    Exit Function
''''
''''Funct_Err:
''''    StoreDependency = False
''''    ReportError Err, strProcName
''''    Resume Funct_Exit
''''End Function
''''
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Debug.prints the databases properties
'''''''
''''Public Function ListDBProperties()
''''Dim oProp As DAO.Property
''''Dim oDb As DAO.Database
''''Dim sVal As String
''''
''''    On Error Resume Next
''''
''''    Set oDb = CurrentDb()
''''
''''    For Each oProp In oDb.Properties
''''        sVal = oProp.Value
''''        Debug.Print "Prop: " & oProp.Name & " = (" & sVal & ")"
''''    Next
''''
''''
''''End Function
''''
''''
''''
''''''' ############################################################
''''''' ############################################################
''''''' ############################################################
''''''' Finds and attempts to delete queries that won't prepare
''''''' USE WITH EXTREME CAUTION!!
'''''''
''''Public Function DeleteBadQueries()
''''Dim oDb As DAO.Database
''''Dim oQDef As DAO.QueryDef
''''On Error Resume Next
''''Dim cDelteCol As Collection
''''Dim iIndex As Integer
''''
''''
''''    Set oDb = CurrentDb()
''''    Set cDelteCol = New Collection
''''
''''    For Each oQDef In oDb.QueryDefs
''''
''''        Debug.Print oQDef.SQL
''''
''''        oQDef.Prepare = dbQPrepare
''''
''''        If Err.Number <> 0 Then
''''            Debug.Print "Delete this puppy: " & oQDef.Name
''''
''''            cDelteCol.Add oQDef.Name
''''        End If
''''        Err.Clear
''''
''''    Next
''''
''''    If cDelteCol.Count > 0 Then
''''        Stop    ' Hammer time!
''''        For iIndex = 0 To cDelteCol.Count
''''            oDb.QueryDefs.Delete (cDelteCol.Item(iIndex))
''''        Next
''''    End If
''''
''''End Function
''''
''''
''''''' ##############################################################################
''''''' ##############################################################################
''''''' ##############################################################################
'''''''
''''''' will go through linked tables and change the server to strDestinationSvr
''''''' Skips sIGNORESERVER
'''''''
''''''' BE CAREFUL USING THIS...
''''''' I tend to make quick hard coded changes in here
''''''' So, first thing - make a backup before you run this!
'''''''
''''Public Function ChangeLinkedTablesToDiffServer(Optional strDestinationSvr As String = "DC-BIGSKY") As Integer
''''On Error GoTo Block_Err
''''Dim strProcName As String
''''Dim oTbl As DAO.TableDef
''''Dim oRegEx As RegExp, oDbRegX As RegExp
''''Dim oRegX2 As RegExp
''''Dim sNewConnect As String
''''Dim cTblsNotFound As Collection
''''Dim sMsg As String
''''Dim sDbName As String
''''Dim i As Integer
''''Dim oRs As DAO.RecordSet
''''Dim sActiveServer As String
''''
''''Dim sCurDb As String
''''Dim sCurServer As String
''''Dim sCurApp As String
''''Dim sSql As String, oDb As DAO.Database
''''
''''Const sIGNORESERVER As String = "Claims.sql.ccaintranet.com"
''''
''''    Set oDb = CurrentDb()
''''
''''    strProcName = ClassName & ".ChangeLinkedTablesToDiffServer"
''''    Set oRegEx = New RegExp
'''''    oRegEx.Pattern = "SERVER=([^\=\;]+);"
''''    oRegEx.Pattern = "^.*?SERVER=([^\=\;]+);*.*?$"
''''    oRegEx.IgnoreCase = True
''''
''''    Set oDbRegX = New RegExp
'''''    oDbRegX.Pattern = "DATABASE=([^\=\;]+);"
''''    oDbRegX.Pattern = "^.*?DATABASE=([^\=\;]+);*.*?$"
''''    oDbRegX.IgnoreCase = True
''''
''''    Set oRegX2 = New RegExp
''''    oRegX2.Pattern = "^.*?APP=([^\=\;]+);*.*?$"
''''    oRegX2.IgnoreCase = True
''''
''''    Set cTblsNotFound = New Collection
''''
''''    For Each oTbl In CurrentDb.TableDefs
''''        If oTbl.Connect <> "" Then
''''            '' only do this if we need to:
''''
'''''            Debug.Assert left(oTbl.Name, 3) <> "dbo"
''''
''''
'''''            If oTbl.Name = "XrefAttachmentTypes" Then
'''''            If left(oTbl.Name, 3) = "dbo" Then
'''''                    oTbl.Connect = "DRIVER=SQL Server;SERVER=DS-FLD-009;Trusted_Connection=Yes;APP=AuditProbe;WSID=TS-CMS-DEV-001;DATABASE=CMS_AUDITORS_ERAC"
'''''                    oTbl.RefreshLink
''''''                Stop
'''''            End If
''''
''''
''''            If InStr(1, oTbl.Connect, sIGNORESERVER, vbTextCompare) > 0 Then
''''                Debug.Print "Skipping table: " & oTbl.Name & " : " & sIGNORESERVER
''''            ElseIf InStr(1, oTbl.Connect, "SERVER=" & strDestinationSvr, vbTextCompare) < 1 Then
''''
''''                Debug.Print oTbl.Name & ": " & oTbl.Connect
''''
''''            sCurDb = oDbRegX.Replace(oTbl.Connect, "$1")
''''            sCurServer = oRegEx.Replace(oTbl.Connect, "$1")
''''            sCurApp = oRegX2.Replace(oTbl.Connect, "$1")
'''''Stop
''''            sNewConnect = "DRIVER=SQL Server;SERVER=" & sCurServer & ";Trusted_Connection=Yes;APP=McrClaimAdmin;WSID=TS-CMS-DEV-001;DATABASE=" & sCurDb
''''
''''
''''            sSql = "UPDATE Link_Table_Config SET CurLinked = -1 WHERE Table = '" & oTbl.Name & "' AND Database = '" & sCurDb & "'"
''''            Debug.Print sSql
''''
'''''Stop
''''            oDb.Execute sSql
''''
'''''                sDbName = oDbRegX.Replace(oTbl.Connect, "$1")
'''''
'''''                sNewConnect = oRegEx.Replace(oTbl.Connect, "SERVER=" & strDestinationSvr & ";")
'''''                If sNewConnect <> "" Then
''''                    oTbl.Connect = sNewConnect
''''                    oTbl.RefreshLink
'''''                Else
'''''                    'Stop
'''''
'''''                End If
''''TableErrResume:
''''            End If
''''
''''
''''        End If
''''    Next
''''
''''        '' update the table to show users that we are pointing to the correct server
''''    Set oRs = CurrentDb().OpenRecordSet("SELECT * FROM Link_Table_Location")
''''
''''    Select Case strDestinationSvr
''''    Case "DEV-SQL-002"
''''        sActiveServer = "MCRDEV"
''''    Case "DBPRDSQL-004"
''''        sActiveServer = "MCRPROD"
''''    End Select
''''
''''
''''    While Not oRs.EOF
''''        If oRs("LocationID") = sActiveServer Then
''''            oRs.Edit
''''            oRs("Active").Value = 1
''''            oRs.Update
''''        Else
''''            oRs.Edit
''''            oRs("Active").Value = 0
''''            oRs.Update
''''        End If
''''        oRs.MoveNext
''''    Wend
''''
''''
''''
''''    If cTblsNotFound.Count > 0 Then
''''        sMsg = "Tables not found:" & vbCrLf
''''        For i = 1 To cTblsNotFound.Count
''''            sMsg = sMsg & cTblsNotFound.Item(i) & vbCrLf
''''        Next
''''
''''        Debug.Print vbCrLf & vbCrLf & vbCrLf
''''        Debug.Print sMsg
''''
''''        MsgBox sMsg
''''
''''    End If
''''
''''Block_Exit:
''''    Set oRs = Nothing
''''    Exit Function
''''Block_Err:
''''    If Not oTbl Is Nothing Then
''''
''''        cTblsNotFound.Add sDbName & ".dbo." & oTbl.Name
''''        Err.Clear
''''        GoTo TableErrResume
''''    End If
''''
''''    Err.Clear
''''    Resume Next
''''End Function
''''
'''''
'''''Public Function TestDlg()
'''''Dim dlg As clsDialogs
'''''Set dlg = New clsDialogs
'''''Dim sFilePath As String
'''''
'''''
'''''    With dlg
'''''
'''''        sFilePath = .OpenPath("C:\", WordDox, , "Pick your word document please!")
'''''
'''''        If sFilePath = "" Then
'''''            MsgBox "Canceled?"
'''''        Else
'''''            MsgBox "Picked: " & sFilePath
'''''        End If
'''''
'''''    End With
'''''
'''''End Function
''''
''''
''''Public Sub TurnOffDeveloperErrorHandling(bTurnOff As Boolean)
''''Static iCurSetting As Integer
''''Static bDefaultCaptured As Boolean
''''
''''
''''    If bTurnOff = True Then
''''        If bDefaultCaptured = False Then
''''            iCurSetting = Application.GetOption("Error Trapping")
''''            bDefaultCaptured = True
''''        End If
''''        Application.SetOption "Error Trapping", 2
''''
''''    Else
''''        Application.SetOption "Error Trapping", iCurSetting
''''    End If
''''
''''End Sub
''''
'''''' I used this to deploy the shortcuts to everyone's folder
'''''' Just need to make sure that this is part of the deployment
'''''' so new users will have their folders set up with the shortcut.
''''Public Function DeployShortcuts()
''''Dim oAdo As clsADO
''''Dim oRs As ADODB.RecordSet
''''Dim oFso As Scripting.FileSystemObject
''''Dim sSql As String
''''Const sRootFldr As String = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\AUDITOR_FOLDERS\CLAIM_ADMIN\"
''''Const sDbName As String = "CLAIMS ADMIN 2010.accde"
''''Const slDbName As String = "CLAIMS ADMIN 2010.laccde"
''''Const sShortcutPath As String = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\AUDITOR_FOLDERS\CLAIM_ADMIN\_Production_copy\_Launch_Claim_Admin"
''''Const sARMSShortcutPath As String = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\ARMS\_PRODUCTION\Open_ARMS"
''''Dim sUserFldr As String
'''''Const sShortcutName As String = "_Launch_Claim_Admin_2010"
''''
'''''    ' right now we aren't going to do this in MCR
'''''Exit Function
''''
''''    sSql = "SELECT UA.UserID FROM ADMIN_User_Account UA " & _
''''        " INNER Join ADMIN_User_Profile UP ON UA.UserID = UP.UserID " & _
''''        " WHERE SUBSTRING(UP.ProfileID,1,3) <> 'Sub'"
''''
''''
''''    Set oAdo = New clsADO
''''    With oAdo
''''        .ConnectionString = GetConnectString("V_DATA_DATABASE")
''''        .SQLTextType = sqltext
''''        .sqlString = sSql
''''        Set oRs = .ExecuteRS
''''    End With
''''
''''    Set oFso = New Scripting.FileSystemObject
''''
''''    While Not oRs.EOF
''''        ' If there isn't a dot in the userid, then skip it
''''        If InStr(1, CStr("" & oRs("UserID").Value), ".") > 0 Then
''''            sUserFldr = sRootFldr & CStr("" & oRs("UserId").Value) & "\"
''''
''''            If oFso.FolderExists(sRootFldr & CStr("" & oRs("UserID").Value)) = False Then
''''                ' create the folder
''''                CreateFolders (sRootFldr & CStr("" & oRs("UserId").Value))
''''            End If
''''
''''            ' Copy the shortcut:
''''            Call CopyFile(sShortcutPath & ".lnk", sUserFldr, False)
''''
''''                ' Give them the arms shortcut too..
''''            Call CopyFile(sARMSShortcutPath & ".lnk", sUserFldr, False)
''''
''''            ' Now, if the mde is there but isn't locked, delete it:
''''            If FileExists(sUserFldr & slDbName) = False Then
''''                If FileExists(sUserFldr & sDbName) = True Then
''''                    DeleteFile sUserFldr & sDbName, False
''''                End If
''''            Else
''''                LogMessage "DeployShortcuts", "LOCKED DB", sUserFldr & slDbName, oRs("UserID").Value & "@connolly.com;"
''''            End If
''''
''''        End If
''''
''''        oRs.MoveNext
''''    Wend
''''
''''End Function
''''
''''
''''
''''Public Sub SubdatasheetFix()
''''Dim oDb As DAO.Database
''''Dim oProp As DAO.Property
''''Dim sPropName As String
''''Dim sPropVal As String
''''Dim sReplaceVal As String
''''Dim iPropType As Integer
''''Dim i As Integer
''''Dim iCount As Integer
''''Dim strProcName As String
''''On Error GoTo Block_Err
''''
''''
''''    strProcName = ClassName & ".SubdatasheetFix"
''''
''''
''''    Set oDb = CurrentDb
''''    sPropName = "SubDataSheetName"
''''    iPropType = 10
''''    sPropVal = "[None]"
''''    sReplaceVal = "[Auto]"
''''    iCount = 0
''''
''''    For i = 0 To oDb.TableDefs.Count - 1
''''        If (oDb.TableDefs(i).Attributes And dbSystemObject) = 0 Then
''''            If oDb.TableDefs(i).Properties(sPropName).Value = sReplaceVal Then
''''                oDb.TableDefs(i).Properties(sPropName).Value = sPropVal
''''                iCount = iCount + 1
''''            End If
''''        End If
''''ReturnFromErrHandler:
''''    Next
''''
''''Block_Exit:
''''    Exit Sub
''''Block_Err:
''''    If Err.Number = 3270 Then
''''        Set oProp = oDb.TableDefs(i).CreateProperty(sPropName)
''''        oProp.Type = iPropType
''''        oProp.Value = sPropVal
''''        oDb.TableDefs(i).Properties.Append oProp
''''        iCount = iCount + 1
''''        Resume ReturnFromErrHandler
''''    Else
''''        MsgBox Err.Description & vbCrLf & vbCrLf & " in " & strProcName & " routine."
''''        GoTo Block_Exit
''''    End If
''''
''''End Sub
''''
''''
''''
''''
''''
''''Public Function RelinkTables() As Boolean
''''On Error GoTo Block_Err
''''Dim strProcName As String
''''Dim oRs As ADODB.RecordSet
''''
''''    strProcName = ClassName & ".RelinkTables"
''''
''''    LogMessage strProcName, , "Relinking tables.. " & CurrentDb.Name
''''
''''    If LCase(GetUserName) <> "kevin.dearing" Then
''''        TurnOffDeveloperErrorHandling True
''''    Else
''''        Application.visible = True
''''
''''    End If
''''
''''            '    Call UnLinkTables
''''
''''    CurrentDb.Execute ("Update Link_Table_Location SET Active = 0 ")
''''
''''    Set oRs = New ADODB.RecordSet
''''    oRs.ActiveConnection = CurrentProject.Connection
''''
''''    oRs.Open ("Select * from Link_Table_Config WHERE Location = 'MCRPROD' AND Server = 'DBPRDSQL-004' AND Database Like 'MCR_AUDITORS%' AND Database <> 'MCR_AUDITORS_CONFIG' ")
''''
''''    While Not oRs.EOF
''''        With oRs
'''''            If chkDB Then
'''''                LinkTable "SQL", ![Server], ![Database]
'''''            Else
''''LogMessage strProcName, , "Relinking table: " & oRs("Table").Value
''''
''''                Call UnLinkTables(oRs("Table").Value)
''''
''''                CurrentDb.TableDefs.Refresh
''''
''''                LinkTable "SQL", ![Server], ![Database], ![Table]
'''''            End If
''''            .MoveNext
''''        End With
''''    Wend
''''
''''    oRs.Close
''''    Set oRs = Nothing
''''
''''    DoCmd.Hourglass False
''''
''''
''''
''''
''''Block_Exit:
''''    CurrentDb.Execute ("Update Link_Table_Location SET Active = 1 WHERE LocationID = 'MCRPROD'")
''''
''''    TurnOffDeveloperErrorHandling False
''''    Exit Function
''''Block_Err:
''''    If LCase(GetUserName) = "kevin.dearing" Then
''''        Application.visible = True
''''        Stop
''''    End If
''''    ReportError Err, strProcName
''''    GoTo Block_Exit
''''
''''End Function
''''
''''
''''
''''Public Function BackFill_Concept_LCD()
''''On Error GoTo Block_Err
''''Dim strProcName As String
''''Dim oAdo As clsADO
''''Dim oRs As ADODB.RecordSet
''''Dim oRegEx As RegExp
''''
''''    strProcName = ClassName & ".BackFill_Concept_LCD"
''''
''''    '' need a regex to:
''''    ' find all instances of the LCD in the code
''''    ' some samples are like:
''''    Set oRegEx = New RegExp
''''    With oRegEx
''''        .IgnoreCase = True
''''        .Global = True
''''        .MultiLine = False
''''        .Pattern = "LCD[\s\t ]+\(*[lL]*([0-9]+?)[\s\t \r\n\)]"
''''    End With
''''
''''    '' We have to go through both tables: _hdr and _Payer_dtl
''''    ''
''''    Set oAdo = New clsADO
''''    With oAdo
''''        .ConnectionString = GetConnectString("v_Data_Database")
''''        .SQLTextType = sqltext
''''        .sqlString = "SELECT PayerNameID = 1000, ConceptId, ConceptReferences FROM CONCEPT_Hdr WHERE ConceptReferences LIKE '%LCD%' AND ConceptStatus NOT IN ('990','995','350')"
''''        Set oRs = .ExecuteRS
''''        If .GotData = False Then
''''            Stop
''''        End If
''''    End With
''''
''''Dim sConceptRef As String
''''Dim sThisConcept As String
''''Dim lPayerNameId As Long
''''Dim oMatches As MatchCollection
''''Dim oMatch As Match
''''
''''
''''    While Not oRs.EOF
''''        '
''''        sThisConcept = oRs("ConceptID").Value
''''        lPayerNameId = oRs("PayerNameID").Value
''''        sConceptRef = oRs("ConceptReferences").Value
''''
''''        Set oMatches = oRegEx.Execute(sConceptRef)
''''
''''        For Each oMatch In oMatches
'''''            Debug.Print oMatch
'''''                        Stop
'''''            Debug.Print oMatch.SubMatches(0)
''''            If IsNumeric(oMatch.SubMatches(0)) Then
''''                Call PopLCD(sThisConcept, CLng(oMatch.SubMatches(0)), lPayerNameId)
''''                Debug.Print sThisConcept & ", " & oMatch.SubMatches(0) & ", " & CStr(lPayerNameId)
''''            End If
''''        Next
''''
''''
''''        oRs.MoveNext
''''    Wend
''''
''''    '' Now, the payer detail table:
''''
''''    oRs.Close
''''    Set oRs = Nothing
''''
''''    Set oAdo = New clsADO
''''    With oAdo
''''        .ConnectionString = GetConnectString("v_Data_Database")
''''        .SQLTextType = sqltext
''''        .sqlString = "SELECT PayerNameID, ConceptId, ConceptReferences FROM CONCEPT_Payer_Dtl WHERE ConceptReferences LIKE '%LCD%' AND ConceptStatus NOT IN ('990','995','350')"
''''        Set oRs = .ExecuteRS
''''        If .GotData = False Then
''''            Stop
''''        End If
''''    End With
''''
''''
''''    While Not oRs.EOF
''''        '
''''        sThisConcept = oRs("ConceptID").Value
''''        lPayerNameId = oRs("PayerNameID").Value
''''        sConceptRef = oRs("ConceptReferences").Value
''''
''''        Set oMatches = oRegEx.Execute(sConceptRef)
''''
''''        For Each oMatch In oMatches
'''''            Debug.Print oMatch
'''''                        Stop
'''''            Debug.Print oMatch.SubMatches(0)
''''            If IsNumeric(oMatch.SubMatches(0)) Then
''''                Call PopLCD(sThisConcept, CLng(oMatch.SubMatches(0)), lPayerNameId)
''''                Debug.Print sThisConcept & ", " & oMatch.SubMatches(0) & ", " & CStr(lPayerNameId)
''''            End If
''''        Next
''''
''''
''''        oRs.MoveNext
''''    Wend
''''
''''
''''Block_Exit:
''''    Exit Function
''''Block_Err:
''''    ReportError Err, strProcName
''''    GoTo Block_Exit
''''End Function
''''
''''
''''Private Function PopLCD(sConceptId As String, lLcdId As Long, lPayerNameId As Long)
''''Dim oAdo As clsADO
''''
''''    Set oAdo = New clsADO
''''    With oAdo
''''        .ConnectionString = GetConnectString("v_Code_Database")
''''        .SQLTextType = StoredProc
''''        .sqlString = "usp_CONMGNT_Add_LCD_To_Concept"
''''        .Parameters.Refresh
''''        .Parameters("@pConceptId") = sConceptId
''''        .Parameters("@pPayerNameId") = lPayerNameId
''''        .Parameters("@pLCD") = lLcdId
''''        .Execute
''''        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
''''            LogMessage strProcName, "BACKFILL WARNING", "Problem adding an LCD: " & CStr(lLcdId) & " Concept: " & sConceptId & " Payer: " & CStr(lPayerNameId), .Parameters("@pErrMsg").Value, False, sConceptId
''''        End If
''''
''''    End With
''''    Set oAdo = Nothing
''''End Function
''''
''''
'''''Public Function PrinterTest()
'''''Dim oPtr As clsPrinter
'''''Dim lPrinter As Long
'''''
'''''    Set oPtr = New clsPrinter
'''''
''''''    lPrinter = oPtr.TestExample
'''''    oPtr.SelectPrinter
'''''
'''''    'Debug.Print oPtr.PrinterStatus()
'''''
'''''
'''''
'''''End Function
''''
'''''''
'''''''Sub SelectPrinter()
''''''''   Dim sPrinter As String
''''''''   sPrinter = Application.ActivePrinter
''''''''   Application.Dialogs(xlDialogPrinterSetup).show
''''''''   ActiveSheet.PrintPreview
''''''''   Application.ActivePrinter = sPrinter
'''''''End Sub
''''
''''
''''
''''
''''
''''Public Function FileSpecs()
''''Dim oDb As DAO.Database
''''Dim oPrj As CurrentProject
''''Dim oVar As Variant
''''
''''    Debug.Print "Current " & TypeName(Application.CurrentProject)
''''
''''    Set oPrj = Application.CurrentProject
''''
''''Debug.Print oPrj.ImportExportSpecifications.Count
''''
''''    For Each oVar In oPrj.ImportExportSpecifications
''''        Debug.Print oVar.Name
''''        Stop
''''    Next
''''
''''End Function
''''
''''
''''
Public Function TestConn()
Dim oCn As ADODB.Connection
Dim oRs As ADODB.RecordSet
Dim oFld As ADODB.Field

    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=DS-FLD-009;Initial Catalog=CMS_AUDITORS_CLAIMS;"
        .CursorLocation = adUseNone
        .Open
    End With

    Set oRs = New ADODB.RecordSet
    oRs.Open "SELECT TOP 100 * FROM AuditClm_hdr", oCn, adOpenStatic, adLockReadOnly
    
    While Not oRs.EOF
        For Each oFld In oRs.Fields
            Debug.Print oFld.Name & " = " & oRs(oFld.Name).Value
        Next
        oRs.MoveNext
    Wend
    
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If

End Function
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'THIS CLASS IS SEALED
'it should not be modified or methods added without the knowledge of the
'Development department. Its methods are implemented based in specific requirements.
'Modifying any part of this class might disable add-ins and generate unexpected behavior in Decipher

'This class manages all the related tasks with the management of add-ins for Decipher


Private Const AddInVersionTableName = "CT_ADDINRANGEVERSION"

'Gets the name of the Ribbon Designer add-in that is used to call the add-in
Private Const RibbonDesignerAddIn As String = "DecipherRibbonDesigner.Connect"
'Gets the name of the Function Name that will be called in the Ribbon Designer add-in
Private Const RibbonAddInFunctionName As String = "RibbonDesigner"

'Determines if a progId is part of the available add-ins for the current Decipher version
Public Function CanAddinBeLoaded(progId As String, Optional isDcUser As Boolean = True) As Boolean
    On Error GoTo CanAddinBeLoadedError
    Dim retval As Boolean
    Dim SQL As String
    Dim addinName As String
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
        
    SQL = GetLoadAddinsSql(isDcUser)
    
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot)
    
    retval = False
    If Not (rs.EOF And rs.BOF) Then
       Do Until rs.EOF
            addinName = Nz(rs.Fields("AddInProgIdName").Value, "")
            
            If UCase(addinName) = UCase(progId) Then
                'add in found no need to keep checking
                retval = True
                GoTo CanAddinBeLoadedExit
            End If
            rs.MoveNext
       Loop
    End If
    
   
CanAddinBeLoadedExit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    CanAddinBeLoaded = retval
    Exit Function
CanAddinBeLoadedError:
    'leave for testing
    Debug.Print Err.Number & "  " & Err.Description
    Debug.Print SQL
    Resume CanAddinBeLoadedExit
End Function

'Gets the Decipher add-in version from the properties/description CT_AddinRangeVersion table
Public Function GetDecipherAddinVersion(dbName As String) As Double
    On Error GoTo GetDecipherAddinVersionError
    Dim db As DAO.Database
    Dim Tbl As DAO.TableDef
    Dim Version As Double
    Version = -1
    Set db = DBEngine.OpenDatabase(dbName)
    
    For Each Tbl In db.TableDefs
        If UCase(Tbl.Name) = UCase(AddInVersionTableName) Then
            Version = Tbl.Properties("Description")
            Exit For
        End If
    Next Tbl
    
   
GetDecipherAddinVersionExit:
    On Error Resume Next
    Set Tbl = Nothing
    Set db = Nothing
    GetDecipherAddinVersion = Version
    Exit Function

GetDecipherAddinVersionError:
    'leave for testing
    'Debug.Print err.Number & "  " & err.Description
    Resume GetDecipherAddinVersionExit

End Function

'loads all Add-ins register in the table CT_AddinDecipherVersion that are enabled if the parameter isDcUser is set to false
'then gets all addins that are marked in the Load column as true
Public Sub LoadAddins(Optional isDcUser As Boolean = True)
    On Error GoTo LoadAddinsError
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim SQL As String
    Dim sqlUpdate As String
    Dim addinName As String
    Dim addinID As Integer
    
    SQL = GetLoadAddinsSql(isDcUser)
  
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot)
    
    If Not (rs.EOF And rs.BOF) Then
        Do Until rs.EOF
            addinName = Nz(rs.Fields("AddInProgIdName").Value, "")
            addinID = rs.Fields("AddinID").Value
                   
            'load addin
            If (Not LoadAddin(addinName, True)) Then
               'if not loaded disable add-in in the table.
                sqlUpdate = "UPDATE CT_AddinDecipherVersion SET Disabled = true,  DisabledDate = #" & Now() & "# WHERE addinID=" & addinID
                CurrentDb.Execute (sqlUpdate)
            End If
            
            rs.MoveNext
        Loop
    End If
    
LoadAddinsExit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
LoadAddinsError:
    'leave for testing
    'Debug.Print err.Number & "  " & err.Description
    Resume LoadAddinsExit
End Sub

'Unloads all Add-ins register in the table CT_AddinDecipherVersion that are enabled if the parameter isDcUser is set to false
'then gets all addins that are marked in the Load column as true
Public Sub UnloadAddins(Optional isDcUser As Boolean = True)
    On Error GoTo UnloadAddinsError
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim SQL As String
    Dim addinName As String
    Dim addinID As Integer
    
    SQL = GetLoadAddinsSql(isDcUser)
    
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot)
    
    If Not (rs.EOF And rs.BOF) Then
        Do Until rs.EOF
            addinName = Nz(rs.Fields("AddInProgIdName").Value, "")
            addinID = rs.Fields("AddinID").Value
                   
            'unload addin
            LoadAddin addinName, False
            rs.MoveNext
        Loop
    End If
    
UnloadAddinsExit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
UnloadAddinsError:
    'leave for testing
    'Debug.Print err.Number & "  " & err.Description
    Resume UnloadAddinsExit

End Sub

'This function enables the add-in locally. it updates the 'enabled' so the add-in can be loaded at start up.
Public Function EnabledAddinLocally(addinName As String) As Boolean
    On Error GoTo EnabledAddinLocallyError
    Dim sqlUpdate As String
    Dim Result As Boolean
    Result = False
    
    sqlUpdate = "UPDATE CT_AddinDecipherVersion SET Disabled = false WHERE AddInProgIdName='" & addinName & "'"
    CurrentDb.Execute (sqlUpdate)
    Result = True

EnabledAddinLocallyExit:

    EnabledAddinLocally = Result
    Exit Function
EnabledAddinLocallyError:
    'Debug.Print err.Number & " " & err.Description
    Debug.Print sqlUpdate
    Resume EnabledAddinLocallyExit
End Function

'Determines the status of the add-in this method does not check if the user is allowed (isDCUser)to load or not the add-in
Public Function AddinStatus(addinName As String) As CnlyAddinStatus
    On Error GoTo AddinStatusError
    Dim retval As CnlyAddinStatus
    retval = CnlyAddinStatus.None
     With COMAddIns(addinName)
        If .Connect = True Then
           retval = Loaded
        Else
            retval = Unloaded
            If (IsAddinDisabledLocal(addinName)) Then
                retval = DisabledLocal
            End If
        End If
    End With
    
AddinStatusExit:
    AddinStatus = retval
    Exit Function
AddinStatusError:
    On Error Resume Next
       retval = NotFound
    Resume AddinStatusExit
End Function

'this methods determines if a particular add-in has been disabled in the table CT_AddinDecipherVersion
'add-in name should be unique but if two or more add-in names (progID)are found and one is disabled it returns true for all
'if the add-in name is not found it returns false
Public Function IsAddinDisabledLocal(addinName As String) As Boolean
    On Error GoTo IsAddinDisabledLocalError
    Dim retval As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim SQL As String
    retval = True
    SQL = "SELECT AddInID, AddinProgIdName From CT_AddinDecipherVersion WHERE Disabled = true AND AddinProgIdName = '" & addinName & "';"
    
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot)
    
    'check if records were returned that match the addinName, if records are returned add-in is disabled
    If rs.EOF And rs.BOF Then
        'no records returned add-in is not disabled
        retval = False
    End If
  
IsAddinDisabledLocalExit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    IsAddinDisabledLocal = retval
    Exit Function
IsAddinDisabledLocalError:
    retval = True
    Resume IsAddinDisabledLocalExit
End Function


'loads an unloads an add-in depending on the loadTheAddin parameter
'if true loads the addin otherwise it unloads the addin
'returns true if successfull otherwise false
Private Function LoadAddin(progId As String, loadTheAddin As Boolean) As Boolean
    On Error GoTo LoadAddinError
    Dim retval As Boolean
    Dim isConnect As Boolean
    isConnect = Not loadTheAddin
    retval = False
       
    With COMAddIns(progId)
        If .Connect = isConnect Then
            'if not connected connect
            .Connect = loadTheAddin
         
        End If
        retval = True
        
    End With
        
LoadAddinExit:
    On Error Resume Next
    LoadAddin = retval
    Exit Function
LoadAddinError:
    On Error Resume Next
    Dim sqlUpdate As String
     sqlUpdate = "UPDATE CT_AddinDecipherVersion SET DisabledError = '" & Err.Number & " - " & Err.Description & " - Connect value=" & CStr(loadTheAddin) & "' WHERE AddinProgIdName=" & "'" & progId & "'"
     Debug.Print sqlUpdate
     CurrentDb.Execute (sqlUpdate)
    'leave for testing
    'Debug.Print err.Number & "  " & err.Description
    Resume LoadAddinExit
        
End Function


'Gets the sql needed to get all the add-ins available in the table
'isDcUser determines if all enabled valid addi-ins will be retrieved
'isDcUser = true all enabled valid add-ins are retrieved
'isDcUser = false all enabled valid add-ins where DcUser column is true
Private Function GetLoadAddinsSql(Optional isDcUser As Boolean = True) As String
    Dim SQL As String
    Dim currentVersion As Double
    currentVersion = GetDecipherAddinVersion(CurrentDb.Name)
        
    SQL = " Select DA.AddinID, DA.AddInVersion, DA.AddInProgIdName "
    SQL = SQL & " From CT_AddinRangeVersion As DV INNER Join "
    SQL = SQL & "CT_AddinDecipherVersion As DA ON DV.VerID = DA.VerID "
    SQL = SQL & "WHERE " & currentVersion & " BETWEEN DV.MinDecVer AND DV.MaxDecVer "
    SQL = SQL & "AND DA.Disabled = false "
    If isDcUser = False Then
        SQL = SQL & "AND DA.DcUser = true "
    End If
    SQL = SQL & "ORDER BY DA.AddInProgIdName;"
    
    GetLoadAddinsSql = SQL
End Function

'It runs the Ribbon Designer add-in and gives the appropriate information to the user based in the status (loaded, unloaded, disabled) of the add-in
'progId is the name of the add-in to run
Private Function RunRibbonrAddIn(ByVal progId As String) As Boolean
  On Error GoTo RunRibbonrAddInError:
    
    Dim Status As CnlyAddinStatus
    Dim sucess As Boolean
    sucess = False
  
     'determines if the user can access this particular add-in
    Status = AddinStatus(progId)
    
    Select Case (Status)
        Case CnlyAddinStatus.Loaded
            If CanAddinBeLoaded(progId) Then
                'if loaded it can be use it
                sucess = True
                
            Else
                MsgBox "You do not have sufficient permissions to use the Ribbon Designer add-in"
            End If
       Case CnlyAddinStatus.Unloaded
            MsgBox "The Ribbon Designer add-in is not currently loaded. You might not have sufficient rights to use this add-in. Contact Tech Support for more information"
            
       Case CnlyAddinStatus.DisabledLocal
            Dim Result As VbMsgBoxResult
            Result = MsgBox("The Ribbon Designer add-in was unable to load and has been disabled in the current Database. " _
                       & vbNewLine & vbNewLine & "Would you like to try and fix this problem automatically?" _
                       & vbNewLine & "(If this error continues please contact Tech Support for assistance.)", vbYesNoCancel + vbCritical, "Add-in Disabled")
            If Result = vbYes Then
                If EnabledAddinLocally(progId) Then
                    MsgBox "You must close and reopen the current database for the specified option to take effect.", vbInformation
                Else
                    MsgBox "There was an error trying to update the table CT_AddinDecipherVersion." & vbNewLine & "Close this message and try again. ", vbCritical
                End If
                
            End If
            
       Case CnlyAddinStatus.NotFound
            MsgBox "The Ribbon Designer add-in was not found in Access. Make sure the add-in is installed"
    End Select
    

RunRibbonrAddInExit:
    RunRibbonrAddIn = sucess
     Exit Function
RunRibbonrAddInError:
    'Debug.Print err.Number & " " & err.Description 'leave for testing
    MsgBox Err.Description
    Resume RunRibbonrAddInExit
End Function

'It executes the Execute method of the MS Access add-in.
'progId is the name of the add-in to run
'functionName the function to run within the add-in
'specialFunction is any special instruction for the function to run
Private Sub ExecuteAddIn(ByVal progId As String, ByVal FunctionName As String, Optional ByVal specialFunction As String = "")
    On Error GoTo ExecuteAddInError:
    COMAddIns(progId).Object.GetFunction(FunctionName).Execute (specialFunction)
         
ExecuteAddInExit:
     
         Exit Sub
ExecuteAddInError:
        
        MsgBox Err.Description
        Resume ExecuteAddInExit

End Sub

'It executes the Run function of the MS Access add-in. It returns true when successfully completed otherwise false.
'progId is the name of the add-in to run
'functionName the function to run within the add-in
'specialFunction is any special instruction for the function to run
Private Function RunAddIn(ByVal progId As String, ByVal FunctionName As String, Optional ByVal specialFunction As String = "") As Boolean
    On Error GoTo RunAddInError:
    Dim success As Boolean
    success = False
    success = COMAddIns(progId).Object.GetFunction(FunctionName).Execute(specialFunction)
         
RunAddInExit:
        RunAddIn = success
         Exit Function
RunAddInError:
        
        MsgBox Err.Description
        Resume RunAddInExit

End Function

'Gets an array with messages from the MS Access add-in
'progId is the name of the add-in to run
'functionName the function to run within the add-in
'specialFunction is any special instruction for the function to run
Private Function GetMessageAddIn(ByVal progId As String, ByVal FunctionName As String, Optional ByVal specialFunction As String = "") As String()
    On Error GoTo GetMessageAddInError:
    Dim Messages() As String
    Messages = COMAddIns(progId).Object.GetFunction(FunctionName).GetMessages
         
GetMessageAddInExit:
        GetMessageAddIn = Messages
         Exit Function
GetMessageAddInError:
        
        MsgBox Err.Description
        Resume GetMessageAddInExit

End Function

'It rebuilds the ribbon bar and display the appropiate messages (error or confirmation of complete) when the operation is completed
'It does not provede error handling since this is done in the RunRibbonrAddIn method.
Public Sub BuildRibbonBar(Optional ByVal buildType As String = "Build")
    If RunRibbonrAddIn(RibbonDesignerAddIn) Then
        ExecuteAddIn RibbonDesignerAddIn, RibbonAddInFunctionName, buildType
    End If
End Sub

'It opens the Ribbon Designer Form. This Form is part of the Decipher Ribbon Designer add-in.
'It does not provede error handling since this is done in the RunRibbonrAddIn method.
Public Sub OpenRibbonDesigner()
    If RunRibbonrAddIn(RibbonDesignerAddIn) Then
        ExecuteAddIn RibbonDesignerAddIn, RibbonAddInFunctionName
    End If
End Sub

'It Runs the specified type of process in the ribbon bar designer
'if successfuly executed it return true otherwise false.
'updateType the process name to execute.
Public Function RunRibbonUpdate(updateType As String) As Boolean
   If RunRibbonrAddIn(RibbonDesignerAddIn) Then
        RunRibbonUpdate = RunAddIn(RibbonDesignerAddIn, RibbonAddInFunctionName, updateType)
   End If

End Function

'Gets an array with messages from the MS Access add-in
'theType the type of messages to retrieve optional.
Public Function GetMessages(Optional ByVal theType As String = "") As String()
   If RunRibbonrAddIn(RibbonDesignerAddIn) Then
        GetMessages = GetMessageAddIn(RibbonDesignerAddIn, RibbonAddInFunctionName, theType)
   End If
End Function
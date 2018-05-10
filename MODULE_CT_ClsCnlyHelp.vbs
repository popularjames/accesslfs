Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Event StatusMessage(ByVal Src As String, ByVal Status As String, ByVal Msg As String)
Private WithEvents ClsGenUtil As CT_ClsGeneralUtilities
Attribute ClsGenUtil.VB_VarHelpID = -1


'To flag whether the selected prefix should be included or excluded for Export/Update
Private PrefixInclude As Boolean

' VT 04/21/2009 Declaring private property variable for HelpFile
Private MvHelpFile As String
Private MvHelpPath As String
Private MvHelpTable As String
Private MvAppName As String
Private MvDecipherVer As String

'Application Name
Public Property Let AppName(data As String)
    MvAppName = data
End Property
Public Property Get AppName() As String
    AppName = MvAppName
End Property

'Help file property
Public Property Let helpFile(data As String)
    MvHelpFile = data
End Property
Public Property Get helpFile() As String
    helpFile = MvHelpFile
End Property
'Help file path property
Public Property Let helpPath(data As String)
    MvHelpPath = data
End Property
Public Property Get helpPath() As String
    helpPath = MvHelpPath
End Property
'Help table name property
Public Property Let HelpTable(data As String)
    MvHelpTable = data
End Property
Public Property Get HelpTable() As String
    HelpTable = MvHelpTable
End Property

Public Property Get DecipherVersion() As String
   
   'Find out Decipher version and set the value to private variable "MvDecipherVer".
    If Nz(MvDecipherVer, "") = "" Or MvDecipherVer = "Nill" Then
        GetDecipherVersion
    End If
    
    DecipherVersion = MvDecipherVer
End Property


Public Function SetMapIds() As Boolean
On Error GoTo ErrorHandler
    Set ClsGenUtil = New CT_ClsGeneralUtilities

    ' Check to see if table even exist before continuing.
    If Not ObjectExists(objTable, MvHelpTable) Then
        RaiseEvent StatusMessage("Mapping", "Error", "Table does not exit: " & MvHelpTable)
        SetMapIds = False
        Exit Function
    End If
    
    'Compare table's version with Decipher version. If the versions don't match do not continue.
    If Not (CompareTblVersion) Then
        SetMapIds = False
        Exit Function
    End If
    
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim frm As Form, ctl As Control
    Dim SQL, curFrm, prevFrm, ctlName As String
    Dim Prefixes() As String
    Dim counter As Integer
    
    Prefixes = GetAppPrefix
    
    ' Building select query that only retrieves the desired form objects' controls.
    SQL = "Select * from " & MvHelpTable & " Where "
       
    If (PrefixInclude = True) Then
        SQL = SQL & "ObjectName Like """ & Prefixes(0) & "*"" "
    Else
        counter = 0
        While counter <= UBound(Prefixes)
            If counter <> 0 Then
                SQL = SQL & " OR "
            End If
            
            SQL = SQL & "Not (ObjectName Like """ & Prefixes(counter) & "*"" ) "
            counter = counter + 1
        Wend
    End If
    
    SQL = SQL & " order by ObjectName, ControlName"
    curFrm = ""
    prevFrm = ""
        
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot + dbReadOnly)
      
    RaiseEvent StatusMessage("Mapping ", "Running", "Mapping Controls...")
               
    While Not rs.EOF
        
    ' VT 04/22/2009 Skipping the current cnlyHelpUpdateAndExport form to avoid going into design mode while running
        If (rs!ObjectName <> "cnlyHelpUpdateAndExport") Then
            
             curFrm = rs!ObjectName

             ' If the current form is different then the previous form, open new form and close prev form.
             If curFrm <> prevFrm Then
             ' If previouse form is still open, close form.
                 If Not frm Is Nothing Then
                    DoCmd.Close acForm, prevFrm, acSaveYes
                End If
        
                ' Open form
                DoCmd.OpenForm curFrm, acDesign
                
                'RaiseEvent StatusMessage("Mapping ", "Running", "Form: " & curFrm)
            
                ' Set Form object to active form.
                Set frm = screen.ActiveForm
                
                ' Set Form help file
                frm.helpFile = MvHelpPath & MvHelpFile
                            
            End If
        
            If Nz(rs!ControlName, "") = "" Then
                    frm.HelpContextId = Nz(rs!MapID, 0)     ' VT 04/22/2009 Setting the first controls context ID to the form
            End If
        
            'If Not IsNull(rs!ControlName) Then
            If Nz(rs!ControlName, "") <> "" Then
                ' Set Control object to control.
                Set ctl = frm.Controls(rs!ControlName)
                               
                ' Set Control HelpContextId
                ctl.HelpContextId = Nz(rs!MapID, 0)
                'Debug.Print frm.Name & "   " & ctl.Name & "   " & ctl.HelpContextId
            End If
        End If
        
ResumeToNextRecord:
        prevFrm = curFrm
        
        ' Move to next record in table
        rs.MoveNext
    Wend
    
    ' Close last form that maybe open
    If prevFrm <> "" Then
        DoCmd.Close acForm, prevFrm, acSaveYes
        RaiseEvent StatusMessage("Mapping ", "Completed", "Mapping Completed Successfully!")
        SetMapIds = True
    Else
        RaiseEvent StatusMessage("Mapping ", "Completed", "No matching records were found in the help table to update forms.")
        SetMapIds = False
    End If
    
ExitFunction:
On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set ctl = Nothing
    Set frm = Nothing

Exit Function
    
ErrorHandler:
'On Error Resume Next
    If Err.Number = 2102 Then
        Debug.Print "Form Error: " & Err.Description
        RaiseEvent StatusMessage("Mapping ", "Form Information", Err.Description)
        
   ElseIf Err.Number = 2465 Then
        RaiseEvent StatusMessage("Mapping ", "Control Information", Err.Description)
        Debug.Print "Control Error: " & Err.Description
        
    ElseIf Err.Number = 438 Then
        'Object doesn't support this property or method
        'RaiseEvent StatusMessage("Mapping ", "Information", err.Description)
        'Debug.Print "Form Error: " & err.Number & err.Description
        
    ElseIf Err.Number = 2467 Then
        'Trying to reference a form control that doesn't exist
        'RaiseEvent StatusMessage("Mapping ", "Information", err.Description)
        
    Else
        Debug.Print "Form Error: " & Err.Description
        RaiseEvent StatusMessage("Mapping ", "Error: Form " & curFrm & "/" & ctlName, "Error Number: " & Err.Number & " - " & Err.Description)
        'RaiseEvent StatusMessage("Mapping ", "Information: " & curFrm & "/" & ctlName, err.Number & "-" & err.Description)
    End If
    
    If (curFrm <> "") Then
        Resume ResumeToNextRecord
    Else
        SetMapIds = False
        Resume ExitFunction
    End If
    
End Function


Public Function ExportObjectList() As Boolean
On Error GoTo ErrorHandling
    ' Check to see if table exist before continuing.
    If Not ObjectExists(objTable, MvHelpTable) Then
        ' Create Table
        If Not CreateTable(MvHelpTable) Then
            ExportObjectList = False
            Exit Function
        End If
    Else
        ' Clear all records to prevent dub records.
        DeleteTable MvHelpTable
    End If
    
    Dim frmCurr As Form
    Dim objForm As AccessObject
    Dim FormsList() As AccessObject
    Dim Prefixes() As String
    Dim frmcounter As Integer
    Dim prefCounter As Integer
        
    ReDim FormsList(0 To CurrentProject.AllForms.Count - 1)
    
    Prefixes = GetAppPrefix
    
    frmcounter = 0
    For Each objForm In CurrentProject.AllForms
                
        ' Building list of form objects based on "prefixInclude" flag and each object name's prefix
        If PrefixInclude = True Then
                           
            ' VT 04/22/2009 Skipping the current cnlyHelpUpdateAndExport form to avoid going into design mode while running
            If (objForm.Name <> "cnlyHelpUpdateAndExport") And (InStr(1, objForm.Name, Prefixes(0)) <> 0) Then
                Set FormsList(frmcounter) = objForm
                frmcounter = frmcounter + 1
            End If
                
        Else
            prefCounter = 0
            While prefCounter <= UBound(Prefixes)
                If (objForm.Name <> "cnlyHelpUpdateAndExport") And (InStr(1, objForm.Name, Prefixes(prefCounter)) = 0) Then
                    Set FormsList(frmcounter) = objForm
                    frmcounter = frmcounter + 1
                End If
                prefCounter = prefCounter + 1
            Wend
        End If
            
    Next
    
    ReDim Preserve FormsList(0 To frmcounter - 1)
    
    RaiseEvent StatusMessage("Exporting ", "Running", "Exporting Controls' Properties...")
    
    ' Loop through current project forms
    frmcounter = 0
    While frmcounter <= UBound(FormsList)

        DoCmd.OpenForm FormsList(frmcounter).Name, acDesign

        ' Set Form object
        Set frmCurr = Forms(FormsList(frmcounter).Name)

        'VT 04/22/2009 Retrieving the forms context help id
        InsertObj frmCurr.Name, "", frmCurr.HelpContextId, "", ""

        ' Loop through controls
        ActiveCtrl frmCurr

        ' Close Form
        DoCmd.Close acForm, FormsList(frmcounter).Name, acSaveNo

        frmcounter = frmcounter + 1
    Wend
    
    'Set the exported table's Description property to Decipher's major and minor version
    SetTblversion
    
    RaiseEvent StatusMessage("Exporting ", "Completed", "Exporting Completed Successfully!")
    
    ExportObjectList = True
        
ExitFunction:
On Error Resume Next
    Set frmCurr = Nothing
    Set objForm = Nothing
    frmcounter = 0
    While frmcounter >= UBound(FormsList)
        Set FormsList(frmcounter) = Nothing
        frmcounter = frmcounter + 1
    Wend

Exit Function

ErrorHandling:
On Error Resume Next
    'MsgBox err.Number & ": " & err.Description
    RaiseEvent StatusMessage("Exporting ", "Error", "Unknown error during exporting")
    ExportObjectList = False
    If Nz(FormsList(frmcounter).Name, "") <> "" Then
        DoCmd.Close acForm, FormsList(frmcounter).Name, acSaveNo
    End If
    Resume ExitFunction
    
End Function

Private Sub ActiveCtrl(ByVal frm As Form)
On Error Resume Next
    Dim cxtId As Integer
    Dim ctlName As String
    Dim frmName As String
    Dim StatusText As String
    Dim ctlTipText As String
    
    cxtId = 0
    ctlName = ""
    frmName = ""
    StatusText = ""
    ctlTipText = ""
         
    ' Loop through each control in frm
    For Each ctrl In frm.Controls
        ' Set variables
        frmName = frm.Name
        ctlName = ctrl.Name
        cxtId = ctrl.HelpContextId
        ctlTipText = ctrl.ControlTipText
        StatusText = ctrl.StatusBarText
            
        ' Insert record into table "?"
        InsertObj frmName, ctlName, cxtId, ctlTipText, StatusText
    Next
        
Exit_ErrorHandler:
    On Error Resume Next
    Exit Sub
    
ErrorHandler:
    'MsgBox err.Number & ": " & err.Description
    RaiseEvent StatusMessage("ActiveControl ", "Error", "Unknown error during retrieving active controls")
    Resume Exit_ErrorHandler
End Sub


Private Sub InsertObj(ByVal ObjectName As String, ByVal ControlName As String, ByVal ContextID As String, ByVal ControlTipText As String, ByVal StatusText As String)
On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim SQL As String
    
    Set db = CurrentDb
    
    ' VT 04/23/2009 Added using replace function to mask any literal double quotes in the strings
    SQL = "Insert Into " & MvHelpTable & " (ObjectName , ControlName, MapID, ControlTipText, StatusBarText) " & _
            "Values (" & Chr(34) & ObjectName & Chr(34) & ", " & Chr(34) & ControlName & Chr(34) & ", " & Chr(34) & ContextID & Chr(34) & ", " & Chr(34) & Replace(ControlTipText, """", """""") & Chr(34) & ", " & Chr(34) & Replace(StatusText, """", """""") & Chr(34) & ")"
       
    db.Execute (SQL)
    
    If db.RecordsAffected = 0 Then
        Debug.Print "***Problem Entering Record"
    Else
        Debug.Print "SuccessFully Inserted Record"
    End If
    
    Set db = Nothing
    
Exit_ErrorHandler:
    On Error Resume Next
    db.Close
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Dim errX As DAO.Error
    
    If Errors.Count > 1 Then
        For Each errX In DAO.Errors
            Debug.Print "ODBC Error"
            Debug.Print errX.Number
            Debug.Print errX.Description
        Next errX
    Else
        Debug.Print "VBA Error"
        Debug.Print Err.Number
        Debug.Print Err.Description
    End If
    
    Debug.Print Err.Description
    'MsgBox err.Number & ": " & err.Description
    RaiseEvent StatusMessage("Inserting ", "Error", "Unknown error during inserting into help table")
    
    Resume Exit_ErrorHandler
End Sub

'01/24/2011 JL Replaced to use ObjectExists from CnlyDtFunction module
'Public Function ObjectExists(ByVal sObjType As Integer, ByVal sObjName As String) As Boolean
'    Dim db As DAO.Database
'    Dim Rs As DAO.Recordset
    
'    Set db = CurrentDb
'    Set Rs = db.OpenRecordset("Select id From msysobjects Where type=" & sObjType & " and name='" & sObjName & "';", dbOpenSnapshot)
    
'    If Not Rs.EOF Then
'      ObjectExists = True
'    End If
    
'    Rs.Close
'    Set Rs = Nothing
'    Set db = Nothing
'End Function

Private Function CreateTable(ByVal sTblName As String) As Boolean
On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim fld As DAO.Field
    Dim tdf As DAO.TableDef
    Dim idx As DAO.index
    Dim SQL As String
      
    Set db = CurrentDb

    SQL = "CREATE TABLE " & sTblName
    SQL = SQL & "(ObjectName TEXT(225), ControlName TEXT(225),  ControlTipText TEXT(225), StatusBarText TEXT(225));"
       
    'Create Help Table
    db.Execute SQL
    Set tdf = db.TableDefs(sTblName)
    
    'Add MapID field
    With tdf
        Set fld = .CreateField("MapID", dbInteger)
        fld.DefaultValue = 0
        .Fields.Append fld
        .Fields.Refresh
    End With
    
    'Add HelpID field to the table
    Set idx = tdf.CreateIndex("PrimaryKey")
    Set fld = tdf.CreateField("HelpID", dbLong)
    
    With fld
        .Attributes = .Attributes Or dbAutoIncrField
    End With
    
    With tdf.Fields
        .Append fld
        .Refresh
    End With

    'Add primary key contraint on the HelpID field
    SQL = "Alter Table " & sTblName & " Add Constraint HelpID_PK Primary Key (HelpID);"
    db.Execute SQL
        
    CreateTable = True
    RaiseEvent StatusMessage("Create Table", "Information", "A new help table for " & Chr(34) & MvAppName & Chr(34) & " application named " & Chr(34) & sTblName & Chr(34) & " has been created.")
    
Exit_ErrorHandler:
    On Error Resume Next
    db.Close
    Set db = Nothing
    Set idx = Nothing
    Set fld = Nothing
    Set tdf = Nothing
   
    Exit Function
ErrorHandler:
    CreateTable = False
    'MsgBox err.Number & ": " & err.Description
    RaiseEvent StatusMessage("Create Table ", "Error", "Unknown error during creating help table")
    
    Resume Exit_ErrorHandler
End Function


Private Sub DeleteTable(ByVal sTblName As String)
On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim SQL As String
    
    SQL = "Delete From " & sTblName & " "
    Set db = CurrentDb
    
    db.Execute (SQL)

    db.Close
    Set db = Nothing
Exit_ErrorHandler:
    On Error Resume Next
    db.Close
    Set db = Nothing
   
    Exit Sub
ErrorHandler:
    Dim errX As DAO.Error
    
    If Errors.Count > 1 Then
        For Each errX In DAO.Errors
            Debug.Print "ODBC Error"
            Debug.Print errX.Number
            Debug.Print errX.Description
        Next errX
    Else
        Debug.Print "VBA Error"
        Debug.Print Err.Number
        Debug.Print Err.Description
    End If
    RaiseEvent StatusMessage("Delete Table ", "Error", "Unknown error during deleting records from help table")
    'MsgBox err.Number & ": " & err.Description
    Resume Exit_ErrorHandler
End Sub

Private Sub SetTblversion()
On Error GoTo CreateProp

    Dim db As DAO.Database
    Dim tblDef As DAO.TableDef
    
    Set db = CurrentDb
    Set tblDef = db.TableDefs(MvHelpTable)
    'Setting an existing Description property's value
    tblDef.Properties("Description") = MvDecipherVer
    
ExitFunction:
On Error Resume Next
    Set db = Nothing
    Set tblDef = Nothing

Exit Sub

CreateProp:
On Error GoTo ErrorHandling
    If Err.Number = 3270 Or Err.Number = 0 Then
        Dim prp As Property
        'Creating the Description property and setting its value
        Set prp = tblDef.CreateProperty("Description", dbText, MvDecipherVer)
        tblDef.Properties.Append prp
        Resume ExitFunction
    Else
        Resume ErrorHandling
    End If

ErrorHandling:
On Error Resume Next
    'MsgBox err.Number & ": " & err.Description
    RaiseEvent StatusMessage("Table Description", "Error", "Error Setting Description to the help table " & MvHelpTable)
    Resume ExitFunction
   
End Sub

Private Sub GetDecipherVersion()
On Error GoTo GetFromTblData
    Dim db As DAO.Database
    Dim tblDef As DAO.TableDef
'    Dim recs As DAO.Recordset
'    Dim Sql As String
    
    Set db = CurrentDb
    Set tblDef = db.TableDefs("CT_AppStartupSeq")
        
    MvDecipherVer = left(tblDef.Properties("Description"), 3)
    
ExitSub:
On Error Resume Next
'    recs.Close
'    Set recs = Nothing
    Set db = Nothing
    Set tblDef = Nothing

Exit Sub

GetFromTblData:
'On Error Resume Next
    'If the table's description property is not set get Decipher version from the table data
'    If err.Number = 3270 Then
'        Sql = "Select Top 1 Version From SCR_ScreensVersions Order by VersionID Desc"
'
'        Set recs = DB.OpenRecordset(Sql, dbOpenSnapshot + dbReadOnly)
'        MvDecipherVer = Left(recs!version, 3)
'        Resume ExitSub
'    End If

'    RaiseEvent StatusMessage("Get Decipher Version", "Error", "Unable to retrieve Decipher version")
    MvDecipherVer = "Nill"
    Resume ExitSub
End Sub

Private Function CompareTblVersion() As Boolean
On Error GoTo ErrorHandling
    
    If ClsGenUtil.GetTblVersion(MvHelpTable) >= MvDecipherVer Then
        CompareTblVersion = True
    Else
        CompareTblVersion = False
        RaiseEvent StatusMessage("HelpTableVersion", "Error", "The " & Chr(34) & MvHelpTable & Chr(34) & " table version should be the same as Decipher's Major and Minor version")
    End If
    
Exit Function

ErrorHandling:
    'MsgBox err.Number & ": " & err.Description
    RaiseEvent StatusMessage("Compare version ", "Error", "Unknown error during comparing help table and Decipher version")
    CompareTblVersion = False
    
End Function


Private Function GetAppPrefix() As String()
On Error GoTo ErrorHandling
    Dim db As DAO.Database
    Dim recs As DAO.RecordSet
    Dim SQL As String
    Dim Prefix() As String
    Dim i As Integer
        
    Set db = CurrentDb
    
    SQL = "Select AppPrefix From CT_HelpConfig Where AppName = '" & MvAppName & "' "
    
    Set recs = db.OpenRecordSet(SQL, dbOpenSnapshot + dbReadOnly)
    
    If Nz(recs!AppPrefix = "") = "" Or recs!AppPrefix = "-No-" Then
        recs.Close
        SQL = "Select Count(AppPrefix) As RecCount From CT_HelpConfig Where (AppPrefix <> " & Chr(34) & Chr(34)
        SQL = SQL & " Or Not(AppPrefix) Is Null) And AppPrefix <> '-No-'"
        Set recs = db.OpenRecordSet(SQL, dbOpenSnapshot + dbReadOnly)
        
        'Set array length to the number of records retrieved
        ReDim Prefix(0 To recs!RecCount - 1)
        
        recs.Close
        
        SQL = "Select AppPrefix  From CT_HelpConfig Where (AppPrefix <> " & Chr(34) & Chr(34)
        SQL = SQL & " Or Not(AppPrefix) Is Null) And AppPrefix <> '-No-'"
        
        Set recs = db.OpenRecordSet(SQL, dbOpenSnapshot + dbReadOnly)
        
        PrefixInclude = False
        i = 0
        While Not recs.EOF
           Prefix(i) = recs!AppPrefix
           i = i + 1
           recs.MoveNext
        Wend
        GetAppPrefix = Prefix
    Else
        PrefixInclude = True
        ReDim Prefix(0 To 0)
        Prefix(0) = recs!AppPrefix
        GetAppPrefix = Prefix
    End If

ExitFunction:
On Error Resume Next
    recs.Close
    Set recs = Nothing
    Set db = Nothing

Exit Function

ErrorHandling:
    'MsgBox err.Number & ": " & err.Description
    RaiseEvent StatusMessage("Application Prefix", "Error", "Unknown error during finding application prefix(s)")
    Resume ExitFunction

End Function

Private Sub ClsGenUtil_StatusMessage(Src As String, Msg As String, lvl As Integer)
'If Lvl = 10 Then Stop
If lvl >= 10 Then
    RaiseEvent StatusMessage(Src, "Error", Msg)
End If

If lvl <= 0 Then 'FINISHED
    RaiseEvent StatusMessage(Src, "Information", Msg)
End If
End Sub

Function IsArrayEmpty(varArray As Variant) As Boolean
   ' Determines whether an array contains any elements.
   ' Returns False if it does contain elements, True
   ' if it does not.

   Dim ArrCount As Integer
   
   On Error Resume Next
   ' If the array is empty, an error occurs
   ArrCount = UBound(varArray)
   If Err.Number <> 0 Then
      IsArrayEmpty = True
   Else
      IsArrayEmpty = False
   End If
End Function
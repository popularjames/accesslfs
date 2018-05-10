Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This class is used to import objects from an Access DB
' By default object that exist in the destination will not be overwritten
'
' Basic example:
'    Dim clsImport as New MUT_ClsImportObjects
'    With clsImport
'        .SourceDatabase = "Path to your database file"
'        .LoadObjectList
'        .ImportAllObjects
'    End With
'
' Example with excludes:
'    Dim clsImport as New MUT_ClsImportObjects
'    With clsImport
'        .ExcludeObjectsTable "Name of table with export list", "Object Name Field", "Object Type Field"
'        .SourceDatabase = "Path to your database file"
'        .AddExcludeTableItem = "NotThisTable"
'        .AddExcludeFormItem = "CNLY*" 'No forms that start with CNLY
'        .LoadObjectList
'        .ImportAllObjects
'    End With
'
' Notes: There are several properties to exclude objects by name.
'        Use an * at the end to exclude a prefix.
'        You can edit the object collections outside this class before the import.
'
' SA 7/17/2012 - Created class
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Event ImportStatus(ByVal ObjectType As String, ByVal ObjectName As String, ByVal Status As String)
Public Event StatusMessage(ByVal Message As String)

Private SourceDB As String
Private OverwriteExisting As Boolean

Private CopyTable As Boolean
Private CopyQuery As Boolean
Private CopyForm As Boolean
Private CopyModule As Boolean
Private CopyMacro As Boolean
Private CopyReport As Boolean
Private CopyTableRelationship As Boolean

Private TableList As Collection
Private QueryList As Collection
Private FormList As Collection
Private ModuleList As Collection
Private MacroList As Collection
Private ReportList As Collection

Private TableExclude As Collection
Private QueryExclude As Collection
Private FormExclude As Collection
Private ModuleExclude As Collection
Private MacroExclude As Collection
Private ReportExclude As Collection

Private ImportErrors As Collection

'* jc tmp
Private Enum AccessObjectType
    Table = 1
    TableLinked = 4
    Query = 5
    Form = -32768
    Module = -32761
    Report = -32764
    Macro = -32766
End Enum

Private Sub Class_Initialize()
    Set ImportErrors = New Collection
    
    'Object lists
    Set TableList = New Collection
    Set QueryList = New Collection
    Set FormList = New Collection
    Set ModuleList = New Collection
    Set MacroList = New Collection
    Set ReportList = New Collection
    
    'Exclude lists
    Set TableExclude = New Collection
    Set QueryExclude = New Collection
    Set FormExclude = New Collection
    Set ModuleExclude = New Collection
    Set MacroExclude = New Collection
    Set ReportExclude = New Collection
    
    'Options
    CopyTable = True
    CopyQuery = True
    CopyForm = True
    CopyModule = True
    CopyMacro = True
    CopyReport = True
    CopyTableRelationship = True
    
    'Filter system objects
    TableExclude.Add "MSYS*"
    TableExclude.Add "USYS*"
    QueryExclude.Add "~*"
    
End Sub

'''''''''''''''''''''''''''''''' LET Properties '''''''''''''''''''''''''''''''''''''''
Public Property Let SourceDatabase(ByVal FilePath As String)
    SourceDB = FilePath
End Property

Public Property Let OverwriteExistingObjects(ByVal Overwrite As Boolean)
    OverwriteExisting = Overwrite
End Property

Public Sub AddExcludeTableItem(ByVal ExcludeText As String)
    If LenB(ExcludeText) > 0 Then
        TableExclude.Add ExcludeText
    End If
End Sub

Public Sub AddExcludeFormItem(ByVal ExcludeText As String)
    If LenB(ExcludeText) > 0 Then
        FormExclude.Add ExcludeText
    End If
End Sub

Public Sub AddExcludeModuleItem(ByVal ExcludeText As String)
    If LenB(ExcludeText) > 0 Then
        ModuleExclude.Add ExcludeText
    End If
End Sub

Public Sub AddExcludeQueryItem(ByVal ExcludeText As String)
    If LenB(ExcludeText) > 0 Then
        QueryExclude.Add ExcludeText
    End If
End Sub

Public Sub AddExcludeMacroItem(ByVal ExcludeText As String)
    If LenB(ExcludeText) > 0 Then
        MacroExclude.Add ExcludeText
    End If
End Sub

Public Sub AddExcludeReportItem(ByVal ExcludeText As String)
    If LenB(ExcludeText) > 0 Then
        ReportExclude.Add ExcludeText
    End If
End Sub

Public Property Let CopyTables(ByVal CopyObject As Boolean)
    CopyTable = CopyObject
End Property

Public Property Let CopyQueries(ByVal CopyObject As Boolean)
    CopyQuery = CopyObject
End Property

Public Property Let CopyForms(ByVal CopyObject As Boolean)
    CopyForm = CopyObject
End Property

Public Property Let CopyModules(ByVal CopyObject As Boolean)
    CopyModule = CopyObject
End Property

Public Property Let CopyMacros(ByVal CopyObject As Boolean)
    CopyMacro = CopyObject
End Property

Public Property Let CopyReports(ByVal CopyObject As Boolean)
    CopyReport = CopyObject
End Property

Public Property Let CopyTableRelationships(ByVal CopyObject As Boolean)
    CopyTableRelationship = CopyObject
End Property

Public Property Let SetTableList(ByRef NewList As Collection)
    Set TableList = NewList
End Property

Public Property Let SetQueryList(ByRef NewList As Collection)
    Set QueryList = NewList
End Property

Public Property Let SetFormList(ByRef NewList As Collection)
    Set FormList = NewList
End Property

Public Property Let SetReportList(ByRef NewList As Collection)
    Set ReportList = NewList
End Property

Public Property Let SetMacroList(ByRef NewList As Collection)
    Set MacroList = NewList
End Property

Public Property Let SetModuleList(ByRef NewList As Collection)
    Set ModuleList = NewList
End Property


'''''''''''''''''''''''''''''''' GET Properties '''''''''''''''''''''''''''''''''''''''
Public Property Get GetTableList() As Collection
    Set GetTableList = TableList
End Property

Public Property Get GetQueryList() As Collection
    Set GetQueryList = QueryList
End Property

Public Property Get GetFormList() As Collection
    Set GetFormList = FormList
End Property

Public Property Get GetReportList() As Collection
    Set GetReportList = ReportList
End Property

Public Property Get GetMacroList() As Collection
    Set GetMacroList = MacroList
End Property

Public Property Get GetModuleList() As Collection
    Set GetModuleList = ModuleList
End Property

Public Property Get GetTotalObjectCount() As Integer
    GetTotalObjectCount = TableList.Count + QueryList.Count + _
        FormList.Count + ReportList.Count + MacroList.Count + ModuleList.Count
End Property

Public Property Get GetImportErrorCount() As Integer
    GetImportErrorCount = ImportErrors.Count
End Property


'''''''''''''''''''''''''''''''' Methods '''''''''''''''''''''''''''''''''''''''

Public Sub ExcludeObjectsTable(ByVal TableName As String, ByVal NameField As String, ByVal TypeField As String)
'Add exclude objects to list from a table
On Error GoTo ErrorHappened
    Dim SQL As String
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet

    Set db = CurrentDb
    SQL = "SELECT " & NameField & "," & TypeField & " FROM " & TableName & " WHERE CheckNamingConvention = 0 ORDER BY " & TypeField
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)
    
    Do Until rs.EOF
        Select Case rs.Fields(TypeField)
            Case "Table"
                AddExcludeTableItem rs.Fields(NameField)
            Case "Query"
                AddExcludeQueryItem rs.Fields(NameField)
            Case "Macro"
                AddExcludeMacroItem rs.Fields(NameField)
            Case "Form"
                AddExcludeFormItem rs.Fields(NameField)
            Case "Module"
                AddExcludeModuleItem rs.Fields(NameField)
            Case "Report"
                AddExcludeReportItem rs.Fields(NameField)
        End Select
        
        rs.MoveNext
    Loop
ExitNow:
On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

Public Sub LoadObjectList()
'Load table objects into collections
On Error GoTo ErrorHappened
    Dim SQL As String
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim ObjectName As String
    
    'Build SQL
    SQL = "SELECT [Name],[Type] FROM MSysObjects WHERE [Type] IN("
    If CopyTable Then
        SQL = SQL & AccessObjectType.Table & ","
    End If
    If CopyQuery Then
        SQL = SQL & AccessObjectType.Query & ","
    End If
    If CopyForm Then
        SQL = SQL & AccessObjectType.Form & ","
    End If
    If CopyModule Then
        SQL = SQL & AccessObjectType.Module & ","
    End If
    If CopyMacro Then
        SQL = SQL & AccessObjectType.Macro & ","
    End If
    If CopyReport Then
        SQL = SQL & AccessObjectType.Report & ","
    End If
    
    If Right(SQL, 1) = "," Then
        'Remove comma and close SQL
        SQL = left(SQL, Len(SQL) - 1) & ") ORDER BY [Type],[Name]"
        
        'Reset collections
        If TableList.Count > 0 Then
            Set TableList = New Collection
        End If
        If QueryList.Count > 0 Then
            Set QueryList = New Collection
        End If
        If FormList.Count > 0 Then
            Set FormList = New Collection
        End If
        If ModuleList.Count > 0 Then
            Set ModuleList = New Collection
        End If
        If MacroList.Count > 0 Then
            Set MacroList = New Collection
        End If
        If ReportList.Count > 0 Then
            Set ReportList = New Collection
        End If
        
        Set db = DBEngine.OpenDatabase(SourceDB)
        Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)
        
        Do Until rs.EOF
            ObjectName = rs![Name]
            Select Case rs![Type]
                Case AccessObjectType.Table
                    If Not IsExcluded(ObjectName, acTable) Then
                        TableList.Add ObjectName, ObjectName
                    End If
                Case AccessObjectType.Query
                    If Not IsExcluded(ObjectName, acQuery) Then
                        QueryList.Add ObjectName, ObjectName
                    End If
                Case AccessObjectType.Form
                    If Not IsExcluded(ObjectName, acForm) Then
                        FormList.Add ObjectName, ObjectName
                    End If
                Case AccessObjectType.Module
                    If Not IsExcluded(ObjectName, acModule) Then
                        ModuleList.Add ObjectName, ObjectName
                    End If
                Case AccessObjectType.Macro
                    If Not IsExcluded(ObjectName, acMacro) Then
                        MacroList.Add ObjectName, ObjectName
                    End If
                Case AccessObjectType.Report
                    If Not IsExcluded(ObjectName, acReport) Then
                        ReportList.Add ObjectName, ObjectName
                    End If
            End Select
            rs.MoveNext
        Loop
    End If
ExitNow:
On Error Resume Next
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

Public Sub ImportAllObjects()
    ImportFormObjects
    ImportModuleObjects
    ImportTableObjects
    ImportQueryObjects
    ImportReportObjects
    ImportMacroObjects
End Sub

Public Sub ImportTables()
    ImportTableObjects
End Sub

Public Sub ImportQueries()
    ImportQueryObjects
End Sub

Public Sub ImportForms()
    ImportFormObjects
End Sub

Public Sub ImportReports()
    ImportReportObjects
End Sub

Public Sub ImportMacros()
    ImportMacroObjects
End Sub

Public Sub ImportModules()
    ImportModuleObjects
End Sub

Private Function ImportTableObjects() As Boolean
'Import all tables in collection
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim i As Integer
    
    For i = 1 To TableList.Count
        ImportObject TableList.Item(i), acTable
        ImportTableRelationships TableList.Item(i)
    Next

    Result = True
ExitNow:
On Error Resume Next
    ImportTableObjects = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

Private Sub ImportTableRelationships(ByVal TableName As String)
'Import table relationships
On Error GoTo ErrorHappened
    Dim DbDest As DAO.Database
    Dim DbSource As DAO.Database
    Dim RelSource As Relation
    Dim RelTarget As Relation
    Dim i As Integer
    
    If CopyTableRelationship Then
        Set DbDest = CurrentDb
        Set DbSource = DBEngine.OpenDatabase(SourceDB)
        
        DbSource.Relations.Refresh
    
        For Each RelSource In DbSource.Relations
            If RelSource.ForeignTable = TableName Then
                Set RelTarget = DbDest.CreateRelation(RelSource.Name, RelSource.Table, TableName, RelSource.Attributes)
                For i = 0 To RelSource.Fields.Count - 1
                    RelTarget.Fields.Append RelTarget.CreateField(RelSource.Fields(i).Name)
                    RelTarget.Fields(i).ForeignName = RelSource.Fields(i).ForeignName
                Next i
                DbDest.Relations.Append RelTarget
            End If
        Next RelSource
    End If
ExitNow:
On Error Resume Next
    DbSource.Close
    Set DbDest = Nothing
    Set DbSource = Nothing
    Set RelSource = Nothing
    Set RelTarget = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

Private Function ImportQueryObjects() As Boolean
'Loop through all queries in selected db
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim i As Integer
    
    For i = 1 To QueryList.Count
        ImportObject QueryList.Item(i), acQuery
    Next
    
    Result = True
ExitNow:
On Error Resume Next
    ImportQueryObjects = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

Private Function ImportFormObjects() As Boolean
'Loop through all forms in selected db
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim i As Integer

    For i = 1 To FormList.Count
        ImportObject FormList.Item(i), acForm
    Next
    
    Result = True
ExitNow:
On Error Resume Next
    ImportFormObjects = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

Private Function ImportReportObjects() As Boolean
'Loop through all modules in selected db
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim i As Integer
    
    For i = 1 To ReportList.Count
        ImportObject ReportList.Item(i), acReport
    Next
    
    Result = True
ExitNow:
On Error Resume Next
    ImportReportObjects = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

Private Function ImportMacroObjects() As Boolean
'Loop through all modules in selected db
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim i As Integer
    
    For i = 1 To MacroList.Count
        ImportObject MacroList.Item(i), acMacro
    Next
    
    Result = True
ExitNow:
On Error Resume Next
    ImportMacroObjects = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

Private Function ImportModuleObjects() As Boolean
'Loop through all modules in selected db
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim i As Integer
    
    For i = 1 To ModuleList.Count
        ImportObject ModuleList.Item(i), acModule
    Next
    
    Result = True
ExitNow:
On Error Resume Next
    ImportModuleObjects = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

Private Sub ImportObject(ByVal ObjectName As String, ByVal ObjectType As AcObjectType)
'Import object
On Error GoTo ErrorHappened
    Dim ObjTypeText As String
    DoEvents
    
    ObjTypeText = ObjecTypeToString(ObjectType)
    
    'Delete object first for overwrites
    If OverwriteExisting Then
        DeleteObject ObjectName, ObjectType
    End If
    
    RaiseEvent StatusMessage("Importing " & ObjTypeText & ": " & ObjectName)

    'Copy object
    DoCmd.TransferDatabase acImport, "Microsoft Access", SourceDB, ObjectType, ObjectName, ObjectName
    
    RaiseEvent ImportStatus(ObjTypeText, ObjectName, "IMPORTED")
ExitNow:
On Error Resume Next
    
Exit Sub
ErrorHappened:
    ImportErrors.Add "Error importing " & ObjTypeText & " " & ObjectName
    RaiseEvent ImportStatus(ObjTypeText, ObjectName, "ERROR!")
    Resume ExitNow
End Sub

Private Sub DeleteObject(ByVal ObjectName As String, ByVal ObjectType As AcObjectType)
'Delete object
On Error GoTo ErrorHappened
    DoCmd.DeleteObject ObjectType, ObjectName
    
ExitNow:

Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

Private Function ObjectExists(ByVal ObjectName As String, ByVal ObjectType As AcObjectType) As Boolean
'Check to see if object exists in current database
On Error GoTo ErrorHappened
    Dim Result As Boolean

    Select Case ObjectType
        Case acTable
            ObjectName = CurrentData.AllTables.Item(ObjectName).Name
        Case acQuery
            ObjectName = CurrentData.AllQueries.Item(ObjectName).Name
        Case acForm
            ObjectName = CurrentProject.AllForms.Item(ObjectName).Name
        Case acModule
            ObjectName = CurrentProject.AllModules.Item(ObjectName).Name
        Case acReport
            ObjectName = CurrentProject.AllReports.Item(ObjectName).Name
        Case acMacro
            ObjectName = CurrentProject.AllMacros.Item(ObjectName).Name
    End Select
    
    Result = True
ExitNow:
On Error Resume Next
    ObjectExists = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

Private Function IsExcluded(ByVal ObjectName As String, ByVal ObjectType As AcObjectType) As Boolean
'Check to see if object is excluded from being imported
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim TempCollection As Collection
    Dim i As Integer
    
    Select Case ObjectType
        Case acTable
            Set TempCollection = TableExclude
        Case acQuery
            Set TempCollection = QueryExclude
        Case acForm
            Set TempCollection = FormExclude
        Case acModule
            Set TempCollection = ModuleExclude
        Case acReport
            Set TempCollection = ReportExclude
        Case acMacro
            Set TempCollection = MacroExclude
    End Select

    With TempCollection
        For i = 1 To TempCollection.Count
            'Check for matching name
            If ObjectName = .Item(i) Then
                Result = True
                Exit For
            End If
            
            If Not Result Then
                'Check for matching prefix
                If Right(.Item(i), 1) = "*" Then
                    If left(ObjectName, Len(.Item(i)) - 1) = left(.Item(i), Len(.Item(i)) - 1) Then
                        Result = True
                        Exit For
                    End If
                End If
            End If
            
            If Not Result And Not OverwriteExisting Then
                'Check to see if object exists in current db
                If ObjectExists(ObjectName, ObjectType) Then
                    Result = True
                    Exit For
                End If
            End If
        Next
    End With
ExitNow:
On Error Resume Next
    Set TempCollection = Nothing
    IsExcluded = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

Private Function ObjecTypeToString(ByVal ObjectType As AcObjectType) As String
    Dim ObjType As String
    Select Case ObjectType
        Case acTable
            ObjType = "TABLE"
        Case acQuery
            ObjType = "QUERY"
        Case acForm
            ObjType = "FORM"
        Case acModule
            ObjType = "MODULE"
        Case acReport
            ObjType = "REPORT"
        Case acMacro
            ObjType = "MACRO"
    End Select

    ObjecTypeToString = ObjType
End Function
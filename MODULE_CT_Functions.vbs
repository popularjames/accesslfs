Option Compare Database
Option Explicit
'CCA Screens Globals
'CnlyDtFunctions
'CnlyTotalsCustom

'Moved from CCA Screens Globals
Global Identity As New CT_ClsIdentity
Global Telemetry As New CT_ClsTelemetry

'Moved from CnlyDtFunctions
Private addInManager As New CT_ClsCnlyAddinSupport

'Moved from CCA Screens Globals
Public Type CnlyFldDef
    Name As String
    Alias As String
    ControlSrc As String
    Type As Byte
    Format As String
    Width As Single
    Height As Single
    left As Single
    Align As Byte
    Decimal As Integer
End Type

'Moved from CCA Screens Globals
'Describes the Status of Decipher add-ins
Public Enum CnlyAddinStatus
    None = 0            'undetermined state
    Loaded = 1          'add-in is loaded in Access and available to unload
    Unloaded = 2        'add-in is not loaded in Access and available to load
    NotFound = 4        'add-in was not found in the avaliable list for Access. it might not have been installed
    DisabledLocal = 8   'add-in is Disabled in the table CT_AddinDecipherVersion not available to load or unload
End Enum

'Moved from CCA Screens Globals
Public Enum ObjType
    objTable = acTable
    objQuery = acQuery
    objForm = acForm
    objReport = acReport
    objModule = acModule
    objMacro = acMacro
End Enum

'Moved from CCA Screens Globals
'USED FOR UTILITY SQL in ClsImport
Public Type ReplacePairs
    From As String
    To As String
End Type

'Moved from CCA Screens Globals
Public Enum CnlyExportObj
    MainGrid = 1
    ActiveTab = 2
    Totals = 3
End Enum

'Moved from CCA Screens Globals
Public Type CnlyDataSheetStyle
    GridlinesBehavior As Byte
    GridlinesColor As Long
    BackGroundColor As Long
    BorderLineStyle As Byte
    CellsEffect As Byte
    HeaderUnderlineStyle As Byte
    fontsize As Byte
    FontFamily As String
    FontItalic As Boolean
    FontUnderline As Boolean
    FontWeight As Integer
    ForeColor As Long
End Type

'Moved from CnlyDtFunctions
Public Function PutXML(Xml As String, AdRst As Object)
    'Use ADO to and an in-memory stream to read XML into a recordset
    'XML is the data to open as a recordset.
    'Results returned as AdRst
    
    On Error GoTo ErrorHappened

    Dim AdStrm 'As New ADODB.Stream
    
    'NOTE use the Stream.WriteText (StrVar) method to get the XML into the Stream object
    'Then use Recorset.Open Stream to get the XML as a recorset.
    
    Set AdRst = CreateObject("ADODB.Recordset")
    Set AdStrm = CreateObject("ADODB.Stream")
    
    AdStrm.Type = 2 'Text
    AdStrm.Charset = "ascii"
    AdStrm.Open
    
    'DEBUG LINES Start
        'AdStrm.LoadFromFile "c:\dev\Screens 1.4\Unified\cnlyscreens.xml"
    'DEBUG LINES End
    
    AdStrm.WriteText Xml
    
        
    AdStrm.position = 0
    
    AdRst.Open AdStrm
    
    

ExitNow:
    On Error Resume Next
    AdStrm.Close
    Set AdStrm = Nothing
    Exit Function
    
ErrorHappened:
    MsgBox Err.Description, vbCritical, "PutXML()"
    Resume ExitNow
    Resume
End Function

Public Function InCollection(vKey As Variant, ThisCollection As Collection) As Boolean
    'Returns True if the Key vKey is in ThisCollection.  False otherwise
    
    On Error Resume Next
    Dim msValue
    
    msValue = ThisCollection.Item(vKey).Item(1)
    
    If Err.Number = 0 Then
        InCollection = True
    Else
        InCollection = False
    End If
    
End Function

Public Sub CreateItemExportSQL(DbSource As Database, BaseTable As String, baseID As String, ItemName As String, colExportScripts As Collection, Optional Level As Integer = 0)
'Store SQL for exporting "BaseTable" and all related tables  in a collection object "ColExportScripts".
'SA 11/13/12 - Reworked to work with new user tables with broken relationships
On Error GoTo ErrorHappened
    
    Dim relLoop As Relation
    Dim fldArray() As String
    Dim SQL As String
    Dim SQL_BaseId As String
    Dim i As Integer
    Dim foreignID As String
    Dim strSQL As String
    Dim db As Database
    Dim ColTable As New Collection
    Dim ColForeignTable As New Collection
    Dim r As Integer
    
    Dim exportScript As CT_ClsExportScript
           
    SQL = "SELECT "
    
    If Level = 0 Then
        'Create export script for the "BaseTable"
        Set exportScript = New CT_ClsExportScript
        fldArray = GetFieldArray(DbSource, BaseTable)
        
        'If the table has a field called "RefID", then it is a "Base" table.  Base tables Primary key fields that are foreign
        '   key fields in other "child" tables.  It is important to restore all of the base tables first.
        ' hc 8/9/2010 - changed to use a contains rather than last field in the list
        If ContainsRefId(fldArray) Then ' base table
            exportScript.ExpTyp = 0 'Base
            strSQL = "UPDATE " & BaseTable & " SET RefID = " & baseID & " WHERE " & GetPrimKeyName(DbSource, BaseTable) & " = " & baseID
            Set db = CurrentDb
            DoCmd.SetWarnings False
            db.Execute strSQL, dbSeeChanges
            DoCmd.SetWarnings True
            Set db = Nothing
        Else 'Is leaf table
            ' DS mar 26 2010 changed Leaf to ExpTypeVal.Leaf
            exportScript.ExpTyp = ExpTypeVal.Leaf
        
        End If
        For i = 1 To UBound(fldArray) - 1
            SQL = SQL & "[" & fldArray(i) & "], "
        Next i
            
            ' add the last field to the sql
        SQL = SQL & fldArray(UBound(fldArray)) & " FROM " & BaseTable & " "
        SQL = SQL & "WHERE [" & GetPrimKeyName(DbSource, BaseTable) & "] IN (" & baseID & ")"

        exportScript.ExpTable = BaseTable
        exportScript.ExpScript = SQL
        exportScript.ExpLevel = 0
        exportScript.ExpName = ItemName
        Level = 1
        
        On Error Resume Next
        colExportScripts.Add exportScript, BaseTable
        On Error GoTo ErrorHappened
    End If

    ' Enumerate the Relations collection of the current
    ' database to report on the property values of
    ' the Relation objects and their Field objects.
    
    'Load table relationships into collection
    For Each relLoop In DbSource.Relations
        With relLoop
            If .Table = BaseTable Then
                ColTable.Add .Table
                ColForeignTable.Add .ForeignTable
            End If
        End With
    Next
    #If ccSCR = 1 Then
        SCR_ScreensExportKeysOverride ColTable, ColForeignTable
    #ElseIf ccSCR = 2 Then
        SCR_AppScreensExportKeysOverride ColTable, ColForeignTable
    #End If
    
    For r = 1 To ColTable.Count
        Set exportScript = New CT_ClsExportScript
        SQL = "SELECT "
        
        If ColTable.Item(r) = BaseTable Then
            fldArray = GetFieldArray(DbSource, ColForeignTable.Item(r))
            ' HC 8/9/2010 - changed to use a contains rather than
            If ContainsRefId(fldArray) Then ' base table
                exportScript.ExpTyp = ExpTypeVal.Base
                strSQL = "UPDATE " & ColForeignTable.Item(r) & " SET RefID = " & GetPrimKeyName(DbSource, ColForeignTable.Item(r)) & " WHERE ScreenId=" & baseID
                Set db = CurrentDb
                DoCmd.SetWarnings False
                db.Execute strSQL
                DoCmd.SetWarnings True
                Set db = Nothing
            Else 'Is leaf table
                 exportScript.ExpTyp = Leaf
            End If
            
            For i = 1 To UBound(fldArray) - 1
                SQL = SQL & "[" & fldArray(i) & "], "
            Next i
            
            ' add the last field to the sql
            SQL = SQL & fldArray(UBound(fldArray)) & " FROM " & ColForeignTable.Item(r) & " "
            SQL = SQL & "WHERE [" & GetPrimKeyName(DbSource, BaseTable) & "] IN (" & baseID & ")"
                                 
            exportScript.ExpScript = SQL
            exportScript.ExpLevel = Level
            exportScript.ExpName = ItemName
            exportScript.ExpTable = ColForeignTable.Item(r)
            
            On Error Resume Next
            colExportScripts.Add exportScript, ColForeignTable.Item(r)
            On Error GoTo ErrorHappened
            
            If exportScript.ExpTyp = ExpTypeVal.Base Then
                'recurse child tables for the current foreign table
                '
                'Need the Primary key IDs for the tables we are getting child records for
                SQL_BaseId = "SELECT " & GetPrimKeyName(DbSource, ColForeignTable.Item(r)) & " FROM " & ColForeignTable.Item(r) & " "
                SQL_BaseId = SQL_BaseId & "WHERE " & GetPrimKeyName(DbSource, BaseTable) & " IN (" & baseID & ")"
                
                foreignID = BuildStringFromQuery(SQL_BaseId, DbSource)
                'Only process child tables if there are matching records in the parent
                If foreignID <> "" Then
                    'recurse, preserving the item name (i.e. screen name), bump the level by 1
                    CreateItemExportSQL DbSource, ColForeignTable.Item(r), foreignID, ItemName, colExportScripts, Level + 1
                End If
            End If
        End If
    Next
ExitNow:
On Error Resume Next
    
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Creating Table Export Scripts"
    Resume ExitNow
    Resume
End Sub

Public Function GetXML(DbPath As String, strSQL As String) As String
'Use ADO to and an in-memory stream to generate XML fo for a geven Query against the specified Access DB.
'Return the resulting XML string.
On Error GoTo ErrorHappened
Dim AdRst 'As New ADODB.Recordset
Dim AdCon 'As New ADODB.Connection
Dim AdStrm 'As New ADODB.Stream


Set AdRst = CreateObject("ADODB.Recordset")
Set AdCon = CreateObject("ADODB.Connection")
Set AdStrm = CreateObject("ADODB.Stream")

Dim SQL As String
Dim Xml As String

' HC changed to use the variable for the access provider
AdCon.Open LINK_SRC_ACCESS & "Data Source=" & DbPath & ";"

AdRst.CursorLocation = 3 '3 adUseClient

SQL = strSQL
AdRst.Open SQL, AdCon, 3 'adOpenStatic

' Note that if you don't specify
' adPersistXML, a binary format (ADTG) will be used by default.
AdRst.Save AdStrm, 1 '1 adPersistXML  'Save to an ADO stream for memory retrieval

Xml = AdStrm.ReadText(AdStrm.Size)

'If InStr(1, Sql, "SCR_ScreensLayoutsFormats") > 0 Then
'    Debug.Print Sql
'    Debug.Print XML
'End If

AdStrm.Close
AdRst.Close
AdCon.Close


GetXML = Xml

ExitNow:
    On Error Resume Next
    Set AdStrm = Nothing
    Set AdRst = Nothing
    AdCon.Close
    Set AdCon = Nothing
    Exit Function
    
ErrorHappened:
    MsgBox Err.Description, vbCritical, "GetXML"
    Resume ExitNow
    Resume
End Function

Public Function GetFieldArray(SrcDb As DAO.Database, strTable As String) As Variant
'returns an array of field names for the given table in the specified Db.

    On Error GoTo eCatch
    
    Dim TbDef As TableDef
    Dim strFields() As String
    Dim i As Integer
    Dim fldCount As Integer
    

    Set TbDef = SrcDb.TableDefs(strTable)
    With TbDef
    
    
    fldCount = .Fields.Count
    ReDim strFields(fldCount - 1)
        
        For i = 0 To (fldCount - 1)
            strFields(i) = .Fields(i).Name

        Next i
        
    End With
    
    GetFieldArray = strFields

eCatch:
    On Error Resume Next
    Set TbDef = Nothing

    Exit Function
    Resume


End Function

Private Function ContainsRefId(ByRef fldArray() As String) As Boolean
' HC 8/9/2010 - function to determine if the field array contains the field RefID
    Dim bReturn As Boolean
    Dim i As Integer
    bReturn = False
    For i = 1 To UBound(fldArray)
        If fldArray(i) = "RefId" Then
            bReturn = True
            i = UBound(fldArray) + 1
        End If
    Next
    ContainsRefId = bReturn
End Function

Public Function FormIsOpen(FormName As String) As Boolean
    Dim Result As Boolean
    Dim i As Integer
    Result = False
    For i = 0 To Application.Forms.Count - 1
        If Application.Forms(i).Name = FormName Then
            Result = True
            Exit For
        End If
    Next i
    
    'tmpStr = Application.Forms(FormName).Name
    'FormIsOpen = True
    
     FormIsOpen = Result

End Function

Public Function ObjectExists(ByVal typeValue As ObjType, ByVal strObjectName As String) As Boolean
    On Error Resume Next
     Dim db As Database
     Dim Tbl As TableDef
     Dim Qry As QueryDef
     Dim i As Integer
     Dim ReturnVal As Boolean
     ReturnVal = False
     
     Set db = CurrentDb()
    
     Select Case typeValue
        Case ObjType.objTable
          For Each Tbl In db.TableDefs
               If Tbl.Name = strObjectName Then
                    ReturnVal = True
                    Exit For
               End If
          Next Tbl
        
        Case ObjType.objQuery
          For Each Qry In db.QueryDefs
               If Qry.Name = strObjectName Then
                    ReturnVal = True
                    Exit For
               End If
          Next Qry

        Case ObjType.objMacro
          For i = 0 To db.Containers("Scripts").Documents.Count - 1
               If db.Containers("Scripts").Documents(i).Name = strObjectName Then
                    ReturnVal = True
                    Exit For
               End If
          Next i
        
        Case Else
            Dim strObjectType As String
            strObjectType = ""
            
            If typeValue = objForm Then
                strObjectType = "Forms"
            ElseIf typeValue = objModule Then
                strObjectType = "Modules"
            ElseIf typeValue = objReport Then
                strObjectType = "Reports"
            End If
            
            For i = 0 To db.Containers(strObjectType).Documents.Count - 1
               If db.Containers(strObjectType).Documents(i).Name = strObjectName Then
                    ReturnVal = True
                    Exit For
               End If
            Next i
           
     End Select
     
   ObjectExists = ReturnVal
      
End Function

Public Function GetFieldFormatCustomTotals(TotalID As Long, FldNum As Integer, DataType As Byte, ByRef Decimals As Integer, FieldName As String, ScreenID As Long) As String
On Error GoTo ErrorHappened
Dim TmpFormat As String, CalcFormat As String
Dim NumDecimals As Integer, CalcDecimals As Byte
Dim SQL As String, db As DAO.Database, rs As DAO.RecordSet
'Left 1
'Center 2
'Right 3

SQL = "SELECT TOP " & FldNum & " * FROM ("
SQL = SQL & "SELECT FldName, Alias,Ordinal, 2 as SRC , Format, FieldWidth, Align, Decimals "
SQL = SQL & "FROM SCR_ScreensTotalsCalculations  "
SQL = SQL & "Where TotalID = " & TotalID & " "

SQL = SQL & "UNION ALL "

SQL = SQL & "SELECT FldName, Alias,Ordinal, 1 as SRC , Format, FieldWidth, Align, Decimals "
SQL = SQL & "From SCR_ScreensTotalsFields "
SQL = SQL & "Where FldType = 1 and TotalID = " & TotalID & " "
SQL = SQL & ") "
SQL = SQL & "ORDER BY SRC, Ordinal, Alias, FldName;"


Set db = CurrentDb
Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)

rs.MoveLast

NumDecimals = Nz(rs.Fields("Decimals"), -1)
TmpFormat = "" & rs.Fields("Format")
rs.Close
If NumDecimals = -1 Or TmpFormat = "" Then
    Select Case DataType
        Case dbDate
            If IsUnitedStates = True Then
                CalcFormat = "mm/dd/yy"
                CalcDecimals = 255 'Auto
            End If
        Case dbByte
            CalcFormat = ""
            CalcDecimals = 0 'Auto
        Case dbLong, dbInteger
            CalcFormat = "Standard"
            CalcDecimals = 0 'Auto
        Case dbCurrency, dbSingle, dbDouble, dbDecimal, dbNumeric
            CalcFormat = "Standard"
            CalcDecimals = 2 'Auto
        Case dbText, dbLongBinary, dbMemo, dbBoolean
            CalcFormat = ""
        Case Else
            CalcFormat = ""
    End Select
End If
If TmpFormat <> "" Then
    GetFieldFormatCustomTotals = TmpFormat
Else
    GetFieldFormatCustomTotals = CalcFormat
End If
If NumDecimals <> -1 Then
    Decimals = NumDecimals
Else
    Decimals = CalcDecimals
End If

ExitNow:
    On Error Resume Next
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error getting custom total formats"
    Resume ExitNow
    Resume
End Function
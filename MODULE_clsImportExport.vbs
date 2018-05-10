Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsImportExport
' DateTime  : 11/14/2005 16:11
' Author    : joseph.casella
' Purpose   : Creates a simplified interface to import/export Access & Excel Files
' Modifications : 8/18/08 Added ExportExcelRecordset
'            12/1/05 Added .ExportExcelSql so that a sql string can be used as a query source
' To Do :   Add Text import/export and Excel Import
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit



Public Function ExportAccess(SourceObject As String, DestDB As String, DestinationTbl As String) As Long

    On Error GoTo HandleError

    '* Not using DoCmd.TransferDatabase because it doesn't work with linked tables.  It only copies the link,
    '* not the data

    Dim strSQL As String

    strSQL = "SELECT * INTO " & DestinationTbl & " IN '" & DestDB & "' FROM " & SourceObject & " src"

    ExportAccess = cIeRunDAO(strSQL)
exitHere:
    Exit Function

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    Resume exitHere

End Function

Public Function ExportExcel(SourceObject As String, FilePath As String)
    On Error GoTo HandleError

    DoCmd.TransferSpreadsheet acExport, , SourceObject, FilePath

exitHere:
    Exit Function

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    Resume exitHere

End Function

Public Function ExportExcelSql(strSQL As String, FilePath As String)
    On Error GoTo HandleError

    '* DoCmd.TransferSpreadsheet doesn't work with a sql string.  This creates a "temporary" query, performs
    '* the export and then deletes the query.

    Dim qdf As QueryDef

    Set qdf = CurrentDb.CreateQueryDef("temp_Export", strSQL)

    qdf.Close

    RefreshDatabaseWindow

    ExportExcel qdf.Name, FilePath


exitHere:
    On Error Resume Next

    DoCmd.DeleteObject acQuery, qdf.Name

    Set qdf = Nothing
    Exit Function

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    Resume exitHere


End Function

Public Function ExportExcelRecordset(rst As ADODB.RecordSet, FilePath As String, Optional IncludeHeader As Boolean = True)
    On Error GoTo HandleError

Dim oExcel As Object '* Excel.Application
Dim oBook As Object '* Excel.Workbook
Dim oSheet As Object '* Excel.Worksheet

Dim iField As Integer
Dim iRow As Integer

Set oExcel = CreateObject("Excel.Application")

Set oBook = oExcel.Workbooks.Add
Set oSheet = oBook.Worksheets(1)

'Insert column headings
    If IncludeHeader = True Then
        oSheet.Rows(1).Font.Bold = True

        For iField = 0 To rst.Fields.Count - 1
            oSheet.Cells(1, iField + 1).Value = rst.Fields(iField).Name
        Next iField
        
    End If

'* Transfer
    
     oSheet.Cells(2, 1).CopyFromRecordset rst
   
     oBook.SaveAs FilePath
     oExcel.visible = True
      
exitHere:
    On Error Resume Next
    
    Set oExcel = Nothing
    Set oBook = Nothing
    Set oSheet = Nothing

    Exit Function

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    Resume exitHere
End Function


Public Function ImportAccessTbl(SourceDB As String, SourceTbl As String, DestinationTbl As String) As Long
    On Error GoTo HandleError

    Dim strSQL As String

    DoCmd.TransferDatabase acImport, "Microsoft Access", SourceDB, acTable, SourceTbl, DestinationTbl


exitHere:
    Exit Function

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    Resume exitHere

End Function


Public Function ImportAccessSql(Optional SQL As String, Optional Source As String, Optional DestinationTbl As String, Optional CreateDestination As Boolean = False) As Long
    On Error GoTo HandleError

    Dim strStmt As String
    Dim dbError As Error
    Dim strError As String

    Dim db As Database

    Set db = CurrentDb

    If SQL = "" Then

        If CreateDestination = False Then

            strStmt = " INSERT * INTO " & DestinationTbl & " FROM (" & Source & ")"

        Else

            strStmt = " SELECT * INTO " & DestinationTbl & " FROM (" & Source & ")"

        End If

    Else
        strStmt = SQL

    End If

    db.Execute strStmt, dbSeeChanges + dbFailOnError

    ImportAccessSql = db.RecordsAffected

exitHere:


    Set dbError = Nothing
    Set db = Nothing

    Exit Function

HandleError:

    DBEngine.Errors.Refresh

    If DBEngine.Errors(DBEngine.Errors.Count - 1).Number = Err.Number Then

        For Each dbError In DBEngine.Errors

            If dbError.Number = 2601 Or dbError.Number = 2627 Then    '*Primary Key Violation

                MsgBox " One or more rows aleady exist in your data.  Insert Cancelled", vbCritical + vbOKOnly, "Record(s) Exist"

                Err.Raise dbError.Number

                Resume Next

            Else

                strError = vbCr & strError & dbError.Number & " (" & dbError.Description & ")"

            End If

        Next dbError
    Else

        MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"

    End If

    Debug.Print strError

    MsgBox "Error In Access Import" & vbCr & vbCr & strError

    Resume exitHere

End Function

Public Function CreateAccessDB(FilePath As String)

    Dim ws As DAO.Workspace
    Dim db As Database

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.CreateDatabase(FilePath, dbLangGeneral)
    db.Close

exitHere:

    Set ws = Nothing
    Set db = Nothing
    Exit Function


HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"

    Resume exitHere

End Function

Private Function cIeRunDAO(strSQL As String) As Long

    Dim db As Database

    '* This is the same as the regular RunDAO function, copied here so there is no external dependency.

    On Error GoTo HandleErrors
    Set db = CurrentDb

    db.Execute strSQL, dbFailOnError + dbSeeChanges

    cIeRunDAO = db.RecordsAffected

exitHere:

    Exit Function

HandleErrors:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"

End Function


Public Sub ViewExcel(FileName As String)
    On Error GoTo HandleErrors

    Dim objXL   'As Excel.Application
    Set objXL = CreateObject("Excel.Application")

    objXL.visible = True
    objXL.Workbooks.Open FileName

exitHere:
    Set objXL = Nothing
    Exit Sub

HandleErrors:
    MsgBox "Error: " & Err.Number & " ( " & Err.Description & ")" & vbCr & vbCr & "Filename: " & FileName, vbOKOnly
    GoTo exitHere

End Sub

Public Sub ExportText(ExportFile As String, ByVal FilePath As String)

    Dim strSQL As String
    Dim db As Database
    Dim strLine As String
    Dim rst As RecordSet
    Dim intI As Integer


    On Error GoTo ErrHandler

    'Open Output file
    Open FilePath For Output As #2

    Set db = CurrentDb

    strSQL = "  SELECT * from " & ExportFile & " "
    Set rst = db.OpenRecordSet(strSQL)

    While Not rst.EOF
        intI = 0
        strLine = ""


        For intI = 0 To rst.Fields.Count - 1
            strLine = strLine & rst.Fields(intI)

        Next intI

        Print #2, strLine    'Trim(strLine)

        rst.MoveNext
    Wend

    Close #2

    Exit Sub

ErrHandler:
    On Error Resume Next
    Close #2
    MsgBox "Error Writing File - " & FilePath & " : " & Err.Description, vbOKCancel + vbCritical
End Sub
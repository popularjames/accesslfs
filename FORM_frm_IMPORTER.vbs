Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


''' Last Modified: 06/29/2012
''' KDearing: Created form
'''

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal _
        lpFileName As String) As Long
        
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
        
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long


Private Const cstrDefaultConnString As String = ";Extended Properties=""text;HDR=YES"""
Private Const cstrDriver As String = "Provider=Microsoft.Jet.OLEDB.5.0;Data Source="

Private WithEvents coCN As ADODB.Connection
Attribute coCN.VB_VarHelpID = -1
Private cdtProcStartDt As Date

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Private Sub cmdImport_Click()
    If Nz(Me.cmbFileType, "") = "" Then
        LogMessage ClassName & ".cmdImport_Click", , "Please select the document type first!", , True
        Exit Sub
    End If
    
    Call ImportNow
    
End Sub

Private Sub cmdPickFile_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oFileDialog As FileDialog
'Dim oFso As Scripting.FileSystemObject

    strProcName = ClassName & ".cmdPickFile_Click"
'    Set oFso = New Scripting.FileSystemObject
    Set oFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    If Me.txtFullPathToFileToImport <> "" Then
        oFileDialog.InitialFileName = Me.txtFullPathToFileToImport
    End If
    
    oFileDialog.Title = "Select the file to import"
    
    oFileDialog.show
    If oFileDialog.SelectedItems.Count > 0 Then
        Me.txtFullPathToFileToImport = oFileDialog.SelectedItems(1)
        Call EnableNextStep(2)
    End If
    ' otherwise they canceld..
    

Block_Exit:
    Set oFileDialog = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub



Private Sub EnableNextStep(intStepToEnable As Integer)
On Error GoTo Block_Err
Dim strProcName As String
Dim sUserPrompt As String
Dim bValidated As Boolean

    strProcName = ClassName & ".EnableNextStep"
        
    Select Case intStepToEnable
    Case 1
    Case 2
        
        If Nz(Me.txtFullPathToFileToImport, "") = "" Then
            sUserPrompt = "Please select a file to import first"
            bValidated = False
        Else
            If FileExists(Me.txtFullPathToFileToImport) = False Then
                sUserPrompt = "Cannot find the file specified, please click the 'Select file' button to verify"
                bValidated = False
            End If
            bValidated = True
        End If
            
        Me.cmbFileType.Enabled = bValidated
        Me.CmdImport.Enabled = bValidated
        
    Case Else
    
    End Select

    
Block_Exit:
    If sUserPrompt <> "" Then
        LogMessage strProcName, , sUserPrompt, , True
    End If
    
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub

'' Returns the path to a copy of the file with the header fixed
'' also sNewHdrRow will be populated
Private Function FixHeader(ByVal sOrigFilePath As String, oMappingRS As ADODB.RecordSet, ByRef sNewHdrRow As String, sDelimiter As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oTxtStreamFrom As Scripting.TextStream
Dim oTxtStreamTo As Scripting.TextStream
Dim oFso As Scripting.FileSystemObject
Dim strFileDir As String
Dim sFileExtension As String
Dim sFileCopy As String
Dim strFileCpyName As String
Dim sOldHdrRow As String
Dim saryFileFields() As String
Dim iAryIdx As Integer

    strProcName = ClassName & ".FixHeader"
    
    Set oFso = New Scripting.FileSystemObject
    
    sOrigFilePath = Me.txtFullPathToFileToImport
    strFileDir = QualifyFldrPath(oFso.GetParentFolderName(sOrigFilePath))
    
    sFileExtension = oFso.GetExtensionName(sOrigFilePath)
    
    sFileCopy = Replace(sOrigFilePath, "." & sFileExtension, "_IMPORT." & sFileExtension)
    FixHeader = sFileCopy
    
        ' don't need to do this since we're doing it via the text stream..
        '    oFso.CopyFile sOrigFilePath, sFileCopy, True
    
    strFileCpyName = Replace(sFileCopy, strFileDir, "")
    

    
        '' Open the file, get the first line
    Set oTxtStreamFrom = oFso.OpenTextFile(FileName:=sOrigFilePath, ioMode:=ForReading, Create:=False)
    sOldHdrRow = oTxtStreamFrom.ReadLine
    
        '' Create our new header:
    saryFileFields = Split(sOldHdrRow, sDelimiter)
    
    For iAryIdx = 0 To UBound(saryFileFields)
        oMappingRS.MoveFirst
        oMappingRS.Find "FileFieldName = '" & Trim(saryFileFields(iAryIdx)) & "'"  ' I get paranoid about white space :D
        
        If Not oMappingRS.EOF And Not oMappingRS.BOF Then
            sNewHdrRow = sNewHdrRow & oMappingRS("Fld_FieldName").Value & sDelimiter
        Else
            Stop ' Hammer time! Problem !!
        End If
        
    Next
    
    If Right(sNewHdrRow, 1) = sDelimiter Then
        sNewHdrRow = left(sNewHdrRow, Len(sNewHdrRow) - 1)  ' remove final delimiter
    End If
    
        '' Create our copy
    If FileExists(sFileCopy) Then
        DeleteFile sFileCopy, False
    End If
    
    Set oTxtStreamTo = oFso.OpenTextFile(FileName:=sFileCopy, ioMode:=ForWriting, Create:=True)
        
    oTxtStreamTo.WriteLine sNewHdrRow
    
        ' now write the rest of the file:
    oTxtStreamTo.Write oTxtStreamFrom.ReadAll
    
        ' Close up shop
    oTxtStreamTo.Close
    oTxtStreamFrom.Close
    
    
Block_Exit:
    Set oTxtStreamTo = Nothing
    Set oTxtStreamFrom = Nothing
    Set oFso = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function


Public Function ImportNow() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sOrigFile As String
Dim strFileDir  As String
Dim sFileCopy As String
Dim sFileExtension As String
Dim oFso As Scripting.FileSystemObject
Dim sNewHdrRow As String
Dim sInputTableName As String
Dim strFileCpyName As String
Dim sSelectClause As String
Dim sPostImportSproc As String
Dim sDelimiter As String
Dim slnkdTName As String

    strProcName = ClassName & ".ImportNow"

        '' ok, how are we going to do this?
        '' I say, make a copy of the file to the users tmp folder
    Set oFso = New Scripting.FileSystemObject

        '' Get some details about the file that we'll need:
    sOrigFile = Me.txtFullPathToFileToImport
    strFileDir = QualifyFldrPath(oFso.GetParentFolderName(sOrigFile))

        '' fix the header based on our field mapping table in FLD-009.CMS_AUDITORS_CLAIMS
        '' although, I'm not sure we'll need to do this since we're using a Schema.ini file..
        '' eh, good measure..
    
    Set oRs = GetFieldMapping(Me.cmbFileType)

    sDelimiter = Nz(oRs("Delimiter").Value, ",")

        ' let's grab the destination table name:
    sInputTableName = Nz(oRs("DestinationTableName").Value, "")

    If sInputTableName = "" Then
        Stop ' Hammer time! Problem fix yer stuff!
    End If

        ' the name of the proc to run after the import
    sPostImportSproc = Nz(oRs("PostImportSproc").Value, "")

        '' Note: sNewHdrRow will be returned populated
    sFileCopy = FixHeader(sOrigFile, oRs, sNewHdrRow, sDelimiter)
    strFileCpyName = Replace(sFileCopy, strFileDir, "")
    
    
        '' Since the Jet provider isn't installed on our term servers (!!%$#@@$!@#)
        '' we can't load straight to an ADODB recordset, so we'll create a linked table
        '' to the file, then we can get our ado recordset, or, God forbid, just use
        '' Jet to select * into SqlServer Linked table
        ''  But, in order to do that we should make a schema.ini file:
        '' So we need to get an ADO RS with the tables meta data (field defs, etc)
        '' Other options are to import via an ImportSpecification set up once per file (and edited each time the thing changes!
        '' technically, this SHOULD be faster while a maintenance nightmare (and Claim Admin would have to be deployed
        '' each time an ImportSpec is changed unless we put those in a separate database and import them each time.. ugh.
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString(sInputTableName)
        .SQLTextType = sqltext
        .sqlString = "SELECT " & sNewHdrRow & " FROM [" & sInputTableName & "] WHERE 1 = 2 "    ' Note, we don't want any data
                                                                                                ' just field info
        Set oRs = .ExecuteRS
        If oRs Is Nothing Then
            Stop ' why nothing ? should be something just no RecordCount
        End If
    End With
    
        '' For this we need to set up a Schema.ini file in the same directory as the file
    If CreateSchemaIniFromAdoRs(oRs, True, strFileDir, strFileCpyName, sDelimiter) = False Then
        Stop ' hammer time!
    End If
    
    slnkdTName = CreateLinkedTable(strFileCpyName, sFileExtension, strFileDir)
    
          ' Not sure this is ok or not.. joe?
    oAdo.sqlString = "TRUNCATE TABLE " & sInputTableName
    oAdo.Execute

    '' the date is just for optimizing..
Dim dtStart As Date
    dtStart = Now
    Call UseADOToTransfer(slnkdTName, sInputTableName, sNewHdrRow)
        ' Another means to do the same thing but SHOULD (in Theory) take longer, but then again, we're still limited
        ' with the whole linked table thing..
            '    Call UseJetToTransfer("tmp_lnk_" & Replace(strFileCpyName, "." & sFileExtension, ""), sInputTableName, sNewHdrRow)
    Debug.Print ProcessTookHowLong(dtStart)
        ' Process took 3:57 with Page Size = 20 CacheSize = 20
    
    ' unlink the table:
    Call RemoveLinkedTable(slnkdTName)

    '' This proc seems to take a while to run so I'm going to run it asynchronously (my spellng sucks)!

    '' Now if we got this far we can execute the stored proc:
    If sPostImportSproc <> "" Then
    
        Set coCN = New ADODB.Connection
        coCN.ConnectionString = GetConnectString("V_CODE_DATABASE")
        coCN.CommandTimeout = 0 ' no timeout thanks!
        coCN.CursorLocation = adUseServer
        coCN.Open
        cdtProcStartDt = Now
        coCN.Execute sPostImportSproc, , adCmdStoredProc + adExecuteNoRecords + adAsyncExecute
    
    End If
    
    MsgBox "Finished (but the stored procedure is still runnin!"
    
Block_Exit:
    Set oFso = Nothing
    Set oRs = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function



Private Sub RemoveLinkedTable(sLinkedTblName As String)
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oTDef As DAO.TableDef

    strProcName = ClassName & ".RemoveLinkedTable"
    
    Set oDb = CurrentDb
    If IsTable(sLinkedTblName) Then
        oDb.TableDefs.Delete sLinkedTblName
        oDb.TableDefs.Refresh
    End If
    oDb.TableDefs.Refresh
    
    
Block_Exit:
    Set oDb = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub

Private Function CreateLinkedTable(strFileCpyName As String, sFileExtension As String, sFileDir As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oTDef As DAO.TableDef

    strProcName = ClassName & ".FixHeader"
    
    Set oDb = CurrentDb
    If IsTable("tmp_lnk_" & Replace(strFileCpyName, "." & sFileExtension, "")) Then
        oDb.TableDefs.Delete "tmp_lnk_" & Replace(strFileCpyName, "." & sFileExtension, "")
        oDb.TableDefs.Refresh
    End If
    Set oTDef = oDb.CreateTableDef("tmp_lnk_" & Replace(strFileCpyName, "." & sFileExtension, ""))
    oTDef.Connect = "Text;DATABASE=" & sFileDir & ";TABLE=" & strFileCpyName
    oTDef.SourceTableName = strFileCpyName
    oDb.TableDefs.Append oTDef
    oDb.TableDefs.Refresh
    
    CreateLinkedTable = "tmp_lnk_" & Replace(strFileCpyName, "." & sFileExtension, "")
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    CreateLinkedTable = ""
    GoTo Block_Exit
End Function

Private Function UseJetToTransfer(sLocalTable As String, sDestTable As String, sSelString As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String
Dim oDb As DAO.Database

    strProcName = ClassName & ".UseJetToTransfer"
    
    Set oDb = CurrentDb
    
    sSql = "INSERT INTO " & sDestTable & " (" & sSelString & ") SELECT " & sSelString & " FROM [" & sLocalTable & "]"
    oDb.Execute sSql

Block_Exit:
    Set oDb = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function


Private Function UseADOToTransfer(sLocalTable As String, sDestTable As String, sSelString As String) As Boolean
On Error GoTo Block_Err
Dim oCn As ADODB.Connection
Dim oRs As ADODB.RecordSet
Dim oDestRs As ADODB.RecordSet
Dim strProcName As String
Dim sSql As String
Dim oDb As DAO.Database
Dim oFld As ADODB.Field
Dim oCmd As ADODB.Command

    strProcName = ClassName & ".UseADOToTransfer"
    
    Set oCn = New ADODB.Connection
    Set oCn = Application.CurrentProject.Connection

    Set oRs = New ADODB.RecordSet
    oRs.CursorLocation = adUseClientBatch
    oRs.CursorType = adOpenKeyset
    oRs.LockType = adLockBatchOptimistic
    
    Call oRs.Open("SELECT " & sSelString & " FROM [" & sLocalTable & "]", oCn)
    
    Set oRs.ActiveConnection = Nothing
    
        '' insert that now
    
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = GetConnectString(sDestTable)
    oCn.Open
    
    Set oCmd = New ADODB.Command
    With oCmd
        
    End With
    
    Set oDestRs = New ADODB.RecordSet
    With oDestRs
        .CursorLocation = adUseClientBatch
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        
            '' Need to play around with these to see how much quicker we can make this..
            '' Server was too busy today for me to get any
            '' kind of good results, of course I should have used  but kind of in the middle of other
            '' things too! :D
        .PageSize = 10
        .CacheSize = 30
        
        .ActiveConnection = oCn
        .Open "SELECT " & sSelString & " FROM " & sDestTable
    End With

    While Not oRs.EOF
        oDestRs.AddNew
        For Each oFld In oRs.Fields
            oDestRs(oFld.Name) = oRs(oFld.Name).Value
        Next
        oRs.MoveNext
    Wend
        
    oDestRs.UpdateBatch
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function



'' Need to change this by:
''  Make it a sproc
''  Use the file ID from the drop down instead of the file name
Private Function GetFieldMapping(SFileName As String) As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String

    strProcName = ClassName & ".ImportNow"
    
    If SFileName = "" Then GoTo Block_Exit
    
    
    sSql = "SELECT F.FileNameDesc, F.DestinationTableName, F.DestinationDb, FM.FieldPosition, FM.FileFieldName, " & _
            " FM.Fld_FieldName, F.Delimiter, F.PostImportSproc " & _
            " FROM ADMIN_Import_Files F INNER JOIN ADMIN_Import_Files_FieldMaps FM ON F.FileId = FM.FileID " & _
            " WHERE F.FileNameDesc = '" & SFileName & "'"
    
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_DATA_DATABASE")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            Stop    ' hammer time! Problem!
        End If
    End With
    
    Set GetFieldMapping = oRs
Block_Exit:
    Set oAdo = Nothing
    Set oRs = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function




Private Function CreateSchemaIniFromDAOTable(bIncFldNames As Boolean, sPath As String, _
        sSectionName As String, sTblQryName As String) As Boolean
Dim Msg As String ' For error handling.
On Error GoTo Block_Err
Dim strProcName As String
Dim oWs As DAO.Workspace, oDb As DAO.Database
Dim tblDef As DAO.TableDef, fldDef As DAO.Field
Dim i As Integer, handle As Integer
Dim FldName As String, fldDataInfo As String
    
    
    strProcName = ClassName & ".CreateSchemaIniFromDAOTable"
        ' -----------------------------------------------
        ' Set DAO objects.
        ' -----------------------------------------------
    Set oDb = CurrentDb()
        ' -----------------------------------------------
        ' Open schema file for append.
        ' -----------------------------------------------
    handle = FreeFile
    Open sPath & "schema.ini" For Output Access Write As #handle
        ' -----------------------------------------------
        ' Write schema header.
        ' -----------------------------------------------
    Print #handle, "[" & sSectionName & "]"
    Print #handle, "ColNameHeader = " & _
                    IIf(bIncFldNames, "True", "False")
    Print #handle, "CharacterSet = ANSI"
    Print #handle, "Format = TabDelimited"
        ' -----------------------------------------------
        ' Get data concerning schema file.
        ' -----------------------------------------------
    Set tblDef = oDb.TableDefs(sTblQryName)
    With tblDef
       For i = 0 To .Fields.Count - 1
          Set fldDef = .Fields(i)
          With fldDef
             FldName = .Name
             Select Case .Type
                Case dbBoolean
                   fldDataInfo = "Bit"
                Case dbByte
                   fldDataInfo = "Byte"
                Case dbInteger
                   fldDataInfo = "Short"
                Case dbLong
                   fldDataInfo = "Integer"
                Case dbCurrency
                   fldDataInfo = "Currency"
                Case dbSingle
                   fldDataInfo = "Single"
                Case dbDouble
                   fldDataInfo = "Double"
                Case dbDate
                   fldDataInfo = "Date"
                Case dbText
                   fldDataInfo = "Char Width " & Format$(.Size)
                Case dbLongBinary
                   fldDataInfo = "OLE"
                Case dbMemo
                   fldDataInfo = "LongChar"
                Case dbGUID
                   fldDataInfo = "Char Width 16"
             End Select
             Print #handle, "Col" & Format$(i + 1) _
                             & "=" & FldName & Space$(1) _
                             & fldDataInfo
          End With
       Next i
    End With
    
    CreateSchemaIniFromDAOTable = True

Block_Exit:
    Close handle
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Function CreateSchemaIniFromAdoRs(oRs As ADODB.RecordSet, bIncFldNames As Boolean, sPath As String, _
        SFileName As String, Optional sDelimiter As String = ",") As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdoFld As ADODB.Field
Dim i As Integer, handle As Integer
Dim FldName As String, fldDataInfo As String
Dim sDelimFormat As String
    
    strProcName = ClassName & ".CreateSchemaIniFromAdoRs"
    
    Select Case sDelimiter
    Case ","
        sDelimFormat = "CSVDelimited"
    Case "\t", vbTab
        sDelimFormat = "TabDelimited"
    Case Else
        sDelimFormat = "Delimited(" & sDelimiter & ")"
    End Select
        
        ' -----------------------------------------------
        ' Set DAO objects.
        ' -----------------------------------------------
'    Set oDb = CurrentDb()
        ' -----------------------------------------------
        ' Open schema file for append.
        ' -----------------------------------------------
    handle = FreeFile
    If FileExists(sPath & "schema.ini") Then
        DeleteFile sPath & "schema.ini", False
    End If
    Open sPath & "schema.ini" For Output Access Write As #handle
        ' -----------------------------------------------
        ' Write schema header.
        ' -----------------------------------------------
    Print #handle, "[" & SFileName & "]"
    Print #handle, "ColNameHeader = " & IIf(bIncFldNames, "True", "False")
    Print #handle, "CharacterSet = ANSI"
    Print #handle, "Format = " & sDelimFormat
        
        ' -----------------------------------------------
        ' Get data concerning schema file.
        ' -----------------------------------------------

        For Each oAdoFld In oRs.Fields
            FldName = oAdoFld.Name
            Select Case AdoTypeToDaoType(oAdoFld)
            Case dbBoolean
                fldDataInfo = "Bit"
            Case dbByte
                fldDataInfo = "Byte"
            Case dbInteger
                fldDataInfo = "Short"
            Case dbLong
                fldDataInfo = "Integer"
            Case dbCurrency
                fldDataInfo = "Currency"
            Case dbSingle
                fldDataInfo = "Single"
            Case dbDouble
                fldDataInfo = "Double"
            Case dbDate
                fldDataInfo = "Date"
            Case dbText
                fldDataInfo = "Char Width " & Format$(oAdoFld.DefinedSize)
            Case dbLongBinary
                fldDataInfo = "OLE"
            Case dbMemo
                fldDataInfo = "LongChar"
            Case dbGUID
                fldDataInfo = "Char Width 16"
            End Select

            Print #handle, "Col" & Format$(i + 1) & "=" & FldName & Space$(1) _
                    & fldDataInfo
            i = i + 1

        Next
    
    
    CreateSchemaIniFromAdoRs = True

Block_Exit:
    Close handle
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function














Private Function CreateImportSpec(sImportSpecName As String, sFullPathToImportFile As String, oRs As ADODB.RecordSet, sTableName As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sXml As String
Dim oFso As Scripting.FileSystemObject
Dim SFileName As String

    strProcName = ClassName & ".ImportNow"
    Set oFso = New Scripting.FileSystemObject
    
        ' Create the import table
    SFileName = oFso.GetFileName(sFullPathToImportFile)
    
        ' Now create the import spec but only if we don't have one already
        ' this needs to be modified but it illustrates how to create an Import Spec via code

    sXml = ""
    sXml = sXml & "[?sXml version=""1.0""?]" & vbCrLf
    sXml = sXml & "[ImportExportSpecification Path=""" & sFullPathToImportFile & """ sXmlns=""urn:www.microsoft.com/office/access/imexspec""]" & vbCrLf
    sXml = sXml & "    [ImportExcel FirstRowHasNames=""true"" Destination=""my_tble_name"" Range=""'sheet1$'""]" & vbCrLf
    sXml = sXml & "        [Columns PrimaryKey=""{Auto}""]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col1"" FieldName=""name_of_column1"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col2"" FieldName=""name_of_column2"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col3"" FieldName=""name_of_column3"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col4"" FieldName=""name_of_column4"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col5"" FieldName=""name_of_column5"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col6"" FieldName=""name_of_column6"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col7"" FieldName=""name_of_column"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col8"" FieldName=""name_of_column8"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col9"" FieldName=""name_of_column9"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col10"" FieldName=""name_of_column10"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col11"" FieldName=""name_of_column11"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col12"" FieldName=""name_of_column12"" Indexed=""NO"" SkipColumn=""false"" DataType=""Text""/]" & vbCrLf
    sXml = sXml & "            [Column Name=""Col13"" FieldName=""name_of_column13"" Indexed=""NO"" SkipColumn=""false"" DataType=""Double""/]" & vbCrLf
    sXml = sXml & "        [/Columns]" & vbCrLf
    sXml = sXml & "    [/ImportExcel]" & vbCrLf
    sXml = sXml & "[/ImportExportSpecification]"



Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function




Private Sub coCN_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.RecordSet, ByVal pConnection As ADODB.Connection)
Dim sTimeTook As String

    sTimeTook = ProcessTookHowLong(cdtProcStartDt)
    MsgBox "The stored proc has finished it seems!!! it is now safe to go back into the water! (and close this form!)", , "Proc Complete - it took: " & sTimeTook

End Sub

Private Sub coCN_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Info message: " & pError.Description
End Sub

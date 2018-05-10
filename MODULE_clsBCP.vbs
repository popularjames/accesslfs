Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 07/10/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''
'''
'''  TODO:
'''  =====================================
'''  - add -e
'''  - check the output file capturing bit
'''
'''  HISTORY:
'''  =====================================
'''  - 03/29/2012 - Created...
'''
''' AUTHOR
'''  =====================================
''' Kevin Dearing
'''
'''
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
'' To get a format file made for me..
''
'' bcp CMS_AUDITORS_WORKSPACE.dbo.Subs_Invoice_imm format nul -T -c -S DS-FLD-009 -f subs_invoice_imm.txt
''
                'bcp [database_name.] schema.{table_name | view_name | "query" {in data_file | out data_file | queryout data_file | format nul}
                '
                '   [-a packet_size]
                '   [-b batch_size]
                '   [-c]
                '   [-C { ACP | OEM | RAW | code_page } ]
                '   [-d database_name]
                '   [-e err_file]
                '   [-E]
                '   [-f format_file]
                '   [-F first_row]
                '   [-h"hint [,...n]"]
                '   [-i input_file]
                '   [-k]
                '   [-K application_intent]
                '   [-L last_row]
                '   [-m max_errors]
                '   [-N]
                '   [-N]
                '   [-o output_file]
                '   [-P password]
                '   [-q]
                '   [-r row_term]
                '   [-R]
                '   [-S [server_name[\instance_name]]
                '   [-t field_term]
                '   [-T]
                '   [-U login_id]
                '   [-v]
                '   [-V (80 | 90 | 100 )]
                '   [-w]
                '   [-x]
                '   /?
                '



Private coRs As ADODB.RecordSet

Private csInTable As String
Private csSqlStatement As String
Private csServer As String
Private csDatabaseName As String
Private csRowTerminator As String
Private csFieldTerminator As String
Private csDataFilePath As String
Private csBCPCmdOutFilePath As String
Private csCommandResponse As String
Private csLastCmd As String
Private csOwner As String
Private csHint As String

Private ciPackageSize As Integer
Private ciBatchSize As Integer

Private csFormatFile As String

Private cbnlQuotedId As Boolean

Private cblnUseUnicode As Boolean
Private cblnUseCharData As Boolean
Private cblnUseNativeData As Boolean
Private cblnUseUnicodeAndNative As Boolean

Private cblnSelfCleanupFiles As Boolean


Private cblnUseQuotedIdentifiers As Boolean
Private cblnCaptureCmdOut As Boolean
Private cblnIdentityInsert As Boolean

Private gbImporting As Boolean

Private ccolCreatedFiles As Collection


''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get RecordSet() As ADODB.RecordSet
    Set RecordSet = coRs
End Property
Public Property Let RecordSet(oRs As ADODB.RecordSet)
    Set coRs = oRs
End Property


Public Property Get UseQuotedIdentifiers() As Boolean       ' - q switch
    UseQuotedIdentifiers = cblnUseQuotedIdentifiers
End Property
Public Property Let UseQuotedIdentifiers(bUseQuotedIdentifiers As Boolean)
    cblnUseQuotedIdentifiers = bUseQuotedIdentifiers
End Property




''' ############################################################
''' ############################################################
'''     Mutually exclusive options
''' ############################################################
''' ############################################################

        Public Property Get UseUnicode() As Boolean       ' - w switch
            UseUnicode = cblnUseUnicode
        End Property
        Public Property Let UseUnicode(bUseUnicode As Boolean)
            cblnUseUnicode = bUseUnicode
            If bUseUnicode = True Then
                cblnUseCharData = False
                cblnUseNativeData = False
                cblnUseUnicodeAndNative = False
            End If
        End Property


        Public Property Get UseCharData() As Boolean       ' - w switch
            UseCharData = cblnUseCharData
        End Property
        Public Property Let UseCharData(blnUseCharData As Boolean)
            cblnUseCharData = blnUseCharData
            If blnUseCharData = True Then
                cblnUseUnicode = False
                cblnUseNativeData = False
                cblnUseUnicodeAndNative = False
            End If
        End Property


        Public Property Get UseNativeData() As Boolean       ' - w switch
            UseNativeData = cblnUseNativeData
        End Property
        Public Property Let UseNativeData(blnUseNativeData As Boolean)
            cblnUseNativeData = blnUseNativeData
            If blnUseNativeData = True Then
                cblnUseUnicode = False
                cblnUseCharData = False
                cblnUseUnicodeAndNative = False
            End If
        End Property


        Public Property Get UseUnicodeAndNative() As Boolean       ' - w switch
            UseUnicodeAndNative = cblnUseUnicodeAndNative
        End Property
        Public Property Let UseUnicodeAndNative(blnUseUnicodeAndNative As Boolean)
            cblnUseUnicodeAndNative = blnUseUnicodeAndNative
            If blnUseUnicodeAndNative = True Then
                cblnUseUnicode = False
                cblnUseCharData = False
                cblnUseNativeData = False
            End If
        End Property

''' ############################################################
''' ############################################################
'''     / Mutually exclusive options
''' ############################################################
''' ############################################################





Public Property Get IdentityInsert() As Boolean       ' -e switch
    IdentityInsert = cblnIdentityInsert
End Property
Public Property Let IdentityInsert(bIdentityInsert As Boolean)
    cblnIdentityInsert = bIdentityInsert
End Property



Public Property Get CaptureCmdOut() As Boolean       ' - w switch
    CaptureCmdOut = cblnCaptureCmdOut
End Property
Public Property Let CaptureCmdOut(blnCaptureCmdOut As Boolean)
    cblnCaptureCmdOut = blnCaptureCmdOut
End Property




Public Property Get BatchSize() As Integer  '   - b switch
    BatchSize = ciBatchSize
End Property
Public Property Let BatchSize(iBatchSize As Integer)
    ciBatchSize = iBatchSize
End Property




Public Property Get PackageSize() As Integer    ' -a switch
    PackageSize = ciPackageSize
End Property
Public Property Let PackageSize(iPackageSize As Integer)
    ciPackageSize = iPackageSize
End Property


Public Property Get CommandResponse() As String
    CommandResponse = csCommandResponse
End Property
Public Property Let CommandResponse(sCommandResponse As String)
    csCommandResponse = sCommandResponse
End Property

Public Property Get DataFilePath() As String
    DataFilePath = csDataFilePath
End Property
Public Property Let DataFilePath(sDataFilePath As String)
    csDataFilePath = sDataFilePath
End Property


Public Property Get FormatFile() As String
    FormatFile = csFormatFile
End Property
Public Property Let FormatFile(sFormatFile As String)
    csFormatFile = sFormatFile
End Property


Public Property Get Hint() As String
    Hint = csHint
End Property
Public Property Let Hint(sHint As String)
    csHint = sHint
End Property


Public Property Get BCPCmdOutFilePath() As String       ' - o switch
    BCPCmdOutFilePath = csBCPCmdOutFilePath
End Property
Public Property Let BCPCmdOutFilePath(sBCPCmdOutFilePath As String)
    csBCPCmdOutFilePath = sBCPCmdOutFilePath
End Property


Public Property Get LastCmd() As String
    LastCmd = csLastCmd
End Property



Public Property Get Owner() As String
    If csOwner = "" Then csOwner = "dbo"
    Owner = csOwner
End Property
Public Property Let Owner(sOwner As String)
    csOwner = sOwner
End Property



Public Property Get Server() As String      ' - S Switch
    Server = csServer
End Property
Public Property Let Server(sServer As String)
    csServer = sServer
End Property


Public Property Get RowTerminator() As String   ' - r switch
    RowTerminator = csRowTerminator
End Property
Public Property Let RowTerminator(sRowTerminator As String)
    csRowTerminator = sRowTerminator
End Property


Public Property Get FieldTerminator() As String     ' - t switch
    FieldTerminator = csFieldTerminator
End Property
Public Property Let FieldTerminator(sFieldTerminator As String)
    csFieldTerminator = sFieldTerminator
End Property



Public Property Get DatabaseName() As String        ' -d switch
    DatabaseName = csDatabaseName
End Property
Public Property Let DatabaseName(sDatabaseName As String)
    csDatabaseName = sDatabaseName
End Property



Public Property Get InTable() As String
    InTable = csInTable
End Property
Public Property Let InTable(sInTable As String)
    csInTable = sInTable
End Property



Public Property Get SqlStatement() As String

    If csSqlStatement = "" Then
        csSqlStatement = coRs.Source
    End If

    '' Make sure all paths and queries are wrapped in double quotes
    If left(csSqlStatement, 1) <> """" Then csSqlStatement = """" & csSqlStatement & """"

    SqlStatement = csSqlStatement
End Property
Public Property Let SqlStatement(sSqlStatement As String)
    csSqlStatement = sSqlStatement
End Property


Public Property Get DeleteFilesAfterObjectDestroyed() As Boolean
    DeleteFilesAfterObjectDestroyed = cblnSelfCleanupFiles
End Property
Public Property Let DeleteFilesAfterObjectDestroyed(blnSelfCleanupFiles As Boolean)
    cblnSelfCleanupFiles = blnSelfCleanupFiles
End Property



Public Property Get FullyQualifiedInObject() As String
Dim sWork As String

'    If Me.Server = "" Then GoTo Block_Exit
    If Me.DatabaseName = "" Then GoTo Block_Exit
    If Me.Owner = "" Then GoTo Block_Exit
    If Me.InTable = "" Then GoTo Block_Exit

    sWork = QuoteIfWhiteSpace(Me.DatabaseName) & "." & _
        QuoteIfWhiteSpace(Me.Owner) & "." & _
        QuoteIfWhiteSpace(Me.InTable)

Block_Exit:
    FullyQualifiedInObject = sWork
    Exit Property
End Property


''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################



''' ############################################################
''' ############################################################
''' ############################################################
'''
'''
'''
Public Function ImportOutFile() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sCmd As String
Dim sSql As String
Dim sSwitches As String
Dim sImportDirectory As String


    strProcName = ClassName & ".ImportOutFile"
    gbImporting = True

    If Me.DataFilePath = "" Then GoTo Block_Exit
    If FileExists(Me.DataFilePath) = False Then GoTo Block_Exit

    sImportDirectory = ParentFolderPath(Me.DataFilePath)
    If FolderExists(sImportDirectory) = False Then GoTo Block_Exit

        '' Get our switches
    sSwitches = BuildOptionSwitches

    sCmd = "bcp " & FullyQualifiedInObject & " in " & _
        QuoteIfWhiteSpace(Me.DataFilePath) & " " & _
        sSwitches

    csLastCmd = sCmd
        Debug.Print sCmd

    ShellWait sCmd

    Call ReadBCPCmdOutputFile
    ImportOutFile = True

Block_Exit:
    gbImporting = False

    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function



''' ############################################################
''' ############################################################
''' ############################################################
'''
'''
'''
Public Function ExportToFile(sOutDirectory As String, Optional ByVal sNameDesc As String, Optional bCreateFormatFile As Boolean = True) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sCmd As String
Dim sSql As String
Dim sUniqueNess As String
Dim sSwitches As String


    strProcName = ClassName & ".ExportToFile"
    gbImporting = False

    sOutDirectory = QualifyFldrPath(sOutDirectory)
    If FolderExists(sOutDirectory) = False Then
        '' create the directory
        If CreateFolders(sOutDirectory) = False Then GoTo Block_Exit
    End If


    sUniqueNess = "_" & Format(Now(), "yyyymmddhhnnss")

    If sNameDesc = "" Then sNameDesc = "bcp_query_out"


    Me.DataFilePath = sOutDirectory & sNameDesc & sUniqueNess & ".txt"

    ccolCreatedFiles.Add Me.DataFilePath


    sSql = SqlStatement

    If sSql = "" Then
        GoTo Block_Exit
    End If

    If Me.CaptureCmdOut = True Then
        Me.BCPCmdOutFilePath = sOutDirectory & sNameDesc & sUniqueNess & "_CMDOUTPUT.txt"
        ccolCreatedFiles.Add Me.BCPCmdOutFilePath
    End If

        '' Get our switches
    sSwitches = BuildOptionSwitches()


    sCmd = "bcp " & sSql & " queryout " & _
        QuoteIfWhiteSpace(Me.DataFilePath) & _
        sSwitches


    csLastCmd = sCmd
            Debug.Print sCmd

    ShellWait sCmd


    Call ReadBCPCmdOutputFile

    If FileExists(Me.DataFilePath) = True Then
        ExportToFile = Me.DataFilePath
    Else
        ExportToFile = ""
    End If

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    ExportToFile = ""
    GoTo Block_Exit
End Function



''' ############################################################
''' ############################################################
''' ############################################################
'''
'''
'''
Private Function BuildOptionSwitches() As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sRet As String
Dim sWork As String

    strProcName = ClassName & ".BuildOptionSwitches"

    With Me
        If .PackageSize <> 0 Then sRet = sRet & " -a" & CStr(.PackageSize) & " "

        If .BatchSize <> 0 Then sRet = sRet & " -b" & CStr(.BatchSize) & " "

            '' The -d database name option is not supported when a 3 part dbtable name is specified.
        If gbImporting = True Then
            If InStr(1, Me.FullyQualifiedInObject, ".", vbBinaryCompare) < 2 Then
                If .DatabaseName <> "" Then sRet = sRet & " -d" & .DatabaseName & " "
            End If
        Else
            If .DatabaseName <> "" Then sRet = sRet & " -d" & .DatabaseName & " "
        End If



        ' -e switch (error file)
        ' -E switch (identity insert on)
        If .IdentityInsert = True Then sRet = sRet & " -E "

        ' -f Format file..
        If Me.FormatFile <> "" Then
            sWork = Me.FormatFile
            If left(sWork, 1) <> """" Then sWork = """" & sWork & """"
            sRet = sRet & " -f " & sWork & " "
            sWork = ""
        End If
        ' -F first row..

        ' -h hint
        If Me.Hint <> "" Then sRet = sRet & " " & Me.Hint & " "

        ' -i Input_file (responses to command prompt questions for interactive mode
        ' -k : retain nulls
        sRet = sRet & " -k "
        ' -K application_intent

        ' -L last row

        ' -m max errors
        Call InsureOneTypeIsSpecified

        ' -c = char data
        If UseCharData = True Then sRet = sRet & " -c "

        ' -n use native data types
        If Me.UseNativeData = True Then sRet = sRet & " -n "

        ' -N use unicode for characar data and native type for non character types
        If Me.UseUnicodeAndNative = True Then sRet = sRet & " -N "

        ' -w Use Unicode
        If .UseUnicode = True Then sRet = sRet & " -w "

        ' -o output_file cmd output is piped to this file
        If .BCPCmdOutFilePath <> "" Then
            sWork = .BCPCmdOutFilePath
            If left(sWork, 1) <> """" Then sWork = """" & sWork & """"
            sRet = sRet & " -o" & sWork & " "
            sWork = ""
        End If

        ' -P passeord

        ' -q ' quoted identifiers on
        If UseQuotedIdentifiers = True Then sRet = sRet & " -q "

        ' -r row_terminator
        ' -R use regional format for dates currency, etc

        ' -S servername:
        If .Server <> "" Then sRet = sRet & " -S" & .Server & " "


        ' -t Field terminator

        ' -T trusted connection (integrated security)
        sRet = sRet & " -T "

        ' -U login id
        ' -v BCP version



    End With


Block_Exit:
    BuildOptionSwitches = sRet
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function



''' ############################################################
''' ############################################################
''' ############################################################
'''
''' If one of the below aren't set then BCP is in "interactive" mode
''' - waiting for input from you (unless you supply an "answer" file
'''
Private Sub InsureOneTypeIsSpecified()

    ' -c = char data
    If cblnUseCharData = False And _
            cblnUseNativeData = False And _
            cblnUseUnicode = False And _
            cblnUseUnicodeAndNative = False Then
        Me.UseCharData = True
    End If

End Sub



''' ############################################################
''' ############################################################
''' ############################################################
'''
'''
'''
Private Function ReadBCPCmdOutputFile() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Dim oTxt As Scripting.TextStream
Dim sMsg As String

    strProcName = ClassName & ".ReadBCPCmdOutputFile"

    If Me.CaptureCmdOut = False Then GoTo Block_Exit

    If FileExists(Me.BCPCmdOutFilePath) = True Then
        Set oFso = New Scripting.FileSystemObject
        If oFso.GetFile(Me.BCPCmdOutFilePath).Size > 0 Then
            Set oTxt = oFso.OpenTextFile(Me.BCPCmdOutFilePath)
            sMsg = oTxt.ReadAll
            oTxt.Close
        End If
    End If

    CommandResponse = sMsg


Block_Exit:
    Set oTxt = Nothing
    Set oFso = Nothing
    ReadBCPCmdOutputFile = IIf(sMsg = "", False, True)
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function


''' ############################################################
''' ############################################################
''' ############################################################
'''
'''
'''
Private Function UniquenessForFileName() As String
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".UniquenessForFileName"

    UniquenessForFileName = "_" & Format(Now(), "yyyymmddhhnnss")


Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function




''' ############################################################
''' ############################################################
''' ############################################################
'''
'''
'''
Private Function QuoteIfWhiteSpace(ByVal sIn As String) As String

    If InStr(1, sIn, " ", vbTextCompare) > 0 Then
        sIn = """" & sIn & """"
        GoTo Block_Exit
    End If
    If InStr(1, sIn, vbTab, vbTextCompare) > 0 Then
        sIn = """" & sIn & """"
        GoTo Block_Exit
    End If


Block_Exit:
    QuoteIfWhiteSpace = sIn
    Exit Function
End Function


''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
'''
'''
'''
Public Function CleanUpCreatedFiles() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim vFile As Variant

    strProcName = ClassName & ".CleanUpCreatedFiles"

    For Each vFile In ccolCreatedFiles
        If FileExists(CStr(vFile)) Then
            If DeleteFile(CStr(vFile), False) = True Then
                CleanUpCreatedFiles = True
            End If
        End If
    Next

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function



''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################


Private Sub Class_Initialize()
    Set ccolCreatedFiles = New Collection
    '' set default cblnSelfCleanupFiles to true
    cblnSelfCleanupFiles = True
End Sub

Private Sub Class_Terminate()
    If DeleteFilesAfterObjectDestroyed = True Then
        Call CleanUpCreatedFiles
    End If
    Set ccolCreatedFiles = Nothing
    Set coRs = Nothing
End Sub
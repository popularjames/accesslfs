Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' DLC 10/25/2011 - Added ability to export to Excel (XLSX format)

'-------- START EVENTS ----------------'
    Public Event ExportStarted()
    Public Event ExportPercentComplete(ByVal intPctComplete As Integer)
    Public Event ExportMessage(ByVal Msg As String)
    Public Event ExportComplete()
    Public Event exportError(ByVal ErrMsg As String)
'-------- END EVENTS ------------------'

'-------- START ENUMERATIONS ----------'
    Public Enum CPExportFormat
        Unspecified
        TextCSV
        TextTab
        TextDelimited
        Xml
        Xls
    End Enum
    
    Public Enum XMLFormats
        Unspecified
        elements
        Attributes
        Raw
    End Enum
'-------- END ENUMERATIONS ------------'

'-------- START VARIABLES -------------'
    '-------- START PRIVATE VARIABLES -----'
        Private genUtils As New CT_ClsGeneralUtilities
        Private iRecordCount As Long
        Private iCurrentRecord As Long
        Private bCancel As Boolean
    '-------- END PRIVATE VARIABLES -------'
    
    '-------- START PROPERTY VARIABLES ----'
        Private lHwnd As Long
        Private eExportFormat As CPExportFormat
        Private sSourceQuery As String
        Private sOutputFileName As String
        Private bAddFieldHeaders As Boolean
        Private sFieldDelimiter As String
        Private sRecordDelimiter As String
        Private sTextQualifier As String
        Private bAppendToFile As Boolean
        Private eXMLFormat As XMLFormats
    '-------- END PROPERTY VARIABLES ------'
'-------- END VARIABLES ---------------'

'-------- START PROPERTIES ------------'
    '-------- START MANDATORY PROPERTIES ------------'
        Public Property Let ExportFormat(ByVal fmt As CPExportFormat)
            eExportFormat = fmt
            Initialize
        End Property
        Public Property Get ExportFormat() As CPExportFormat
            ExportFormat = eExportFormat
        End Property

        Public Property Let SourceQuery(ByVal Qry As String)
            sSourceQuery = Qry
        End Property
        Public Property Get SourceQuery() As String
            SourceQuery = sSourceQuery
        End Property
        
        Public Property Let hwnd(ByRef handle As Long)
            lHwnd = handle
        End Property
        Public Property Get hwnd() As Long
            hwnd = lHwnd
        End Property
    '-------- END MANDATORY PROPERTIES --------------'
    
    '-------- START OPTIONAL PROPERTIES -------------'
        Public Property Let OutputFileName(ByVal outName As String)
            sOutputFileName = outName
        End Property
        Public Property Get OutputFileName() As String
            OutputFileName = sOutputFileName
        End Property
        
        Public Property Let AddFieldHeaders(ByVal addHeaders As Boolean)
            bAddFieldHeaders = addHeaders
        End Property
        Public Property Get AddFieldHeaders() As Boolean
            AddFieldHeaders = bAddFieldHeaders
        End Property
        
        Public Property Let FieldDelimiter(ByVal delimiter As String)
            sFieldDelimiter = delimiter
        End Property
        Public Property Get FieldDelimiter() As String
            FieldDelimiter = sFieldDelimiter
        End Property
        
        Public Property Let RecordDelimiter(ByVal delimiter As String)
            sRecordDelimiter = delimiter
        End Property
        Public Property Get RecordDelimiter() As String
            RecordDelimiter = sRecordDelimiter
        End Property
        
        Public Property Let TextQualifier(ByVal qualifier As String)
            sTextQualifier = qualifier
        End Property
        Public Property Get TextQualifier() As String
            TextQualifier = sTextQualifier
        End Property
        
        Public Property Let AppendToFile(ByVal Append As Boolean)
            bAppendToFile = Append
        End Property
        Public Property Get AppendToFile() As Boolean
            AppendToFile = bAppendToFile
        End Property
        
        Public Property Let XMLFormat(fmt As XMLFormats)
            eXMLFormat = fmt
        End Property
        Public Property Get XMLFormat() As XMLFormats
            XMLFormat = eXMLFormat
        End Property
    '-------- END OPTIONAL PROPERTIES ---------------'
'-------- END PROPERTIES --------------'

'-------- START PUBLIC FUNCTIONS ------'
    Public Function GetExportFormatStrings() As String()
        Dim Result As String
        
        ' Create CSV of Enums
        Result = "CSV=Text - Comma Separated Values,TAB=Text - Tab Delimited,DEL=Text - Custom Delimited,XML=XML Document,XLS=Excel Spreadsheet"
        
        ' Return split (string array) of enums
        GetExportFormatStrings = Split(Result, ",")
    End Function
    
    Public Function GetExportFormatEnum(ByVal exportType As String) As CPExportFormat
        Select Case UCase(exportType)
            Case "CSV"
                GetExportFormatEnum = TextCSV
            Case "TAB"
                GetExportFormatEnum = TextDelimited
            Case "DEL"
                GetExportFormatEnum = TextDelimited
            Case "XML"
                GetExportFormatEnum = Xml
            Case "XLS"
                GetExportFormatEnum = Xls
            Case Else
                GetExportFormatEnum = CPExportFormat.Unspecified
        End Select
    End Function
    
    Public Function GetXMLFormatStrings() As String()
        Dim Result As String
        
        Result = "ELE=Elements - Element Centric Mapping,ATR=Attributes - Attribute Centric Mapping,RAW=Raw - Raw XML Format"
        
         GetXMLFormatStrings = Split(Result, ",")
    End Function
    
    Public Function GetXMLFormatEnum(ByVal xmlType As String) As XMLFormats
        Select Case UCase(xmlType)
            Case "ELE"
                GetXMLFormatEnum = elements
            Case "ATR"
                GetXMLFormatEnum = Attributes
            Case "RAW"
                GetXMLFormatEnum = Raw
            Case Else
                GetXMLFormatEnum = XMLFormats.Unspecified
        End Select
    End Function
    
    Public Function GetCurrentExportFormat() As String
        Dim Result As String
        
        Select Case eExportFormat
            Case CPExportFormat.TextDelimited
                Result = "Text - Delimited"
            Case CPExportFormat.TextCSV
                Result = "Text - Comma Separated Values (CSV)"
            Case CPExportFormat.TextTab
                Result = "Text - Tab Delimited"
            Case CPExportFormat.Xml
                Result = "XML Document"
            Case Else
                Result = "Unspecified"
        End Select
    End Function
    
    Public Function GetCurrentExportFormatEnum() As CPExportFormat
        Dim Result As CPExportFormat
        
        Select Case eExportFormat
            Case CPExportFormat.TextDelimited
                Result = TextDelimited
            Case CPExportFormat.TextCSV
                Result = TextCSV
            Case CPExportFormat.TextTab
                Result = TextTab
            Case CPExportFormat.Xml
                Result = Xml
            Case Else
                Result = CPExportFormat.Unspecified
        End Select
        
        GetCurrentExportFormatEnum = Result
    End Function
    
    Public Function GetCurrentXMLFormat() As String
        Dim Result As String
        
        Select Case eXMLFormat
            Case XMLFormats.Attributes
                Result = "Attribute Centric Mapping"
            Case XMLFormats.elements
                Result = "Element Centric Mapping"
            Case XMLFormats.Raw
                Result = "Raw XML"
            Case Else
                Result = "Unspecified"
    End Select
    End Function
'-------- END PUBLIC FUNCTIONS ------'

'-------- START PUBLIC SUBS ---------'
    Public Sub Export()
    On Error GoTo ErrorHandler
        Dim db As DAO.Database
        Dim rs As DAO.RecordSet
        Dim ioMode As Byte
        Dim fso
        Dim OutFile
        Dim currPct
        Dim lastPct
        
        bCancel = False
        lastPct = -1
        
        RaiseEvent ExportStarted
        
        ' Validate minimum requirements
        RaiseEvent ExportMessage("Validating Export Parameters")
        DoEvents
        
        If Not ValidateExportProperties Then
            Exit Sub
        End If

        RaiseEvent ExportMessage("Opening recordset")
        DoEvents
        
        Set db = CurrentDb
        Set rs = db.OpenRecordSet(sSourceQuery, dbOpenSnapshot)
        
        RaiseEvent ExportMessage("Acquiring Total Record Count..." & vbCrLf & "This may take several minutes depending on number of fields and records.")
        DoEvents
        
        rs.MoveLast
        rs.MoveFirst
        
        ' Set the I/O Mode (Write or Append)
        ioMode = 2
        If bAppendToFile Then
            ioMode = 8
        End If
        
        ' Create the file objects
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set OutFile = fso.OpenTextFile(sOutputFileName, ioMode, True, 0)
        
        ' Aquire record count, initialize record counter
        iRecordCount = rs.recordCount
        iCurrentRecord = 1
        
        ' Determine if we will be adding field headers
        If bAddFieldHeaders And eExportFormat <> CPExportFormat.Xml Then
            RaiseEvent ExportMessage("Export Field Headers")
            DoEvents
            
            OutFile.Write CreateExportRecordHeader(rs)
        End If
        
        ' Determine if XML format
        If eExportFormat = CPExportFormat.Xml And Not bCancel Then
            ' Write XML Declaration and root node
            OutFile.WriteLine "<?xml version=""1.0""  encoding=""windows-1252""?>"
            OutFile.WriteLine "<Export>"
        End If
        
        ' Loop until the end of recordset or user termination
        While Not rs.EOF And Not bCancel
            ' Calculate current percent complete
            currPct = CInt((iCurrentRecord / iRecordCount) * 100)
            
            ' If percent complete has changed
            If currPct > lastPct Then
                ' Notify caller of percent change
                RaiseEvent ExportPercentComplete(currPct)
                DoEvents
                
                ' Set last percent complete
                lastPct = currPct
            End If
            
            ' Give user feedback
            RaiseEvent ExportMessage("Exporting Record " & iCurrentRecord & " of " & iRecordCount)
            DoEvents
            
            ' Create export record
            OutFile.Write CreateExportRecord(rs)
            
            ' Move to next record
            rs.MoveNext
            
            ' Increment record count
            iCurrentRecord = iCurrentRecord + 1
        Wend
        
        ' Determine if XML format
        If eExportFormat = CPExportFormat.Xml And Not bCancel Then
            ' Close root node
            OutFile.WriteLine "</Export>"
        End If
        
        ' Notify user of export complete
        RaiseEvent ExportComplete
                
ErrorHandlerExit:
    On Error Resume Next
        
        ' Determine if user cancelled process
        If bCancel Then
            ' Raise error
            RaiseEvent exportError("Export Terminated By User")
            DoEvents
        End If
        
        ' Close recordset, destory object
        rs.Close
        Set rs = Nothing
        
        ' Close output file, destory object, destory FSO
        OutFile.Close
        Set OutFile = Nothing
        Set fso = Nothing
    
        Exit Sub
ErrorHandler:
        ' Notify user of error
        RaiseEvent exportError(Err.Number & ": " & Err.Description)
        DoEvents
        
        Resume ErrorHandlerExit
    End Sub
    
    Public Sub Cancel()
        ' Set cancel to true
        bCancel = True
        DoEvents
    End Sub
'-------- END PUBLIC SUBS -----------'

'-------- START PRIVATE FUNCTIONS ---'
    Private Function ValidateExportProperties() As Boolean
        Dim filter As String
        
        ' Make sure that the hwnd is specified
        If lHwnd = 0 Then
            RaiseEvent exportError("No Hwnd Specified")
            DoEvents
            
            ValidateExportProperties = False
            Exit Function
        End If
        
        ' Make sure that an export format is specified
        If eExportFormat = 0 Then
            RaiseEvent exportError("No Export Format Specified")
            DoEvents
            ValidateExportProperties = False
            Exit Function
        End If
        
        If eExportFormat = Xml And eXMLFormat = 0 Then
            RaiseEvent exportError("No XML Format Specified")
            DoEvents
            ValidateExportProperties = False
            Exit Function
        End If
        
        ' Make sure that a source query is specified
        If sSourceQuery = vbNullString Then
            RaiseEvent exportError("No Source Query Specified")
            DoEvents
            ValidateExportProperties = False
            Exit Function
        End If
        
        ' Determine if a OutputFileName was specified
        If sOutputFileName = vbNullString Then
            ' No output file specified, determine what file extension to use based
            ' on the export format
            Select Case eExportFormat
                Case CPExportFormat.TextCSV
                    filter = "Comma Separated Values (*.csv)" & Chr(0) & "*.csv" & Chr(0)
                Case CPExportFormat.Xml
                    filter = "XML Document (*.xml)" & Chr(0) & "*.xml" & Chr(0)
                Case Else
                    filter = "Text File (*.txt)" & Chr(0) & "*.txt" & Chr(0)
            End Select
        
            ' Open file dialog
            sOutputFileName = FileDialog(1, "Save Export File As", Me.hwnd, "C:\", filter)
            
            ' Test if user did not specify a file name
            If sOutputFileName = vbNullString Then
                ' Notify user of error
                RaiseEvent exportError("No Output File Name Specified")
                DoEvents
                
                ValidateExportProperties = False
                Exit Function
            End If
        End If
        
        ' Test for delimited format
        If eExportFormat = TextDelimited Or eExportFormat = TextCSV Or eExportFormat = TextTab Then
            ' Make sure that there is a field delimiter specified
            If sFieldDelimiter = vbNullString Then
                ' Notify user of error
                RaiseEvent exportError("No Field Delimiter Specified")
                DoEvents
                
                ValidateExportProperties = False
                Exit Function
            End If
            
            ' Make sure that there is a record delimiter
            If sRecordDelimiter = vbNullString Then
                ' Notify user of error
                RaiseEvent exportError("No Record Delimiter Specified")
                DoEvents
                
                ValidateExportProperties = False
                Exit Function
            End If
        End If
        
        ' Validated
        ValidateExportProperties = True
    End Function
    
    Private Function CreateExportRecordHeader(ByRef rs As DAO.RecordSet) As String
        Dim Result As String
        Dim ctr As Integer
        
        Result = ""
        
        ' Loop through the fields
        For ctr = 0 To rs.Fields.Count - 1
            ' Add field name and field delimiter to result string
             Result = Result & rs.Fields(ctr).Name & sFieldDelimiter
            
            ' Make sure not cancelled
            If bCancel Then
                ' Cancelled, step out
                Exit For
            End If
        Next
        
        ' Drop trailing field delimiter (if necessary)
        If Len(Result) > 0 Then
            Result = left(Result, Len(Result) - Len(sFieldDelimiter))
        Else
            Result = Result
        End If
        
        ' Return value with record delimiter
        CreateExportRecordHeader = Result & sRecordDelimiter
    End Function

    Private Function CreateExportRecord(ByRef rs As DAO.RecordSet) As String
        Dim Result As String
        Dim FieldValue As String
        Dim ctr As Integer
        Dim fld As DAO.Field
        
        Result = ""
        
        ' Test if export format is XML
        If eExportFormat = CPExportFormat.Xml Then
            ' Determine which XML format to create
            Select Case eXMLFormat
                Case XMLFormats.Attributes
                    Result = Result & vbTab & "<Export " & vbCrLf
                Case XMLFormats.elements
                    Result = Result & vbTab & "<Export>" & vbCrLf
                Case XMLFormats.Raw
                    Result = Result & vbTab & "<row "
            End Select
        End If
        
        ' Loop through the fields
        For ctr = 0 To rs.Fields.Count - 1
            ' Set the field variable (makes my life easier)
            Set fld = rs.Fields(ctr)
            
            ' Aquire the field value (always use NZ, because the field value can be null)
            FieldValue = Nz(rs.Fields(ctr).Value, "")
            
            ' Look for a GUID
            If left(FieldValue, 7) = "{guid {" Then
                ' Convert GUID to a readable value
                FieldValue = genUtils.ConvertGuid(FieldValue)
            End If
            
            ' Determine if the record delimiter is a CRLF
            If sRecordDelimiter = vbCrLf Then
                ' Replace any CRLF's in the field with just LF
                FieldValue = Replace(FieldValue, vbCrLf, vbLf)
            End If
            
            ' Determine the export format
            Select Case eExportFormat
                ' Delimited format
                Case CPExportFormat.TextCSV, CPExportFormat.TextDelimited, CPExportFormat.TextTab
                    ' Determine if this is a text type field
                    If IsTextField(rs.Fields(ctr).Type) Then
                        ' Add the text qualifier
                        Result = Result & sTextQualifier & Replace(FieldValue, sTextQualifier, sTextQualifier & sTextQualifier) & sTextQualifier & sFieldDelimiter
                    Else
                        ' Just add the field value
                        Result = Result & FieldValue & sFieldDelimiter
                    End If
                    
                ' XML Format
                Case CPExportFormat.Xml
                    ' Determine if we are using elements
                    If eXMLFormat = elements Then
                        ' Create a new element (make sure that the field name is a legal XML name, if not correct it)
                        ' HC 5/2010 - replaced formatforxmldata with general utilities item xmlencode
                        Result = Result & vbTab & vbTab & "<" & XML_LegalName(rs.Fields(ctr).Name) & ">" & genUtils.XMLEncode(FieldValue) & "</" & XML_LegalName(rs.Fields(ctr).Name) & ">" & vbCrLf
                    Else
                        ' Create an attribute (make sure that the field name is a legal XML name, if not correct it)
                        ' HC 5/2010 - replaced formatforxmldata with general utilities item xmlencode
                        Result = Result & vbTab & vbTab & XML_LegalName(rs.Fields(ctr).Name) & "=""" & genUtils.XMLEncode(FieldValue) & """ " & vbCrLf
                    End If
            End Select
            
            ' Make sure that the user has not cancelled the process
            If bCancel Then
                ' User cancelled, step out
                Exit For
            End If
        Next
        
        ' Determine if we are using an XML format
        If eExportFormat = CPExportFormat.Xml Then
            ' Determine XML format
            Select Case eXMLFormat
                ' Attribute centric, close element
                Case XMLFormats.Attributes
                    Result = Result & "/>" & vbCrLf
                ' Element centric, close the Export element
                Case XMLFormats.elements
                    Result = Result & vbTab & "</Export>" & vbCrLf
                ' Raw format, close the element
                Case XMLFormats.Raw
                    Result = Result & "/>" & vbCrLf
            End Select
            
            ' Set return value
            CreateExportRecord = Result
        Else
            ' Drop trailing delimiter (if necessary)
            If Len(Result) > 0 Then
                Result = left(Result, Len(Result) - Len(sFieldDelimiter))
            Else
                Result = Result
            End If
            
            ' Set return value
            CreateExportRecord = Result & sRecordDelimiter
        End If
    End Function
    
    Private Function XML_LegalName(ByVal theName As String) As String
        Dim Result As String
        Dim tstChr As String
        Dim iPos As Integer
    
        ' Loop through the characters
        For iPos = 1 To Len(theName)
            ' Aquire the character at current position
            tstChr = UCase(Mid(theName, iPos, 1))
            
            ' Make sure that the character is in the valid range (0-9), (A-Z)
            If (((Asc(tstChr) > 47) And (Asc(tstChr) < 59)) Or ((Asc(tstChr) > 64) And (Asc(tstChr) < 91))) Then
                Result = Result & Mid(theName, iPos, 1)
            End If
        Next iPos
        
        ' Aquire the first character
        tstChr = left(Result, 1) 'Mid(result, 1, 1)
        
        ' Test if first characeter is a number
        If ((Asc(tstChr) > 47) And (Asc(tstChr) < 59)) Then
            ' It is a number, prefix name with an "x" (per XML specification)
            Result = "x" & Result
        End If
        
        ' Return result
        XML_LegalName = Result
    End Function
        
    Private Function IsTextField(typ As Integer) As Boolean
        ' Test the type of field
        If typ = dbText Or _
            typ = dbMemo Or _
            typ = dbGUID Or _
            typ = dbChar Then
            
            ' This is a text field
            IsTextField = True
        Else
            ' This is not a text field
            IsTextField = False
        End If
    End Function
    '-------- END PRIVATE FUNCTIONS -----'
    '-------- START PRIVATE SUBS --------'
        Private Sub Initialize()
            ' Initialize the field headers and append flags
            bAddFieldHeaders = False
            bAppendToFile = False

            ' Determine the export format
            Select Case eExportFormat
                ' CSV format
                Case CPExportFormat.TextCSV
                    ' Set the field delimiter
                    sFieldDelimiter = ","
                    ' Set the record delimiter
                    sRecordDelimiter = vbCrLf
                    ' Set the text qualifier
                    sTextQualifier = Chr(34)
                
                ' TAB delimited format
                Case CPExportFormat.TextTab
                    ' Set the field delimiter
                    sFieldDelimiter = vbTab
                    ' Set the record delimiter
                    sRecordDelimiter = vbCrLf
                    ' Set the text qualifier
                    sTextQualifier = Chr(34)
                
                ' Other type
                Case Else
                    ' Clear the field delimiter
                    sFieldDelimiter = vbNullString
                    ' Clear the record delimiter
                    sRecordDelimiter = vbNullString
                    ' Set the text qualifier
                    sTextQualifier = Chr(34)
            End Select
        End Sub
    '-------- END PRIVATE SUBS ----------'
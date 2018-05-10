Option Compare Database
Option Explicit

' This is mostly code I've taken from the internet to deal with
' large binary files
' -- Kevin D. Dearing


Const BlockSize = 32768
'Const BlockSize = 16384

'Const acSysCmdInitMeter = SYSCMD_INITMETER
'Const acSysCmdUpdateMeter = SYSCMD_UPDATEMETER
'Const acSysCmdRemoveMeter = SYSCMD_REMOVEMETER

Private Const ClassName As String = "mod_Blobs"

      '**************************************************************
      ' FUNCTION: ReadBLOB()
      '
      ' PURPOSE:
      '   Reads a BLOB from a disk file and stores the contents in the
      '   specified table and field.
      '
      ' PREREQUISITES:
      '   The specified table with the OLE object field to contain the
      '   binary data must be opened in Visual Basic code (Access Basic
      '   code in Microsoft Access 2.0 and earlier) and the correct record
      '   navigated to prior to calling the ReadBLOB() function.
      '
      ' ARGUMENTS:
      '   Source - The path and filename of the binary information
      '            to be read and stored.
      '   T      - The table object to store the data in.
      '   Field  - The OLE object field in table T to store the data in.
      '
      ' RETURN:
      '   The number of bytes read from the Source file.
      '**************************************************************
Function ReadBLOB(Source As String, T As DAO.RecordSet, sField As String, Optional blnAddNew As Boolean = True, Optional varOtherFieldsNValues As Variant) As Boolean
On Error GoTo Funct_Err
Dim strProcName As String
Dim NumBlocks As Integer, SourceFile As Integer, i As Integer
Dim FileLength As Long, LeftOver As Long
Dim FileData As String
Dim retval As Variant
Dim iLoop As Integer

'    On Error GoTo Err_ReadBLOB
    strProcName = ClassName & ".ReadBLOB"
    ReadBLOB = True
    
    ' Open the source file.
    SourceFile = FreeFile
    Open Source For Binary Access Read As SourceFile

    ' Get the length of the file.
    FileLength = LOF(SourceFile)
    If FileLength = 0 Then
        ReadBLOB = 0
        Exit Function
    End If

    ' Calculate the number of blocks to read and leftover bytes.
    NumBlocks = FileLength \ BlockSize
    LeftOver = FileLength Mod BlockSize

    ' SysCmd is used to manipulate status bar meter.
    retval = SysCmd(acSysCmdInitMeter, "Reading BLOB", _
             FileLength \ 1000)

    ' Put the record in edit mode.
    If blnAddNew = True Then
        T.AddNew
    Else
        T.Edit
    End If
 
    ' Read the leftover data, writing it to the table.
    FileData = String$(LeftOver, 32)
    Get SourceFile, , FileData
    T(sField).AppendChunk (FileData)

    retval = SysCmd(acSysCmdUpdateMeter, LeftOver / 1000)

    ' Read the remaining blocks of data, writing them to the table.
    FileData = String$(BlockSize, 32)
    For i = 1 To NumBlocks
        Get SourceFile, , FileData
        T(sField).AppendChunk (FileData)

        retval = SysCmd(acSysCmdUpdateMeter, BlockSize * i / 1000)
    Next i

    ' Update the record and terminate function.
    If IsArray(varOtherFieldsNValues) = True Then
        If Not IsEmpty(varOtherFieldsNValues) = True Then
                ' Only use if there are an even number of items in the array...
            If (UBound(varOtherFieldsNValues) + 1) Mod 2 = 0 Then
                For iLoop = 0 To UBound(varOtherFieldsNValues) Step 2
                    T(varOtherFieldsNValues(iLoop)).Value = varOtherFieldsNValues(iLoop + 1)
                Next
            End If
        End If
    End If
    
    
    T.Update
    retval = SysCmd(acSysCmdRemoveMeter)
    Close SourceFile
'    ReadBLOB = FileLength
    
Funct_Exit:
    Exit Function

Funct_Err:
    ReadBLOB = False
    ReportError Err, strProcName
'    ReadBLOB = -Err
    Resume Funct_Exit
End Function

      '**************************************************************
      ' FUNCTION: WriteBLOB()
      '
      ' PURPOSE:
      '   Writes BLOB information stored in the specified table and field
      '   to the specified disk file.
      '
      ' PREREQUISITES:
      '   The specified table with the OLE object field containing the
      '   binary data must be opened in Visual Basic code (Access Basic
      '   code in Microsoft Access 2.0 or earlier) and the correct
      '   record navigated to prior to calling the WriteBLOB() function.
      '
      ' ARGUMENTS:
      '   T           - The table object containing the binary information.
      '   sField      - The OLE object field in table T containing the
      '                 binary information to write.
      '   Destination - The path and filename to write the binary
      '                 information to.
      '
      ' RETURN:
      '   The number of bytes written to the destination file.
      '**************************************************************
Function WriteBLOB(T As DAO.RecordSet, sField As String, Destination As String)
Dim NumBlocks As Integer, DestFile As Integer, i As Integer
Dim FileLength As Long, LeftOver As Long
Dim FileData As String
Dim retval As Variant

    On Error GoTo Err_WriteBLOB

    ' Get the size of the field.
    FileLength = T(sField).FieldSize()
    If FileLength = 0 Then
        WriteBLOB = 0
        Exit Function
    End If

    ' Calculate number of blocks to write and leftover bytes.
    NumBlocks = FileLength \ BlockSize
    LeftOver = FileLength Mod BlockSize

    ' Remove any existing destination file.
    DestFile = FreeFile
    If FileExists(Destination) = True Then
        WriteBLOB = 1234
        Exit Function
    End If
    Open Destination For Output As DestFile
    Close DestFile

    ' Open the destination file.
    Open Destination For Binary As DestFile

    ' SysCmd is used to manipulate the status bar meter.
    retval = SysCmd(acSysCmdInitMeter, "Writing BLOB", FileLength / 1000)

    ' Update the status bar meter.
'    RetVal = SysCmd(acSysCmdUpdateMeter, LeftOver / 1000)

    ' Write the remaining blocks of data to the output file.

Dim lOffSet As Long
    lOffSet = 0
    
    For i = 1 To NumBlocks
        ' Reads a chunk and writes it to output file.
        'FileData = T(sField).GetChunk((i - 1) * BlockSize, BlockSize)
        FileData = T(sField).GetChunk(lOffSet, BlockSize)
        Put DestFile, , FileData

        lOffSet = lOffSet + BlockSize

        retval = SysCmd(acSysCmdUpdateMeter, lOffSet / 1000)
    
    Next i

    ' Write the leftover data to the output file.
    FileData = T(sField).GetChunk(lOffSet, LeftOver)
    Put DestFile, , FileData

    ' Terminates function
    retval = SysCmd(acSysCmdRemoveMeter)
    Close DestFile
    WriteBLOB = FileLength
    Exit Function

Err_WriteBLOB:
    WriteBLOB = -Err
    Exit Function

End Function
'
'Function WriteBLOB(T As Recordset, sField As String, Destination As String)
'Dim NumBlocks As Integer, DestFile As Integer, i As Integer
'Dim FileLength As Long, LeftOver As Long
'Dim FileData As String
'Dim RetVal As Variant
'
'    On Error GoTo Err_WriteBLOB
'
'    ' Get the size of the field.
'    FileLength = T(sField).FieldSize()
'    If FileLength = 0 Then
'        WriteBLOB = 0
'        Exit Function
'    End If
'
'    ' Calculate number of blocks to write and leftover bytes.
'    NumBlocks = FileLength \ BlockSize
'    LeftOver = FileLength Mod BlockSize
'
'    ' Remove any existing destination file.
'    DestFile = FreeFile
'    Open Destination For Output As DestFile
'    Close DestFile
'
'    ' Open the destination file.
'    Open Destination For Binary As DestFile
'
'    ' SysCmd is used to manipulate the status bar meter.
'    RetVal = SysCmd(acSysCmdInitMeter, "Writing BLOB", FileLength / 1000)
'
'    ' Write the leftover data to the output file.
'    FileData = T(sField).GetChunk(0, LeftOver)
'    Put DestFile, , FileData
'
'    ' Update the status bar meter.
'    RetVal = SysCmd(acSysCmdUpdateMeter, LeftOver / 1000)
'
'    ' Write the remaining blocks of data to the output file.
'    For i = 1 To NumBlocks
'        ' Reads a chunk and writes it to output file.
'        FileData = T(sField).GetChunk((i - 1) * BlockSize + LeftOver, BlockSize)
'        Put DestFile, , FileData
'
'        RetVal = SysCmd(acSysCmdUpdateMeter, ((i - 1) * BlockSize + LeftOver) / 1000)
'    Next i
'
'    ' Terminates function
'    RetVal = SysCmd(acSysCmdRemoveMeter)
'    Close DestFile
'    WriteBLOB = FileLength
'    Exit Function
'
'Err_WriteBLOB:
'    WriteBLOB = -Err
'    Exit Function
'
'End Function

      '**************************************************************
      ' SUB: CopyFile
      '
      ' PURPOSE:
      '   Demonstrates how to use ReadBLOB() and WriteBLOB().
      '
      ' PREREQUISITES:
      '   A table called BLOB that contains an OLE Object field called
      '   Blob.
      '
      ' ARGUMENTS:
      '   Source - The path and filename of the information to copy.
      '   Destination - The path and filename of the file to write
      '                 the binary information to.
      '
      ' EXAMPLE:
      '   CopyBlob "c:\windows\winfile.hlp", "c:\windows\winfil_1.hlp"
      '**************************************************************
'''Public Sub CopyBlob(Source As String, Destination As String)
'''Dim BytesRead As Variant, BytesWritten As Variant
'''Dim Msg As String
'''Dim db As Database
'''Dim T As Recordset
'''
'''    ' Open the BLOB table.
'''    Set db = CurrentDb()
'''    Set T = db.OpenRecordset("BLOB", dbOpenTable)
'''
'''    ' Create a new record and move to it.
'''    T.AddNew
'''    T.Update
'''    T.MoveLast
'''
'''    BytesRead = ReadBLOB(Source, T, "Blob")
'''
'''    Msg = "Finished reading """ & Source & """"
'''    Msg = Msg & Chr$(13) & ".. " & BytesRead & " bytes read."
'''    MsgBox Msg, 64, "Copy File"
'''
'''    BytesWritten = WriteBLOB(T, "Blob", Destination)
'''
'''    Msg = "Finished writing """ & Destination & """"
'''    Msg = Msg & Chr$(13) & ".. " & BytesWritten & " bytes written."
'''    MsgBox Msg, 64, "Copy File"
'''End Sub
                

Public Function ExtractDependencies() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim sSql As String
Dim sDestPath As String
Dim bAtLeastOneFileNotFound As Boolean

    strProcName = ClassName & ".ExtractDependencies"
    bAtLeastOneFileNotFound = True

    
    Set oDb = CurrentDb()
    sSql = "SELECT * FROM tbl_App_Dependencies WHERE Active = True "
    
    Set oRs = oDb.OpenRecordSet(sSql)
    
    If oRs.EOF And oRs.BOF Then
        LogMessage "No active dependencies found to extract", strProcName
        GoTo Block_Exit
    End If

    While Not oRs.EOF
        If IsNull(oRs("ExtractPath").Value) Or CStr("" & oRs("ExtractPath").Value) = "" Then
            sDestPath = CurrentDBDir() & "\" & oRs("DependencyName").Value
        Else
            sDestPath = oRs("ExtractPath").Value
        End If
        
'        sDestPath = FixPath(sDestPath, Now())
        If FileExists(sDestPath) = False Then
            WriteBLOB oRs, "DependencyOLE", sDestPath
        End If
        
        If bAtLeastOneFileNotFound = True Then
            bAtLeastOneFileNotFound = FileExists(sDestPath)
        End If
        
        If FileExists(sDestPath) = True Then
            SetAttr sDestPath, vbHidden
        End If
        
        oRs.MoveNext
    Wend

    ExtractDependencies = True
Block_Exit:
    Set oRs = Nothing
    Set oDb = Nothing
    
    Exit Function
    
Block_Err:
    ExtractDependencies = False
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
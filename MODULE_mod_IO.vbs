Option Compare Database
Option Explicit

'' HISTORY:
'' 03/12/2012  KD: Added ClassName and CreateFolder_s_
'' 03/07/2013 - Get_TIFF_COUNT_2 Added


'**************DPR 3/7/2013*********************
Private Const ClassName As String = "mod_IO"


Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Private Const Pathlen = 256
Private Const MaxPath = 256


Private Declare Function GetTempFileName _
    Lib "kernel32" Alias "GetTempFileNameA" _
    (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long
 




Private Type LongType
    lLong As Long
End Type

Private Type IntType
    iInt As Integer
End Type

Private Type FourBytes
    bByte1 As Byte
    bByte2 As Byte
    bByte3 As Byte
    bByte4 As Byte
End Type

Private Type TwoBytes
    bByte1 As Byte
    bByte2 As Byte
End Type
'**************DPR 3/7/2013*********************


Public Function SetFileReadOnly(FileName As String) As Boolean
    Dim fso As New FileSystemObject
    Dim mFile As file
    
    On Error GoTo Err_handler
    
    Set mFile = fso.GetFile(FileName)
    mFile.Attributes = mFile.Attributes Or ReadOnly
    
    SetFileReadOnly = True
    
Exit_Function:
    Set mFile = Nothing
    Set fso = Nothing
    Exit Function
    
Err_handler:
    SetFileReadOnly = False
    Resume Exit_Function
End Function


Public Function CreateFolder(Path) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim fso
    Dim strChkPath
    Dim i, j, iPathLen
   
    Set fso = CreateObject("scripting.filesystemobject")

    iPathLen = Len(Path)
    If iPathLen = 0 Then
        CreateFolder = False
        Exit Function
    End If
    
    If fso.FolderExists(Path) Then
        CreateFolder = True
        Exit Function
    End If
    
    If InStr(1, Path, "\") = 0 Then
        CreateFolder = False
        Exit Function
    End If
        
    If left(Path, 2) = "\\" Then
        j = InStr(3, Path, "\") + 1
    Else
        j = 1
    End If
    
    Do
        i = InStr(j, Path, "\")
        If i > 0 Then
            strChkPath = left(Path, i - 1)
            If Not fso.FolderExists(strChkPath) Then
                '' 20121003 KD Bug fix - fso won't create the directory if it ends with a slash...
                If Right(strChkPath, 1) = "\" Then
                    fso.CreateFolder left(strChkPath, Len(strChkPath) - 1)
                Else
                    Sleep 1000
                    fso.CreateFolder strChkPath
                End If
            End If
            j = i + 1
        Else
            j = iPathLen
        End If
    Loop Until j = iPathLen
    
    If Not fso.FolderExists(Path) Then
        Call fso.CreateFolder(Path)
    End If

CreateFolder = True

Exit Function

ErrHandler:
    CreateFolder = False
End Function


Public Function FolderExist(FolderPath As String) As Boolean
    Dim fso As New FileSystemObject
    
    FolderExist = fso.FolderExists(FolderPath)
End Function


'' KD Added: 20120416
Public Function FileExists(sFilePath As String) As Boolean
    Dim fso As New FileSystemObject
    
    FileExists = fso.FileExists(sFilePath)
    Set fso = Nothing
End Function


'' KD Added: 20120416
Public Function FolderExists(sFolderPath As String) As Boolean
    Dim fso As New FileSystemObject
    
    FolderExists = fso.FolderExists(sFolderPath)
    Set fso = Nothing
End Function


Public Function GetFileName(ByVal FileName As String) As String
    If InStrRev(FileName, "\") Then
        FileName = Mid(FileName, InStrRev(FileName, "\") + 1, Len(FileName))
    End If
    GetFileName = FileName
End Function


'pulled from mod LETTER FileSystem
Public Function DeleteFolder(Path) As String
    Dim fso
    Dim strChkPath
    Dim i, j, iPathLen
   
    Set fso = CreateObject("scripting.filesystemobject")

    iPathLen = Len(Path)
    If iPathLen = 0 Then
        DeleteFolder = "Folder name is empty."
        Exit Function
    End If
    
    
    If Right(Path, 1) = "\" Then
        Path = left(Path, Len(Path) - 1)
    End If
    
    If fso.FolderExists(Path) Then
        DeleteFolder = "Folder Deleted"
        On Error Resume Next
        fso.DeleteFolder Path
        On Error GoTo 0
        'Exit Function
    End If
Set fso = Nothing
Exit Function
    j = 0
End Function


Public Function DeleteFullFolder(sPath As String, Optional bFolderItself As Boolean = False) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Dim oFile As Scripting.file
Dim oSubFldr As Scripting.Folder
Dim oFldr As Scripting.Folder

    strProcName = ClassName & ".DeleteFullFolder"


    If sPath = "" Then
        LogMessage strProcName, "WARNING", "Code error, no path sent to sub"
        DeleteFullFolder = False
        GoTo Block_Exit
    End If

    Set oFso = New Scripting.FileSystemObject

    If FolderExists(sPath) = False Then
        DeleteFullFolder = True ' end result is no folder there.. so we're good
        GoTo Block_Exit
    End If

    Set oFldr = oFso.GetFolder(sPath)
    If oFldr Is Nothing Then
        DeleteFullFolder = True ' end result is no folder there.. so we're good
        GoTo Block_Exit
    End If

        ' Delete any subfolders
    For Each oSubFldr In oFldr.SubFolders
        Call DeleteFolder(oSubFldr.Path)
    Next

        ' delete any files:
    For Each oFile In oFldr.Files
        oFile.Delete True
    Next

        ' And finally, the folder itself
    If bFolderItself = True Then
        If oFso.FolderExists(sPath) Then
            oFso.DeleteFolder sPath, True
            DeleteFullFolder = True
        End If
    End If

Block_Exit:
    Set oFile = Nothing
    Set oFldr = Nothing
    Set oSubFldr = Nothing
    Set oFso = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    DeleteFullFolder = False
    GoTo Block_Exit
End Function


Public Function RemoveEmptyFolders(Path As String, Optional RemoveRoot As Boolean = False) As Integer
    Dim fso
    Dim fld
    Dim subfld

    Set fso = CreateObject("scripting.filesystemobject")
    Set fld = fso.GetFolder(Path)

    For Each subfld In fld.SubFolders
        Call RemoveEmptyFolders(subfld.Path, True)
    Next

    If fld.Files.Count = 0 And fld.SubFolders.Count = 0 Then
        If RemoveRoot Then
            fld.Delete
        End If
    End If
End Function

Public Function ExportDetails(rst As ADODB.RecordSet, strFilePath As String) As Boolean

    Dim dlg As clsDialogs
    Dim cie As clsImportExport

    Set cie = New clsImportExport
    Set dlg = New clsDialogs

    With dlg
    
        strFilePath = .SavePath(Identity.CurrentFolder, xlsf, strFilePath)
        strFilePath = .CleanFileName(strFilePath, CleanPath)
     
        If strFilePath <> "" Then
        
            If .FileExists(strFilePath) = True Then
            
                If MsgBox("Overwrite existing file?", vbYesNo) = vbYes Then
                    .DeleteFile strFilePath
                Else
                    GoTo exitHere
                End If
            
            End If
            
            If rst.recordCount > 65535 Then
                MsgBox "Warning: Your recordset contains more than 65535 rows, the maximum number of rows allowed in Excel.  " & _
                Trim(str(rst.recordCount - 65535)) & " rows will not be displayed.", vbCritical
            End If
                        
        Else
            GoTo exitHere
        End If
     
        With cie
            .ExportExcelRecordset rst, strFilePath, True
        End With
         
    End With
    
    ExportDetails = True

exitHere:
    Set cie = Nothing
    Set dlg = Nothing
    Exit Function
    
HandleError:
    ExportDetails = False
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    GoTo exitHere

End Function


Public Function Count_TIF_Pages(FileName As String) As Long
   Dim fso As FileSystemObject
    Dim strTempFileName As String
    Dim strLocalPath As String
    Dim strChkFile As String
    Dim bErrFlag As Boolean
    Dim oTIF As Object
    Dim Person As New CT_ClsIdentity
    
    ' TK removed 8/4/2010
'    strLocalPath = "C:\Documents and Settings\" & Identity.UserName & "\My Documents\Scanning"
'    CreateFolder (strLocalPath)
    
    Set fso = New FileSystemObject
    If fso.FileExists(FileName) Then
        strChkFile = FileName
        strTempFileName = "M:\" & gstrAcctAbbrev & "_SCANNING_TEMP.TIF"
        bErrFlag = False
       
        On Error GoTo Err_handler
        Set oTIF = CreateObject("MODI.Document")
    
        On Error GoTo Copy_File
        oTIF.Create strChkFile
    
        On Error GoTo Err_handler
    
        If bErrFlag = True Then
            oTIF.Create strChkFile
        End If
        Count_TIF_Pages = oTIF.BuiltInDocumentProperties("Number of Pages")
    Else
        Count_TIF_Pages = -1
    End If
    
    
Exit_Function:
    Set oTIF = Nothing
    If fso.FileExists(strTempFileName) Then
        fso.DeleteFile strTempFileName
    End If
    
    Set fso = Nothing
    '* jc Set Person = Nothing
    Exit Function
    
Copy_File:
    bErrFlag = True
    'Debug.Print "PROBLEM -- " & FileName
    fso.CopyFile FileName, strTempFileName, True
    strChkFile = strTempFileName
    Resume Next
    
Err_handler:
    Count_TIF_Pages = -1
    Resume Exit_Function
End Function


Public Function Count_PDF_Pages(FileName As String) As Long
    Dim oPDF As Object 'As Acrobat.CAcroPDDoc
    
    On Error GoTo Err_handler
    
    Set oPDF = CreateObject("AcroExch.PDDoc")
    oPDF.Open (FileName)
    Count_PDF_Pages = oPDF.GetNumPages

Exit_Function:
    Set oPDF = Nothing
    Exit Function
    
Err_handler:
    Count_PDF_Pages = -1
    Resume Exit_Function
End Function


Public Function ConvertWordToTif(strInputFilePath As String, strOutputFilePath As String)
    Dim wrdApp As Object
    Dim wrdDoc As Object
    Set wrdApp = CreateObject("Word.Application")
    Set wrdDoc = CreateObject("Word.Document")
    wrdApp.visible = False
    
    Dim strDefaultPrinter As String
    Dim strNewPrinter As String
   
    strDefaultPrinter = Application.Printer.DeviceName
    strNewPrinter = "Microsoft Office Document Image Writer"
    
    On Error GoTo FailHere
    
    ' Temporary changing printer
    wrdApp.WordBasic.FilePrintSetup Printer:=strNewPrinter, DoNotSetAsSysDefault:=1

    ' Get input file
    Set wrdDoc = wrdApp.Documents.Open(strInputFilePath)
    ' Convert .doc to .tif
    wrdDoc.PrintOut PrintToFile:=True, OutputFileName:=strOutputFilePath, Background:=False

    ' close the Word application
    wrdApp.Quit

    ' TIF conversion success
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    ConvertWordToTif = True
    Exit Function
    
FailHere:
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    ConvertWordToTif = False
End Function


Public Function DocumentToPDF(strInputFilePath As String, strOutputFilePath As String)
    Dim AcroApp As Object
    Dim AcroAVDoc As Object
    Dim AcroPDoc As Object
    Dim oJS As Object
    Dim nPages As Long
    
    Set AcroApp = CreateObject("AcroExch.App")
    Set AcroAVDoc = CreateObject("AcroExch.AVDoc")
    Set AcroPDoc = CreateObject("AcroExch.PDDoc")
    
    AcroApp.Hide
    
    ' Get input file
    If AcroAVDoc.Open(strInputFilePath, "") = False Then
        GoTo FailHere
    End If
    
    Set AcroPDoc = AcroAVDoc.GetPDDoc
    
    If AcroPDoc.GetFileName = "" Then
        GoTo FailHere
    End If
    
    ' Convert to pdf file
    AcroPDoc.Save 1, strOutputFilePath
    AcroPDoc.Close
    AcroAVDoc.Close True
    AcroApp.CloseAllDocs
    AcroApp.Exit
    
    Set AcroPDoc = Nothing
    Set AcroAVDoc = Nothing
    Set AcroApp = Nothing

    ' PDF conversion success
    DocumentToPDF = True
    Exit Function
    
FailHere:
    Set AcroPDoc = Nothing
    Set AcroAVDoc = Nothing
    Set AcroApp = Nothing
    DocumentToPDF = False
End Function





Public Function CreateFolders(ByVal strFolderToCreate As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sDrive As String
Dim sNewFolderPath As String
Dim asFolders() As String
Dim l As Integer
Dim sThisFolder As String
Dim iFoundAt As Integer
Dim oFso As Scripting.FileSystemObject
Dim strOrigFolder As String

    strProcName = ClassName & ".CreateFolders"
    CreateFolders = True
    l = 0
    
    If Right(strFolderToCreate, 1) = "\" Then strFolderToCreate = left(strFolderToCreate, Len(strFolderToCreate) - 1)
    
    strOrigFolder = strFolderToCreate
    
    Set oFso = New Scripting.FileSystemObject
        ' if it's a file, reduce it to the parent..
    
    If InStr(Len(strFolderToCreate) - 3, strFolderToCreate, ".", vbTextCompare) > 0 Then
        strFolderToCreate = oFso.GetParentFolderName(strFolderToCreate)
'        LogMessage strProcName, "Original path passed to sub appears to be a file: (" & strOrigFolder & ")", "WARNING", strFolderToCreate
        ' we are  going to say that the orig folder should be strFoldertocreate...
'        strOrigFolder = strFolderToCreate
        ' because we aren't creating a file here..
    End If
    
    sDrive = oFso.GetDriveName(strFolderToCreate)
    If left$(sDrive, 2) = "\\" Then     ' UNC path...
        ' So, we have the "Drive" part in sDrive so let's remove that from the folder path
        ' so we can later generically so soemthing like: Createfolder sDrive & sNewFolderPath
        sNewFolderPath = Replace(strFolderToCreate, sDrive, "", , , vbTextCompare)
    Else            ' Regular path
        sNewFolderPath = strFolderToCreate
    End If
    
    
    '' Build an array of folders that we need to create..
    '' I could and perhaps should just split() the path on "\"...
    
    While sNewFolderPath <> ""
        ReDim Preserve asFolders(l)
        iFoundAt = InStr(1, StrReverse(sNewFolderPath), "\", vbTextCompare)
        sThisFolder = Right$(sNewFolderPath, iFoundAt - 1)
        
        asFolders(l) = sThisFolder
        sNewFolderPath = oFso.GetParentFolderName(sNewFolderPath)
        If sNewFolderPath = sDrive & "\" Then sNewFolderPath = ""
        l = l + 1
    Wend
    
    sNewFolderPath = sDrive & "\"
    
    For l = UBound(asFolders) To 0 Step -1  ' have to start from the root..
        sNewFolderPath = sNewFolderPath & asFolders(l) & "\"
        Debug.Print "folder " & asFolders(l)
        If Not oFso.FolderExists(sNewFolderPath) Then
            ' Create the folder..
            Debug.Print "Creating the folder: " & sNewFolderPath
'            LogMessage strProcName, "DEBUG", "Folder to create: " & sNewFolderPath, sNewFolderPath
            
            ' just in case..
            If Right(sNewFolderPath, 1) = "\" Then
                oFso.CreateFolder (left(sNewFolderPath, Len(sNewFolderPath) - 1))
            Else
                oFso.CreateFolder (sNewFolderPath)
            End If
        End If
    Next
    
    'CreateFolders = oFso.FolderExists(strOrigFolder)
    CreateFolders = oFso.FolderExists(strFolderToCreate)
    
    Set oFso = Nothing
Block_Exit:
    Exit Function

Block_Err:
    CreateFolders = False
    ReportError Err, strProcName, strOrigFolder
    Select Case Err.Number
    Case 76     ' Path not found
        
    Case Else
        Debug.Print "Error: " & Err.Number & " " & Err.Description
    End Select
    GoTo Block_Exit
End Function
'**************DPR 3/7/2013*********************




Public Function TifPageCount_Damon(sFile As String) As Long

    Dim bBytes() As Byte, uOffset As LongType, uTemp As FourBytes, uITemp As TwoBytes, bLEndian As Boolean
    Dim iDirCount As Integer, uInt As IntType, lNext As Long
    
    'Per .tif specification at http://partners.adobe.com/public/developer/en/tiff/TIFF6.pdf
    'The general structure is an 8 byte header that has the starting IFD location.  Each IFD has a pointer
    'to the address of the next, with the last one having a 0.  The first 2 IFD bytes are the number of
    '12 byte entries, and these are followed by the next pointer.
       
    On Error GoTo ErrHand

    bBytes = OpenFileAsArray(sFile)                     'Read the file into a byte array.
    
    uITemp.bByte1 = bBytes(0)                           'These 2 bytes are the byte order.
    uITemp.bByte2 = bBytes(1)
    LSet uInt = uITemp                                  'Convert to an integer.
    If uInt.iInt = 18761 Then                           '18761 = "II"
        bLEndian = True                                 'The file is little-endian byte order.
    Else
        If uInt.iInt <> 19789 Then                      'The other valid setting is "MM", or 19789.
            Exit Function                               'If it isn't either, it's not a valid .tif.
        End If
    End If
            
    If bLEndian Then
        uITemp.bByte1 = bBytes(2)                       'These 2 header bytes are the file identifier.
        uITemp.bByte2 = bBytes(3)
    Else                                                'Big-endian order.
        uITemp.bByte1 = bBytes(3)
        uITemp.bByte2 = bBytes(2)
    End If
    LSet uInt = uITemp                                  'Convert to an integer.
    If uInt.iInt <> 42 Then Exit Function               'If this is not 42, it is not a valid .tif
        
    If bLEndian Then
        uTemp.bByte1 = bBytes(4)                        'The 4-7 bytes of the header are a
        uTemp.bByte2 = bBytes(5)                        'pointer to the first Image File Directory.
        uTemp.bByte3 = bBytes(6)
        uTemp.bByte4 = bBytes(7)
    Else                                                'Big-endian order.
        uTemp.bByte1 = bBytes(7)
        uTemp.bByte2 = bBytes(6)
        uTemp.bByte3 = bBytes(5)
        uTemp.bByte4 = bBytes(4)
    End If
    LSet uOffset = uTemp                                'Convert to a long.
    
    Do Until uOffset.lLong = 0                          'Stop on the null pointer.
        uITemp.bByte1 = bBytes(uOffset.lLong)           'Read the first 2 bytes of the IFD header for
        uITemp.bByte2 = bBytes(uOffset.lLong + 1)       'the number of directory entries for this page.
        LSet uInt = uITemp                              'Convert to an integer.
        lNext = uOffset.lLong + 2 + (12 * uInt.iInt)    'Next pointer location is the number of entries
                                                        'at 12 bytes each plus 2 for the header.
        If bLEndian Then
            uTemp.bByte1 = bBytes(lNext)                'Read the next pointer (4 bytes).
            uTemp.bByte2 = bBytes(lNext + 1)
            uTemp.bByte3 = bBytes(lNext + 2)
            uTemp.bByte4 = bBytes(lNext + 3)
        Else                                            'Big-endian order.
            uTemp.bByte1 = bBytes(lNext + 3)
            uTemp.bByte2 = bBytes(lNext + 2)
            uTemp.bByte3 = bBytes(lNext + 1)
            uTemp.bByte4 = bBytes(lNext)
        End If
        LSet uOffset = uTemp                            'Convert to a long.
        TifPageCount_Damon = TifPageCount_Damon + 1                 'Increment the page count
    Loop

ErrHand:
    Err.Clear               'Not really a problem to just return.  Most errors will be array bounds (last page count
                            'should still be fine, or file opening problems (0 return is appropriate).
End Function

Private Function OpenFileAsArray(sFile As String) As Byte()

    'Utility function that opens a passed filename and returns it as an array of bytes.  Used for the
    'tif and pdf page counting functions.

    Dim bTemp() As Byte, iFile As Integer, sBuffer As String

    iFile = FreeFile
    Open sFile For Binary As #iFile                     'Open the file.
    sBuffer = String$(LOF(iFile), Chr$(0))              'Create a buffer.
    Get #iFile, , sBuffer                               'Write it to the buffer.
    Close #iFile                                        'Close it.

    bTemp = StrConv(sBuffer, vbFromUnicode)             'Convert to a byte array.
    OpenFileAsArray = bTemp

End Function



'**************DPR 3/7/2013*********************


'**************JS 05/09/2013*********************
'Modified function that Damon provided. It was crashing with String Out of Memory Space ERROR when the TIF file was too big because it loaded the whole file into a variable
'Now it reads the file per chunks, it is slower! but at least it doesnt crash.

Public Function TifPageCount(sFile As String) As Long

    Dim bBytes() As Byte, uOffset As LongType, xOffset As LongType, uTemp As FourBytes, uITemp As TwoBytes, bLEndian As Boolean, CurrentByte As Long
    Dim iDirCount As Integer, uInt As IntType, lNext As Long
    
    'Per .tif specification at http://partners.adobe.com/public/developer/en/tiff/TIFF6.pdf
    'The general structure is an 8 byte header that has the starting IFD location.  Each IFD has a pointer
    'to the address of the next, with the last one having a 0.  The first 2 IFD bytes are the number of
    '12 byte entries, and these are followed by the next pointer.
       
    Dim bTemp() As Byte, iFile As Integer, sBuffer As String
    Dim LenghtOfFile As Long

    CurrentByte = 0
       
    On Error GoTo ErrHand
    
    iFile = FreeFile
    Open sFile For Binary As #iFile                     'Open the file.

   
        sBuffer = String$(8, Chr$(0))
        Get #iFile, , sBuffer                               'Write it to the buffer.
        bBytes = StrConv(sBuffer, vbFromUnicode)              'Convert to a byte array.
        CurrentByte = 8
    
        'bBytes = OpenFileAsArray(sFile)                     'Read the file into a byte array.
        
        uITemp.bByte1 = bBytes(0)                           'These 2 bytes are the byte order.
        uITemp.bByte2 = bBytes(1)
        LSet uInt = uITemp                                  'Convert to an integer.
        If uInt.iInt = 18761 Then                           '18761 = "II"
            bLEndian = True                                 'The file is little-endian byte order.
        Else
            If uInt.iInt <> 19789 Then                      'The other valid setting is "MM", or 19789.
                GoTo ExitFunction                               'If it isn't either, it's not a valid .tif.
            End If
        End If
                
        If bLEndian Then
            uITemp.bByte1 = bBytes(2)                       'These 2 header bytes are the file identifier.
            uITemp.bByte2 = bBytes(3)
        Else                                                'Big-endian order.
            uITemp.bByte1 = bBytes(3)
            uITemp.bByte2 = bBytes(2)
        End If
        LSet uInt = uITemp                                  'Convert to an integer.
        If uInt.iInt <> 42 Then GoTo ExitFunction               'If this is not 42, it is not a valid .tif
            
        If bLEndian Then
            uTemp.bByte1 = bBytes(4)                        'The 4-7 bytes of the header are a
            uTemp.bByte2 = bBytes(5)                        'pointer to the first Image File Directory.
            uTemp.bByte3 = bBytes(6)
            uTemp.bByte4 = bBytes(7)
        Else                                                'Big-endian order.
            uTemp.bByte1 = bBytes(7)
            uTemp.bByte2 = bBytes(6)
            uTemp.bByte3 = bBytes(5)
            uTemp.bByte4 = bBytes(4)
        End If
        LSet uOffset = uTemp                                'Convert to a long. We have next pointer here
        xOffset = uOffset
        
        sBuffer = String$(uOffset.lLong - 8, Chr$(0))       'Advance to the next pointer, minus the 8 bytes we already read initially
        Get #iFile, , sBuffer
        CurrentByte = CurrentByte + uOffset.lLong - 8
        
        Do Until xOffset.lLong = 0                          'Stop on the null pointer.
            uOffset.lLong = 0
            
            sBuffer = String$(2, Chr$(0))
            Get #iFile, , sBuffer                               'Write it to the buffer.
            bBytes = StrConv(sBuffer, vbFromUnicode)              'Convert to a byte array.
            CurrentByte = CurrentByte + 2
            
            uITemp.bByte1 = bBytes(0)           'Read the first 2 bytes of the IFD header for
            uITemp.bByte2 = bBytes(1)       'the number of directory entries for this page.
            
            
            LSet uInt = uITemp                              'Convert to an integer.
            lNext = (12 * uInt.iInt)     'Next pointer location is the number of entries
                                                            'at 12 bytes each plus 2 for the header.

                                                            
            sBuffer = String$(lNext, Chr$(0))
            Get #iFile, , sBuffer                               'Write it to the buffer.
            bBytes = StrConv(sBuffer, vbFromUnicode)              'Convert to a byte array.
            CurrentByte = CurrentByte + lNext
                                                            
                                                            
            sBuffer = String$(4, Chr$(0))
            Get #iFile, , sBuffer                               'Write it to the buffer.
            bBytes = StrConv(sBuffer, vbFromUnicode)              'Convert to a byte array.
            CurrentByte = CurrentByte + 4
                                                             
            If bLEndian Then
                uTemp.bByte1 = bBytes(0)                'Read the next pointer (4 bytes).
                uTemp.bByte2 = bBytes(1)
                uTemp.bByte3 = bBytes(2)
                uTemp.bByte4 = bBytes(3)
            Else                                            'Big-endian order.
                uTemp.bByte1 = bBytes(3)
                uTemp.bByte2 = bBytes(2)
                uTemp.bByte3 = bBytes(1)
                uTemp.bByte4 = bBytes(0)
            End If
            
            
            LSet xOffset = uTemp                            'Convert to a long.
            
            If xOffset.lLong > 0 Then
                sBuffer = String$(xOffset.lLong - CurrentByte, Chr$(0))
                Get #iFile, , sBuffer                               'Write it to the buffer.
                bBytes = StrConv(sBuffer, vbFromUnicode)              'Convert to a byte array.
                CurrentByte = CurrentByte + (xOffset.lLong - CurrentByte)
            End If
            
            TifPageCount = TifPageCount + 1                 'Increment the page count
            
        Loop

    GoTo ExitFunction

    'Close #iFile                                        'Close it.

ErrHand:
    'When this image gets converted to PDF will loose a page due to it having an error
    TifPageCount = TifPageCount - 1
    On Error Resume Next
    Err.Clear               'Not really a problem to just return.  Most errors will be array bounds (last page count
                        'should still be fine, or file opening problems (0 return is appropriate).
    
ExitFunction:
    Close #iFile

End Function

'**************JS 05/09/2013*********************

Public Function GetSystemTempFolder() As String
Dim lRet As Long
Dim strPath As String

    ' Init our string buffer
    strPath = Space(Pathlen)
    
    ' Returns Length of path(num characters in path)
    lRet = GetTempPath(Pathlen, strPath)
    ' Chop off NULLS and trailing "\"
    strPath = left(strPath, lRet) & Chr(0)
    
    GetSystemTempFolder = TrimNull(strPath)
    
End Function


Public Function GetUniqueFilename(Optional Path As String = "", Optional Prefix As String = "", Optional UseExtension As String = "") As String
Dim wUnique As Long
Dim lpTempFileName As String
Dim lngRet As Long

    wUnique = 0
    If Path = "" Then Path = CurDir
    
    lpTempFileName = String(MaxPath, 0)
    lngRet = GetTempFileName(Path, Prefix, wUnique, lpTempFileName)
    
    lpTempFileName = left(lpTempFileName, InStr(lpTempFileName, Chr(0)) - 1)
    Call Kill(lpTempFileName)
    If Len(UseExtension) > 0 Then
        lpTempFileName = left(lpTempFileName, Len(lpTempFileName) - 3) & UseExtension
    End If
    GetUniqueFilename = lpTempFileName
    
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' Returns path to the file created (assuming it's been saved)
'''
Public Function ExportRsToExcel(oRs As DAO.RecordSet) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oExcel As Excel.Application
Dim iCol As Integer
Dim oFld As DAO.Field
Dim oWb As Excel.Workbook
Dim oWs As Excel.Worksheet
Dim oNotifyFrm As Form_frm_CMS_User_Notification

    '   Dim oExcel as Object
    '   Dim oWb As Object
    '   Dim oWs As Object

    strProcName = ClassName & ".ExportRsToExcel"
    LogMessage strProcName, "DEBUG TRAIL", "Starting to export " & oRs.recordCount & " records to excel"
    

    Set oNotifyFrm = New Form_frm_CMS_User_Notification
    
    With oNotifyFrm
        .Title = "Working..."
        .UserMessage = "Exporting, this should just take a moment or two.."
        .visible = True
        .Repaint
    End With
    Sleep 500
    DoEvents
    
    
    DoCmd.Hourglass True
    
    
'    Set oNotifyFrm = New Form_frm_CMS_User_Notification
'    With oNotifyFrm
'        .title = "Working..."
'        .UserMessage = "Selecting subset, please wait a moment..."
'        .visible = True
'    End With
    
    ' Open Excel
    Set oExcel = New Excel.Application
    oExcel.visible = False  ' don't show them yet!
    
    ' Dump the RS into a new workbook
    ' maybe we should prompt them to create a new one or select an existing one?
    Set oWb = oExcel.Workbooks.Add
    
    
    Set oWs = oWb.Sheets(1)
    oWs.Activate
    '' First dump the header
    iCol = 1
    For Each oFld In oRs.Fields
        oWs.Cells(1, iCol).Value = oFld.Name
        iCol = iCol + 1
    Next
    
    LogMessage strProcName, "DEBUG TRAIL", "Dumping recordset now"

    oRs.MoveFirst
    oWs.Cells(2, 1).CopyFromRecordset oRs
    LogMessage strProcName, "DEBUG TRAIL", "Formatting Header row"
    
    '' Let's format it a tiny bit:
    ' first the header row..
    oWs.Rows("1:1").Select
    oExcel.selection.Font.Bold = True
    With oExcel.selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With

    LogMessage strProcName, "DEBUG TRAIL", "Filtering  / freeze pane"


        ' Now, fit the columns and stuff
    oWs.Cells.Select
        
    oExcel.selection.AutoFilter
    oExcel.ActiveWindow.SplitRow = 0.7
    oExcel.ActiveWindow.FreezePanes = True
    
    LogMessage strProcName, "DEBUG TRAIL", "Autofit"
    
    
    oWs.Cells.EntireColumn.AutoFit
    
    
    oWs.Cells(2, 1).Select
    LogMessage strProcName, "DEBUG TRAIL", "Finished exporting to excel"

    oExcel.visible = True
    
 
Block_Exit:
        ' Make sure excel isn't hanging invisibly
    DoCmd.Close acForm, oNotifyFrm.Name, acSaveNo
    Set oNotifyFrm = Nothing
    DoCmd.Hourglass False
    If Not oExcel Is Nothing Then
        If oExcel.visible = False Then oExcel.visible = True
    End If
    Exit Function

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


            '' ##############################################################################
            '' ##############################################################################
            '' ##############################################################################

Private Sub QuietExcel(Optional oWb As Excel.Workbook)
On Error GoTo Block_Err
Dim strProcName As String

    ' --- This sub makes sure that Excel is "quiet" so code can run
    '       quicker and more effectively
    '       to reset, call ReleaseExcel() below..

    strProcName = ClassName & ".QuietExcel"
    
    If oWb Is Nothing Or IsMissing(oWb) = True Then
        Set oWb = ThisWorkbook
    End If


    With oWb.Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Interactive = False
'        .Calculation = xlCalculationManual
        .Cursor = xlWait
    End With

Block_Exit:
    Exit Sub

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

            '' ##############################################################################
            '' ##############################################################################
            '' ##############################################################################


Private Sub ReleaseExcel(Optional oWb As Excel.Workbook)
On Error GoTo Block_Err
Dim strProcName As String

    ' --- This sub releases excel from the QuietExcel sub above..
    
    strProcName = ClassName & ".ReleaseExcel"
    
    If oWb Is Nothing Or IsMissing(oWb) = True Then
        Set oWb = ThisWorkbook
    End If

    With oWb.Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Interactive = True
'        .Calculation = IIf(slCurrentUsersCalcMethod = 0, xlCalculationAutomatic, slCurrentUsersCalcMethod)
        .Cursor = xlDefault
    End With

Block_Exit:
    Exit Sub

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Public Function SetFolderHidden(ByVal sFldrPath As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject

    strProcName = ClassName & ".SetFolderHidden"
    ' Set it invisible so they have to click the shortcut to the script..
    Set oFso = New Scripting.FileSystemObject
    oFso.GetFolder(sFldrPath).Attributes = Hidden   '  2 =  Hidden

    SetFolderHidden = True

Block_Exit:
    Set oFso = Nothing
    Exit Function

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
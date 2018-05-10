Option Compare Database
Option Explicit



''' Last Modified: 05/16/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  This is free code to create a PDF file (Acrobat is not required)
'''     What IS required is a dll which is found in the hidden
'''     tbl_App_Dependencies table. It doesn't need to be registered
'''     therefore this app extracts it and uses it when needed
'''
'''  REQUIREMENTS:
'''  =====================================
'''     - tbl_App_Dependencies (and the contents!)
'''     - mod_Blobs (to extract the stuff in the above table)
'''
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 05/16/2012 - updated loadlib to make sure it only looks for the dlls
'''     in the right place for Connolly
'''  - 03/27/2012 - Added to Claim Admin
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


'DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 97 through A2003
'
'Copyright: Stephen Lebans - Lebans Holdings 1999 Ltd.


'Distribution:

' Plain and simple you are free to use this source within your own
' applications. whether private or commercial, without cost or obligation, other that keeping
' the copyright notices intact. No public notice of copyright is required.
' You may not resell this source code by itself or as part of a collection.
' You may not post this code or any portion of this code in electronic format.
' The source may only be downloaded from:
' www.lebans.com
'
'Name:      ConvertReportToPDF
'
'Version:   7.51
'
'Purpose:
'
'а1) Export report to Snapshot and then to PDF. Output exact duplicate of a Report to PDF.
'
'ннннннннннннннннннннннннннннннннннннннннннннннннна
'
'Author:    Stephen Lebans
'Email:     Stephen@lebans.com
'Web Site:  www.lebans.com
'Date:      Feb 21, 2006, 11:11:11 AM
'Dependencies: DynaPDF.dll  StrStorage.dll  clsCommonDialog
'Inputs:    See inline Comments for explanation
'Output:    See inline Comments for explanation
'Credits:   Anyone who wants some!
'BUGS:      Please report any bugs to my email address.
'
'What's Missing:
'           Enhanced Error Handling
'
'How it Works:
' A SnapShot file is created in the normal manner by code like:
'       'Export the selected Report to SnapShot format
'       DoCmd.OutputTo acOutputReport, rptName, "SnapshotFormat(*.snp)", _
'       strPathandFileName
'
' rptName is the desired Report we are working with.
' strPathandFileName can be anything, in this Class it is a
' Temporary FileName and Path created with calls to the
' GetTempPath and GetUniqueFileName API's.
'
' We then pass the FileName to the SetupDecompressOrCopyFile API.
' This will decompress the original SnapShot file into a
' Temporary file with the same name but a "tmp" extension.
'
' The decompressed Temp SnapShot file is then passed to the
' ConvertUncompressedSnapshotToPDF function exposed by the StrStorage DLL.
' The declaration for this call is at the top of this module.
' The function uses the Structured Storage API's to
' open and read the uncompressed Snapshot file. Within this file,
' there is one Enhanced Metafile for each page of the original report.
' Additionally, there is a Header section that contains, among other things,
' a copy of the Report's Printer Devmode structure. We need this to
' determine the page size of the report.

'The StrStorage DLL exposes one function.
'Public Function ConvertUncompressedSnapshotToPDF( _
'UnCompressedSnapShotName As String, _
'OutputPDFname As String = "", _
'Optional CompressionLevel As Long = 0, _
'Optional PasswordOwner As String = "" _
'Optional PasswordOpenAs String = "" _
'Optional PasswordRestrictions as Long = 0, _
'Optional PDFNoFontEmbedding As Long = 0 _
') As Boolean

' Now we call the ConvertUncompressedSnapshotToPDF funtion exposed by the StrStorage DLL.
'
'blRet = ConvertUncompressedSnapshot(sFileName as String, sPDFFileName as String)
'
'
'Have Fun!
'
'
'
' ******************************************************

Private Const ClassName As String = "modNewReportToPDF"

Public Declare Function ConvertUncompressedSnapshot Lib "StrStorage.dll" _
    (ByVal UnCompressedSnapShotName As String, _
    ByVal OutputPDFname As String, _
    Optional ByVal CompressionLevel As Long = 0, _
    Optional ByVal PasswordOwner As String = "", _
    Optional ByVal PasswordOpen As String = "", _
    Optional ByVal PasswordRestrictions As Long = 0, _
    Optional PDFNoFontEmbedding As Long = 0 _
    ) As Boolean

' For debugging with Visual C++
'Lib "C:\VisualCsource\Debug\StrStorage.dll"

Private Declare Function ShellExecuteA Lib "shell32.dll" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" _
    Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
    (ByVal hLibModule As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" _
    Alias "GetTempPathA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Private Declare Function GetTempFileName _
    Lib "kernel32" Alias "GetTempFileNameA" _
    (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long
 
Private Declare Function SetupDecompressOrCopyFile _
    Lib "setupAPI" _
    Alias "SetupDecompressOrCopyFileA" ( _
    ByVal SourceFileName As String, _
    ByVal TargetFileName As String, _
    ByVal CompressionType As Integer) As Long

Private Declare Function SetupGetFileCompressionInfo _
    Lib "setupAPI" _
    Alias "SetupGetFileCompressionInfoA" ( _
    ByVal SourceFileName As String, _
    TargetFileName As String, _
    SourceFileSize As Long, _
    DestinationFileSize As Long, _
    CompressionType As Integer _
    ) As Long

 
'Compression types
Private Const FILE_COMPRESSION_NONE = 0
Private Const FILE_COMPRESSION_WINLZA = 1
Private Const FILE_COMPRESSION_MSZIP = 2

Private Const Pathlen = 256
Private Const MaxPath = 256


' Allow user to set FileName instead
' of using API Temp Filename or
' popping File Dialog Window
Private mSaveFileName As String

' Full path and name of uncompressed SnapShot file
Private mUncompressedSnapFile As String

' Name of the Report we ' working with
Private mReportName As String

' Instance returned from LoadLibrary calls
Private hLibDynaPDF As Long
Private hLibStrStorage As Long



Public Function ConvertReportToPDF( _
    Optional rptName As String = "", _
    Optional ReportWhereCondition As String = "", _
    Optional SnapshotName As String = "", _
    Optional OutputPDFname As String = "", _
    Optional ShowSaveFileDialog As Boolean = False, _
    Optional StartPDFViewer As Boolean = True, _
    Optional CompressionLevel As Long = 0, _
    Optional PasswordOwner As String = "", _
    Optional PasswordOpen As String = "", _
    Optional PasswordRestrictions As Long = 0, _
    Optional PDFNoFontEmbedding As Long = 0 _
    ) As Boolean
Dim S As String
Dim blret As Boolean
Dim strPath  As String
Dim strPathandFileName  As String
Dim strEMFUncompressed As String

Dim sOutFile As String
Dim lngRet As Long
Dim sPath As String * 512
    
            ' RptName is the name of a report contained within this MDB
            ' SnapshotName is the name of an existing Snapshot file
            ' OutputPDFname is the name you select for the output PDF file
            ' ShowSaveFileDialog is a boolean param to specify whether or not to display
            ' the standard windows File Dialog window to select an exisiting Snapshot file
            ' CompressionLevel - not hooked up yet
            ' PasswordOwner  - not hooked up yet
            ' PasswordOpen - not hooked up yet
            ' PasswordRestrictions - not hooked up yet
            ' PDFNoFontEmbedding - Do not Embed fonts in PDF. Set to 1 to stop the
            ' default process of embedding all fonts in the output PDF. If you are
            ' using ONLY - any of the standard Windows fonts
            ' using ONLY - any of the standard 14 Fonts natively supported by the PDF spec
            'The 14 Standard Fonts
            'All version of Adobe's Acrobat support 14 standard fonts. These fonts are always available
            'independent whether they're embedded or not.
            'Family name PostScript name Style
            'Courier Courier fsNone
            'Courier Courier-Bold fsBold
            'Courier Courier-Oblique fsItalic
            'Courier Courier-BoldOblique fsBold + fsItalic
            'Helvetica Helvetica fsNone
            'Helvetica Helvetica-Bold fsBold
            'Helvetica Helvetica-Oblique fsItalic
            'Helvetica Helvetica-BoldOblique fsBold + fsItalic
            'Times Times-Roman fsNone
            'Times Times-Bold fsBold
            'Times Times-Italic fsItalic
            'Times Times-BoldItalic fsBold + fsItalic
            'Symbol Symbol fsNone, other styles are emulated only
            'ZapfDingbats ZapfDingbats fsNone, other styles are emulated only

    If FileExists(QualifyFldrPath(CurrentDBDir) & "DynaPDF.dll") = False Then
        Call ExtractDependencies
    End If

' Let's see if the DynaPDF.DLL is available.
    blret = LoadLib()
    If blret = False Then
        ' Cannot find DynaPDF.dll or StrStorage.dll file
        LogMessage "modNewReportToPDF.ConvertReportToPDF", "ERROR", "Couldn't load the library"
        Exit Function
    End If

On Error GoTo ERR_CREATSNAP

    ' Init our string buffer
    strPath = Space(Pathlen)

    
    ' Let's kill any existing Temp SnapShot file
    If Len(mUncompressedSnapFile & vbNullString) > 0 Then
        Kill mUncompressedSnapFile
        mUncompressedSnapFile = ""
    End If
    
    ' If we have been passed the name of a Snapshot file then
    ' skip the Snapshot creation process below
    If Len(SnapshotName & vbNullString) = 0 Then
          
        ' Make sure we were passed a ReportName
        If Len(rptName & vbNullString) = 0 Then
            ' No valid parameters - FAIL AND EXIT!!
            ConvertReportToPDF = ""
            Exit Function
        End If
            
        ' Get the Systems Temp path
        ' Returns Length of path(num characters in path)
        lngRet = GetTempPath(Pathlen, strPath)
        ' Chop off NULLS and trailing "\"
        strPath = left(strPath, lngRet) & Chr(0)
        
        ' Now need a unique Filename
        ' locked from a previous aborted attemp.
        ' Needs more work!
        strPathandFileName = GetUniqueFilename(strPath, "SNP" & Chr(0), "snp")
        
        DoCmd.OpenReport rptName, acViewPreview, , ReportWhereCondition, acHidden
        
        ' Export the selected Report to SnapShot format
        DoCmd.OutputTo acOutputReport, rptName, "SnapshotFormat(*.snp)", strPathandFileName
        ' Make sure the process has time to complete
        DoEvents
    
    Else
        strPathandFileName = SnapshotName
     
    End If
    
    ' Let's decompress into same filename but change type to ".tmp"
    lngRet = GetTempPath(512, sPath)
    
    strEMFUncompressed = GetUniqueFilename(sPath, "SNP", "tmp")
    
    lngRet = SetupDecompressOrCopyFile(strPathandFileName, strEMFUncompressed, 0&)
    
    If lngRet <> 0 Then
        Err.Raise vbObjectError + 525, "ConvertReportToPDF.SetupDecompressOrCopyFile", _
        "Sorry...cannot Decompress SnapShot File" & vbCrLf & _
        "Please select a different Report to Export"
    End If
    
    ' Set our uncompressed SnapShot file name var
    mUncompressedSnapFile = strEMFUncompressed
    
    ' Remember to Cleanup our Temp SnapShot File if we were NOT passed the
    ' Snapshot file as the optional param
    If Len(SnapshotName & vbNullString) = 0 Then
        Kill strPathandFileName
    End If
    
    
    ' Do we name output file the same as the input file name
    ' and simply change the file extension to .PDF or
    ' do we show the File Save Dialog
    If ShowSaveFileDialog = False Then
    
        ' let's decompress into same filename but change type to ".tmp"
        ' But first let's see if we were passed an output PDF file name
        If Len(OutputPDFname & vbNullString) = 0 Then
            sOutFile = Mid(strPathandFileName, 1, Len(strPathandFileName) - 3)
            sOutFile = sOutFile & "PDF"
        Else
            sOutFile = OutputPDFname
        End If
    
    Else
    
    End If
    
    ' Call our function in the StrStorage DLL
    ' Note the Compression and Password params are not hooked up yet.
    blret = ConvertUncompressedSnapshot(mUncompressedSnapFile, sOutFile, _
    CompressionLevel, PasswordOwner, PasswordOpen, PasswordRestrictions, PDFNoFontEmbedding)
    
    If blret = False Then
        Err.Raise vbObjectError + 526, "ConvertReportToPDF.ConvertUncompressedSnaphot", _
            "Sorry...damaged SnapShot File" & vbCrLf & _
            "Please select a different Report to Export"
    End If
    
    ' Do we open new PDF in registered PDF viewer on this system?
    If StartPDFViewer = True Then
        ShellExecuteA Application.hWndAccessApp, "open", sOutFile, vbNullString, vbNullString, 1
    End If
    
    ' Success
    ConvertReportToPDF = True
    
    
EXIT_CREATESNAP:
    
    ' Let's kill any existing Temp SnapShot file
    If Len(mUncompressedSnapFile & vbNullString) > 0 Then
         On Error Resume Next
       Kill mUncompressedSnapFile
        mUncompressedSnapFile = ""
    End If
    
    ' If we aready loaded then free the library
    If hLibStrStorage <> 0 Then
        hLibStrStorage = FreeLibrary(hLibStrStorage)
    End If
    
    If hLibDynaPDF <> 0 Then
        hLibDynaPDF = FreeLibrary(hLibDynaPDF)
    End If

    Exit Function

ERR_CREATSNAP:
    MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
    mUncompressedSnapFile = ""
    ConvertReportToPDF = False
    Resume EXIT_CREATESNAP
End Function



Private Function LoadLib() As Boolean
Dim S As String
Dim blret As Boolean
Dim strProcName As String
Dim blnAlreadyExtracted As Boolean

On Error Resume Next

    strProcName = ClassName & ".LoadLib"

    LoadLib = False
    
    ' If we aready loaded then free the library
    If hLibDynaPDF <> 0 Then
        hLibDynaPDF = FreeLibrary(hLibDynaPDF)
    End If
    
    
    ' Our error string
    S = "Sorry...cannot find the DynaPDF.dll file" & vbCrLf
    S = S & "Please copy the DynaPDF.dll file to your Windows System32 folder or into the same folder as this Access MDB."
    
    ' OK Try to load the DLL assuming it is in the Window System folder
'    hLibDynaPDF = LoadLibrary("DynaPDF.dll")
'    If hLibDynaPDF = 0 Then
'        ' See if the DLL is in the same folder as this MDB
'        ' CurrentDB works with both A97 and A2K or higher
TryDynaAgain:
        hLibDynaPDF = LoadLibrary(QualifyFldrPath(CurrentDBDir()) & "DynaPDF.dll")
        If hLibDynaPDF = 0 Then
            ' not found, let's extract it from the tbl_App_Dependencies and try again
            If blnAlreadyExtracted = True Then
                ''UpdateUser "Cannot create PDF due to the lack of DLL dependencies", strProcName, "ERROR", , "DynaPDF.dll"
                LogMessage strProcName, "ERROR", "Cannot create PDF due to lack of DLL dependencies", "DynaPDF.dll"
                MsgBox S, vbOKOnly, "MISSING DynaPDF.dll FILE"
                LoadLib = False
            Else
                ExtractDependencies
                Sleep 1000
                blnAlreadyExtracted = True
                GoTo TryDynaAgain
            End If
            Exit Function
        End If
'    End If
    
    
    ' Our error string
    S = "Sorry...cannot find the StrStorage.dll file" & vbCrLf
    S = S & "Please copy the StrStorage.dll file to your Windows System32 folder or into the same folder as this Access MDB."
    
    ' ** Commented out for Debugging only - Must be active
    ' ***************************************************************************
    '
    ' OK Try to load the DLL assuming it is in the Window System folder
'    hLibStrStorage = LoadLibrary("StrStorage.dll")
'    If hLibStrStorage = 0 Then
        ' See if the DLL is in the same folder as this MDB
        ' CurrentDB works with both A97 and A2K or higher
TryStorageAgain:
        hLibStrStorage = LoadLibrary(QualifyFldrPath(CurrentDBDir()) & "StrStorage.dll")
        If hLibStrStorage = 0 Then
            If blnAlreadyExtracted = True Then
                ''UpdateUser "Cannot create PDF due to the lack of DLL dependencies", strProcName, "ERROR", , "DynaPDF.dll"
                LogMessage strProcName, "ERROR", "Cannot create PDF due to lack of DLL dependencies", "DynaPDF.dll"
                MsgBox S, vbOKOnly, "MISSING StrStorage.dll FILE"
            Else
                ExtractDependencies
                blnAlreadyExtracted = True
                GoTo TryStorageAgain
            End If
            Exit Function
        End If
'    End If
    
    
    ' RETURN SUCCESS
    LoadLib = True
End Function


''******************** Code Begin ****************
''Code courtesy of
''Terry Kreft & Ken Getz
''
'Private Function CurrentDBDir() As String
'Dim strDBPath As String
'Dim strDBFile As String
'    strDBPath = CurrentDb.Name
'    strDBFile = Dir(strDBPath)
'    CurrentDBDir = Left$(strDBPath, Len(strDBPath) - Len(strDBFile))
'End Function
''******************** Code End ****************



Private Function GetUniqueFilename(Optional Path As String = "", _
    Optional Prefix As String = "", _
    Optional UseExtension As String = "") _
    As String

' originally Posted by Terry Kreft
' to: comp.Databases.ms -Access
' Subject:  Re: Creating Unique filename ??? (Dev code)
' Date: 01/15/2000
' Author: Terry Kreft <terry.kreft@mps.co.uk>

' SL Note: Input strings must be NULL terminated.
' Here it is done by the calling function.

Dim wUnique As Long
Dim lpTempFileName As String
Dim lngRet As Long

  wUnique = 0
  If Path = "" Then Path = CurDir
  lpTempFileName = String(MaxPath, 0)
  lngRet = GetTempFileName(Path, Prefix, _
                            wUnique, lpTempFileName)

  lpTempFileName = left(lpTempFileName, _
                        InStr(lpTempFileName, Chr(0)) - 1)
  Call Kill(lpTempFileName)
  If Len(UseExtension) > 0 Then
    lpTempFileName = left(lpTempFileName, Len(lpTempFileName) - 3) & UseExtension
  End If
  GetUniqueFilename = lpTempFileName
End Function


'''Private Function fFileDialog() As String
'''' Calls the API File Save Dialog Window
'''' Returns full path to new File
'''
'''On Error GoTo Err_fFileDialog
'''
'''' Call the File Common Dialog Window
'''Dim clsDialog As Object
'''Dim strTemp As String
'''Dim strFname As String
'''
'''    Set clsDialog = New clsCommonDialog
'''
'''    ' Fill in our structure
'''    ' I'll leave in how to select Gif and Jpeg to
'''    ' show you how to build the Filter in case you want
'''    ' to use this code in another project.
'''    clsDialog.Filter = "PDF (*.PDF)" & Chr$(0) & "*.PDF" & Chr$(0)
'''    'clsDialog.Filter = clsDialog.Filter & "Gif (*.GIF)" & Chr$(0) & "*.GIF" & Chr$(0)
'''    'clsDialog.Filter = "ALL (*.*)" & Chr$(0) & "*.*" & Chr$(0)
'''    clsDialog.hdc = 0
'''    clsDialog.MaxFileSize = 256
'''    clsDialog.Max = 256
'''    clsDialog.FileTitle = vbNullString
'''    clsDialog.DialogTitle = "Please Select a path and Enter a Name for the PDF File"
'''    clsDialog.InitDir = vbNullString
'''    clsDialog.DefaultExt = vbNullString
'''
'''    ' Display the File Dialog
'''    clsDialog.ShowSave
'''
'''    ' See if user clicked Cancel or even selected
'''    ' the very same file already selected
'''    strFname = clsDialog.FileName
'''    'If Len(strFname & vbNullString) = 0 Then
'''    ' Raise the exception
'''     ' Err.Raise vbObjectError + 513, "clsPrintToFit.fFileDialog", _
'''      '"Please type in a Name for a New File"
'''    'End If
'''
'''    ' Return File Path and Name
'''    fFileDialog = strFname
'''
'''Exit_fFileDialog:
'''
'''    Err.Clear
'''    Set clsDialog = Nothing
'''    Exit Function
'''
'''Err_fFileDialog:
'''    fFileDialog = ""
'''    MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
'''    Resume Exit_fFileDialog
'''
'''End Function




'Public Function fFileDialogSnapshot() As String
'' Calls the API File Open Dialog Window
'' Returns full path to existing Snapshot File
'
'On Error GoTo Err_fFileDialog
'
'' Call the File Common Dialog Window
'Dim clsDialog As Object
'Dim strTemp As String
'Dim strFname As String
'
'    Set clsDialog = New clsCommonDialog
'
'    ' Fill in our structure
'    ' I'll leave in how to select Gif and Jpeg to
'    ' show you how to build the Filter in case you want
'    ' to use this code in another project.
'    clsDialog.Filter = "SNAPSHOT (*.SNP)" & Chr$(0) & "*.SNP" & Chr$(0)
'    'clsDialog.Filter = "ALL (*.*)" & Chr$(0) & "*.*" & Chr$(0)
'    clsDialog.hdc = 0
'    clsDialog.MaxFileSize = 256
'    clsDialog.Max = 256
'    clsDialog.FileTitle = vbNullString
'    clsDialog.DialogTitle = "Please Select a Snapshot File"
'    clsDialog.InitDir = vbNullString
'    clsDialog.DefaultExt = vbNullString
'
'    ' Display the File Dialog
'    clsDialog.ShowOpen
'
'    ' See if user clicked Cancel or even selected
'    ' the very same file already selected
'    strFname = clsDialog.FileName
'    If Len(strFname & vbNullString) = 0 Then
'    ' Do nothing. Add your desired error logic here.
'    End If
'
'    ' Return File Path and Name
'    fFileDialogSnapshot = strFname
'
'Exit_fFileDialog:
'
'    Err.Clear
'    Set clsDialog = Nothing
'    Exit Function
'
'Err_fFileDialog:
'    fFileDialogSnapshot = ""
'    MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
'    Resume Exit_fFileDialog
'
'End Function



'Public Function fFileDialogSavePDFname() As String
'' Calls the API File Open Dialog Window
'' Returns full path to existing Snapshot File
'
'On Error GoTo Err_fFileDialog
'
'' Call the File Common Dialog Window
'Dim clsDialog As Object
'Dim strTemp As String
'Dim strFname As String
'
'    Set clsDialog = New clsCommonDialog
'
'    ' Fill in our structure
'    ' I'll leave in how to select Gif and Jpeg to
'    ' show you how to build the Filter in case you want
'    ' to use this code in another project.
'    clsDialog.Filter = "PDF (*.PDF)" & Chr$(0) & "*.PDF" & Chr$(0)
'    'clsDialog.Filter = "ALL (*.*)" & Chr$(0) & "*.*" & Chr$(0)
'    clsDialog.hdc = 0
'    clsDialog.MaxFileSize = 256
'    clsDialog.Max = 256
'    clsDialog.FileTitle = vbNullString
'    clsDialog.DialogTitle = "Please Select a name for the PDF File"
'    clsDialog.InitDir = vbNullString
'    clsDialog.DefaultExt = vbNullString
'
'
'
'    ' Display the File Dialog
'    clsDialog.ShowOpen
'
'    ' See if user clicked Cancel or even selected
'    ' the very same file already selected
'    strFname = clsDialog.FileName
'    If Len(strFname & vbNullString) = 0 Then
'    ' Do nothing. Add your desired error logic here.
'    End If
'
'    ' Return File Path and Name
'    fFileDialogSavePDFname = strFname
'
'Exit_fFileDialog:
'
'    Err.Clear
'    Set clsDialog = Nothing
'    Exit Function
'
'Err_fFileDialog:
'    fFileDialogSavePDFname = ""
'    MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
'    Resume Exit_fFileDialog
'
'End Function
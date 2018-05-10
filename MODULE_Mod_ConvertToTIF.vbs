Option Compare Database
Option Explicit


Private Const ClassName As String = "modConvertMethods"


Private Const cs_TIFF_CP_EXE_PATH As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\AUDITOR_FOLDERS\FAX_REPOSITORY\EXE\" & "TiffCP.exe"
Private Const cs_BIN_PATH As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\AUDITOR_FOLDERS\FAX_REPOSITORY\EXE\"

Public Function GetFilePartName(FileName As String)

Dim sMid As String
Dim intPos As Integer
Dim EndPos As Integer

sMid = (FileName)

intPos = InStrRev(sMid, ".")
EndPos = InStrRev(sMid, "\") + 1
GetFilePartName = Mid(sMid, EndPos, intPos - EndPos)

End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' Accepts a full path to a .doc or .docx file
''' and creates a .tiff in the same directory (unless sOutPath is specified
''' which would be the full path to a file)
''' Incidentially, we first have to convert to PDF (at the moment)
''' Returns true on success AND sOutPath is changed to the TIFF's full path (if not
''' already)
'''
Public Function TiffToPdf(sInFilePath As String, Optional sOutPath As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sReturnedPath As String
Dim sPdfPath As String
Dim sCmd As String

    strProcName = ClassName & ".TiffToPdf"


    sPdfPath = sOutPath
    If sPdfPath = "" Then
        sPdfPath = Replace(sInFilePath, ".tif", ".pdf", , , vbTextCompare)
    End If
    sPdfPath = Replace(sPdfPath, ".tif", ".pdf")
    sPdfPath = Replace(sPdfPath, ".tiff", ".pdf") ' just in case
    
    sCmd = cs_BIN_PATH & "tiff2pdf.exe -d -o """ & sOutPath & """  """ & sInFilePath & """"
    Debug.Print sCmd
    Sleep 250
'    Shell sCmd, vbHide
Debug.Print Now() & "Converting...    "
    ShellWait sCmd, vbHide
Debug.Print Now() & "Finished...    "
    
    Sleep 350
    
    TiffToPdf = FileExists(sPdfPath)

    If TiffToPdf = False Then
        LogMessage strProcName, "ERROR", "The resultant PDF file doesn't exist after conversion...", sPdfPath
    End If

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    TiffToPdf = False
    GoTo Block_Exit
End Function

   


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' Accepts a full path to a .doc or .docx file
''' and creates a .tiff in the same directory (unless sOutPath is specified
''' which would be the full path to a file)
''' Incidentially, we first have to convert to PDF (at the moment)
''' Returns true on success AND sOutPath is changed to the TIFF's full path (if not
''' already)
'''
Public Function DocToTiff(sInFilePath As String, Optional sOutPath As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sReturnedPath As String
Dim sPdfPath As String

    strProcName = ClassName & ".DocToTiff"

    sPdfPath = sOutPath
    sPdfPath = Replace(sPdfPath, ".tif", ".pdf")
    sPdfPath = Replace(sPdfPath, ".tiff", ".pdf") ' just in case
    
    ' This (for now) is going to have to call Doc To Pdf, then PDf to Tiff..
   '''' If DocToPdf(sInFilePath, sPdfPath) = False Then
   ''''     LogMessage strProcName, "ERROR", "Problem converting word doc file to PDF"
   ''''     GoTo Block_Exit
   '''' End If
'Stop
    
    Sleep 500
    
    If PdfToTiff(sPdfPath, sOutPath) = False Then
        LogMessage strProcName, "ERROR", "Problem converting pdf to tif"
        GoTo Block_Exit
    End If
    ' Now, clean up the pdf:
    If FileExists(sOutPath) And FileExists(sPdfPath) Then
        DeleteFile sPdfPath, False
    End If
    
    DocToTiff = sOutPath
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    DocToTiff = False
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' Accepts a full path to a .pdf file
''' and creates a .tiff in the same directory (unless sOutPath is specified
''' which would be the full path to a file)
''' Returns true on success AND sOutPath is changed to the TIFF's full path (if not
''' already)
'''
Public Function PdfToTiff(sInFilePath As String, Optional sOutPath As String) As String
On Error GoTo Block_Err

Dim oAcroApp As AcroApp
Dim oAcroPdDoc As AcroPDDoc

Dim oAcroJSO As Object  '' hmm.. can't find the actual type yet.. Oh. that's because it's not really a .. Type..
Dim oFso As Scripting.FileSystemObject

Dim iPageCount As Integer

Dim pdfFileName As String
Dim strProcName As String

    strProcName = ClassName & ".PdfToTiff"


    Set oFso = New Scripting.FileSystemObject
    
    If sOutPath = "" Then
        sOutPath = oFso.GetParentFolderName(sInFilePath)
    End If
    
    If Right(sOutPath, 4) <> ".tif" Then
        sOutPath = QualifyFldrPath(sOutPath)
    End If
    
    
    Set oAcroApp = New AcroApp
    Set oAcroPdDoc = New AcroPDDoc
    
    If LCase(Right(sInFilePath, 4)) = ".pdf" Then
        If oAcroPdDoc.Open(sInFilePath) = True Then
                ' How many pages?
            iPageCount = oAcroPdDoc.GetNumPages()
            
            Set oAcroJSO = oAcroPdDoc.GetJSObject()
            
            oAcroJSO.SaveAs sOutPath & oFso.GetBaseName(sInFilePath) & ".tif", "com.adobe.acrobat.tiff"
            'oAcroJSO.SaveAs "\\ccaintranet.com\dfs-cms-ds\Data\CMS\AnalystFolders\Viktoria\test.tif", "com.adobe.acrobat.tiff"
            
            Set oAcroJSO = Nothing
            oAcroPdDoc.Close
            
            Sleep 600
            
            If iPageCount > 1 Then
                sOutPath = Replace(sInFilePath, ".pdf", ".tif")

                Call ConcatAdobeTiffExportFiles(sOutPath, iPageCount)
                
            End If
            
        Else
            sOutPath = ""
        End If
    Else
        sOutPath = ""
    End If
    
    If sOutPath <> "" Then
        If FileExists(sOutPath) = True Then
            PdfToTiff = sOutPath
        Else
            PdfToTiff = ""
        End If
    End If
    
Block_Exit:
    oAcroPdDoc.Close
    Set oAcroPdDoc = Nothing
    
    oAcroApp.CloseAllDocs
    oAcroApp.Exit
    Set oAcroApp = Nothing

    Exit Function
Block_Err:
    ReportError Err, strProcName
    MsgBox (Err.Description)
    Err.Clear
    PdfToTiff = False
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''  Conversion helper functions / subs

Private Sub SearchPDFs(oFolder As Scripting.Folder)
Dim oFile As Scripting.file
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Dim oAcroPdDoc As AcroPDDoc
Dim oAcroJSO As Object
Dim oSubFolder As Scripting.Folder
Dim intPages As Integer
Dim sDestDir As String


    strProcName = ClassName & ".FunctionName"



    For Each oFile In oFolder.Files
            If Right(LCase(oFile.Name), 4) = ".pdf" Then
                If Not oFso.FileExists(sDestDir & oFso.GetBaseName(oFile.Path) & ".tif") Then
                    If oAcroPdDoc.Open(oFile.Path) Then

                        intPages = oAcroPdDoc.GetNumPages()
                        'wscript.echo "Pages" & cstr(intPages)
                        Set oAcroJSO = oAcroPdDoc.GetJSObject
                        oAcroJSO.SaveAs sDestDir & oFso.GetBaseName(oFile.Path) & ".tif", "com.adobe.acrobat.tiff"
                        oAcroPdDoc.Close
                        ConcatAdobeTiffExportFiles sDestDir & oFso.GetBaseName(oFile.Path) & ".tif", intPages
                    Else
                        'wscript.echo "  - *** ERROR opening " & oFile.path
                    End If
        
                Else
                    'wscript.echo "* Found:      " & oFile.path
                End If
            End If
        
    Next
        

    
    For Each oSubFolder In oFolder.SubFolders
        SearchPDFs oSubFolder
    Next
    

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Sub ConcatAdobeTiffExportFiles(sDestFileName As String, intPages As Integer)
On Error GoTo Block_Err
Dim strProcName As String
Dim strPad As String
Dim oFso As Scripting.FileSystemObject
Dim sDestFileBase As String
Dim oFolder As Scripting.Folder
Dim oFile As Scripting.file
Dim sPageOneFile As String
Dim sCurFile As String
Dim sOutFolder As String


    strProcName = ClassName & ".ConcatAdobeTiffExportFiles"
    
    strPad = String(Len(CStr(intPages)) - 1, "0")
    
    Set oFso = New Scripting.FileSystemObject
    sOutFolder = QualifyFldrPath(oFso.GetParentFolderName(sDestFileName))

    sDestFileBase = oFso.GetParentFolderName(sDestFileName) & "\" & oFso.GetBaseName(sDestFileName)
    Set oFolder = oFso.GetFolder(oFso.GetParentFolderName(sDestFileName))
    
    If oFso.FileExists(sDestFileName) Then
        oFso.DeleteFile (sDestFileName)
    End If

    sPageOneFile = sOutFolder & oFso.GetBaseName(sDestFileName) & "_Page_" & strPad & "1.tif"
    If InStr(1, sPageOneFile, sOutFolder, vbTextCompare) < 1 Then
        sPageOneFile = sOutFolder & sPageOneFile
    End If

        ' If we don't have page 1 then we have a bit of a problem don't we?
    If oFso.FileExists(sPageOneFile) = False Then GoTo Block_Exit

    For Each oFile In oFolder.Files
    
        If InStr(oFile.Name, oFso.GetBaseName(sDestFileName) & "_Page_") = 1 Then
                ' we don't want to concat the first page with the first page so skip that
            If LCase(oFile.Path) <> LCase(sPageOneFile) Then
                Sleep 250
                Shell cs_TIFF_CP_EXE_PATH & " -a """ & oFile.Path & """ """ & sPageOneFile & """", vbHide
                Sleep 500
            End If
        End If
    
    Next

        ' We are done concat'ing so rename the page 1 file to the final file name
    Sleep 650
    oFso.MoveFile sPageOneFile, sDestFileName
    
        ' ANd delete the other pages that have already been concatenated
    For Each oFile In oFolder.Files
        If InStr(oFile.Name, oFso.GetBaseName(sDestFileName) & "_Page_") = 1 Then
            oFso.DeleteFile (oFile.Path)
        End If
    Next

Block_Exit:
    Set oFile = Nothing
    Set oFolder = Nothing
    Set oFso = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub


' Note to self: probably should pass in the pdf doc's pagect & use that to determine the filenames
' of the individual tiff docs. The current routine is fine up to page 99, and then our dependence on
' the filenames be listed alphabetically may cause issues (<pdffile>_Page_00.pdf)
'
' ON SECOND THOUGHT -- we're fine. Acrobat automatically pads the number w/zero based on the largest
' pagenum present (0, 00, 000, etc...)


Private Sub ConcatAdobeTiffExportFiles_LEGACY(sDestFileName As String, intPages As Integer)
On Error GoTo Block_Err
Dim strProcName As String
Dim strPad As String
Dim oFso As Scripting.FileSystemObject
Dim sDestFileBase As String
Dim oFolder As Scripting.Folder
Dim oFile As Scripting.file
Dim strTempFile As String

    strProcName = ClassName & ".ConcatAdobeTiffExportFiles"
    
    If intPages < 10 Then strPad = ""
    If intPages >= 10 And intPages < 100 Then strPad = "0"
    If intPages >= 100 And intPages < 1000 Then strPad = "00"
    If intPages >= 1000 And intPages < 10000 Then strPad = "000"
    
    sDestFileBase = oFso.GetParentFolderName(sDestFileName) & "\" & oFso.GetBaseName(sDestFileName)
    Set oFolder = oFso.GetFolder(oFso.GetParentFolderName(sDestFileName))
    
    If oFso.FileExists(sDestFileName) Then
        oFso.DeleteFile (sDestFileName)
    End If
    
'Dim sCurPath
'sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")


    For Each oFile In oFolder.Files
        If InStr(oFile.Name, oFso.GetBaseName(sDestFileName) & "_Page_") = 1 Then
            Shell cs_TIFF_CP_EXE_PATH & " -a """ & oFile.Path & """ """ & Replace(sDestFileName, ".tif", "_Page_" & strPad & "1.tif") & """", vbHide
'            oShell.Run sCurPath & "\bin\tiffcp.exe -a """ & oFile.path & """ """ & Replace(sDestFileName, ".tif", "_Page_" & strPad & "1.tif") & """", 0, True
        End If
    
    Next



    strTempFile = Replace(sDestFileName, ".tif", "_Page_" & strPad & "1.tif")
    'wscript.echo strTempFile & " " & strDestFileName
    oFso.MoveFile strTempFile, sDestFileName
                   

    For Each oFile In oFolder.Files
        If InStr(oFile.Name, oFso.GetBaseName(sDestFileName) & "_Page_") = 1 Then
            oFso.DeleteFile (oFile.Path)
        End If
    Next




Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub




Private Sub CreateFolderPath(ByVal sPath)
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".CreateFolderPath"
    
    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(.GetParentFolderName(sPath)) Then
            CreateFolderPath .GetParentFolderName(sPath)
        End If
        If Not .FolderExists(sPath) Then
            .CreateFolder (sPath)
        End If
    End With
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub


'#
'# Usage:        ClmPkg_FileInUse(sFileName)
'#
'# Parameters:   sFileName:    Filename to check
'#
'# Returns:      True if file currently locked by another process, false if not.
'#
'# Purpose:      Checks to see if the provided sFileName is in use by another
'#               process (so that we can surface an appropriate message to the user).
'#

Public Function ClmPkg_FileInUse(SFileName As String) As Boolean
Dim iFileNum As Integer
Dim lErrNum As Long
Dim sErrDesc As String

    ' Attempt to open the file and lock it.
    On Error Resume Next
    iFileNum = FreeFile()
    Sleep 250
    Open SFileName For Input Lock Read As iFileNum
    Close iFileNum
    lErrNum = Err.Number
    On Error GoTo 0
    
    ClmPkg_FileInUse = (lErrNum = 70) ' 70 = Permission Denied. - File is opened by another user.
    
End Function
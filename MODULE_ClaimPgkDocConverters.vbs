'#########################################################################################################
'#                                                                                                       #
'# Module Name:          CnlyClmPkgConverters (Paperless Claims)                                         #
'#                                                                                                       #
'# Description:          Contains all document conversion code (e.g. DOC to PDF). These functions are    #
'#                       publically declared and are standalone -- although they have dependencies       #
'#                       on the rest of the Claim Packager code.                                         #
'#                                                                                                       #
'#                       If you have a need to perform conversions to PDF outside of a package, go       #
'#                       ahead and call these routines. Even if we change the technologies used here,    #
'#                       the calls should remain the same.                                               #
'#                                                                                                       #
'# Original Author:      Karl Erickson           (06/01/2010)                                            #
'# Last Update By:       Karl Erickson           (06/01/2010)                                            #
'#                                                                                                       #
'# Change History:       [#] [MM/DD/YYYY]  [Author Name]    [Explanation of Change]                      #
'#                       --- ------------  ---------------  -------------------------------------------- #
'#                       000 06/01/2010    Karl Erickson    Created                                      #
'#                       001 02/24/2011    Barbara Dyroff   Commented out the following functions due    #
'#                                                          to missing referenced classes and functions. #
'#                                                          ClmPkg_Xls2Pdf, ClmPkg_Eml2Pdf,              #
'#                                                          ClmPkg_Jpg2Pdf, ClmPkg_Tif2Pdf,              #
'#                                                          ClmPkg_Bmp2Pdf                               #
'#                                                                                                       #
'#########################################################################################################

Option Compare Database
'Option Explicit

'#
'# Usage:        ClmPkg_Xls2Pdf(sXlsFile, sPdfFile, sErrorTxt)
'#
'# Parameters:   sXlsFile:          Full path/file name of source XLS/XLSX file
'#               sPdfFile:          Full path/file name of the destination PDF file (that we're converting to)
'#               sErrorTxt:         Error text (returns false if error present as well)
'#
'# Returns:      True if completely successful, False if not.
'#
'# Purpose:      Converts the given MS Excel (up to 2007) file into a PDF file
'#
'
'Public Function ClmPkg_Xls2Pdf(sXlsFile As String, sPdfFile As String, ByRef sErrorTxt As String) As Boolean
'
'    Dim oFSO As Object         'System.FileSystemObject
'    Dim oExcelApp As Object    'Excel.Application
'    Dim oExcelWB As Object     'Excel.Workbook
'    Dim sFolder As String
'
'    ClmPkg_Xls2Pdf = False
'    sErrorTxt = "Xls2Pdf: Unable to access Microsoft Excel" ' Default error message
'
'    If ClmPkg_FileInUse(sXlsFile) Then
'        sErrorTxt = "Xls2Pdf: Document file aready in use"
'    Else
'
'        On Error Resume Next
'        Set oFSO = CreateObject("Scripting.FileSystemObject")
'
'        Set oExcelApp = CreateObject("Excel.Application")
'
'        ' Office Versions:
'        '   12.0 = Office 2007
'        '   11.0 = Office 2003
'        '   10.0 = Office 2002/XP
'        '    9.0 = Office 2000
'
'        ' Ensure we're at Word 2007 (at least), and that the "Save as PDF" addon (or SP2) is installed.
'        ' (The referenced dll could change in future Word versions, but there isn't any other method
'        ' exposed to determine if it's installed. Presumably it'll be standard with Office 2010+...)
'
'        If oExcelApp.Version >= 12# And _
'           oFSO.FileExists(Environ("commonprogramfiles") & "\Microsoft Shared\OFFICE" & Format(Val(oExcelApp.Version), "00") & "\EXP_PDF.DLL") Then
'
'            Set oExcelWB = oExcelApp.Workbooks.Open(sXlsFile)
'
'            oExcelWB.CheckCompatibility = False
'            'oExcelApp.visible = True
'            oExcelApp.DisplayAlerts = False
'            'oExcelWB.SaveAs sPdfFile, xlTypePDF
'            oExcelWB.ActiveSheet.ExportAsFixedFormat FileName:=sPdfFile, Type:=xlTypePDF
'            oExcelWB.Close wdDoNotSaveChanges
'            'oExcelApp.Quit WdDoNotSaveChanges
'            ClmPkg_Xls2Pdf = True
'            sErrorTxt = ""
'        ElseIf oExcelApp.Version >= 12# Then
'            sErrorTxt = "Xls2Pdf: Excel does not have 'Save as PDF' component installed"
'            ClmPkg_Xls2Pdf = False
'        Else
'            sErrorTxt = "Xls2Pdf: Excel 2007 (or newer) is not installed"
'            ClmPkg_Xls2Pdf = False
'        End If
'
'        Set oExcelApp = Nothing
'        Set oFSO = Nothing
'    End If
'
'    On Error GoTo 0
'
'End Function

'#
'# Usage:        ClmPkg_Doc2Pdf(sDocFile, sPdfFile, sErrorTxt)
'#
'# Parameters:   sDocFile:          Full path/file name of source DOC/DOCX file
'#               sPdfFile:          Full path/file name of the destination PDF file (that we're converting to)
'#               sErrorTxt:         Error text (returns false if error present as well)
'#
'# Returns:      True if completely successful, False if not.
'#
'# Purpose:      Converts the given MS Word (up to 2007) file into a PDF file
'#

Public Function ClmPkg_Doc2Pdf(ByVal sDocFile As String, ByVal sPdfFile As String, ByRef sErrorTxt As String) As Boolean

    Dim oFso As Object         'System.FileSystemObject
    Dim oWordApp As Object     'Word.Application
    Dim oWordDoc As Object     'Word.Document
    Dim oWordDocs As Object    'Word.Documents
    Dim sFolder As String
    
    
    ClmPkg_Doc2Pdf = False
    sErrorTxt = "Doc2Pdf: Unable to access Microsoft Word" ' Default error message
    
    If ClmPkg_FileInUse(sDocFile) Then
        sErrorTxt = "Doc2Pdf: Document file aready in use"
    Else
    
        On Error Resume Next
        
        Set oFso = CreateObject("Scripting.FileSystemObject")
        Set oWordApp = CreateObject("Word.Application")
        'Set oWordDocs = oWordApp.Documents
        
        ' Office Versions:
        '   12.0 = Office 2007
        '   11.0 = Office 2003
        '   10.0 = Office 2002/XP
        '    9.0 = Office 2000
        
        ' Ensure we're at Word 2007 (at least), and that the "Save as PDF" addon (or SP2) is installed.
        ' (The referenced dll could change in future Word versions, but there isn't any other method
        ' exposed to determine if it's installed. Presumably it'll be standard with Office 2010+...)
        
        
        If oWordApp.Version >= 12# And _
           oFso.FileExists(Environ("commonprogramfiles") & "\Microsoft Shared\OFFICE" & Format(val(oWordApp.Version), "00") & "\EXP_PDF.DLL") Then
        
                Set oWordDoc = oWordApp.Documents.Open(sDocFile)
                oWordDoc.SaveAs sPdfFile, 17 'wdFormatPDF
                oWordDoc.Close 0 'wdDoNotSaveChanges
                
                ClmPkg_Doc2Pdf = True
                sErrorTxt = ""
        ElseIf oWordApp.Version >= 12# Then
            sErrorTxt = "Doc2Pdf: Word does not have 'Save as PDF' component installed"
        Else
            sErrorTxt = "Doc2Pdf: Word 2007 (or newer) is not installed"
        End If
        
        oWordApp.Quit 0 'wdDoNotSaveChanges
        
        Set oWordApp = Nothing
        Set oFso = Nothing
        
        On Error GoTo 0
    End If
End Function


'#
'# Usage:        ClmPkg_Txt2Pdf(sDocFile, sPdfFile, sErrorTxt)
'#
'# Parameters:   sTxtFile:          Full path/file name of source text file
'#               sPdfFile:          Full path/file name of the destination PDF file (that we're converting to)
'#               sErrorTxt:         Error text (returns false if error present as well)
'#
'# Returns:      True if completely successful, False if not.
'#
'# Purpose:      Converts the given plain text file into a PDF file
'#

Public Function ClmPkg_Txt2Pdf(sTxtFile As String, sPdfFile As String, ByRef sErrorTxt As String) As Boolean
    ' sneaky, aren't we??
    ClmPkg_Txt2Pdf = ClmPkg_Doc2Pdf(sTxtFile, sPdfFile, sErrorTxt)
End Function

'#
'# Usage:        ClmPkg_Rtf2Pdf(sRtfFile, sPdfFile, sErrorTxt)
'#
'# Parameters:   sRtfFile:          Full path/file name of source RTF file
'#               sPdfFile:          Full path/file name of the destination PDF file (that we're converting to)
'#               sErrorTxt:         Error text (returns false if error present as well)
'#
'# Returns:      True if completely successful, False if not.
'#
'# Purpose:      Converts the given RTF file into a PDF file
'#

Public Function ClmPkg_Rtf2Pdf(sRtfFile As String, sPdfFile As String, ByRef sErrorTxt As String) As Boolean
    ' sneaky, aren't we??
    ClmPkg_Rtf2Pdf = ClmPkg_Doc2Pdf(sRtfFile, sPdfFile, sErrorTxt)
End Function

'#
'# Usage:        ClmPkg_Htm2Pdf(sHtmFile, sPdfFile, sErrorTxt)
'#
'# Parameters:   sHtmFile:          Full path/file name of source HTML file
'#               sPdfFile:          Full path/file name of the destination PDF file (that we're converting to)
'#               sErrorTxt:         Error text (returns false if error present as well)
'#
'# Returns:      True if completely successful, False if not.
'#
'# Purpose:      Converts the given HTML file into a PDF file
'#

Public Function ClmPkg_Htm2Pdf(sHtmFile As String, sPdfFile As String, ByRef sErrorTxt As String) As Boolean
    ' sneaky, aren't we??
    ClmPkg_Htm2Pdf = ClmPkg_Doc2Pdf(sHtmFile, sPdfFile, sErrorTxt)
End Function

'#
'# Usage:        ClmPkg_Tif2Pdf(sTifFile, sDestPdfFile, sErrorTxt)
'#
'# Parameters:   sTifFile:          Full path/file name of source TIFF file
'#               sDestPdfFile:      Full path/file name of the destination PDF file (that we're converting to)
'#               sErrorTxt:         Error text (returns false if error present as well)
'#
'# Returns:      True if completely successful, False if not.
'#
'# Purpose:      Converts the given TIFF file into a PDF file (non-OCR'd)
'#
'
'Public Function ClmPkg_Tif2Pdf(ByVal sTifFile As String, ByVal sDestPdfFile As String, ByRef sErrorTxt As String) As Boolean
'    Dim oFSO As Object
'    Dim oAcroApp As Object
'    Dim oAcroPdDoc As Object
'    Dim oAcroAvDoc As Object
'    Dim oAcroJSO As Object
'
'    ClmPkg_Tif2Pdf = False
'    sErrorTxt = "Tif2Pdf: Unable to access Adobe Acrobat" ' Default error message
'
'    If ClmPkg_FileInUse(sTifFile) Then
'        sErrorTxt = "Tif2Pdf: Document file aready in use"
'    Else
'
'        On Error Resume Next
'
'        Set oFSO = CreateObject("Scripting.FileSystemObject")
'
'        If oFSO.FileExists(sTifFile) Then
'
'            Set oAcroApp = CreateObject("AcroExch.App")
'            Set oAcroPdDoc = CreateObject("AcroExch.PDDoc")
'
'            If Err.Number = 0 Then
'                Set oAcroAvDoc = CreateObject("AcroExch.AVDoc")
'                'MsgBox "Before PDF Conversion"
'                If oAcroAvDoc.Open(sTifFile, "") Then
'                    'MsgBox "TIFF File should be open now"
'                    Set oAcroPdDoc = oAcroAvDoc.GetPDDoc
'                    'oAcroApp.show
'
'                    'MsgBox oAcroPdDoc.GetFileName()
'
'                    oAcroPdDoc.Save PDSaveFull, sDestPdfFile
'                    'oAcroPdDoc.Save PDSaveCopy, sDestPdfFile
'                    'oAcroPdDoc.Save PDSaveLinearized, sDestPdfFile
'                    'oAcroPdDoc.Save PDSaveCollectGarbage, sDestPdfFile
'
'
'
'                    'Set oAcroJSO = oAcroPdDoc.GetJSObject
'                    'MsgBox oAcroPdDoc.GetFileName()
'
'                    ' "SaveAs" the document into  PDF format & close
'                    'oAcroJSO.SaveAs sDestPdfFile, "com.adobe.acrobat.tiff"
'
'                    oAcroPdDoc.Close
'
'                    ClmPkg_Tif2Pdf = True
'                    sErrorTxt = ""
'
'                    'Set oAcroJSO = Nothing
'
'                    oAcroAvDoc.Close True
'                    oAcroPdDoc.Close True
'                Else
'                    sErrorTxt = "Tif2Pdf: Acrobat could not open the requested file."
'                End If
'            Else
'                sErrorTxt = "Tif2Pdf: Unable to locate Acrobat"
'            End If
'
'            Set oAcroAvDoc = Nothing
'            Set oAcroPdDoc = Nothing
'            Set oAcroApp = Nothing
'
'        Set oFSO = Nothing
'
'        On Error GoTo 0
'
'        End If
'    End If
'End Function

'#
'# Usage:        ClmPkg_Jpg2Pdf(sJpgFile, sDestPdfFile, sErrorTxt)
'#
'# Parameters:   sJpgFile:          Full path/file name of source JPG file
'#               sDestPdfFile:      Full path/file name of the destination PDF file (that we're converting to)
'#               sErrorTxt:         Error text (returns false if error present as well)
'#
'# Returns:      True if completely successful, False if not.
'#
'# Purpose:      Converts the given JPG file into a PDF file
'#

'Public Function ClmPkg_Jpg2Pdf(sJpgFile As String, sDestPdfFile As String, ByRef sErrorTxt As String) As Boolean
'    ClmPkg_Jpg2Pdf = ClmPkg_Tif2Pdf(sJpgFile, sDestPdfFile, sErrorTxt)
'End Function

'#
'# Usage:        ClmPkg_Bmp2Pdf(sJpgFile, sDestPdfFile, sErrorTxt)
'#
'# Parameters:   sBmpFile:          Full path/file name of source BMP file
'#               sDestPdfFile:      Full path/file name of the destination PDF file (that we're converting to)
'#               sErrorTxt:         Error text (returns false if error present as well)
'#
'# Returns:      True if completely successful, False if not.
'#
'# Purpose:      Converts the given BMP file into a PDF file
'#

'Public Function ClmPkg_Bmp2Pdf(sBmpFile As String, sDestPdfFile As String, ByRef sErrorTxt As String) As Boolean
'    ClmPkg_Bmp2Pdf = ClmPkg_Tif2Pdf(sBmpFile, sDestPdfFile, sErrorTxt)
'End Function

'#
'# Usage:        ClmPkg_Eml2Pdf(sEmlFile, sPdfFile, sErrorTxt)
'#
'# Parameters:   sEmlFile:          Full path/file name of source EML file
'#               sPdfFile:          Full path/file name of the destination PDF file (that we're converting to)
'#               sErrorTxt:         Error text (returns false if error present as well)
'#
'# Returns:      True if completely successful, False if not.
'#
'# Purpose:      Converts the given EML file into a PDF file
'#
'
'Public Function ClmPkg_Eml2Pdf(sEmlFile As String, sPdfFile As String, ByRef sErrorTxt As String) As Boolean
'
'    Dim oFSO As Object         'System.FileSystemObject
'    Dim oWordApp As Object     'Word.Application
'    Dim oWordDoc As Object     'Word.Document
'    Dim oWordDocs As Object    'Word.Documents
'    Dim sFolder As String
'    Dim oWordSel As Object
'    Dim oCdoMsg As Object      'CDO.Message
'    Dim oCdoStream As Object
'    Dim oCdoAttach As Object
'    Dim sEmlAttachLst As String
'
'    Dim sTempPlainText As String
'    Dim oPackage As New ClsClmPkgPackageAPI
'    Dim oTxtFile As Object
'
'    ClmPkg_Eml2Pdf = False
'    sErrorTxt = "Eml2Pdf: Unable to access Microsoft Word" ' Default error message
'
'    If ClmPkg_FileInUse(sEmlFile) Then
'        sErrorTxt = "Eml2Pdf: Document file aready in use"
'    Else
'
'        On Error Resume Next
'
'        Set oFSO = CreateObject("Scripting.FileSystemObject")
'        Set oWordApp = CreateObject("Word.Application")
'
'        ' Office Versions:
'        '   12.0 = Office 2007
'        '   11.0 = Office 2003
'        '   10.0 = Office 2002/XP
'        '    9.0 = Office 2000
'
'
'        ' Ensure we're at Word 2007 (at least), and that the "Save as PDF" addon (or SP2) is installed.
'        ' (The referenced dll could change in future Word versions, but there isn't any other method
'        ' exposed to determine if it's installed. Presumably it'll be standard with Office 2010+...)
'
'        If oWordApp.Version >= 12# And _
'           oFSO.FileExists(Environ("commonprogramfiles") & "\Microsoft Shared\OFFICE" & Format(Val(oWordApp.Version), "00") & "\EXP_PDF.DLL") Then
'
'            ' When we use Word to load/convert the EML file, it doesn't present the emails "top-level"
'            ' header info (since that's normally presented in the email bar in Word/Outlook); we use CDO
'            ' to read in the relevant information & then place it at the top of the Word doc.
'
'                Set oCdoMsg = CreateObject("CDO.Message")
'                Set oCdoStream = oCdoMsg.GetStream
'                oCdoStream.LoadFromFile sEmlFile
'                oCdoStream.FLush
'
'                sTempPlainText = ""
'                If Nz(oCdoMsg.HTMLBody, "") = "" Then
'                    ' This is a plaintext email; output the decoded body text & use that for input to Word...
'                    sTempPlainText = oPackage.GetUniqueTempFileName("txt")
'                    Set oTxtFile = oFSO.CreateTextFile(sTempPlainText, True)
'                    oTxtFile.Write oCdoMsg.textbody
'                    oTxtFile.Close
'                    Set oWordDoc = oWordApp.Documents.Open(sTempPlainText)
'                Else
'                    ' This is a html/rtf encoded email; loda it directly into Word
'                    Set oWordDoc = oWordApp.Documents.Open(sEmlFile)
'                End If
'
'                oWordApp.DisplayAlerts = wdAlertsNone
'                With oWordApp.ActiveDocument.PageSetup
'                    .TopMargin = oWordApp.InchesToPoints(0.75)
'                    .LeftMargin = oWordApp.InchesToPoints(0.75)
'                    .RightMargin = oWordApp.InchesToPoints(0.75)
'                    .BottomMargin = oWordApp.InchesToPoints(0.75)
'                End With
'
'                sEmlAttachLst = ""
'                For Each oCdoAttach In oCdoMsg.Attachments
'                    If sEmlAttachLst > "" Then
'                        sEmlAttachLst = sEmlAttachLst & "; "
'                    End If
'                    sEmlAttachLst = sEmlAttachLst & oCdoAttach.FileName
'                Next
'
'                ' Add header into to Word Doc
'                Set oWordSel = oWordApp.Selection
'                oWordSel.HomeKey wdStory, wdMove
'
'                oWordSel.Font.Name = "Tahoma"
'                oWordSel.Font.Size = "10"
'                oWordSel.Font.Color = wdColorBlack
'
'                If oCdoMsg.From > "" Then
'                    oWordSel.Font.Bold = True
'                    oWordSel.TypeText "From: "
'                    oWordSel.Font.Bold = False
'                    oWordSel.TypeText oCdoMsg.From
'                    oWordSel.TypeParagraph
'                End If
'
'                If oCdoMsg.SentOn > "" Then
'                    oWordSel.Font.Bold = True
'                    oWordSel.TypeText "Sent: "
'                    oWordSel.Font.Bold = False
'                    oWordSel.TypeText Format(oCdoMsg.SentOn, "DDDD, MMMM dd, yyyy h:mm AMPM")
'                    oWordSel.TypeParagraph
'                End If
'
'                If oCdoMsg.To > "" Then
'                    oWordSel.Font.Bold = True
'                    oWordSel.TypeText "To: "
'                    oWordSel.Font.Bold = False
'                    oWordSel.TypeText oCdoMsg.To
'                    oWordSel.TypeParagraph
'                End If
'
'                If oCdoMsg.cc > "" Then
'                    oWordSel.Font.Bold = True
'                    oWordSel.TypeText "Cc: "
'                    oWordSel.Font.Bold = False
'                    oWordSel.TypeText oCdoMsg.cc
'                    oWordSel.TypeParagraph
'                End If
'
'                If sEmlAttachLst > "" Then
'                    oWordSel.Font.Bold = True
'                    oWordSel.TypeText "Attach: "
'                    oWordSel.Font.Bold = False
'                    oWordSel.TypeText sEmlAttachLst
'                    oWordSel.TypeParagraph
'                End If
'
'                If oCdoMsg.Subject > "" Then
'                    oWordSel.Font.Bold = True
'                    oWordSel.TypeText "Subject: "
'                    oWordSel.Font.Bold = False
'                    oWordSel.TypeText oCdoMsg.Subject
'                    oWordSel.TypeParagraph
'                End If
'
'                oWordSel.TypeParagraph
'
'                oWordDoc.SaveAs sPdfFile, wdFormatPDF
'                oWordDoc.Close wdDoNotSaveChanges
'
'                ClmPkg_Eml2Pdf = True
'                sErrorTxt = ""
'
'        ElseIf oWordApp.Version >= 12# Then
'            sErrorTxt = "Eml2Pdf: Word does not have 'Save as PDF' component installed"
'
'        Else
'            sErrorTxt = "Eml2Pdf: Word 2007 (or newer) is not installed"
'
'        End If
'
'        oWordApp.Quit wdDoNotSaveChanges
'
'        Set oWordApp = Nothing
'        Set oFSO = Nothing
'
'        On Error GoTo 0
'    End If
'End Function

'#
'# Usage:        ClmPkg_OcrTiff2Tiff(sOrigTiffFile, sDestTiffFile)
'#
'# Parameters:   sOrigTiffFile:     Full path/file name of source TIFF file
'#               sDestTiffFile:     Full path/file name of the destination TIFF file (output from the OCR process)
'#
'# Returns:      True if completely successful, False if not.
'#
'# Purpose:      OCRs the given TIFF file
'#

Public Sub ClmPkg_OcrTiff2Tiff(sOrigTiffFile As String, sDestTiffFile As String)
    Dim oModiDoc As Object
    Dim iPage As Integer
    
    Set oModiDoc = CreateObject("MODI.Document")
    oModiDoc.Create (sOrigTiffFile)
    For iPage = 0 To (oModiDoc.Images.Count - 1)
        oModiDoc.Images(iPage).OCR
    Next
      
    oModiDoc.SaveAs sDestTiffFile
    oModiDoc.Close False 'Don't save changes (just in case...)

    Set oModiDoc = Nothing
End Sub

Public Sub testme()
    Dim sErrorTxt As String
    
    ClmPkg_Bmp2Png "\\ccaintranet.com\DFS-FLD-TS\Users\Karl.Erickson\Desktop\test.bmp", _
                   "\\ccaintranet.com\DFS-FLD-TS\Users\Karl.Erickson\Desktop\test.pdf", _
                   sErrorTxt
                   

End Sub

Public Function ClmPkg_Bmp2Png(sBmpFile As String, sDestPngFile As String, ByRef sErrorTxt As String) As Boolean
    Dim oFso As Object
    Dim oAcroApp As Object
    Dim oAcroPdDoc As Object
    Dim oAcroAvDoc As Object
    Dim oAcroJSO As Object
    
    ClmPkg_Bmp2Png = False
    sErrorTxt = "Bmp2Png: Unable to access Adobe Acrobat" ' Default error message

    If ClmPkg_FileInUse(sBmpFile) Then
        sErrorTxt = "Bmp2Png: Document file aready in use"
    Else
        
        On Error Resume Next
    
        Set oFso = CreateObject("Scripting.FileSystemObject")
        
        If oFso.FileExists(sBmpFile) Then
            
            Set oAcroApp = CreateObject("AcroExch.App")
            Set oAcroPdDoc = CreateObject("AcroExch.PDDoc")
            If Err.Number = 0 Then
                oAcroApp.Hide
                
                Set oAcroAvDoc = CreateObject("AcroExch.AVDoc")
        
                If oAcroAvDoc.Open(sBmpFile, "") Then
                    Set oAcroPdDoc = oAcroAvDoc.GetPDDoc
            
                    Set oAcroJSO = oAcroPdDoc.GetJSObject
                    
                    ' "SaveAs" the document into  PDF format & close
                    oAcroJSO.SaveAs sDestPngFile ', "com.adobe.acrobat.png"
                    oAcroPdDoc.Close
                    
                    ClmPkg_Bmp2Png = True
                    sErrorTxt = ""
                
                    Set oAcroJSO = Nothing
                    
                    oAcroAvDoc.Close True
                Else
                    sErrorTxt = "Bmp2Png: Acrobat could not open the requested file."
                End If
            Else
                sErrorTxt = "Bmp2Png: Unable to locate Acrobat"
            End If
            
            Set oAcroAvDoc = Nothing
            Set oAcroPdDoc = Nothing
            Set oAcroApp = Nothing
            
        Set oFso = Nothing
    
        On Error GoTo 0
    
        End If
    End If
End Function





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
    Open SFileName For Input Lock Read As iFileNum
    Close iFileNum
    lErrNum = Err.Number
    On Error GoTo 0
    
    ClmPkg_FileInUse = (lErrNum = 70) ' 70 = Permission Denied. - File is opened by another user.
    
End Function


'#
'# Usage:        ClmPkg_Tif2Pdf(sTifFile, sDestPdfFile, sErrorTxt)
'#
'# Parameters:   sTifFile:          Full path/file name of source TIFF file
'#               sDestPdfFile:      Full path/file name of the destination PDF file (that we're converting to)
'#               sErrorTxt:         Error text (returns false if error present as well)
'#
'# Returns:      True if completely successful, False if not.
'#
'# Purpose:      Converts the given TIFF file into a PDF file (non-OCR'd)
'#

Public Function ClmPkg_Tif2Pdf(ByVal sTifFile As String, ByVal sDestPdfFile As String, ByRef sErrorTxt As String) As Boolean
    Dim oFso As Object
    Dim oAcroApp As Object
    Dim oAcroPdDoc As Object
    Dim oAcroAvDoc As Object
    Dim oAcroJSO As Object
    
    ClmPkg_Tif2Pdf = False
    sErrorTxt = "Tif2Pdf: Unable to access Adobe Acrobat" ' Default error message

    If ClmPkg_FileInUse(sTifFile) Then
        sErrorTxt = "Tif2Pdf: Document file aready in use"
    Else
        
        On Error Resume Next
    
        Set oFso = CreateObject("Scripting.FileSystemObject")
        
        If oFso.FileExists(sTifFile) Then
            
            Set oAcroApp = CreateObject("AcroExch.App")
            Set oAcroPdDoc = CreateObject("AcroExch.PDDoc")
            
            If Err.Number = 0 Then
                Set oAcroAvDoc = CreateObject("AcroExch.AVDoc")
                'MsgBox "Before PDF Conversion"
                If oAcroAvDoc.Open(sTifFile, "") Then
                    'MsgBox "TIFF File should be open now"
                    Set oAcroPdDoc = oAcroAvDoc.GetPDDoc
                    'oAcroApp.show
                    
                    'MsgBox oAcroPdDoc.GetFileName()
                    
                    oAcroPdDoc.Save PDSaveFull, sDestPdfFile
                    'oAcroPdDoc.Save PDSaveCopy, sDestPdfFile
                    'oAcroPdDoc.Save PDSaveLinearized, sDestPdfFile
                    'oAcroPdDoc.Save PDSaveCollectGarbage, sDestPdfFile
                    
                    
                    
                    'Set oAcroJSO = oAcroPdDoc.GetJSObject
                    'MsgBox oAcroPdDoc.GetFileName()
                    
                    ' "SaveAs" the document into  PDF format & close
                    'oAcroJSO.SaveAs sDestPdfFile, "com.adobe.acrobat.tiff"
                    
                    oAcroPdDoc.Close
                    
                    ClmPkg_Tif2Pdf = True
                    sErrorTxt = ""
                
                    'Set oAcroJSO = Nothing
                    
                    oAcroAvDoc.Close True
                    oAcroPdDoc.Close True
                Else
                    sErrorTxt = "Tif2Pdf: Acrobat could not open the requested file."
                End If
            Else
                sErrorTxt = "Tif2Pdf: Unable to locate Acrobat"
            End If
            
            Set oAcroAvDoc = Nothing
            Set oAcroPdDoc = Nothing
            Set oAcroApp = Nothing
            
        Set oFso = Nothing
    
        On Error GoTo 0
    
        End If
    End If
End Function
Option Compare Database
Option Explicit

Private Const ClassName As String = "mod_Template_Tool_POC"


''' Last Modified: 09/23/2015
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
'''  - Maybe make actual properties instead of using GetField()
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 09/23/2015 - Created class
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


'Private Const cs_PROCESS_FLDR_ROOT As String = "\\ccaintranet.com\dfs-cms-ds\Data\CMS\AnalystFolders\KevinD\BOLD_Letter_Tool\iHT_Takeover\Templates"
Private Const cs_PROCESS_FLDR_ROOT As String = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\AStep Migration\CRS to Conshy\Original Templates"

'
Public Function ReplaceWithFields()
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oSubFld As Scripting.Folder ' for james structure
Dim oFile As Scripting.file
Dim oWordApp As Word.Application
Dim oWordDoc As Word.Document
Dim objWordField As Word.Field
Dim objWordSection As Word.Section
Dim oFooter As Word.HeaderFooter
Dim oRng As Word.Range
Dim oRegEx As RegExp
Dim oMatchs As MatchCollection
Dim oMatch As Match
Dim oAdo As clsADO
Dim oCn As ADODB.Connection
Dim oRs As ADODB.RecordSet
Dim vsLines() As String
Dim iLineIdx As Integer
Dim sLine As String
Dim sExt As String
Dim iFoot As Integer
    
    
    strProcName = ClassName & ".ReplaceWithFields"

    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = GetConnectString("v_Workspace_Database")
        .CursorLocation = adUseClientBatch
        .Open
    End With

    Set oRs = New ADODB.RecordSet
    With oRs
        Set .ActiveConnection = oCn
        .Open "SELECT * FROM KD_iHT_Template_Fields_Investigation WHERE 1 = 2", , adOpenDynamic, adLockBatchOptimistic
        Set .ActiveConnection = Nothing
    End With

    Set oRegEx = New RegExp
    With oRegEx
        .IgnoreCase = True
        .MultiLine = False
        .Global = True
        .Pattern = "[\<]{3}([^\>]+?)[\>]{3}"         ' <<<name>>>
    End With
    
    Set oWordApp = New Word.Application
oWordApp.visible = True

    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder(cs_PROCESS_FLDR_ROOT)


    For Each oSubFld In oFldr.SubFolders
        
        
        For Each oFile In oSubFld.Files
        
            ' steps:
            ' Open the document
            If left(oFile.Name, 1) = "~" Then GoTo NxtFile
            sExt = oFso.GetExtensionName(oFile.Path)
            
            If sExt <> "doc" And sExt <> "docx" Then
                Stop
                GoTo NxtFile
            End If
            Debug.Print oFile.Name
            Set oWordDoc = oWordApp.Documents.Open(oFile.Path)
            ' look for all of the FAKE fields (<<<name>>>)
            ' replace the name
            
            With oWordDoc
    '                            If oWordDoc.Sections.Count > 1 Then
    '                                Debug.Print "More than 1 section!!!"
    '
    '                                Stop
    '                            End If
                For Each oRng In .StoryRanges
                    vsLines() = Split(oRng.Text, vbCr)
                    
                    For iLineIdx = 0 To UBound(vsLines)
                        sLine = vsLines(iLineIdx)
                        Debug.Print sLine
                        
                        Set oMatchs = oRegEx.Execute(sLine)
                        For Each oMatch In oMatchs
                            oRs.AddNew
                            oRs("TemplateFileName").Value = oFile.Name
                            oRs("FieldName").Value = left(oMatch.SubMatches(0), 255)
                            oRs("Section").Value = "StoryRange"
                            oRs.Update
                        Next
                        
                    Next

    
                    ' headers
                    For Each objWordSection In oWordDoc.Sections
                        For iFoot = 1 To objWordSection.Headers.Count
    '                            Stop
                            Set oFooter = objWordSection.Footers.Item(iFoot)
                            vsLines = Split(oFooter.Range.Text, vbCrLf)
                            For iLineIdx = 0 To UBound(vsLines)
                                sLine = vsLines(iLineIdx)
                                Debug.Print sLine
                                Set oMatchs = oRegEx.Execute(sLine)
                                For Each oMatch In oMatchs
                                    oRs.AddNew
                                    oRs("TemplateFileName").Value = oFile.Name
                                    oRs("FieldName").Value = oMatch.SubMatches(0)
                                    oRs("Section").Value = "Header"
                                    oRs.Update
                                Next
    
                            Next
                        Next
                    Next
    '                for each oFooter in oworddoc.Sections.h
                    ' footers
                    For Each objWordSection In oWordDoc.Sections
                        For iFoot = 1 To objWordSection.Footers.Count
    '                            Stop
                            Set oFooter = objWordSection.Footers.Item(iFoot)
                            vsLines = Split(oFooter.Range.Text, vbCrLf)
                            For iLineIdx = 0 To UBound(vsLines)
                                sLine = vsLines(iLineIdx)
                                Debug.Print sLine
                                Set oMatchs = oRegEx.Execute(sLine)
                                For Each oMatch In oMatchs
                                    oRs.AddNew
                                    oRs("TemplateFileName").Value = oFile.Name
                                    oRs("FieldName").Value = oMatch.SubMatches(0)
                                    oRs("Section").Value = "Footer"
                                    oRs.Update
                                Next
    
                            Next
                        Next
                    Next
                    ' anything else to check????
    '                Stop
                Next
            End With
            oWordDoc.Close False
NxtFile:
        Next
NxtBusiness:
    Next

    With oRs
        Set .ActiveConnection = oCn
        .UpdateBatch
        .Close
    End With


Block_Exit:
    
    Set oRs = Nothing
    Set oFile = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing

    If Not oWordApp Is Nothing Then
        For Each oWordDoc In oWordApp.Documents
            oWordDoc.Close False
        Next
        oWordApp.Quit
        Set oWordApp = Nothing
    End If

    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function ProcessFolder_COPY(sFldr As String, Optional oForm As Form_frm_LETTER_Reconciliation_Tool) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oFile As Scripting.file
Dim oWordApp As Word.Application
Dim oWordDoc As Word.Document
Dim objWordField As Word.Field
Dim objWordSection As Word.Section

Dim lSections As Long
Dim i As Integer
Dim oRng As Word.Range
Dim sTxt As String

Dim oRegEx As RegExp
Dim oMatchs As MatchCollection
Dim oMatch As Match
Dim iSubM As Integer
Dim sDocNum As String, iDocNum As Integer
Dim sInstanceId As String
Dim dtStarted As Date

Const lUpdateModulous As Long = 1
Dim lProcessedCnt As Long
Dim sPctDone As String
Dim lMaxCnt As Long
Dim sElapsedTime As String

    strProcName = ClassName & ".ProcessFolder"

    If Not oForm Is Nothing Then
        oForm.Status = ""
        sPctDone = ""
        lProcessedCnt = 0
        lMaxCnt = oForm.MaxToProcess
        dtStarted = Now()
    End If

'    sFldr = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\Scan Ops\Automated Letters\ToPrint\2014-12-03\VADRA"

    Set oDb = CurrentDb
    oDb.Execute "DELETE FROM LETTER_Reconciliation_Tool"
    Set oRs = oDb.OpenRecordSet("SELECT * FROM LETTER_Reconciliation_Tool WHERE 1 = 2")
    
    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder(sFldr)
    
    Set oWordApp = New Word.Application
'oWordApp.Visible = True
    
    Set oRegEx = New RegExp
    With oRegEx
        .Global = False
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "^Re:\s*?(.*?)\s*?\- ([a-zA-Z0-9]+?)$"
    End With
    
    For Each oFile In oFldr.Files
        lProcessedCnt = lProcessedCnt + 1
        DoEvents
        If Not oForm Is Nothing Then
            oForm.CurrentLetterNum = lProcessedCnt

            DoEvents
            If oForm.StopNow = True Then
                GoTo Block_Exit
            End If
            oForm.UpdateStatus
            While oForm.Paused = True
                SleepEvents 2
'                Stop
            Wend
            
        End If
        Set oWordDoc = oWordApp.Documents.Open(oFile.Path)
        sDocNum = left(oFile.Name, 4)
        iDocNum = sDocNum
'Stop
        sInstanceId = Replace(oFile.Path, sFldr & "\", "", , , vbTextCompare)
        sInstanceId = Replace(sInstanceId, sDocNum & "_", "", , , vbTextCompare)
'Stop
'        For Each objWordSection In oWordApp.ActiveDocument.Sections
'            For Each objWordField In objWordSection.Range.Fields
'                Stop
'
'                Debug.Print objWordField.Type
'
'                Stop
'
'            Next
'        Next

        DoEvents

        With oWordDoc
            ' Loop through Story Ranges and update.
            ' Note that this may trigger interactive fields (eg ASK and FILLIN).
            For Each oRng In .StoryRanges
                Do
'                    Stop
'                    Debug.Print oRng.Text
                    sTxt = left(oRng.Text, 2000)
                    Debug.Print sTxt
                    Set oMatchs = oRegEx.Execute(sTxt)
                    
                    If oMatchs.Count = 0 Then
                        Stop ' fix it!
                    End If
                    
                    For Each oMatch In oMatchs
                        Debug.Print oMatch
                        Debug.Print oMatch.SubMatches(0)
                        Debug.Print oMatch.SubMatches(1)
                        oRs.AddNew
                        oRs("DocNum") = iDocNum
                        oRs("DocName") = oFile.Path
                        oRs("InstanceId") = sInstanceId
                        oRs("ProvName") = oMatch.SubMatches(0)
                        oRs("ProvNum") = oMatch.SubMatches(1)
                        oRs.Update
                        GoTo NextDoc
                    Next

'                    oRng.Fields.Unlink
'                    For Each hLink In oRng.Hyperlinks
'                        hLink.Delete
'                    Next
                    Set oRng = oRng.NextStoryRange
                Loop Until oRng Is Nothing

            Next
        End With
NextDoc:
        oWordDoc.Close False
        DoEvents
    Next
    

    
Block_Exit:
    Set oDb = Nothing
    
    Set oRs = Nothing
    Set oFile = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing

    If Not oWordApp Is Nothing Then
        For Each oWordDoc In oWordApp.Documents
            oWordDoc.Close False
        Next
        oWordApp.Quit
        Set oWordApp = Nothing
    End If

    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
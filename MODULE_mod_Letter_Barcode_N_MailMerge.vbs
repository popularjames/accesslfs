Option Compare Database
Option Explicit

Private Const ClassName As String = "mod_Letter_Barcode_Example_Code"


Public Sub SamplePseudoWordMailMergeCode()
'
'    ' Set data source for mail merge.  Data will be from new Temp Table
'    objWordDoc.MailMerge.OpenDataSource Name:=pstrODCFile, _
'                        SQLStatement:="exec usp_LETTER_Get_Info '" & oLetterInst.InstanceID & "'"
'
'
'    ' Perform mail merge.
'    objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
'    objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
'    objWordDoc.MailMerge.Execute Pause:=False
'    If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
'        objWordApp.visible = True
'        MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
'        bMergeError = True
'        objWordApp.ActiveDocument.Activate
'        strErrMsg = "Error encountered with mail merge."
'        GoTo Block_Err
'    End If
'
'    ' 20130219 KD: Add the Sec Pages field in the footer for barcodes
'    Call AddSecPagesCode(objWordApp.ActiveDocument)
'
'
'    ' Save the output doc
'    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
'    strOutputPath = pstrOutputBasePath & "\" & pstrProvNum & "\"
'    Call CreateFolders(strOutputPath)
'
'    If Not FolderExists(strOutputPath) Then
'        strErrMsg = "Provider folder for letter was not created for instance: " & oLetterInst.InstanceID & vbNewLine + vbNewLine & "Process will continue for the instances left."
'        GoTo Block_Err
'    End If
'
'    If oLetterInst.LetterQueueStatus = "R" Then
'    'If pstrInstanceStatus = "R" Then
'        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-Reprint-" & oLetterInst.InstanceID & ".doc"
'    Else
'        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-" & oLetterInst.InstanceID & ".doc"
'    End If
'
'    objWordMergedDoc.spellingchecked = True
'    objWordMergedDoc.Repaginate
'
'    If UnlinkWordFields(objWordApp, objWordMergedDoc) = False Then
'        LogMessage strProcName, "ERROR", "There was an error unlinking the fields. Check that the fields are correct!", pstrOutputFileName, True
'    End If
'
'    objWordMergedDoc.SaveAs pstrOutputFileName
'    SleepEvents 1
'
'    With oLetterInst
'        If .LetterBatchId = 0 Then
'            .LetterBatchId = Me.MostRecentBatchId
'        End If
'        If objWordMergedDoc.BuiltInDocumentProperties(14) = 1 Then
'            Stop
'        End If
'        .PageCount = objWordMergedDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
'        .LetterPath = pstrOutputFileName
'        .SaveStaticDetails
'    End With
'
'    objWordMergedDoc.Close
'
'    Set objWordMergedDoc = Nothing
'
'    If Not FileExists(pstrOutputFileName) Then
'        strErrMsg = "PrintLetterInstance: Letter was not created for instance " & oLetterInst.InstanceID & vbNewLine + vbNewLine & "Process will continue for the instances left."
'        LogMessage strProcName, "ERROR", "Generated letter does not exist where it should", pstrOutputFileName
'
'        GoTo Block_Err
'    End If
'
'
'Block_Exit:
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''
''''' Required functions
'''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''' This function will look for a bookmark named 'SecPages' and will replace that with
''' the Sec Pages field
'Public Function AddSecPagesCode(objWordDoc As Object) As Boolean
''Private Function AddSecPagesCode(objWordDoc As Word.Document) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim objWordField As FormField
'
''Dim objWordSection As Object
'Dim objWordSection As Object
'Dim iFoot As Integer
'Dim oFooter As Object
'Dim oField As Object
'Dim oRange As Object
'Dim sBookmarkName As String
'Dim saryBkmarks(1) As String
'Dim iBkmark As Integer
'
'
'    strProcName = ClassName & ".AddSecPagesCode"
'
'    saryBkmarks(0) = "SecPages"
'    saryBkmarks(1) = "SecPages2"
'    DoEvents
'    DoEvents
'    DoEvents
'    SleepEvents 1
'
'    For iBkmark = 0 To UBound(saryBkmarks)
'        sBookmarkName = saryBkmarks(iBkmark)
'
'        If IsBookMark(objWordDoc, sBookmarkName) = False Then
'            ' nothing else to do
'            GoTo NextBkmark
'        End If
'
'        For Each objWordSection In objWordDoc.Sections
'            For iFoot = 1 To objWordSection.Footers.Count
'                    ' shape range should take care of the text box..
'                    '' Note: may want to do the Headers here too
'                Set oFooter = objWordSection.Footers.Item(iFoot)
'                    ' Looks like 2007 + use a different means..
'                If oFooter.Shapes.Count > 0 Then
'
'                    If IsBookMark(objWordDoc, sBookmarkName) = False Then
'                        ' nothing else to do
'                        GoTo NextBkmark
'                    End If
'
'                    Set oRange = objWordDoc.Bookmarks(sBookmarkName).Range
'
'                    Set oField = oFooter.Shapes(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
'
'                    ' now lets remove the bookmark all together - I've been finding examples where this is replaced several times
'                    ' resulting in barcodes like 01030303
'                    ' when it should be 0103
''                    oField.Unlink
'                    objWordDoc.Bookmarks(sBookmarkName).Delete
'
'                ElseIf oFooter.Range.ShapeRange.Count > 0 Then
'
'                    If IsBookMark(objWordDoc, sBookmarkName) = False Then
'                        ' nothing else to do
'                        GoTo NextBkmark
'                    End If
''                    oField.Unlink
'                    objWordDoc.Bookmarks(sBookmarkName).Delete
'
'
'                    Set oRange = objWordDoc.Bookmarks(sBookmarkName).Range
'                    Set oField = oFooter.Range.ShapeRange.Item(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
'
'
'
'                End If
'            Next
'        Next
'NextBkmark:
'    Next
'
'    AddSecPagesCode = True  ' In this case true = no error.. :)
'
'Block_Exit:
'    Set oField = Nothing
'    Set oRange = Nothing
'    Set oFooter = Nothing
'    Set objWordSection = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function
'
'
'
'Public Function UnlinkWordFields(oWordApp As Word.Application, oDoc As Word.Document, Optional sLetterType As String = "") As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim objWordField As Word.Field
'Dim objWordSection As Word.Section
'Dim i As Integer
'
'    strProcName = ClassName & ".UnlinkWordFields"
'
'    ' 20130219 KD: Make sure that the section pages start at 1
'
'    oDoc.Repaginate
'    SleepEvents 1
'    DoEvents
'    DoEvents
'    DoEvents
'
'    With oDoc
'
'
'        For i = 1 To .Sections.Count
'            .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
'            .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
'            .Repaginate
'            .Sections(i).Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
'            .Sections(i).Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
'            .Repaginate
'
'            '' how about shapes?
'            Dim oShape As Word.Shape
'            'For Each oShape In .Sections(i).Footers(wdHeaderFooterPrimary).Shapes
'             '   oShape.TextFrame.TextRange.Fields.Unlink
'
'            'Next
'
''            For Each oShape In .Sections(i).headers(wdHeaderFooterPrimary).Shapes
''                oShape.TextFrame.TextRange.Fields.Unlink
''            Next
'
'        Next i
'        '.Fields.Unlink
'    End With
'
'    oDoc.Activate
'
'        '' Hardcoded (shame) for QR barcodes: need to make this data driven at some point..
'    If sLetterType <> "VADRA_QR" Then
'    ' by the way, this breaks the ADR footer's Page X of Y (even though it doesn't break the
''        For Each objWordSection In oWordApp.ActiveDocument.Sections
''            For i = 1 To objWordSection.Footers.Count
''                For Each objWordField In objWordSection.Footers.Item(i).Range.Fields
''                    Debug.Print objWordField.Code
''                    objWordField.Update
''                    objWordField.Unlink
''                Next
''
''            Next
''        Next
'
'    Else
'        '' this should be unlinking the bar codes
'        For Each objWordSection In oWordApp.ActiveDocument.Sections
'            For Each objWordField In objWordSection.Range.Fields
'                '            objWordField.Update
'                objWordField.Unlink
'            Next
'        Next
'    End If
'
'    Dim oRng As Word.Range, hLink As Word.Hyperlink
'
'    With oDoc
'        ' Loop through Story Ranges and update.
'        ' Note that this may trigger interactive fields (eg ASK and FILLIN).
'        For Each oRng In .StoryRanges
'            Do
'                oRng.Fields.Unlink
'                For Each hLink In oRng.Hyperlinks
'                    hLink.Delete
'                Next
'                Set oRng = oRng.NextStoryRange
'            Loop Until oRng Is Nothing
'        Next
'    End With
'
'
'
'      UnlinkWordFields = True
'Block_Exit:
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function


' For the life of me, I don't know who keeps changing my code to late bound
' I mean, it's not difficult to go from one version of word to another and the
' benefits of early bound outweigh any kind of other issues (unless you have
' mixed users of course - we don't here in CMS!!!!)
'Public Function IsBookMark(objWordDoc As Object, sBookmarkName As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oBkmk As Object
'
'    strProcName = ClassName & ".IsBookMark"
'
'    For Each oBkmk In objWordDoc.Bookmarks
'        If UCase(oBkmk.Name) = UCase(sBookmarkName) Then
'            IsBookMark = True
'            GoTo Block_Exit
'        End If
'    Next
'
'Block_Exit:
'    Set oBkmk = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function
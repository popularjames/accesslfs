Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#########################################################################################################
'#                                                                                                       #
'# Module Name:          Form_MUT_CodeUpgradeManager                                                     #
'#                                                                                                       #
'# Description:          Provides an interface to Upgrade naming conventions introduced by Decipher 3.0  #
'#                                                                                                       #
'#                                                                                                       #
'# Original Author:      Karl Erickson           (08/13/2010)                                            #
'# Last Update By:       Lino Gomes              (09/06/2012)                                            #
'#                                                                                                       #
'# Change History:       [#] [MM/DD/YYYY]  [Author Name]    [Explanation of Change]                      #
'#                       --- ------------  ---------------  -------------------------------------------- #
'#                       000 08/13/2010    Karl Erickson    Created                                      #
'#                       001 07/26/2012    Karl Erickson    Updated for Decipher 3.0/AppSource           #
'#                       001 09/06/2012    Lino Gomes       Updated for Decipher 3.0/AppSource           #
'#                                                                                                       #
'#########################################################################################################

Option Compare Database
Option Explicit

' Display mode of status bar at bottom of main GUI dialog.
Private Enum MUT_STATUS
    MUT_STATUS_NORMAL = 0
    MUT_STATUS_ERROR = 1
End Enum

Private Sub SetProgress(iPctComplete As Integer, Optional sProgressText As String, Optional iBarType As MUT_STATUS)
    Dim lBarColor As Long
    
    If iPctComplete < 0 Then
        iPctComplete = 0
    ElseIf iPctComplete > 100 Then
        iPctComplete = 100
    End If
    
    If IsMissing(sProgressText) Then
        sProgressText = ""
    End If

    If IsMissing(iBarType) Or iBarType = MUT_STATUS_NORMAL Then
        lBarColor = 16756398
    ElseIf iBarType = MUT_STATUS_ERROR Then
        lBarColor = RGB(255, 119, 119)
    End If
    
    If iPctComplete > 0 Then
        lblProgress.Caption = sProgressText
        
        bxStatusFill.Width = bxStatusBorder.Width * (iPctComplete / 100#)
        bxStatusFill.BackColor = lBarColor
        
        lblProgress.visible = True
        bxStatusFill.visible = True
    Else
        lblProgress.visible = False
        bxStatusFill.visible = False
        
        lblProgress.Caption = ""
        bxStatusFill.Width = bxStatusBorder.Width
        
    End If

    DoEvents
End Sub


Private Sub UpgradeCodeReferences(bPreviewOnly As Boolean)
On Error GoTo ErrorHappened
    Dim oRs As DAO.RecordSet
    Dim bFound As Boolean
    Dim lStartLine As Long
    Dim lStartCol As Long
    Dim lEndLine As Long
    Dim lEndCol As Long
    Dim sNewLineTxt As String
    Dim sOldLineTxt As String
    Dim sTypeName As String
    Dim sObjName As String
    Dim iObjType As ObjType
    Dim objComponent As Object
    Dim bModuleEdited As Boolean
    Dim iCt As Integer
    Dim iTotal As Integer
    
    ' First thing: Backup!!!
    

    iTotal = Application.VBE.ActiveVBProject.Collection.Item("DecipherCoreTemplate").VBComponents.Count
    iCt = 0
    

    For Each objComponent In Application.VBE.ActiveVBProject.Collection.Item("DecipherCoreTemplate").VBComponents
        iCt = iCt + 1
        
        If bPreviewOnly Then
            SetProgress (iCt / iTotal * 100#), "Scanning VBA: " & objComponent.Name & "...", MUT_STATUS_NORMAL
        Else
            SetProgress (iCt / iTotal * 100#), "Updating VBA: " & objComponent.Name & "...", MUT_STATUS_NORMAL
        End If
        
        'Exclude AppSource Apps
        Set oRs = CurrentDb.OpenRecordSet("Select Count(1) as Ct From MUT_CodeUpgradeExclude WHERE InStr(1, '" & objComponent.Name & "', ObjectName)", , dbOpenSnapshot)
        oRs.MoveFirst
        bFound = oRs!CT
        oRs.Close
        
        If bFound = 0 Then
            sTypeName = ""
            bModuleEdited = False
           
            Select Case objComponent.Type
                Case 1, 2 'Module / Class
                    iObjType = objModule
                    sTypeName = "Module"
                Case 100 '
                    If InStr(objComponent.Name, "Form_") = 1 Then
                        iObjType = objForm
                        sTypeName = "Form"
                        sObjName = Mid(objComponent.Name, 6)
                    ElseIf InStr(objComponent.Name, "Report") = 1 Then
                        iObjType = objReport
                        sTypeName = "Report"
                        sObjName = Mid(objComponent.Name, 8)
                    End If
            End Select
            
            Set oRs = CurrentDb.OpenRecordSet("SELECT * FROM MUT_CodeUpgrade ORDER BY Ordinal", , dbOpenSnapshot)
            oRs.MoveFirst
            
            bFound = True
            DoEvents
           
            Do Until oRs.EOF
                With objComponent.CodeModule
            
                    lStartLine = 1
                    lStartCol = 1
                    lEndLine = .CountOfLines
                    lEndCol = -1
                    
                    bFound = True
            
                    Do Until bFound = False
                        bFound = .Find(oRs!FindString, lStartLine, lStartCol, lEndLine, lEndCol)
                        
                        If bFound Then
                            sNewLineTxt = ""
                            lStartLine = lEndLine
                            lStartCol = lEndCol
                            lEndLine = .CountOfLines
                            lEndCol = -1
            
                            sOldLineTxt = .Lines(lStartLine, 1)
                            
                            'Look for Whole Word Only
                            If oRs!WholeWordOnly = -1 Then
                                sNewLineTxt = FindWholeWordOnly(sOldLineTxt, oRs!FindString, oRs!ReplaceString)
                            Else
                                sNewLineTxt = Replace(sOldLineTxt, oRs!FindString, Nz(oRs!ReplaceString, ""))
                            End If
                            
                            If sNewLineTxt > "" Then
                            
                                If Not bPreviewOnly Then
                                    .ReplaceLine lStartLine, sNewLineTxt
                                    bModuleEdited = True
                                    On Error Resume Next
                                    Select Case sTypeName
                                        Case "Module"
                                            DoCmd.Save acModule, sObjName
                                        Case "Form"
                                            DoCmd.Save acForm, sObjName
                                        Case "Report"
                                            DoCmd.Save acReport, sObjName
                                    End Select
                                    On Error GoTo ErrorHappened
                                End If
            
                                lstCodeChanges.AddItem sTypeName & ";" & .Name & ";" & lStartLine & ";" & MUT_ListQuote(sOldLineTxt) & ";" & MUT_ListQuote(Replace(sNewLineTxt, vbCrLf, "  ")) & ";"
                                
                            End If
                        End If
                    Loop
            
                End With
                oRs.MoveNext
            Loop
            oRs.Close
        End If
        If Not bPreviewOnly Then
            On Error Resume Next
            Select Case sTypeName
                Case "Form"
                    DoCmd.Close acForm, sObjName, acSaveYes
                Case "Report"
                    DoCmd.Close acReport, sObjName, acSaveYes
            End Select
            On Error GoTo ErrorHappened
        End If
    Next

    
ExitNow:
    On Error Resume Next
    oRs.Close
    Set oRs = Nothing
    Set objComponent = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description
    Resume ExitNow
End Sub

Private Sub UpgradeFormProperties(bPreviewOnly As Boolean)
On Error GoTo ErrorHappened
    
    Dim oRs As DAO.RecordSet
    Dim rs As DAO.RecordSet
    Dim bFound As Boolean
    Dim SQL As String
    Dim frm As Form
    Dim ctrl As Control
    Dim sOldLineTxt As String
    Dim sNewLineTxt As String
    Dim iCt As Integer
    Dim iTotal As Integer

    
    SQL = "SELECT Switch([Type]=5,1,[Type]=1,0,[Type]=-32766,4,[Type]=-32764,3,[Type]=-32761,5,[Type]=-32768,2) AS ObjectType, " & _
            "Switch([Type]=5,'QUERY',[Type]=1,'TABLE',[Type]=-32766,'MACRO',[Type]=-32764,'REPORT',[Type]=-32761,'MODULE',[Type]=-32768,'FORM') AS ObjectTypeName, " & _
            "O2.Name AS ObjectName, (Select Count(1) as Ct From MUT_CodeUpgradeExclude WHERE InStr(1, O2.Name, ObjectName)) AS ct " & _
            "FROM MSysObjects AS O2 " & _
            "WHERE ((((Select Count(1) as Ct From MUT_CodeUpgradeExclude WHERE InStr(1, O2.Name, ObjectName)))=0) AND ((O2.Type) In (-32768)) AND ((O2.Name) Not Like '~sq*' And (O2.Name) Not Like 'MSys*'))" & _
            "ORDER BY O2.Type, O2.Name;"

    Set rs = CurrentDb.OpenRecordSet(SQL, , dbOpenSnapshot)
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveLast
        rs.MoveFirst
        iTotal = rs.recordCount
        
        Do Until rs.EOF
            iCt = iCt + 1
            
            If bPreviewOnly Then
                SetProgress (iCt / iTotal * 100#), "Scanning Form: " & rs!ObjectName & "...", MUT_STATUS_NORMAL
            Else
                SetProgress (iCt / iTotal * 100#), "Updating Form: " & rs!ObjectName & "...", MUT_STATUS_NORMAL
            End If
            
            DoCmd.OpenForm rs!ObjectName, acDesign, , , acFormPropertySettings, acHidden
            Set frm = Forms(rs!ObjectName)
            For Each ctrl In frm.Controls
                With ctrl
                    Select Case .ControlType
                        Case acSubform
                            sOldLineTxt = ctrl.Properties("SourceObject")
                            If Nz(sOldLineTxt, vbNullString) <> vbNullString Then
                                Set oRs = CurrentDb.OpenRecordSet("SELECT * FROM MUT_CodeUpgrade WHERE ObjectPropertyOnly = -1 ORDER BY Ordinal", , dbOpenSnapshot)
                                oRs.MoveFirst
                                Do Until oRs.EOF
                                    bFound = True
                                    
                                    If oRs!WholeWordOnly = -1 Then
                                        sNewLineTxt = FindWholeWordOnly(sOldLineTxt, oRs!FindString, oRs!ReplaceString)
                                    Else
                                        sNewLineTxt = Replace(sOldLineTxt, oRs!FindString, Nz(oRs!ReplaceString, ""))
                                    End If
                            
                                    If sOldLineTxt <> sNewLineTxt And Nz(sNewLineTxt, vbNullString) <> vbNullString Then
                                        If Not bPreviewOnly Then
                                            ctrl.Properties("SourceObject") = sNewLineTxt
                                            DoCmd.Save acForm, rs!ObjectName
                                        End If
                                        lstCodeChanges.AddItem rs!ObjectTypeName & ";" & rs!ObjectName & "." & .Name & ";-;" & MUT_ListQuote(oRs!FindString) & ";" & MUT_ListQuote(Replace(Nz(oRs!ReplaceString, ""), vbCrLf, "  ")) & ";"
                                    End If
                                    oRs.MoveNext
                                Loop
                                oRs.Close
                            End If
                        Case acComboBox, acListBox
                            sOldLineTxt = ctrl.RowSource
                            If Nz(sOldLineTxt, vbNullString) <> vbNullString Then
                                Set oRs = CurrentDb.OpenRecordSet("SELECT * FROM MUT_CodeUpgrade WHERE ObjectPropertyOnly = -1 ORDER BY Ordinal", , dbOpenSnapshot)
                                oRs.MoveFirst
                                Do Until oRs.EOF
                                    bFound = True
                                    
                                    If oRs!WholeWordOnly = -1 Then
                                        sNewLineTxt = FindWholeWordOnly(sOldLineTxt, oRs!FindString, oRs!ReplaceString)
                                    Else
                                        sNewLineTxt = Replace(sOldLineTxt, oRs!FindString, Nz(oRs!ReplaceString, ""))
                                    End If
                            
                                    If sOldLineTxt <> sNewLineTxt And Nz(sNewLineTxt, vbNullString) <> vbNullString Then
                                        If Not bPreviewOnly Then
                                            ctrl.ctrl.RowSource = sNewLineTxt
                                            DoCmd.Save acForm, rs!ObjectName
                                        End If
                                        lstCodeChanges.AddItem rs!ObjectTypeName & ";" & rs!ObjectName & "." & .Name & ";-;" & MUT_ListQuote(oRs!FindString) & ";" & MUT_ListQuote(Replace(Nz(oRs!ReplaceString, ""), vbCrLf, "  ")) & ";"
                                    End If
                                    oRs.MoveNext
                                Loop
                                oRs.Close
                            End If
                    End Select
                End With
            Next ctrl
                        
            If Nz(frm.RecordSource, vbNullString) <> vbNullString Then
                sOldLineTxt = frm.RecordSource
                
                Set oRs = CurrentDb.OpenRecordSet("SELECT * FROM MUT_CodeUpgrade WHERE ObjectPropertyOnly = -1 ORDER BY Ordinal", , dbOpenSnapshot)
                oRs.MoveFirst
                Do Until oRs.EOF
                    If oRs!WholeWordOnly = -1 Then
                        sNewLineTxt = FindWholeWordOnly(sOldLineTxt, oRs!FindString, oRs!ReplaceString)
                    Else
                        sNewLineTxt = Replace(sOldLineTxt, oRs!FindString, Nz(oRs!ReplaceString, ""))
                    End If
                    
                    If sOldLineTxt <> sNewLineTxt And Nz(sNewLineTxt, vbNullString) <> vbNullString Then
                        If Not bPreviewOnly Then
                            frm.RecordSource = sNewLineTxt
                            On Error Resume Next
                            DoCmd.Save acForm, rs!ObjectName
                            On Error GoTo ErrorHappened
                        End If
                        lstCodeChanges.AddItem rs!ObjectTypeName & ";" & rs!ObjectName & ".RecordSource;-;" & MUT_ListQuote(oRs!FindString) & ";" & MUT_ListQuote(Replace(Nz(oRs!ReplaceString, ""), vbCrLf, "  ")) & ";"
                    End If
                    oRs.MoveNext
                Loop
                oRs.Close
            End If
                        

                
            On Error Resume Next
            If Not bPreviewOnly Then
                DoCmd.Close acForm, rs!ObjectName, acSaveYes
            Else
                DoCmd.Close acForm, rs!ObjectName, acSaveNo
            End If
            On Error GoTo ErrorHappened
            rs.MoveNext
        Loop
    End If
    
ExitNow:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    oRs.Close
    Set oRs = Nothing
    Set ctrl = Nothing
    Set frm = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description
    Resume ExitNow
End Sub

Private Sub UpgradeReportProperties(bPreviewOnly As Boolean)
On Error GoTo ErrorHappened
    
    Dim oRs As DAO.RecordSet
    Dim rs As DAO.RecordSet
    Dim bFound As Boolean
    Dim SQL As String
    Dim Rpt As Report
    Dim ctrl As Control
    Dim sOldLineTxt As String
    Dim sNewLineTxt As String
    Dim iCt As Integer
    Dim iTotal As Integer

    
    SQL = "SELECT Switch([Type]=5,1,[Type]=1,0,[Type]=-32766,4,[Type]=-32764,3,[Type]=-32761,5,[Type]=-32768,2) AS ObjectType, " & _
            "Switch([Type]=5,'QUERY',[Type]=1,'TABLE',[Type]=-32766,'MACRO',[Type]=-32764,'REPORT',[Type]=-32761,'MODULE',[Type]=-32768,'FORM') AS ObjectTypeName, " & _
            "O2.Name AS ObjectName, (Select Count(1) as Ct From MUT_CodeUpgradeExclude WHERE InStr(1, O2.Name, ObjectName)) AS ct " & _
            "FROM MSysObjects AS O2 " & _
            "WHERE ((((Select Count(1) as Ct From MUT_CodeUpgradeExclude WHERE InStr(1, O2.Name, ObjectName)))=0) AND ((O2.Type) In (-32764)) AND ((O2.Name) Not Like '~sq*' And (O2.Name) Not Like 'MSys*'))" & _
            "ORDER BY O2.Type, O2.Name;"

    Set rs = CurrentDb.OpenRecordSet(SQL, , dbOpenSnapshot)
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveLast
        rs.MoveFirst
        iTotal = rs.recordCount
        
        Do Until rs.EOF
            iCt = iCt + 1
            
            If bPreviewOnly Then
                SetProgress (iCt / iTotal * 100#), "Scanning Report: " & rs!ObjectName & "...", MUT_STATUS_NORMAL
            Else
                SetProgress (iCt / iTotal * 100#), "Updating Report: " & rs!ObjectName & "...", MUT_STATUS_NORMAL
            End If
            
            DoCmd.OpenReport rs!ObjectName, acDesign, , , acFormPropertySettings, acHidden
            Set Rpt = Reports(rs!ObjectName)
            For Each ctrl In Rpt.Controls
                With ctrl
                    Select Case .ControlType
                        Case acSubform
                            sOldLineTxt = ctrl.Properties("SourceObject")
                            
                            If Nz(sOldLineTxt, vbNullString) <> vbNullString Then
                                Set oRs = CurrentDb.OpenRecordSet("SELECT * FROM MUT_CodeUpgrade WHERE ObjectPropertyOnly = -1 ORDER BY Ordinal", , dbOpenSnapshot)
                                oRs.MoveFirst
                                Do Until oRs.EOF
                                    bFound = True
                                    
                                    If oRs!WholeWordOnly = -1 Then
                                        sNewLineTxt = FindWholeWordOnly(sOldLineTxt, oRs!FindString, oRs!ReplaceString)
                                    Else
                                        sNewLineTxt = Replace(sOldLineTxt, oRs!FindString, Nz(oRs!ReplaceString, ""))
                                    End If
                            
                                    If sOldLineTxt <> sNewLineTxt And Nz(sNewLineTxt, vbNullString) <> vbNullString Then
                                        If Not bPreviewOnly Then
                                            ctrl.Properties("SourceObject") = sNewLineTxt
                                            DoCmd.Save acReport, rs!ObjectName
                                        End If
                                        lstCodeChanges.AddItem rs!ObjectTypeName & ";" & rs!ObjectName & "." & .Name & ";-;" & MUT_ListQuote(oRs!FindString) & ";" & MUT_ListQuote(Replace(Nz(oRs!ReplaceString, ""), vbCrLf, "  ")) & ";"
                                    End If
                                    oRs.MoveNext
                                Loop
                                oRs.Close
                            End If
                    End Select
                End With
            Next ctrl
                        
            If Nz(Rpt.RecordSource, vbNullString) <> vbNullString Then
                sOldLineTxt = Rpt.RecordSource
                
                Set oRs = CurrentDb.OpenRecordSet("SELECT * FROM MUT_CodeUpgrade WHERE ObjectPropertyOnly = -1 ORDER BY Ordinal", , dbOpenSnapshot)
                oRs.MoveFirst
                Do Until oRs.EOF
                    If oRs!WholeWordOnly = -1 Then
                        sNewLineTxt = FindWholeWordOnly(sOldLineTxt, oRs!FindString, oRs!ReplaceString)
                    Else
                        sNewLineTxt = Replace(sOldLineTxt, oRs!FindString, Nz(oRs!ReplaceString, ""))
                    End If
                    
                    If sOldLineTxt <> sNewLineTxt And Nz(sNewLineTxt, vbNullString) <> vbNullString Then
                        If Not bPreviewOnly Then
                            Rpt.RecordSource = sNewLineTxt
                            On Error Resume Next
                            DoCmd.Save acReport, rs!ObjectName
                            On Error GoTo ErrorHappened
                        End If
                        lstCodeChanges.AddItem rs!ObjectTypeName & ";" & rs!ObjectName & ".RecordSource;-;" & MUT_ListQuote(oRs!FindString) & ";" & MUT_ListQuote(Replace(Nz(oRs!ReplaceString, ""), vbCrLf, "  ")) & ";"
                    End If
                    oRs.MoveNext
                Loop
                oRs.Close
            End If
                        

                
            On Error Resume Next
            If Not bPreviewOnly Then
                DoCmd.Close acReport, rs!ObjectName, acSaveYes
            Else
                DoCmd.Close acReport, rs!ObjectName, acSaveNo
            End If
            On Error GoTo ErrorHappened
            rs.MoveNext
        Loop
    End If
    
ExitNow:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    oRs.Close
    Set oRs = Nothing
    Set ctrl = Nothing
    Set Rpt = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description
    Resume ExitNow
End Sub

Private Sub UpgradeQueryProperties(bPreviewOnly As Boolean)
On Error GoTo ErrorHappened
    
    Dim oRs As DAO.RecordSet
    Dim rs As DAO.RecordSet
    Dim bFound As Boolean
    Dim SQL As String
    Dim Qdef As QueryDef
    Dim sOldLineTxt As String
    Dim sNewLineTxt As String
    Dim iCt As Integer
    Dim iTotal As Integer
    Dim QrySql As String
    
    SQL = "SELECT Switch([Type]=5,1,[Type]=1,0,[Type]=-32766,4,[Type]=-32764,3,[Type]=-32761,5,[Type]=-32768,2) AS ObjectType, " & _
            "Switch([Type]=5,'QUERY',[Type]=1,'TABLE',[Type]=-32766,'MACRO',[Type]=-32764,'REPORT',[Type]=-32761,'MODULE',[Type]=-32768,'FORM') AS ObjectTypeName, " & _
            "O2.Name AS ObjectName, (Select Count(1) as Ct From MUT_CodeUpgradeExclude WHERE InStr(1, O2.Name, ObjectName)) AS ct " & _
            "FROM MSysObjects AS O2 " & _
            "WHERE ((((Select Count(1) as Ct From MUT_CodeUpgradeExclude WHERE InStr(1, O2.Name, ObjectName)))=0) AND ((O2.Type) In (5)) AND ((O2.Name) Not Like '~sq*' And (O2.Name) Not Like 'MSys*'))" & _
            "ORDER BY O2.Type, O2.Name;"

    iCt = 0
    Set rs = CurrentDb.OpenRecordSet(SQL, , dbOpenSnapshot)
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveLast
        rs.MoveFirst
        iTotal = rs.recordCount
    
        Do Until rs.EOF
            iCt = iCt + 1
            
            If bPreviewOnly Then
                SetProgress (iCt / iTotal * 100#), "Scanning Query: " & rs!ObjectName & "...", MUT_STATUS_NORMAL
            Else
                SetProgress (iCt / iTotal * 100#), "Updating Query: " & rs!ObjectName & "...", MUT_STATUS_NORMAL
            End If
            
            Set Qdef = CurrentDb.QueryDefs(rs!ObjectName)
            sOldLineTxt = Qdef.SQL
            
            Set oRs = CurrentDb.OpenRecordSet("SELECT * FROM MUT_CodeUpgrade WHERE ObjectPropertyOnly = -1 ORDER BY Ordinal", , dbOpenSnapshot)
            oRs.MoveFirst
            
            Do Until oRs.EOF
                bFound = True
                
                    If oRs!WholeWordOnly = -1 Then
                        sNewLineTxt = FindWholeWordOnly(sOldLineTxt, oRs!FindString, oRs!ReplaceString)
                    Else
                        sNewLineTxt = Replace(sOldLineTxt, oRs!FindString, Nz(oRs!ReplaceString, ""))
                    End If
                    
                    If sOldLineTxt <> sNewLineTxt And Nz(sNewLineTxt, vbNullString) <> vbNullString Then
                        If Not bPreviewOnly Then
                            QrySql = sNewLineTxt
                            Qdef.SQL = QrySql
                            On Error Resume Next
                            DoCmd.Save acQuery, rs!ObjectName
                            On Error GoTo ErrorHappened
                        End If
                        
                        lstCodeChanges.AddItem rs!ObjectTypeName & ";" & rs!ObjectName & ".SQL;-;" & MUT_ListQuote(oRs!FindString) & ";" & MUT_ListQuote(Replace(Nz(oRs!ReplaceString, ""), vbCrLf, "  ")) & ";"
                    End If
                    
                oRs.MoveNext
            Loop
            oRs.Close
            rs.MoveNext
        Loop
    End If

ExitNow:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    oRs.Close
    Set oRs = Nothing
    Set Qdef = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description
    Resume ExitNow
End Sub

Function BackupDatabase() As Boolean
On Error GoTo ErrorHappened

    Dim fso As Object
    Dim SourceDB As String
    Dim DestDB As String
    
    SourceDB = CurrentDb.Name
    DestDB = Replace(SourceDB & "_Backup_" & Format(Now, "yyyymmddhhmmss"), ".accdb", vbNullString) & ".accdb"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile SourceDB, DestDB
    

    BackupDatabase = True
    
ExitNow:
    On Error Resume Next
    Set fso = Nothing
    Exit Function
ErrorHappened:
    BackupDatabase = False
    Resume ExitNow
End Function


Private Sub ChkForm_AfterUpdate()
If lblProgress.Caption = "Update Complete. Please review and save the updated code." Then
    Me.cmdPreview.Enabled = True
End If
End Sub

Private Sub ChkQueries_AfterUpdate()
If lblProgress.Caption = "Update Complete. Please review and save the updated code." Then
    Me.cmdPreview.Enabled = True
End If
End Sub

Private Sub ChkReport_AfterUpdate()
If lblProgress.Caption = "Update Complete. Please review and save the updated code." Then
    Me.cmdPreview.Enabled = True
End If
End Sub

Private Sub ChkVba_AfterUpdate()
If lblProgress.Caption = "Update Complete. Please review and save the updated code." Then
    Me.cmdPreview.Enabled = True
End If
End Sub

'Private Sub cmdAbout_Click()
'    cmdClose.SetFocus
'    CommandBars(PCLM_MENU_USAGE_ABOUT).ShowPopup 'cmdAbout.Left, cmdAbout.Top + cmdAbout.Height
'End Sub
'
'Public Sub SetUpContextMenus()
'    Dim oBar As Object
'
'    On Error Resume Next ' ignore error if command bar does not exist to be deleted
'    CommandBars(PCLM_MENU_USAGE_ABOUT).Delete
'
'    On Error GoTo 0
'
'    '
'    ' About...
'    '
'
'    Set oBar = CommandBars.Add(Name:=PCLM_MENU_USAGE_ABOUT, Position:=msoBarPopup)
'    With oBar.Controls.Add(Type:=msoControlButton)
'        .Tag = PCLM_CTX_ABOUT_HELP
'        .Caption = "About Paperless Claims..."
'        .FaceID = 487
'        .onAction = "=PCLM_InvokeContextMenu(" & .Tag & ")"
'    End With
'
'End Sub


Private Sub cmdClose_Click()
    DoCmd.Close acForm, Form.Name
End Sub

Private Sub ExportLogFile(ByVal FileName As String, ByRef Lst As listBox)
'Write screen import log to file
On Error GoTo ErrorHappened
    Dim fso As Object
    Dim fts As Object
    Dim r As Integer
    Dim c As Integer
    Dim temp As String

    If (Not Lst.ColumnHeads And Lst.ListCount > 0) Or (Lst.ColumnHeads And Lst.ListCount > 1) Then
        Application.Echo False
        Set fso = CreateObject("Scripting.Filesystemobject")
        '2=Overwrite, 8=Append
        Set fts = fso.OpenTextFile(FileName, 2, True)
        
        'Header
        fts.WriteLine String(Len(CurrentProject.FullName) + 2, "*")
        fts.WriteLine "* " & Now
        fts.WriteLine "* " & Identity.UserName
        fts.WriteLine "* " & CurrentProject.FullName
        fts.WriteLine String(Len(CurrentProject.FullName) + 2, "*")
        fts.WriteLine
        
        'Write listbox to file
        For r = 0 To Lst.ListCount - 1
            temp = vbNullString
            
            'Build string based on column count
            For c = 0 To Lst.ColumnCount - 1
                temp = temp & Lst.Column(c, r) & vbTab
            Next c
            
            fts.WriteLine temp
        Next r
        
        MsgBox "The log was exported to:" & vbCrLf & FileName, vbInformation, "Log export complete"
    Else
        MsgBox "There isn't any information in the list to export.", vbInformation, "Nothing to export"
    End If

ExitNow:
On Error Resume Next
    fts.Close
    Set fts = Nothing
    Set fso = Nothing
    Application.Echo True
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error exporting log file"
    Resume ExitNow
End Sub

Private Sub cmdExport_Click()
    ExportLogFile CurrentProject.Path & "\Code Upgrade Log.xls", lstCodeChanges
End Sub

Private Sub cmdPreview_Click()

    CmdClose.SetFocus
    cmdPreview.Enabled = False
    CmdExport.Enabled = False
    
    If ChkVba + ChkForm + ChkReport + ChkQueries < 0 Then
        DoCmd.Hourglass True
        
        lstCodeChanges.RowSource = ""
        lstCodeChanges.AddItem "Type; Name; Line#; Old Code; New Code;"

        If Me.ChkVba Then
            UpgradeCodeReferences True
        End If
        If Me.ChkForm Then
            UpgradeFormProperties True
        End If
        If Me.ChkReport Then
            UpgradeReportProperties True
        End If
        If Me.ChkQueries Then
            UpgradeQueryProperties True
        End If
        
        SetProgress 100, "Scan Complete", MUT_STATUS_NORMAL
        
        DoCmd.Hourglass False
    Else
        MsgBox "No selections have been selected to scan."
    End If
    
    cmdPreview.Enabled = True
    CmdUpdate.Enabled = (lstCodeChanges.ListCount > 1) ' 1=header row
    CmdExport.Enabled = (lstCodeChanges.ListCount > 1) ' 1=header row

End Sub

Private Sub cmdUpdate_Click()
    Dim iRet As Integer
    
    CmdClose.SetFocus
    cmdPreview.Enabled = False
    CmdUpdate.Enabled = False
    CmdExport.Enabled = False

    If ChkVba + ChkForm + ChkReport + ChkQueries < 0 Then
    
        iRet = MsgBox("Would you like to create a backup of the current database before perfoming any code updates?" & vbCrLf & vbCrLf & "(Hint: 'Yes' is usually a good answer....)", vbYesNo + vbQuestion + vbDefaultButton1, Form.Caption)
        
        If iRet = vbYes Then
            BackupDatabase
        End If
        
        DoCmd.Hourglass True
        
        lstCodeChanges.RowSource = ""
        lstCodeChanges.AddItem "Type; Name; Line#; Old Code; New Code;"
        
        If Me.ChkVba Then
            UpgradeCodeReferences False
        End If
        If Me.ChkForm Then
            UpgradeFormProperties False
        End If
        If Me.ChkReport Then
            UpgradeReportProperties False
        End If
        If Me.ChkQueries Then
            UpgradeQueryProperties False
        End If
        
        SetProgress 100, "Update Complete. Please review and save the updated code.", MUT_STATUS_NORMAL
    
        DoCmd.Hourglass False
    Else
        MsgBox "No selections have been selected to scan."
    End If
   
    CmdExport.Enabled = (lstCodeChanges.ListCount > 1) ' 1=header row
    Me.cmdPreview.Enabled = True

End Sub

Private Sub Form_Load()
    cmdPreview.Enabled = True
    CmdUpdate.Enabled = False
    
    lstCodeChanges.RowSource = ""
    lstCodeChanges.AddItem "Type; Name; Line#; Old Code; New Code;"
    
    'SetUpContextMenus
End Sub

Private Sub Form_Open(Cancel As Integer)
    Telemetry.RecordOpen "Form", Me.Name, "v1.07"
End Sub

Private Function MUT_ListQuote(sValue As String) As String
    
    sValue = Nz(sValue, "")
    sValue = Replace(sValue, """", """""")  'Escape out existing double-quotes
    
    If sValue > "" Then
        sValue = """" & sValue & """" 'Wrap in double-quotes if value present
    End If
    
    MUT_ListQuote = sValue
    
End Function

Private Function FindWholeWordOnly(sOldLineTxt As String, sFindString As String, sReplaceString As String) As String
    Dim pos As Long
    Dim bLWholeWord As Boolean
    Dim bRWholeWord As Boolean
    
    bLWholeWord = False
    bRWholeWord = False
    pos = InStr(1, sOldLineTxt, sFindString)
    If pos > 1 Then
        Select Case Asc(Mid(sOldLineTxt, pos - 1, 1))
            Case 0, 10, 13, 32, 33, 34, 39, 46, 91
                bLWholeWord = True
            Case Else
                bLWholeWord = False
        End Select
    ElseIf pos = 1 Then
        bLWholeWord = True
    End If
                                        
    If (pos + Len(sFindString) - 1) < Len(sOldLineTxt) Then
        Select Case Asc(Mid(sOldLineTxt, pos + Len(sFindString), 1))
            Case 0, 10, 13, 32, 33, 34, 39, 46, 93
                bRWholeWord = True
            Case Else
                bRWholeWord = False
        End Select
    ElseIf (pos + Len(sFindString) - 1) = Len(sOldLineTxt) Then
        bRWholeWord = True
    End If
    
    If bLWholeWord = True And bRWholeWord = True Then
        FindWholeWordOnly = Replace(sOldLineTxt, sFindString, Nz(sReplaceString, ""))
    Else
        FindWholeWordOnly = ""
    End If

End Function

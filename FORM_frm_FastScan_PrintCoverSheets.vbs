Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim mstrCalledFrom As String

Dim mstrMatchINPath As String
Dim mstrMatchOUTPath As String
Dim mstrSplitINPath As String
Dim mstrSplitOUTPath As String
Dim mstrTIFViewerPath As String
Dim mstrAcrobatPath As String



Private Sub cmdPrintCoverSheets_Click()
Dim strDefaultPrinter As String

If Me.lstChosenFolder.ListIndex = -1 Then
    MsgBox "You must select a Folder first! Cannot continue.", vbInformation, "Error: No Folder Selected"
    Exit Sub
End If

If Me.lstChosenFolder = "Error" Then
    MsgBox "You must select a CoverSheet type first!.", vbExclamation, "Error"
    Exit Sub
End If

Sleep 1000 'making sure the generated Coversheet numbers are not equal to the second

' get current default printer.
strDefaultPrinter = Application.Printer.DeviceName

' switch to printer of your choice:
Set Application.Printer = Application.Printers(lstPrinters.Value)

DoCmd.OpenReport "rpt_FastScan_PrintCoverSheets", acViewNormal



Set Application.Printer = Application.Printers(strDefaultPrinter)


End Sub

Private Sub Form_Load()
On Error Resume Next


    If Me.OpenArgs() & "" <> "" Then
        mstrCalledFrom = Me.OpenArgs
    End If

    If Nz(gintAccountID, 0) = 0 Then
        MsgBox "There is not a currently selected Account ID! Cannot continue.", vbInformation, "Error: Account not selected"
        DoCmd.Close acForm, Me.Name
        Exit Sub
    End If
    
    mstrMatchINPath = "" & DLookup("MatchINPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)
    mstrMatchOUTPath = "" & DLookup("MatchOUTPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)
    mstrSplitINPath = "" & DLookup("SplitINPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)
    mstrSplitOUTPath = "" & DLookup("SplitOUTPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)
    mstrTIFViewerPath = "" & DLookup("TIFViewerPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)
    mstrAcrobatPath = "" & DLookup("AcrobatPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)
    
    If Nz(mstrMatchINPath, "") = "" Or Nz(mstrMatchOUTPath, "") = "" Or Nz(mstrTIFViewerPath, "") = "" Or Nz(mstrAcrobatPath, "") = "" Then
        MsgBox "One or more of the FastScan config values are not setup for this Account. Please check the FastScan_Config table.", vbInformation, "Error with FastScan config values"
        DoCmd.Close acForm, Me.Name
        Exit Sub
    End If

    If Right$(mstrMatchINPath, 1) <> "\" Then mstrMatchINPath = mstrMatchINPath & "\"
    If Right$(mstrMatchOUTPath, 1) <> "\" Then mstrMatchOUTPath = mstrMatchOUTPath & "\"
    If Right$(mstrMatchINPath, 1) <> "\" Then mstrMatchINPath = mstrMatchINPath & "\"
    If Right$(mstrMatchINPath, 1) <> "\" Then mstrMatchINPath = mstrMatchINPath & "\"
    If Right$(mstrMatchINPath, 1) <> "\" Then mstrMatchINPath = mstrMatchINPath & "\"
    If Right$(mstrMatchINPath, 1) <> "\" Then mstrMatchINPath = mstrMatchINPath & "\"

    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim ErrorReturned As String
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
        .SQLTextType = sqltext
        .sqlString = "select distinct FolderName, FolderPriority from FastScanMaint.v_FastScan_UserAuthFolders where accountid = " & gintAccountID & " and Printable = 1 order by FolderName"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            MsgBox "There are no printable FastScan folders setup for this account or you don't have permission!" & vbNewLine & vbNewLine & "Cannot continue.", vbInformation, "Error: FastScan folders missing"
            DoCmd.Close acForm, Me.Name
            GoTo Cleanup
            Exit Sub
        End If
        'Me.cmbFoldersType = Me.cmbIssueType.ItemData(1)
    End With
    
    Set Me.lstChosenFolder.RecordSet = oRs
    
    
'    If Me.lstChosenFolder.RecordSet Is Nothing Then
'        MsgBox "There are no FastScan folders setup for this account! Cannot continue.", vbInformation, "Error: FastScan folders missing"
'        DoCmd.Close acForm, Me.name
'        Exit Sub
'    ElseIf Me.lstChosenFolder.RecordSet.EOF And Me.lstChosenFolder.RecordSet.BOF Then
'        MsgBox "There are no FastScan folders setup for this account! Cannot continue.", vbInformation, "Error: FastScan folders missing"
'        DoCmd.Close acForm, Me.name
'        Exit Sub
'    End If
    
    'Me.lstChosenFolder.value = Me.lstChosenFolder.ItemData(0)
    
    AuditName.Caption = UCase(Nz(DLookup("ClientName", "Admin_Account_config", "accountid = " & gintAccountID), "ERROR"))
    
   
    'MsgBox (gintAccountID)
    Dim PrinterCounter As Integer
    For PrinterCounter = 0 To Application.Printers.Count - 1
        Me.lstPrinters.AddItem Application.Printers(PrinterCounter).DeviceName
    Next
    Me.lstPrinters = Application.Printer.DeviceName

Cleanup:
    oAdo.DisConnect
    Set oAdo = Nothing
    Set oRs = Nothing
    
End Sub

Private Sub Form_Close()
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
End Sub

Private Sub txtNumberOfCoverSheets_AfterUpdate()
    If Me.txtNumberOfCoverSheets > 1000 Then Me.txtNumberOfCoverSheets = 1000
End Sub


   

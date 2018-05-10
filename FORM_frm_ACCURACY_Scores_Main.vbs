Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const mstrUserID = "james.segura"

Dim objExcelApp As Excel.Application
Dim wb As Excel.Workbook
Dim ws As Excel.Worksheet
Dim mstrSessionID As String
Dim mstrAuditType As String
Dim mstrAuditScope As String
Dim mstrResultVersion As String
Dim mstrSourceExcelFileName As String
Dim mstrClaimTableVer As String
Dim mstrClaimDetailTableVer As String
Dim mstrNewIssueTableVer As String
Dim mstrNewIssueDetailTableVer As String
Dim mstrClaimDetailColumnHeader As String
Dim mstrNewIssueColumnHeader As String
Dim mstrNewIssueDetailColumnHeader As String



Private Sub Command0_Click()
    If ReadyToStart Then
        UpdateProcessStatus "Starting..."
        Initialize
        ProcessDataWorkbook
    End If
End Sub
Sub Initialize()
    Set objExcelApp = CreateObject("Excel.Application")
    mstrSessionID = ""
    mstrAuditType = ""
    mstrAuditScope = ""
    mstrResultVersion = ""
    mstrSourceExcelFileName = ""
    mstrClaimTableVer = "NotFound"
    mstrClaimDetailTableVer = "NotFound"
    mstrNewIssueTableVer = "NotFound"
    mstrNewIssueDetailTableVer = "NotFound"
    mstrClaimDetailColumnHeader = ""
    mstrNewIssueColumnHeader = ""
    mstrNewIssueDetailColumnHeader = ""
    ClearTmpTable
End Sub

Sub ProcessDataWorkbook()

    Sleep (1000)
    mstrSessionID = Format$(Now(), "yyyymmddhhnnss")
    
    mstrSourceExcelFileName = txtSourceFileName  '"September 2011.xlsx"
    mstrAuditType = Me.lstAuditType '"Accuracy"
    mstrAuditScope = Me.txtAuditScope ' "2011-09"
    mstrResultVersion = Me.lstResultVersion ' "Score1"
    
    UpdateProcessStatus "Starting load for SessionID = " & mstrSessionID
    
    Dim CurrCellX As Integer
    Dim CurrCellY As Integer
    Dim CurrReportRowNum As Integer
    Dim strClaimColumnHeader As String
    Dim fs As FileSystemObject
    
    UpdateProcessStatus "Check file exists: " & mstrSourceExcelFileName
    
    If Not FileExists(Me.txtSourceFileFolder & "\" & mstrSourceExcelFileName) Then
        GoTo ExitWithError
    End If
    
    If Not PreLoadValidate Then
        GoTo ExitWithError
    End If
     
    Set wb = objExcelApp.Workbooks.Open(Me.txtSourceFileFolder & "\" & mstrSourceExcelFileName)
    Set ws = wb.Sheets(1)
    
    CurrCellX = 1
    CurrCellY = 1
    CurrReportRowNum = 0

    'record header information
    If Not RecordToTable("Header", "AuditType", mstrAuditType, 0) Then GoTo ExitWithError
    If Not RecordToTable("Header", "AuditScope", mstrAuditScope, 0) Then GoTo ExitWithError
    If Not RecordToTable("Header", "ResultVersion", mstrResultVersion, 0) Then GoTo ExitWithError
    If Not RecordToTable("Header", "SourceExcelFileName", mstrSourceExcelFileName, 0) Then GoTo ExitWithError
    
    UpdateProcessStatus "Searching for Header Fields"
    
    'RA NAME
    Dim bolRANameFound As Boolean
    bolRANameFound = False
    
    If ScanForText("RA Name", 1, 1, 100, 100, CurrCellX, CurrCellY) Then
        bolRANameFound = True
    ElseIf ScanForText("RAC Name", 1, 1, 100, 100, CurrCellX, CurrCellY) Then
        bolRANameFound = True
    ElseIf ScanForText("Recovery Auditor Name", 1, 1, 100, 100, CurrCellX, CurrCellY) Then
        bolRANameFound = True
    End If
    If bolRANameFound Then
        If Len(ws.Cells(CurrCellY, CurrCellX)) > 50 Then
            If Not RecordToTable("Cell", "RAName", ws.Cells(CurrCellY, CurrCellX), 0) Then GoTo ExitWithError
        Else
            If Not RecordToTable("Cell", "RAName", NextCellRight(CurrCellX, CurrCellY), 0) Then GoTo ExitWithError
        End If
    Else
        GoTo ExitWithError
    End If
    
    'PackageID
    If ScanForText("Package ID", 1, 1, 100, 100, CurrCellX, CurrCellY) Then
        If Len(ws.Cells(CurrCellY, CurrCellX)) > 50 Then
            If Not RecordToTable("Cell", "PackageID", ws.Cells(CurrCellY, CurrCellX), 0) Then GoTo ExitWithError
        Else
            If Not RecordToTable("Cell", "PackageID", NextCellRight(CurrCellX, CurrCellY), 0) Then GoTo ExitWithError
        End If
    Else
        GoTo ExitWithError
    End If
    
    'Date Submitted
    If ScanForText("Date Submitted", 1, 1, 100, 100, CurrCellX, CurrCellY) Then
        If Len(ws.Cells(CurrCellY, CurrCellX)) > 50 Then
            If Not RecordToTable("Cell", "DateSubmitted", ws.Cells(CurrCellY, CurrCellX), 0) Then GoTo ExitWithError
        Else
            If Not RecordToTable("Cell", "DateSubmitted", NextCellRight(CurrCellX, CurrCellY), 0) Then GoTo ExitWithError
        End If
    Else
        GoTo ExitWithError
    End If
    
    'Accuracy Review number
    If ScanForText("Accuracy Review #", 1, 1, 100, 100, CurrCellX, CurrCellY) Then
        If Len(ws.Cells(CurrCellY, CurrCellX)) > 50 Then
            If Not RecordToTable("Cell", "AccuracyReview", ws.Cells(CurrCellY, CurrCellX), 0) Then GoTo ExitWithError
        Else
            If Not RecordToTable("Cell", "AccuracyReview", NextCellRight(CurrCellX, CurrCellY), 0) Then GoTo ExitWithError
        End If
    Else
        GoTo ExitWithError
    End If
    
    'Report Summary
    If ScanForText("Report Summary", 1, 1, 100, 100, CurrCellX, CurrCellY) Then
        If Not RecordToTable("Cell", "ReportSummary", ws.Cells(CurrCellY, CurrCellX), 0) Then GoTo ExitWithError
    Else
        GoTo ExitWithError
    End If
    
    'Agreement Rate
    If ScanForText(" Rate", 1, 1, 100, 500, CurrCellX, CurrCellY) Then
        If LoadAgreementRateValues(CurrCellX, CurrCellY) Then
            
        Else
            GoTo ExitWithError
        End If
    Else
        GoTo ExitWithError
    End If
    
    If Not FullScanReconnaissance Then GoTo ExitWithError
    
    'record version information
    If Not RecordToTable("Version", "ClaimTableVer", mstrClaimTableVer, 0) Then GoTo ExitWithError
    If Not RecordToTable("Version", "ClaimDetailTableVer", mstrClaimDetailTableVer, 0) Then GoTo ExitWithError
    If Not RecordToTable("Version", "NewIssueTableVer", mstrNewIssueTableVer, 0) Then GoTo ExitWithError
    If Not RecordToTable("Version", "NewIssueDetailTableVer", mstrNewIssueDetailTableVer, 0) Then GoTo ExitWithError
    
    'main claim table
    strClaimColumnHeader = "Claim Number"
    
    If ScanForText(strClaimColumnHeader, 1, 1, 100, 500, CurrCellX, CurrCellY) Then
    
'        If Not DetermineClaimTableVersion(CurrCellX, CurrCellY) Then
'            GoTo ExitWithError
'        End If
    
        If Not LoadMainClaimTable(CurrCellX, CurrCellY) Then
            GoTo ExitWithError
        End If
    Else
        GoTo ExitWithError
    End If

    'load claim detail table
    If mstrClaimDetailTableVer <> "NotFound" Then
        If mstrClaimDetailColumnHeader <> "" Then
            If ScanForText(mstrClaimDetailColumnHeader, 1, CurrCellY, 100, 500, CurrCellX, CurrCellY) Then
        '        If Not DetermineClaimDetailTableVersion(CurrCellX, CurrCellY) Then
        '            GoTo ExitWithError
        '        End If
                If Not LoadDetailClaimTable(CurrCellX, CurrCellY) Then
                    GoTo ExitWithError
                End If
            Else
                GoTo ExitWithError
            End If
        Else
            UpdateProcessStatus "Empty Claim Detail Header Master for table version = " & mstrClaimDetailTableVer
            GoTo ExitWithError
        End If
    End If
    
    'load concept table
    If mstrNewIssueTableVer <> "NotFound" Then 'not all have issue table
        If mstrNewIssueColumnHeader <> "" Then
            If ScanForText(mstrNewIssueColumnHeader, 1, CurrCellY, 100, 500, CurrCellX, CurrCellY) Then
    '            If Not DetermineNewIssueTableVersion(CurrCellX, CurrCellY) Then
    '                GoTo ExitWithError
    '            End If
                If Not LoadNewIssueTable(CurrCellX, CurrCellY) Then
                    GoTo ExitWithError
                End If
            Else
                GoTo ExitWithError
            End If
        Else
            UpdateProcessStatus "Empty Issue Column Header Master for table version = " & mstrNewIssueTableVer
            GoTo ExitWithError
        End If
    End If
    
    'load concept detail table
    If mstrNewIssueDetailTableVer <> "NotFound" Then
        If mstrNewIssueDetailColumnHeader <> "" Then 'not all have issue table
            If ScanForText(mstrNewIssueDetailColumnHeader, 1, CurrCellY, 25, 500, CurrCellX, CurrCellY) Then
    '            If Not DetermineNewIssueTableVersion(CurrCellX, CurrCellY) Then
    '                GoTo ExitWithError
    '            End If
                If Not LoadNewIssueDetailTable(CurrCellX, CurrCellY) Then
                    GoTo ExitWithError
                End If
            Else
                GoTo ExitWithError
            End If
        Else
            UpdateProcessStatus "Empty Issue Detail Column Header Master for table version = " & mstrNewIssueDetailTableVer
            GoTo ExitWithError
        End If
    End If
    
   
    If Not LoadedDataValidation Then
        GoTo ExitWithError
    Else
        If Not SaveResultsToTables Then
            GoTo ExitWithError
        End If
    End If

    UpdateProcessStatus "Done with file: " & Me.txtSourceFileName
    
    ClearFormFields clearFileName:=True

ExitSub:
    
    'Close the workbook
    Set ws = Nothing
    If Not wb Is Nothing Then
        wb.Close False
    End If
    Set wb = Nothing
    objExcelApp.Quit
    Set objExcelApp = Nothing

    Exit Sub

ExitWithError:
    MsgBox "Error occurred", vbExclamation, "Error"
    GoTo ExitSub

End Sub

Function LoadNewIssueTable(CurrCellX As Integer, CurrCellY As Integer) As Boolean

On Error GoTo ErrorHandler

    Dim FoundAtCellX As Integer
    Dim FoundAtCellY As Integer

    Dim strReportRow As String
    Dim strNINumber As String
    Dim strNIClaims As String
    Dim strNIReceived As String
    Dim strNIReviewed As String
    Dim strNISource As String
    Dim intMasterCellFontX As Integer
    Dim intMasterCellFontY As Integer
    
    Dim bolFoundData As Boolean
    Dim bolSavedData As Boolean
    Dim iRows As Integer
    
    bolFoundData = False
    bolSavedData = False
    strReportRow = "0"
    UpdateProcessStatus "Load Concept Table"
    
    Call NextCellDown(CurrCellX, CurrCellY, FoundAtCellY)
    intMasterCellFontX = CurrCellX
    intMasterCellFontY = FoundAtCellY
    
    'here we just go down the cells looking for the correct words
    For iRows = FoundAtCellY To FoundAtCellY + 300
        strNINumber = Nz(ws.Cells(iRows, CurrCellX).Text, "")
        If CellBelongsToGroup(CurrCellX, iRows, intMasterCellFontX, intMasterCellFontY) Then
            If Not bolFoundData Then
                bolFoundData = True
            End If
        Else
            If Not bolFoundData Then
                GoTo ErrorHandler
            Else
                Exit For
            End If
        End If
        
        If mstrNewIssueTableVer = "2015-04" Then
        
            strNIClaims = NextCellRight(CurrCellX, iRows, FoundAtCellX)
            strNIReceived = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strNIReviewed = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strNISource = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
        
            If Not RecordToTable("NewIssue", "NINumber", strNINumber, iRows) Then GoTo ErrorHandler
            If Not RecordToTable("NewIssue", "NIClaims", strNIClaims, iRows) Then GoTo ErrorHandler
            If Not RecordToTable("NewIssue", "NIReceived", strNIReceived, iRows) Then GoTo ErrorHandler
            If Not RecordToTable("NewIssue", "NIReviewed", strNIReviewed, iRows) Then GoTo ErrorHandler
            If Not RecordToTable("NewIssue", "NISource", strNISource, iRows) Then GoTo ErrorHandler
            
            bolSavedData = True
            
        End If
        
    Next
    
    CurrCellY = iRows 'update the current Y cell
    
    If bolSavedData Then LoadNewIssueTable = True
    
    Exit Function
    
ErrorHandler:
    
    LoadNewIssueTable = False


End Function

Function CellBelongsToGroup(CurrCellX As Integer, CurrCellY As Integer, MasterCellFontX As Integer, MasterCellFontY As Integer) As Boolean

On Error GoTo ErrorHandler

    Dim strCurrCellValue As String
    Dim strCurrCellFontName As String
    Dim intCurrCellFontSize As Integer
    Dim strCurrCellFontStyle As String
    Dim intCurrCellCellInteriorColor As Integer
    Dim strMasterCellValue As String
    Dim strMasterCellFontName As String
    Dim intMasterCellFontSize As Integer
    Dim strMasterCellFontStyle As String
    Dim intMasterCellCellInteriorColor As Integer
    
    
    With ws.Cells(CurrCellY, CurrCellX)
        strCurrCellValue = Nz(.Text, "")
        strCurrCellFontName = Nz(.Font.Name, "")
        intCurrCellFontSize = Nz(.Font.Size, 0)
        strCurrCellFontStyle = Nz(.Font.FontStyle, "")
        intCurrCellCellInteriorColor = Nz(.Interior.ColorIndex, 0)
    End With
    
    With ws.Cells(MasterCellFontY, MasterCellFontX)
        strMasterCellValue = Nz(.Text, "")
        strMasterCellFontName = Nz(.Font.Name, "")
        intMasterCellFontSize = Nz(.Font.Size, 0)
        strMasterCellFontStyle = Nz(.Font.FontStyle, "")
        intMasterCellCellInteriorColor = Nz(.Interior.ColorIndex, 0)
    End With
    
    ' And Not (intCurrCellCellInteriorColor = intMasterCellCellInteriorColor))
    If ((strCurrCellValue = "" Or strCurrCellValue = "0") And Not (strMasterCellValue = "" Or strMasterCellValue = "0")) _
        Or (strCurrCellFontName <> strMasterCellFontName And Not (strCurrCellValue = "" Or strCurrCellValue = "0")) _
        Or (intCurrCellFontSize > intMasterCellFontSize And Abs(intCurrCellFontSize - intMasterCellFontSize) > 1) Then 'Or strCurrCellFontStyle <> strMasterCellFontStyle
        GoTo ErrorHandler
    End If
    
    CellBelongsToGroup = True
    
    Exit Function
    
ErrorHandler:
    
    CellBelongsToGroup = False
    
End Function


Function LoadNewIssueDetailTable(CurrCellX As Integer, CurrCellY As Integer) As Boolean

On Error GoTo ErrorHandler

    Dim FoundAtCellX As Integer
    Dim FoundAtCellY As Integer

    Dim strReportRow As String
    Dim strNINumber As String
    Dim strNIName As String
    Dim strNIRecommendation As String

    Dim bolFoundData As Boolean
    Dim bolSavedData As Boolean
    Dim iRows As Integer
    
    Dim intMasterCellFontX As Integer
    Dim intMasterCellFontY As Integer
    
    bolFoundData = False
    bolSavedData = False
    strReportRow = 0
    
    UpdateProcessStatus "Load Concept Detail Table"
    
    Call NextCellDown(CurrCellX, CurrCellY, FoundAtCellY)
    
    intMasterCellFontX = CurrCellX
    intMasterCellFontY = FoundAtCellY
    
    'here we just go down the cells looking for the correct words
    For iRows = FoundAtCellY To FoundAtCellY + 300
        strNINumber = Nz(ws.Cells(iRows, CurrCellX).Text, "")
        If CellBelongsToGroup(CurrCellX, iRows, intMasterCellFontX, intMasterCellFontY) Then
            If Not bolFoundData Then
                bolFoundData = True
            End If
        Else
            If Not bolFoundData Then
                GoTo ErrorHandler
            Else
                Exit For
            End If
        End If
        
        If mstrNewIssueDetailTableVer = "2015-04" Then
        
            strNIName = NextCellRight(CurrCellX, iRows, FoundAtCellX)
            strNIRecommendation = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
        
            If Not RecordToTable("NewIssueDetail", "NINumber", strNINumber, iRows) Then GoTo ErrorHandler
            If Not RecordToTable("NewIssueDetail", "NIName", strNIName, iRows) Then GoTo ErrorHandler
            If Not RecordToTable("NewIssueDetail", "NIRecommendation", strNIRecommendation, iRows) Then GoTo ErrorHandler
            
            bolSavedData = True
            
        End If
        
    Next
    
    CurrCellY = iRows 'update the current Y cell
    
    If bolSavedData Then LoadNewIssueDetailTable = True
    
    Exit Function
    
ErrorHandler:
    
    LoadNewIssueDetailTable = False


End Function



Function LoadDetailClaimTable(CurrCellX As Integer, CurrCellY As Integer) As Boolean

On Error GoTo ErrorHandler

    Dim FoundAtCellX As Integer
    Dim FoundAtCellY As Integer

    Dim strReportRow As String
    Dim strClaimNumber As String
    Dim str2ndLevel As String
    Dim str3rdLevel As String
    Dim strRVCQA As String
    Dim strRVCRationale As String
    Dim strRVCFinalDet As String
    

    Dim bolFoundData As Boolean
    Dim bolSavedData As Boolean
    Dim iRows As Integer
    Dim intMasterCellFontX As Integer
    Dim intMasterCellFontY As Integer
    
    Dim objCellFont As Excel.Font
    
    bolFoundData = False
    bolSavedData = False
    strReportRow = 0
    
    UpdateProcessStatus "Load Detail Claim Table"
    
    Call NextCellDown(CurrCellX, CurrCellY, FoundAtCellY)
    intMasterCellFontX = CurrCellX
    intMasterCellFontY = FoundAtCellY
    
    'here we just go down the cells looking for the correct words
    For iRows = FoundAtCellY To FoundAtCellY + 300
    
        strClaimNumber = Nz(ws.Cells(iRows, CurrCellX).Text, "")
        
        If CellBelongsToGroup(CurrCellX, iRows, intMasterCellFontX, intMasterCellFontY) Then
            If Not bolFoundData Then
                bolFoundData = True
            End If
        Else
            If Not bolFoundData Then
                GoTo ErrorHandler
            Else
                Exit For
            End If
        End If
        
        'we need to get the row number from the report, this will be used later to match to the second claim table
        strReportRow = NextCellLeft(CurrCellX, iRows)
        'the row number must be numeric
        If Not IsNumeric(strReportRow) Then
            GoTo ErrorHandler
        End If
        
        If mstrClaimDetailTableVer = "2011-09" Then
        
            str2ndLevel = NextCellRight(CurrCellX, iRows, FoundAtCellX, NotEmpty:=False)
            str3rdLevel = NextCellRight(FoundAtCellX, iRows, FoundAtCellX, NotEmpty:=False)
            strRVCQA = NextCellRight(FoundAtCellX, iRows, FoundAtCellX, NotEmpty:=False)
            strRVCRationale = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
        
            If Not RecordToTable("ClaimDetail", "ClaimNumber", strClaimNumber, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("ClaimDetail", "2ndLevel", str2ndLevel, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("ClaimDetail", "3rdLevel", str3rdLevel, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("ClaimDetail", "RVCQA", strRVCQA, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("ClaimDetail", "RVCRationale", strRVCRationale, CInt(strReportRow)) Then GoTo ErrorHandler
            
            bolSavedData = True
            
        End If
        
        If mstrClaimDetailTableVer = "2015-04" Or mstrClaimDetailTableVer = "2014-09" Then
        
            strRVCFinalDet = NextCellRight(CurrCellX, iRows, FoundAtCellX, NotEmpty:=False)
            strRVCQA = NextCellRight(FoundAtCellX, iRows, FoundAtCellX, NotEmpty:=False)
            strRVCRationale = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
        
            If Not RecordToTable("ClaimDetail", "ClaimNumber", strClaimNumber, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("ClaimDetail", "RVCFinalDet", strRVCFinalDet, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("ClaimDetail", "RVCQA", strRVCQA, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("ClaimDetail", "RVCRationale", strRVCRationale, CInt(strReportRow)) Then GoTo ErrorHandler
            
            bolSavedData = True
            
        End If
    
    Next
    
    CurrCellY = iRows 'update the current Y cell
    
    If bolSavedData Then LoadDetailClaimTable = True
    
    Exit Function
    
ErrorHandler:
    
    LoadDetailClaimTable = False


End Function

Function LoadMainClaimTable(CurrCellX As Integer, CurrCellY As Integer) As Boolean

On Error GoTo ErrorHandler
    
    LoadMainClaimTable = False
    
    Dim FoundAtCellX As Integer
    Dim FoundAtCellY As Integer
    
    Dim strReportRow As String
    Dim strClaimNumber As String
    Dim strNewIssueNumber As String
    Dim strProviderType As String
    Dim strReviewType As String
    Dim strMeetAuditParms As String
    Dim strRVCRecAuditParms As String
    Dim strBeneLiability As String
    Dim strRAerrorCode As String
    Dim strRVCerrorCode As String
    Dim strHCPCSCodeDRG As String
    Dim strRAImproperPayType As String
    Dim strRVCImproperPayType As String
    Dim strRAImproperAmt As String
    Dim strClaimDeniedAcc As String
    Dim strImproperPayDet As String
    Dim strRAvsRACinfo As String
    Dim strNumReviews As String
    Dim strRVCAgreeBeneLiability As String
    Dim strRVCAgreeErrorCode As String
    Dim strRVCAgreeImproperAmt As String
    Dim strChosenQA As String
    
    Dim bolFoundData As Boolean
    Dim bolSavedData As Boolean
    Dim iRows As Integer
    
    Dim objCellFont As Excel.Font
    
    Dim intMasterCellFontX As Integer
    Dim intMasterCellFontY As Integer
    
    UpdateProcessStatus "Load Main Claim Table"
    
    bolFoundData = False
    bolSavedData = False
    strReportRow = 0
    
    Call NextCellDown(CurrCellX, CurrCellY, FoundAtCellY)
    
    intMasterCellFontX = CurrCellX
    intMasterCellFontY = FoundAtCellY

    
    'here we just go down the cells looking for the correct words
    For iRows = FoundAtCellY To FoundAtCellY + 300
    
        If CellBelongsToGroup(CurrCellX, iRows, intMasterCellFontX, intMasterCellFontY) Then
            If Not bolFoundData Then
                bolFoundData = True
            End If
        Else
            If Not bolFoundData Then
                GoTo ErrorHandler
            Else
                Exit For
            End If
        End If

        strClaimNumber = Nz(ws.Cells(iRows, CurrCellX).Text, "")
        
        
        'we need to get the row number from the report, this will be used later to match to the second claim table
        strReportRow = NextCellLeft(CurrCellX, iRows)
        
        'the row number must be numeric
        If Not IsNumeric(strReportRow) Then
            GoTo ErrorHandler
        End If
        
        'now we get the rest of columns to the right
        
        If mstrClaimTableVer = "2011-09" Then
            strNewIssueNumber = NextCellRight(CurrCellX, iRows, FoundAtCellX)
            strProviderType = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strMeetAuditParms = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRVCRecAuditParms = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strBeneLiability = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRAerrorCode = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRVCerrorCode = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strHCPCSCodeDRG = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRAImproperPayType = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRVCImproperPayType = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRAImproperAmt = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strClaimDeniedAcc = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strImproperPayDet = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRAvsRACinfo = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strNumReviews = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            If Not RecordToTable("Claim", "ClaimNumber", strClaimNumber, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "NewIssueNumber", strNewIssueNumber, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "ProviderType", strProviderType, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "MeetAuditParms", strMeetAuditParms, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RVCRecAuditParms", strRVCRecAuditParms, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "BeneLiability", strBeneLiability, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RAerrorCode", strRAerrorCode, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RVCerrorCode", strRVCerrorCode, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "HCPCSCodeDRG", strHCPCSCodeDRG, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RAImproperPayType", strRAImproperPayType, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RVCImproperPayType", strRVCImproperPayType, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RAImproperAmt", strRAImproperAmt, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "ClaimDeniedAcc", strClaimDeniedAcc, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "ImproperPayDet", strImproperPayDet, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RAvsRACinfo", strRAvsRACinfo, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "NumReviews", strNumReviews, CInt(strReportRow)) Then GoTo ErrorHandler
            bolSavedData = True
        End If

        If mstrClaimTableVer = "2015-04" Then
            strNewIssueNumber = NextCellRight(CurrCellX, iRows, FoundAtCellX)
            strReviewType = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strMeetAuditParms = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRVCAgreeBeneLiability = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRVCAgreeErrorCode = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRAImproperPayType = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRVCImproperPayType = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRAImproperAmt = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRVCAgreeImproperAmt = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strRAvsRACinfo = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            strChosenQA = NextCellRight(FoundAtCellX, iRows, FoundAtCellX, NotEmpty:=False)
            strClaimDeniedAcc = NextCellRight(FoundAtCellX, iRows, FoundAtCellX)
            If Not RecordToTable("Claim", "ClaimNumber", strClaimNumber, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "NewIssueNumber", strNewIssueNumber, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "ReviewType", strReviewType, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "MeetAuditParms", strMeetAuditParms, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RVCAgreeBeneLiability", strRVCAgreeBeneLiability, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RVCAgreeErrorCode", strRVCAgreeErrorCode, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RAImproperPayType", strRAImproperPayType, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RVCImproperPayType", strRVCImproperPayType, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RAImproperAmt", strRAImproperAmt, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RVCAgreeImproperAmt", strRVCAgreeImproperAmt, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "RAvsRACinfo", strRAvsRACinfo, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "ChosenQA", strChosenQA, CInt(strReportRow)) Then GoTo ErrorHandler
            If Not RecordToTable("Claim", "ClaimDeniedAcc", strClaimDeniedAcc, CInt(strReportRow)) Then GoTo ErrorHandler
            bolSavedData = True
        End If
        
    Next
    
    
    CurrCellY = iRows 'update the current Y cell
    If bolSavedData Then LoadMainClaimTable = True
    
    Exit Function
    
ErrorHandler:
    
    LoadMainClaimTable = False
        
End Function


Function LoadAgreementRateValues(CurrCellX, CurrCellY) As Boolean

On Error GoTo ErrorHandler
    
    LoadAgreementRateValues = False
    
    Dim iRows As Integer
    Dim iColumns As Integer
    Dim iTotalReviewed As Integer
    Dim iAgreement As Integer
    Dim iDisagreement As Integer
    Dim iRate As Integer
    Dim strCurrentCell As String
    Dim strRightFromCurrentCell As String
    
    Dim bolTotalReviewedFound As Boolean
    Dim bolAgreementFound As Boolean
    Dim bolDisagreementFound As Boolean
    Dim bolRateFound As Boolean
    
    UpdateProcessStatus "Load Agreement Rate values"
    
    iTotalReviewed = 0
    iAgreement = 0
    iDisagreement = 0
    iRate = 0
    
    bolTotalReviewedFound = False
    bolAgreementFound = False
    bolDisagreementFound = False
    bolRateFound = False
    
    'here we just go down the cells looking for the correct words
    For iRows = CurrCellY + 1 To CurrCellY + 150
        Debug.Print strCurrentCell
        strCurrentCell = ws.Cells(iRows, CurrCellX).Text
        strRightFromCurrentCell = NextCellRight(CurrCellX, iRows)
        'check for TotalReviewed
        If iTotalReviewed = 0 And InStr(1, Nz(strCurrentCell, ""), "Total", vbTextCompare) > 0 And InStr(1, Nz(strCurrentCell, ""), "Reviewed", vbTextCompare) > 0 And IsNumeric(Nz(strRightFromCurrentCell, "")) Then
            iTotalReviewed = Abs(CInt(strRightFromCurrentCell))
            bolTotalReviewedFound = True
        End If
        'Check for All Agrees
        If iAgreement = 0 And (InStr(1, Nz(strCurrentCell, ""), "Accurately", vbTextCompare) > 0 Or InStr(1, Nz(strCurrentCell, ""), "in Agreement", vbTextCompare) > 0) And IsNumeric(Nz(strRightFromCurrentCell, "")) Then
            iAgreement = Abs(CInt(strRightFromCurrentCell))
            bolAgreementFound = True
        End If
        'Check for All Disagrees
        If iDisagreement = 0 And (InStr(1, Nz(strCurrentCell, ""), "Inaccurate", vbTextCompare) > 0 Or InStr(1, Nz(strCurrentCell, ""), "in Disagreement", vbTextCompare) > 0) And IsNumeric(Nz(strRightFromCurrentCell, "")) Then
            iDisagreement = Abs(CInt(strRightFromCurrentCell))
            bolDisagreementFound = True
        End If
        'Check for Rate
        If iRate = 0 And InStr(1, Nz(strCurrentCell, ""), " rate", vbTextCompare) > 0 And IsNumeric(Replace(Nz(strRightFromCurrentCell, ""), "%", "")) Then
            iRate = Abs(CInt(Replace(strRightFromCurrentCell, "%", "")))
            bolRateFound = True
        End If
        'For iColumns = CurrCellX To CurrCellX + 1
        'Next
        
        If bolTotalReviewedFound And bolAgreementFound And bolDisagreementFound And bolRateFound Then
            GoTo OutOfLoop
        End If

    Next
    
OutOfLoop:
    
    If iTotalReviewed > 0 And ((iAgreement = 0 And iRate = 0) Or (iDisagreement = 0 And iRate = 100) Or (iAgreement > 0 And iRate > 0 And iDisagreement > 0)) Then
        If Not RecordToTable("Cell", "TotalReviewed", iTotalReviewed, 0) Then GoTo ErrorHandler
        If Not RecordToTable("Cell", "Agreement", iAgreement, 0) Then GoTo ErrorHandler
        If Not RecordToTable("Cell", "Disagreement", iDisagreement, 0) Then GoTo ErrorHandler
        If Not RecordToTable("Cell", "Rate", iRate, 0) Then GoTo ErrorHandler
        LoadAgreementRateValues = True
        
    End If
        
    Exit Function
    
ErrorHandler:
    
    LoadAgreementRateValues = False
    
End Function

Function NextCellRight(CurrentCellX, CurrentCellY, Optional FoundCellX, Optional NotEmpty As Boolean = True, Optional NotHidden As Boolean = True, Optional MaxCellsToLook As Integer = 100) As String

    Dim i As Integer
    Dim strCellContents As String

    NextCellRight = ""
    
    i = 0
    
    While i <= MaxCellsToLook
    
        'This is to take care of the combined cells
        If ws.Cells(CurrentCellY, CurrentCellX + i).MergeArea.Columns.Count > 1 Then
            i = i + (ws.Cells(CurrentCellY, CurrentCellX).MergeArea.Columns.Count)
        Else
            i = i + 1
        End If
    
        'this is to take care of the hidden columns
        If (NotHidden And ws.Cells(CurrentCellY, CurrentCellX + i).EntireColumn.Hidden = False) Or NotHidden = False Then
        
            'this is to take care of the empty cells
            If left(Nz(ws.Cells(CurrentCellY, CurrentCellX + i).Text, "  "), 1) = "#" Then
                strCellContents = Nz(ws.Cells(CurrentCellY, CurrentCellX + i).Value, "")
            Else
                strCellContents = Nz(ws.Cells(CurrentCellY, CurrentCellX + i).Text, "")
            End If
            
            
            If (NotEmpty And strCellContents <> "") Or NotEmpty = False Then
                NextCellRight = strCellContents
                FoundCellX = CurrentCellX + i
                GoTo OutOfTheLoop
            End If
        End If

    Wend
    
OutOfTheLoop:

End Function

Function NextCellLeft(CurrentCellX, CurrentCellY, Optional FoundCellX, Optional NotHidden As Boolean = True, Optional NotEmpty As Boolean = True) As String

    Dim i As Integer
    Dim strCellContents As String

    NextCellLeft = ""
    
    'This is to take care of the combined cells

    i = 0
    
    While i <= 100
        
        If CurrentCellX - i - 1 < 1 Then GoTo OutOfTheLoop
        
        If ws.Cells(CurrentCellY, CurrentCellX - i - 1).MergeArea.Columns.Count > 1 Then
            i = (i + ws.Cells(CurrentCellY, CurrentCellX - i - 1).MergeArea.Columns.Count)
        Else
            i = i + 1
        End If
        
        
        
        If (NotHidden And ws.Cells(CurrentCellY, CurrentCellX - i).EntireColumn.Hidden = False) Or NotHidden = False Then
            If left(Nz(ws.Cells(CurrentCellY, CurrentCellX - i).Text, "  "), 1) = "#" Then
                strCellContents = Nz(ws.Cells(CurrentCellY, CurrentCellX - i).Value, "")
            Else
                strCellContents = Nz(ws.Cells(CurrentCellY, CurrentCellX - i).Text, "")
            End If
            
            If (NotEmpty And strCellContents <> "") Or NotEmpty = False Then
                NextCellLeft = strCellContents
                FoundCellX = CurrentCellX - i
                GoTo OutOfTheLoop
            End If
        End If

    Wend
    
OutOfTheLoop:

End Function

Function NextCellDown(CurrentCellX, CurrentCellY, Optional FoundCellY, Optional NotHidden As Boolean = True, Optional NotEmpty As Boolean = True) As String

    Dim i As Integer
    Dim strCellContents As String

    NextCellDown = ""
    
    i = 0
    
    While i <= 100
    
        If ws.Cells(CurrentCellY, CurrentCellX).MergeArea.Rows.Count > 1 Then
            i = (i + ws.Cells(CurrentCellY, CurrentCellX).MergeArea.Rows.Count)
        Else
            i = i + 1
        End If
    
        If (NotHidden And ws.Cells(CurrentCellY + i, CurrentCellX).EntireRow.Hidden = False) Or NotHidden = False Then
            If left(Nz(ws.Cells(CurrentCellY + i, CurrentCellX).Text, "  "), 2) = "##" Then
                strCellContents = Nz(ws.Cells(CurrentCellY + i, CurrentCellX).Value, "")
            Else
                strCellContents = Nz(ws.Cells(CurrentCellY + i, CurrentCellX).Text, "")
            End If
            If (NotEmpty And strCellContents <> "") Or NotEmpty = False Then
                NextCellDown = strCellContents
                FoundCellY = CurrentCellY + i
                GoTo OutOfTheLoop
            End If
        End If

    Wend
    
OutOfTheLoop:

End Function


Function RecordToTable(KeyType As String, KeyName As String, ByVal KeyValue As String, KeyReportRow As Integer) As Boolean

On Error GoTo ErrorHandler

    RecordToTable = False


    Dim MyCodeAdo As clsADO
    Dim cmd As ADODB.Command
    Dim sqlString As String
    Dim spReturnVal As Integer
    Dim ErrorReturned As String
    
    KeyReportRow = Nz(KeyReportRow, 0)
   
    UpdateProcessStatus "Record To Table: KeyReportRow = '" & CStr(KeyReportRow) & "' KeyType = '" & KeyType & "' KeyName = '" & KeyName & "' " & " KeyValue = '" & KeyValue & "'"
    
    Set MyCodeAdo = New clsADO
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 600
    With cmd
        .ActiveConnection = MyCodeAdo.CurrentConnection
        .CommandText = "usp_ACCURACY_LoadToTmpTable"
        .commandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters("@pUserID") = mstrUserID
        .Parameters("@pSessionID") = mstrSessionID
'        .Parameters("@pAuditType") = mstrAuditType
'        .Parameters("@pAuditScope") = mstrAuditScope
'        .Parameters("@pResultVersion") = mstrResultVersion
'        .Parameters("@pSourceExcelFileName") = mstrSourceExcelFileName
        .Parameters("@pKeyReportRow") = KeyReportRow
        .Parameters("@pKeyType") = KeyType
        .Parameters("@pKeyName") = KeyName
        .Parameters("@pKeyValue") = KeyValue
        .Execute
        spReturnVal = .Parameters("@Return_Value")
        ErrorReturned = Nz(.Parameters("@pErrMsg"), "")
        If spReturnVal <> 0 Or ErrorReturned <> "" Then
            GoTo ErrorHandler
        End If
    End With
          
    RecordToTable = True
          
ExitFunction:

    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    
    Exit Function
    
ErrorHandler:
    
    If Nz(ErrorReturned, "") = "" Then ErrorReturned = Nz(Err.Description, "")
    UpdateProcessStatus "Error: usp_ACCURACY_LoadToTmpTable returned ReturnVal = '" & Nz(spReturnVal, 0) & "' ErrMsg = '" & ErrorReturned & "'"
    
    RecordToTable = False
    
    GoTo ExitFunction
    
End Function

Function FullScanReconnaissance() As Boolean

On Error GoTo ErrorHandler

Dim ABCHeaderFound As Boolean
Dim ClaimHeaderFound As Boolean
Dim ClaimDetailHeaderFound As Boolean
Dim NewIssueTrackingHeaderFound As Boolean
Dim RecommendationsHeaderFound As Boolean
Dim ClaimDetailTitleFound As Boolean

Dim RowCounter As Integer
Dim CurrCellX As Integer
Dim i As Integer
Dim ThisRow As String
Dim strCellValue As String
Dim intLastLineHeaderDetected As Integer
Dim objCellFont As Excel.Font

Dim intMasterCellFontX As Integer
Dim intMasterCellFontY As Integer

Dim bolHeaderNotDetected As Boolean

UpdateProcessStatus "Identify sections and remove duplicate column headers rows"

FullScanReconnaissance = False

ABCHeaderFound = False
ClaimHeaderFound = False
ClaimDetailHeaderFound = False
NewIssueTrackingHeaderFound = False
RecommendationsHeaderFound = False
ClaimDetailTitleFound = False

RowCounter = 1
CurrCellX = 1

intMasterCellFontY = 1

bolHeaderNotDetected = False

Set objCellFont = ws.Cells(RowCounter, 5).Font

While RowCounter <= 500
    
        UpdateProcessStatus "Identify sections and remove duplicate column headers rows, Row = " & CStr(RowCounter)
        ThisRow = ""
        CurrCellX = 1
        ThisRow = Nz(ws.Cells(RowCounter, CurrCellX).Text, "")
        
        If Not CellBelongsToGroup(4, RowCounter, 4, intMasterCellFontY) Then
            intLastLineHeaderDetected = RowCounter
            bolHeaderNotDetected = True
        End If
        
        If bolHeaderNotDetected And (RowCounter - intLastLineHeaderDetected > 5) Then
            UpdateProcessStatus "5 rows after the last section started a new header was not recognized. This might be an new unrecognized version of the report. Current Row = " & str(RowCounter)
            GoTo ErrorHandler
        End If
        
        For i = 1 To 20
            strCellValue = NextCellRight(CurrCellX, RowCounter, CurrCellX, NotEmpty:=False)
            ThisRow = ThisRow & strCellValue
        Next
                
                
        Debug.Print "RowNumber = " & str(RowCounter) & " - Text = " & ThisRow
        
        
        
        ThisRow = Replace(Replace(Replace(Replace(Replace(Replace(ThisRow, " ", ""), Chr(34), ""), Chr(13), ""), Chr(10), ""), Chr(9), ""), "'", "")
        
        If ThisRow = "" Then
            bolHeaderNotDetected = False
            GoTo ToTheNextRow
        End If

'        If IsNumeric(ThisRow) Then
'            If val(ThisRow) < 500 Then
'                bolHeaderNotDetected = False
'                GoTo DeleteThisRow
'            End If
'        End If


        'First case, for the letter column headers
        
        If left(ThisRow, 3) = "ABC" Then
            bolHeaderNotDetected = False
            If Not ABCHeaderFound Then
                ABCHeaderFound = True
                GoTo ToTheNextRow
            Else
                GoTo DeleteThisRow
            End If
        End If

        'second case, Claim Headers
        If InStr(1, ThisRow, "ClaimNumberNewIssueNumberReviewType(Automated,Semi-AutomatedorComplex)Didtheclaimmeettheauditparameters?(YesorNo)DoestheRVCagreewiththeRAsbeneficiaryliabilitydetermination?DoestheRVCagreewiththeRAserrorcodedetermination?RAImproperPaymentType(Full,PartialorUnder)RVCImproperPaymentType(Full,PartialorUnder)RAImproperPaymentAmountDoestheRVCagreewiththeRAsImproperPaymentAmount?DoesRAsubmittedinformationmatchtheRDW?WastheclaimchosenforQA?Wasthisclaimdeniedaccurately?", vbTextCompare) > 0 _
            Or InStr(1, ThisRow, "ClaimNumberNewIssueNumberReviewType(Automated,Semi-AutomatedorComplex)Didtheclaimmeettheauditparameters?(YesorNo)DoestheRVCagreewiththeRAsbeneficiaryliabilitydetermination?DoestheRVCagreewiththeRAserrorcodedetermination?RAImproperPaymentType(Full,PartialorUnder)RVCImproperPaymentType(Full,PartialorUnder)RAImproperPaymentAmountDoestheRVCagreewiththeRAsImproperPaymentAmount?DoesRAsubmittedinformationmatchtheRACDW?WastheclaimchosenforQA?Wasthisclaimdeniedaccurately?", vbTextCompare) > 0 _
            Or InStr(1, ThisRow, "ClaimNumberNewIssueNumberReviewType(Automated,Semi-AutomatedorComplex)Didtheclaimmeettheauditparameters?(YesorNo)DoestheRVCagreewiththeRecoveryAuditorsbeneficiaryliabilitydetermination?DoestheRVCagreewiththeRecoveryAuditorserrorcodedetermination?RecoveryAuditorImproperPaymentType(Full,Partial,orUnder)RVCImproperPaymentType(Full,Partial,orUnder)RecoveryAuditorImproperPaymentAmountDoestheRVCagreewiththeRecoveryAuditorsImproperPaymentAmount?DoesRecoveryAuditorsubmittedinformationmatchtheRDW?WastheclaimchosenforQA?Wasthisclaimdeniedaccurately?", vbTextCompare) > 0 Then
            bolHeaderNotDetected = False
            mstrClaimTableVer = "2015-04"
            If Not ClaimHeaderFound Then
                ClaimHeaderFound = True
                GoTo ToTheNextRow
            Else
                GoTo DeleteThisRow
            End If
        End If
        
        If InStr(1, ThisRow, "ClaimNumberNewIssueNumberProviderTypeDidtheclaimmeettheauditparameters?(YES,NO,orN/A)DidtheRVCreceiveauditparameters?Istherebeneficiaryliability?(YESorNO)RAerrorcodeRVCerrorcodeHCPCSCode(DRG)RAImproperPaymentType(FULL,PARTIALUNDER)RVCImproperPaymentType(FULL,PARTIAL,UNDER)RAImproperPaymentAmountWasclaimdeniedaccurately?ImproperPaymentdeterminationCorrect?(YESorNO)RAInfo=RACDWInfo#Reviews", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "ClaimNumberNewIssueNumberProviderTypeDidtheclaimmeettheauditparameters?(YES,NO,orNA)DidtheRVCreceiveauditparameters?Istherebeneficiaryliability?(YESorNO)RAerrorcodeRVCerrorcodeHCPCSCode(DRG)RAImproperPaymentType(FULL,PARTIALUNDER)RVCImproperPaymentType(FULL,PARTIAL,UNDER)RAImproperPaymentAmountWasclaimdeniedaccurately?ImproperPaymentdeterminationCorrect?(YESorNO)RAInfo=RACDWInfo#Reviews", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "ClaimNumberNewIssueNumberProviderTypeDidtheclaimmeettheauditparameters?(YES,NO,orNA)DidtheRVCreceiveauditparameters?Istherebeneficiaryliability?(YESorNO)RAerrorcodeRVCerrorcodeHCPCSCode(DRG)RAImproperPaymentType(FULL,PARTIAL,UNDER)RVCImproperPaymentType(FULL,PARTIAL,UNDER)RAImproperPaymentAmountWasclaimdeniedaccurately?ImproperPaymentdeterminationCorrect?(YESorNO)RAInfo=RACDWInfo#Reviews", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "ClaimNumberNewIssueNumberProviderTypeDidtheclaimmeettheauditparameters?(YES,NO,orN/A)DidtheRVCreceiveauditparameters?Istherebeneficiaryliability?(YESorNO)RACerrorcodeRVCerrorcodeHCPCSCode(DRG)RACImproperPaymentType(FULL,PARTIALUNDER)RVCImproperPaymentType(FULL,PARTIAL,UNDER)RACImproperPaymentAmountWasclaimdeniedaccurately?ImproperPaymentdeterminationCorrect?(YESorNO)RACInfo=RACDWInfo#Reviews", vbTextCompare) > 0 Then
            bolHeaderNotDetected = False
            mstrClaimTableVer = "2011-09"
            If Not ClaimHeaderFound Then
                ClaimHeaderFound = True
                GoTo ToTheNextRow
            Else
                GoTo DeleteThisRow
            End If
        End If




        'third case, Claim Detail titles and headers
        
        
        If InStr(1, ThisRow, "ClaimDetail", vbTextCompare) > 0 Then
            bolHeaderNotDetected = False
            If Not ClaimDetailTitleFound Then
                ClaimDetailTitleFound = True
                GoTo ToTheNextRow
            Else
                GoTo DeleteThisRow
            End If
        End If
        
        
        If InStr(1, ThisRow, "Claim#RVCFinalDetermination(Agree/Disagree)QARationaleforRVCDetermination", vbTextCompare) > 0 _
            Or InStr(1, ThisRow, "Claim#RVCFinalDeterminationAgreeDisagreeQARationaleforRVCDetermination", vbTextCompare) > 0 _
            Or InStr(1, ThisRow, "Claim#RVCFinalDetermination(AgreeDisagree)QARationaleforRVCDetermination", vbTextCompare) > 0 _
            Or InStr(1, ThisRow, "Claim#RVCFinalDeterminationAgree/DisagreeQARationaleforRVCDetermination", vbTextCompare) > 0 Then
            bolHeaderNotDetected = False
            mstrClaimDetailTableVer = "2015-04"
            mstrClaimDetailColumnHeader = "Claim #"
            If Not ClaimDetailHeaderFound Then
                ClaimDetailHeaderFound = True
                GoTo ToTheNextRow
            Else
                GoTo DeleteThisRow
            End If
        End If
        
        If InStr(1, ThisRow, "ClaimNumberRVCFinalDetermination(Agree/Disagree)QARationaleforRVCDetermination", vbTextCompare) > 0 Then
            bolHeaderNotDetected = False
            mstrClaimDetailTableVer = "2014-09"
            mstrClaimDetailColumnHeader = "Claim Number"
            If Not ClaimDetailHeaderFound Then
                ClaimDetailHeaderFound = True
                GoTo ToTheNextRow
            Else
                GoTo DeleteThisRow
            End If
        End If
            
        If InStr(1, ThisRow, "ClaimNumber2ndLevel3rdLevelQARationaleforRVCDetermination", vbTextCompare) > 0 Then
            bolHeaderNotDetected = False
            mstrClaimDetailTableVer = "2011-09"
            mstrClaimDetailColumnHeader = "Claim Number"
            If Not ClaimDetailHeaderFound Then
                ClaimDetailHeaderFound = True
                GoTo ToTheNextRow
            Else
                GoTo DeleteThisRow
            End If
        End If

        'fourth case, new issue tracking headers

        If InStr(1, ThisRow, "NINumber#claimsEditParameters", vbTextCompare) > 0 Then
            bolHeaderNotDetected = False
            mstrNewIssueTableVer = "2015-04"
            mstrNewIssueColumnHeader = "NI Number"
            If Not NewIssueTrackingHeaderFound Then
                NewIssueTrackingHeaderFound = True
                GoTo ToTheNextRow
            Else
                GoTo DeleteThisRow
            End If
        End If
        
        'fifth case, recommendations headers
        If InStr(1, ThisRow, "NewIssueNewIssueNameRecommendations", vbTextCompare) > 0 Then
            bolHeaderNotDetected = False
            mstrNewIssueDetailTableVer = "2015-04"
            mstrNewIssueDetailColumnHeader = "New Issue"
            If Not RecommendationsHeaderFound Then
                RecommendationsHeaderFound = True
                GoTo ToTheNextRow
            Else
                GoTo DeleteThisRow
            End If
        End If

        'these are the tables we just dont care to record
        If InStr(1, ThisRow, "ReceivedReviewedSource", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "ACCURACYCLAIMDETAIL", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "ProviderTypes", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "CodeTypeName", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "AccuracyReviewScorecard", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "ClaimAccuracyRate(CumulativeforRA'smonthlyreport)", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "ProviderTypeCode", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "NewIssue#CMSassignedNI#", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "TotalclaimsreviewedClaimslevelonly", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "AccurateIdentificationRate", vbTextCompare) > 0 Or _
            InStr(1, ThisRow, "Lab/Ambo1", vbTextCompare) > 0 Then
            bolHeaderNotDetected = False
            GoTo ToTheNextRow
        End If

'
'                'End If
'            'End If
'        End If


    GoTo ToTheNextRow

DeleteThisRow:
    ws.Cells(RowCounter, 1).EntireRow.Delete
    GoTo DoTheLoop

ToTheNextRow:

    intMasterCellFontY = RowCounter

    RowCounter = RowCounter + 1
    
    
    
DoTheLoop:

   
Wend

If Not (ClaimHeaderFound Or ClaimDetailHeaderFound) Then
    UpdateProcessStatus "Claim Header or Claim Detail Header not found."
    GoTo ErrorHandler
End If

   
FullScanReconnaissance = True

Exit Function

ErrorHandler:

FullScanReconnaissance = False

End Function


Function ScanForText(TextToSearch As String, StartCellX, StartCellY, EndCellX, EndCellY, ByRef FoundCellX, ByRef FoundCellY) As Boolean

On Error GoTo ErrorHandler

Dim RowCounter As Integer
Dim ColumnCounter As Integer

UpdateProcessStatus "Searching for text: '" & TextToSearch & "'"

For RowCounter = StartCellY To EndCellY
    For ColumnCounter = StartCellX To EndCellX
        If InStr(1, ws.Cells(RowCounter, ColumnCounter).Text, TextToSearch) Then
            GoTo OutOfTheLoop
        End If
    Next
Next
    
GoTo ErrorHandler
    
OutOfTheLoop:
    ScanForText = True
    FoundCellX = ColumnCounter
    FoundCellY = RowCounter
    Exit Function
    
ErrorHandler:

    ScanForText = False

End Function


Function ClearTmpTable() As Boolean

On Error GoTo ErrorHandler

    Dim MyCodeAdo As clsADO
    Dim cmd As ADODB.Command
    Dim sqlString As String
    Dim spReturnVal As Integer
    Dim ErrorReturned As String
    
    ClearTmpTable = False
    
    UpdateProcessStatus "Clearing tmp table"
    
    Set MyCodeAdo = New clsADO
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 600
    With cmd
        .ActiveConnection = MyCodeAdo.CurrentConnection
        .CommandText = "usp_ACCURACY_ClearTmpTable"
        .commandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters("@pUserID") = mstrUserID
        .Execute
        spReturnVal = .Parameters("@Return_Value")
        ErrorReturned = Nz(.Parameters("@pErrMsg"), "")
        If spReturnVal <> 0 Or ErrorReturned <> "" Then
            GoTo ErrorHandler
        End If
    End With
    
    ClearTmpTable = True
    
    Exit Function
    
ErrorHandler:

    If Nz(ErrorReturned, "") = "" Then ErrorReturned = Nz(Err.Description, "")
    UpdateProcessStatus "Error: usp_ACCURACY_ClearTmpTable returned ReturnVal = '" & Nz(spReturnVal, 0) & "' ErrMsg = '" & ErrorReturned & "'"


    ClearTmpTable = False

End Function


Sub UpdateProcessStatus(StatusText As String)
    Me.txtProcessStatus = StatusText
    DoEvents
End Sub


Function LoadedDataValidation() As Boolean

On Error GoTo ErrorHandler

    Dim MyCodeAdo As clsADO
    Dim cmd As ADODB.Command
    Dim sqlString As String
    Dim spReturnVal As Integer
    Dim ErrorReturned As String
    
    LoadedDataValidation = False
    
    UpdateProcessStatus "Validating Loaded Data..."
    
    Set MyCodeAdo = New clsADO
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 600
    With cmd
        .ActiveConnection = MyCodeAdo.CurrentConnection
        .CommandText = "usp_ACCURACY_ValidateTmpData"
        .commandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters("@pSessionID") = mstrSessionID
        .Parameters("@pClaimTableVer") = mstrClaimTableVer
        .Parameters("@pClaimDetailTableVer") = mstrClaimDetailTableVer
        .Parameters("@pNewIssueTableVer") = mstrNewIssueTableVer
        .Parameters("@pNewIssueDetailTableVer") = mstrNewIssueDetailTableVer
        .Parameters("@pUserID") = mstrUserID
        '.Parameters("@pAuditType") = mstrAuditType
        .Execute
        spReturnVal = .Parameters("@Return_Value")
        ErrorReturned = Nz(.Parameters("@pErrMsg"), "")
        If spReturnVal <> 0 Or ErrorReturned <> "" Then
            GoTo ErrorHandler
        End If
    End With
    
    LoadedDataValidation = True
    
    Exit Function
    
ErrorHandler:

    If Nz(ErrorReturned, "") = "" Then ErrorReturned = Nz(Err.Description, "")
    UpdateProcessStatus "Error: usp_ACCURACY_ValidateTmpData returned ReturnVal = '" & Nz(spReturnVal, 0) & "' ErrMsg = '" & ErrorReturned & "'"
    LoadedDataValidation = False


End Function

'
'Function DetermineClaimTableVersion(CurrCellX, CurrCellY) As Boolean
'
'    Dim i As Integer
'    Dim strCellValue As String
'    Dim ThisRow As String
'    Dim FoundAtCellX As Integer
'
'    FoundAtCellX = CurrCellX
'
'    mstrClaimTableVer = ""
'
'    ThisRow = Nz(ws.Cells(CurrCellY, CurrCellX).text, "")
'    For i = 1 To 20
'        strCellValue = NextCellRight(FoundAtCellX, CurrCellY, FoundAtCellX)
'        ThisRow = ThisRow & strCellValue
'    Next
'
'    ThisRow = Replace(Replace(Replace(Replace(Replace(ThisRow, " ", ""), Chr(34), ""), Chr(13), ""), Chr(10), ""), Chr(9), "")
'
'    If InStr(1, ThisRow, "ClaimNumberNewIssueNumberReviewType(Automated,Semi-AutomatedorComplex)Didtheclaimmeettheauditparameters?(YesorNo)DoestheRVCagreewiththeRAsbeneficiaryliabilitydetermination?DoestheRVCagreewiththeRAserrorcodedetermination?RAImproperPaymentType(Full,PartialorUnder)RVCImproperPaymentType(Full,PartialorUnder)RAImproperPaymentAmountDoestheRVCagreewiththeRAsImproperPaymentAmount?DoesRAsubmittedinformationmatchtheRDW?WastheclaimchosenforQA?Wasthisclaimdeniedaccurately?", vbTextCompare) > 0 Then
'        mstrClaimTableVer = "2015-04"
'    End If
'
'    If InStr(1, ThisRow, "ClaimNumberNewIssueNumberProviderTypeDidtheclaimmeettheauditparameters?(YES,NO,orN/A)DidtheRVCreceiveauditparameters?Istherebeneficiaryliability?(YESorNO)RAerrorcodeRVCerrorcodeHCPCSCode(DRG)RAImproperPaymentType(FULL,PARTIALUNDER)RVCImproperPaymentType(FULL,PARTIAL,UNDER)RAImproperPaymentAmountWasclaimdeniedaccurately?ImproperPaymentdeterminationCorrect?(YESorNO)RAInfo=RACDWInfo#Reviews", vbTextCompare) > 0 Then
'        mstrClaimTableVer = "2011-09"
'    End If
'
'    If mstrClaimTableVer <> "" Then
'        DetermineClaimTableVersion = True
'    Else
'        DetermineClaimTableVersion = False
'    End If
'
'End Function
'
'
'Function DetermineClaimDetailTableVersion(CurrCellX, CurrCellY) As Boolean
'
'    Dim i As Integer
'    Dim strCellValue As String
'    Dim ThisRow As String
'    Dim FoundAtCellX As Integer
'
'    FoundAtCellX = CurrCellX
'
'    mstrClaimTableVer = ""
'
'    ThisRow = Nz(ws.Cells(CurrCellY, CurrCellX).text, "")
'    For i = 1 To 20
'        strCellValue = NextCellRight(FoundAtCellX, CurrCellY, FoundAtCellX)
'        ThisRow = ThisRow & strCellValue
'    Next
'
'    ThisRow = Replace(Replace(Replace(Replace(Replace(ThisRow, " ", ""), Chr(34), ""), Chr(13), ""), Chr(10), ""), Chr(9), "")
'
'    If InStr(1, ThisRow, "Claim#RVCFinalDetermination(Agree/Disagree)QARationaleforRVCDetermination", vbTextCompare) > 0 Then
'        mstrClaimDetailTableVer = "2015-04"
'    End If
'
'    If InStr(1, ThisRow, "ClaimNumber2ndLevel3rdLevelQARationaleforRVCDetermination", vbTextCompare) > 0 Then
'        mstrClaimDetailTableVer = "2011-09"
'    End If
'
'    If mstrClaimDetailTableVer <> "" Then
'        DetermineClaimDetailTableVersion = True
'    Else
'        DetermineClaimDetailTableVersion = False
'    End If
'
'End Function
'
'
'Function DetermineNewIssueTableVersion(CurrCellX, CurrCellY) As Boolean
'
'    Dim i As Integer
'    Dim strCellValue As String
'    Dim ThisRow As String
'    Dim FoundAtCellX As Integer
'
'    FoundAtCellX = CurrCellX
'
'    mstrClaimTableVer = ""
'
'    ThisRow = Nz(ws.Cells(CurrCellY, CurrCellX).text, "")
'    For i = 1 To 20
'        strCellValue = NextCellRight(FoundAtCellX, CurrCellY, FoundAtCellX)
'        ThisRow = ThisRow & strCellValue
'    Next
'
'    ThisRow = Replace(Replace(Replace(Replace(Replace(ThisRow, " ", ""), Chr(34), ""), Chr(13), ""), Chr(10), ""), Chr(9), "")
'
'    If InStr(1, ThisRow, "NINumber#claimsEditParameters", vbTextCompare) > 0 Then
'        mstrNewIssueTableVer = "2015-04"
'    End If
'
'
'    If mstrNewIssueTableVer <> "" Then
'        DetermineNewIssueTableVersion = True
'    Else
'        DetermineNewIssueTableVersion = False
'    End If
'
'End Function


Function SaveResultsToTables() As Boolean
'    Dim mycode_ADO As clsADO
'

    
    Dim MyCodeAdo As clsADO
    Dim cmd As ADODB.Command
    Dim sqlString As String
    Dim spReturnVal As Integer
    Dim ErrorReturned As String
    
    UpdateProcessStatus "Saving Results to Tables..."
    
    Set MyCodeAdo = New clsADO
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")

    
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 600
    With cmd
        .ActiveConnection = MyCodeAdo.CurrentConnection
        .CommandText = "usp_ACCURACY_SaveScore"
        .commandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters("@pSessionID") = mstrSessionID
        .Parameters("@pUserID") = mstrUserID
        .Execute
        spReturnVal = .Parameters("@Return_Value")
        ErrorReturned = Nz(.Parameters("@pErrMsg"), "")
        If spReturnVal <> 0 Or ErrorReturned <> "" Then
            GoTo ErrorHandler
        End If
    End With
    
    
'    mycode_ADO.CommitTrans
    
    SaveResultsToTables = True
    
    Exit Function
    
ErrorHandler:

    SaveResultsToTables = False
End Function

Private Sub txtSourceFileFolder_Change()
    ClearFormFields clearFileName:=True
End Sub

Private Sub txtSourceFileName_Change()

    ClearFormFields
    Me.lstAuditType = "Accuracy"
    
    Dim iYears As Integer
    Dim iMonths As Integer
    Dim strPeriod As String

    
    iYears = 2009
    iMonths = 1
    
    While iYears < "2020"
        iMonths = 1
        While iMonths < 13
            strPeriod = Trim$(str(iYears)) & "-" & Right("00" & Trim$(str(iMonths)), 2)
            Debug.Print strPeriod
            If InStr(1, Me.txtSourceFileName, strPeriod) > 0 Then
                GoTo YearFound
            End If
            iMonths = iMonths + 1
        Wend
        iYears = iYears + 1
    Wend
    
    GoTo YearNotFound
    
YearFound:
    Me.txtAuditScope = strPeriod

YearNotFound:

    If InStr(1, Me.txtSourceFileName, "V01") > 0 Or InStr(1, Me.txtSourceFileName, "V1") > 0 Then
        Me.lstResultVersion = "Score1"
    ElseIf InStr(1, Me.txtSourceFileName, "V02") > 0 Or InStr(1, Me.txtSourceFileName, "V2") > 0 Then
        Me.lstResultVersion = "Score2"
    ElseIf InStr(1, Me.txtSourceFileName, "V03") > 0 Or InStr(1, Me.txtSourceFileName, "V3") > 0 Then
        Me.lstResultVersion = "Score3"
    ElseIf InStr(1, Me.txtSourceFileName, "V04") > 0 Or InStr(1, Me.txtSourceFileName, "V4") > 0 Then
        Me.lstResultVersion = "Score4"
    End If

    
End Sub

Sub ClearFormFields(Optional clearFileName As Boolean = False)
    If clearFileName Then
        Me.txtSourceFileName = ""
    End If
    Me.lstAuditType = ""
    Me.lstResultVersion = ""
    Me.txtAuditScope = ""
    If left(Me.txtProcessStatus, 4) <> "Done" Then Me.txtProcessStatus = ""
End Sub

Function ReadyToStart() As Boolean
    ReadyToStart = False
    If Me.lstAuditType = "" Or Me.lstResultVersion = "" Or Me.txtAuditScope = "" Then
        MsgBox "You need to enter all the parameters before processing.", vbExclamation, "Error"
        GoTo ErrorHandler
    End If
    
    ReadyToStart = True
    Exit Function
    
ErrorHandler:

End Function


Function PreLoadValidate() As Boolean
On Error GoTo ErrorHandler

    Dim MyCodeAdo As clsADO
    Dim cmd As ADODB.Command
    Dim sqlString As String
    Dim spReturnVal As Integer
    Dim ErrorReturned As String
    
    PreLoadValidate = False
    
    UpdateProcessStatus "Pre Load Validation..."
    
    Set MyCodeAdo = New clsADO
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 600
    With cmd
        .ActiveConnection = MyCodeAdo.CurrentConnection
        .CommandText = "usp_ACCURACY_PreLoadValidate"
        .commandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters("@pSessionID") = mstrSessionID
        .Parameters("@pUserID") = mstrUserID
        .Parameters("@pAuditType") = mstrAuditType
        .Parameters("@pAuditScope") = mstrAuditScope
        .Parameters("@pResultVersion") = mstrResultVersion
        .Parameters("@pSourceFileName") = mstrSourceExcelFileName
        .Execute
        spReturnVal = .Parameters("@Return_Value")
        ErrorReturned = Nz(.Parameters("@pErrMsg"), "")
        If spReturnVal <> 0 Or ErrorReturned <> "" Then
            GoTo ErrorHandler
        End If
    End With
    
    PreLoadValidate = True
    
    Exit Function
    
ErrorHandler:

    If Nz(ErrorReturned, "") = "" Then ErrorReturned = Nz(Err.Description, "")
    UpdateProcessStatus "Error: usp_ACCURACY_PreLoadValidate returned ReturnVal = '" & Nz(spReturnVal, 0) & "' ErrMsg = '" & ErrorReturned & "'"
    PreLoadValidate = False

End Function

Private Sub txtSourceFileName_Click()
    Me.txtSourceFileName.SelStart = 0
    Me.txtSourceFileName.SelLength = Len(Me.txtSourceFileName.Text)
End Sub

Private Sub txtSourceFileName_Exit(Cancel As Integer)
    txtSourceFileName_Change
End Sub

Private Sub txtSourceFileName_GotFocus()
    txtSourceFileName_Click
End Sub

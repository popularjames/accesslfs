Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#If ccSCR = 1 Then
Private MvFrmScr As Form_SCR_MainScreens
#End If

' DLC 11/01/2011 - Save in Excel 97-2003 format if the filename ends with .xls
Private MvRunning As Boolean
Private mvFormId As Byte
Private MvShowUI As Boolean
Private MvFileName As String
Private MvQuiet As Boolean
Private MvQryName As String
Private MvIncludeFormats As Boolean
Private MvAutoStart As Boolean
Private MvOverWrite As Boolean
Private MvExpObject As CnlyExportObj

Private WithEvents MvFrmExcel As Form_CT_Excel
Attribute MvFrmExcel.VB_VarHelpID = -1

Private Type CnlySelColumn
    Caption As String
    CtrlName As String
    Calculated As Boolean
    ControlSource As String
    order As Long
    Value As String
    AllignAc As Long
    AllignXl As Long
    Format As String
    FormatXl As String
    Decimals As Integer
End Type

Private Type CnlySelRange
    top As Long
    left As Long
    Right As Long
    Height As Long
    Columns As Integer
End Type
Public Event Status(Task As String, Msg As String, lvl As Integer)


Public Function Access2ExcelFormat(AcFormat As String, AcDecimals As Integer)
Dim StDecimals As String, StFormat As String

If AcDecimals = 0 Or AcDecimals = 255 Then
    StDecimals = ""
Else
    StDecimals = "." & String(AcDecimals, "0")
End If

Select Case UCase(AcFormat)
Case "CURRENCY"
    If AcDecimals = 255 Then 'OverRide
        StDecimals = "." & String(2, "0")
    End If
    StFormat = "$#,##0" & StDecimals
Case "STANDARD"
    StFormat = "#,##0" & StDecimals
Case "PERCENT"
    StFormat = "0" & StDecimals & "%"
Case Else
    'TAKE THE FORMAT AS IS (AND PRAY)
    StFormat = AcFormat & StDecimals
End Select

Access2ExcelFormat = StFormat
End Function

Public Function Access2ExcelAllign(AcAllign As Long) As Long
    Select Case AcAllign
    Case 0 'General
        Access2ExcelAllign = 1 'xlHAlignGeneral
    Case 1 ' Left
        Access2ExcelAllign = -4131 'xlLeft
    Case 2 ' Center
        Access2ExcelAllign = -4108 'xlCenter
    Case 3 'Right
        Access2ExcelAllign = -4152 'xlright
    Case 4 'Distribute
        Access2ExcelAllign = -4117 'xlHAlignDistributed
    Case Else
        Access2ExcelAllign = 0 'Return general is something else passed in
    End Select
End Function

Public Property Let ExpObject(data As CnlyExportObj)
    MvExpObject = data
End Property
Public Property Get ExpObject() As CnlyExportObj
    ExpObject = MvExpObject
End Property

'Changed to late binding so it works with any size datasheet
Private Function GetColumnsOrdered(Grid As Object) As CnlySelColumn()
On Error GoTo ErrorHappened
Dim tmpCol As New Collection, SelCol As CnlySelColumn, ArySort() As String
Dim AryCols() As CnlySelColumn
Dim TxtFld As Access.TextBox

Dim X As Long

ReDim ArySort(Grid.FldCT + 1) As String
'SET THE COLUMN HEADERS

'Retrieve an array of columns to be included
For X = Grid.FldCT To 1 Step -1
    Set TxtFld = Grid.Controls("Field" & X)
    If TxtFld.ColumnWidth <> 0 And TxtFld.ColumnHidden = False Then
        ArySort(TxtFld.ColumnOrder) = TxtFld.Name
    End If
    Set TxtFld = Nothing
Next X

'Put them in order in to a collection  (COULD NOT STORE THE STRUCTURE IN COLLETION)
For X = 0 To UBound(ArySort)
    If "" & ArySort(X) <> "" Then
        'Debug.Print ArySort(X) & " - " & Grid.Controls(ArySort(X)).ControlSource
        If "" & Grid.Controls(ArySort(X)).ControlSource <> "" Then  'REMOVE BLANK COLUMNS
            tmpCol.Add ArySort(X), CStr(X)
        End If
    End If
Next X

'Now Make an Array of Definitions
ReDim AryCols(tmpCol.Count - 1) As CnlySelColumn
For X = 1 To tmpCol.Count
    Set TxtFld = Grid.Controls(tmpCol(X))
    With TxtFld
        SelCol.CtrlName = .Name
        If "" & .Tag <> "" Then
            SelCol.Calculated = True
            If left(.ControlSource, 1) = "=" Then
                SelCol.ControlSource = Mid(.ControlSource, 2)
            Else
                SelCol.ControlSource = .ControlSource
            End If
        Else
            SelCol.Calculated = False
            SelCol.ControlSource = .ControlSource
        End If
        SelCol.AllignAc = .TextAlign
        SelCol.Decimals = .DecimalPlaces
        SelCol.order = .ColumnOrder
        SelCol.Format = .Format
        SelCol.Caption = Grid.Controls("CapField" & Mid(SelCol.CtrlName, 6)).Caption
        SelCol.AllignXl = Access2ExcelAllign(SelCol.AllignAc)
        SelCol.FormatXl = Access2ExcelFormat(SelCol.Format, SelCol.Decimals)
    End With
    Set TxtFld = Nothing
    AryCols(X - 1) = SelCol
Next X

GetColumnsOrdered = AryCols

ExitNow:
    On Error Resume Next
    Set tmpCol = Nothing 'Free my memory!
    Set TxtFld = Nothing
    Exit Function
ErrorHappened:
    RaiseEvent Status("GetColumnsOrdered", Err.Description, 10)
    If MvQuiet = False Then
        MsgBox Err.Description, vbCritical, "ClsExcel.GetColumnsOrdered"
    End If
    Resume ExitNow
    
End Function

Public Property Let Overwrite(data As Boolean)
    MvOverWrite = data
End Property
Public Property Get Overwrite() As Boolean
    Overwrite = MvOverWrite
End Property



Public Property Let AutoStart(data As Boolean)
    MvAutoStart = data
End Property
Public Property Get AutoStart() As Boolean
    AutoStart = MvAutoStart
End Property
Public Function ExportSelected(Grid As Form_CT_SubGenericDataSheet) As Boolean
On Error GoTo ErrorHappened
'Dim tmpCol As New Collection, SelRge As CnlySelRange, SelCol As CnlySelColumn, ArySort() As String
Dim SelRge As CnlySelRange
Dim AryTmp() As CnlySelColumn
Dim AryCols() As CnlySelColumn

Dim rst As DAO.RecordSet, TxtFld As Access.TextBox
Dim xlApp 'As Excel.Application
Dim XlWbk 'As Excel.Workbook
Dim XlSht 'As Excel.Worksheet
Dim Row As Long, X As Long
Dim Status As Form_CT_Status
'Dim Tmp As Variant
With Grid
    SelRge.left = .SelLeft - 1
    SelRge.Height = .SelHeight
    SelRge.top = .SelTop
    SelRge.Right = .SelWidth + (SelRge.left - 1)
End With

If SelRge.Height = 0 Then 'Nothing to do here
    GoTo ExitNow
End If

DoCmd.Hourglass True

Set Status = New Form_CT_Status
With Status
    .ShowCancelAll = False
    .ShowMessage = False
    .ShowCancel = True
    .ShowTime = True
    .ProgMax = SelRge.Height
    .show
End With
'Call Application.SysCmd(acSysCmdInitMeter, "Exporting to Excel", SelRge.Height)

If SelRge.Right > Grid.FldCT Then
    SelRge.Right = Grid.FldCT
End If


ReDim ArySort(Grid.FldCT + 1) As String
'SET THE COLUMN HEADERS


AryTmp = GetColumnsOrdered(Grid)
ReDim AryCols(SelRge.Columns) As CnlySelColumn

'TRIM TO ONLY SELECTED
For X = 0 To UBound(AryTmp)
    If AryTmp(X).order >= SelRge.left And AryTmp(X).order <= SelRge.Right Then  'IT IS A KEEPER
        ReDim Preserve AryCols(SelRge.Columns) As CnlySelColumn
        'Debug.Print AryTmp(X).CtrlName
        AryCols(SelRge.Columns) = AryTmp(X)
        SelRge.Columns = SelRge.Columns + 1
    End If
Next X


'Create The Basic Excel Objects
Set xlApp = CreateObject("Excel.Application")
With xlApp
    .visible = False
    Set XlWbk = .Workbooks.Add
    .UserControl = False
End With
Set XlSht = XlWbk.Worksheets(1)
        
'FIRST ROW
With XlSht
    For X = 0 To SelRge.Columns - 1
        With .Cells(1, X + 1)
            .Value = AryCols(X).Caption
            .Font.Bold = True
            .Borders.color = 12632256
            .Font.color = 0

            '.Width = AryCols(X).Width
            With .Interior
                .ColorIndex = 15
                .Pattern = 1 'xlSolid
                .PatternColorIndex = -4105 'xlAutomatic
            End With
        End With
    Next X
End With

'Get the form and its recordset.
Set rst = Grid.RecordSet
With rst
    ' Move to the first record in the recordset.
    .MoveFirst
    ' Move to the first selected record.
    .Move SelRge.top - 1
    For Row = 1 To SelRge.Height
        If Row Mod 5 = 0 Then
            DoEvents
            With Status
                .ProgVal = Row
                If .EvalStatus(Canceled) = True Then
                    Exit For
                End If
            End With
'            Call Application.SysCmd(acSysCmdInitMeter, "Exporting to Excel " & CStr(Row) & "/" & CStr(SelRge.Height), SelRge.Height)
'            Call Application.SysCmd(acSysCmdUpdateMeter, Row)
            DoEvents
        End If
        For X = 0 To SelRge.Columns - 1
            With XlSht.Cells(Row + 1, X + 1)
                .NumberFormat = AryCols(X).FormatXl
                .Value = Grid.Controls(AryCols(X).CtrlName)
            End With
        Next X
      .MoveNext
    Next Row
    '.Close  *** DO NOT CLOSE AS IT REMOVES THE FORM RECORDSET
End With


'FORMAT THE COLUMNS
For X = 1 To SelRge.Columns
    With XlSht.Columns(X)
        .HorizontalAlignment = AryCols(X - 1).AllignXl
        .AutoFit
    End With
Next X


With xlApp
    .visible = True
    .UserControl = True
End With

ExportSelected = True
ExitNow:
    On Error Resume Next
'    Call Application.SysCmd(acSysCmdClearStatus)
    DoCmd.Hourglass False
    Set XlSht = Nothing
    Set XlWbk = Nothing
    Set xlApp = Nothing
    Set TxtFld = Nothing
    Set Status = Nothing
    Set rst = Nothing
    Exit Function
ErrorHappened:
    RaiseEvent Status("ExportSelected", Err.Description, 10)
    If MvQuiet = False Then
        MsgBox Err.Description, vbCritical, "ClsExcel.ExportSelected"
    End If
    Resume ExitNow
    
End Function

Public Property Let IncludeFormats(data As Boolean)
    MvIncludeFormats = data
End Property
Public Property Get IncludeFormats() As Boolean
    IncludeFormats = MvIncludeFormats
End Property

Public Property Get TmpQryName() As String
    TmpQryName = MvQryName
End Property
Public Property Let QuietMode(data As Boolean)
    MvQuiet = data
End Property
Public Property Get QuietMode() As Boolean
    QuietMode = MvQuiet
End Property

Public Property Let FileName(data As String)
    MvFileName = data
End Property
Public Property Get FileName() As String
    FileName = MvFileName
End Property

'DLC 12/01/2011 - Enabled error handling and standardized message when cancelling export using Esc
Public Function ExportQuery(QryName As String) As Boolean
On Error GoTo ErrorHappened
    Dim db As DAO.Database, Qry As DAO.QueryDef
    Dim ExportAsXLS As Boolean
    If "" & MvFileName = "" Then
        RaiseEvent Status("ExportQuery", "No output file specified", 10)
        If MvQuiet = False Then
            MsgBox "No output file specified", vbCritical, "ClsExcel.ExportQuery"
        End If
        GoTo ExitNow
    Else  'FILE SPECIFIED
        If MvOverWrite = True Then
            If CreateObject("Scripting.FileSystemObject").FileExists(MvFileName) = True Then
                Kill MvFileName
            End If
        Else
            If CreateObject("Scripting.FileSystemObject").FileExists(MvFileName) = True Then
                If MsgBox("File Already Exists:" & vbCrLf & vbCrLf & MvFileName & vbCrLf & vbCrLf & "Overwrite?", vbQuestion + vbYesNo, "Overwrite File") = vbYes Then
                    Kill MvFileName
                Else
                    RaiseEvent Status("ExportQuery", "Output file exists and operation canceled", 10)
                    If MvQuiet = False Then
                        MsgBox "Output file exists and operation canceled", vbCritical, "ClsExcel.ExportQuery"
                    End If
                    GoTo ExitNow
                End If
            End If
        End If
    End If
    
    Set db = CurrentDb
    For Each Qry In db.QueryDefs
        If UCase(Qry.Name) = UCase(QryName) Then
            Exit For
        End If
    Next Qry
    
    If UCase(Qry.Name) = UCase(QryName) Then
        
        If Len(MvFileName) > 4 Then
            ExportAsXLS = Right(MvFileName, 4) = ".xls"
        Else
            ExportAsXLS = False
        End If
        
        If MvIncludeFormats = False Then
            'if format was not selected export
            DoCmd.TransferSpreadsheet acExport, IIf(ExportAsXLS, acSpreadsheetTypeExcel9, acSpreadsheetTypeExcel12Xml), QryName, MvFileName, True ', , True
            If MvAutoStart = True Then
                'open file after export
                OpenURL MvFileName, vbMaximizedFocus
            End If
        Else
            'if format was selected export
            DoCmd.OutputTo acOutputQuery, QryName, IIf(ExportAsXLS, acFormatXLS, acFormatXLSX), MvFileName, MvAutoStart
        End If
    End If
    
    ExportQuery = True
    
ExitNow:
    On Error Resume Next
    Set Qry = Nothing
    Set db = Nothing
    Exit Function
ErrorHappened:
    RaiseEvent Status("ExportQuery", Err.Description, 10)
    If MvQuiet = False Then
        If Err = 2501 Or Err = 3059 Then  '"The OutputTo action was canceled" or "Operation canceled by user"
            MsgBox "Export cancelled by user", vbCritical, "ClsExcel.ExportQuery"
        Else
            MsgBox Err.Description, vbCritical, "ClsExcel.ExportQuery"
        End If
    End If
    Resume ExitNow

End Function


Public Property Let ShowUI(data As Boolean)
    MvShowUI = data
End Property
Public Property Get ShowUI() As Boolean
    ShowUI = MvShowUI
End Property

Public Property Let FormID(data As Byte)
    mvFormId = data
    #If ccSCR = 1 Then
    Set MvFrmScr = Scr(mvFormId)
    #End If
End Property

Public Property Get FormID() As Byte
    FormID = mvFormId
End Property

Public Sub RunOther(ByRef oFrm As Form)
    Dim sTitle As String
    If MvShowUI = True Then
        Set MvFrmExcel = New Form_CT_Excel
        MvRunning = True
        With MvFrmExcel
            .cmbExportGrid.RowSource = "SELECT GridDesc,GridName FROM CT_ExportExcel WHERE FormName='" & oFrm.Name & "'"
            If .cmbExportGrid.ListCount > 0 Then
                .cmbExportGrid.Value = .cmbExportGrid.Column(0, 0)
            End If
            .visible = True
            .tbGrid.Pages(1).PageIndex = 0
             sTitle = DLookup("Title", "CT_ExportExcel", "Formname='" & oFrm.Name & "'")
             .LblScreenName1.ControlSource = "='Excel -> " & sTitle & "'"
            .LblScreenName2.ControlSource = .LblScreenName1.ControlSource
            If Nz(MvFileName, "") = "" Then
                MvFileName = Identity.FolderOutput & "\" & Identity.Computer & "." & sTitle & ".xlsx"
            End If
            .ExcelClass = Me
            .Refresh
            Do Until .ReturnValue > -1
                DoEvents
            Loop
            
            If .ReturnValue = vbOK Then
                SetupExportCustom .cmbExportGrid.Column(1, .cmbExportGrid.ListIndex), oFrm
            End If
        End With
        Set MvFrmExcel = Nothing
        
    End If

End Sub

Public Sub Run()
    'TODO - Remove compiler switches from this method when possible
    If MvShowUI = True Then
        Set MvFrmExcel = New Form_CT_Excel
        MvRunning = True
        With MvFrmExcel
            .visible = True
            .tbGrid.Pages(0).PageIndex = 0
            #If ccSCR = 1 Then
            .LblScreenName1.ControlSource = "='Excel -> " & MvFrmScr.ScreenName & "'"
            #End If
            .LblScreenName2.ControlSource = .LblScreenName1.ControlSource
            If "" & MvFileName = "" Then
                #If ccSCR = 1 Then
                MvFileName = Identity.FolderOutput & "\" & Identity.Computer & "." & MvFrmScr.ScreenName & ".xlsx"
                #End If
            End If
            .ExcelClass = Me
            .Recalc
            Do Until .ReturnValue > -1
                DoEvents
            Loop

            If .ReturnValue = vbOK Then
                SetupExport ExpObject
                .visible = False
            Else
                .visible = False
            End If

        End With
        Set MvFrmExcel = Nothing

    End If
End Sub


' DLC 10/25/2011 - Added ability to export based on a sql select statement
Public Sub ExportSql(ByVal strSQL As String)
On Error GoTo ErrorHappened
    Dim MvQryName As String
    Dim db As DAO.Database
    Dim Qry As DAO.QueryDef
    MvQryName = "ExcelTmp" & Identity.Computer
    Set db = CurrentDb
    Set Qry = New DAO.QueryDef
    With Qry
        .Name = MvQryName
        .SQL = strSQL
    End With
    On Error Resume Next
    db.QueryDefs.Delete MvQryName
On Error GoTo ErrorHappened
    db.QueryDefs.Append Qry
    ExportQuery MvQryName
    db.QueryDefs.Delete MvQryName
ExitNow:
    On Error Resume Next
    Set db = Nothing
    Set Qry = Nothing
    Exit Sub
ErrorHappened:
    RaiseEvent Status("SetupExport", Err.Description, 10)
    If Not MvQuiet Then
        MsgBox Err.Description, vbCritical, "ClsExcel.ExportSql"
    End If
    Resume ExitNow
End Sub


'this function creates the SQL query used for exporting data to excel
Public Function SetupExportCustom(ByVal sGrid As String, ByRef oFrm As Form) As Boolean
    Dim ArgCols() As CnlySelColumn
    Dim SQL As String
    Dim X As Integer
    Dim oGrid As Form_CT_SubGenericDataSheet
    Dim MvQryName As String
    Dim db As DAO.Database
    Dim Qry As DAO.QueryDef
    Dim fld As DAO.Field
    Dim prp As DAO.Property

    
    Set oGrid = oFrm.Controls(sGrid).Form
    ArgCols = GetColumnsOrdered(oGrid)
    
    SQL = "Select "
    'columns
    For X = 0 To UBound(ArgCols)
        With ArgCols(X)
            If .Calculated = False Then  'Standard Field
                SQL = SQL & " [" & .ControlSource & "]"
            Else 'Calculated Field -- 'Stip the Equal sign
                SQL = SQL & " (" & .ControlSource & ") as " & .Caption & "Calc"
            End If
        End With
        If X <> UBound(ArgCols) Then
            SQL = SQL & ","
        End If
    Next X
    
    'create sql
    SQL = SQL & " FROM (" & oGrid.RecordSource & ") "
    
    If oGrid.FilterOn And oGrid.filter <> "" Then
        SQL = SQL & " WHERE " & oGrid.filter
    End If
    If oGrid.OrderByOn And oGrid.OrderBy <> "" Then
        SQL = SQL & " ORDER BY " & oGrid.OrderBy
    End If
    
    If "" & SQL <> "" Then
        MvQryName = "ExcelTmp" & Identity.Computer
        Set db = CurrentDb
        Set Qry = New DAO.QueryDef
        With Qry
            .Name = MvQryName
            .SQL = SQL
        End With
        On Error Resume Next
        db.QueryDefs.Delete MvQryName
        On Error GoTo ErrorHappened
        db.QueryDefs.Append Qry
        
        With Qry
            For Each fld In Qry.Fields
                'Debug.Print ArgCols(Fld.OrdinalPosition).Caption
                If "" & ArgCols(fld.OrdinalPosition).Format <> "" Then
                    Set prp = fld.CreateProperty("Format", dbText, "" & ArgCols(fld.OrdinalPosition).Format)
                    fld.Properties.Append prp
                End If
                If "" & ArgCols(fld.OrdinalPosition).Decimals <> "" Then
                    Set prp = fld.CreateProperty("DecimalPlaces", dbText, "" & ArgCols(fld.OrdinalPosition).Decimals)
                    fld.Properties.Append prp
                End If
            Next fld
        End With
        db.QueryDefs.Refresh
    End If
    
    'export the created SQL query
    ExportQuery MvQryName
    db.QueryDefs.Delete MvQryName
ExitNow:
        On Error Resume Next
        Set Qry = Nothing
        Set db = Nothing
        Set prp = Nothing
        Set fld = Nothing
        Exit Function
ErrorHappened:
        RaiseEvent Status("SetupExport", Err.Description, 10)
        If MvQuiet = False Then
            MsgBox Err.Description, vbCritical, "ClsExcel.SetupExport"
        End If
        Resume ExitNow
  
End Function


Public Function SetupExport(Exp As CnlyExportObj) As Boolean
On Error GoTo ErrorHappened
Dim db As DAO.Database, Qry As DAO.QueryDef, TxtFld As Access.TextBox, fld As DAO.Field
Dim AryCols() As CnlySelColumn
Dim SQL As String, X As Long, Formats(1) As String, FormatsCol As New Collection
Dim MasterFields() As String, ChildFields() As String, StQual As String, SqlWhere As String
Dim prp As DAO.Property

'If active tab selected and totals are active then change type
'TODO Remove compiler switch if possible
#If ccSCR = 1 Then
If Exp = CnlyExportObj.ActiveTab And MvFrmScr.Tabs = 0 Then
    Exp = CnlyExportObj.Totals
End If

Select Case Exp
Case CnlyExportObj.MainGrid  '****** EXPORT THE MAIN GRID ****************************************
    AryCols = GetColumnsOrdered(MvFrmScr.GridForm)
Case CnlyExportObj.ActiveTab  '****** EXPORT A TAB ****************************************
    AryCols = GetColumnsOrdered(MvFrmScr.Tabs.Pages(MvFrmScr.Tabs.Value).Controls(0).Form)
Case CnlyExportObj.Totals   '****** EXPORT THE TOTALS ****************************************
    AryCols = GetColumnsOrdered(MvFrmScr.Controls("Subform1").Form)
End Select
#End If

'*** WRITE THE SQL SELECT STATEMENT ***
SQL = "Select "
For X = 0 To UBound(AryCols)
    With AryCols(X)
        If .Calculated = False Then  'Standard Field
            SQL = SQL & " [" & .ControlSource & "]"
        Else 'Calculated Field -- 'Stip the Equal sign
            SQL = SQL & " (" & .ControlSource & ") as " & .Caption & "Calc"
        End If
    End With
    If X <> UBound(AryCols) Then
        SQL = SQL & ","
    End If
Next X

'TODO Remove compiler switch if possible
#If ccSCR = 1 Then
Select Case Exp
Case CnlyExportObj.MainGrid  '****** EXPORT THE MAIN GRID ****************************************
    SQL = SQL & " " & MvFrmScr.SQL.From
    SQL = SQL & "Where "
    SQL = SQL & MvFrmScr.BuildWhere
    MvFrmScr.BuildWhere
    With MvFrmScr.SQL
        If MvFrmScr.GridForm.OrderByOn = False Then
            If "" & .OrderBy <> "" Then
                SQL = SQL & " ORDER BY " & .OrderBy
            End If
        Else
            If "" & MvFrmScr.GridForm.OrderBy <> "" Then
                SQL = SQL & " ORDER BY " & MvFrmScr.GridForm.OrderBy
            Else
                ' DS Mar 11 2010 fix from Dave Brady Change Request ID#1367
                If "" & .OrderBy <> "" Then
                    SQL = SQL & " ORDER BY " & .OrderBy
                End If
            End If
        End If
    End With

Case CnlyExportObj.ActiveTab  '****** EXPORT A TAB ****************************************
    With MvFrmScr.Tabs.Pages(MvFrmScr.Tabs).Controls(0)
        SQL = SQL & " FROM (" & .Form.RecordSource & ") "

        MasterFields = Split(Replace(.LinkMasterFields, "SubForm.Form!", ""), ";")
        ChildFields = Split(Replace(.LinkChildFields, "SubForm.Form!", ""), ";")

        If UBound(MasterFields) >= 0 Then
            For X = 0 To UBound(MasterFields)
                If "" & SqlWhere <> "" Then
                    SqlWhere = SqlWhere & " AND"
                End If

                SqlWhere = SqlWhere & " [" & ChildFields(X) & "] = "
                Select Case .Form.RecordsetClone(ChildFields(X)).Type
                Case dbText
                    StQual = "'"
                Case dbDate
                    StQual = "#"
                Case Else
                    StQual = ""
                End Select
                SqlWhere = SqlWhere & StQual & MvFrmScr.GridForm(MasterFields(X)) & StQual
            Next X
        End If

        If .Form.FilterOn And "" & .Form.filter <> "" Then
            If "" & SqlWhere <> "" Then
                SqlWhere = SqlWhere & " AND "
            End If
            SqlWhere = SqlWhere & "(" & .Form.filter & ")"
        End If
        If "" & SqlWhere <> "" Then
            SQL = SQL & " WHERE " & SqlWhere
        End If
        If .Form.OrderByOn And "" & .Form.OrderBy <> "" Then
            SQL = SQL & " ORDER BY "
            SQL = SQL & .Form.OrderBy
        End If

    End With
Case CnlyExportObj.Totals   '****** EXPORT THE TOTALS ****************************************
    With MvFrmScr.Controls("Subform1")
        SQL = SQL & " FROM (" & .Form.RecordSource & ") as Ttl "
        If .Form.FilterOn And "" & .Form.filter <> "" Then
            SQL = SQL & " WHERE "
            SQL = SQL & "(" & .Form.filter & ")"
        End If
        If .Form.OrderByOn And "" & .Form.OrderBy <> "" Then
            SQL = SQL & " ORDER BY "
            SQL = SQL & .Form.OrderBy
        End If
    End With
End Select
#End If

If "" & SQL <> "" Then
    MvQryName = "ExcelTmp" & Identity.Computer
    Set db = CurrentDb
    Set Qry = New DAO.QueryDef
    With Qry
        .Name = MvQryName
        .SQL = SQL
    End With
    On Error Resume Next
    db.QueryDefs.Delete MvQryName
    On Error GoTo ErrorHappened
    db.QueryDefs.Append Qry
    
    With Qry
        For Each fld In Qry.Fields
            'Debug.Print AryCols(Fld.OrdinalPosition).Caption
            If "" & AryCols(fld.OrdinalPosition).Format <> "" Then
                Set prp = fld.CreateProperty("Format", dbText, "" & AryCols(fld.OrdinalPosition).Format)
                fld.Properties.Append prp
            End If
            If "" & AryCols(fld.OrdinalPosition).Decimals <> "" Then
                Set prp = fld.CreateProperty("DecimalPlaces", dbText, "" & AryCols(fld.OrdinalPosition).Decimals)
                fld.Properties.Append prp
            End If
        Next fld
    End With
    db.QueryDefs.Refresh
End If

ExportQuery MvQryName
'DB.QueryDefs.Refresh
'DBEngine.Idle dbRefreshCache + dbForceOSFlush
db.QueryDefs.Delete MvQryName
ExitNow:
    On Error Resume Next
    Set Qry = Nothing
    Set db = Nothing
    Set TxtFld = Nothing
    Set prp = Nothing
    Set fld = Nothing
    Set FormatsCol = Nothing
    Exit Function
ErrorHappened:
    RaiseEvent Status("SetupExport", Err.Description, 10)
    If MvQuiet = False Then
        MsgBox Err.Description, vbCritical, "ClsExcel.SetupExport"
    End If
    Resume ExitNow
End Function

Private Sub Class_Initialize()
    MvShowUI = True
End Sub

Private Sub Class_Terminate()
'Delete the temporary query
Set MvFrmExcel = Nothing
'Set MvFrmExcelOther = Nothing
If "" & MvQryName <> "" Then
    On Error Resume Next
    'CurrentDb.QueryDefs.Delete MvQryName
End If

End Sub
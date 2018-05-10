Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Const CnstCtlSpacing As Single = 0.01

Private subGen As CT_ClsSubGenericDataSheet

'Change number to match number of fields in this grid
Private Const SheetFieldCount As Integer = 255

Private csPrimaryCriteria As String

Public Event Activate()
Public Event ApplyFilter(filter As String)
Public Event Current()
Public Event Click()
Public Event Deactivate()
Public Event FocusLost()
Public Event FocusGot()
Public Event Message(Txt As String)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Unload()
Public Event KeyPressed(AsciiKey As Integer)
Private bAdo As Boolean
Private cbInitialLoad As Boolean

Private MvFldCT As Integer
Private MvIsCustomTotal As Boolean
'**********
'** Added by David.Brady to support multiple row selection from the grid.
Private MvSelTop As Long
Private MvSelHeight As Long

Public Property Get ScreenID() As Long
    ScreenID = 1
End Property

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get PrimaryCriteria() As String
    PrimaryCriteria = csPrimaryCriteria
End Property

Public Property Get CompletedLoad() As Boolean
    CompletedLoad = Not cbInitialLoad
End Property


'********** Stuff to support the layout saving
Public Sub LayoutClear()
    subGen.LayoutClear
End Sub

Public Sub LayoutField(Name As String, Ordinal As Long, Width As Single, CalcFld As Boolean)
    subGen.LayoutField Name, Ordinal, Width, CalcFld
End Sub


'**********
Property Get SelectionTop() As Long
    SelectionTop = MvSelTop
End Property

Property Get SelectionHeight() As Long
    SelectionHeight = MvSelHeight
End Property

Property Get FldCT() As Integer
    FldCT = MvFldCT
End Property

Public Sub InitData(RecordSource As String, RecordSourceType As Byte)
On Error GoTo Block_Err
Dim strProcName As String
Dim InitDb As DAO.Database
Dim oDb As DAO.Database
Dim oQDef As DAO.QueryDef
Dim sQueryName As String
Dim sSql As String
Dim strConnect As String
Dim InitRst As DAO.RecordSet
Dim InitFld As DAO.Field
Dim CurLeft As Long
Dim CurField As CnlyFldDef
Dim NewCtrl As Control
Dim CurNum As Integer
Dim ErrMsg As String
Dim ctl As Control
Dim sConnString As String
    
    strProcName = ClassName & ".InitData"
    
    subGen.SetMyForm = Me
    subGen.SetFieldCount = SheetFieldCount
    
    Set InitDb = CurrentDb


    If "" & RecordSource = "" Then GoTo Block_Exit

    If InStr(1, RecordSource, " ", vbTextCompare) > 1 Then
        Stop
    Else
        csPrimaryCriteria = RecordSource
    End If

    Select Case RecordSourceType
        Case 0 'TABLE
            Me.RecordSource = RecordSource
            
'            Set InitRst = InitDb.TableDefs(RecordSource)
            Set InitRst = Me.RecordSet
        Case 1 'Query
            Set InitRst = InitDb.QueryDefs(RecordSource)
        Case 2  ' Passthrough query

            ' if the query we are about to build exists, delete it first
            



'            strConnect = "DRIVER=SQL Server;SERVER=DS-FLD-009;DATABASE=CMS_AUDITORS_CMS;TRUSTED_CONNECTION=YES;"
            strConnect = "DRIVER=SQL Server;DSN=DS-FLD-009;DATABASE=CMS_AUDITORS_CMS;Trusted_Connection=Yes;APP=ClaimAdmin;TABLE=dbo.CONCEPT_ConMgmtSearch"
            strConnect = "ODBC;DSN=DS-FLD-009;UID=kevin.dearing;Trusted_Connection=Yes;DATABASE=CMS_AUDITORS_CLAIMS"
            
            sQueryName = "qry_pt_" & RecordSource
            sSql = "SELECT * FROM " & RecordSource & " WHERE SearchUserId = '" & Identity.UserName & "'"
            
            Set oDb = CurrentDb()
            If IsQuery(sQueryName) = True Then
                oDb.QueryDefs.Delete sQueryName
            End If
            Set oQDef = oDb.CreateQueryDef(sQueryName, sSql)
            oQDef.Connect = "ODBC;" & strConnect
            oQDef.ReturnsRecords = (Len(sQueryName) > 0)
            If oQDef.ReturnsRecords = False Then
                oQDef.Execute
            Else

'                Call oDB.QueryDefs.Append(oQDef)
                
                oDb.QueryDefs.Refresh
            End If
            oQDef.Close

            ' So now let's just set the recordsource to the name of our query:
            Me.RecordSource = sSql
            Set InitRst = Me.RecordSet
        Case Else
            Set InitRst = CurrentDb.OpenRecordSet(RecordSource, dbOpenSnapshot)
    End Select

    Me.Section(acHeader).visible = False
    Me.Section(acDetail).visible = False

    'Configure The User Preferences for the grid
    With Identity.DataSheetStyle
        Me.DatasheetBackColor = CLng(.BackGroundColor)
        Me.DatasheetBorderLineStyle = .BorderLineStyle
        Me.DatasheetCellsEffect = .CellsEffect
        Me.DatasheetColumnHeaderUnderlineStyle = .HeaderUnderlineStyle
        Me.DatasheetFontHeight = .fontsize
        Me.DatasheetFontItalic = .FontItalic
        Me.DatasheetFontName = .FontFamily
        Me.DatasheetFontUnderline = .FontUnderline
        Me.DatasheetFontWeight = .FontWeight
        Me.DatasheetForeColor = .ForeColor
        Me.DatasheetGridlinesBehavior = .GridlinesBehavior
        Me.DatasheetGridlinesColor = .GridlinesColor
    End With

    For Each InitFld In InitRst.Fields
        CurNum = CurNum + 1
        With CurField
            .Type = InitFld.Type
            .Name = InitFld.Name
            .left = .left + .Width + CnstCtlSpacing
            .Width = GetFieldWidth(.Type, IIf(InitFld.Type = 10, InitFld.Size, 0), RecordSource, .Name, 0)
            .Height = GetFieldHeight(Identity.DataSheetStyle.fontsize)
        End With
        AddField CurNum, CurField
    Next InitFld

    MvFldCT = CurNum
    For CurNum = CurNum + 1 To 255
        With Me.Controls.Item("Field" & CStr(CurNum))
            .ColumnOrder = CurNum
            .ColumnHidden = True
            .ColumnWidth = 0
            .Enabled = False
            .visible = False
        End With
    Next CurNum

    For Each ctl In Me.Controls
        If ctl.ControlType = acTextBox Then
            ctl.ColumnWidth = -2
        End If
    Next

Block_Exit:
    On Error Resume Next
    Me.Section(acHeader).visible = True
    Me.Section(acDetail).visible = True
    Me.Repaint
    Set InitRst = Nothing
    Set InitFld = Nothing
    Set InitDb = Nothing
    Exit Sub
    
Block_Err:
    Select Case Err.Number
    Case 3061 'Too few Params specified error
        If MvIsCustomTotal Then
            ErrMsg = "Field(s) specified in your Custom Total are misspelled or no longer exist."
        Else
            ErrMsg = "Field(s) specified in one of the grids are misspelled or no longer exist."
            ErrMsg = ErrMsg & vbCrLf & "RecordSource: " & RecordSource
        End If
    Case Else
        ErrMsg = ""
    End Select
    
    ErrMsg = ErrMsg & vbCrLf & vbCrLf & Err.Description
    
    MsgBox ErrMsg, vbCritical, "Error Initializing Tabbed Fields"
    GoTo Block_Exit
End Sub


''' KD: Added below because above is using DAO (ancient chineses secret) which essentially
''' is using the linked tables via JET - SLOWWWWWW!!!
Public Sub InitDataADO(oRs As ADODB.RecordSet)
On Error GoTo InitDataError
Dim oFld As ADODB.Field
Dim CurLeft As Long
Dim CurField As CnlyFldDef
Dim NewCtrl As Control
Dim CurNum As Integer
Dim ErrMsg As String
Dim ctl As Control

    bAdo = True

    Me.Section(acHeader).visible = False
    Me.Section(acDetail).visible = False

    'Configure The User Preferences for the grid
    With Identity.DataSheetStyle
        Me.DatasheetBackColor = CLng(.BackGroundColor)
        Me.DatasheetBorderLineStyle = .BorderLineStyle
        Me.DatasheetCellsEffect = .CellsEffect
        Me.DatasheetColumnHeaderUnderlineStyle = .HeaderUnderlineStyle
        Me.DatasheetFontHeight = .fontsize
        Me.DatasheetFontItalic = .FontItalic
        Me.DatasheetFontName = .FontFamily
        Me.DatasheetFontUnderline = .FontUnderline
        Me.DatasheetFontWeight = .FontWeight
        Me.DatasheetForeColor = .ForeColor
        Me.DatasheetGridlinesBehavior = .GridlinesBehavior
        Me.DatasheetGridlinesColor = .GridlinesColor
    End With

    For Each oFld In oRs.Fields
        CurNum = CurNum + 1
        With CurField
            .Type = AdoTypeToDaoType(oFld)
            .Name = oFld.Name
            .left = .left + .Width + CnstCtlSpacing
            '.width = GetFieldWidth(.Type, IIf(oFld.Type = 10, oFld.Precision, 0), RecordSource, .Name, 0)
            .Width = GetFieldWidth(AdoTypeToDaoType(oFld), IIf(AdoTypeToDaoType(oFld) = 10, oFld.Precision, 0), RecordSource, .Name, 0)
            .Height = GetFieldHeight(Identity.DataSheetStyle.fontsize)
        End With
        AddField CurNum, CurField
    Next

    MvFldCT = CurNum
    For CurNum = CurNum + 1 To 255
        With Me.Controls.Item("Field" & CStr(CurNum))
            .ColumnOrder = CurNum
            .ColumnHidden = True
            .ColumnWidth = 0
            .Enabled = False
            .visible = False
        End With
    Next CurNum

    For Each ctl In Me.Controls
        If ctl.ControlType = acTextBox Then
            ctl.ColumnWidth = -2
        End If
    Next

InitDataExit:
    On Error Resume Next
    Me.Section(acHeader).visible = True
    Me.Section(acDetail).visible = True
    Me.Repaint
    Set oFld = Nothing
    Exit Sub
    
InitDataError:
    Select Case Err.Number
    Case 3061 'Too few Params specified error
        If MvIsCustomTotal Then
            ErrMsg = "Field(s) specified in your Custom Total are misspelled or no longer exist."
        Else
            ErrMsg = "Field(s) specified in one of the grids are misspelled or no longer exist."
            ErrMsg = ErrMsg & vbCrLf & "RecordSource: " & RecordSource
        End If
    Case Else
        ErrMsg = ""
    End Select
    
    ErrMsg = ErrMsg & vbCrLf & vbCrLf & Err.Description
    
    MsgBox ErrMsg, vbCritical, "Error Initializing Tabbed Fields"
    Resume InitDataExit
    Resume
End Sub


''' KD: Added below because above is using DAO (ancient chineses secret) which essentially
''' is using the linked tables via JET - SLOWWWWWW!!!
Public Sub InitDataDAO(oRs As DAO.RecordSet)
On Error GoTo InitDataError
Dim oFld As ADODB.Field
Dim CurLeft As Long
Dim CurField As CnlyFldDef
Dim NewCtrl As Control
Dim CurNum As Integer
Dim ErrMsg As String
Dim ctl As Control

    bAdo = True

    Me.Section(acHeader).visible = False
    Me.Section(acDetail).visible = False

    'Configure The User Preferences for the grid
    With Identity.DataSheetStyle
        Me.DatasheetBackColor = CLng(.BackGroundColor)
        Me.DatasheetBorderLineStyle = .BorderLineStyle
        Me.DatasheetCellsEffect = .CellsEffect
        Me.DatasheetColumnHeaderUnderlineStyle = .HeaderUnderlineStyle
        Me.DatasheetFontHeight = .fontsize
        Me.DatasheetFontItalic = .FontItalic
        Me.DatasheetFontName = .FontFamily
        Me.DatasheetFontUnderline = .FontUnderline
        Me.DatasheetFontWeight = .FontWeight
        Me.DatasheetForeColor = .ForeColor
        Me.DatasheetGridlinesBehavior = .GridlinesBehavior
        Me.DatasheetGridlinesColor = .GridlinesColor
    End With

    For Each oFld In oRs.Fields
        CurNum = CurNum + 1
        With CurField
            .Type = AdoTypeToDaoType(oFld)
            .Name = oFld.Name
            .left = .left + .Width + CnstCtlSpacing
            '.width = GetFieldWidth(.Type, IIf(oFld.Type = 10, oFld.Precision, 0), RecordSource, .Name, 0)
            .Width = GetFieldWidth(AdoTypeToDaoType(oFld), IIf(AdoTypeToDaoType(oFld) = 10, oFld.Precision, 0), RecordSource, .Name, 0)
            .Height = GetFieldHeight(Identity.DataSheetStyle.fontsize)
        End With
        AddField CurNum, CurField
    Next

    MvFldCT = CurNum
    For CurNum = CurNum + 1 To 255
        With Me.Controls.Item("Field" & CStr(CurNum))
            .ColumnOrder = CurNum
            .ColumnHidden = True
            .ColumnWidth = 0
            .Enabled = False
            .visible = False
        End With
    Next CurNum

    For Each ctl In Me.Controls
        If ctl.ControlType = acTextBox Then
            ctl.ColumnWidth = -2
        End If
    Next

InitDataExit:
    On Error Resume Next
    Me.Section(acHeader).visible = True
    Me.Section(acDetail).visible = True
    Me.Repaint
    Set oFld = Nothing
    Exit Sub
    
InitDataError:
    Select Case Err.Number
    Case 3061 'Too few Params specified error
        If MvIsCustomTotal Then
            ErrMsg = "Field(s) specified in your Custom Total are misspelled or no longer exist."
        Else
            ErrMsg = "Field(s) specified in one of the grids are misspelled or no longer exist."
            ErrMsg = ErrMsg & vbCrLf & "RecordSource: " & RecordSource
        End If
    Case Else
        ErrMsg = ""
    End Select
    
    ErrMsg = ErrMsg & vbCrLf & vbCrLf & Err.Description
    
    MsgBox ErrMsg, vbCritical, "Error Initializing Tabbed Fields"
    Resume InitDataExit
    Resume
End Sub


Private Sub AddField(FldNum As Integer, fld As CnlyFldDef)
    With Me("Field" & CStr(FldNum))
        .ControlSource = fld.Name
        .Tag = ""
        .Width = fld.Width * 1440 'DOES NOT NEED TO BE FIXED AND IS USED FOR LAYOUTS
        .ColumnOrder = FldNum
        If fld.Width * 1440 = 0 Then
            .ColumnHidden = True
            .ColumnWidth = 0
            .visible = False
            .Locked = True
        Else
            .ColumnHidden = False
            .visible = True
            .Locked = False
            'Damon Added to fix bug with disabled fields
            .Enabled = True
            .ColumnWidth = fld.Width * 1440  'IN POINTS - NOW SURE WHY
        End If
        .Height = fld.Height * 1440
        If fld.Format = "Hyperlink" Then
            .IsHyperlink = True
        Else
            .Format = fld.Format
        End If

        'Alex Added 9/22/08
        'MsgBox fld.Name + " " + CStr(fld.Type)
        Select Case fld.Type
          Case Is = 10 'Text
            .TextAlign = 1
          Case Is = 12 'Some other type of text
            .TextAlign = 1
          Case Is = 8 'Date
            .TextAlign = 2
            .Format = "MM/DD/YY"
          Case Is = 4 'Integer
            .TextAlign = 3
            .Format = "#,###"
            .DecimalPlaces = 0
          Case Is = 5 'Money
            .TextAlign = 3
            .Format = "#,###.##"
            .DecimalPlaces = 2
          Case Is = 7 'Decimal
            .TextAlign = 3
            .Format = "#,###.##"
            .DecimalPlaces = 2
          Case Is = 20 'some sort of other decimial
            .TextAlign = 3
            .Format = "#,###.##"
            .DecimalPlaces = 2
         End Select
    End With
    
    With Me("LblField" & CStr(FldNum))
        .Caption = GetSplitFieldNameForLabel(fld.Alias, fld.Name)
        .TextAlign = fld.Align
        '.Left = Fld.Left * 1440  -FIX SIZE ISSUE
        '.Width = Fld.Width * 1440  -FIX SIZE ISSUE
        .Height = fld.Height * 1440 * 1.78
        .visible = True
    End With
    
    With Me("CapField" & CStr(FldNum))
        '.Width = Fld.Width * 1440  -FIX SIZE ISSUE
        If "" & fld.Alias = "" Then
            .Caption = fld.Name
        Else
            .Caption = fld.Alias
        End If
    End With
    MvFldCT = MvFldCT + 1
End Sub


Private Sub Field1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field100_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub Field101_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field102_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub




Private Sub Field103_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Field11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub Field12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub Field16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub Field18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Field9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Dim StFilter As String
    
    Select Case ApplyType
        Case acShowAllRecords '0
            RaiseEvent ApplyFilter("")
            Me.FilterOn = False
        Case acApplyFilter '1
            StFilter = Replace(Me.filter, Me.Name & ".", "")
            If StFilter <> Me.filter Then
                Me.filter = Replace(Me.filter, Me.Name & ".", "")
            End If
            If bAdo Then
                Me.filter = Replace(Me.filter, "*", "%")
'                Stop
            End If
            RaiseEvent ApplyFilter(Me.filter)
        Case acCloseFilterWindow '2
    End Select
End Sub

Public Sub ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Form_ApplyFilter Cancel, ApplyType
End Sub

Private Sub Form_Click()
    RaiseEvent Click
End Sub

Private Sub Form_Current()
    'PLACE HOLDER FOR EVENT CAPTURE
    RaiseEvent Current
    cbInitialLoad = False   ' this ends the initial form load, focus goes back to the main form now, eventually after that form is finished loading we'll get GotFocus()
End Sub

Private Sub Form_DblClick(Cancel As Integer)
    'Damon added to control navigation
    
    On Error GoTo ErrHandler
        
    Dim strParameter As String
    Dim strParameterString As String
    
    Dim strError As String
    Dim strParent As String
    Dim arrParameters() As String
    Dim intI As Integer
    Dim strAppID As String
    
    
    strParameterString = ""
    
    If Me.Parent.Form.Name = "frm_GENERAL_Tab" Then
        strParent = Me.Parent.Form.Parent.Name
        strAppID = Me.Parent.Form.Parent.frmAppID
    Else
        strParent = Me.Parent.Form.Name
        strAppID = Me.Parent.frmAppID
        
    End If
    
    strParameter = Nz(DLookup("Parameter", "GENERAL_Navigate", "SearchType = '" & strAppID & "' and ActionName = 'dblClick' and parentform = '" & strParent & "'"), "")
    arrParameters = Split(strParameter, "|")
    
    If UBound(arrParameters) > 0 Then
        For intI = 0 To UBound(arrParameters)
           strParameterString = strParameterString & Me.RecordSet(arrParameters(intI)) & "|"
        Next intI
    Else
          strParameterString = strParameterString & Me.RecordSet(arrParameters(0))
    End If
    
    If strParameter <> "" Then
        Navigate strParent, strAppID, "DblClick", strParameterString
    End If

    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_General_Search : Navigate"
End Sub

Private Sub Form_Deactivate()
    RaiseEvent Deactivate
End Sub

Private Sub Form_GotFocus()
    RaiseEvent FocusGot
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '**********
    '** Added by David.Brady to support multiple row selection from the grid.
    MvSelHeight = Me.SelHeight
    MvSelTop = Me.SelTop
    '**********
End Sub

Private Sub Form_LostFocus()
    RaiseEvent FocusLost
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If screen.MousePointer = 7 Then screen.MousePointer = 0
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '**********
    '** Added by David.Brady to support multiple row selection from the grid.
    MvSelHeight = Me.SelHeight
    MvSelTop = Me.SelTop
    '**********
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.RecordSource = ""
    Set subGen = New CT_ClsSubGenericDataSheet
    cbInitialLoad = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    screen.MousePointer = 0
    Set subGen = Nothing
    RaiseEvent Unload
End Sub

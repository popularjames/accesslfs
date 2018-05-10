Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : CT_ClsSubGenericDataSheet
' Author    : SA
' Date      : 11/6/2012
' Purpose   : Moved code from forms CT_SubGenericDataSheet... into this common class
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private Const CnstCtlSpacing As Single = 0.01

Private myForm As Form
Private SheetFieldCount As Integer
Private FieldCount As Integer
Private SelectionTop As Long
Private SelectionHeight As Long
Private IsCustomTotal As Boolean

Public Property Let SetMyForm(ByRef frm As Form)
    Set myForm = frm
End Property

Public Property Let SetSheetFieldCount(ByVal Value As Integer)
    SheetFieldCount = Value
End Property
Public Property Get GetSheetFieldCount() As Integer
    GetSheetFieldCount = SheetFieldCount
End Property

Public Property Let SetFieldCount(ByVal Value As Integer)
    FieldCount = Value
End Property
Public Property Get GetFieldCount() As Integer
    GetFieldCount = FieldCount
End Property

Public Property Let SetSelectionTop(ByVal Value As Long)
    SelectionTop = Value
End Property
Public Property Get GetSelectionTop() As Long
    GetSelectionTop = SelectionTop
End Property

Public Property Let SetSelectionHeight(ByVal Value As Long)
    SelectionHeight = Value
End Property
Public Property Get GetSelectionHeight() As Long
    GetSelectionHeight = SelectionHeight
End Property

Public Property Let SetIsCustomTotal(ByVal Value As Boolean)
    IsCustomTotal = Value
End Property
Public Property Get GetIsCustomTotal() As Boolean
    GetIsCustomTotal = IsCustomTotal
End Property

Public Sub InitData(ByVal RecordSource As String, ByVal RecordSourceType As Byte, Optional ByVal useDataSource As String = vbNullString)
On Error GoTo InitDataError
    'SA 5/16/12 - Optimized code for faster screen loading
    Dim InitDb As DAO.Database
    Dim InitRst As Object
    Dim InitFld As Field
    Dim CurField As CnlyFldDef
    Dim CurNum As Integer
    Dim ErrMsg As String
    Dim FieldSettings As New CT_ClsFieldSettings
    Dim RowHeight As Integer
    
    If LenB(useDataSource) = 0 Then
        useDataSource = GetDataSource(RecordSource)
    End If
    
    Set InitDb = CurrentDb
    If LenB(Nz(RecordSource, vbNullString)) = 0 Then
        GoTo InitDataExit
    End If
    
    Select Case RecordSourceType
        Case 0 'TABLE
            Set InitRst = InitDb.TableDefs(RecordSource)
        Case 1 'Query
            Set InitRst = InitDb.QueryDefs(RecordSource)
        Case Else
            Set InitRst = CurrentDb.OpenRecordSet(RecordSource, dbOpenSnapshot)
    End Select
    
    myForm.Section(acHeader).visible = False
    myForm.Section(acDetail).visible = False
    
    'Configure The User Preferences for the grid
    With Identity.DataSheetStyle
        myForm.DatasheetBackColor = CLng(.BackGroundColor)
        myForm.DatasheetBorderLineStyle = .BorderLineStyle
        myForm.DatasheetCellsEffect = .CellsEffect
        myForm.DatasheetColumnHeaderUnderlineStyle = .HeaderUnderlineStyle
        myForm.DatasheetFontHeight = .fontsize
        myForm.DatasheetFontItalic = .FontItalic
        myForm.DatasheetFontName = .FontFamily
        myForm.DatasheetFontUnderline = .FontUnderline
        myForm.DatasheetFontWeight = .FontWeight
        myForm.DatasheetForeColor = .ForeColor
        myForm.DatasheetGridlinesBehavior = .GridlinesBehavior
        myForm.DatasheetGridlinesColor = .GridlinesColor
        
        'SA 11/15/2012 - Set row height (-1 = auto size)
        RowHeight = Nz(DLookup("Value", "CT_Options", "OptionName='DataSheetRowHeight'"), -1)
        If RowHeight <> 0 Then
            myForm.RowHeight = RowHeight
        End If
    End With
    
    FieldSettings.LoadFieldSettings myForm.Parent.ScreenID, useDataSource
    
    For Each InitFld In InitRst.Fields
        CurNum = CurNum + 1
        With CurField
            .Type = InitFld.Type
            .Name = InitFld.Name
            .Alias = FieldSettings.GetFieldAlias(.Name)
            .left = .left + .Width + CnstCtlSpacing
            .Width = FieldSettings.GetFieldWidth(.Name, .Type, IIf(InitFld.Type = 10, InitFld.Size, 0))
            .Height = GetFieldHeight(Identity.DataSheetStyle.fontsize)
            .Align = FieldSettings.GetFieldAlign(.Name, .Type)
            
            If IsCustomTotal Then
                .Format = GetFieldFormatCustomTotals(vbNullString & myForm.Tag, CurNum, .Type, .Decimal, .Name, myForm.Parent.ScreenID)
            Else
                .Format = FieldSettings.GetFieldFormat(.Name, .Type, .Decimal)
            End If
        End With
        
        AddField CurNum, CurField
    Next InitFld
    
    FieldCount = CurNum
    
    For CurNum = CurNum + 1 To SheetFieldCount
        With myForm.Controls.Item(FieldTextBoxString(CurNum))
            .ColumnOrder = CurNum
            .ColumnHidden = True
            .ColumnWidth = 0
            .Enabled = False
            .visible = False
        End With
    Next CurNum
InitDataExit:
On Error Resume Next
    myForm.Section(acHeader).visible = True
    myForm.Section(acDetail).visible = True
    myForm.Repaint
    Set InitRst = Nothing
    Set InitFld = Nothing
    Set InitDb = Nothing
Exit Sub
InitDataError:
    If Err.Number = 3061 Then
        'Too few Params specified error
        If IsCustomTotal Then
            ErrMsg = "Field(s) specified in your Custom Total are misspelled or no longer exist."
        Else
            ErrMsg = "Field(s) specified in one of the grids are misspelled or no longer exist." & _
                    vbCrLf & "RecordSource: " & RecordSource
        End If
        ErrMsg = ErrMsg & vbCrLf & vbCrLf & Err.Description
    Else
        ErrMsg = Err.Description
    End If
    
    MsgBox ErrMsg, vbCritical, "Error Initializing Tabbed Fields"
    Resume InitDataExit
    Resume
End Sub

Public Sub CalcFieldsAdd(ByRef fld As CnlyFldDef)
    AddField FieldCount + 1, fld
    myForm.InsideWidth = myForm.Parent.InsideWidth
End Sub

Public Sub FormatsClear()
'Remove all of the existing formats
On Error GoTo ErrorHappened
    Dim TxtFld As Access.TextBox
    Dim X As Long
    
    For X = 1 To FieldCount
        Set TxtFld = FieldTextBox(X)
        TxtFld.FormatConditions.Delete
    Next X
    
ExitNow:
On Error Resume Next
    Set TxtFld = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Public Function CalcFieldsClear()
'Clear calculated fields
On Error GoTo ErrorHappened
    Dim X As Integer
    Dim RmCT As Integer
    Dim LastVisible As Integer

    'Application.Echo False
    'CREATE A DUMMY FIELD FOR FOCUS
    With FieldTextBox(FieldCount)
        .ColumnWidth = 100
        .ColumnHidden = False
        .Locked = False
        .Enabled = True
        .SetFocus 'SET FOCUS TO DUMMY FIELD
    End With
    For X = FieldCount To 1 Step -1
        With FieldTextBox(X)
            If LenB(Nz(.Tag, vbNullString)) > 0 Then
            
                .ControlSource = vbNullString
                .Width = 0
                .visible = False
                .ColumnWidth = 0
                .ColumnHidden = True
                .visible = False
                .Enabled = False
                .ColumnOrder = X  'Put it back in order
                With FieldLabel(X)
                    .Width = 0
                    .Caption = vbNullString
                End With
                RmCT = RmCT + 1
            End If
            If .ColumnWidth > 0 And .ColumnHidden = False Then 'GET THE LAST TO RESET FOCUS TO
                LastVisible = X
            End If
        End With
    Next X
    FieldTextBox(LastVisible).SetFocus 'REHIDE DUMMY COLUMN
    
    With FieldTextBox(FieldCount)
        .ColumnWidth = 0
        .ColumnHidden = True
        .Locked = True
        .Enabled = False
    End With
    
    'Set The Number Fields To Current Minus 1
    FieldCount = FieldCount - RmCT
    myForm.InsideWidth = myForm.Parent.InsideWidth
ExitNow:
On Error Resume Next
    'Application.Echo True
Exit Function
ErrorHappened:
    Resume ExitNow
    Resume
End Function

Public Sub LayoutField(ByVal Name As String, ByVal Ordinal As Long, ByVal Width As Single, ByVal CalcFld As Boolean)
'Apply field layout
On Error GoTo ErrorHappened
    Dim TxtFld As Access.TextBox
    Dim X As Long

    For X = 1 To FieldCount
        Set TxtFld = FieldTextBox(X)
        If (CalcFld And Nz(TxtFld.Tag, vbNullString) = Name) Or (Not CalcFld And TxtFld.ControlSource = Name) Then
        'If (CalcFld = True And "" & TxtFld.Tag = Name) Or (CalcFld = False And TxtFld.ControlSource = Name) Then
            With TxtFld
                .ColumnWidth = Width
                .ColumnOrder = Ordinal
            End With
        End If
    Next X
    
ExitNow:
On Error Resume Next
    Set TxtFld = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Public Sub LayoutClear()
'Clear layout
On Error GoTo ErrorHappened
    Dim TxtFld As Access.TextBox
    Dim X As Long
    
    'Application.Echo False
    For X = FieldCount To 1 Step -1
        Set TxtFld = FieldTextBox(X)
        With TxtFld
            .ColumnWidth = .Width
            .ColumnOrder = X
            If .Width <> 0 And .ColumnHidden = True Then
                .ColumnHidden = False
            End If
        End With
    Next X
    
    'Fix the Crazy Size Events
    DoEvents
    myForm.InsideWidth = myForm.Parent.InsideWidth

    'Application.Echo True
ExitNow:
On Error Resume Next
    Set TxtFld = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Public Sub AddField(ByVal FldNum As Integer, ByRef fld As CnlyFldDef)
'Add field
On Error GoTo ErrorHappened
    Dim genUtils As New CT_ClsGeneralUtilities
    Const TwipsPerInch As Integer = 1440

    With myForm(FieldTextBoxString(FldNum))
        If LenB(Nz(fld.ControlSrc, vbNullString)) > 0 Then 'CALC FIELD
            .ControlSource = "=(" & fld.ControlSrc & ")"
            .Tag = fld.Name
            .visible = True
            .Locked = False
            .ColumnHidden = False
            .ColumnWidth = -2
            .Enabled = True
        Else
            .ControlSource = fld.Name
            .Tag = vbNullString
            'DLC 06/09/2010 - Access 2010 Upgrade : Ensure that the field is displayed in the correct regional layout
            If fld.Type = 8 Then
                .Format = genUtils.GetRegionalShortDateFormat()
                .InputMask = .Format
            End If
            '-------------------------------------
        End If
        .Width = fld.Width * TwipsPerInch 'DOES NOT NEED TO BE FIXED AND IS USED FOR LAYOUTS
        .ColumnOrder = FldNum
        If fld.Width = 0 Then
            .ColumnHidden = True
            .ColumnWidth = 0
            .visible = False
            .Locked = True
            .Enabled = False
        Else
            .ColumnHidden = False
            .visible = True
            .Locked = False
            .Enabled = True
            .ColumnWidth = fld.Width * TwipsPerInch  'IN POINTS - NOW SURE WHY
        End If
        .Height = fld.Height * TwipsPerInch
        .TextAlign = fld.Align
        If fld.Format = "Hyperlink" Then
            .IsHyperlink = True
        Else
            .Format = fld.Format
        End If
        .DecimalPlaces = fld.Decimal
    End With

    With myForm(FieldLabelString(FldNum))
        If LenB(Nz(fld.Alias, vbNullString)) = 0 Then
            .Caption = fld.Name
        Else
            .Caption = fld.Alias
        End If
    End With
    FieldCount = FieldCount + 1

ExitNow:
On Error Resume Next
    Set genUtils = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Public Sub ResetMousePointer()
    If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Public Function GetDataSource(ByVal RecordSource As String) As String
'Get datasource for sheet
On Error GoTo ErrorHappened
    Dim Result As String
    Dim position As Integer
    Dim newString As String
    
    newString = UCase(RecordSource)
    
    position = InStr(RecordSource, " FROM ")
    If position > 0 Then
        newString = Mid$(RecordSource, position + 6)
        position = InStr(newString, " ")
        If position = 0 Then
            Result = newString
        ElseIf position < 0 Then
            Result = RecordSource
        Else
            Result = Mid$(newString, 1, position)
        End If
    Else
        Result = RecordSource
    End If
ExitNow:
On Error Resume Next
    GetDataSource = Result
Exit Function
ErrorHappened:
    Resume ExitNow
    Resume
End Function

Public Sub KeyPress(ByVal KeyAscii As Integer)
'Export to Excel on Control + E
On Error GoTo ErrorHappened
    If KeyAscii = 5 Then 'CTRL E
        Dim locExcel As New CT_ClsExcel
        If SelectionHeight > 0 Then
            With locExcel
                .ShowUI = False
                .ExportSelected myForm
            End With
        End If
        Set locExcel = Nothing
    End If
ExitNow:
On Error Resume Next
    Set locExcel = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Private Function FieldLabel(ByVal index As Integer) As Access.Label
    'Return label object with specifed index
    Set FieldLabel = myForm.Controls(FieldLabelString(index))
End Function

Private Function FieldLabelString(ByVal index As Integer) As String
    'Return a string with the name of the label based on index
    FieldLabelString = "CapField" & CStr(index)
End Function

Private Function FieldTextBox(ByVal index As Integer) As Access.TextBox
    'Return textbox object with specifed index
    Set FieldTextBox = myForm.Controls(FieldTextBoxString(index))
End Function

Private Function FieldTextBoxString(ByVal index As Integer) As String
    'Return a string with the name of the textbox based on index
    FieldTextBoxString = "Field" & CStr(index)
End Function
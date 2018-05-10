Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'SA 11/6/2012 - Changed to CT_SubGenericDataSheet050

Const MaxDynamicListSize = 50000 'used to "turn off" dynamic list filetering in response to user typing in Quick-Add area.
'                                   When the list is too large, the dynamic filtering becomes intolerably slow.

Public Results As Integer
Public MvSelections As Collection
Public WithEvents frmDataSheet As Form_CT_SubGenericDataSheet050
Attribute frmDataSheet.VB_VarHelpID = -1

Public maskedQuery As String
Public wrapQuery As String
Public ViewSource As String
Public orderVal As String
Public SqlWhere As String
Public searchFragment As String
Public lookupField As String
Public SortField As String
Public SortOrder As String
Public BuildWhere As String
Public buildOrder As String
Public BuildSQL As String
Public iAuditID As String

Private genUtils As New CT_ClsGeneralUtilities

Private mvTitle As String
Private MvListTitle As String
Private MvStartupWidth As Single
Private MvRecordCount As Long
Private MvScreenID As Long

Private Const DecipherMultiSelectRightClick As String = "DecipherMultiSelectRightClick"

Private rightMenu As CommandBar
Private WithEvents copyButton As CommandBarButton
Attribute copyButton.VB_VarHelpID = -1
Private WithEvents pasteButton As CommandBarButton
Attribute pasteButton.VB_VarHelpID = -1

Private mvFormId As Byte
Private MvListBy As Access.ComboBox
Private MvList As Access.ComboBox
Private MvCaption As Access.Label
Private MvListMulti As MSForms.listBox
Private MvQual As String
Private WithEvents TB As Access.TextBox
Attribute TB.VB_VarHelpID = -1
Private ActiveColumn As String
Private MvDataSource As String
Private MvMultiInfo As Access.Label

Const adVarChar = 200
Const adChar = 129

Private MouseX As Single
Private MouseY As Single

Private ListCol As Object
Property Let qualifier(data As String)
     MvQual = data
End Property
Property Let list(data As ComboBox)
     Set MvList = data
End Property
Property Let listBy(data As ComboBox)
    Set MvListBy = data
End Property
Property Let ListMulti(data As MSForms.listBox)
    Set MvListMulti = data
End Property
Property Let lblCaption(data As Access.Label)
    Set MvCaption = data
End Property
' property added to format grid
Property Let DataSource(Value As String)
    MvDataSource = Value
End Property
Property Let MultiInfo(Value As Access.Label)
    Set MvMultiInfo = Value
End Property
Public Property Let StartupWidth(data As Single)
On Error GoTo ErrorHappened

    ' AUTO CALCULATE THE WIDTH BASED ON COLUMN WIDTHS
    MvStartupWidth = data
    Me.InsideWidth = MvStartupWidth
    
ExitNow:
    On Error Resume Next
    Exit Property
ErrorHappened:
    MsgBox "Error Setting Startup For Width." & vbCrLf & vbCrLf & Err.Description, vbCritical, CodeContextObject.Name & ".StartupWidth"
    Resume ExitNow
End Property
Public Property Get Selections() As Collection
   Set Selections = MvSelections
End Property
Public Property Let Title(data As String)
    mvTitle = data
    Me.Caption = mvTitle
End Property
Public Property Get Title() As String
    Title = mvTitle
End Property
Public Property Let ListTitle(data As String)
    MvListTitle = data
    Me.LblLst.Caption = MvListTitle
End Property
Public Property Get ListTitle() As String
    ListTitle = MvListTitle
End Property
Property Get ScreenID() As Long
    ScreenID = MvScreenID
End Property
Public Sub Cancel()
    On Error Resume Next
    Results = vbCancel
    Me.visible = False
End Sub
Public Sub Ok()
On Error GoTo ErrorHappened

    Const MaxCaptionLength As Integer = 255
    'SA 1/18/2012 - CR1042 Variable used to set label
    Dim FrmScreen As Form_SCR_MainScreens
    Dim X As Long
    Dim strCaption As String
    Dim BlnTruncated As Boolean
    
    DoCmd.Hourglass True
    Set FrmScreen = Scr(mvFormId)
    
    With MvListMulti
        .Clear
        
        For X = 0 To CmboMulti.ListCount - 1
            .AddItem CmboMulti.list(X)

            If Not BlnTruncated Then
                If Len(strCaption) > MaxCaptionLength Then
                    BlnTruncated = True
                Else
                    strCaption = strCaption & CmboMulti.list(X) & ", "
                End If
            End If
        Next X
        
        'Remove the trailing comma
        If CmboMulti.ListCount > 0 Then
            strCaption = left(strCaption, Len(strCaption) - 2)
        End If

        If BlnTruncated Or Len(strCaption) > MaxCaptionLength Then
            strCaption = left(strCaption, MaxCaptionLength - 3) & "..."
        End If
    
        MvMultiInfo.ControlTipText = strCaption
        
        If Len(strCaption) > 21 Then
            strCaption = "Items selected: " & CmboMulti.ListCount
        End If
        
        If CmboMulti.ListCount > 0 Then
            MvMultiInfo.Caption = strCaption
        Else
            MvMultiInfo.Caption = "(No Items Selected)"
            MvMultiInfo.ControlTipText = vbNullString
        End If
            
    End With
    
ExitNow:
    On Error Resume Next
    DoCmd.Hourglass False
    Results = vbOK
    Me.visible = False
    
    Exit Sub
ErrorHappened:
    'Display error in label
    MvMultiInfo.ControlTipText = vbNullString
    MvMultiInfo.Caption = "(Error)"
    
    Resume ExitNow
    Resume
End Sub
Private Sub CmboMulti_DblClick(ByVal Cancel As Object)
    cmdRemove_Click
End Sub
Private Sub CmboMulti_KeyDown(ByVal KeyCode As Object, ByVal Shift As Integer)
    'Debug.Print KeyCode, Shift
    If KeyCode = 86 And Shift = 2 Then 'Ctl V' was pressed, do a paste
        CmdPaste_Click
    End If
    
    If KeyCode = 67 And Shift = 2 Then 'Ctl c' was pressed, do a copy
        CopyList
    End If
End Sub
Private Sub CmboMulti_KeyPress(ByVal KeyAscii As Object)
    'Debug.Print KeyAscii
End Sub
Private Sub CmdAdd_Click()
    Dim i As Integer
    
    On Error GoTo eTrap
    
    DoCmd.Hourglass True

    With frmDataSheet
        If .SelectionHeight = 0 Or .SelectionHeight = 1 Then
            CmboMultiAddItem .RecordSet.Fields(MvList.BoundColumn - 1)
        Else
            If .SelectionHeight > 10000 Then
                If MsgBox("You are about to add " & .SelectionHeight & " items.", vbOKCancel, "Add Selected Items") = vbCancel Then
                    GoTo eSuccess
                End If
            End If
            
            'need to iterate from select top to selectheight using a recordset clone.
            With Me.frmDataSheet.RecordsetClone
                .MoveFirst
                .Move Me.frmDataSheet.SelectionTop - 1
                For i = 1 To Me.frmDataSheet.SelectionHeight
                    CmboMultiAddItem .Fields(MvList.BoundColumn - 1)
                    .MoveNext
                Next i
            End With
        End If
    End With

eSuccess:
    On Error Resume Next
    DoCmd.Hourglass False
    Exit Sub

eTrap:
    MsgBox Err.Description
    Resume eSuccess
    
End Sub
Private Sub CmdAddAll_Click()
    DoCmd.Hourglass True
    With frmDataSheet
            'need to iterate from select top to selectheight using a recordset clone.
            With Me.frmDataSheet.RecordsetClone
                .MoveFirst
                Do Until .EOF
                    CmboMultiAddItem .Fields(MvList.BoundColumn - 1)
                    .MoveNext
                Loop
            End With
    End With
    DoCmd.Hourglass False
End Sub
Private Sub CmdCancel_Click()
    Cancel
End Sub
Private Sub cmdClear_Click()
    DoCmd.Hourglass True
    CmboMultiClear
    DoCmd.Hourglass False
End Sub
Private Sub cmdKeyA_Click()
    reFilter ("A")
End Sub
Private Sub cmdKeyB_Click()
    reFilter ("B")
End Sub
Private Sub cmdKeyC_Click()
    reFilter ("C")
End Sub
Private Sub cmdKeyD_Click()
    reFilter ("D")
End Sub
Private Sub cmdKeyE_Click()
    reFilter ("E")
End Sub
Private Sub cmdKeyF_Click()
    reFilter ("F")
End Sub
Private Sub cmdKeyG_Click()
    reFilter ("G")
End Sub
Private Sub cmdKeyH_Click()
    reFilter ("H")
End Sub
Private Sub cmdKeyI_Click()
    reFilter ("I")
End Sub
Private Sub cmdKeyJ_Click()
    reFilter ("J")
End Sub
Private Sub cmdKeyK_Click()
    reFilter ("K")
End Sub
Private Sub cmdKeyL_Click()
    reFilter ("L")
End Sub
Private Sub cmdKeyM_Click()
    reFilter ("M")
End Sub
Private Sub cmdKeyN_Click()
    reFilter ("N")
End Sub
Private Sub cmdKeyO_Click()
    reFilter ("O")
End Sub
Private Sub cmdKeyP_Click()
    reFilter ("P")
End Sub
Private Sub cmdKeyQ_Click()
    reFilter ("Q")
End Sub
Private Sub cmdKeyR_Click()
    reFilter ("R")
End Sub
Private Sub cmdKeyS_Click()
    reFilter ("S")
End Sub
Private Sub cmdKeyT_Click()
    reFilter ("T")
End Sub
Private Sub cmdKeyU_Click()
    reFilter ("U")
End Sub
Private Sub cmdKeyV_Click()
    reFilter ("V")
End Sub
Private Sub cmdKeyW_Click()
    reFilter ("W")
End Sub
Private Sub cmdKeyX_Click()
    reFilter ("X")
End Sub
Private Sub cmdKeyY_Click()
    reFilter ("Y")
End Sub
Private Sub cmdKeyZ_Click()
    reFilter ("Z")
End Sub
Private Sub cmdNbr0_Click()
    reFilter ("0")
End Sub
Private Sub cmdNbr1_Click()
    reFilter ("1")
End Sub
Private Sub cmdNbr2_Click()
    reFilter ("2")
End Sub
Private Sub cmdNbr3_Click()
    reFilter ("3")
End Sub
Private Sub cmdNbr4_Click()
    reFilter ("4")
End Sub
Private Sub cmdNbr5_Click()
    reFilter ("5")
End Sub
Private Sub cmdNbr6_Click()
    reFilter ("6")
End Sub
Private Sub cmdNbr7_Click()
    reFilter ("7")
End Sub
Private Sub cmdNbr8_Click()
    reFilter ("8")
End Sub
Private Sub cmdNbr9_Click()
    reFilter ("9")
End Sub
Private Sub cmdOk_Click()
    Ok
End Sub

Private Function BuildQuery(fragment As String) As String
'SA 11/26/2012 - Removed unused parameter "ranged"
Dim SQL As String
Dim SqlBase As String

Dim Where As String
Dim sqlOrderBy As String
Dim WhereInstr As Integer
Dim OrderByInstr As Integer
Dim FieldIsString As Boolean

    Where = ""
    
    SQL = MvList.RowSource
    ' HC corrected search to include the spaces
    WhereInstr = InStr(1, SQL, " where ", vbTextCompare)
    OrderByInstr = InStr(1, SQL, " order by ", vbTextCompare)
    
    If WhereInstr > 0 Then
        SqlBase = Mid$(SQL, 1, WhereInstr - 1)
    ElseIf OrderByInstr > 0 Then
        SqlBase = Mid$(SQL, 1, OrderByInstr - 1)
    Else
        SqlBase = SQL
    End If
    
    
    If WhereInstr > 0 Then
        Where = Mid$(SQL, WhereInstr, IIf(OrderByInstr > 0, OrderByInstr, Len(SQL)) - WhereInstr)
    Else
        Where = ""
    End If
    
    If OrderByInstr > 0 Then
        sqlOrderBy = Mid$(SQL, OrderByInstr, Len(SQL) - OrderByInstr + 1)
    Else
        sqlOrderBy = ""
    End If

    If (fragment <> "") Then
        Select Case frmDataSheet.RecordSet.Fields(ActiveColumn).Type
            Case adBSTR, adChar, adVarChar, adWChar, _
               adVarWChar, adLongVarChar, adLongVarWChar
                FieldIsString = True
            Case Else
                FieldIsString = False
        End Select


        If Where <> "" Then
            Where = Where & " AND "
        
        Else
            Where = " WHERE "
        End If
        
        'The "filter" is different if the field is a string or a number.
        If FieldIsString Then
            Where = Where & ActiveColumn & " >= '" & fragment & "' AND " & ActiveColumn & " < '" & fragment & "zzzzzzzzzz' "
        Else
            Where = Where & ActiveColumn & " Like '" & fragment & "*' "
        End If
    End If
    
    BuildSQL = SqlBase
    BuildWhere = Where
    buildOrder = sqlOrderBy
    BuildQuery = SqlBase & Where & sqlOrderBy 'FIX no range where clause RRS
End Function

Private Sub reFilter(fragment As String)
    Dim theQuery As String
    ActiveColumn = GetActiveColumn
    'clearGrid
    theQuery = BuildQuery(fragment)
    
    frmDataSheet.RecordSource = theQuery
End Sub
Private Sub CmdPaste_Click()
    'Grab and format contents of clipboard and, if possible, paste it in to the CmboMulti list.
    Dim strBuf As String

    On Error Resume Next
    
    DoCmd.Hourglass True
    strBuf = ClipBoard_GetData()
    If Trim(strBuf) <> "" Then
        PasteList strBuf
    End If
    DoCmd.Hourglass False
End Sub
Private Sub cmdQuickAdd_Click()
    'add the item to the list and clear the filter.
    Dim i As Long
    Dim fldType As Integer
    Dim strDelimit As String
    
    
    On Error GoTo eTrap
    
    If Trim(Me.txtQuickAdd) <> "" Then
        'First, is this a wildcard search or literal search?
        If InStr(1, Me.txtQuickAdd, "*") > 0 Or InStr(1, Me.txtQuickAdd, "?") > 0 Then
            'Wildcard match
            'if > 10000 items then list was not filtered yet. Do it now to be sure.
            reFilter Me.txtQuickAdd.Value
            'Add anything in the grid that matches the quick-add text to the select list.
            CmdAddAll_Click
            Me.txtQuickAdd.Value = ""
        Else
            'Literal Match
            'Check to see if we are supposed to validate the value first.
            If chkValidateQuickAdd <> 0 Then
                If BuildWhere <> "" Then reFilter ""
                'check to see if the quick add value is in the list of items to choose from
                ' HC v2.5.1200, the quick add value is always based on the sorted column, so this needs to be the mvlistby value
                fldType = Me.frmDataSheet.RecordSet.Fields(MvListBy.BoundColumn - 1).Type
                If fldType = 10 Or fldType = 12 Then
                    'Is an alpha field, delimit with quotes.
                    strDelimit = """"
                Else
                    'is a numeric field, no delimiter.
                    strDelimit = ""
                End If
                
                ' HC v 2.5.1200, changed mvList to mvlistby so the collect column name is included in the sql stmt.
                i = GetRowCountFromSql(BuildSQL & " where " & Me.frmDataSheet.RecordSet.Fields(MvListBy.BoundColumn - 1).Name & " = " & strDelimit & Me.txtQuickAdd & strDelimit)
                If i > 0 Then 'Matches an item in the pick list.
                    CmboMultiAddItem Me.txtQuickAdd
                    Me.txtQuickAdd.Value = ""
                Else
                    MsgBox "Your Quick Add value is not in the list." & vbCrLf & "Please enter a new value.", vbExclamation, "Quick Add Value Warning"
                End If
            Else
                CmboMultiAddItem Me.txtQuickAdd
                Me.txtQuickAdd.Value = ""
            End If
        End If
    End If
    Me.txtQuickAdd.SetFocus
Exit Sub
eTrap:
    MsgBox Err.Description, vbInformation, "Quick Add Error"
End Sub

Private Sub cmdRefresh_Click()
    Dim theQuery As String
    
    theQuery = MvList.RowSource
    frmDataSheet.InitData theQuery, 2, MvDataSource
    frmDataSheet.RecordSource = theQuery
    FormatGrid
End Sub
Private Sub cmdSearchPhrase_Click()
    searchFragment = Replace(Replace(txtSearchPhrase.Value & "", "'", ""), Chr(34), "")
    reFilter (searchFragment)
End Sub
Private Sub CmdSQL_Click()
On Error Resume Next

    DoCmd.OpenForm CCAText, , , , , , Scr(mvFormId).BuildMultiItemSQL(Me.CmboMulti.Object, MvQual)

    If Not TB Is Nothing Then
    Set TB = Nothing
    End If
    
    Set TB = Forms(CCAText).TxtText
    TB.OnExit = "[Event Procedure]"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    'The following defines the behavior of the enter key.
    
    If KeyAscii = 13 Then
        'If the user has just updated the filter phrase and hits <Enter>, apply the filter.
        If Me.ActiveControl.Name = Me.cmdSearchPhrase.Name Then
            cmdSearchPhrase_Click
        End If
        
        If Me.ActiveControl.Name = Me.cmdQuickAdd.Name Then
            cmdQuickAdd_Click
        End If
    End If
End Sub
Private Sub CmboMulti_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Button = 2 Then
        'Set rightMenu = CommandBars.FindControl(Type:=msoBarPopup, Name:=DecipherMultiSelectRightClick)
        If Not (rightMenu Is Nothing) Then
            If CmboMulti.ListCount > 0 Then
                copyButton.Enabled = True
            Else
                copyButton.Enabled = False
            End If
            pasteButton.Enabled = True
            rightMenu.ShowPopup
        End If
    End If
        
    
End Sub

Private Sub Form_Load()
       
    Set frmDataSheet = Lst.Form
   
    genUtils.SuspendLayout Me
    
    If (ListCol Is Nothing) Then
        Set ListCol = CreateObject("Scripting.Dictionary")
    End If
    
    Me.txtQuickAdd.SetFocus
    MakeRightClickMenu
    genUtils.ResumeLayout Me
    
        
End Sub
Private Sub Form_Resize()
On Error Resume Next
Dim SgWidth As Single
Dim sgHeight As Single
Dim SgLeft As Single

    If Me.InsideHeight < 5205 Then Me.InsideHeight = 5205
    If Me.InsideWidth < 12870 Then Me.InsideWidth = 12870

    SgWidth = Me.InsideWidth
    sgHeight = Me.InsideHeight '  Me.Section(acDetail).Height
    Me.Section(acDetail).Height = sgHeight
    Me.LblLst.Width = SgWidth - (SgLeft * 2)
    Me.Repaint
    
    Me.Lst.Height = Me.Detail.Height - Me.FormHeader.Height - Me.FormFooter.Height - 90
    Me.CmboMulti.Height = Me.Lst.Height
    Me.CmdOK.left = Me.InsideWidth - 45 - Me.CmdCancel.Width
    
    With CmdCancel
        .left = CmdOK.left - (SgLeft * 2) - .Width
        .top = CmdOK.top
    End With
       
    Me.CmboMulti.left = Me.InsideWidth - 45 - Me.CmboMulti.Width
    Me.CmdAdd.left = Me.CmboMulti.left - 420 '540
    Me.CmdAddAll.left = Me.CmboMulti.left - 420 '540
    Me.CmdClear.left = Me.CmboMulti.left - 420 '540
    Me.CmdPaste.left = Me.CmboMulti.left - 420 '540
    Me.CmdRemove.left = Me.CmboMulti.left - 420 '540
    Me.CmdSQL.left = Me.CmboMulti.left - 420 '540
     
    Me.Lst.Width = Me.CmboMulti.left - 555 - Me.Lst.left
    Me.lblSelectedItemsPanel.left = Me.CmboMulti.left
    Me.lblSelectedItems.left = Me.CmboMulti.left
    Me.pnlToolbar.left = Lst.Width + 100
    Me.pnlToolbar.Height = Lst.Height
    Me.lineDetail.top = Lst.Height + 50
End Sub
Public Sub InitData(FormID As Byte, ByVal lvl As Byte)
    Dim FrmScreen As Form_SCR_MainScreens
    Dim X As Long
        
    'Set the Module Level Variable For the Calling Screen
    mvFormId = FormID
    Set FrmScreen = Scr(FormID)
    MvScreenID = FrmScreen.ScreenID
    
    
    With FrmScreen
        Select Case lvl
            Case 1
                Me.list = .CmboPrimary
                Me.listBy = .CmboListPrimaryBy
                Me.ListMulti = .CmboPrimaryMulti.Object
                Me.lblCaption = .LblPrimary
                Me.qualifier = .Config.PrimaryQualifier
                Me.DataSource = .Config.PrimaryListBoxRecordSource
                Me.MultiInfo = .lblPrimaryMulti
            Case 2
                Me.list = .CmboSecondary
                Me.listBy = .CmboListSecondaryBy
                Me.ListMulti = .CmboSecondaryMulti.Object
                Me.lblCaption = .LblSecondary
                Me.qualifier = .Config.SecondaryQualifier
                Me.DataSource = .Config.SecondaryListBoxRecordSource
                Me.MultiInfo = .lblSecondaryMulti
            Case 3
                Me.list = .CmboTertiary
                Me.listBy = .CmboListTertiaryBy
                Me.ListMulti = .CmboTertiaryMulti.Object
                Me.lblCaption = .LblTertiary
                Me.qualifier = .Config.TertiaryQualifier
                Me.DataSource = .Config.TertiaryListBoxRecordSource
                Me.MultiInfo = .lblTertiaryMulti
        End Select
    End With
    
    SortField = MvListBy
    
    reFilter ""
    
    MvRecordCount = GetRowCountFromSql(MvList.RowSource)
    
    If MvRecordCount > MaxDynamicListSize Then  ' HC 10/17/2008 - to use constant as in original
        chkValidateQuickAdd = 0
    Else
        chkValidateQuickAdd = -1
    End If
    
    'Me.Lbl.Caption = MvCaption.Caption
    
    'SYNC THE CURRENTLY SELECTED ITEMS
    With MvListMulti
        'Sync the Items in the lists
        For X = 0 To .ListCount - 1
            CmboMultiAddItem .list(X)
        Next X
    End With
    
    frmDataSheet.InitData MvList.RowSource, 2, MvDataSource
    FormatGrid
End Sub
Private Sub frmDataSheet_DblClick()
    ' 4/20/2009 DEB: Added fix bug that caused all items to be added to the select list when the user intends
    ' to only resize the columns by double clicking on the header
    ' Only add the selection if the user clicked on a row not that header
    ' if the mouse is "below" the height of one of the caption fields, we're on a row
    ' The 1.09 factor is just to provide a little insurance that we are really not on the header row
    If MouseY > (frmDataSheet.CapField1.Height * 1.09) Then
        CmdAdd_Click
    End If
End Sub
Private Sub cmdRemove_Click()
    Dim N As Long

    DoCmd.Hourglass True
    
    With CmboMulti
        For N = .ListCount - 1 To 0 Step -1
            'If Selected Then Remove it
            If .Selected(N) Then CmboMultiRemoveItem N
        Next N
    End With
    

    DoCmd.Hourglass False
End Sub

Public Sub PasteList(strBuf)
    'Paste contents of string buffer "strBuf" into the selected items list box.
    'This will typically be the contents of the clipboard cut from an excel spreadsheet, etc.
    'The list may either be comma or crlf separated.
    
    Dim staBuf As Variant
    Dim i As Long
    
   
    If InStr(1, strBuf, ",") Then 'is comma separated
        'strip out any un-needed formatting.
        strBuf = Replace$(strBuf, Chr$(34), "")
        strBuf = Replace$(strBuf, "(", "")
        strBuf = Replace$(strBuf, ")", "")
        
        staBuf = Split(strBuf, ",")
    ElseIf InStr(1, strBuf, vbCrLf) Then 'is crlf separated list
        staBuf = Split(strBuf, vbCrLf)
    End If
    
    If IsArray(staBuf) Then
        For i = 0 To UBound(staBuf)
            If staBuf(i) <> "" Then CmboMultiAddItem Trim(staBuf(i))
        Next i
    End If

End Sub

Private Sub frmDataSheet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseX = X
MouseY = Y
End Sub

Private Sub tb_Exit(Cancel As Integer)
'used to fish data from a form opened as a dialog box.  This event will fire when the dialog form is closing
'allowing us to get the value.  "tb" was SET to reference a text box on the dialog form.

    Dim strReturn As String
    Dim X As Integer
    Dim strUpdatedList As Variant
    Dim strItem As String
    
    strReturn = "" & TB.Value
    
    Set TB = Nothing
    
    'Did the user change the list?
    If strReturn <> Scr(mvFormId).BuildMultiItemSQL(Me.CmboMulti.Object, MvQual) Then
        'Yes update the form list.
        DoCmd.Hourglass True
        
        'Clear the list box
        CmboMultiClear
        
        'trim off the enclosing "()" on the string.
        strReturn = left$(strReturn, Len(strReturn) - 1)
        strReturn = Right$(strReturn, Len(strReturn) - 1)
        strReturn = Replace$(strReturn, Chr$(34), "")
        
        strUpdatedList = Split(strReturn, ",")
        
        For X = 0 To UBound(strUpdatedList)
            strItem = strUpdatedList(X)
            If strItem <> "" Then
                CmboMultiAddItem strItem
            End If
        Next X
        DoCmd.Hourglass False
    Else
    'do nothing
    End If
End Sub
Private Sub txtQuickAdd_Change()
Beep
End Sub
Private Sub txtQuickAdd_KeyUp(KeyCode As Integer, Shift As Integer)
    If MvRecordCount < MaxDynamicListSize And chkValidateQuickAdd <> 0 Then
        reFilter Me.txtQuickAdd.Text
    End If
End Sub
Private Sub txtSearchPhrase_AfterUpdate()
    'Beep
End Sub
Public Sub CopyList()
'Copy contents of cmboMulti to the clip board as a CRLF separated value list.
    Dim strClip As String
    Dim X As Long
    
    On Error GoTo eTrap
    
    DoCmd.Hourglass True
        
    'Build the list string from the cmboMuli list contents
    For X = 0 To CmboMulti.ListCount - 1
        strClip = strClip & CmboMulti.list(X) & vbCrLf
    Next X

    DoCmd.Hourglass False
    
    If Trim$(strClip) <> "" Then
        'put it in the Clipboard
        ClipBoard_SetData strClip
    End If
    
eSuccess:
    On Error Resume Next
    DoCmd.Hourglass False
    
    Exit Sub
    
eTrap:
Resume eSuccess

End Sub
Public Function GetRowCountFromSql(SQL As String) As Long

    Dim SqlBase As String
    Dim Where As String
    Dim WhereInstr As Integer
    Dim OrderByInstr As Integer
    Dim BaseTable As String
    
    On Error GoTo eTrap
    
    'Sql = MvList.RowSource
    ' HC corrected search to include the spaces
    WhereInstr = InStr(1, SQL, " where ", vbTextCompare)
    OrderByInstr = InStr(1, SQL, " order by ", vbTextCompare)
    
    If WhereInstr > 0 Then
        SqlBase = Mid$(SQL, 1, WhereInstr - 1)
    ElseIf OrderByInstr > 0 Then
        SqlBase = Mid$(SQL, 1, OrderByInstr - 1)
    Else
        SqlBase = SQL
    End If
    
    BaseTable = Trim(Split(SqlBase, " from ", , vbTextCompare)(1))
    
    If WhereInstr > 0 Then
        'add 5 to starting point becuase we don't want the word "where" in the statement.
        Where = Mid$(SQL, WhereInstr + 6, IIf(OrderByInstr > 0, OrderByInstr, Len(SQL)) - WhereInstr - 5)
    Else
        Where = ""
    End If
        
    GetRowCountFromSql = DSum(1, BaseTable, Where)
    
ExitSuccess:
    On Error Resume Next
    Exit Function
    
eTrap:
    GetRowCountFromSql = 0
    
    Resume ExitSuccess
    Resume

End Function
Private Function GetActiveColumn() As String
    'derive the name of the field associated with active control on the data grid.
    On Error GoTo eTrap
    
    Dim strActiveControlName As String
    
    strActiveControlName = Me.frmDataSheet.ActiveControl.ControlSource
    
    If strActiveControlName = "" Then
        strActiveControlName = SortField
    Else
        SortField = strActiveControlName

    End If
    
    GetActiveColumn = strActiveControlName
    
eSuccess:
    Exit Function
    
eTrap:
    GetActiveColumn = SortField

End Function

Private Sub CmboMultiAddItem(sItem As String)
    'Add the item to the list box if it is not already present.
    'Use a collection "ListCol" to keep track of list box contents as it is much faster for a large list.
    On Error GoTo eTrap
    
    If ListCol.Exists(sItem) = False Then
        CmboMulti.AddItem sItem
        ListCol.Add sItem, sItem
    End If
eSuccess:
    Exit Sub
    
eTrap:
    MsgBox "Item could not be added to the select list." & vbCrLf & Err.Description, vbExclamation
End Sub
Private Sub CmboMultiRemoveItem(index As Variant)
    'Remove item from list box by Index reference.
    'Maintain ListCol collection to keep it in sync.
    
    On Error GoTo eTrap
    
        ListCol.Remove CmboMulti.list(index)
        CmboMulti.RemoveItem (index)
        
eSuccess:
    Exit Sub
eTrap:
    MsgBox "Item could not be removed to the select list." & vbCrLf & Err.Description, vbExclamation
End Sub
Private Sub CmboMultiClear()
    CmboMulti.Clear
    ListCol.RemoveAll
End Sub
Private Sub copyButton_Click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    CopyList
End Sub

Private Sub pasteButton_Click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    CmdPaste_Click
End Sub

Private Sub MakeRightClickMenu()
On Error GoTo ErrorHappened
    
    Dim objCommandBar As CommandBar
    Dim objCommandBarButton As CommandBarButton
    Dim genUtils As New CT_ClsGeneralUtilities
            
    ' clear the existing DecipherMultiSelectRightClick
    genUtils.ClearMenu (DecipherMultiSelectRightClick)
    Set rightMenu = Nothing
    
    Set objCommandBar = CommandBars.Add(Name:=DecipherMultiSelectRightClick, position:=msoBarPopup, Temporary:=False, MenuBar:=False)
    Set rightMenu = objCommandBar
    
    Set objCommandBarButton = objCommandBar.Controls.Add(msoControlButton, , , , False)
    With objCommandBarButton
        .Caption = "Copy"
        .Tag = "Copy"
        .FaceId = 19
        .style = msoButtonIconAndCaption
    End With
    Set copyButton = objCommandBarButton
    Set objCommandBarButton = objCommandBar.Controls.Add(msoControlButton, , , , False)
    With objCommandBarButton
        .Caption = "Paste"
        .Tag = "Paste"
        .FaceId = 22
        .style = msoButtonIconAndCaption

    End With
    Set pasteButton = objCommandBarButton
ErrorHappened:
    Exit Sub
End Sub
Private Sub FormatGrid()

    Title = MvCaption.Caption
    ListTitle = SortOrder       '"Select Vendor:"
    StartupWidth = Me.CmboMulti.left + Me.CmboMulti.Width + 30 'AUTO SIZE THE FORM TO LIST WIDTH

End Sub

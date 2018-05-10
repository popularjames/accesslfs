Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Form to capture a history of Ad-Hoc Filters and allow user to move through/edit/apply them.
'David.Brady 01.08.2009
'Updated and reviewed by Development 11/04/2010

Private WithEvents MvGridMain As Form_CT_SubGenericDataSheet 'to get screen events.
Attribute MvGridMain.VB_VarHelpID = -1
Private WithEvents MvMainScreen As Form_SCR_MainScreens 'to get grid events.
Attribute MvMainScreen.VB_VarHelpID = -1

Private Const Dictionary As String = "Scripting.Dictionary"
Private Const FiltersCleared As String = "- Clear All Filters -" 'displayed when all filters have been cleared

Private cListItems As Object
Private cScreens As Object
Private isUnregistered As Boolean

Public Property Get screenCount() As Integer
    screenCount = cScreens.Count
End Property

'It binds the Filter History form to the active ScrMainScreen form
Public Function SetActiveScreen(FormID As Byte)
    '021809 DB: Fixed bug that resulted in FiltersCleared being added to the histroy repeatedly.
    '021809 DB: Form reved to 1.1
    On Error Resume Next
    Dim strFilter As String
        
    If MvMainScreen Is Nothing Then
        InitScreen FormID
        Set MvMainScreen = Scr(FormID)
        Set MvGridMain = MvMainScreen.GridForm
        Me.Caption = MvMainScreen.Caption
    
    Else
        If MvMainScreen.FormID <> FormID Then
            InitScreen FormID
            Set MvMainScreen = Scr(FormID)
            Set MvGridMain = MvMainScreen.GridForm
            Me.Caption = MvMainScreen.Caption
            
            If MvGridMain.FilterOn = True Then
                strFilter = MvGridMain.filter
            Else
                strFilter = ""
            End If
            
            'If the grid's current filter is in the list, select it
            If (strFilter <> "") Or (CmboMulti.ListCount > 0) Then
                strFilter = IIf(strFilter = "", FiltersCleared, strFilter)
                'remove the ScrSubGenericDataSheet reference if any to avoid duplicates
                strFilter = Replace(strFilter, "[ScrSubGenericDataSheet].", "")
                If FilterExistsInList(strFilter) = False Then
                    'add it to the list
                    CmboMulti.AddItem strFilter
                    CmboMulti.ListIndex = CmboMulti.ListCount - 1
                End If
            End If
        End If
    End If
End Function

Private Function InitScreen(FormID As Byte) As Boolean
    'If this is a new screen, add it, else grab the existing one.
    Dim filterList As Object 'dictionary that holds the list of filters for a screen
    
    If cScreens Is Nothing Then
        Set cScreens = CreateObject(Dictionary)
    End If
    
    If cScreens.Exists(CStr(FormID)) = False Then
        Set cListItems = CreateObject(Dictionary)
        cScreens.Add CStr(FormID), cListItems
        Set cListItems = Nothing
    End If
    
    'If this is the the screen already in use, do nothing, else, save existing values and load values for this screen.
    If MvMainScreen Is Nothing Then 'ensure there is a curren screen.
        GetList cScreens.Item(CStr(FormID)), CmboMulti
    Else
        'NOTE: code failing here.  Scr(FormID) not initialized for the active screen sometimes.
        'did not see this happending JL 04/11/2011
        If Nz(MvMainScreen.Caption, "") <> Scr(FormID).Caption Then
           'save the history for the old screen if exist
           Set filterList = cScreens.Item(CStr(MvMainScreen.FormID))
           PutList filterList, CmboMulti
            'load the history for this screen if any
            GetList cScreens.Item(CStr(FormID)), CmboMulti
            
            Me.Caption = Scr(FormID).Caption
        End If

    End If
End Function

Private Sub CmboMulti_Click()
    Me.txtEdit = CmboMulti.list(CmboMulti.ListIndex)
End Sub

Private Sub CmboMulti_DblClick(ByVal Cancel As Object)
On Error GoTo CmboMultiError
    Dim filter As String
    
    If CmboMulti.ListCount > 0 Then
    
        filter = CmboMulti.list(CmboMulti.ListIndex)
        filter = IIf(filter = FiltersCleared, "", filter)
        MvGridMain.filter = filter
        
        MvGridMain.FilterOn = True
        
        MvGridMain.ApplyFilter 0, 1
    End If

CmboMultiExit:
    
    Exit Sub
CmboMultiError:
    MsgBox "The specified filter has an error in it and could not be applied." & vbCrLf & "Please revise the filter before re-applying it.", vbExclamation
    Resume CmboMultiExit
End Sub

Private Sub CmboMulti_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim filter As String
    If CmboMulti.ListCount > 0 Then
        filter = CmboMulti.list(CmboMulti.ListIndex)
        filter = IIf(filter = FiltersCleared, "", filter)
        Me.txtEdit = filter
    End If

End Sub

Private Sub cmdApply_Click()
    On Error GoTo cmdApplyError
    Dim filter As String
    filter = Me.txtEdit.Value
    filter = IIf(filter = FiltersCleared, "", filter)
      
    MvGridMain.filter = filter
    MvGridMain.FilterOn = True
    MvGridMain.ApplyFilter 0, 1
    
cmdApplyExit:
     Exit Sub
cmdApplyError:
    MsgBox "The specified filter has an error in it and could not be applied." & vbCrLf & "Please revise the filter before re-applying it.", vbExclamation
    Resume cmdApplyExit
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.InsideHeight = 6000
    Me.InsideWidth = 4320
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cScreens = Nothing
End Sub


'It caputeres Apply Filter Event from the SubGenericDataSheet form
Private Sub MvGridMain_ApplyFilter(filter As String)
    Dim FilterCopy As String
    'For the sake of clairity, change an empty filter into the words FiltersCleared = "- Clear All Filters -"
    'to keep blank lines from appearing in the list.
    
    FilterCopy = IIf(filter = "", FiltersCleared, filter)

    If FilterExistsInList(FilterCopy) = False Then
        CmboMulti.AddItem FilterCopy
        CmboMulti.ListIndex = CmboMulti.ListCount - 1
    End If
End Sub

Private Sub cmdClearCriteria_Click()
    DoCmd.Hourglass True
    CmboMulti.Clear
    DoCmd.Hourglass False
End Sub

Private Sub cmdOk_Click()
    'hide form
    Me.visible = False
End Sub


'It allows the user to hear a sound when the Form Filter list is updated
Private Sub CmboMulti_Updated(Code As Integer)
    'leave for easy implementation when required
    'Beep
End Sub

'Captures the isVisible Event from ScrMainScreens to determine what screen is visible
'and adjust the History Information accordantly
Private Sub MvMainScreen_isVisible(data As Boolean)
    DoEvents
    On Error GoTo MvMainScreen_isVisibleError
    'leave for testing
    'Debug.Print MvMainScreen.Caption & " visible = " & Data
    'Debug.Print "Active form is now" & Screen.ActiveForm.Caption

    'make it the active form.
    Dim FormID As Byte

    'Get the formid for the active form if any by comparing the caption of the active form
    'with against the name of enumerated screens in the Scr() array.
    If data = False Then 'then curren screen just lost focus, attach to the new screen.
       FormID = GetActiveFormID()
       'If a matching screen was found, bind to it.
       If FormID > 0 Then
            SetActiveScreen FormID
       End If
    End If

MvMainScreen_isVisibleExit:
    Exit Sub
MvMainScreen_isVisibleError:
    Me.visible = False
    Resume MvMainScreen_isVisibleExit
End Sub

'Removes specified screen from collection.
Public Function UnregisterScreen(FormID As Byte)
   
    'if regestered screen count = 0 then close form.
    If cScreens.Exists(CStr(FormID)) Then
        cScreens.Remove (CStr(FormID))
    End If
    isUnregistered = True
    If cScreens.Count = 0 Then
        Me.visible = False
    End If
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo FormKeyUpError

    Dim itemsCount As Integer
    Dim idx As Integer
    
    If KeyCode = 65 And Shift = 2 Then 'Control and A key pressed select all values
        With CmboMulti
            itemsCount = .ListCount - 1
            
            For idx = 0 To itemsCount
                .Selected(idx) = True
            Next
        End With
        CmboMulti.SetFocus
    End If
        
    If Me.ActiveControl Is Me.CmboMulti And KeyCode = 46 And Shift = 0 Then 'Delete key pressed remove selected value
        With CmboMulti
            itemsCount = .ListCount - 1
            
            For idx = itemsCount To 0 Step -1
                If .Selected(idx) Then
                    .RemoveItem idx
                End If
            Next
        End With
    End If
    
    If KeyCode = 27 Then 'Escape key pressed hide form
        Me.visible = False
    End If
    

FormKeyUpExit:
    Exit Sub
    
FormKeyUpError:
    MsgBox "An error occurred. " & Err.Description
    Resume FormKeyUpExit
End Sub


Private Sub Form_Load()
   
    Set cScreens = CreateObject(Dictionary)
End Sub

Private Function FilterExistsInList(StNewData As String) As Boolean
    Dim N As Integer, ItemFound As Boolean
    
    For N = 0 To CmboMulti.ListCount - 1
        If CmboMulti.list(N) = StNewData Then
            ItemFound = True
            CmboMulti.ListIndex = N 'to show that this is the "active" filter.
            Exit For
        End If
    Next N
    FilterExistsInList = ItemFound

End Function

Private Sub PutList(dicFilterList As Variant, cmboList As Object)
   'populate collection sc from listbox lsb.
   Dim i As Long
   
   dicFilterList.RemoveAll
   For i = 0 To cmboList.ListCount - 1
       dicFilterList.Add CStr(i), cmboList.list(i)
   Next
End Sub


Private Sub GetList(dicFilterList As Variant, cmboList As Object)
   'populate the list "cmboList" from the collection "sc".
   Dim Item As Variant
   'clear the list first
   cmboList.Clear
   'clear the edit area
   Me.txtEdit = ""
   
   For Each Item In dicFilterList
       cmboList.AddItem dicFilterList.Item(Item)
   Next
   
End Sub

Private Sub txtEdit_DblClick(Cancel As Integer)
    On Error GoTo txtEditDblClickError
    Dim filter As String
    Dim FilterCopy As String
    
    'For the sake of clairity, change an empty filter into the words FiltersCleared = "- Clear All Filters -"
    'to keep blank lines from appearing in the list.
    
    txtEdit.Value = txtEdit.Text
    filter = Me.txtEdit.Value
    
    FilterCopy = IIf(filter = "", FiltersCleared, filter)
    filter = IIf(filter = FiltersCleared, "", filter)
    
    MvGridMain.filter = filter
    MvGridMain.FilterOn = True
    MvGridMain.ApplyFilter 0, 1
    
    If FilterExistsInList(FilterCopy) = False Then
        CmboMulti.AddItem FilterCopy
    End If
    
txtEditDblClickExit:
    Exit Sub
txtEditDblClickError:
    MsgBox "The specified filter has an error in it and could not be applied." & vbCrLf & "Please revise the filter before re-applying it.", vbExclamation
    Resume txtEditDblClickExit
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo txtEditKeyDownError
    Dim filter As String
    Dim FilterCopy As String
    
    If KeyCode = 13 Then
        'the user has just updated the filter text and hits <Enter>, apply the filter.
        
        txtEdit.Value = txtEdit.Text
        filter = Me.txtEdit.Value
        FilterCopy = IIf(filter = "", FiltersCleared, filter)
        filter = IIf(filter = FiltersCleared, "", filter)
    
        MvGridMain.filter = filter
        MvGridMain.FilterOn = True
        'add the edited filter to the list if it is unique
        MvGridMain.ApplyFilter 0, 1
        If FilterExistsInList(FilterCopy) = False Then
            'only insert unique filters
            CmboMulti.AddItem FilterCopy
            
        End If
        
        DoEvents
    End If
    
txtEditKeyDownExit:
    Exit Sub
txtEditKeyDownError:
    MsgBox "The specified filter has an error in it and could not be applied." & vbCrLf & "Please revise the filter before re-applying it.", vbExclamation
    Resume txtEditKeyDownExit
End Sub

'it gets the Form ID of the active screen
Public Function GetActiveFormID() As Byte
    On Error GoTo GetActiveFormIDError
    Dim FormID As Byte
    Dim i As Byte
    Dim FormName As String
    Dim screenCount As Integer
    screenCount = 0
    FormID = 0
    
    For i = 1 To 20
        'If a screen enum exists for this "i", check to see if its name matches this one.
        If Not Scr(i) Is Nothing Then
            screenCount = screenCount + 1
            FormName = screen.ActiveForm.Caption
            If Scr(i).Caption = FormName Then
                FormID = i
                Exit For
            End If
        End If
    Next i
    
    If isUnregistered = True Then
        'run this only if a screen is being closed to avoid unnecessary calls to code
        isUnregistered = False
        If FormID = 0 And screenCount >= 1 And cScreens.Count >= 1 Then
            'no formID was not found check if there is an active screen in Scr Array
            'and if this is the only screen in the Filter History Form
                Set MvMainScreen = Nothing 'set the main screen to nothing so data is refreshed in the form
                Dim j As Integer
                For j = 20 To 1 Step -1
                    If Not Scr(j) Is Nothing Then
                        'guess the next id of the screen
                        FormID = j
                        Exit For
                    End If
                Next j
        End If
    End If
    
    
GetActiveFormIDExit:
    GetActiveFormID = FormID
    Exit Function
    
GetActiveFormIDError:
    GetActiveFormID = 0
    Resume GetActiveFormIDExit
End Function


Private Sub Form_Resize()
    '03.12.2009 DB: Added code to prevent error when form resized vertically to very small size.
    'Reved form to 1.2
    
    Dim ButtonIntraSpacing As Integer
    Dim ButtonSpacing As Integer
    
    ButtonSpacing = 60 'Me.pnlSurround.Width - Me.CmdOk.Left - Me.CmdOk.Width
    ButtonIntraSpacing = 120 'Me.CmdOk.Left - Me.CmdApply.Left - CmdApply.Width
    
    'don't let it get any smaller than...
    If Me.InsideWidth < 4320 Then
        Me.InsideWidth = 4320
    End If
    
    'fit the width
    Me.pnlSurround.Width = Me.InsideWidth
    Me.CmboMulti.Width = Me.InsideWidth - (2 * Me.CmboMulti.left)
    imgHeaderBackground.Width = Me.InsideWidth - (2 * Me.imgHeaderBackground.left)
    lblTitle.Width = Me.InsideWidth - (2 * lblTitle.left)
    Me.txtEdit.Width = Me.InsideWidth - (2 * Me.txtEdit.left)
    Me.imgFooterBackground.Width = Me.InsideWidth - (2 * Me.imgFooterBackground.left)
    
    'don't let the pnlSurround.Height be set less than 0
    If (Me.InsideHeight - Me.FormFooter.Height - Me.FormHeader.Height) > 0 Then
        Me.pnlSurround.Height = Me.InsideHeight - Me.FormFooter.Height - Me.FormHeader.Height
    End If
    
    'don't let the pnlSurround.Height be set less than 0
    If (Me.pnlSurround.Height - Me.txtEdit.Height - lblTitle.top) > 0 Then
        Me.txtEdit.top = Me.pnlSurround.Height - Me.txtEdit.Height - lblTitle.top
    End If
    
    'don't let the CmboMulti.Height be set less than 0
    If (Me.pnlSurround.Height - Me.txtEdit.Height - Me.lblTitle.Height - (3 * lblTitle.top)) > 0 Then
        Me.CmboMulti.Height = Me.pnlSurround.Height - Me.txtEdit.Height - Me.lblTitle.Height - (3 * lblTitle.top)
    End If
    Me.CmdOK.left = Me.InsideWidth - Me.CmdOK.Width - ButtonSpacing
    Me.CmdApply.left = Me.CmdOK.left - Me.CmdApply.Width - ButtonIntraSpacing
    
End Sub

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private boundControl As Control
Private rightMenu As CommandBar
Private WithEvents copyButton As CommandBarButton
Attribute copyButton.VB_VarHelpID = -1
Private WithEvents pasteButton As CommandBarButton
Attribute pasteButton.VB_VarHelpID = -1

Public Property Get Title() As String
    Title = mvTitle
End Property
Public Property Let Title(val As String)
    mvTitle = val
    Me.Caption = mvTitle
End Property

Public Sub Bind(ctrl As Control)
On Error GoTo BindError

    Dim CT As Integer
    Dim idx As Integer
    
    Set boundControl = ctrl
    
    CT = ctrl.ListCount - 1
    CmboMulti.Clear
    
    For idx = 0 To CT
        CmboMulti.AddItem ctrl.list(idx)
    Next
    
    
BindExit:
    Exit Sub
    
BindError:
    MsgBox "Error Loading Selections to Zoom Window", vbOKOnly + vbExclamation, "Load Zoom Error"
    GoTo BindExit
    
End Sub

Private Sub CmdCancel_Click()
    'RaiseEvent ListCancelled
    Me.visible = False
End Sub

Private Sub cmdClearCriteria_Click()
    DoCmd.Hourglass True
    
    CmboMulti.Clear
    
    DoCmd.Hourglass False
End Sub

Private Sub cmdCopy_Click()
    CopyToClipBoard vbCrLf
End Sub

Private Sub cmdOk_Click()
    Dim CT As Integer
    Dim idx As Integer
    'Dim vals As String
    
    CT = CmboMulti.ListCount - 1
    
    boundControl.Clear
    
    For idx = 0 To CT
        'vals = vals & "'" & CmboMulti.List(x) & "',"
        boundControl.AddItem CmboMulti.list(idx)
    Next
    
    'If Len(vals) > 0 Then
    '    vals = left(vals, Len(vals) - 1)
    'End If
    
    Me.visible = False
    'RaiseEvent ListUpdated
    'lstValues = vals
End Sub

Private Sub CmdPaste_Click()
    PasteFromClipBoard
End Sub

Private Sub copyButton_Click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    CopyToClipBoard vbCrLf
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Form_KeyUp_Error

    Dim CT As Integer
    Dim idx As Integer
    
    If KeyCode = 65 And Shift = 2 Then
        With CmboMulti
            CT = .ListCount - 1
            
            For idx = 0 To CT
                .Selected(idx) = True
            Next
        End With
        
        CmboMulti.SetFocus
    End If
        
    If Me.ActiveControl Is Me.CmboMulti And KeyCode = 46 And Shift = 0 Then
        With CmboMulti
            CT = .ListCount - 1
            
            For idx = CT To 0 Step -1
                If .Selected(idx) Then
                    .RemoveItem idx
                End If
            Next
        End With
    End If
    
    If KeyCode = 27 Then
        Me.visible = False
    End If
    
    If KeyCode = 13 Then
        cmdOk_Click
    End If
    

Form_KeyUp_Exit:
    Exit Sub
    
Form_KeyUp_Error:
    Resume Form_KeyUp_Exit
End Sub

Private Sub pasteButton_Click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    PasteFromClipBoard
End Sub


Private Sub CopyToClipBoard(delimiter As String)
On Error GoTo CopyToClipBoard_Error

    'Copy contents of cmboMulti to the clip board as a CRLF separated value list.
    Dim strClip As String
    Dim X As Long
    
    DoCmd.Hourglass True
        
    'Build the list string from the cmboMuli list contents
    For X = 0 To CmboMulti.ListCount - 1
        strClip = strClip & CmboMulti.list(X) & delimiter
    Next X

    DoCmd.Hourglass False
    
    If Trim$(strClip) <> "" Then
        'put it in the Clipboard
        ClipBoard_SetData strClip
    End If
    
CopyToClipBoard_Success:
    On Error Resume Next
    
    DoCmd.Hourglass False
    
    Exit Sub
    
CopyToClipBoard_Error:
    Resume CopyToClipBoard_Success
End Sub

Private Sub PasteFromClipBoard()
On Error Resume Next

    'Grab and format contents of clipboard and, if possible, paste it in to the CmboMulti list.
    Dim strBuf As String
    
    DoCmd.Hourglass True
    
    strBuf = ClipBoard_GetData()
    
    If Trim(strBuf) <> "" Then
        PasteList strBuf
    End If
        
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
    ElseIf strBuf <> "" Then
        strBuf = strBuf & ","
        staBuf = Split(strBuf, ",")
    End If
    
    If IsArray(staBuf) Then
        For i = 0 To UBound(staBuf)
                'Check to see if it is already in the list
                'ItemFound = CheckExistsInList(Trim(staBuf(i)))
                'If not in the list then add it
                If CheckExistsInList(Trim(staBuf(i))) = False And staBuf(i) <> "" Then CmboMulti.AddItem Trim(staBuf(i))
        Next i
    End If
End Sub

Private Sub CmboMulti_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
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
    CreateContextMenu
End Sub

Private Function CheckExistsInList(StNewData As String) As Boolean
    Dim N As Integer, ItemFound As Boolean
    
    For N = 0 To CmboMulti.ListCount - 1
        If CmboMulti.list(N) = StNewData Then
            ItemFound = True
            Exit For
        End If
    Next N
    CheckExistsInList = ItemFound

End Function

Private Sub CreateContextMenu()
On Error GoTo CreateContextMenu_Error
    
    Dim objCommandBar As CommandBar
    Dim objCommandBarButton As CommandBarButton
    Dim genUtils As New CT_ClsGeneralUtilities
            
    ' clear the existing DecipherMultiSelectRightClick
    genUtils.ClearMenu DecipherMultiSelectRightClick
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
    
CreateContextMenu_Error:
    Exit Sub
End Sub

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'' Form / subform open event orders for your viewing pleasure (I can never remember the order!! Am I getting old?)

''  Form1_EventTracker_SubF 1) Form_Open
''  Form1_EventTracker_SubF 2) Form_Load
''  Form1_EventTracker_SubF 3) Form_Resize
''  Form1_EventTracker_SubF 4) Form_Current
''  Form1_EventTracker_Main 1) Form_Open
''  Form1_EventTracker_Main 2) Form_Load
''  Form1_EventTracker_Main 3) Form_Resize
''  Form1_EventTracker_Main 4) Form_Activate
''  Form1_EventTracker_Main 5) Form_Current
''  Form1_EventTracker_SubF 5) Form_GotFocus
''



Dim lEventCount As Long

Private Sub Form_Activate()
PrintEvent "Form_Activate"
End Sub

Private Sub Form_AfterFinalRender(ByVal drawObject As Object)
PrintEvent "Form_AfterFinalRender"
End Sub

Private Sub Form_AfterLayout(ByVal drawObject As Object)
PrintEvent "Form_AfterLayout"
End Sub

Private Sub Form_AfterRender(ByVal drawObject As Object, ByVal chartObject As Object)
PrintEvent "Form_AfterRender"
End Sub

Private Sub Form_BeforeQuery()
PrintEvent "Form_BeforeQuery"
End Sub

Private Sub Form_BeforeRender(ByVal drawObject As Object, ByVal chartObject As Object, ByVal Cancel As Object)
PrintEvent "Form_BeforeRender"
End Sub

Private Sub Form_BeforeScreenTip(ByVal ScreenTipText As Object, ByVal SourceObject As Object)
PrintEvent "Form_BeforeScreenTip"
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
PrintEvent "Form_BeforeUpdate"
End Sub

Private Sub Form_Current()
PrintEvent "Form_Current"
End Sub

Private Sub Form_DataChange(ByVal Reason As Long)
PrintEvent "Form_DataChange"
End Sub

Private Sub Form_DataSetChange()
PrintEvent "Form_DataSetChange"
End Sub

Private Sub Form_Deactivate()
PrintEvent "Form_Deactivate"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
PrintEvent "Form_Error"
End Sub

Private Sub Form_Filter(Cancel As Integer, FilterType As Integer)
PrintEvent "Form_Filter"
End Sub

Private Sub Form_GotFocus()
PrintEvent "Form_GotFocus"
End Sub

Private Sub Form_Load()
PrintEvent "Form_Load"
End Sub

Private Sub Form_LostFocus()
PrintEvent "Form_LostFocus"
End Sub

Private Sub Form_OnConnect()
PrintEvent "Form_OnConnect"
End Sub

Private Sub Form_Open(Cancel As Integer)
PrintEvent "Form_Open"
End Sub


Private Sub Form_Query()
PrintEvent "Form_Query"
End Sub

Private Sub Form_Resize()
PrintEvent "Form_Resize"
End Sub

Private Sub Form_ViewChange(ByVal Reason As Long)
PrintEvent "Form_ViewChange"
End Sub





Private Sub PrintEvent(sEventName As String)
    lEventCount = lEventCount + 1
    Debug.Print Me.Name & " " & CStr(lEventCount) & ") " & sEventName
End Sub

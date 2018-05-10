Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'SA 11/5/2012 - Moved most code out to class CT_ClsSubGenericDataSheet

'SA 05/17/12 - Added field count constant to support multiple subgenericdatasheets
'Change number to match number of fields in this grid
Private Const SheetFieldCount As Integer = 150

Private subGen As New CT_ClsSubGenericDataSheet

Public Event Activate()
Public Event ApplyFilter(filter As String)
Public Event Current()
Public Event Click()
Public Event DblClick()
Public Event Deactivate()
Public Event FocusLost()
Public Event FocusGot()
Public Event Message(Txt As String)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Unload()
Public Event KeyPressed(AsciiKey As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event BeforeUpdate()
Public Event AfterUpdate()


Property Let IsCustomTotal(Value As Boolean)
    subGen.SetIsCustomTotal = Value
End Property
Property Get IsCustomTotal() As Boolean
    IsCustomTotal = subGen.GetIsCustomTotal
End Property

Property Get SelectionTop() As Long
    SelectionTop = subGen.GetSelectionTop
End Property
Property Get SelectionHeight() As Long
    SelectionHeight = subGen.GetSelectionHeight
End Property
Property Get FldCT() As Integer
    FldCT = subGen.GetFieldCount
End Property

Public Sub InvokeEvent(eventName As String, Args() As Variant)
    CallByName Me, "MouseUp", VbGet
End Sub

Public Function CalcFieldsAdd(fld As CnlyFldDef)
    subGen.CalcFieldsAdd fld
End Function

Public Sub LayoutClear()
    subGen.LayoutClear
End Sub

Public Sub LayoutField(Name As String, Ordinal As Long, Width As Single, CalcFld As Boolean)
    subGen.LayoutField Name, Ordinal, Width, CalcFld
End Sub

Public Sub FormatsClear()
    subGen.FormatsClear
End Sub

Public Function CalcFieldsClear()
    subGen.CalcFieldsClear
End Function

Public Sub InitData(ByVal RecordSource As String, ByVal RecordSourceType As Byte, Optional ByVal useDataSource As String = vbNullString)
    subGen.SetMyForm = Me
    subGen.SetSheetFieldCount = SheetFieldCount
    subGen.InitData RecordSource, RecordSourceType, useDataSource
End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub Field1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub Field2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub Field3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub Field4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub Field5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub Field6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub Field7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub Field8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub Field9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub Field10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub Form_Activate()
    RaiseEvent Activate
End Sub

Private Sub Form_AfterRender(ByVal drawObject As Object, ByVal chartObject As Object)
    Application.Echo True
End Sub

Private Sub Form_AfterUpdate()
    RaiseEvent AfterUpdate
End Sub

Public Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Dim StFilter As String
    Select Case ApplyType
    Case acShowAllRecords '0
        RaiseEvent ApplyFilter("")
        Me.FilterOn = False
    Case acApplyFilter '1
        'Cancel = True
        StFilter = Replace(Me.filter, "[" & Me.Name & "].", "")
        If StFilter <> Me.filter Then
            Me.filter = Replace(Me.filter, "[" & Me.Name & "].", "")
        End If
        RaiseEvent ApplyFilter(Me.filter)
    Case acCloseFilterWindow '2
        'RaiseEvent ApplyFilter("")
    End Select
End Sub

Public Sub ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Form_ApplyFilter Cancel, ApplyType
End Sub
Private Sub Form_BeforeRender(ByVal drawObject As Object, ByVal chartObject As Object, ByVal Cancel As Object)
    Application.Echo False
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    RaiseEvent BeforeUpdate
End Sub

Private Sub Form_Click()
    RaiseEvent Click
End Sub
Private Sub Form_Current()
    'PLACE HOLDER FOR EVENT CAPTURE
    subGen.SetSelectionHeight = Me.SelHeight
    subGen.SetSelectionTop = Me.SelTop
    RaiseEvent Current
End Sub
Private Sub Form_DblClick(Cancel As Integer)
    RaiseEvent DblClick
End Sub
Private Sub Form_Deactivate()
    RaiseEvent Deactivate
End Sub
Private Sub Form_GotFocus()
    RaiseEvent FocusGot
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    subGen.KeyPress KeyAscii
    RaiseEvent KeyPressed(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    subGen.SetSelectionHeight = Me.SelHeight
    subGen.SetSelectionTop = Me.SelTop

    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Form_LostFocus()
    RaiseEvent FocusLost
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.SetSelectionHeight = Me.SelHeight
    subGen.SetSelectionTop = Me.SelTop
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    screen.MousePointer = 0
    Set subGen = Nothing
    RaiseEvent Unload
End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub
Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    subGen.ResetMousePointer
End Sub

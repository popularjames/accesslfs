Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub FieldName_AfterUpdate()
On Error GoTo FieldName_AfterUpdateError
If Me.Parent Is Nothing Then
    Exit Sub
Else
    Me!FieldType = CurrentDb.TableDefs(Me!FieldName.RowSource).Fields(Me.FieldName).Type
End If
FieldName_AfterUpdateExit:
    On Error Resume Next
    Exit Sub
FieldName_AfterUpdateError:
    If Err.Number = 3265 Then
        Resume TryQueryDef
    End If
    MsgBox Err.Description & String(2, vbCrLf) & "Error getting field data type!", vbInformation, "Self Config"
    Resume FieldName_AfterUpdateExit

TryQueryDef:
On Error GoTo FieldName_AfterUpdateExit
    Me!FieldType = CurrentDb.QueryDefs(Me!FieldName.RowSource).Fields(Me.FieldName).Type
    GoTo FieldName_AfterUpdateExit
    
End Sub

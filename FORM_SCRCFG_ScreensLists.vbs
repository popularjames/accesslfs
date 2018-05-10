Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Sub UpdateToggles(FocusField)
On Error GoTo UpdateTogglesError
Dim FldRst As RecordSet, FieldName As String

DoCmd.Hourglass True

Set FldRst = Me.RecordsetClone
FieldName = "" & Me!FieldName

If Me(FocusField) Then
    With FldRst
        .MoveFirst
        Do Until .EOF
            If (!FieldName <> FieldName) And FldRst(FocusField) = True Then
                .Edit
                FldRst(FocusField) = False
                .Update
            End If
            .MoveNext
        Loop
    End With
End If

UpdateTogglesExit:
    On Error Resume Next
    DoCmd.Hourglass False
    FldRst.Close
    Set FldRst = Nothing
    Exit Sub
    
UpdateTogglesError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Updating Toggle Field: " & FocusField, vbCritical, "Auto Updater"
    Resume UpdateTogglesExit
End Sub

Private Sub AlternateDisplay_AfterUpdate()
UpdateToggles "AlternateDisplay"
End Sub

Private Sub Bound_AfterUpdate()
UpdateToggles "Bound"
End Sub

Private Sub FieldName_AfterUpdate()
On Error GoTo FieldName_AfterUpdateError
If Me.Parent Is Nothing Then
    Exit Sub
Else
    Me.FieldType = CurrentDb.TableDefs(Me!FieldName.RowSource).Fields(Me.FieldName).Type
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

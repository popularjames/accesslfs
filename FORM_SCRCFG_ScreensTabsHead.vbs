Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdBrowse_Click()
'Browse for app file
On Error GoTo ErrorHandler
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select icon (16x16)"
        .InitialFileName = CurrentProject.Path
        .AllowMultiSelect = False
        .filters.Clear
        .filters.Add "Icons", "*.ico"
        If .show Then
            txtTabImage = .SelectedItems(1)
        End If
    End With
ExitNow:
On Error Resume Next
    Set fd = Nothing
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error loading file"
    Resume ExitNow
End Sub

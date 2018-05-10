Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event ExcelClosed()

Private MvReturn As Integer
Private MvExcel As CT_ClsExcel


Property Get ReturnValue() As Integer
    ReturnValue = MvReturn
End Property

Private Sub CmdCancel_Click()
    MvReturn = vbCancel
    RaiseEvent ExcelClosed
End Sub

Private Sub CmdGetFile_Click()
    Dim stName As String
    Dim stDir As String
    stName = "" & Me.TxtFileName
    
    If stName <> "" Then
        stDir = Mid(stName, 1, InStrRev(stName, "\"))
    End If
    stName = FileDialog(1, "SAVE EXCEL FILE AS", Me.hwnd, stDir, "EXCEL FILE (*.xls)" & Chr(0) & "*.xls" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0), stName)
    
    If stName <> TxtFileName Then
        TxtFileName = stName
        TxtFileName_Change
    End If
End Sub

Private Sub cmdOk_Click()
    If Me.tbGrid.Pages(1).PageIndex = 0 Then
        If Nz(cmbExportGrid.Value, "") = "" Then
             MsgBox "Please select the grid data you want to export!"
             Exit Sub
         End If
    End If
With MvExcel
    .AutoStart = Me.CkAutoLaunch
    .IncludeFormats = Me.CkSendFormatted
    .FileName = Me.TxtFileName
    .Overwrite = Me.CkOverWrite
    .ExpObject = Me.OptGrpSendData
End With
    MvReturn = vbOK
    RaiseEvent ExcelClosed
End Sub

Private Sub Form_Close()
'PLACE HOLDER FOR EVENT CATCHING
RaiseEvent ExcelClosed
End Sub

Private Sub Form_Load()
'PLACE HOLDER FOR EVENT CATCHING
MvReturn = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'PLACE HOLDER FOR EVENT CATCHING
RaiseEvent ExcelClosed
End Sub


Property Let ExcelClass(data As CT_ClsExcel)
    Set MvExcel = data
    Me.TxtFileName = MvExcel.FileName
End Property

Property Get ExcelClass() As CT_ClsExcel
    Set ExcelClass = MvExcel
End Property

Private Sub TxtFileName_Change()
    If Not MvExcel Is Nothing Then
        MvExcel.FileName = "" & TxtFileName
    End If
End Sub

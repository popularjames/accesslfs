Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'SA 03/22/2012 - CR1782 Changed SortList from ActiveX object to Access ListBox

Private SortList As listBox
Private ScreenName As String
Public ScreenID As Long

Public Sub InitData()
On Error GoTo InitDataError
    Dim SQL As String
    SQL = "SELECT SortName FROM SCR_ScreensSorts WHERE ScreenID=" & _
            ScreenID & " GROUP BY SortName ORDER BY SortName"
    Me.ListSorts.RowSource = SQL

InitDataExit:
On Error Resume Next
    Exit Sub
InitDataError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Initializing Sort Lists!", vbCritical
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DeleteSortError
    If ListSorts.ListIndex <> -1 Then
        If MsgBox("Are you sure you want to delete sort '" & Me.ListSorts & "'", _
                vbQuestion + vbYesNo + vbDefaultButton2, "Delete Sort") = vbYes Then
            CurrentDb.Execute "DELETE FROM SCR_ScreensSorts WHERE ScreenID=" & _
                ScreenID & " AND SortName=" & Chr(34) & Me.ListSorts & Chr(34)
        End If
    End If

    InitData
    
    Me.ListSortList.RowSource = vbNullString

DeleteSortExit:
On Error Resume Next
    Exit Sub
DeleteSortError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Removing Sort!", vbCritical, "Delete Error"
    Resume DeleteSortExit
End Sub

Private Sub CmdDeleteAll_Click()
On Error GoTo DeleteSortError

    If MsgBox("Are you sure you want to delete all sorts for screen '" & ScreenName & "'?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Sort") = vbYes Then
        CurrentDb.Execute "DELETE FROM SCR_ScreensSorts WHERE ScreenID=" & ScreenID
    End If

    InitData
    
    Me.ListSortList.RowSource = vbNullString

DeleteSortExit:
On Error Resume Next
    Exit Sub
DeleteSortError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Removing Sort!", vbCritical, "Delete Error"
    Resume DeleteSortExit
End Sub

Private Sub cmdLoad_Click()
On Error GoTo LoadError
    'SA 03/22/2012 - CR1782 Add sort fields to control tip
    Dim i As Integer
    Dim sortTip As String
    sortTip = "ORDER BY "
    
    SortList.RowSource = vbNullString
    For i = 0 To Me.ListSortList.ListCount - 1
        SortList.AddItem Me.ListSortList.Column(0, i) & ";" & Me.ListSortList.Column(1, i)
        
        If Me.ListSortList.Column(0, i) = "A" Then
            sortTip = sortTip & Me.ListSortList.Column(1, i) & ", "
        Else
            sortTip = sortTip & Me.ListSortList.Column(1, i) & " DESC, "
        End If
    Next
    
    sortTip = left(sortTip, Len(sortTip) - 2)
    SortList.ControlTipText = left(sortTip, 255)

LoadExit:
On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
LoadError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Loading Sort List!", vbCritical, "Load Error"
    Resume LoadExit
End Sub

Property Let BoundSortList(Criteria As listBox)
    Set SortList = Criteria
End Property
Property Let BoundScreenName(Criteria As String)
    ScreenName = Criteria
    Me.Caption = ScreenName & " Sorts"
End Property

Property Get BoundSortList() As listBox
    Set BoundSortList = SortList
End Property

Property Get BoundScreenName() As String
    BoundScreenName = ScreenName
End Property

Private Sub ListSorts_AfterUpdate()
    Dim SQL As String
    SQL = "SELECT SortOrder, FieldName FROM SCR_ScreensSorts WHERE ScreenID=" & _
            ScreenID & " AND SortName=" & Chr(34) & Me.ListSorts & Chr(34) & " ORDER BY SortIndex"
    Me.ListSortList.RowSource = SQL
End Sub

Private Sub ListSorts_DblClick(Cancel As Integer)
    cmdLoad_Click
End Sub

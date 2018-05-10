Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'2014-07-01 VS: Show ALJ Claims To Add Screen
'2014-07-03 VS: Added more functions

Public Sub cmdAddSelected_Click()

Dim i As Long
Dim addlist() As Variant
addlist = selectedClaims()

For i = 0 To UBound(addlist, 2) - 1

    algAppealNum = addlist(1, i)
    algCnlyClaimNum = addlist(0, i)
    Add_To_Existing_ALJ_Package (algPackageName)
    
    Clear_Exception_ALJ (algCnlyClaimNum)

Next i

End Sub

Private Sub cmdExit_Click()

    DoCmd.Close

End Sub

Private Sub cmdSelectAll_Click()

    Call SelectAll

End Sub

Function selectedClaims() As Variant
    Dim PackageName As String
    Dim Msg As String
    Dim i As Integer
    Dim ind As Integer
    Dim max As Integer
    Dim oItem As Variant
    Dim X As Variant
    Dim Y As Variant
    Dim claimList() As Variant
    ReDim claimList(0 To 1, 0 To 1) As Variant
    
    max = Me.lstExistClaimsToAdd.ItemsSelected.Count
    ind = 0
    
     If Me.lstExistClaimsToAdd.ItemsSelected.Count <> 0 Then
        For Each oItem In Me.lstExistClaimsToAdd.ItemsSelected

            claimList(0, ind) = Me.lstExistClaimsToAdd.Column(1, oItem)
            claimList(1, ind) = Me.lstExistClaimsToAdd.Column(0, oItem)
        
            ind = ind + 1
            ReDim Preserve claimList(0 To 1, 0 To ind) As Variant
            
        Next oItem
     End If

selectedClaims = claimList
End Function

Public Function SelectAll() As Boolean
On Error GoTo Err_handler
    'Purpose:   Select all items in the multi-select list box.
    'Return:    True if successful
    Dim lngRow As Long

    If Me.lstExistClaimsToAdd.MultiSelect Then
        For lngRow = 0 To Me.lstExistClaimsToAdd.ListCount - 1
            Me.lstExistClaimsToAdd.Selected(lngRow) = True
        Next
        SelectAll = True
    End If

Exit_Handler:
    Exit Function

Err_handler:
    MsgBox (Err.Description)
    Resume Exit_Handler
End Function

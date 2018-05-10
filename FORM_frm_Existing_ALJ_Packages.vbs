Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'2014-06-18 VS: Show existing ALJ Pakages Screen
'2014-07-03 VS: Added Add Multiple Claims button
'2014-07-16 VS: Added Edit Package button

Public Sub CmdAdd_Click()
      
Call Add_To_Existing_ALJ_Package(selectedPackage())
End Sub

Public Sub cmdCreate_Click()

Call Create_New_ALJ_Package
Me.lstExistPkgs.Requery
Me.Refresh
End Sub

Public Sub cmdView_Click()

algPackageName = selectedPackage()

DoCmd.OpenForm "frm_ALJ_Package_Details", , , , , , Me.Name

End Sub

Public Sub cmdAddMultiple_Click()

Found_ALJ_Claims (selectedPackage)

End Sub

Public Sub cmdEdit_Click()

algPackageName = selectedPackage()

DoCmd.OpenForm "frm_Edit_ALJ_Package", , , , , , Me.Name

End Sub

Public Sub cmdDelete_Click()

algPackageName = selectedPackage()

If MsgBox("Are you sure you want to delete " + algPackageName + " package?", vbQuestion + vbYesNo, "Delete Selected Package") = vbNo Then
            Exit Sub
        End If

Call Delete_ALJ_Package(algPackageName)
Me.lstExistPkgs.Requery
Me.Refresh
End Sub

Private Sub cmdExit_Click()

    DoCmd.Close
    
End Sub

Private Sub cmdRefreshList_Click()

   Call Find_ALJ_Package
    
End Sub

Function selectedPackage() As String
    Dim PackageName As String
    Dim Msg As String
    Dim i As Integer
    
      Msg = "You selected" & vbNewLine
            For i = 0 To Me.lstExistPkgs.ListCount - 1
            If Me.lstExistPkgs.Selected(i) Then
              Msg = Msg & Me.lstExistPkgs.Column(0) & vbNewLine
              PackageName = Me.lstExistPkgs.Column(0)
              algPackageName = Me.lstExistPkgs.Column(0)
              algJudgeName = Me.lstExistPkgs.Column(1)
              algHearDate = Me.lstExistPkgs.Column(2)
            End If
Next i

selectedPackage = PackageName
End Function

 

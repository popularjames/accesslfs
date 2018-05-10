Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'2014-07-16 VS: Added ALJ Fax Queue Document Level prototype
Private Sub cmdViewImage_Click()
    If Me.txtDocPath.Value <> "" Then
        Application.FollowHyperlink (Me.txtDocPath.Value)
    Else
        MsgBox ("There is no document associated with this package")
    End If
End Sub

Private Sub cmdSend_Click()
        Call sendALJFax(Me.txtDocPath, Me.txtFaxNumber, Me.txtPackageNameDtl, Me.txtDocType)
End Sub

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private cintLinkedPayerNameId As Integer
Private cstPayerNameIdsForThisConcept As String

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get PayersForThisConcept() As String
    PayersForThisConcept = cstPayerNameIdsForThisConcept
End Property
Public Property Let PayersForThisConcept(sPayersForThisConcept As String)
    cstPayerNameIdsForThisConcept = sPayersForThisConcept
    Call RefreshData
End Property


Public Property Get LinkedPayerNameId() As Integer
    If cintLinkedPayerNameId = 0 Then
        ' either default to all: 1000
        ' or reach up to the parent form... yeah, that..

        cintLinkedPayerNameId = Nz(Me.Parent.Controls("cmbPayer").Value, 1000)
    End If
    LinkedPayerNameId = cintLinkedPayerNameId
    
End Property
Public Property Let LinkedPayerNameId(iLinkedPayerNameId As Integer)
    cintLinkedPayerNameId = iLinkedPayerNameId
End Property


Public Sub RefreshData()

    Me.cmbPayerNameId.RowSource = "SELECT XREF_PAYERNAMES.PayerNameId, XREF_PAYERNAMES.PayerName FROM XREF_PAYERNAMES WHERE XREF_PAYERNAMES.ForUserDisplay <> 0 AND " & _
            " PayerNameID IN (1000," & cstPayerNameIdsForThisConcept & ") ORDER BY XREF_PAYERNAMES.PayerName;"
    Me.Requery
End Sub


Private Sub Form_AfterInsert()
Debug.Print "After Insert"
End Sub

Private Sub Form_AfterUpdate()
Debug.Print "AfterUpdate"
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me.cmbPayerNameId = Me.LinkedPayerNameId
End Sub

Private Sub TrackComment_Change()
    mbRecordChanged = True
End Sub

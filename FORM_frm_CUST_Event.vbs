Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbSelectedProviderContact_Change()
Dim ProvAddrID, AddrType As String
Dim rs As ADODB.RecordSet

    Set rs = Parent.rsEvent
    rs.MoveFirst
    
    ProvAddrID = cmbSelectedProviderContact.Column(4, cmbSelectedProviderContact.ListIndex)
    AddrType = cmbSelectedProviderContact.Column(0, cmbSelectedProviderContact.ListIndex)
    
    rs("ProvAddrID") = ProvAddrID
    rs("AddrType") = AddrType
    rs.UpdateBatch
    txtSelectedProviderContact = cmbSelectedProviderContact.Column(1, cmbSelectedProviderContact.ListIndex) & IIf(Trim(cmbSelectedProviderContact.Column(2, cmbSelectedProviderContact.ListIndex)) = "", "", " - " & cmbSelectedProviderContact.Column(2, cmbSelectedProviderContact.ListIndex) & " - " & cmbSelectedProviderContact.Column(3, cmbSelectedProviderContact.ListIndex))
    
End Sub

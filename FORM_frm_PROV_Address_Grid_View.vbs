Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents frmAddrDetail As Form_frm_PROV_Addr
Attribute frmAddrDetail.VB_VarHelpID = -1
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1

Public Event RecordChanged()

Private strRowSource As String

Private mrsPROVAddr As ADODB.RecordSet
Private mrsPROVAddrPortal As ADODB.RecordSet
Private mrsPROVAddrDeleted As ADODB.RecordSet

Private mbRecordLocked As Boolean
Private miAppPermission As Integer

Const CstrFrmAppID As String = "ProvAddr"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property
Property Let CnlyRowSource(data As String)
     strRowSource = data
End Property
Property Get CnlyRowSource() As String
     CnlyRowSource = strRowSource
End Property
Property Set PortalAddrRecordSource(data As ADODB.RecordSet)
     Set mrsPROVAddrPortal = data
End Property
Property Get PortalAddrRecordSource() As ADODB.RecordSet
     Set PortalAddrRecordSource = mrsPROVAddrPortal
End Property

Property Set AddrRecordSource(data As ADODB.RecordSet)
     Set mrsPROVAddr = data
     strRowSource = ""
End Property
Property Get AddrRecordSource() As ADODB.RecordSet
     Set AddrRecordSource = mrsPROVAddr
End Property

Property Set DeletedAddrRecord(data As ADODB.RecordSet)
     Set mrsPROVAddrDeleted = data
End Property

Property Get DeletedAddrRecord() As ADODB.RecordSet
     Set DeletedAddrRecord = mrsPROVAddrDeleted
End Property

Property Let RecordLocked(data As Boolean)
     mbRecordLocked = data
End Property

Property Get RecordLocked() As Boolean
     RecordLocked = mbRecordLocked
End Property

Public Sub RefreshData()
    Dim OldBookMark
    Dim strError As String
    On Error GoTo ErrHandler
    
    If strRowSource <> "" And mrsPROVAddr Is Nothing Then
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        MyAdo.sqlString = CnlyRowSource
        
        Set mrsPROVAddr = MyAdo.OpenRecordSet()
    End If
    
    If Not (mrsPROVAddr.BOF And mrsPROVAddr.EOF) Then
        OldBookMark = mrsPROVAddr.Bookmark
    End If
    
    Set Me.RecordSet = mrsPROVAddr
    
    If Not (mrsPROVAddr.BOF And mrsPROVAddr.EOF) Then
        mrsPROVAddr.Bookmark = OldBookMark
    End If
    
exitHere:
    Set MyAdo = Nothing
    Exit Sub
    
ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub



Private Sub Form_Close()
    DoCmd.SetWarnings True
End Sub

Private Sub Form_DblClick(Cancel As Integer)
    If mrsPROVAddr.recordCount > 0 Then
        Set frmAddrDetail = New Form_frm_PROV_Addr
        Set frmAddrDetail.AddrRecordSource = mrsPROVAddr
        Set frmAddrDetail.PortalRecordSource = mrsPROVAddrPortal
        frmAddrDetail.RecordLocked = Me.RecordLocked
        frmAddrDetail.DisableMousewheel = True
        frmAddrDetail.RefreshMain
        ShowFormAndWait frmAddrDetail
        Set frmAddrDetail = Nothing
        Set Me.Parent.myPROV.mrsPROVAddrPortal = mrsPROVAddrPortal
        RefreshData
    End If
End Sub

Private Sub Form_Delete(Cancel As Integer)
    If MsgBox("Are you sure you want to delete this record?", vbYesNo) = vbYes Then
        Cancel = False
        If IsSubForm(Me) Then
            Me.Parent.RecordChanged = True
        End If
        With mrsPROVAddrDeleted
            .AddNew
            !cnlyProvID = Me.cnlyProvID
            !AddrType = Me.AddrType
            !EffDt = Me.EffDt
        End With
    Else
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    Dim strCnlyProvID As String
    
    Call Account_Check(Me)
    miAppPermission = UserAccess_Check(Me)
    If miAppPermission = 0 Then Exit Sub
      
    Set mrsPROVAddr = CreateObject("ADODB.Recordset")
    Set mrsPROVAddrDeleted = CreateObject("ADODB.Recordset")
    
    If IsSubForm(Me) = False Then
        strCnlyProvID = InputBox("Please enter CnlyProvID: ")
        If strCnlyProvID <> "" Then
            Set MyAdo = New clsADO
            MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
            MyAdo.sqlString = "select * from prov_address where cnlyprovid = '" & strCnlyProvID & "'"
            Set mrsPROVAddr = MyAdo.OpenRecordSet
            MyAdo.sqlString = "select * from PROV_ADDRESS_PortalPending where cnlyprovid = '" & strCnlyProvID & "' AND ReqStatusCode = '01'"
            Set mrsPROVAddrPortal = MyAdo.OpenRecordSet
            Set MyAdo = Nothing
        End If
        RefreshData
    End If
    DoCmd.SetWarnings False
End Sub

Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    Err.Raise ErrNum, ErrSource, ErrMsg
End Sub
Private Sub frmAddrDetail_RecordChanged()
    RaiseEvent RecordChanged
    If IsSubForm(Me) Then
        Me.Parent.RecordChanged = True
    End If
End Sub
Private Function HighlightField(Optional myfld As String) As Boolean
' KD COMEBACK 20120911
If Not mrsPROVAddrPortal Is Nothing Then
If mrsPROVAddrPortal.recordCount > 0 Then
    Dim MyAdo As clsADO
    Set MyAdo = New clsADO
    Dim temprs As ADODB.RecordSet
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "exec CMS_Auditors_Code.dbo.usp_GetProvAddress '" & Me.cnlyProvID & "','" & Me.AddrType & "','" & Me.EffDt & "','" & Me.TermDt & "'"
    Set temprs = MyAdo.OpenRecordSet
    If myfld = "" Then
        HighlightField = IIf(temprs.recordCount > 0, True, False)
    Else
        HighlightField = IIf(temprs.recordCount > 0 And temprs.Fields(myfld) <> Me.Controls(myfld), True, False)
    End If
    Set temprs = Nothing
End If
End If
End Function

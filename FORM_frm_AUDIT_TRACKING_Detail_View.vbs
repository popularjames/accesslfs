Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private MyAdo As clsADO

Private mstrRecSQL As String
Private mrsBEFORE As ADODB.RecordSet
Private mrsAFTER As ADODB.RecordSet

Public Property Let RecordSQL(data As String)
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    mstrRecSQL = data
    MyAdo.sqlString = mstrRecSQL + " and HistImage = 'BEFORE'"
    Set mrsBEFORE = MyAdo.OpenRecordSet
    
    MyAdo.sqlString = mstrRecSQL + " and HistImage = 'AFTER'"
    Set mrsAFTER = MyAdo.OpenRecordSet
    
    Set MyAdo = Nothing
End Property

Public Sub RefreshData()
    Dim i As Integer
    Dim iMaxCol As Integer
    Dim vBefore, vAfter
    
    Dim rs As ADODB.RecordSet
    
    Set rs = CreateObject("ADODB.Recordset")
    With rs
        .Fields.Append "ColName", adChar, 255
        .Fields.Append "BEFORE", adChar, 8000
        .Fields.Append "AFTER", adChar, 8000
    End With
    
    iMaxCol = mrsBEFORE.Fields.Count
    rs.Open
    
    For i = 5 To iMaxCol - 1
        With rs
            If mrsBEFORE.recordCount > 0 Then vBefore = Nz(mrsBEFORE(i), "null") Else vBefore = ""
            If mrsAFTER.recordCount > 0 Then vAfter = Nz(mrsAFTER(i), "null") Else vAfter = ""
            If vBefore <> vAfter Then
                .AddNew
                !colName = mrsBEFORE.Fields(i).Name
                !Before = vBefore
                !After = vAfter
            End If
        End With
    Next i
    rs.UpdateBatch
    
    Set Me.RecordSet = rs
    
End Sub

Private Sub Form_Load()
    Me.Parent.DetailFormLoaded
End Sub

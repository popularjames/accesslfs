Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrAuditTableName As String
Private mstrAuditKey As String

Public Property Let AuditTableName(data As String)
    mstrAuditTableName = data
End Property

Public Property Let AuditKey(data As String)
    mstrAuditKey = data
End Property


Public Sub RefreshData()
    Dim MyAdo As clsADO
    Dim strSQL As String
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    If mstrAuditTableName <> "" Then
        strSQL = "select min(HistRowID) as HistRowID, HistCreateDt, HistUser, HistAction from " & mstrAuditTableName
        If mstrAuditKey <> "" Then
            strSQL = strSQL & " where " & mstrAuditKey
        End If
        
        strSQL = strSQL & " group by HistCreateDt, HistUser, HistAction order by HistCreateDt desc, HistUser, HistAction"
        
        MyAdo.sqlString = strSQL
        Set Me.RecordSet = MyAdo.OpenRecordSet
        MyAdo.DisConnect
    End If
    
    Set MyAdo = Nothing
End Sub

Private Sub Form_Current()
    Dim strSQL As String
    If IsSubForm(Me) And mstrAuditTableName <> "" Then
        If Not (Me.RecordSet Is Nothing) Then
            strSQL = "SELECT * from " & mstrAuditTableName & _
                     " where HistCreateDt = '" & myTime(Me.HistCreateDt) & _
                     "' and HistUser = '" & Me.HistUser & _
                     "' and HistAction = '" & Me.HistAction & "'"
        
            If mstrAuditKey <> "" Then
                strSQL = strSQL & " and " & mstrAuditKey
            End If
                     
            Me.Parent.Form.txtSQLSource = strSQL
            Me.Parent.RefreshDetail
        End If
    End If
End Sub


Public Function myTime(InTime As Date) As String
    Dim myDate As String
    Dim myHour As Integer
    Dim myMinute As Integer
    Dim mySecond As Integer
    Dim myMillisecond As Double
    Dim myTimeValue As Double
    
    myTimeValue = InTime - DateValue(InTime)
    
    myDate = Format(InTime, "mm-dd-yyyy")
    myHour = Int(myTimeValue * 24)
    myMinute = Int(myTimeValue * 24 * 60 - Int(myTimeValue * 24) * 60)
    mySecond = Int(myTimeValue * 24 * 3600 - Int(myTimeValue * 24 * 60) * 60)
    myMillisecond = myTimeValue * 24 * 3600 - Int(myTimeValue * 24 * 3600)
    If Format(myMillisecond * 1000, "#000") = "1000" Then
        mySecond = mySecond + 1
        myMillisecond = 0
        
        If mySecond = 60 Then
            mySecond = 0
            myMinute = myMinute + 1
        End If
        
        If myMinute = 60 Then
            myMinute = 0
            myHour = myHour + 1
        End If
    End If
        
    myTime = myDate & " " & Format(myHour, "#00") & ":" & Format(myMinute, "#00") & ":" & Format(mySecond, "#00") & "." & Format(myMillisecond * 1000, "#000")
End Function

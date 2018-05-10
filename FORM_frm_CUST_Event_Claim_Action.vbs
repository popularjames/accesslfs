Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private mstrErrMsg As String


Private Sub Form_AfterInsert()
    'Me.CnlyClaimNum = Parent.mstrCnlyClaimNum
    'Me.EventID = Parent.miEventID
    'Me.LastUpdateUser = Parent.mstrUserName
    Me.LastUpdateDt = Now()
 
    
    
    'Me.Recordset.Update
End Sub



Private Sub Form_Close()
Dim stuff As String

    stuff = "arf!"
End Sub


Private Sub Form_Current()

Dim currActionID As Long


On Error GoTo Err_handler


    Me.lblActionID.Caption = ""

    If Not Me.Controls("ActionID") Is Nothing Then
        lngActionID = 0
    Else
        lngActionID = Nz(Me.Controls("ActionID").Value, 0)
    End If




    Dim bResult As Boolean
    Dim rs As ADODB.RecordSet
    Dim MyAdo As clsADO
    
    Dim combinedString As String
    Dim strDisplay As String
    Dim strDiv As String
    
    strDiv = String(100, "-")

   combinedString = "SELECT ActionID, EventID, LastUpdateUser, LastUpdateDt, Notes FROM CUST_Event_Claim_Action_Notes where ActionID = '" & lngActionID & "' and EventID = '" & lngEventID & "' "
    
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = combinedString
    Set rs = MyAdo.OpenRecordSet
    
    Do Until rs.EOF
        
        strDisplay = strDisplay & "Added By: " & UCase(rs.Fields("LastUpdateUser")) & " @ " & UCase(rs.Fields("LastUpdateDt")) & vbCrLf & strDiv & vbCrLf & rs.Fields("Notes") & vbCrLf & vbCrLf
       rs.MoveNext
    Loop
        Me.txtNotes.Value = strDisplay
        
Me.lblActionID.Caption = lngActionID
        
Exit_Sub:
    Set MyAdo = Nothing

    Set rs = Nothing
    Exit Sub

Err_handler:
    If Err.Number = 2424 Then
        GoTo Exit_Sub
     Else
        mstrErrMsg = Err.Description
        MsgBox mstrErrMsg
        GoTo Exit_Sub
    End If

End Sub


Private Sub Form_Dirty(Cancel As Integer)

setNotificaton Me.ActionID, Me.EventID

End Sub



Private Sub Form_LostFocus()

    'Update any changes before leaving this form, so they are not lost when focus changes
    'to another claim for the event
    
    
    Me.Parent.SaveEvent



End Sub

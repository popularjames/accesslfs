Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit ' KD added 2/12/2015!  How in the world could a "DEVELOPER" not have this set?!?!?


Private Sub cboLocation_Click()
    Me.RecordSource = "select * from Link_Table_Config where Location = '" & Me.cboLocation & "' order by Server, Database, Table"
    Me.Requery
End Sub

Private Sub cmdExit_Click()
    DoCmd.Close
End Sub

Private Sub cmdLinkAllTable_Click()
    Dim rs As ADODB.RecordSet
    
    UnLinkTables
    
    CurrentDb.Execute ("Update Link_Table_Location SET Active = 0 ")
    
    If Nz(Me.cboLocation, "") = "" Then
        MsgBox "No Location Selected!  Please Choose a Location", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    
    DoCmd.Hourglass True
    
    Set rs = New ADODB.RecordSet
    rs.ActiveConnection = CurrentProject.Connection
    
    If chkDB Then
        rs.Open ("Select * from Link_Table_Config WHERE Location ='" & Me.cboLocation & "'")
    Else
        rs.Open ("Select * from Link_Table_Config WHERE Location ='" & Me.cboLocation & "'")
    End If
    
    While Not rs.EOF
        With rs
            If chkDB Then
                LinkTable "SQL", ![Server], ![Database], , ![Schema]
            Else
                LinkTable "SQL", ![Server], ![Database], ![Table], ![Schema]
            End If
            .MoveNext
        End With
    Wend
    
    rs.Close
    Set rs = Nothing
    
    DoCmd.Hourglass False
    
    MsgBox "Done"
    CurrentDb.Execute ("Update Link_Table_Location SET Active = 1 WHERE LocationID = '" & Me.cboLocation & "'")
    SetApplicationTitle
    
End Sub

Private Sub cmdLinkThisTable_Click()
    On Error GoTo Err_handler
    With Me.RecordSet
    If chkDB Then
        LinkTable "SQL", ![Server], ![Database], , ![Schema]
    Else
        LinkTable "SQL", ![Server], ![Database], ![Table], ![Schema]
    End If
    End With
    MsgBox "Done"

Err_handler:

End Sub

Private Sub cmdRemove_Click()
    Dim iAns
    If Me.RecordSet.recordCount > 0 Then
        iAns = MsgBox("Are you sure you want to delete this table?", vbYesNo)
        If iAns = vbYes Then
            With Me.RecordSet
                UnLinkTables .Table
                .Delete
                .MoveNext
            End With
        End If
    End If
End Sub

Private Sub cmdUnlinkTable_Click()
    UnLinkTables
    CurrentDb.Execute ("Update Link_Table_Location SET Active = 0 ")
    SetApplicationTitle
    MsgBox "Done"
    
End Sub

Private Sub Combo13_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub Form_Load()
    If Not Nz(DLookup("LocationDesc", "Link_Table_Location", "Active = 1"), "") = "" Then
        Me.cboLocation = Trim(DLookup("LocationID", "Link_Table_Location", "Active = 1"))
    Else
        Me.cboLocation = "CMSPROD"
    End If
    cboLocation_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
    chkDB = 0
End Sub

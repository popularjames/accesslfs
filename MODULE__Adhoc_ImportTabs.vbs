Option Compare Database
Option Explicit


Public Sub ImportTabs()

Dim rst As DAO.RecordSet
Dim strSQL As String
Dim arrProfile() As String
Dim intRowID As Long
Dim strTabName As String
Dim intI As Integer

strSQL = "SELECT * FROM GENERAL_Tabs_Local"

Set rst = CurrentDb.OpenRecordSet(strSQL)

While Not rst.EOF
    arrProfile = Split(rst!ProfileID, ";")
    strTabName = rst!TabName
    
    intRowID = DLookup("RowID", "GENERAL_Tabs", "TabName = '" & strTabName & "'")
        
        
    For intI = 0 To UBound(arrProfile())
        strSQL = "Insert INTO GENERAL_Tabs_Linked_ProfileIDs (RowID, ProfileID)"
        strSQL = strSQL & " VALUES ( " & intRowID & ", '" & arrProfile(intI) & "' )"
        CurrentDb.Execute (strSQL)
    Next intI
    
    
        
        
    rst.MoveNext
Wend


End Sub
Option Compare Database
Option Explicit


''' Last Modified: 08/14/2013
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''
'''
'''  TODO:
'''  =====================================
'''  -
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 08/14/2013 - Created class
'''
''' AUTHOR
'''  =====================================
''' Kevin Dearing
'''
'''
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################




Private Const ClassName As String = "mod_SQL_Functions"

' KD Note: This is done really quickly..
' I'm sure I'll need to modify and build on this as time goes by
Public Function AccessSqlToSqlServer(ByVal sSql As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sReturn As String

    strProcName = ClassName & ".AccessSqlToSqlServer"
    
    sReturn = Replace(sSql, """", "'")
    ' How about wildcards: but only when it's in a LIKE
    
'    sReturn = Replace(sReturn, "*", "%")
    sReturn = Replace(sReturn, "#", "'")
    
    
Block_Exit:
    AccessSqlToSqlServer = sReturn
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function GetADOFieldOrdinalPosition(oRs As ADODB.RecordSet) As Scripting.Dictionary
On Error GoTo Block_Err
Dim strProcName As String
Dim dctRet As Scripting.Dictionary
Dim oFld As ADODB.Field
Dim iPos As Integer

    strProcName = ClassName & ".GetADOFieldOrdinalPosition"
    
    Set dctRet = New Scripting.Dictionary
    For Each oFld In oRs.Fields
'        oFld.Properties
        If dctRet.Exists(UCase(oFld.Name)) = True Then
            Stop
        Else
            dctRet.Add UCase(oFld.Name), iPos
        End If
        iPos = iPos + 1
    Next
    
    
Block_Exit:
    Set GetADOFieldOrdinalPosition = dctRet
    Set dctRet = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
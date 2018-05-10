Option Compare Database
Option Explicit

Private Const ClassName As String = "mod_MSAccess_Tables"


Private csTmpTableName As String

Private Function GetTmpTableName() As String
    csTmpTableName = "zzz_" & Replace(Identity.UserName(), ".", "_")
    GetTmpTableName = csTmpTableName
End Function


' Copy ADODB field structure to a new ADODB Recordset
Private Function CopyFields(rs As ADODB.RecordSet) As ADODB.RecordSet
On Error GoTo Err_CopyFields
    Dim newRS As New ADODB.RecordSet, fld As ADODB.Field
    
    For Each fld In rs.Fields
        newRS.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes 'Note Attribute 104 changes to 108
    Next
    Set CopyFields = newRS
    
Exit_CopyFields:
    Exit Function

Err_CopyFields:
    MsgBox Err.Description, vbOKOnly + vbCritical
    Resume Exit_CopyFields
End Function

' Copy an ADODB Recordset to a new Recordset instance.
Private Function CopyRecordset(rs As ADODB.RecordSet) As ADODB.RecordSet
On Error GoTo Err_CopyRecordset
    Dim newRS As New ADODB.RecordSet, fld As ADODB.Field
    Set newRS = CopyFields(rs)
    newRS.Open  'You must open the Recordset before adding new records.
    
    rs.MoveFirst
    Do Until rs.EOF
        newRS.AddNew
        For Each fld In rs.Fields
            newRS(fld.Name) = fld.Value  'Assumes no BLOB fields
        Next
        rs.MoveNext
    Loop
    Set CopyRecordset = newRS

Exit_CopyRecordset:
    Exit Function

Err_CopyRecordset:
    MsgBox Err.Description, vbOKOnly + vbCritical
    Resume Exit_CopyRecordset
End Function

    'Populate the Local temp Table for the current claim.
Public Function CopyDataToLocalTmpTableFromADORS(oRs As ADODB.RecordSet, bForceRemake As Boolean, Optional sTableName As String) As String
On Error GoTo Block_Err
    Dim strProcName As String
    Dim oFld As ADODB.Field
    Dim oDaoRs As DAO.RecordSet

    strProcName = ClassName & ".CopyDataToLocalTmpTableFromADORS"
    If sTableName <> "" Then
        csTmpTableName = sTableName
    End If
    
    If csTmpTableName = "" Then Call GetTmpTableName
    
    If IsTable(csTmpTableName) = False Or bForceRemake = True Then
        Call CreateTableFromADORS(oRs, csTmpTableName, bForceRemake)
    End If
    
    ' Make sure it's empty
    CurrentDb.Execute "DELETE FROM [" & csTmpTableName & "]"
    
    Set oDaoRs = CurrentDb.OpenRecordSet(csTmpTableName, dbOpenTable)
    
    ' populate it:
    If oRs.EOF And oRs.BOF Then
        CopyDataToLocalTmpTableFromADORS = True
        GoTo Block_Exit
    Else
        oRs.MoveFirst
        While Not oRs.EOF
            oDaoRs.AddNew
            For Each oFld In oRs.Fields
                If oFld.Name = "AutoId" Then
                    If isDAOField(oDaoRs, oFld.Name) = True Then
                        oDaoRs(oFld.Name) = CStr(oFld.Value)
                    End If
                Else
                    If isDAOField(oDaoRs, oFld.Name) = True Then
                        oDaoRs(oFld.Name) = oFld.Value
                    End If
                End If
            Next
            oDaoRs.Update
            oRs.MoveNext
        Wend
    End If
    CopyDataToLocalTmpTableFromADORS = csTmpTableName
    
Block_Exit:
    Set oDaoRs = Nothing
    Set oFld = Nothing
    Exit Function
Block_Err:
    MsgBox Err.Description, vbOKOnly + vbCritical
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


' Create a SQL Server temp Table for the current claim Recordset.
Private Function CreateTableFromADORS(oRs As ADODB.RecordSet, sTblName As String, Optional bForceRemake As Boolean = False) As String
On Error GoTo Block_Err
    Dim strProcName As String
    Dim oTDef As DAO.TableDef
    Dim oAdoField As ADODB.Field
    Dim oTblFld As DAO.Field

    strProcName = ClassName & ".CreateTableFromADORS"
    
    If bForceRemake = True Then
        If IsTable(sTblName) = True Then
            CurrentDb.TableDefs.Delete (sTblName)
            CurrentDb.TableDefs.Refresh
        End If
    ElseIf IsTable(sTblName) = True Then
            ' already created. nothing to do
        CreateTableFromADORS = sTblName
        GoTo Block_Exit
    End If
    
    Set oTDef = New DAO.TableDef
    With oTDef
        .Name = sTblName
        For Each oAdoField In oRs.Fields
            Set oTblFld = New DAO.Field
            oTblFld.Name = oAdoField.Name

            If oAdoField.Name = "AutoId" Then 'Trouble displaying dbBigInt on form and then trouble using ado adInteger with larger numbers.
                oTblFld.Type = dbText ' Convert to a string for display.
            ElseIf (oAdoField.Type = adVarChar) Then 'Allow Zero Length strings.
                oTblFld.Type = AdoTypeToDaoType(oAdoField)
                oTblFld.AllowZeroLength = True
            Else
                oTblFld.Type = AdoTypeToDaoType(oAdoField)
            End If
            .Fields.Append oTblFld
        Next

    End With
    
    CurrentDb.TableDefs.Append oTDef
    CreateTableFromADORS = sTblName
    CurrentDb.TableDefs.Refresh
    
Block_Exit:
    Exit Function
Block_Err:
    MsgBox Err.Description, vbOKOnly + vbCritical
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
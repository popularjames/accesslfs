Option Compare Database
Option Explicit

' 2016/01/12    KD: NIRF - remove confidentiality watermark
' 2015/09/12    KD: Fixed 'Linked to' in app title
' 2015/09/12    KD: deployed Next contract version!!!
' 2015/02/18    KD: Fixing check in process.. hopefully.
' 2015/02/12    KD: New contract check in / check out a final test?
' 2015/02/12    KD: New contract check in / check out more testing
' 2015/02/12    KD: New contract check in / check out testing

' 2014/08/11    KD: testing
' 2014/06/19     VS: Link APPEAL_XREF_ALJ_Hearing_AppealsAnalyst table
' 2014/04/09    KD: Deploying for Viktoria
' 2014/03/17    KD: didn't deploy last time..
' 2014/01/30    KD: forgot to deploy last time..
' 2013/10/21    KD: Adding a form that got left out by a bug..
' 2013/08/14    KD:  Deployed for Tuan
'
' 2013/05/01    KD:     Added this module for comments about what was done
'   (add a comment here and the version control will think you've changed
'   something even if you haven't really done anything other than wanting
'   to relink the tables or whatever
'
' 2013/05/01    KD: Change for the Check In video
'
' 2013/05/02    KD: Added a linked table. removed another
' 2013/07/10    BD: Relink for update to v_AR_SETUP_Hdr.
' 2013/07/26    KC: Relink for update to v_Queue_Hdr for updates for the Manager Queue Maintenance
' 2013/12/12    BD: Relink for XREF_ActivityType, XREF_PayType, XREF_InvoiceClass.
' 2014/08/22    BD: Relink for v_PayerNumber_PayerName.


Public Function TestKD()
Dim bSuccess As Boolean

    bSuccess = AddConverterQueueJob("\\ccaintranet.com\dfs-cms-ds\Data\CMS\AnalystFolders\KevinD\CONVERTER_Queue\Bad_Tifs\test.tif", "PDF", "\\ccaintranet.com\dfs-cms-ds\Data\CMS\AnalystFolders\KevinD\CONVERTER_Queue\Bad_Tifs\out", "test-output4.pdf", , True, False, , , , True)
    
    MsgBox "Success: " & CStr(bSuccess)
    
End Function


'Public Function TestKD()
'Dim oAdo As clsADO
'Dim sSql As String
'Dim oRs As ADODB.RecordSet
'Dim sErrors As String
'
'    Set oAdo = New clsADO
'
'
'    sSql = "SELECT top 10 * FROM AUDITCLM_References r WHERE R.RefType = 'LETTER' ORDER BY CreateDt DESC "
'
'
'    With oAdo
'        .ConnectionString = GetConnectString("v_Data_Database")
'        .SQLTextType = SQLtext
'        .sqlString = sSql
'        Set oRs = .ExecuteRS
'    End With
'
'
'    If CombineDocsFromRs(oRs, "RefLink", "\\ccaintranet.com\DFS-CMS-DS\Data\CMS\AnalystFolders\KevinD\CONVERTER_Queue\Bad_Tifs\Combine_Test\combined.doc", True, sErrors) = False Then
'        Stop
'    Else
'        Stop
'    End If
'
'End Function
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------------------------------------------------
'Author :
'Description:
'Create Date:
'Last Modified:
' PD,Oct 26 2011 - CR # 2120 fix - An ODBC error occurs when adding a single fields multiple times
'---------------------------------------------------------------------------------------------------------------------------------

Public Function BuildCriteriaFromSubform(MySubform As Control, ReportID As Long, ByVal CurSQL As String) As String
On Error GoTo BuildCriteriaError
Dim FieldsRst As RecordSet
Dim SubformRst As RecordSet

Dim SQL As String

SQL = "SELECT * "
SQL = SQL & "FROM SCR_ScreensReportsFields "
SQL = SQL & "WHERE ReportID= " & ReportID & " "

Set FieldsRst = CurrentDb.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)
Set SubformRst = MySubform.Form.RecordsetClone

SubformRst.Bookmark = MySubform.Form.Bookmark
SQL = CurSQL
With FieldsRst
    If .EOF And .BOF Then
        'MsgBox "Unable to locate fields for specified report!" & vbCrLf & vbCrLf & "Check the configuration!", vbCritical, "FETCH ERROR"
        GoTo BuildCriteriaExit
    End If
    .MoveFirst
    Do Until .EOF
        If Len(SQL) > 0 Then
            SQL = SQL & " and "
        End If
        
        'DLC 06/10/2010 : Updated to handle international date settings
        If !FieldType = dbDate Then
            SQL = SQL & !RptFieldName & " =#" & Format(SubformRst(!FieldName), "yyyy-mm-dd") & "#"
        Else
            SQL = SQL & !RptFieldName & " =" & GetIdentifier(!FieldType) & SubformRst(!FieldName) & GetIdentifier(!FieldType)
        End If
        .MoveNext
    Loop
End With

BuildCriteriaExit:
    On Error Resume Next
    BuildCriteriaFromSubform = SQL
    Set SubformRst = Nothing
    Set FieldsRst = Nothing
    Exit Function

BuildCriteriaError:
    Select Case Err.Number
    Case 7951 ' NO Recordset in Subform - Do Nothing
        MsgBox "You must have valid queried data first!"
    Case 3021  ' No Data (No Current Record)
        MsgBox "The data you select has no records!"
    Case Else
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error building fields criteria from subform!", vbInformation, "BAD FUNCTION"
    End Select
    SQL = "ERROR SETTING CRITERIA"
    Resume BuildCriteriaExit
    Resume
End Function


Public Function BuildReportSort(myForm As Form) As String
On Error GoTo BuildSortError
Dim SQL As String
'Build Sort Order
SQL = ""
If Len(myForm!CmboSort1.Value) > 0 Then
    SQL = myForm!CmboSort1 & IIf(myForm!ToggleSort1, " Desc", "")
    If Len(myForm!CmboSort2.Value) > 0 Then
        SQL = SQL & ", " & myForm!CmboSort2 & IIf(myForm!ToggleSort2, " Desc", "")
        If Len(myForm!CmboSort3.Value) > 0 Then
            SQL = SQL & ", " & myForm!CmboSort3 & IIf(myForm!ToggleSort3, " Desc", "")
        End If
    End If
End If
BuildSortExit:
    On Error Resume Next
    BuildReportSort = SQL
    Exit Function
    
BuildSortError:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error building report sort string!", vbInformation, "BAD FUNCTION"
    Resume BuildSortExit
End Function
Public Function BuildDateCriteria(myForm As Form) As String
On Error GoTo BuildCriteriaError
Dim SQL As String
'Build Sort Order
SQL = ""
'Build Criteria
If Len(myForm!CmboFilterDte.Value) > 0 Then
'   HC 5/2010 - removed 2010
    'SQL = myForm!CmboFilterDte & " Between #" & Format(myForm.StartDte.Object.Value, "yyyy-mm-dd") & "# And #" & Format(myForm.EndDte.Object.Value, "yyyy-mm-dd") & "#"
    ' HC 5/2010 - updated 2010
    SQL = myForm!CmboFilterDte & " Between #" & Format(myForm.StartDte.Value, "yyyy-mm-dd") & "# And #" & Format(myForm.EndDte.Value, "yyyy-mm-dd") & "#"

End If

BuildCriteriaExit:
    On Error Resume Next
    BuildDateCriteria = SQL
    Exit Function
    
BuildCriteriaError:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error building report sort string!", vbInformation, "BAD FUNCTION"
    Resume BuildCriteriaExit
End Function
Public Function BuildDetailSort(myForm As Form, Optional SkipOrderBy) As String
On Error GoTo BuildSortError
Dim SQL As String, X As Byte
Dim ordDict As Object
Dim ListItem As String

'Build Sort Order
SQL = ""
'CR # 2120 fix - Use a dictionary object to build a SQL construct with unique values.
Set ordDict = CreateObject("Scripting.Dictionary")
                                                   
'Build Sort Order
If myForm.SortList.ListCount > 0 Then
    If IsMissing(SkipOrderBy) Then SQL = SQL & " ORDER BY "
    For X = 0 To myForm.SortList.ListCount - 1
        ListItem = myForm.SortList.Column(1, X)
        'CR # 2120 fix - Check if field for sorting has already been added to the
        '                dictionary object,if not add it and append to SQL construct.
        If Not ordDict.Exists(ListItem) Then
            ordDict.Add ListItem, X
            If X > 0 Then SQL = SQL & ", "
            SQL = SQL & ListItem & IIf(myForm.SortList.Column(0, X) = "D", " Desc", "")
        End If
    Next X
End If

BuildSortExit:
    On Error Resume Next
    BuildDetailSort = SQL
    Set ordDict = Nothing
    Exit Function
    
BuildSortError:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error building detail sort string!", vbInformation, "BAD FUNCTION"
    Resume BuildSortExit
End Function

Public Function GetReportCfg(ByVal ReportID As Long, ByVal RptConfig As CT_ClsRpt)
On Error GoTo GetReportCfgError
Dim ReportRst As RecordSet
Dim SQL As String

SQL = "SELECT * "
SQL = SQL & "FROM SCR_ScreensReports "
SQL = SQL & "WHERE ReportID= " & ReportID & ""

Set ReportRst = CurrentDb.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)

With ReportRst
    If .EOF And .BOF Then
        MsgBox "Unable to locate specified report!" & vbCrLf & vbCrLf & "Check the configuration!", vbCritical, "FETCH ERROR"
        GoTo GetReportCfgExit
    End If
    RptConfig.ReportName = !ReportName
    RptConfig.EnableSort = !EnableSort
    RptConfig.EnableFilter = !EnableFilter
    RptConfig.EnablePrimary = !EnablePrimary
    RptConfig.EnableSecondary = !EnableSecondary
    RptConfig.ExtraSQL = "" & !ExtraSQL
    RptConfig.EnableTertiary = !EnableTertiary
End With

GetReportCfgExit:
    On Error Resume Next
    ReportRst.Close
    Set ReportRst = Nothing
    Exit Function
    
GetReportCfgError:
    MsgBox "Unable to retreive report Config!" & vbCrLf & vbCrLf & "Check the configuration!", vbCritical, "PROGRAM ERROR"
    Resume GetReportCfgExit
    Resume
End Function


Public Sub LaunchItemGraph(cfg As CnlyScreenCfg, FunctionID As Long, RecordSource As String, ExtraCriteria As String)
On Error GoTo ErrorHappened
Dim locForm As Form_SCR_MainScreens

DoCmd.OpenForm CCAGraphItemCost, acNormal
Set locForm = Scr(cfg.FormID)

With Forms(CCAGraphItemCost)
    .GroupField = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'GroupField'")
    .VenField = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'VenField'")
    .ItemField = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'ItemField'")
    .ListCostField = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'ListCostField'")
    .NetCostField = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'NetCostField'")
    .QtyField = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'QtyField'")
    .VenType = DLookup("FieldType", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'VenField'")
    .ItemType = DLookup("FieldType", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'ItemField'")
    .GraphSource = RecordSource
    .VenNum = locForm.SubForm.Form(DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'VenField'"))
    .ItemNum = locForm.SubForm.Form(DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'ItemField'"))
    .ExtraCriteria = ExtraCriteria
    .DateCriteria = "" & locForm.CmboFilterDte
    ' HC 5/2010 - removed 2010
'    .DateCriteriaFrom = "" & locForm.StartDte.Object.Value
'    .DateCriteriaTo = "" & locForm.EndDte.Object.Value
    ' HC 5/2010 - updated 2010
    .DateCriteriaFrom = "" & locForm.StartDte.Value
    .DateCriteriaTo = "" & locForm.EndDte.Value
    .InitGraph True
End With

ExitNow:
    On Error Resume Next
    Set locForm = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "LaunchItemGraph"
    Resume ExitNow
    Resume
End Sub
Public Sub LaunchDiscGraph(cfg As CnlyScreenCfg, FunctionID As Long, RecordSource As String, ExtraCriteria As String)
On Error GoTo GraphError
Dim locForm As Form_SCR_MainScreens

DoCmd.OpenForm CCAGraphDisc, acNormal
Set locForm = Scr(cfg.FormID)

With Forms(CCAGraphDisc)
    .VenField = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'VenField'")
    .InvDateField = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'InvDateField'")
    .ChkDateField = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'ChkDateField'")
    .InvAmtField = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'InvAmtField'")
    .DiscAmtField = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'DiscAmtField'")
    .VenType = DLookup("FieldType", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'VenField'")
    .GraphSource = RecordSource
    .ExtraCriteria = ExtraCriteria
    .VenNum = "" & locForm.SubForm.Form(.VenField)
    .DateCriteria = "" & locForm.CmboFilterDte
    ' HC 5/2010 - removed 2010
    '.DateCriteriaFrom = "" & locForm.StartDte.Object.Value
    '.DateCriteriaTo = "" & locForm.EndDte.Object.Value
    ' HC 5/2010 - updated 2010
    .DateCriteriaFrom = "" & locForm.StartDte.Value
    .DateCriteriaTo = "" & locForm.EndDte.Value
    .InitGraph False
End With
GraphExit:
    On Error Resume Next
    Set locForm = Nothing
    Exit Sub
GraphError:
    Resume GraphExit
End Sub

Public Sub LaunchVendorNotes(ByVal FormID As Long, FunctionID As Long)
On Error GoTo ErrorHappend
Dim frmNotes As Form_SCR_PopupVendorNotes
Dim FrmScr As Form_SCR_MainScreens
Dim VenFld1 As String, VenFld2 As String
DoCmd.OpenForm CCAVenNotes, acNormal

Set frmNotes = Forms(CCAVenNotes)
Set FrmScr = Scr(FormID)
With frmNotes
    VenFld1 = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'VenID1'")
    VenFld2 = DLookup("FieldName", "SCR_ScreensFunctionsFields", "FunctionID = " & FunctionID & " and FieldDef = 'VenID2'")
    
    If FrmScr.GridForm.RecordSet Is Nothing Then
        MsgBox "You must have data in the main grid to run this function.", vbCritical, "Error Loading Vendor Notes"
        DoCmd.Close acForm, "", acSaveNo
        GoTo ExitNow
    End If
    .VenID1 = FrmScr.GridForm.RecordSet(VenFld1)
    .VenID2 = FrmScr.GridForm.RecordSet(VenFld2)
    
End With
ExitNow:
    On Error Resume Next
    Set frmNotes = Nothing
    Set FrmScr = Nothing
    Exit Sub
ErrorHappend:
    MsgBox Err.Description, vbCritical, "Error Loading Vendor Notes"
    DoCmd.Close acForm, CCAVenNotes, acSaveNo
    Resume ExitNow
    Resume
End Sub
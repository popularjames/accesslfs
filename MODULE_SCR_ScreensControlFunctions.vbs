Option Compare Database
Option Explicit

Public Sub SaveFilter(MyControl As Control)
On Error GoTo SaveFilterError
Dim FilterName As String, tmpFilter As String
Dim FilterRst As RecordSet, tmpStr As String
Dim frm As Form_SCR_MainScreens, SQL As String

Set frm = MyControl.Parent.Parent.Parent
tmpFilter = frm.SubForm.Form.filter
If tmpFilter = "" Then GoTo SaveFilterExit
tmpFilter = Replace(tmpFilter, "[" & frm.SubForm.Form.Name & "].", "")
tmpStr = "Please enter a name for the filter below:" & String(2, vbCrLf) & tmpFilter
GetName:
'filterName = InputBox(tmpStr, "Save Filter", Replace(TmpFilter, Chr(34), Chr(39)))
FilterName = InputBox(tmpStr, "Save Filter", Replace(tmpFilter, "'", ""))
FilterName = Replace(FilterName, Chr(34), "")
If FilterName = "" Then GoTo SaveFilterExit

SQL = "SELECT * FROM SCR_ScreensFilters WHERE ScreenID = " & frm.ScreenID
Set FilterRst = CurrentDb.OpenRecordSet(SQL, dbOpenDynaset, dbSeeChanges)

With FilterRst
    If .BOF And .EOF Then GoTo AddIt
    .FindFirst ("FilterName = " & Chr(34) & FilterName & Chr(34) & " AND UserName = " & Chr(34) & Identity.UserName & Chr(34))
    If .NoMatch Then
AddIt:
        .AddNew
        !ScreenID = frm.ScreenID
        !FilterName = FilterName
        !FilterSQL = tmpFilter
        !UserName = Identity.UserName
        .Update
        .Close
    Else
        If MsgBox("Filter: " & FilterName & "  already exists!" & String(2, vbCrLf) & "Do you want to overwrite it?", vbQuestion + vbYesNo + vbDefaultButton2, "Replace Filter") = vbYes Then
            .Edit
            !ScreenID = frm.ScreenID
            !FilterName = FilterName
            !FilterSQL = tmpFilter
            !UserName = Identity.UserName
            .Update
            .Close
        Else
            tmpStr = "Please enter a different name for the filter below:" & String(2, vbCrLf) & tmpFilter & String(2, vbCrLf) & FilterName & " is already taken!"
            .Close
            GoTo GetName
        End If
    End If
End With


SaveFilterExit:
    On Error Resume Next
    Set frm = Nothing
    MyControl.Requery
    FilterRst.Close
    Set FilterRst = Nothing
    Exit Sub

SaveFilterError:
    If Err.Number = 3022 Then 'Duplicate Query String
        MsgBox "The same filter already exists under a different name!" & String(2, vbCrLf) & "Error saveing filter:" & String(2, vbCrLf) & tmpFilter, vbInformation, "Control Fucntions"
    Else
        MsgBox Err.Description & String(2, vbCrLf) & "Error saveing filter:" & String(2, vbCrLf) & tmpFilter, vbInformation, "Control Fucntions"
    End If
    Resume SaveFilterExit
    Resume
End Sub

Public Sub ClearValues(myForm As Form_SCR_MainScreens)
On Error Resume Next
    'SA 03/22/2012 - CR1782 Updated methods to use Access ListBox
    With myForm
        .StartDte = myForm.Config.StartDate
        .EndDte = myForm.Config.EndDate
        .CmboPrimary = vbNullString
        .LblPrimaryAlternate.Caption = "Primary Alternate Selection"
        .CmboSecondary = vbNullString
        .LblSecondaryAlternate.Caption = "Secondary Alternate Selection"
        .CmboSortFieldList = vbNullString
        .CmboFilterDte = vbNullString
        .CmboFilters = vbNullString
        .CmboReports = vbNullString
        .CmboFunction = vbNullString
        .SortList.RowSource = vbNullString
        .ToggleSort.Value = 0
        .SubForm.Form.filter = vbNullString
        .SubForm.Form.FilterOn = False
        .CmboListPrimaryBy = myForm.Config.PrimaryField
        .CmboListSecondaryBy = myForm.Config.SecondaryField
    End With
End Sub
Public Sub TogleAscDesc(Toggle As Control)
On Error GoTo TogleAscDescError
    Toggle.Caption = IIf(Toggle, "Desc", "Asc")

TogleAscDescExit:
    On Error Resume Next
    Exit Sub
TogleAscDescError:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error Toggling Sort!", vbInformation, "Toggle Error"
    Resume TogleAscDescExit
End Sub
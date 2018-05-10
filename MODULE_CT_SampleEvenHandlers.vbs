Option Compare Database
Option Explicit

Public Function SAMPLE_LoadPowerBar(FormID As Byte)
On Error GoTo ErrorHappened
    #If ccSCR = 1 Then
    Scr(FormID).SubformPowerBar.SourceObject = "PowerBar-Sample"
    #End If
ExitNow:
    On Error Resume Next
    Exit Function
ErrorHappened:
    MsgBox Err.Description, vbCritical
    Resume ExitNow
End Function

' HC 5/2010 - removed 2010
'Public Function SAMPLE_LocalizeDateControls(FormID As Byte)
''0 = Long Date
''1 = Short DAte
''2 = Time
''3 = Custom
'        Scr(FormID).StartDte.Format = 1 'SHORT DATE
'
'        'OR YOU CAN DO THIS
'
'        Scr(FormID).EndDte.Format = 3 'CUSTOM
'        Scr(FormID).EndDte.CustomFormat = "dd/MM/yy" 'CUSTOM
'
'End Function


Public Function SAMPLE_SetPrimaryComboToAuditNumber(FormID As Byte)
#If ccSCR = 1 Then
    Scr(FormID).CmboPrimary = Identity.AuditNum
    Scr(FormID).CmboPrimary_AfterUpdate
#End If
End Function

Public Function SAMPLE_CatchReport(FormID As Byte, ByVal oAny As CT_ClsRpt)

Select Case oAny.ReportName
Case "SOME REPORT"
With oAny
    .EnableSort = True
    .SortString = "SomeField, SomeOtherField"
End With
Case "Some Other Report"
    #If ccSCR = 1 Then
    If Scr(FormID).CmboSortFieldList.ListIndex = 1 Then 'NO SORT SPECIFIED
        'GO LOOK UP SOME SORT BY SOME SET OF VARIABLES (DUP SCREENS SHOULD USE THIS WAY)
    End If
    #End If
End Select



End Function

Public Function SAMPLE_LoadConditionalFormat(FormID As Byte)
#If ccSCR = 1 Then

Dim FrmFormats As Form_SCR_CondFormats
Dim ScreenID As Long
Dim FormatID As Long
Dim FormatName As String
Dim SQL As String
ScreenID = Scr(FormID).ScreenID
FormatName = "Past 30 Days"

SQL = "ScreenID = " & ScreenID & " and FormatName = " & Chr(34) & FormatName & Chr(34)
FormatID = Nz(DLookup("CondFormatID", "SCR_ScreensCondFormats", "ScreenID = " & ScreenID & " and FormatName = " & Chr(34) & FormatName & Chr(34)), 0)

If FormatID = "0" Then
    MsgBox "Unable to set the default condition format to (" & FormatName & ")", vbInformation, "LoadMeijersClaimTrackerConditionalFormat"

Else
       Set FrmFormats = Scr(FormID).SubformCondFormats.Form
       
       
       With FrmFormats

            .ApplyFormatAdd FormatID, FormatName
            .ApplyFormats
       End With
End If

#End If
End Function

Public Function SAMPLE_KeyPressed(FormID As Byte, ByVal KeyAscii As Integer)
    Select Case KeyAscii
        Case 4 'CTL-D: Flag/Unflag current Dup group
            'Code for flagging group
        Case 12 'CTL-L: Flag/Unflag entire list.
            'Code for flagging list
        Case 18 'CTL-R: Flag/Unflag Current Record.
            'Code for flagging record
    End Select

End Function

Public Sub timetest()
Dim dStart As Double
dStart = GetTimer()
'Telemetry.RecordEvent "OpenStreen", "Screen 1"
'Debug.Print GetTimer - dStart & " Seconds"
End Sub

#If ccSCR = 1 Then
Public Sub MakeClaimsScreen()
    RestoreScreenFromXML "", "", "Claims"
End Sub
#End If
Option Compare Database
Option Explicit

'SA 05/21/2012 - CR2782 Reworked SetTabs and InitTabs to allow for loading of different sub generic data sheets
'SA 11/14/2012 - Added ability to load forms other than the SubGenericDatasheet into bottom tabs

Public Const glTabsUsed As Byte = 1 ' The count of the tabs that are used for non-dynamic purposes
Public Const glTabTotalsCustom As Byte = 0

Public Sub SetTabs(ByRef myForm As Form, ByRef cfg As CnlyScreenCfg)
On Error GoTo SetTabsInitError
    Dim CurTab As Byte
    Dim FieldCount As Integer
    Dim SourceObject As String

    If cfg.TabsCT > 0 Then
        'Set source object for all tabs that are used
        For CurTab = 0 To cfg.TabsCT - 1
            'Get field count
            If cfg.Tabs(CurTab).SourceType = Table Then
                FieldCount = CurrentDb.TableDefs(cfg.Tabs(CurTab).Source).Fields.Count
            ElseIf cfg.Tabs(CurTab).SourceType = Query Then
                FieldCount = CurrentDb.QueryDefs(cfg.Tabs(CurTab).Source).Fields.Count
            Else
                FieldCount = 0
                SourceObject = cfg.Tabs(CurTab).Source
            End If

            'Determine correct source object
            Select Case FieldCount
                'SA 11/5/2012 - Added sheet for 25 fields
                Case 1 To 25
                    SourceObject = "CT_SubGenericDataSheet025"
                Case 26 To 50
                    SourceObject = "CT_SubGenericDataSheet050"
                Case 51 To 100
                    SourceObject = "CT_SubGenericDataSheet100"
                Case 101 To 150
                    SourceObject = "CT_SubGenericDataSheet150"
                Case Is > 150
                    SourceObject = "CT_SubGenericDataSheet"
            End Select

            'Set source object
            myForm.Tabs.Pages(CurTab + 1).Controls(0).SourceObject = SourceObject

            'Init tab
            InitTab myForm.Controls("Page" & CurTab + 2), myForm("Subform" & CStr(CurTab + 2)), cfg.Tabs(CurTab)
        Next
    End If

SetTabsExit:
On Error Resume Next
    Exit Sub
SetTabsInitError:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error Initializing Tab Recordset", vbCritical, "Error"
    Resume SetTabsExit
    Resume
End Sub

Public Sub InitTab(ByRef oPage As Object, ByRef oSF As Access.SubForm, ByRef sctTab As CnlyScreenTab)
On Error GoTo ErrorHappened
    With oPage
        .visible = True
        .Caption = sctTab.Caption
        .Tag = sctTab.TabID
    End With
    With oSF
        .visible = True
        If left(.SourceObject, 22) = "CT_SubGenericDataSheet" Then
            'Code specific to CT_SubGenericDataSheet
            .Form.InitData Nz(sctTab.Source, vbNullString), CByte(sctTab.SourceType)
            .Form.DatasheetFontHeight = Identity.DataSheetStyle.fontsize
            .Tag = Nz(sctTab.Source, vbNullString)
        Else
            'Forms
            .Tag = .Form.RecordSource
        End If
        .LinkChildFields = Nz(sctTab.LinkChild, vbNullString)
        .LinkMasterFields = Nz(sctTab.LinkMaster, vbNullString)
        
    End With
ExitNow:

Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Tab Init Error"
    Resume ExitNow
    Resume
End Sub
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private MvCfg As CnlyScreenCfg
Private MvScr As Form_SCR_MainScreens
Private MvNew As Boolean

Public Property Get NewScreen() As Boolean
    NewScreen = MvNew
End Property

Property Get ScreenForm() As Form_SCR_MainScreens
    Set ScreenForm = MvScr
End Property

Public Function CreateScreen(ByVal ScreenName As String) As Form_SCR_MainScreens
On Error GoTo ErrorHappened
    'SA 03/22/2012 - CR2600 Prevent screen from being opened if it was deleted.
    'SA 03/22/2012 - CR2687 Allow users to open the same screen more than once.
    Dim X As Integer
    Dim NewScreenNum As Integer
    Dim MouseState As Integer
    
    MouseState = screen.MousePointer
    
    'Confirm that screen exists
    If Not ScreenExists(ScreenName) Then
        DoCmd.Hourglass False
        MsgBox "Can't find the screen: " & ScreenName & vbCrLf & vbCrLf & _
            "Restart Decipher to refresh the screens list.", vbInformation, "Can't find screen"
        GoTo ExitNow
    End If
    
    NewScreenNum = -1
    MvNew = True
    MvCfg.ScreenName = ScreenName
    
TryAgain:
    'Check to see if screen is already open
    For X = 1 To UBound(Scr)
        If Not Scr(X) Is Nothing Then
            If UCase(ScreenName) = UCase(Scr(X).ScreenName) Then
                'Screen already open prompt user
                DoCmd.Hourglass False
                If MsgBox("The screen " & Chr(34) & ScreenName & Chr(34) & _
                    " is already open." & vbCrLf & vbCrLf & _
                    "Do you want to open another copy of this screen?", vbQuestion + vbYesNo, _
                    "Duplicate Screen") = vbNo Then
    
                    NewScreenNum = X
                    MvNew = False
                End If
                
                Exit For
            End If
        End If
    Next X

    'Get first available screen index for new screens
    If MvNew Then
        For X = 1 To UBound(Scr)
            If Scr(X) Is Nothing Then
                NewScreenNum = X
                Exit For
            End If
        Next X
    End If
    
    'Too many screens open
    If NewScreenNum = -1 Then
        DoCmd.Hourglass False
        MsgBox "The Maximum (" & UBound(Scr()) & ") Screens Have Been Opened!" & _
                String(3, vbCrLf) & "Contact Development!!!", vbCritical, "Screen Launch Failure!"
        GoTo ExitNow
    End If
    
    'Open screen or set focus
    If MvNew Then
        Set MvScr = New Form_SCR_MainScreens
        Set Scr(NewScreenNum) = MvScr
        MvCfg.FormID = NewScreenNum
        MvScr.ScreenName = MvCfg.ScreenName
        Scr(NewScreenNum).FormID = MvCfg.FormID
        Scr(NewScreenNum).ScreenName = MvCfg.ScreenName
    Else
        Set MvScr = Scr(NewScreenNum)
    End If
    
    Scr(NewScreenNum).SetFocus
    DoCmd.Maximize
    
    Set CreateScreen = Scr(NewScreenNum)
ExitNow:
On Error Resume Next
    screen.MousePointer = MouseState
    Exit Function
ErrorHappened:
    Select Case Err.Number
    Case 2467 'The object does not exits - The screen closed incorrectly
        If MvNew = False Then
            Set Scr(NewScreenNum) = Nothing
            Resume TryAgain
        End If
    Case Else
        DoCmd.Hourglass False
        MsgBox Err.Description, vbCritical, "Error Creating Screens"
        Resume ExitNow
    End Select
    
End Function

Public Function ScreenExists(ByVal ScreenName As String) As Boolean
'Query screens table with screen name to see if screen exists
'SA 03/22/2012 - New function to determine if screen exists
On Error GoTo ErrorHappened
    
    If DCount("ScreenID", "SCR_Screens", "ScreenName='" & Replace(ScreenName, "'", "''") & "' AND Included=TRUE") > 0 Then
        ScreenExists = True
    Else
        ScreenExists = False
    End If

ExitNow:
    Exit Function
ErrorHappened:
    ScreenExists = False
    Resume ExitNow
    Resume
End Function

Public Sub GetConfig()
On Error GoTo ErrorHappened
    'SA 03/22/2012 - Removed rstOpen variable and reworked exit handler
    'SA 8/3/2012 - Added user tab config and removed label jumping
    Dim db As Database
    Dim rst As RecordSet
    Dim SQL As String
    Dim X As Integer
    Dim InitProcess As String
    
    InitProcess = "Opening RecordSource"
    SQL = "SELECT * FROM SCR_Screens WHERE ScreenName='" & Replace(MvCfg.ScreenName, "'", "''") & "'"
    
    Set db = CurrentDb
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)

    If rst.recordCount > 0 Then
        With MvCfg
            .ScreenID = rst!ScreenID
            .ScreenName = rst!ScreenName
            .PrimaryRecordSource = Nz(rst!PrimaryRecordSource, vbNullString)
            .PrimaryRecordSourceType = rst!PrimaryRecordSourceType
            .CustomCriteriaListBoxRecordSource = rst!CustomCriteriaListBoxRecordSource
            .DateUse = rst!DateUse
            .StartDate = rst!StartDte
            .EndDate = rst!EndDte
            .PrimaryListBoxRecordSource = Nz(rst!PrimaryListBoxRecordSource, vbNullString)
            .PrimaryListBoxRecordSourceType = rst!PrimaryListBoxRecordSourceType
            .PrimaryListBoxCaption = Nz(rst!PrimaryListBoxCaption, vbNullString)
            .PrimaryListBoxMulti = rst!PrimaryListBoxMulti
            .SecondaryListBoxUse = rst!SecondaryListBoxUse
            .SecondaryListBoxDependency = rst!SecondaryListBoxDependency
            .SecondaryListBoxMulti = rst!SecondaryListBoxMulti
            .SecondaryListBoxRecordSource = Nz(rst!SecondaryListBoxRecordSource, vbNullString)
            .SecondaryListBoxRecordSourceType = rst!SecondaryListBoxRecordSourceType
            .SecondaryListBoxCaption = Nz(rst!SecondaryListBoxCaption, vbNullString)
            .TertiaryListBoxUse = rst!TertiaryListBoxUse
            .TertiaryListBoxDependency = rst!TertiaryListBoxDependency
            .TertiaryListBoxMulti = rst!TertiaryListBoxMulti
            .TertiaryListBoxRecordSource = Nz(rst!TertiaryListBoxRecordSource, vbNullString)
            .TertiaryListBoxRecordSourceType = rst!TertiaryListBoxRecordSourceType
            .TertiaryListBoxCaption = Nz(rst!TertiaryListBoxCaption, vbNullString)
            .TertiaryListBoxPrimaryDependency = Nz(rst!TertiaryListBoxPrimaryDependency, vbNullString)
            .PrimaryField = DLookup("FieldName", "SCR_ScreensListFields", "ScreenID=" & rst!ScreenID & " AND ListLevel=1 AND Bound=true")
            .PrimaryQualifier = GetIdentifier(DLookup("FieldType", "SCR_ScreensListFields", "ScreenID=" & rst!ScreenID & " AND ListLevel=1 and Bound=true"))
            
            If .SecondaryListBoxUse Then
                .SecondaryField = DLookup("FieldName", "SCR_ScreensListFields", "ScreenID=" & MvCfg.ScreenID & " AND ListLevel=2 AND Bound=true")
                .SecondaryQualifier = GetIdentifier(DLookup("FieldType", "SCR_ScreensListFields", "ScreenID=" & MvCfg.ScreenID & " AND ListLevel=2 AND Bound=true"))
            End If
            
            If .TertiaryListBoxUse Then
                .TertiaryField = DLookup("FieldName", "SCR_ScreensListFields", "ScreenID=" & MvCfg.ScreenID & " AND ListLevel=3 AND Bound=true")
                .TertiaryQualifier = GetIdentifier(DLookup("FieldType", "SCR_ScreensListFields", "ScreenID=" & MvCfg.ScreenID & " AND ListLevel=3 AND Bound=true"))
            End If
            
            If DCount("PwrBarID", "SCR_ScreensPowerBars", "ScreenID=" & .ScreenID) > 0 Then
                .PowerBars = True
            Else
                .PowerBars = False
            End If
            
        End With
        rst.Close
        
        'Load user tabs
        SQL = "SELECT TabNumber,TabCaption,TabControlTip,TabStatusbar,TabImage,Subform FROM SCR_ScreensTabsHead " & _
                "WHERE ScreenID=" & MvCfg.ScreenID & " ORDER BY TabNumber"
        Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)
        
        Do Until rst.EOF
            Select Case rst!TabNumber
                Case 1
                    LoadUserTabconfig MvCfg.TabsHeadUser1, rst
                Case 2
                    LoadUserTabconfig MvCfg.TabsHeadUser2, rst
                Case 3
                    LoadUserTabconfig MvCfg.TabsHeadUser3, rst
            End Select
            rst.MoveNext
        Loop
        rst.Close
        
        'Load bottom tab settings
        SQL = "SELECT TabID,RecordSource,Feature,Type FROM SCR_ScreensTabs " & _
                "WHERE ScreenID=" & MvCfg.ScreenID & " ORDER BY Sort, Feature"
        Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)
        
        If rst.recordCount > 0 Then
            With rst
                .MoveLast
                .MoveFirst
                MvCfg.TabsCT = .recordCount
                ReDim MvCfg.Tabs(MvCfg.TabsCT)
                Do Until .EOF
                    MvCfg.Tabs(.AbsolutePosition).TabID = .Fields("TabID")
                    MvCfg.Tabs(.AbsolutePosition).Source = .Fields("RecordSource")
                    MvCfg.Tabs(.AbsolutePosition).Caption = .Fields("Feature")
                    MvCfg.Tabs(.AbsolutePosition).SourceType = .Fields("Type")
                    MvCfg.Tabs(.AbsolutePosition).LinkChild = vbNullString
                    MvCfg.Tabs(.AbsolutePosition).LinkMaster = vbNullString
                    
                    .MoveNext
                Loop
                .Close
            End With
            
            For X = 0 To MvCfg.TabsCT - 1
                SQL = "SELECT ChildField,MasterField FROM SCR_ScreensTabsFields WHERE TabID=" & MvCfg.Tabs(X).TabID
                Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)
        
                If rst.recordCount > 0 Then
                    With rst
                        Do Until .EOF
                            MvCfg.Tabs(X).LinkChild = MvCfg.Tabs(X).LinkChild & .Fields("ChildField") & ";"
                            MvCfg.Tabs(X).LinkMaster = MvCfg.Tabs(X).LinkMaster & "SubForm.Form!" & .Fields("MasterField") & ";"
                            .MoveNext
                        Loop
                    End With
                    MvCfg.Tabs(X).LinkChild = left(MvCfg.Tabs(X).LinkChild, Len(MvCfg.Tabs(X).LinkChild) - 1)
                    MvCfg.Tabs(X).LinkMaster = left(MvCfg.Tabs(X).LinkMaster, Len(MvCfg.Tabs(X).LinkMaster) - 1)
                End If
                rst.Close
            Next X
        Else
            'No tabs
            MvCfg.TabsCT = 0
        End If

        'Apply settings to form and open
        MvScr.Config = MvCfg
        MvScr.InitData
    Else
        MsgBox "There Appears To Be a Problem With The Screen Configuration File!", vbCritical, "INITIALIZATION ERROR"
    End If
ExitNow:
On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description & String(2, vbCrLf) & "Error in Screen Init: " & InitProcess, vbCritical, "INITIALIZATION ERROR"
    Resume ExitNow
    Resume
End Sub

Private Sub LoadUserTabconfig(ByRef UserTab As CnlyScreenTabsHead, ByRef rst As DAO.RecordSet)
'Load use tab settings from recordset
On Error GoTo ErrorHappened
    With UserTab
        .ShowTab = True
        .Caption = rst!TabCaption
        .ControlTip = Nz(rst!TabControlTip, vbNullString)
        .StatusBar = Nz(rst!TabStatusbar, vbNullString)
        .Image = Nz(rst!TabImage, vbNullString)
        .SubForm = Nz(rst!SubForm, vbNullString)
    End With
ExitNow:
On Error Resume Next

    Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub
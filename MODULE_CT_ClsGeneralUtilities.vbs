Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event StatusMessage(Src As String, Msg As String, lvl As Integer)

'MN - 7/28/2009 - Version 2.1 - additions
Private Const QI As String = "'"

Public Function ApplicationIsReadOnly() As Boolean
'DLC 05/20/2010 Determine whether the .MDB is readonly
On Error GoTo ErrorHandler
    Dim blnReadOnly As Boolean
    Dim db As DAO.Database
    Set db = CurrentDb
    blnReadOnly = Not db.Properties("Updatable")
Exit_ErrorHandler:
    On Error Resume Next
    db.Close
    Set db = Nothing
    ApplicationIsReadOnly = blnReadOnly
    Exit Function
ErrorHandler:
    'Report that database is not readonly in the event of an error
    blnReadOnly = False
    Resume Exit_ErrorHandler
End Function

Public Function InchesToTwips(ByVal InchesProp As String) As Integer
    InchesToTwips = 1440 * CDbl(InchesProp)
End Function

Public Function GetLocalDateFormat(ByVal colName As String) As String
    'MN - used to format date in datagrid where column is of type "Date" (SQL Server) and "Text" in Access
    GetLocalDateFormat = "Format(" & colName & ",""" & GetRegionalShortDateFormat() & """)"
End Function
Public Function GetTextValue(ByRef Txt As Object, Optional ByVal ConvertEmptyStringToNull As Boolean = True) As String
    If Nz(Txt.Value, "") = "" Then
        If ConvertEmptyStringToNull Then
            GetTextValue = "NULL"
        Else
            GetTextValue = ""
        End If
    Else
        GetTextValue = QI & EscapeChars(Nz(Txt.Value, ""), QI) & QI
    End If
End Function
' DPS 10/30/2009 provide version of GetTextValue that does not strip cr/lf
Public Function GetLongTextValue(ByRef Txt As Object, Optional ByVal ConvertEmptyStringToNull As Boolean = True) As String
    If Nz(Txt.Value, "") = "" Then
        If ConvertEmptyStringToNull Then
            GetLongTextValue = "NULL"
        Else
            GetLongTextValue = ""
        End If
    Else
        GetLongTextValue = QI & EscapeQuotes(Nz(Txt.Value, ""), QI) & QI
    End If
End Function
Public Function GetNonTextValue(ByRef Txt As Object, Optional ByVal ConvertEmptyStringToNull As Boolean = True) As String
    If Nz(Txt.Value, "") = "" Then
        If ConvertEmptyStringToNull Then
            GetNonTextValue = "NULL"
        Else
            GetNonTextValue = ""
        End If
    Else
        GetNonTextValue = EscapeChars(Nz(Txt.Value, ""), QI)
    End If
    
End Function
'MN - Use this function when Access column is text and SQL server column is Date
' in update/insert sql statements
Public Function GetDateValue(ByRef dt As Object, Optional ByVal ConvertEmptyStringToNull As Boolean = True) As String
    If Nz(dt.Value, "") = "" Then
        If ConvertEmptyStringToNull Then
            GetDateValue = "NULL"
        Else
            GetDateValue = ""
        End If
    Else
        GetDateValue = QI & Format(dt.Value, "yyyy-mm-dd") & QI
    End If
End Function


Public Function GetGuidValue(ByRef Txt As Object) As String
    If Nz(Txt.Value, "") = "" Then
        GetGuidValue = ""
    Else
        GetGuidValue = "{" & ConvertGuid(Nz(Txt.Value, "")) & "}"
    End If
End Function

'Clear Form
Public Sub ClearForm(ByVal frm As Form)
    Dim cCont As Control
    For Each cCont In frm.Controls
        If TypeName(cCont) = "TextBox" Or TypeName(cCont) = "ComboBox" Then
            cCont = vbNullString
        ElseIf TypeName(cCont) = "CheckBox" Then
            cCont = 0
        ElseIf TypeName(cCont) = "ListBox" Then
            If cCont.RowSourceType = "Value List" Then
                RemoveAllListBoxItems cCont
            End If
        End If
    Next cCont
End Sub

'Remove Selected List Box Items
Public Sub RemoveSelectedListBoxItems(ByRef objLst As Object)
    Dim strRemoveList As String
    Dim astrRemoveItem() As String
    Dim varItem As Variant
    Dim lngi As Long

    With objLst
        For Each varItem In .ItemsSelected
            strRemoveList = strRemoveList & Chr(0) & .ItemData(varItem)
        Next varItem
    
        If Len(strRemoveList) > 0 Then
            astrRemoveItem = Split(Mid$(strRemoveList, 2), Chr(0))
    
        For lngi = LBound(astrRemoveItem) To UBound(astrRemoveItem)
            .RemoveItem astrRemoveItem(lngi)
        Next lngi
    
        End If
    End With
End Sub

'Remove All List Box Items
Public Sub RemoveAllListBoxItems(ByRef objLst As Object)
    Dim N As Integer
    With objLst
        For N = .ListCount - 1 To 0 Step -1
            .RemoveItem (N)
        Next N
    End With
End Sub


' END version 2.1 additions

'*****************************************************************************

Private Function GetScreenProfileSQL(profileName) As String
    GetScreenProfileSQL = "" & _
        "SELECT " & _
        "   P.ProfileName, " & _
        "   C.ControlName, " & _
        "   C.Top, " & _
        "   C.Left, " & _
        "   C.Width, " & _
        "   C.Height, " & _
        "   C.Visible, " & _
        "   C.UpdateVisibility " & _
        "FROM " & _
        "   SCR_ScreensProfiles AS P " & _
        "   INNER JOIN SCR_ScreensProfilesControls AS C " & _
        "       ON P.ProfileID = C.ProfileID " & _
        "WHERE " & _
        "   P.ProfileName=""" & profileName & """;"
End Function

Public Sub SuspendLayout(Optional frm As Form)
    Application.Echo False
End Sub

Public Sub ResumeLayout(Optional frm As Form)
    Application.Echo True
End Sub

Public Sub LoadScreenProfile(profileName As String, frm As Form)
' DS Mar 8, 2010 added error handler
'On Error Resume Next
On Error GoTo LoadError
    Dim rs As RecordSet
    Dim ctrl As Control
    
    Set rs = CurrentDb.OpenRecordSet(GetScreenProfileSQL(profileName), dbOpenSnapshot + dbReadOnly)
    
    While Not rs.EOF
        
        Set ctrl = frm.Controls(rs.Fields("ControlName").Value)
        
        With ctrl
            .top = rs.Fields("Top").Value
            .left = rs.Fields("Left").Value
            .Width = rs.Fields("Width").Value
            .Height = rs.Fields("Height").Value
            If rs.Fields("UpdateVisibility") Then
                .visible = rs.Fields("Visible").Value
            End If
        End With
        
        rs.MoveNext
    Wend
    
LoadErrorResume:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    
    Exit Sub
LoadError:
    MsgBox "Error " & Err.Number & " ( " & Err.Description & ")"
    Resume LoadErrorResume
End Sub
Public Sub ToggleAccessMenus(ByVal Setting As Boolean)
    Dim formViewSetting As AcShowToolbar
    Dim toolbarSetting As AcShowToolbar
    If Setting Then
        toolbarSetting = acToolbarYes
        formViewSetting = acToolbarWhereApprop
    Else
        toolbarSetting = acToolbarNo
        formViewSetting = acToolbarNo
        'formViewSetting = acToolbarWhereApprop
    End If
    
    DoCmd.ShowToolbar "Form Design", formViewSetting
    DoCmd.ShowToolbar "Formatting (Form/Report)", formViewSetting
    DoCmd.ShowToolbar "Toolbox", formViewSetting
    DoCmd.ShowToolbar "Database", formViewSetting
    DoCmd.ShowToolbar "Form View", formViewSetting
    DoCmd.ShowToolbar "Formatting (Datasheet)", formViewSetting
    

End Sub
Public Function NewID() As String
    Dim res As String, resLen As Long, GUID(15) As Byte
    res = Space$(128)
    'CoCreateGuid_Alt GUID(0)
    'resLen = StringFromGUID2_Alt(GUID(0), ByVal StrPtr(res), 128)
    ' HC changed to call centralized winapi functions
    createGUID GUID(0)
    resLen = StringFromGuidALT(GUID(0), ByVal StrPtr(res), 128)
    
    NewID = left$(res, resLen - 1)
End Function

Public Function GUIDToString(ByVal GUID As String) As String
    Dim translated As String
    Dim Result(15) As Byte
    Dim ctr As Integer
    
    If (Len(GUID) <> 8) Then
        GUIDToString = GUID
    Else
        translated = StrConv(GUID, vbUnicode)
        For ctr = 0 To 15
            Result(ctr) = AscB(Mid(translated, ctr + 1, 1))
        Next
        GUIDToString = ConvertGuid(StringFromGUID(Result))
    End If
End Function
Public Function getVersion(StDatabase As String) As Double
On Error GoTo ErrorHappened
Dim db As DAO.Database
Dim Tbl As DAO.TableDef
getVersion = -1  'NOT DETECTED BY DEFAULT
Set db = DBEngine.OpenDatabase(StDatabase)

For Each Tbl In db.TableDefs
    Select Case UCase(Tbl.Name)
    Case "SCR_SCREENS"
        Exit For
    Case "CCACFGSCRSCREENNAMES"
        getVersion = 0
        RaiseEvent StatusMessage("GetVersion", "Version: 0", 0)
        GoTo ExitNow
    End Select
Next Tbl

If UCase(Tbl.Name) <> "SCR_SCREENS" Then
    RaiseEvent StatusMessage("GetVersion", "No Valid Screens Tables Found in Database: " & StDatabase, 8)
    RaiseEvent StatusMessage("GetVersion", "Get Version Failed", 10)
    getVersion = -1
    GoTo ExitNow
End If

getVersion = Tbl.Properties("Description")
RaiseEvent StatusMessage("GetVersion", "Version:" & Tbl.Properties("Description"), 0)

ExitNow:
    On Error Resume Next
    Set Tbl = Nothing
    Set db = Nothing
    Exit Function

ErrorHappened:
    RaiseEvent StatusMessage("GetVersion", "Error Getting Database Version for :" & StDatabase, 10)
    RaiseEvent StatusMessage("GetVersion", Err.Description, 10)
    getVersion = -1
    Resume ExitNow

End Function

Public Sub CreatePK(ByVal TableName As String, ByVal Fields As String)
On Error Resume Next
    CurrentDb.Execute "CREATE UNIQUE INDEX PK_" & TableName & " On " & TableName & "(" & Fields & ")"
End Sub
Public Function EscapeChars(what As String, Optional ByVal Qte As String = "'") As String
On Error GoTo ErrorHandler

    Dim Result As String
    Result = ""
    
    If Nz(what, "") <> "" Then
        Result = what
        If (Qte = "'") Then
            Result = Replace(Result, "'", "''") ' change single quote to two single quote - SQL escaping
        End If
        Result = Replace(Result, Chr(13), " ") 'remove carriage return
        Result = Replace(Result, Chr(10), "") 'remove carriage return
        If (Qte = Chr(34)) Then
            Result = Replace(Result, Chr(34), Chr(34) & Chr(34)) ' Escape double-quotes
        End If
    End If
    EscapeChars = Result

Exit_ErrorHandler:
    Exit Function
ErrorHandler:
    MsgBox Err.Description
    On Error Resume Next
    Resume Exit_ErrorHandler
End Function
Public Function EscapeQuotes(what As String, Optional ByVal Qte As String = "'") As String
On Error GoTo ErrorHandler

    Dim Result As String
    Result = ""
    
    If Nz(what, "") <> "" Then
        Result = what
        If (Qte = "'") Then
            Result = Replace(Result, "'", "''") ' change single quote to two single quote - SQL escaping
        End If
        If (Qte = Chr(34)) Then
            Result = Replace(Result, Chr(34), Chr(34) & Chr(34)) ' Escape double-quotes
        End If
    End If
    EscapeQuotes = Result

Exit_ErrorHandler:
    Exit Function
ErrorHandler:
    MsgBox Err.Description
    On Error Resume Next
    Resume Exit_ErrorHandler
End Function

Public Function GetGridLabel(ByRef theForm As Form, ByVal theName As String) As TextBox
    Dim Result As TextBox
    Dim ictr As Integer
    
    Set Result = Nothing
    For ictr = 0 To theForm.Controls.Count - 1
        If (theForm.Controls(ictr).Name = theName) Then
            Set Result = theForm.Controls(ictr)
            Exit For
        End If
    Next ictr
    Set GetGridLabel = Result
End Function
Public Function IsSubForm(pFrm As Form) As Boolean
On Error Resume Next
    Dim strName As String

    strName = pFrm.Parent.Name
    IsSubForm = (Err = 0)
End Function

'resize column with extra paramter to hide first column --usually ID col
'MN - 7/28/2009 -- added extra optional column to hide first column
'KB - 6/17/2011 -- fixed the colWidth when hiding more than one column.
Public Sub ResizeGridColumns(ByRef detailForm As SubForm, ByVal NumCols As Integer, _
    Optional ByVal HideFirstCol As Boolean = False, _
    Optional ByVal startAt As Integer = 1, _
    Optional ByVal HideMultiple As Boolean = False)
    
    Dim colWidth As Integer
    Dim iCol As Integer
    ' need this as a parm so we can start after 2nd col
    'Dim startAt As Integer
    
    If HideFirstCol Then
        If startAt = 1 Then startAt = 2 ' backward-compatiblity
        colWidth = CInt((detailForm.Width) / (NumCols - IIf(HideMultiple, (startAt - 1), 1)))
        For iCol = 1 To startAt - 1
            If (detailForm.Form.Controls("Field" & CStr(iCol)).ControlSource <> "") Then
                detailForm.Form.Controls("Field" & CStr(iCol)).ColumnWidth = 0
            Else
                Exit For
            End If
        Next iCol
    Else
        colWidth = CInt((detailForm.Width) / NumCols)
        'startAt = 1
    End If
    For iCol = startAt To NumCols
        If (detailForm.Form.Controls("Field" & CStr(iCol)).ControlSource <> "") Then
            detailForm.Form.Controls("Field" & CStr(iCol)).ColumnWidth = colWidth
        Else
            Exit For
        End If
    Next iCol
    detailForm.Form.Controls("Field" & CStr(NumCols)).ColumnWidth = colWidth - 300
End Sub

Public Sub ClearMenu(ByVal menuName As String)

    Dim objCommandBar As CommandBar
    
   For Each objCommandBar In Application.CommandBars
        If objCommandBar.Name = menuName Then
            objCommandBar.Delete
        End If
    Next objCommandBar
    
End Sub
Public Function ConvertGuid(ByVal theGUID As String) As String
   Dim Result As String
   
   Result = Replace(theGUID, "guid", "")
   Result = Replace(Result, "{", "")
   Result = Replace(Result, "}", "")
   
   ConvertGuid = Trim(Result)
End Function

'JL 01/19/11 updated to work with URLs
Public Function URLEncode(ByVal what As String) As String
    
    Dim Result As String
    
    If Nz(what, vbNullString) <> vbNullString Then
        Result = Replace(what, " ", "%20")
        Result = Replace(Result, "<", "%3C")
        Result = Replace(Result, ">", "%3E")
        Result = Replace(Result, "#", "%23")
        Result = Replace(Result, "%", "%25")
        Result = Replace(Result, "&", "%26")
        Result = Replace(Result, Chr(34), "%22")
        Result = Replace(Result, Chr(39), "%27")
    End If
    
    URLEncode = Result
End Function


Public Function XMLEncode(ByVal what As String, Optional ByVal preserveCrLf As Boolean = False) As String
    Dim Result As String
    If Nz(what, vbNullString) <> vbNullString Then
        Result = Replace(what, "&", "&#38;")
        Result = Replace(Result, "<", "&#60;")
        Result = Replace(Result, ">", "&#62;")
        Result = Replace(Result, Chr(34), "&#34;")
        Result = Replace(Result, Chr(39), "&#39;")
        
        If preserveCrLf = False Then 'added due to functionality issues in Claim Plus
          Result = Replace(Result, vbCr, "&#13;")
          Result = Replace(Result, vbLf, "&#10;")
        End If
        
      
    End If
    XMLEncode = Result
End Function



Public Function GetRegionalShortDateFormat() As String
On Error Resume Next
    Dim Result As String
    Dim reg As CT_ClsReg
    Set reg = New CT_ClsReg
    Result = reg.GetRegValueStr(HKEY_CURRENT_USER, "Control Panel\International", "sShortDate")
    Set reg = Nothing

    GetRegionalShortDateFormat = Result
End Function

Public Sub BuildADOCommand(ByRef oCmd As Object, ByVal commandType As Long, ByVal CommandText As String)
    Set oCmd = CreateObject("ADODB.Command")
    
    oCmd.commandType = commandType
    oCmd.CommandTimeout = 0
    oCmd.CommandText = CommandText
    
End Sub
Public Sub BuildADOParam(ByRef oCmd As Object, ByVal ParmName As String, ByVal parmType As Long, _
    ByVal ParmDirection As Long, Optional ByVal Size As Long = -1)
    Dim oParam As Object
    
'Build The Parameters Collection
    If Size = -1 Then
        Set oParam = oCmd.CreateParameter(ParmName, parmType, ParmDirection)
    Else
        Set oParam = oCmd.CreateParameter(ParmName, parmType, ParmDirection, Size)
    End If
        
    oCmd.Parameters.Append oParam
    Set oParam = Nothing

End Sub

Public Function RunADOCommand(ByRef oCmd As Object, ByVal TableName As String) As Integer
On Error GoTo ErrorHappened
    
    Dim StConnect As String
    Dim Server
    Dim Database
    Dim oConn
    Dim Result As Integer

    Server = GetLinkedServer(TableName)
    Database = GetLinkedDatabase(TableName)
          
'Set and open Connection and Command Objects
    StConnect = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;"
    StConnect = StConnect & "Data Source=" & Server & ";"
    StConnect = StConnect & "Initial Catalog=" & Database & ";"

''Set and open Connection and Command Objects
    Set oConn = CreateObject("ADODB.Connection")
    oConn.CursorLocation = 2 ' adUseServer
    oConn.Open StConnect
    oCmd.ActiveConnection = oConn

    oCmd.Execute
    Do Until oCmd.State <> 4 'adStateExecuting
        DoEvents
    Loop

    Do Until oCmd.State = 0 'adStateClosed
        DoEvents
    Loop
    Result = 0

ExitNow:

    On Error Resume Next
    oConn.Close
    Set oConn = Nothing
    RunADOCommand = Result
    Exit Function

ErrorHappened:
    'Debug.Print "Run ADO Command " & err.Description
    Result = -1
    Resume ExitNow

End Function
Public Function ReturnADORecordSet(ByRef oCmd As Object, ByRef rs As Object, ByRef oConn As Object, _
        ByVal TableName As String, Optional ByVal CloseConnection As Boolean = False) As Integer
On Error GoTo ErrorHappened
    
    Dim StConnect As String
    Dim Server
    Dim Database
    Dim Result As Integer

    Server = GetLinkedServer(TableName)
    Database = GetLinkedDatabase(TableName)
          
'Set and open Connection and Command Objects
    StConnect = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;"
    StConnect = StConnect & "Data Source=" & Server & ";"
    StConnect = StConnect & "Initial Catalog=" & Database & ";"

''Set and open Connection and Command Objects
    Set oConn = CreateObject("ADODB.Connection")
    oConn.CursorLocation = 2 ' adUseServer
    oConn.Open StConnect
    oCmd.ActiveConnection = oConn

    Set rs = oCmd.Execute
    Do Until oCmd.State <> 4 'adStateExecuting
        DoEvents
    Loop

    Do Until oCmd.State = 0 'adStateClosed
        DoEvents
    Loop

    Result = 0
ExitNow:

    On Error Resume Next
    If CloseConnection Then
        oConn.Close
        Set oConn = Nothing
    End If
    ReturnADORecordSet = Result
    Exit Function

ErrorHappened:
    'Debug.Print "Return ADO RecordSET " & err.Description
    CloseConnection = True
    Set rs = Nothing
    Result = -1
    Resume ExitNow

End Function

'DPS  4/21/2009  - function to return detailed SQL error message
'A private copy of this function exists in CFG_CfgLink Form make sure to update any changes done here
Public Function CompleteDBExecuteError(Optional ByVal errDescription As String = "", _
    Optional ByVal bStripCRLF As Boolean = True) As String
    Dim strErrMsg As String
    Dim strErrLoopMsg As String
    Dim errLoop As Error
    If DBEngine.Errors.Count > 0 Then
        'If the DBEngine Error Matches the current error, get the deatils
        If DBEngine.Errors(DBEngine.Errors.Count - 1).Number = Err.Number Then
            'Notify user of any errors that result from executing the query.
            For Each errLoop In DBEngine.Errors
                strErrLoopMsg = Replace(errLoop.Description, "[Microsoft][ODBC SQL Server Driver][SQL Server]", "")
                If bStripCRLF Then
                    strErrLoopMsg = Replace(strErrLoopMsg, vbCrLf, " ", 1, 1, vbTextCompare)
                End If
                strErrMsg = strErrMsg & " " & strErrLoopMsg
            Next errLoop
        End If
    End If
    'If there was a non DBEngine error, use specified or standard error desc/#
    If strErrMsg = "" And Err.Number <> 0 Then
        strErrMsg = IIf(errDescription = "", Err.Description & " (" & Err.Number & ")", errDescription)
    End If
    CompleteDBExecuteError = Trim(strErrMsg)
End Function


Public Function GetTblVersion(TblName As String) As Single
On Error GoTo ErrorHappened
    Dim db As DAO.Database
    Dim Tbl As DAO.TableDef
       
    Set db = CurrentDb
    Set Tbl = db.TableDefs(TblName)
    
    GetTblVersion = Tbl.Properties("Description")

ExitNow:
    On Error Resume Next
    Set Tbl = Nothing
    Set db = Nothing
    Exit Function

ErrorHappened:
    GetTblVersion = -1
    If Err.Number = 3270 Then
        RaiseEvent StatusMessage("GetVersion", TblName & " - Error retrieving table version", 10)
        RaiseEvent StatusMessage("GetVersion", TblName & " - Unable to match table version with Decipher version", 10)
        Debug.Print Err.Number
        Debug.Print Err.Description
    Else
        Debug.Print Err.Number
        Debug.Print Err.Description
        RaiseEvent StatusMessage("GetVersion", TblName & ": " & Err.Number & " - " & Err.Description, 10)
    End If
    Resume ExitNow
End Function

'DPS  7/24/2009  - function to selectively un/lock or en/disable controls based on their type
Public Sub lockControl(ByRef cCont As Control, ByVal lockControl As Boolean)
    If cCont.ControlType = acCommandButton Then
        cCont.Enabled = Not lockControl
    ElseIf cCont.ControlType = acCheckBox Or cCont.ControlType = acTextBox Or cCont.ControlType = acComboBox Or _
        cCont.ControlType = acListBox Or cCont.ControlType = acOptionButton Or cCont.ControlType = acOptionGroup Then
        'cCont.Enabled = Not LockFields
        cCont.Locked = lockControl
    ElseIf cCont.ControlType = acCustomControl Then
        ' need to hide/show control when changing enabling or else it freaks out
        cCont.visible = False
        cCont.Enabled = Not lockControl
        cCont.visible = True
    End If
End Sub
' dps 12/14/2009 validate email address format / 11Jan2010 change to use regular expression
Public Function IsEmailAddress(ByVal EmailAddress As Variant) As Boolean
On Error GoTo ErrorHandler
    Dim retval As Boolean
    Dim objRegExp As Object, ValExp As String
    
    retval = True
    
    If Nz(EmailAddress, "") = "" Then
        IsEmailAddress = False
        Exit Function
    End If
    
    EmailAddress = Trim(EmailAddress) ' trim it since spaces are not allowed
    
    ' from http://regexlib.com/REDetails.aspx?regexp_id=26
    ValExp = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
    
    'Create a regular expression object.
    Set objRegExp = CreateObject("VBScript.RegExp")
    With objRegExp
        'Set the pattern by using the Pattern property.
        .Pattern = ValExp
        'Set Case Insensitivity.
        .IgnoreCase = True
        'Set global applicability.
        .Global = True
    End With ' objRegExp
    
    retval = objRegExp.test(EmailAddress)
    
    IsEmailAddress = retval

ResumeErrorHandler:
    On Error Resume Next
    Set objRegExp = Nothing
    Exit Function
    
ErrorHandler:
    retval = False
    Resume ResumeErrorHandler
End Function
' dps 12/15/2009 validate delimited list of email addresses for format
Public Function ValidateEmailList(ByVal listEmails As String, Optional ByVal delim As String = ";") As Boolean
    Dim arEmails() As String
    Dim bRetVal As Boolean
    Dim intCounter As Integer
    
    bRetVal = True
    
    If Nz(listEmails, "") <> "" Then
        arEmails = Split(listEmails, delim)
        For intCounter = LBound(arEmails) To UBound(arEmails)
            If Not IsEmailAddress(arEmails(intCounter)) Then
                bRetVal = False
                Exit For
            End If
        Next
    End If
    
    ValidateEmailList = bRetVal
End Function

'DLC 06/22/09 Return the lowest of the two numbers
Public Function Min(ByVal v1 As Double, ByVal v2 As Double) As Double
    If v1 < v2 Then
       Min = v1
    Else
       Min = v2
    End If
End Function

'DLC 01/06/09 Return the highest of the two numbers
Public Function max(ByVal v1 As Double, ByVal v2 As Double) As Double
    If v1 < v2 Then
       max = v2
    Else
       max = v1
    End If
End Function

' DLC 05/16/11 (based on version by Karl Erickson)
Public Function QuoteWrap(ByVal Value As String) As String
    Const DQI = """"
    If left(Value, 1) = DQI And Right(Value, 1) = DQI Then       'If the value is already wrapped in quotes,
        QuoteWrap = Value
    Else
        Value = Nz(Replace(Value, DQI, DQI & DQI), vbNullString) 'Escape out existing double-quotes
        QuoteWrap = DQI & Value & DQI                            'Wrap in double-quotes if value present
    End If
End Function

' DLC 5/16/11 - Safely wrap a semicolon delimited string within double quotes
'               to prevent list boxes having issues with commas
Public Function WrapListBoxValues(ByVal values As String) As String
    Dim ReturnValue As String
    Dim pos As Integer
    ReturnValue = vbNullString
    pos = InStr(values, ";")
    Do While pos > 0
       ReturnValue = ReturnValue + QuoteWrap(Mid(values, 1, pos - 1)) + ";"
       values = Mid(values, pos + 1)
       pos = InStr(values, ";")
    Loop
    ReturnValue = ReturnValue + QuoteWrap(values)
    WrapListBoxValues = ReturnValue
End Function

Public Sub SetLstErrorWidth(ByVal ListContent As String, ByVal lb As listBox, Optional ByVal MaxWidth As Integer = 32000)
' DLC 07/22/11 - Fix for bug 7659
' This automatically sets the size of the columns and overall width of the LstError listbox.
' The calculation is estimated as the font is proportional but 70 Twips per character + 600 padding seems to give the best result.
Dim errs() As String
Dim c1Max As Integer
Dim c2Max As Integer
Dim i As Integer
errs = Split(ListContent, ";")
'Work out the longest columns in the ListBox
For i = 0 To UBound(errs) - 1 Step 2
    c1Max = max(c1Max, Len(errs(i)))
    c2Max = max(c2Max, Len(errs(i + 1)))
Next i
c1Max = c1Max * 70 + 600
c2Max = c2Max * 70 + 400
lb.Width = Min(c1Max + c2Max + 200, MaxWidth)
lb.ColumnWidths = c1Max & "; " & c2Max
End Sub
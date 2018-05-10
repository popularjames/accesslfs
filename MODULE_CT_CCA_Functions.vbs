Option Compare Database
Option Explicit


Function GetArrayFromHash(ByRef HashArray() As String, ByVal HashChar As String, InString As String) As Integer
On Error Resume Next

Dim tmpStr As String
Dim NumCols As Integer
Dim pos As Integer

pos = 1
If Len(InString) > 0 Then
    Do Until pos = 0
        If InStr(pos + Len(HashChar), InString, HashChar) <> 0 Then
            tmpStr = Mid(InString, IIf(NumCols = 0, pos, IIf(NumCols = 0, 1, pos + Len(HashChar))), IIf(NumCols = 0, InStr(pos + 1, InString, HashChar) - pos, InStr(pos + 1, InString, HashChar) - pos - Len(HashChar)))
        Else
            tmpStr = Mid(InString, IIf(NumCols = 0, 1, pos + Len(HashChar)))
        End If
        pos = InStr(IIf(NumCols = 0, 1, pos + Len(HashChar)), InString, HashChar)
        NumCols = NumCols + 1
        ReDim Preserve HashArray(NumCols)
        HashArray(NumCols) = tmpStr
    Loop
End If

GetArrayFromHash = NumCols

End Function

Public Function GetCurrentLocale() As Integer
On Error Resume Next
    Static Result As String
    Dim reg As CT_ClsReg
    If Result = "" Then
        Set reg = New CT_ClsReg
        Result = reg.GetRegValueStr(HKEY_CURRENT_USER, "Control Panel\International\Geo", "Nation")
        Set reg = Nothing
    End If
    GetCurrentLocale = CInt(Result)
End Function

Public Function IsUnitedStates() As Boolean
    On Error Resume Next
    If GetCurrentLocale = 244 Then
        IsUnitedStates = True
    Else
        IsUnitedStates = False
    End If

End Function

Public Function GetReportParam(Optional NewInstance, Optional StrQuestion, Optional strTitle) As Double
On Error Resume Next
Dim Msg As String
Dim Title As String
Static lastPct As String
Static IsNewInstance As Boolean

If IsMissing(NewInstance) = False Then
    IsNewInstance = CBool(NewInstance)
End If

'If LastPct = "" Or DateDiff("s", LastDate, Now) > 2 Then
If lastPct = "" Or IsNewInstance Then
    Msg = IIf(IsMissing(StrQuestion), "NO QUESTION PROVIDED!", CStr(StrQuestion))
    Title = IIf(IsMissing(strTitle), "NO Title", CStr(strTitle))
    Msg = InputBox(Msg, Title, 0)
    lastPct = Msg
    IsNewInstance = False
End If


GetReportParam = CDbl(IIf(lastPct = "", 0, lastPct))

End Function
Public Function GetReportParam2(Optional NewInstance, Optional StrQuestion, Optional strTitle) As Double

Dim Msg As String
Dim Title As String
Static lastPct As String
Static IsNewInstance As Boolean

If IsMissing(NewInstance) = False Then
    IsNewInstance = CBool(NewInstance)
End If

'If LastPct = "" Or DateDiff("s", LastDate, Now) > 2 Then
If lastPct = "" Or IsNewInstance Then
    Msg = IIf(IsMissing(StrQuestion), "NO QUESTION PROVIDED!", CStr(StrQuestion))
    Title = IIf(IsMissing(strTitle), "NO Title", CStr(strTitle))
    Msg = InputBox(Msg, Title, 0)
    lastPct = Msg
    IsNewInstance = False
End If

GetReportParam2 = CDbl(IIf(lastPct = "", 0, lastPct))

End Function

Public Function CleanFileName(InString) As String
On Error Resume Next
Dim OutString As String, CurChar As String * 1, CurCde As Byte
Dim pos As Long

For pos = 1 To Len(InString)
    CurChar = Mid(InString, pos, 1)
    CurCde = Asc(CurChar)
    Select Case CurCde
    Case 48 To 57, 65 To 90, 97 To 122   'Letters and numbers
        OutString = OutString & CurChar
    Case 32, 39, 40, 41, 45, 46, 95, 96 ' Acceptable other characters and space
        OutString = OutString & CurChar
    Case Else ' Do Nothing
    End Select
Next pos
CleanFileName = OutString
End Function


Public Function GetDirectoryAndFilename(ByRef FileStartLoc As Integer, ByRef FileName As String) As String
On Error Resume Next
Dim ST As Long
Dim StrHash As String

ST = 0
StrHash = IIf(left(FileName, 1) = "\" Or left(FileName, 1) = "/", left(FileName, 1), "\")
Do While InStr(ST + 1, FileName, StrHash) <> 0
    ST = InStr(ST + 1, FileName, StrHash)
Loop
If ST <> 0 Then
    GetDirectoryAndFilename = Mid(FileName, 1, ST)
    FileStartLoc = ST + 1
End If
FileName = Mid(FileName, FileStartLoc, Len(FileName) - FileStartLoc + 1)
End Function

Public Function GetIdentifier(DataType As Byte) As String
    Select Case DataType
        Case dbBoolean, dbByte, dbInteger, dbLong, dbCurrency, dbSingle, dbDouble
            GetIdentifier = ""
        Case dbDate
            GetIdentifier = "#"
        Case dbText, dbMemo
            GetIdentifier = Chr(34)
        Case dbLongBinary    ' OLE
            GetIdentifier = Chr(34)
        Case Else
            GetIdentifier = ""
    End Select
End Function

Public Function NumToText(dblVal As Double) As String
    Static Ones(0 To 9) As String
    Static Teens(0 To 9) As String
    Static Tens(0 To 9) As String
    Static Thousands(0 To 4) As String
    Static bInit As Boolean
    Dim i As Integer, bAllZeros As Boolean, bShowThousands As Boolean
    Dim strVal As String, strBuff As String, strTemp As String
    Dim nCol As Integer, nChar As Integer


    If bInit = False Then
        'Initialize array
        bInit = True
        Ones(0) = "zero"
        Ones(1) = "one"
        Ones(2) = "two"
        Ones(3) = "three"
        Ones(4) = "four"
        Ones(5) = "five"
        Ones(6) = "six"
        Ones(7) = "seven"
        Ones(8) = "eight"
        Ones(9) = "nine"
        Teens(0) = "ten"
        Teens(1) = "eleven"
        Teens(2) = "twelve"
        Teens(3) = "thirteen"
        Teens(4) = "fourteen"
        Teens(5) = "fifteen"
        Teens(6) = "sixteen"
        Teens(7) = "seventeen"
        Teens(8) = "eighteen"
        Teens(9) = "nineteen"
        Tens(0) = ""
        Tens(1) = "ten"
        Tens(2) = "twenty"
        Tens(3) = "thirty"
        Tens(4) = "forty"
        Tens(5) = "fifty"
        Tens(6) = "sixty"
        Tens(7) = "seventy"
        Tens(8) = "eighty"
        Tens(9) = "ninety"
        Thousands(0) = ""
        Thousands(1) = "thousand"   'US numbering
        Thousands(2) = "million"
        Thousands(3) = "billion"
        Thousands(4) = "trillion"
    End If
    'Trap errors
    On Error GoTo NumToTextError
    'Get fractional part
    strBuff = "and " & Format((dblVal - Int(dblVal)) * 100, "00") & "/100"
    'Convert rest to string and process each digit
    strVal = CStr(Int(dblVal))
    'Non-zero digit not yet encountered
    bAllZeros = True
    'Iterate through string
    For i = Len(strVal) To 1 Step -1
        'Get value of this digit
        nChar = val(Mid$(strVal, i, 1))
        'Get column position
        nCol = (Len(strVal) - i) + 1
        'Action depends on 1's, 10's or 100's column
        Select Case (nCol Mod 3)
            Case 1  '1's position
                bShowThousands = True
                If i = 1 Then
                    'First digit in number (last in loop)
                    strTemp = Ones(nChar) & " "
                ElseIf Mid$(strVal, i - 1, 1) = "1" Then
                    'This digit is part of "teen" number
                    strTemp = Teens(nChar) & " "
                    i = i - 1   'Skip tens position
                ElseIf nChar > 0 Then
                    'Any non-zero digit
                    strTemp = Ones(nChar) & " "
                Else
                    'This digit is zero. If digit in tens and hundreds column
                    'are also zero, don't show "thousands"
                    bShowThousands = False
                    'Test for non-zero digit in this grouping
                    If Mid$(strVal, i - 1, 1) <> "0" Then
                        bShowThousands = True
                    ElseIf i > 2 Then
                        If Mid$(strVal, i - 2, 1) <> "0" Then
                            bShowThousands = True
                        End If
                    End If
                    strTemp = ""
                End If
                'Show "thousands" if non-zero in grouping
                If bShowThousands Then
                    If nCol > 1 Then
                        strTemp = strTemp & Thousands(nCol \ 3)
                        If bAllZeros Then
                            strTemp = strTemp & " "
                        Else
                            strTemp = strTemp & ", "
                        End If
                    End If
                    'Indicate non-zero digit encountered
                    bAllZeros = False
                End If
                strBuff = strTemp & strBuff
            Case 2  '10's position
                If nChar > 0 Then
                    If Mid$(strVal, i + 1, 1) <> "0" Then
                        strBuff = Tens(nChar) & "-" & strBuff
                    Else
                        strBuff = Tens(nChar) & " " & strBuff
                    End If
                End If
            Case 0  '100's position
                If nChar > 0 Then
                    strBuff = Ones(nChar) & " hundred " & strBuff
                End If
        End Select
    Next i
    'Convert first letter to upper case
    strBuff = UCase$(left$(strBuff, 1)) & Mid$(strBuff, 2)
EndNumToText:
    'Return result
    NumToText = strBuff
    Exit Function
NumToTextError:
    strBuff = "#Error#"
    Resume EndNumToText
End Function


Public Function EscapeQuotes(strIn As String) As String
    'Converts any quote chars (") found in strIn to "escaped quotes" i.e. ("")
    On Error Resume Next
    EscapeQuotes = Replace$(strIn, Chr$(34), Chr$(34) & Chr$(34))
End Function

Public Function IsDataCenter(Optional ByVal UserName As String) As Boolean
    On Error GoTo IsDataCenter_Err
    

    Dim Result As Boolean
    Dim sComputer As String
    Dim visible As Boolean
    
    sComputer = Identity.Computer
    
    Select Case True
    Case left(sComputer, 5) = "DCWS-"
        visible = True
    Case left(sComputer, 6) = "TS-DC-"
        visible = True
    Case left(sComputer, 3) = "DC-"
        visible = True
    Case Else
        visible = False
    End Select
    
    
    
    IsDataCenter = True
    
IsDataCenter_Exit:
On Error Resume Next

    Exit Function
IsDataCenter_Err:
On Error Resume Next
    MsgBox Err.Description
    Result = False
    Resume IsDataCenter_Exit
End Function

Public Function SavedLocationGet() As String
'Returns the last "Location" linked to using "ConfigLinks".

On Error Resume Next  'If Prop does not exits then just return blank
Dim rtn As String
    
    rtn = ""
    rtn = CurrentDb.Properties("LastLinkLocation")
    SavedLocationGet = rtn
End Function
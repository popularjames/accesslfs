Option Compare Database
Option Explicit

Private Const PointSize As Byte = 8
Private Const ControlTop As Byte = 10
Private Const ControlSpacing As Single = 0.01
Private Const FontName As String = "Courier New"
' DLC - 01/15/10 - Added genUtils
Private genUtils As New CT_ClsGeneralUtilities
'SA 2/2/2012 - CR2636 Instance variable to keep track of date settings once per session
Private RegionalShortDateFormat As String

Public Function GetFieldWidth(ByVal DataType As Byte, ByVal Size As Integer, ByVal TableName As String, ByVal FieldName As String, ByVal ScreenID As Long) As Double
On Error GoTo GetFieldWidthError
    'SA 5/16/12 - Optimized code for faster screen loading
    Dim ReturnValue As Double

    ReturnValue = Nz(DLookup("FieldWidth", "SCR_ScreensFieldFormats", "ScreenID=" & ScreenID & " AND RecordSource='" & TableName & "' AND FieldName='" & Replace(FieldName, "'", "''") & "'"), -1)

    If ReturnValue < 0 Then
        Select Case DataType
            Case dbBoolean
                ReturnValue = GetFieldSize(PointSize, 4)
            Case dbByte
                ReturnValue = GetFieldSize(PointSize, 3)
            Case dbInteger
                ReturnValue = GetFieldSize(PointSize, 5)
            Case dbLong
                ReturnValue = GetFieldSize(PointSize, 9)
            Case dbCurrency
                ReturnValue = GetFieldSize(PointSize, 11)
            Case dbSingle
                ReturnValue = GetFieldSize(PointSize, 10)
            Case dbDouble, dbDecimal, dbNumeric
                ReturnValue = GetFieldSize(PointSize, 11)
            Case dbDate
                ReturnValue = GetFieldSize(PointSize, 8)
            Case dbText
                ReturnValue = GetFieldSize(PointSize, CInt(IIf(Size > 10, 10, Size)))
            Case dbLongBinary    ' OLE
                ReturnValue = 0
            Case dbMemo
                ReturnValue = GetFieldSize(PointSize, 20)
        End Select
    End If
            
GetFieldWidthExit:
    GetFieldWidth = ReturnValue
Exit Function
GetFieldWidthError:
On Error Resume Next
    If ReturnValue < 0 Then
        ReturnValue = 0
    End If
    Resume GetFieldWidthExit
End Function

Public Function GetFieldAlign(DataType As Byte, RecordSource As String, FieldName As String, ScreenID As Long) As Byte
On Error GoTo GetFieldAlignError
    'SA 5/16/12 - Optimized code for faster screen loading
    'General 0
    'Left 1
    'Center 2
    'Right 3
    'Distribute 4
    Dim FieldAlign As Byte
    FieldAlign = Nz(DLookup("Align", "SCR_ScreensFieldFormats", "ScreenID=" & ScreenID & " AND RecordSource='" & RecordSource & "' and FieldName='" & Replace(FieldName, "'", "''") & "'"), 0)

    If FieldAlign = 0 Then
        Select Case DataType
            Case dbByte, dbInteger, dbLong, dbCurrency, dbSingle, dbDouble, dbDate, dbDecimal
                FieldAlign = 3
            Case dbText, dbLongBinary, dbMemo, dbBoolean
                FieldAlign = 1
            Case Else
                FieldAlign = 1
        End Select
    End If
    
GetFieldAlignExit:
    GetFieldAlign = FieldAlign
Exit Function
GetFieldAlignError:
    On Error Resume Next
    If FieldAlign < 0 Then
        FieldAlign = 0
    End If
    Resume GetFieldAlignExit
End Function
Public Function GetFieldFormat(ByVal DataType As Byte, ByRef Decimals As Integer, ByVal RecordSource As String, ByVal FieldName As String, ByVal ScreenID As Long) As String
On Error GoTo GetFieldFormatError
    'SA 5/16/12 - Optimized code for faster screen loading
    'Left 1
    'Center 2
    'Right 3
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim TmpFormat As String
    Dim CalcFormat As String
    Dim NumDecimals As Integer
    Dim CalcDecimals As Byte
    Dim ReturnVal As String
    ReturnVal = vbNullString

    Set db = CurrentDb
    Set rs = db.OpenRecordSet("SELECT Decimals,Format FROM SCR_ScreensFieldFormats WHERE ScreenID=" & ScreenID & " AND RecordSource='" & RecordSource & "' AND FieldName='" & Replace(FieldName, "'", "''") & "'", dbOpenSnapshot, dbForwardOnly)
    If rs.recordCount > 0 Then
        NumDecimals = Nz(rs!Decimals, -1)
        TmpFormat = Nz(rs!Format, vbNullString)
    End If
    
    If NumDecimals = -1 Or LenB(TmpFormat) = 0 Then
        Select Case DataType
            Case dbDate
                If LenB(RegionalShortDateFormat) = 0 Then
                    RegionalShortDateFormat = genUtils.GetRegionalShortDateFormat()
                End If
                CalcFormat = Replace(RegionalShortDateFormat, "yyyy", "yy")  ' we only want 2 digit years for backward compatibility
            Case dbByte
                CalcFormat = vbNullString
                CalcDecimals = 0 'Auto
            Case dbLong, dbInteger
                CalcFormat = "Standard"
                CalcDecimals = 0 'Auto
            Case dbCurrency, dbSingle, dbDouble, dbDecimal, dbNumeric
                CalcFormat = "Standard"
                CalcDecimals = 2 'Auto
            ' Added HC 8/4/2008
            Case dbBoolean
                CalcFormat = "Yes/No"
            Case dbText, dbLongBinary, dbMemo, dbBoolean
                CalcFormat = vbNullString
            Case Else
                CalcFormat = vbNullString
        End Select
    End If
    If LenB(TmpFormat) > 0 Then
        ReturnVal = TmpFormat
    Else
        ReturnVal = CalcFormat
    End If
    
GetFieldFormatExit:
On Error Resume Next

    If NumDecimals <> -1 Then
        Decimals = NumDecimals
    Else
        Decimals = CalcDecimals
    End If
    
    GetFieldFormat = ReturnVal
    
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
Exit Function
GetFieldFormatError:
    On Error Resume Next
    If CalcDecimals < 0 Then
        CalcDecimals = 0
    End If
    Resume GetFieldFormatExit
End Function

Public Function GetFieldAlias(FieldName As String, RecordSource As String, ScreenID As Long) As String
On Error GoTo GetFieldAliasError
    'SA 5/16/12 - Optimized code for faster screen loading
    Dim ReturnVal As String

    ReturnVal = Nz(DLookup("Alias", "SCR_ScreensFieldFormats", "ScreenID=" & ScreenID & " AND RecordSource='" & RecordSource & "' AND FieldName='" & Replace(FieldName, "'", "''") & "'"), vbNullString)
    
GetFieldAliasExit:
    GetFieldAlias = ReturnVal
Exit Function
GetFieldAliasError:
    On Error Resume Next
    If IsNull(ReturnVal) Then
         ReturnVal = ""
    End If
    Resume GetFieldAliasExit
End Function

Public Function GetSplitFieldNameForLabel(Alias As String, FieldName As String) As String
On Error Resume Next
    'SA 5/16/12 - Optimized code for faster screen loading
    Dim strAlias As String
    Dim X As Integer
    Dim CurChar As Integer
    Dim FoundSplit As Boolean

    If LenB(Alias) > 0 Then
        strAlias = Alias
    Else
        strAlias = FieldName
    End If
    
    If Len(strAlias) < 2 Then
        GetSplitFieldNameForLabel = vbCrLf & strAlias
        Exit Function
    End If

    For X = 2 To Len(strAlias)
        CurChar = Asc(Mid(strAlias, X, 1))
        If CurChar >= 65 And CurChar <= 90 Then
            GetSplitFieldNameForLabel = Mid(strAlias, 1, X - 1) & vbCrLf & Mid(strAlias, X, Len(strAlias) - X + 1)
            FoundSplit = True
            Exit For
        End If
    Next X
    
    If Not FoundSplit Then
        GetSplitFieldNameForLabel = vbCrLf & FieldName
    End If
'Debug.Print GetSplitFieldNameForLabel
End Function


Public Function GetFieldHeight(PointSize As Byte) As Single
On Error Resume Next
Select Case PointSize
Case 6
    GetFieldHeight = 0.11
Case 7
    GetFieldHeight = 0.125
Case 8
    GetFieldHeight = 0.16
Case 9
    GetFieldHeight = 0.18
Case 10
    GetFieldHeight = 0.19
Case Else
    GetFieldHeight = 0.19
End Select

End Function

Public Function GetFieldSize(PointSize As Integer, NumChar As Integer) As Double
    On Error Resume Next
    Dim ReturnVal As Double
    ReturnVal = 0
    Select Case PointSize
    Case 6
        If NumChar = 1 Then
            ReturnVal = 0.1
        Else
            ReturnVal = NumChar * 0.5 + 0.5
        End If
    Case 7
        If NumChar = 1 Then
            ReturnVal = 0.1176
        Else
            ReturnVal = NumChar * 0.0588 + 0.1176
        End If
    Case 8
        If NumChar = 1 Then
            ReturnVal = 0.13333
        Else
            ReturnVal = NumChar * 0.06667 + 0.13333
        End If
    Case 9
        If NumChar = 1 Then
            ReturnVal = 0.15333
        Else
            ReturnVal = NumChar * 0.06667 + 0.15333
        End If
    Case 10
        If NumChar = 1 Then
            ReturnVal = 0.17333
        Else
            ReturnVal = NumChar * 0.06667 + 0.17333
        End If
    End Select
    
    GetFieldSize = ReturnVal
End Function
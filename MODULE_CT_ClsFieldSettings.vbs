Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'SA 5/16/12 - Created class to speed up screen loads. Settings are loaded for an entire grid rather than 1 field at a time

Private Const PointSize As Byte = 8
Private RegionalShortDateFormat As String
Private db As DAO.Database
Private rs As DAO.RecordSet
Private genUtils As New CT_ClsGeneralUtilities

Private Sub Class_Terminate()
On Error Resume Next
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

Public Sub LoadFieldSettings(ByVal ScreenID As Integer, ByVal RecordSource As String)
On Error GoTo ErrorHandler
    Dim SQL As String
    If TableExists("SCR_ScreensFieldFormats", CurrentDb) Then
        'Open Screens settings
        SQL = "SELECT FieldName,Alias,FieldWidth,Align,Format,Decimals " & _
            "FROM SCR_ScreensFieldFormats " & _
            "WHERE ScreenID=" & ScreenID & " AND RecordSource='" & RecordSource & "'"
    Else
        'Open empty RS
        SQL = "SELECT 1 FROM CT_InstalledApps WHERE 1<0"
    End If
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot)
ExitNow:

Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error"
    Resume ExitNow
End Sub

Public Function GetFieldWidth(ByVal FieldName As String, ByVal DataType As Byte, ByVal Size As Integer) As Double
On Error GoTo ErrorHandler
    Dim ReturnValue As Double
    
    ReturnValue = -1
    
    With rs.Clone
        If .recordCount > 0 Then
            .FindFirst "FieldName='" & FieldName & "'"
            If Not .NoMatch Then
                ReturnValue = Nz(!FieldWidth, -1)
            End If
        End If
    End With
    
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
            
ExitNow:
On Error Resume Next
    If ReturnValue < 0 Then
        ReturnValue = 0
    End If
    GetFieldWidth = ReturnValue
Exit Function
ErrorHandler:
On Error Resume Next
    Resume ExitNow
End Function

Public Function GetFieldAlias(ByVal FieldName As String) As String
On Error GoTo ErrorHandler
    Dim ReturnVal As String
    ReturnVal = vbNullString
    
    With rs.Clone
        If .recordCount > 0 Then
            .FindFirst "FieldName='" & FieldName & "'"
            If Not .NoMatch Then
                ReturnVal = Nz(!Alias, vbNullString)
            End If
        End If
    End With
ExitNow:
    GetFieldAlias = ReturnVal
Exit Function
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error"
    Resume ExitNow
End Function

Public Function GetFieldAlign(ByVal FieldName As String, ByVal DataType As Byte) As Byte
On Error GoTo ErrorHandler
    'General 0
    'Left 1
    'Center 2
    'Right 3
    'Distribute 4
    Dim FieldAlign As Byte
    
    FieldAlign = 0
    
    With rs.Clone
        If .recordCount > 0 Then
            .FindFirst "FieldName='" & FieldName & "'"
            If Not .NoMatch Then
                FieldAlign = Nz(!Align, 0)
            End If
        End If
    End With

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
    
ExitNow:
On Error Resume Next
    GetFieldAlign = FieldAlign
Exit Function
ErrorHandler:
    On Error Resume Next
    If FieldAlign < 0 Then
        FieldAlign = 0
    End If
    Resume ExitNow
End Function

Public Function GetFieldFormat(ByVal FieldName As String, ByVal DataType As Byte, ByRef Decimals As Integer) As String
On Error GoTo ErrorHandler
    'SA 5/16/12 - Optimized code for faster screen loading
    'Left 1
    'Center 2
    'Right 3
    Dim TmpFormat As String
    Dim CalcFormat As String
    Dim NumDecimals As Integer
    Dim CalcDecimals As Byte
    Dim ReturnVal As String
    ReturnVal = vbNullString
    
    NumDecimals = -1
    TmpFormat = vbNullString
            
    With rs.Clone
        If .recordCount > 0 Then
            .FindFirst "FieldName='" & FieldName & "'"
            If Not .NoMatch Then
                NumDecimals = Nz(!Decimals, -1)
                TmpFormat = Nz(!Format, vbNullString)
            End If
        End If
    End With
    
    If NumDecimals = -1 Or LenB(TmpFormat) = 0 Then
        Select Case DataType
            Case dbDate
                If LenB(RegionalShortDateFormat) = 0 Then
                    RegionalShortDateFormat = genUtils.GetRegionalShortDateFormat()
                    RegionalShortDateFormat = Replace(RegionalShortDateFormat, "yyyy", "yy")  ' we only want 2 digit years for backward compatibility
                End If
                CalcFormat = RegionalShortDateFormat
            Case dbByte
                CalcFormat = vbNullString
                CalcDecimals = 0
            Case dbLong, dbInteger
                CalcFormat = "Standard"
                CalcDecimals = 0
            Case dbCurrency, dbSingle, dbDouble, dbDecimal, dbNumeric
                CalcFormat = "Standard"
                CalcDecimals = 2
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
    
ExitNow:
On Error Resume Next
    If NumDecimals <> -1 Then
        Decimals = NumDecimals
    Else
        Decimals = CalcDecimals
    End If
    GetFieldFormat = ReturnVal
Exit Function
ErrorHandler:
    Resume ExitNow
End Function
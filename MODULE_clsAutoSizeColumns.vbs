Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const TwipsPerInch = 1440
Private Const MouseNormal = 0   '(Default) The shape is determined by Microsoft Access
Private Const MouseArrow = 1
Private Const MouseIBeam = 3
Private Const MouseVerticalResize = 7 ' (Size N, S)
Private Const MouseHorizontalResize = 9 '  Horizontal Resize (Size E, W)
Private Const MouseBusy = 111 ' Busy (Hourglass)

Private Type Size
        cx As Long
        cy As Long
End Type

Private Const LF_FACESIZE = 32

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE
End Type

Private Declare Function apiCreateFontIndirect Lib "gdi32" Alias _
        "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function apiSelectObject Lib "gdi32" _
 Alias "SelectObject" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal hdc As Long) As Long

Private Declare Function apiGetDC Lib "user32" _
  Alias "GetDC" (ByVal hwnd As Long) As Long

Private Declare Function apiReleaseDC Lib "user32" _
  Alias "ReleaseDC" (ByVal hwnd As Long, _
  ByVal hdc As Long) As Long

Private Declare Function apiDeleteObject Lib "gdi32" _
  Alias "DeleteObject" (ByVal hObject As Long) As Long

Private Declare Function apiGetTextExtentPoint32 Lib "gdi32" _
Alias "GetTextExtentPoint32A" _
(ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, _
lpSize As Size) As Long

' Create an Information Context
 Private Declare Function apiCreateIC Lib "gdi32" Alias "CreateICA" _
 (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
 ByVal lpOutput As String, lpInitData As Any) As Long
 
' Close an existing Device Context (or information context)
 Private Declare Function apiDeleteDC Lib "gdi32" Alias "DeleteDC" _
 (ByVal hdc As Long) As Long

 Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
 
 Private Declare Function GetDeviceCaps Lib "gdi32" _
 (ByVal hdc As Long, ByVal nIndex As Long) As Long
 
 ' Constants
 Private Const SM_CXVSCROLL = 2
 Private Const LOGPIXELSX = 88

' Array of strings used to build the ColumnWidth property
Private strWidthArray() As String

' Array of Column Widths.
' The entries are cumulative in order to
' aid matching of the start of each column
Private sngWidthArray() As Single

' Amount of extra space to add to edge of each column
Private m_ColumnMargin As Long

' ListBox/Combo we are resizing
Private m_Control As Access.Control
'


Public Sub SetControl(ctl As Access.Control)
' You must set this property from the calling Form
' in order for this Class to work properly!!!
Dim lngTemp As Long
Dim strTemp As String
Dim intTemp As Integer

' Junk Var for loops
Dim ctr As Long
    ctl.ColumnWidths = "1"
    ' Save a local reference
    Set m_Control = ctl
    
    ' If we access the ListIndex property
    ' then the entire Index for the RowSource
    ' behind each ListBox is loaded.
    ' Allows for smoother initial scrolling.
    lngTemp = m_Control.ListCount
        
    ' Check and see if there is only one entry
    ' for the ColumnWidth property. This would
    ' signify the value is to be repeated for all Columns.
    ' The delimineter is the ";" character
    strTemp = m_Control.ColumnWidths
    intTemp = Split(strWidthArray(), strTemp, ";")
    ' If only one entry then we must redim the array
    ' to hold values for all columns and copy this
    ' value into each element of the array.
    If intTemp = 0 Then
        ReDim Preserve strWidthArray(m_Control.ColumnCount - 1)
        For lngTemp = 1 To UBound(strWidthArray)
            strWidthArray(lngTemp) = strWidthArray(0)
        Next
    End If
    
    ' Build cumulative ColumnWidth positions
    ' Size sngWidthArray to match strWidthArray
    ReDim sngWidthArray(UBound(strWidthArray))
    
    For lngTemp = 0 To UBound(strWidthArray)
'        Debug.Assert lngTemp = 99
'        Debug.Assert lngTemp <> 99
        For ctr = 0 To lngTemp
'            Debug.Assert ctr <> 99
            sngWidthArray(lngTemp) = sngWidthArray(lngTemp) + CSng(strWidthArray(ctr))
        Next ctr

    Next lngTemp

End Sub

' 20130417: KD: Fixed the issue where the property value is too long
' and also properly indented code!! grrrrr
Public Sub AutoSize()
On Error Resume Next
' Junk vars
Dim lngRet As Long
Dim ctr As Long
Dim strTemp As String
Dim lngWidth As Long

' Temp array to hold calculated Column Width
Dim lngArray() As Long
' Temp array to hold calculated Column Widths
Dim strArray() As String
    
    Call TurnOffDeveloperErrorHandling(True)
    ReDim lngArray(UBound(sngWidthArray))
    ReDim strArray(UBound(sngWidthArray))
    
    For ctr = 0 To m_Control.ColumnCount - 1
        lngArray(ctr) = GetColumnMaxWidth(m_Control, ctr) + m_ColumnMargin
       ' MsgBox (m_Control.column(ctr))
    '    ctl.ColumnWidths = Nz(ctl.ColumnWidths, "") & lngArray(ctr) & ";"
    Next ctr
    
    ' Build the ColumnWidths property
    For ctr = 0 To UBound(lngArray)
        ' Init var
        lngWidth = lngArray(ctr)
            
        If ctr <> UBound(strArray) Then
            strArray(ctr) = lngWidth + 100 & ";"
        Else
            strArray(ctr) = lngWidth + 100
        End If
    Next ctr

' Build ColumnWidths property
    strTemp = ""
    m_Control.ColumnWidths = CStr(m_Control.Width + 200) & ";"
    For ctr = 0 To UBound(strArray)
'        Debug.Assert ctr <> 99
        strTemp = strTemp & strArray(ctr)
    
        'access can't handle more than 58 some rows with values in columnwidths, so if there are more than this they get set to default width.
        Call TurnOffDeveloperErrorHandling(True)
        On Error Resume Next
        ' 20130417: KD: Ok, well.... This needs to be more dynamic so...
        If ctr = 0 Then
            m_Control.ColumnWidths = strArray(ctr)
        Else
            m_Control.ColumnWidths = m_Control.ColumnWidths & "; " & strArray(ctr)
        End If
        If Err.Number = 2176 Then   ' The setting for this property is too long.
                '            m_Control.ColumnCount = ctr ' bottom line, we need to stop here, too many columns - actually, doing this
                ' kills the form's recordset
            m_Control.ColumnWidths = m_Control.ColumnWidths & ";"
            ' seem to show up
            Err.Clear
            On Error GoTo 0
            GoTo DoneResizing
        End If
        
    Next
DoneResizing:
    Call TurnOffDeveloperErrorHandling(False)

    'remove the first semicolon
    If left(m_Control.ColumnWidths, 1) = ";" Then
        m_Control.ColumnWidths = Right(m_Control.ColumnWidths, Len(m_Control.ColumnWidths) - 1)
    End If
End Sub



Private Function UpdateColumnWidthProp()
' Build a new ColumnWidth property from our
' array of singles.
Dim strTemp As String
Dim lngTemp As Long
Dim sngTemp As Single
Dim ctr As Long
Dim blBusy As Boolean

On Error Resume Next

    If blBusy = True Then Exit Function
    
    blBusy = True
    ' Build the ColumnWidths property
    For lngTemp = UBound(sngWidthArray) To 0 Step -1
        ' Init var
        sngTemp = sngWidthArray(lngTemp)
        If lngTemp > 0 Then sngTemp = sngTemp - sngWidthArray(lngTemp - 1)
        
        If lngTemp <> UBound(strWidthArray) Then
            strWidthArray(lngTemp) = sngTemp & ";"
        Else
            strWidthArray(lngTemp) = sngTemp
        End If
    
    Next lngTemp

    ' Build ColumnWidths property
    strTemp = ""
    For lngTemp = 0 To UBound(strWidthArray)
        strTemp = strTemp & strWidthArray(lngTemp)
    Next
    
    lngTemp = StrComp(strTemp, m_Control.ColumnWidths, 0)
    ' Only update if there is a change from the current settings
    If lngTemp <> 0 Then m_Control.ColumnWidths = strTemp
    
    ' Clear our Busy Flag
    blBusy = False

End Function

Private Function Split(ArrayReturn() As String, ByVal StringToSplit As String, _
 SplitAt As String) As Integer
' Copyright Terry Kreft
Dim intInstr As Integer
Dim intCount As Long
Dim strTemp As String

   intCount = -1
   intInstr = InStr(StringToSplit, SplitAt)
   Do While intInstr > 0
     intCount = intCount + 1
     ReDim Preserve ArrayReturn(0 To intCount)
     ArrayReturn(intCount) = left(StringToSplit, intInstr - 1)
     StringToSplit = Mid(StringToSplit, intInstr + 1)
     intInstr = InStr(StringToSplit, SplitAt)
   Loop
   
   
   If Len(StringToSplit) > 0 Then
     intCount = intCount + 1
     ReDim Preserve ArrayReturn(0 To intCount)
     ArrayReturn(intCount) = StringToSplit
   End If
   Split = intCount
 End Function


Private Function StringToTwips(ctl As Control, strText As String) As Long
Dim myfont As LOGFONT
Dim stfSize As Size
Dim lngLength As Long
Dim lngRet As Long
Dim hdc As Long
Dim lngscreenXdpi As Long
Dim fontsize As Long
Dim hfont As Long, prevhfont As Long
    
    ' Get Desktop's Device Context
    hdc = apiGetDC(0&)
    
    'Get Current Screen Twips per Pixel
    lngscreenXdpi = GetDPI()
    
    ' Build our LogFont structure.
    ' This  is required to create a font matching
    ' the font selected into the Control we are passed
    ' to the main function.
    'Copy font stuff from Text Control's property sheet
    With myfont
        .lfFaceName = ctl.FontName & Chr$(0)  'Terminate with Null
        fontsize = ctl.fontsize
        .lfWeight = ctl.FontWeight
        .lfItalic = ctl.FontItalic
        .lfUnderline = ctl.FontUnderline
    
        ' Must be a negative figure for height or system will return
        ' closest match on character cell not glyph
        .lfHeight = (fontsize / 72) * -lngscreenXdpi
    End With
                                     
    ' Create our Font
    hfont = apiCreateFontIndirect(myfont)
    ' Select our Font into the Device Context
    prevhfont = apiSelectObject(hdc, hfont)
                
    ' Let's get length and height of output string
    lngLength = Len(strText)
    lngRet = apiGetTextExtentPoint32(hdc, strText, lngLength, stfSize)
    
    ' Select original Font back into DC
    hfont = apiSelectObject(hdc, prevhfont)
    
    ' Delete Font we created
    lngRet = apiDeleteObject(hfont)
        
    ' Release the DC
    lngRet = apiReleaseDC(0&, hdc)
        
    ' Return the Height of the String in Twips
    StringToTwips = stfSize.cy * (1440 / GetDPI())
        
End Function


Private Function GetDPI() As Integer

    ' Determine how many Twips make up 1 Pixel
    ' based on current screen resolution
    
    Dim lngIC As Long
    lngIC = apiCreateIC("DISPLAY", vbNullString, _
     vbNullString, vbNullString)
    
    ' If the call to CreateIC didn't fail, then get the info.
    If lngIC <> 0 Then
        GetDPI = GetDeviceCaps(lngIC, LOGPIXELSX)
        ' Release the information context.
        apiDeleteDC lngIC
    Else
        ' Something has gone wrong. Assume a standard value.
        GetDPI = 96
    End If
 End Function



Private Function GetColumnMaxWidth(ctl As Control, col As Long) As Long
' Loop through passed Column and calculate the
' width of the largest string for all rows of this column.

    ' Junk var
    Dim lngRet As Long
    
    ' Create our Font
    Dim myfont As LOGFONT
    Dim lngscreenXdpi As Long
    Dim fontsize As Long
    Dim hfont As Long, prevhfont As Long
    Dim hdc As Long
    Dim hDC2 As Long
    
    ' Calc size of the string
    Dim strText As String
    Dim lngLength As Long
    Dim stfSize As Size
    
    ' Loop through the rows of the ctl
    Dim ctr As Long
    Dim MaxWidth As Long
    
    ' Get Desktop's Device Context
    hDC2 = apiGetDC(0&)
    ' Create a compatible DC
    hdc = CreateCompatibleDC(hDC2)
    
    ' Release the handle to the Desktop DC
    lngRet = apiReleaseDC(0&, hDC2)
    
    'Get Current Screen Twips per Pixel
    lngscreenXdpi = GetDPI()
    
    ' Build our LogFont structure.
    ' This  is required to create a font matching
    ' the font selected into the Control we are passed
    ' to the main function.
    'Copy font stuff from Control's property sheet
    With myfont
        .lfFaceName = ctl.FontName & Chr$(0)  'Terminate with Null
        fontsize = ctl.fontsize
        .lfWeight = ctl.FontWeight
        .lfItalic = ctl.FontItalic
        .lfUnderline = ctl.FontUnderline
    
        ' Must be a negative figure for height or system will return
        ' closest match on character cell not glyph
        .lfHeight = (fontsize / 72) * -lngscreenXdpi
    End With
                                     
    ' Create our Font
    hfont = apiCreateFontIndirect(myfont)
    ' Select our Font into the Device Context
    prevhfont = apiSelectObject(hdc, hfont)
                
    ' Loop through all of the rows in the ListBox
    ' for the given Column(col) and row(ctr)

    ' Reset our max width var
    MaxWidth = 0
    
'setup to make this handle empty controls. ' KD Comeback and fix this!!! 20130325
    Dim i As Long
    If (ctl.ListCount = 0) Then
        i = 1
    Else
        i = ctl.ListCount
    End If
    
    
    For ctr = 0 To i - 1
    'For ctr = 0 To ctl.ListCount - 1
        strText = ctl.Column(col, ctr)
                       
        ' Let's get the width of output string
        lngLength = Len(strText)
        lngRet = apiGetTextExtentPoint32(hdc, strText, lngLength, stfSize)
    
        ' Now compare with last result and save larger value
        If stfSize.cx > MaxWidth Then MaxWidth = stfSize.cx
    Next ctr
    
    ' Select original Font back into DC
    hfont = apiSelectObject(hdc, prevhfont)
    
    ' Delete Font we created
    lngRet = apiDeleteObject(hfont)
        
    ' Release the DC
    lngRet = apiDeleteDC(hdc)
        
    ' Return the Height of the String in Twips
    GetColumnMaxWidth = MaxWidth * (1440 / GetDPI())
'    strText = ctl.column(col, 0)
'ctl.ColumnWidths = "123;123"
'    MsgBox (ctl.column(col, 0).Value)
  'ctl.ColumnWidths = Nz(ctl.ColumnWidths, "") & GetColumnMaxWidth & ";"
        
End Function


Public Property Let ColumnMargin(m As Long)
    ' This is TWIPS
    m_ColumnMargin = m
End Property

Public Property Get ColumnMargin() As Long
    ColumnMargin = m_ColumnMargin
End Property


Private Sub Class_Terminate()
    ' Release our reference
    Set m_Control = Nothing
End Sub

Private Sub Class_Initialize()
    ' Add a couple of pixels to allow
    ' for a margin at column edges
    m_ColumnMargin = (TwipsPerInch / GetDPI()) * 6
End Sub
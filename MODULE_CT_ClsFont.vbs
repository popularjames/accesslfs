Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'************  Code Start  ***********
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Private Const LF_FACESIZE = 32

Private Const FW_BOLD = 700

Private Const CF_APPLY = &H200&
Private Const CF_ANSIONLY = &H400&
Private Const CF_TTONLY = &H40000
Private Const CF_EFFECTS = &H100&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOSIZESEL = &H200000
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOVERTFONTS = &H1000000
Private Const CF_OEMTEXT = 7
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_SCREENFONTS = &H1
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_SHOWHELP = &H4&
Private Const CF_USESTYLE = &H80&
Private Const CF_WYSIWYG = &H8000
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS

Private Const LOGPIXELSY = 90

Private Type FormFontInfo
  Name As String
  Weight As Integer
  Height As Integer
  UnderLine As Boolean
  Italic As Boolean
  color As Long
End Type

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
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type FONTSTRUC
  lStructSize As Long
  hwnd As Long
  hdc As Long
  lpLogFont As Long
  iPointSize As Long
  flags As Long
  rgbColors As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
  hInstance As Long
  lpszStyle As String
  nFontType As Integer
  MISSING_ALIGNMENT As Integer
  nSizeMin As Long
  nSizeMax As Long
End Type

' HC 5/2010 - left the declare for the font in the ClsFont Class
Private Declare PtrSafe Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As FONTSTRUC) As Long

'MODULE VARIABLES
Private MvFontInfo As FormFontInfo
Private MvShowEffects As Boolean
Private MvShowScripts As Boolean
Private MvShowSize As Boolean
Property Let ShowSize(data As Boolean)
    MvShowSize = data
End Property
Property Get ShowSize() As Boolean
    ShowSize = MvShowSize
End Property

Property Let ShowScripts(data As Boolean)
    MvShowScripts = data
End Property
Property Get ShowScripts() As Boolean
    ShowScripts = MvShowScripts
End Property

Property Let ShowEffects(data As Boolean)
    MvShowEffects = data
End Property
Property Get ShowEffects() As Boolean
    ShowEffects = MvShowEffects
End Property


Property Let Name(data As String)
    MvFontInfo.Name = data
End Property
Property Get Name() As String
    Name = MvFontInfo.Name
End Property

Property Let Weight(data As Integer)
    MvFontInfo.Weight = data
End Property
Property Get Weight() As Integer
    Weight = MvFontInfo.Weight
End Property

Property Let Height(data As Integer)
    MvFontInfo.Height = data
End Property
Property Get Height() As Integer
    Height = MvFontInfo.Height
End Property

Property Let UnderLine(data As Boolean)
    MvFontInfo.UnderLine = data
End Property
Property Get UnderLine() As Boolean
    UnderLine = MvFontInfo.UnderLine
End Property
Property Let Italic(data As Boolean)
    MvFontInfo.Italic = data
End Property
Property Get Italic() As Boolean
    Italic = MvFontInfo.Italic
End Property

Property Let color(data As Long)
    MvFontInfo.color = data
End Property
Property Get color() As Long
    color = MvFontInfo.color
End Property


Private Function MulDiv(In1 As Long, In2 As Long, In3 As Long) As Long
Dim lngTemp As Long
  On Error GoTo MulDiv_err
  If In3 <> 0 Then
    lngTemp = In1 * In2
    lngTemp = lngTemp / In3
  Else
    lngTemp = -1
  End If
MulDiv_end:
  MulDiv = lngTemp
  Exit Function
MulDiv_err:
  lngTemp = -1
  Resume MulDiv_err
End Function
Private Function ByteToString(aBytes() As Byte) As String
  Dim dwBytePoint As Long, dwByteVal As Long, szOut As String
  dwBytePoint = LBound(aBytes)
  While dwBytePoint <= UBound(aBytes)
    dwByteVal = aBytes(dwBytePoint)
    If dwByteVal = 0 Then
      ByteToString = szOut
      Exit Function
    Else
      szOut = szOut & Chr$(dwByteVal)
    End If
    dwBytePoint = dwBytePoint + 1
  Wend
  ByteToString = szOut
End Function

Private Sub StringToByte(InString As String, ByteArray() As Byte)
Dim intLbound As Integer
  Dim intUbound As Integer
  Dim intLen As Integer
  Dim intX As Integer
  intLbound = LBound(ByteArray)
  intUbound = UBound(ByteArray)
  intLen = Len(InString)
  If intLen > intUbound - intLbound Then intLen = intUbound - intLbound
For intX = 1 To intLen
ByteArray(intX - 1 + intLbound) = Asc(Mid(InString, intX, 1))
Next
End Sub


Public Function DialogFont() As Boolean
Dim LF As LOGFONT, fs As FONTSTRUC
Dim lLogFontAddress As Long, lMemHandle As Long

LF.lfWeight = MvFontInfo.Weight
LF.lfItalic = MvFontInfo.Italic * -1
LF.lfUnderline = MvFontInfo.UnderLine * -1
LF.lfHeight = -MulDiv(CLng(MvFontInfo.Height), GetDeviceCapitals(GetDevice(hWndAccessApp), LOGPIXELSY), 72)
Call StringToByte(MvFontInfo.Name, LF.lfFaceName())
fs.rgbColors = MvFontInfo.color
fs.lStructSize = Len(fs)

lMemHandle = GlobalAlloc(GHND, Len(LF))
If lMemHandle = 0 Then
  DialogFont = False
  Exit Function
End If

lLogFontAddress = GlobalLock(lMemHandle)
If lLogFontAddress = 0 Then
  DialogFont = False
  Exit Function
End If

CopyMemory ByVal lLogFontAddress, LF, Len(LF)
fs.lpLogFont = lLogFontAddress


fs.flags = fs.flags Or CF_BOTH Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT

If MvShowEffects = True Then
  fs.flags = fs.flags Or CF_EFFECTS
End If
If MvShowScripts = False Then
  fs.flags = fs.flags Or CF_NOSCRIPTSEL
End If
If MvShowSize = False Then
    fs.flags = fs.flags Or CF_LIMITSIZE
    fs.nSizeMin = MvFontInfo.Height
    fs.nSizeMax = MvFontInfo.Height
End If

If ChooseFont(fs) = 1 Then
  CopyMemory LF, ByVal lLogFontAddress, Len(LF)
  MvFontInfo.Weight = LF.lfWeight
  MvFontInfo.Italic = CBool(LF.lfItalic)
  MvFontInfo.UnderLine = CBool(LF.lfUnderline)
  MvFontInfo.Name = ByteToString(LF.lfFaceName())
  MvFontInfo.Height = CLng(fs.iPointSize / 10)
  MvFontInfo.color = fs.rgbColors
  DialogFont = True
Else
  DialogFont = False
End If
End Function
Sub PropertiesFromControl(ctl As Control)
With MvFontInfo
      .color = ctl.ForeColor
      .Height = ctl.fontsize
      .Weight = ctl.FontWeight
      .Italic = ctl.FontItalic * -1
      .UnderLine = ctl.FontUnderline * -1
      .Name = ctl.FontName
End With
End Sub
Sub PropertiesToControl(ctl As Control)
With MvFontInfo
      ctl.ForeColor = .color
      ctl.fontsize = .Height
      ctl.FontWeight = .Weight
      ctl.FontItalic = CBool(.Italic)
      ctl.FontUnderline = CBool(.UnderLine)
      ctl.FontName = .Name
End With
End Sub

'************  Code End  ***********
Private Sub Class_Initialize()
'SET THE DEFAULTS
With MvFontInfo
      .color = 0
      .Height = 8
      .Weight = 400
      .Italic = False
      .UnderLine = False
      .Name = "Arial"
End With
MvShowSize = True
End Sub
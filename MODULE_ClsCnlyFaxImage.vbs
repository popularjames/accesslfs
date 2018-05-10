Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

''''''''''''''''''''''
' Fax resolution
''''''''''''''''''''''
Public Enum FaxRez
Regular
Fine
SuperFine
UltraFine
End Enum

''''''''''''''''''''''
' Members
''''''''''''''''''''''
Private m_oldPrinter As Object ' Printer
Private m_printerName As String
Private m_oldAppPath As String
Private m_appPath As String
Private m_oldTempDir As String
Private m_tempDir As String
Private m_fileName As String
Private m_fileGen As Long
Private m_oldFileGen As Long
Private m_oldXDpi As Integer
Private m_xDpi As Integer
Private m_oldYDpi As Integer
Private m_yDpi As Integer
Private m_oldFaxOut As Boolean
Private m_faxOut As Boolean
Private m_oldLowFaxOut As Boolean
Private m_lowFaxOut As Boolean
Private m_oldForcePrinterDpi As Boolean
Private m_forcePrinterDpi As Boolean

Private m_outputPrefix As String
Private m_destinationTag As String
Private m_destinationPath As String
Private m_faxrez As FaxRez
Private m_faxSender As ClsCnlyFax

''''''''''''''''''''''
' OutputPath property
''''''''''''''''''''''
Public Property Get OutputPath() As String
    If m_destinationTag <> "" Then
        OutputPath = m_destinationPath
    Else
        OutputPath = m_tempDir
    End If
End Property
Public Property Let OutputPath(ByVal Value As String)
    If Right(Value, 1) = "\" Then Value = left(Value, Len(Value) - 1)
    If Value = m_destinationPath Then Exit Property

    m_destinationTag = ""
    m_destinationPath = ""

    If Value <> "" Then
        If Not DoesPathExist(Value) Then Err.Raise vbObjectError + 1, "Output path does not exist"
        m_destinationTag = m_faxSender.GetDestinationTag(Value)
        If m_destinationTag <> "" Then m_destinationPath = Value
    End If

    FileName = m_fileName
End Property
''''''''''''''''''''''
' OutputPathname property
''''''''''''''''''''''
Public Property Get OutputPathname() As String
    OutputPathname = MakePathname(OutputPath, m_fileName)
End Property
''''''''''''''''''''''
' FileName property
''''''''''''''''''''''
Public Property Get FileName() As String
    FileName = m_fileName
End Property
Public Property Let FileName(ByVal Value As String)
    If DoesOutputFileAlreadyExist(Value) Then Err.Raise vbObjectError + 1, "Output file already exists: rename or delete"
    If InStr(Value, ".") < 2 Then Err.Raise vbObjectError + 1, "Output file must have extension"

    Dim DevMode As ClsCnlyDevMode
    Set DevMode = New ClsCnlyDevMode
    If Not DevMode.Com.SetImageFileName(OutputPrefix + Value, DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to set filename"
    DevMode.Save
    m_fileName = Value
End Property
''''''''''''''''''''''
' Id property
''''''''''''''''''''''
Public Property Get Id() As String
    Id = left(m_fileName, InStr(m_fileName, ".") - 1)
End Property
Public Property Let Id(ByVal Value As String)
    If Value = "" Then Value = m_faxSender.GenerateId()
    FileName = Value + ".tif"
End Property
''''''''''''''''''''''
' FaxSender property
''''''''''''''''''''''
Public Property Get Sender() As ClsCnlyFax
    Set Sender = m_faxSender
End Property
''''''''''''''''''''''
' DoesOutputFileAlreadyExist method
''''''''''''''''''''''
Public Function DoesOutputFileAlreadyExist(ByVal file As String) As Boolean
    DoesOutputFileAlreadyExist = DoesFileExist(MakePathname(OutputPath, file))
End Function
''''''''''''''''''''''
' TempDirectory property
''''''''''''''''''''''
Public Property Get TempDirectory() As String
  TempDirectory = m_tempDir
End Property
Public Property Let TempDirectory(ByVal Value As String)
    If Right(Value, 1) = "\" Then Value = left(Value, Len(Value) - 1)
    If Value = m_tempDir Then Exit Property
    If m_tempDir <> "" Then
      If Not DoesPathExist(Value) Then Err.Raise vbObjectError + 1, "Invalid temp dir"
    End If

    Dim DevMode As ClsCnlyDevMode
    Set DevMode = New ClsCnlyDevMode
    If Not DevMode.Com.SetOutputDirectory(Value, DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to set temp dir"
    DevMode.Save
    m_tempDir = Value
End Property
''''''''''''''''''''''
' ApplicationPath property
''''''''''''''''''''''
Public Property Get ApplicationPath() As String
    ApplicationPath = m_appPath
End Property
Public Property Let ApplicationPath(ByVal Value As String)
    If Value = m_appPath Then Exit Property
    If Not DoesFileExist(Value) Then Err.Raise vbObjectError + 1, "Invalid app path"

    Dim DevMode As ClsCnlyDevMode
    Set DevMode = New ClsCnlyDevMode
    If Not DevMode.Com.SetApplicationPath(Value, DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to set app path"
    DevMode.Save
    m_appPath = Value
End Property
''''''''''''''''''''''
' Resolution property
''''''''''''''''''''''
Public Property Get Resolution() As FaxRez
    Resolution = m_faxrez
End Property
Public Property Let Resolution(ByVal rez As FaxRez)
    Dim X, Y As Integer
    Select Case rez
      Case FaxRez.Regular
        X = 204
        Y = 98
      Case FaxRez.Fine
        X = 204
        Y = 196
      Case FaxRez.SuperFine
        X = 204
        Y = 391
      Case FaxRez.UltraFine
        X = 408
        Y = 391
    End Select

    IsFaxOutput = False
    HorizontalDpi = X
    VerticalDpi = Y
    m_faxrez = rez
End Property
''''''''''''''''''''''
' HorizontalDpi property
''''''''''''''''''''''
Public Property Get HorizontalDpi() As Integer
    HorizontalDpi = m_xDpi
End Property
Public Property Let HorizontalDpi(ByVal Value As Integer)
    If Value = m_xDpi Then Exit Property

    Dim DevMode As ClsCnlyDevMode
    Set DevMode = New ClsCnlyDevMode
    If Not DevMode.Com.SetXDPI(Value, DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to set xdpi"
    DevMode.Save
    m_xDpi = Value
End Property
''''''''''''''''''''''
' VerticalDpi property
''''''''''''''''''''''
Public Property Get VerticalDpi() As Integer
    VerticalDpi = m_yDpi
End Property
Public Property Let VerticalDpi(ByVal Value As Integer)
    If Value = m_yDpi Then Exit Property

    Dim DevMode As ClsCnlyDevMode
    Set DevMode = New ClsCnlyDevMode
    If Not DevMode.Com.SetYDPI(Value, DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to set ydpi"
    DevMode.Save
    m_yDpi = Value
End Property
''''''''''''''''''''''
' IsFaxOutput property
''''''''''''''''''''''
Public Property Get IsFaxOutput() As Boolean
    IsFaxOutput = m_faxOut
End Property
Public Property Let IsFaxOutput(ByVal Value As Boolean)
    If Value = m_faxOut Then Exit Property

    Dim DevMode As ClsCnlyDevMode
    Set DevMode = New ClsCnlyDevMode
    If Value Then
      If Not DevMode.Com.EnableFaxOutput(DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to enable fax output"
    Else
      If Not DevMode.Com.DisableFaxOutput(DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to disable fax output"
    End If
    DevMode.Save
    m_faxOut = Value
End Property
''''''''''''''''''''''
' IsLowFaxOutput property
''''''''''''''''''''''
Public Property Get IsLowFaxOutput() As Boolean
    IsLowFaxOutput = m_lowFaxOut
End Property
Public Property Let IsLowFaxOutput(ByVal Value As Boolean)
    If Value = m_lowFaxOut Then Exit Property

    Dim DevMode As ClsCnlyDevMode
    Set DevMode = New ClsCnlyDevMode
    If Value Then
      If Not DevMode.Com.EnableLowFaxOutput(DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to enable fax output"
    Else
      If Not DevMode.Com.DisableLowFaxOutput(DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to disable fax output"
    End If

    DevMode.Save
    m_lowFaxOut = Value
End Property
''''''''''''''''''''''
' IsForcePrinterDPI property
''''''''''''''''''''''
Public Property Get IsForcePrinterDPI() As Boolean
    IsForcePrinterDPI = m_forcePrinterDpi
End Property
Public Property Let IsForcePrinterDPI(ByVal Value As Boolean)
    If Value = m_forcePrinterDpi Then Exit Property

    Dim DevMode As ClsCnlyDevMode
    Set DevMode = New ClsCnlyDevMode
    If Value Then
      If Not DevMode.Com.EnableForcePrinterDPI(DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to enable force dpi"
    Else
      If Not DevMode.Com.DisableForcePrinterDPI(DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to disable force dpi"
    End If
    DevMode.Save
    m_forcePrinterDpi = Value
End Property
''''''''''''''''''''''
' OutputPrefix property
''''''''''''''''''''''
Public Property Get OutputPrefix() As String
    If m_destinationTag <> "" Then
        OutputPrefix = left(m_outputPrefix, Len(m_outputPrefix) - 2) & "_" & m_destinationTag & Right(m_outputPrefix, 2)
        Exit Sub
    End If

    OutputPrefix = m_outputPrefix
End Property
Public Property Let OutputPrefix(ByVal Value As String)
    m_outputPrefix = Value
    FileName = m_fileName
End Property
''''''''''''''''''''''
' PrinterName property
''''''''''''''''''''''
Public Property Get PrinterName() As String
  PrinterName = m_printerName
End Property
''''''''''''''''''''''
' Reset driver to fax settings.  Slow and requires admin rights.
''''''''''''''''''''''
Public Sub ResetDriver()
    Dim rawDevmode As Object
    Set rawDevmode = CreateObject("BLACKICEDEVMODE.BlackIceDEVMODECtrl.1")
    If Not rawDevmode.ReplaceUserSettings(m_printerName, False) Then Err.Raise vbObjectError + 1, "Cannot reset Black Ice Printer Driver"
End Sub
''''''''''''''''''''''
' Wait until done.  Call after printing to block until completion.
''''''''''''''''''''''
Public Function WaitUntilDone(ByVal milliseconds As Long) As Boolean
    WaitUntilDone = m_faxSender.WaitForFile(OutputPathname, milliseconds, False)
End Function
Private Function MakePathname(ByVal Path As String, ByVal Name As String) As String
    MakePathname = Path & "\" & Name
End Function
Private Function DoesFileExist(ByVal pathname As String)
    DoesFileExist = (Dir(pathname) <> "")
End Function
Private Function DoesPathExist(ByVal Path As String)
    DoesPathExist = (Dir(Path, vbDirectory) <> "")
End Function
Private Sub Class_Initialize()
    m_faxrez = FaxRez.Fine
    m_outputPrefix = "{In Progress}_"
    m_destinationTag = ""
    m_destinationPath = ""

    Set m_faxSender = New ClsCnlyFax
    Dim DevMode As ClsCnlyDevMode
    Set DevMode = New ClsCnlyDevMode

    ' Set default printer.
    Set m_oldPrinter = Application.Printer
    m_printerName = DevMode.PrinterName
    Set Application.Printer = Printers(m_printerName)

    ' Get initial settings so that we can restore them.
    m_oldTempDir = DevMode.Com.GetOutputDirectory(DevMode.Id)
    m_oldFileGen = DevMode.Com.GetFileGenerationMethod(DevMode.Id)
    m_oldAppPath = DevMode.Com.GetApplicationPath(DevMode.Id)
    m_appPath = m_oldAppPath
    m_xDpi = DevMode.Com.GetXDPI(DevMode.Id)
    m_oldXDpi = m_xDpi
    m_yDpi = DevMode.Com.GetYDPI(DevMode.Id)
    m_oldYDpi = m_yDpi
    m_faxOut = DevMode.Com.IsFaxOutputEnabled(DevMode.Id)
    m_oldFaxOut = m_faxOut
    m_lowFaxOut = DevMode.Com.IsLowFaxOutputEnabled(DevMode.Id)
    m_oldLowFaxOut = m_lowFaxOut

    ' Override for the duration of this instance.
    m_fileGen = 3
    If Not DevMode.Com.SetFileGenerationMethod(3, DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to set file generation method"

    DevMode.Save
    Set DevMode = Nothing

    Id = ""
    
    m_tempDir = ""
    TempDirectory = m_faxSender.TempDirectory + "\FAXES"
    m_oldTempDir = m_tempDir

    End Sub
Private Sub Class_Terminate()
    Dim DevMode As ClsCnlyDevMode
    Set DevMode = New ClsCnlyDevMode

    ' Restore default printer.
    Set Application.Printer = m_oldPrinter

    ' Restore changed settings.
    If m_oldFileGen <> -1 Then If Not DevMode.Com.SetFileGenerationMethod(-1, DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to restore file generation mode"
    DevMode.Save
    Set DevMode = Nothing

    TempDirectory = m_oldTempDir
    ApplicationPath = m_oldAppPath
    HorizontalDpi = m_oldXDpi
    VerticalDpi = m_oldYDpi
    IsFaxOutput = m_oldFaxOut
    IsLowFaxOutput = m_oldLowFaxOut

    Set m_faxSender = Nothing
End Sub

Public Property Let killClass(ByVal Value As Integer)

    Dim DevMode As ClsCnlyDevMode
    Dim m_NewFileGen As Integer
    Set DevMode = New ClsCnlyDevMode
       

    ' Restore default printer.
    Set Application.Printer = m_oldPrinter

    m_NewFileGen = Value
    
    ' Restore changed settings.
    If m_oldFileGen <> m_NewFileGen Then If Not DevMode.Com.SetFileGenerationMethod(m_NewFileGen, DevMode.Id) Then Err.Raise vbObjectError + 1, "Unable to restore file generation mode"
    DevMode.Save
    Set DevMode = Nothing

    TempDirectory = m_oldTempDir
    ApplicationPath = m_oldAppPath
    HorizontalDpi = m_oldXDpi
    VerticalDpi = m_oldYDpi
    IsFaxOutput = m_oldFaxOut
    IsLowFaxOutput = m_oldLowFaxOut

    Set m_faxSender = Nothing


End Property
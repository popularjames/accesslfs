Option Compare Database
'DLC 5/20/2010
'----- Default Folder for Decipher Dup Tool Scripts ------
Public Const DUP_TOOL_SCRIPT_FOLDER As String = "Y:\DataProcessing\Templates\Decipher\Scripts"

' HC 5/2010 - added 2 public constants for access and sql
' HC 5/2010 - updated 2010
Public Const LINK_SRC_ACCESS = "Provider=Microsoft.ACE.OLEDB.12.0;"
' HC 5/2010 - removed 2010
'Public Const LINK_SRC_ACCESS = "Provider=Microsoft.Jet.OLEDB.4.0;"
Public Const LINK_SRC_SQL = "Provider=SQLOLEDB.1;Integrated Security='SSPI';"
Public Const adOpenStatic = 3
Public Const adLockBatchOptimistic = 4


Public Const adUseServer = 2
Public Const adUseClient = 3
'---- ParameterDirectionEnum Values ----
Public Const adParamUnknown = &H0
Public Const adParamInput = &H1
Public Const adParamOutput = &H2
Public Const adParamInputOutput = &H3
Public Const adParamReturnValue = &H4

'---- CommandTypeEnum Values ----
Public Const adCmdUnknown = &H8
Public Const adCmdText = &H1
Public Const adCmdTable = &H2
Public Const adCmdStoredProc = &H4
Public Const adCmdFile = &H100
Public Const adCmdTableDirect = &H200
Public Const adCmdURLBind = &H400

Public Const adEmpty = 0
Public Const adTinyInt = 16
Public Const adSmallInt = 2
Public Const adInteger = 3
Public Const adBigInt = 20
Public Const adUnsignedTinyInt = 17
Public Const adUnsignedSmallInt = 18
Public Const adUnsignedInt = 19
Public Const adUnsignedBigInt = 21
Public Const adSingle = 4
Public Const adDouble = 5
Public Const adCurrency = 6
Public Const adDecimal = 14
Public Const adNumeric = 131
Public Const adBoolean = 11
Public Const adError = 10
Public Const adUserDefined = 132
Public Const adVariant = 12
Public Const adIDispatch = 9
Public Const adIUnknown = 13
Public Const adGUID = 72
Public Const adDate = 7
Public Const adDBDate = 133
Public Const adDBTime = 134
Public Const adDBTimeStamp = 135
Public Const adBSTR = 8
Public Const adChar = 129
Public Const adVarChar = 200
Public Const adLongVarChar = 201
Public Const adWChar = 130
Public Const adVarWChar = 202
Public Const adLongVarWChar = 203
Public Const adBinary = 128
Public Const adVarBinary = 204
Public Const adLongVarBinary = 205
Public Const adChapter = 136
Public Const adFileTime = 64
Public Const adDBFileTime = 137
Public Const adPropVariant = 138
Public Const adVarNumeric = 139
Public Const adStateClosed = &H0
Public Const adStateOpen = &H1
Public Const adStateConnecting = &H2
Public Const adStateExecuting = &H4

Public Const PSD_DEFAULTMINMARGINS = &H0 '  default (printer's)
Public Const PSD_DISABLEMARGINS = &H10
Public Const PSD_DISABLEORIENTATION = &H100
Public Const PSD_DISABLEPAGEPAINTING = &H80000
Public Const PSD_DISABLEPAPER = &H200
Public Const PSD_DISABLEPRINTER = &H20
Public Const PSD_ENABLEPAGEPAINTHOOK = &H40000
Public Const PSD_ENABLEPAGESETUPHOOK = &H2000 '  must be same as PD_*
Public Const PSD_ENABLEPAGESETUPTEMPLATE = &H8000 '  must be same as PD_*
Public Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000 '  must be same as PD_*
Public Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8 '  3rd of 4 possible
Public Const PSD_INTHOUSANDTHSOFINCHES = &H4 '  2nd of 4 possible
Public Const PSD_INWININIINTLMEASURE = &H0 '  1st of 4 possible
Public Const PSD_MARGINS = &H2 '  use caller's
Public Const PSD_MINMARGINS = &H1 '  use caller's
Public Const PSD_NOWARNING = &H80 '  must be same as PD_*
Public Const PSD_RETURNDEFAULT = &H400 '  must be same as PD_*
Public Const PSD_SHOWHELP = &H800 '  must be same as PD_*

Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Public Const CnlyAppName As String = "Decipher"

'Constants For Forms
Public Const CCAGraphItemCost As String = "SCR_PopupItemGraph"
Public Const CCAGraphItemCostMulti As String = "SCR_PopupItemGraphMulti"
Public Const CCAPriceAnalysis As String = "SCR_PopupPriceAnalysis"
Public Const CCAGraphDisc As String = "SCR_PopupDiscGraph"
Public Const CCAVenNotes As String = "SCR_PopupVendorNotes"

'Constants For Cfg Forms
Public Const CCASorts As String = "SCR_CfgSorts"
Public Const CCAFilters As String = "SCR_CfgFilters"
Public Const CCAMultiSelect As String = "SCR_CfgMultiItemSelect"
Public Const CCAText As String = "SCR_CfgText"

Public Const CSIDL_WINDOWS = &H24

' Shell constants
Public Const BIF_RETURNONLYFSDIRS = &H1&
Public Const BIF_DONTGOBELOWDOMAIN = &H2&
Public Const BIF_STATUSTEXT = &H4&
Public Const BIF_RETURNFSANCESTORS = &H8&
Public Const BIF_BROWSEFORCOMPUTER = &H1000&
Public Const BIF_BROWSEFORPRINTER = &H2000&

Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_COMMON_STARTUP = &H18
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_PRINTHOOD = &H1B
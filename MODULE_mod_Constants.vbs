Option Compare Database
Option Explicit



Public Const gbRecordingVideo As Boolean = False

Public Const gcClaimAdminName As String = "Claims Admin"    ' *NC*  2/12/2015 KD: This is only used for the Icon stuff - no need to change it unless we change the icon
                                                            ' for the new contract (not a bad idea to have a visual representation I guess, but
                                                            ' we will have 1 claim admin that can connect to old contract and new so...)

Public Const gcLocked = -1
Public Const gcAllowAdd = 1
Public Const gcAllowChange = 2
Public Const gcAllowDelete = 4
Public Const gcAllowView = 8
Public Const gcAllowReAssign = 16
Public Const gcAllowForward = 32
Public Const gcReleaseClaim = 64
Public Const gcPrintLetter = 128

Public gintAuditNum As Integer
Public gstrAuditDesc As String

Public gintAccountID As Integer
Public Const gintAuditId As Integer = 4490
Public gstrAcctAbbrev As String
Public gstrAcctDesc As String

Public gstrProfileID As String


Public Enum SecurityAction
    AllowAdd = 1
    AlllowChange = 2
    AllowDelete = 4
    AllowView = 8
    AllowReassign = 16
    AllowForward = 32
    ReleaseClaim = 64
    PrintLetter = 128
End Enum

' *jc duplicates core function
'Public Type CnlyFldDef
'    Name As String
'    Alias As String
'    ControlSrc As String
'    Type As Byte
'    Format As String
'    Width As Single
'    Height As Single
'    left As Single
'    Align As Byte
'    Decimal As Integer
'End Type

Public Enum OperationMode
    User = 1
    Manager = 2
End Enum

Public Type SQLConstruct
    Select As String
    From As String
    Where As String
    GroupBy As String
    OrderBy As String
End Type

Public Enum CnlyClaimLevel
    ClmDetail = 1
    ClmHeader = 2
End Enum

Public Const gb_VERBOSE_LOGGING As Boolean = False

Public Const gs_LOG_FILE_DIRECTORY As String = "\\ccaintranet.com\dfs-cms-ds\Data\CMS\AnalystFolders\KevinD\_CLAIM_ADMIN_LOGS\"
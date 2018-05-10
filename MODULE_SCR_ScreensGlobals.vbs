Option Compare Database
Option Explicit

'SA 8/3/2012 - Added type CnlyScreenTabsHead

Global Scr(20) As Form_SCR_MainScreens

Public Enum CnlyRecSrcType
    Table = 0
    Query = 1
End Enum

' added WhereTertiary
Public Type CnlyScreenSQL
    Select As String
    From As String
    WhereTertiary As String
    WherePrimary As String
    WhereSecondary As String
    WhereDates As String
    OrderBy As String
    SqlAll As String
    SQL As String 'it hold values Select, From, WherePrimary, WhereSecondary, and WhereTertiary, only
    SqlTotals As String
    filter As String
End Type

Public Type CnlyScreenTab
    TabID As Long
    Caption As String
    Source As String
    SourceType As CnlyRecSrcType
    LinkChild As String
    LinkMaster As String
End Type

Public Type CnlyScreenTabsHead
    ShowTab As Boolean
    Caption As String
    ControlTip As String
    StatusBar As String
    Image As String
    SubForm As String
End Type

' added types for batches
Public Type CnlyScreenCfg
    ScreenID As Long
    FormID As Byte
    ScreenName As String
    FormName As String
    PrimaryRecordSource As String
    PrimaryRecordSourceType As Byte
    CustomCriteriaListBoxRecordSource As String
    DateUse As Boolean
    StartDate As Date
    EndDate As Date
    PrimaryListBoxMulti As Boolean
    PrimaryListBoxRecordSource As String
    PrimaryListBoxRecordSourceType As Byte
    PrimaryListBoxCaption  As String
    SecondaryListBoxUse As Boolean
    SecondaryListBoxMulti As Boolean
    SecondaryListBoxDependency As Boolean
    SecondaryListBoxRecordSource As String
    SecondaryListBoxRecordSourceType As Byte
    SecondaryListBoxCaption As String
    TertiaryListBoxUse As Boolean
    TertiaryListBoxDependency As Boolean
    TertiaryListBoxMulti As Boolean
    TertiaryListBoxRecordSource As String
    TertiaryListBoxRecordSourceType As Byte
    TertiaryListBoxCaption As String
    TertiaryListBoxPrimaryDependency As Boolean
'Built From Sub RecordSets
'Added a set for Batch
    PrimaryField As String
    PrimaryFieldPos As Integer
    PrimaryQualifier As String
    PrimaryAlternatePos As Integer
    SecondaryField As String
    SecondaryFieldPos As Integer
    SecondaryQualifier As String
    SecondaryAlternatePos As Integer
    TertiaryField As String
    TertiaryFieldPos As Integer
    TertiaryQualifier As String
    TertiaryAlternatePos As Integer
    Tabs() As CnlyScreenTab
    TabsCT As Integer
    PowerBars As Boolean
    TabsHeadUser1 As CnlyScreenTabsHead
    TabsHeadUser2 As CnlyScreenTabsHead
    TabsHeadUser3 As CnlyScreenTabsHead
    'DLC 11/13/2012 - Added to support Audit/Platform selection
    Platform As String
End Type
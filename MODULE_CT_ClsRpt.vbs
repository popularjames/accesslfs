Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'THIS WAS PUT IT TO PLACE TO ALLOW THE PASSING OF THE REPORT CONFIG (LATE BOUND) TO THE EVENTS METHOD
'THIS WAS A SHORT CUT AND SHOULD BE REVISITED

' ** Added Tertiary
Public ReportName As String
Public EnableSort As Boolean
Public EnableFilter As Boolean
Public EnablePrimary  As Boolean
Public EnableSecondary  As Boolean
Public EnableTertiary As Boolean
Public EnableDate As Boolean
Public SortString As String
Public Criteria As String
Public ExtraSQL As String
Public OpenArgs As Variant
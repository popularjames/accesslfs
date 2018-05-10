Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function RecordSetToXML(SQL As String, RootNode As String, RowNode As String) As String
On Error GoTo ErrorHappened
Dim oRs As DAO.RecordSet, oDb As DAO.Database, oFld As DAO.Field
Dim oXml 'As MSXML2.DOMDocument
Dim xEl 'As MSXML2.IXMLDOMElement
Dim xRoot 'As MSXML2.IXMLDOMElement
Dim xAt 'As MSXML2.IXMLDOMAttribute


Set oDb = CurrentDb
Set oRs = oDb.OpenRecordSet(SQL, dbOpenSnapshot)
Set oXml = CreateObject("MSXML2.DOMDocument")

Set xRoot = oXml.appendChild(oXml.createElement(RootNode))

Do While Not (oRs.EOF Or oRs.BOF)
    Set xEl = xRoot.appendChild(oXml.createElement(RowNode))

    For Each oFld In oRs.Fields
        Set xAt = oXml.createAttribute(oFld.Name)
        xAt.Value = "" & oFld.Value
        xEl.Attributes.setNamedItem xAt
    Next oFld
    
    oRs.MoveNext
Loop
oRs.Close
RecordSetToXML = oXml.Xml

ExitNow:
    On Error Resume Next
    Set oRs = Nothing
    Set oDb = Nothing
    Set oXml = Nothing
    Set xEl = Nothing
    Set xAt = Nothing
    Exit Function

ErrorHappened:
    MsgBox Err.Description, vbCritical, "RecordSetToXML Error"
    Resume ExitNow
    Resume
End Function
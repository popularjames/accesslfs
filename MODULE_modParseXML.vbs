Option Explicit

Function RetrieveXMLNode(ByVal sXMLString As String, sNodeName As String, sItem As String) As String

    Dim oXMLDoc As New MSXML2.FreeThreadedDOMDocument
    Dim oXMLNodesList As IXMLDOMNodeList
    Dim oXMLNodeAttr As MSXML2.IXMLDOMAttribute
    oXMLDoc.loadXML (sXMLString)
    Set oXMLNodesList = oXMLDoc.getElementsByTagName(sNodeName)
    Dim iCntNode As Integer
    For iCntNode = 0 To oXMLNodesList.Length - 1
        For Each oXMLNodeAttr In oXMLNodesList.Item(iCntNode).Attributes
            If LCase(oXMLNodeAttr.BaseName) = sItem Then
                RetrieveXMLNode = oXMLNodeAttr.nodeTypedValue
            End If
        Next
    Next
'    Debug.Print oXMLDoc.xml
    Set oXMLDoc = Nothing
    Set oXMLNodesList = Nothing
End Function

Function RetrieveXMLNodeXPath(ByVal sXMLString As String, sSelect As String, sNodeNameMatch As String, Optional sTagname As String, Optional sLookup As String)
    Dim oXDoc As New MSXML2.DOMDocument30
    Dim oXNodes As MSXML2.IXMLDOMNodeList
    Dim oXNode As MSXML2.IXMLDOMNode
    Dim oXAttr As MSXML2.IXMLDOMAttribute
    Dim sTag As String
    Dim sValue As String
    
    Dim sCom As String
    oXDoc.SetProperty "SelectionLanguage", "XPath"
    oXDoc.loadXML (sXMLString)
'    oXDoc.loadXML (sCom)
    Set oXNodes = oXDoc.selectNodes(sSelect)
    For Each oXNode In oXNodes
        sTag = ""
        For Each oXAttr In oXNode.Attributes
            Select Case oXAttr.Name
                Case sNodeNameMatch
                    If LCase(oXAttr.Value) = LCase(sTagname) Then
                        sTag = oXAttr.Name
                    ElseIf Len(sTagname) = 0 Then
                        RetrieveXMLNodeXPath = oXAttr.Value
                        Exit For
                    End If
                Case sLookup
                    If sTag <> "" Then
                        RetrieveXMLNodeXPath = oXAttr.Value
                        Exit For
                    End If
                Case Else
            End Select
        Next
    Next
End Function


Function LoadSettings(SFileName As String) As String

    Dim oXDoc As New MSXML2.DOMDocument30

    oXDoc.Load (SFileName)

    LoadSettings = oXDoc.Xml

End Function
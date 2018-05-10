Option Compare Database
Option Explicit

Private Const ClassName As String = "modXML"


'' So, this is to send a stored proc
'' multiple parameters in any order
'' sSprocParams should be a string formatted like: Name=Value|NextName=NextValue
''
Public Function BuildXmlParams(sSprocParams As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oXml As MSXML2.DOMDocument
Dim oRoot As MSXML2.IXMLDOMElement
Dim oNode As MSXML2.IXMLDOMElement
Dim varyParams() As String
Dim iIdx As Integer
Dim sName As String
Dim sVal As String

    strProcName = ClassName & ".BuildXmlParams"
    
    varyParams = Split(sSprocParams, "|")
    
    Set oXml = New MSXML2.DOMDocument

    Set oRoot = oXml.createElement("XML")
    oXml.appendChild oRoot
    
    For iIdx = 0 To UBound(varyParams)
        sName = varyParams(iIdx)
        If Trim(sName) <> "" Then
            
            sVal = Mid(sName, InStr(1, sName, "=", vbTextCompare) + 1)
            sName = Replace(sName, "=" & sVal, "")
            
            Set oNode = oXml.createElement(sName)
            oNode.Text = sVal
            oRoot.appendChild oNode
            
'Debug.Print sName & " = " & sVal
        End If
        
    Next
    
    BuildXmlParams = oXml.Xml
    
Block_Exit:
    Set oNode = Nothing
    Set oRoot = Nothing
    Set oXml = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Private Function GetNameValuePair(sIn As String, Optional sDelimiter As String = "=", Optional sName As String, Optional sValue As String) As Boolean
On Error GoTo Block_Err
Dim saryNameVal() As String

    If InStr(1, sIn, sDelimiter, vbBinaryCompare) > 0 Then
        saryNameVal = Split(sIn, sDelimiter)
        sName = saryNameVal(0)
        sValue = saryNameVal(1)
    Else
        sName = "LastMessage"   ' out default
        sValue = sIn
    End If
    GetNameValuePair = True
Block_Exit:
    Exit Function
Block_Err:
    GetNameValuePair = False
    GoTo Block_Exit
End Function


Public Function MakeXmlString(ByVal strData As String, Optional ByVal sRootName As String = "Root", _
    Optional ByVal sFieldDelimiter As String = ";", Optional sEqualsDelimiter As String = "=", _
    Optional bForceToLowerCase As Boolean = True) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim strReturn As String
Dim oXml As MSXML2.DOMDocument
Dim oRoot As MSXML2.IXMLDOMElement
Dim oNode As MSXML2.IXMLDOMElement
Dim saryPairs() As String
Dim iIdx As Integer
Dim sField As String
Dim sVal As String


    strProcName = ClassName & ".MakeXmlString"
    If strData = "" Then GoTo Block_Exit
    
    strData = TrueTrim(strData)
    
    Set oXml = New MSXML2.DOMDocument
    Set oRoot = oXml.appendChild(oXml.createElement(sRootName))
    
    If InStr(1, strData, sFieldDelimiter, vbBinaryCompare) > 0 Then
        saryPairs = Split(strData, sFieldDelimiter)
        
        For iIdx = 0 To UBound(saryPairs)
            sVal = saryPairs(iIdx)
            If InStr(1, sVal, "=", vbBinaryCompare) > 0 Then
                Call GetNameValuePair(sVal, sEqualsDelimiter, sField, sVal)
            Else
                sField = "LastMessage"
            End If
            If bForceToLowerCase = True Then
                sField = LCase(sField)
            End If
            Set oNode = oRoot.appendChild(oXml.createElement(sField))
            oNode.Text = sVal
        Next
        
        strReturn = oXml.Xml
 
    Else
        ' Assume it's LastMessage = strNote
        Call GetNameValuePair(strData, sEqualsDelimiter, sField, sVal)
        If bForceToLowerCase = True Then
            sField = LCase(sField)
        End If
        strReturn = "<" & sRootName & "><" & sField & ">" & strData & "</" & sField & "></" & sRootName & ">"
    End If
    

    MakeXmlString = strReturn
    
Block_Exit:
    Set oNode = Nothing
    Set oRoot = Nothing
    Set oXml = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
<div align="center">

## XML\_Generator


</div>

### Description

Generate XML from ADO recordsets.
 
### More Info
 
'Set a ref to MS ADO and MSXML3.0

strParentName=name of top level node (usually the table name)

oRS = Recordset

Use as follows...

Create a procedure to connect to and retreive a recorset from a datasource.

Dim a strVariable to hold the returned xml and a boolen to check the ceration process...

dim strXML as string

Dim bOK as boolean

'Use as follows...

bOK=bGenerate_XML("tablename", oRS , strXML)

strXML = The transformed data

bGenerate_XML = Boolean

No error checking.... so there may be some


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Deltaoo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/deltaoo.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/deltaoo-xml-generator__1-31557/archive/master.zip)





### Source Code

```
' Coded by Deltaoo
'  Mail deltaoo@hotmail.com
'-------------------------------
'Use this code to convert a recordset to XML
' Use bGenerate_XML as boolean
Option Explicit
'  -- CONSTANTS --
Const XML_OPEN = "<?xml version=""1.0"" encoding=""UTF-8""?>"
Const XML_CLOSE = "" '"</xml>"
Private Function AddNode(strNodeValue As String, strNodeName As String) As String
Dim strRet     As String
  strRet = "     <" & LCase(ReplaceString(strNodeValue)) & ">"
  strRet = strRet & strNodeName & "</" & LCase(ReplaceString(strNodeValue)) & ">"
  AddNode = strRet
'
End Function
Public Function bGenerate_XML(strParentName As String, oRS As ADODB.Recordset, ByRef strXML As String) As Boolean
Dim strRet     As String
Dim n        As Integer
Dim strRootName   As String
On Error Resume Next ' Must handle the error for NULLS///
  strRootName = Trim(LCase(strParentName)) & "s"
  strParentName = LCase(strParentName)
  strRet = XML_OPEN & vbCrLf
  strRet = strRet & "<" & strRootName & ">" & vbCrLf
    With oRS
    Do Until .EOF
      strRet = strRet & "   <" & strParentName & ">" & vbCrLf
      For n = 0 To .Fields.Count - 1
      strRet = strRet & AddNode(.Fields(n).Name, .Fields(n)) & vbCrLf
      Next n
    .MoveNext
      strRet = strRet & "   </" & strParentName & ">" & vbCrLf
    Loop
    End With
  strRet = strRet & "</" & strRootName & ">" & vbCrLf
  strRet = strRet & XML_CLOSE & vbCrLf
  ' test the XML Before sending it back to the Caller
    bGenerate_XML = b_XML_OK(strRet)
    strXML = strRet
End Function
Private Function ReplaceString(strValue) As String
Dim strRet
  If IsNull(strValue) Then strValue = ""
  strRet = strValue
  strRet = Replace(strRet, "&", "&amp;")
  strRet = Replace(strRet, "<", "&lt;")
  strRet = Replace(strRet, ">", "&gt;")
  strRet = Replace(strRet, """", "&quot;")
  strRet = Replace(strRet, "'", "&apos;")
  '  -- Pass the value back --
  ReplaceString = strRet
End Function
Private Function b_XML_OK(strXMLData As String) As Boolean
Dim oDOM      As MSXML2.DOMDocument
Dim bProcOK     As Boolean
  Set oDOM = CreateObject("MSXML2.DOMDocument")
    bProcOK = oDOM.loadXML(bstrXML:=strXMLData)
    If Not bProcOK Then strXMLData = oDOM.parseError.reason
  Set oDOM = Nothing
    b_XML_OK = bProcOK
End Function
```


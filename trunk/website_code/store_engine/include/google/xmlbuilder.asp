<%
'******************************************************************************
' Copyright (C) 2006 Google Inc.
'  
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'  
'      http://www.apache.org/licenses/LICENSE-2.0
'  
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'******************************************************************************

'******************************************************************************
' Class XmlBuilder is used to build an XML string.
'******************************************************************************
Class XmlBuilder
    Public Xml
    Public Indents
    Public Stack()

	Private StackIndex

    Public Sub Class_Initialize
		StackIndex = 0
		Indents = "  "
        Xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & VbCrLf
    End Sub

    Private Sub Indent()
		Dim i, j
		j = StackIndex
		For i = 0 To j
            Xml = Xml & Indents
        Next
    End Sub

    Public Sub Push(Elem, Attributes) 
        Indent()
        Xml = Xml & "<" & Elem
		If Attributes <> "" Then
			Xml = Xml & " " & Attributes
		End If
        Xml = Xml & ">" & VbCrLf

		ReDim Preserve Stack(StackIndex)
		Stack(StackIndex) = Elem
		StackIndex = StackIndex + 1
    End Sub

    Public Sub AddElement(Elem, Content, Attributes) 
        Indent()
        Xml = Xml & "<" & Elem
		If Attributes <> "" Then
			Xml = Xml & " " & Attributes
		End If
        Xml = Xml & ">" & Server.HTMLEncode(Content) & "</" & Elem & ">" & VbCrLf
    End Sub

    Public Sub AddXmlElement(Elem, Content, Attributes) 
        Indent()
        Xml = Xml & "<" & Elem
		If Attributes <> "" Then
			Xml = Xml & " " & Attributes
		End If
        Xml = Xml & ">" & Content & "</" & Elem & ">" & VbCrLf
    End Sub

    Public Sub EmptyElement(Elem, Attributes) 
        Indent()
        Xml = Xml & "<" & Elem
		If Attributes <> "" Then
			Xml = Xml & " " & Attributes
		End If
        Xml = Xml & " />" & VbCrLf
    End Sub

    Public Sub Pop(Elem)
		StackIndex = StackIndex - 1
		Indent()
		If Elem <> Stack(StackIndex) Then
			Response.Write "XML Error: Tag Mismatch when trying to close " & Elem
		Else
			Xml = Xml & "</" & Elem & ">" & VbCrLf
		End If
    End Sub

	Function Attribute(Key, Value)
		Attribute = Key & "=""" & Value & """"
	End Function

    Public Function GetXml() 
        GetXml = Xml
    End Function

	Private Sub Class_Terminate()
		Erase Stack
	End Sub
End Class


'******************************************************************************
' The GetElementText function returns the value of an XML element.
' If there are two elements with the same tagname, beware that it only returns  
'     the value of the first element.
'
' Input:    Node       The node in which the tagname can be found
'           Tagname    Tagname of the element
'******************************************************************************
Function GetElementText(Node, Tagname)
	Dim NodeList
	Set NodeList = Node.getElementsByTagname(Tagname)
	If NodeList.Length > 0 Then 
		GetElementText = NodeList(0).text
	Else
		GetElementText = ""
	End If
	Set NodeList = Nothing
End Function


'******************************************************************************
' The GetRootNode function returns a DOM object of the root node
'
' Input:    XmlData    Given XML
'
' Returns:  DOM object of the root node
'******************************************************************************
Function GetRootNode(XmlData)
	Dim XmlDOMDoc
	Set XmlDOMDoc = Server.CreateObject("Msxml2.DOMDocument.3.0")
	XmlDOMDoc.loadXml XmlData
	Set GetRootNode = XmlDOMDoc.documentElement
	Set XmlDOMDoc = Nothing
End Function

%>
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
' Class Order is the object representation of an Order Processing API command:
'     Financial Commands:
'         <charge-order>
'         <refund-order>
'         <cancel-order>
'         <authorize-order>
'     Fulfillment Commands:
'         <process-order>
'         <add-merchant-order-number>
'         <deliver-order>
'         <add-tracking-data>
'         <send-buyer-message>
'     Archiving Commands:
'         <archive-order>
'         <unarchive-order>
'******************************************************************************
Class Order

	Private OrderCommand

    '**************************************************************************
	' Creates a <charge-order> command
	' 
	' Input:    GoogleOrderNumber    Required: <google-order-number>
	'           Amount               Optional: <amount>
    '**************************************************************************
	Public Sub ChargeOrder(GoogleOrderNumber, Amount)
		Dim Xml, Attr
		Set Xml = New XmlBuilder
		Attr = Xml.Attribute("xmlns", SchemaUri)
		Attr = Attr & " " & Xml.Attribute("google-order-number", GoogleOrderNumber)
		If Amount <> "" Then
			Xml.Push "charge-order", Attr
			Attr = Xml.Attribute("currency", MerchantCurrency)
			Xml.AddElement "amount", Amount, Attr
			Xml.Pop "charge-order"
		Else
			Xml.EmptyElement "charge-order", Attr
		End If
		OrderCommand = Xml.GetXml
		Set Xml = Nothing
	End Sub

    '**************************************************************************
	' Creates a <refund-order> command
	' 
	' Input:    GoogleOrderNumber    Required: <google-order-number>
	'           Reason               Required: <reason>
	'           Comment              Optional: <comment>
	'           Amount               Optional: <amount>
    '**************************************************************************
	Public Sub RefundOrder(GoogleOrderNumber, Reason, Comment, Amount)
		Dim Xml, Attr
		Set Xml = New XmlBuilder
		Attr = Xml.Attribute("xmlns", SchemaUri)
		Attr = Attr & " " & Xml.Attribute("google-order-number", GoogleOrderNumber)
		Xml.Push "refund-order", Attr
		Xml.AddElement "reason", Reason, ""
		If Comment <> "" Then
			Xml.AddElement "comment", Comment, ""
		End If
		If Amount <> "" Then
			Attr = Xml.Attribute("currency", MerchantCurrency)
			Xml.AddElement "amount", Amount, Attr
		End If
		Xml.Pop "refund-order"
		OrderCommand = Xml.GetXml
		Set Xml = Nothing
	End Sub

    '**************************************************************************
	' Creates a <cancel-order> command
	' 
	' Input:    GoogleOrderNumber    Required: <google-order-number>
	'           Reason               Required: <reason>
	'           Comment              Optional: <comment>
    '**************************************************************************
	Public Sub CancelOrder(GoogleOrderNumber, Reason, Comment)
		Dim Xml, Attr
		Set Xml = New XmlBuilder
		Attr = Xml.Attribute("xmlns", SchemaUri)
		Attr = Attr & " " & Xml.Attribute("google-order-number", GoogleOrderNumber)
		Xml.Push "cancel-order", Attr
		Xml.AddElement "reason", Reason, ""
		If Comment <> "" Then
			Xml.AddElement "comment", Comment, ""
		End If
		Xml.Pop "cancel-order"
		OrderCommand = Xml.GetXml
		Set Xml = Nothing
	End Sub

    '**************************************************************************
	' Creates a <authorize-order> command
	' 
	' Input:    GoogleOrderNumber    Required: <google-order-number>
    '**************************************************************************
	Public Sub AuthorizeOrder(GoogleOrderNumber)
		Dim Xml, Attr
		Set Xml = New XmlBuilder
		Attr = Xml.Attribute("xmlns", SchemaUri)
		Attr = Attr & " " & Xml.Attribute("google-order-number", GoogleOrderNumber)
		Xml.EmptyElement "authorize-order", Attr
		OrderCommand = Xml.GetXml
		Set Xml = Nothing
	End Sub

    '**************************************************************************
	' Creates a <process-order> command
	' 
	' Input:    GoogleOrderNumber    Required: <google-order-number>
    '**************************************************************************
	Public Sub ProcessOrder(GoogleOrderNumber)
		Dim Xml, Attr
		Set Xml = New XmlBuilder
		Attr = Xml.Attribute("xmlns", SchemaUri)
		Attr = Attr & " " & Xml.Attribute("google-order-number", GoogleOrderNumber)
		Xml.EmptyElement "process-order", Attr
		OrderCommand = Xml.GetXml
		Set Xml = Nothing
	End Sub

    '**************************************************************************
	' Creates a <add-merchant-order-number> command
	' 
	' Input:    GoogleOrderNumber      Required: <google-order-number>
	'           MerchantOrderNumber    Required: <merchant-order-number>
    '**************************************************************************
	Public Sub AddMerchantOrderNumber(GoogleOrderNumber, MerchantOrderNumber)
		Dim Xml, Attr
		Set Xml = New XmlBuilder
		Attr = Xml.Attribute("xmlns", SchemaUri)
		Attr = Attr & " " & Xml.Attribute("google-order-number", GoogleOrderNumber)
		Xml.Push "add-merchant-order-number", Attr
		Xml.AddElement "merchant-order-number", MerchantOrderNumber, ""
		Xml.Pop "add-merchant-order-number"
		OrderCommand = Xml.GetXml
		Set Xml = Nothing
	End Sub

    '**************************************************************************
	' Creates a <deliver-order> command
	' 
	' Input:    GoogleOrderNumber      Required: <google-order-number>
	'           Carrier                Optional: <merchant-order-number>
	'           TrackingNumber         Optional: <tracking-number>
	'           SendEmail              Optional: <send-email>; Boolean
    '**************************************************************************
	Public Sub DeliverOrder(GoogleOrderNumber, Carrier, TrackingNumber, SendEmail)
		Dim Xml, Attr
		Set Xml = New XmlBuilder
		Attr = Xml.Attribute("xmlns", SchemaUri)
		Attr = Attr & " " & Xml.Attribute("google-order-number", GoogleOrderNumber)

		If Carrier <> "" Or SendEmail Then
			Xml.Push "deliver-order", Attr
			If Carrier <> "" Then
				Xml.Push "tracking-data", ""
				Xml.AddElement "carrier", Carrier, ""
				Xml.AddElement "tracking-number", TrackingNumber, ""
				Xml.Pop "tracking-data"
			End If 
			If SendEmail Then
				Xml.AddElement "send-email", "true", ""
			End If
			Xml.Pop "deliver-order"
		Else
			Xml.EmptyElement "deliver-order", Attr
		End If
		OrderCommand = Xml.GetXml
		Set Xml = Nothing
	End Sub

    '**************************************************************************
	' Creates a <add-tracking-data> command
	' 
	' Input:    GoogleOrderNumber      Required: <google-order-number>
	'           Carrier                Required: <merchant-order-number>
	'           TrackingNumber         Required: <tracking-number>
    '**************************************************************************
	Public Sub AddTrackingData(GoogleOrderNumber, Carrier, TrackingNumber)
		Dim Xml, Attr
		Set Xml = New XmlBuilder
		Attr = Xml.Attribute("xmlns", SchemaUri)
		Attr = Attr & " " & Xml.Attribute("google-order-number", GoogleOrderNumber)
		Xml.Push "add-tracking-data", Attr
		Xml.Push "tracking-data", ""
		Xml.AddElement "carrier", Carrier, ""
		Xml.AddElement "tracking-number", TrackingNumber, ""
		Xml.Pop "tracking-data"
		Xml.Pop "add-tracking-data"
		OrderCommand = Xml.GetXml
		Set Xml = Nothing
	End Sub
	
    '**************************************************************************
	' Creates a <send-buyer-message> command
	' 
	' Input:    GoogleOrderNumber      Required: <google-order-number>
	'           Message                Required: <message>
	'           SendEmail              Optional: <send-email>; Boolean
    '**************************************************************************
	Public Sub SendBuyerMessage(GoogleOrderNumber, Message, SendEmail)
		Dim Xml, Attr
		Set Xml = New XmlBuilder
		Attr = Xml.Attribute("xmlns", SchemaUri)
		Attr = Attr & " " & Xml.Attribute("google-order-number", GoogleOrderNumber)
		Xml.Push "send-buyer-message", Attr
		Xml.AddElement "message", Message, ""
		If SendEmail Then
			Xml.AddElement "send-email", "true", ""
		End If
		Xml.Pop "send-buyer-message"
		OrderCommand = Xml.GetXml
		Set Xml = Nothing
	End Sub

    '**************************************************************************
	' Creates a <archive-order> command
	' 
	' Input:    GoogleOrderNumber    Required: <google-order-number>
    '**************************************************************************
	Public Sub ArchiveOrder(GoogleOrderNumber)
		Dim Xml, Attr
		Set Xml = New XmlBuilder
		Attr = Xml.Attribute("xmlns", SchemaUri)
		Attr = Attr & " " & Xml.Attribute("google-order-number", GoogleOrderNumber)
		Xml.EmptyElement "archive-order", Attr
		OrderCommand = Xml.GetXml
		Set Xml = Nothing
	End Sub

    '**************************************************************************
	' Creates a <unarchive-order> command
	' 
	' Input:    GoogleOrderNumber    Required: <google-order-number>
    '**************************************************************************
	Public Sub UnarchiveOrder(GoogleOrderNumber)
		Dim Xml, Attr
		Set Xml = New XmlBuilder
		Attr = Xml.Attribute("xmlns", SchemaUri)
		Attr = Attr & " " & Xml.Attribute("google-order-number", GoogleOrderNumber)
		Xml.EmptyElement "unarchive-order", Attr
		OrderCommand = Xml.GetXml
		Set Xml = Nothing
	End Sub

    '**************************************************************************
	' Sends the order processing API command to Google Checkout.
    '**************************************************************************
	Public Sub SendOrderCommand()
		Dim ResponseXml
		ResponseXml = SendRequest(OrderCommand, PostUrl)
		ProcessSynchronousResponse(ResponseXml)
	End Sub

    '**************************************************************************
	' Diagnoses the order processing command XML by posting to 
	'     the Diagnose URL.
    '**************************************************************************
	Public Sub DiagnoseXml()
		Dim ResponseXml
		ResponseXml = SendRequest(OrderCommand, DiagnoseUrl)
		ProcessSynchronousResponse(ResponseXml)
	End Sub

    '**************************************************************************
	' Processes the synchronous response received from Google Checkout 
	'     when an order processing command is issued.
    '**************************************************************************
	Public Sub ProcessSynchronousResponse(ResponseXml)
		Dim ResponseNode, RootTag
		Set ResponseNode = GetRootNode(ResponseXml)
		RootTag = ResponseNode.Tagname


                

		Select Case RootTag
			Case "request-received"
				DisplayResponse ResponseXml
			Case "error"
				DisplayResponse ResponseXml
			Case "diagnosis"
				DisplayResponse ResponseXml
			Case Else
		End Select 
	End Sub

    '**************************************************************************
	' Displays the response in the browser.
    '**************************************************************************
	Public Sub DisplayResponse(ResponseXml)
		Response.Write "<pre>"
		Response.Write Server.HTMLEncode(ResponseXml)
		Response.Write "</pre>"
	End Sub

    '**************************************************************************
	' Returns the order processing API command XML in plain text format.
    '**************************************************************************
	Public Property Get GetXml()
		GetXml = OrderCommand
	End Property

End Class

%>

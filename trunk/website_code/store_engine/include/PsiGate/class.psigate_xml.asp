<%
'
' +----------------------------------------------------------------------+
' | Copyright (c) 2005 Vincent Gibara - vgibara@cwm.ca                   |
' | Version 1.1 - 11/22/2005                                             |
' +----------------------------------------------------------------------+
' | This source file is subject to version 2.0 of the GPL license,			 |
' | that is bundled with this package in the file LICENSE, and is        |
' | available through the world-wide-web at the following url:           |
' | http://www.gnu.org/licenses/gpl.txt.                                 |
' +----------------------------------------------------------------------+
' | This code is inspired from the php class from Jason LeBaron -        |
' | jason@networkdad.com																							   |
' +----------------------------------------------------------------------+
' | Released under GPL                                                   |
' +----------------------------------------------------------------------+
' $Id: class.psigate.asp 2 22/11/2005 13:40:44Z cwm $

Const PSIGATE_TRANSACTION_OK = "APPROVED"
Const PSIGATE_TRANSACTION_DECLINED = "DECLINED"
Const PSIGATE_TRANSACTION_ERROR = "ERROR"

Class PsiGatePayment
	Private http, parser, xmlData, currentTag, xmlRequest, xmlResponse, strResult, items
	Private myGatewayURL, myStoreID, myPassphrase, myPaymentType, myCardAction, mySubtotal, myTaxTotal1, myTaxTotal2, myTaxTotal3, myTaxTotal4, myTaxTotal5, myShipTotal, myCardNumber, myCardExpMonth, myCardExpYear, myCardIDCode, myCardIDNumber, myTestResult, myOrderID, myUserID, myBname, myBcompany, myBaddress1, myBaddress2, myBcity, myBprovince, myBpostalcode, myBcountry, mySname, myScompany, mySaddress1, mySaddress2, myScity, mySprovince, mySpostalcode, myScountry, myPhone, myFax, myEmail, myComments, myCustomerIP, myShippingtotal
	Private myResultTrxnTransTime, myResultTrxnOrderID, myResultTrxnApproved, myResultTrxnReturnCode, myResultTrxnErrMsg, myResultTrxnTaxTotal, myResultTrxnShipTotal, myResultTrxnSubTotal, myResultTrxnFullTotal, myResultTrxnPaymentType, myResultTrxnCardNumber, myResultTrxnCardExpMonth, myResultTrxnCardExpYear, myResultTrxnTransRefNumber, myResultTrxnCardIDResult, myResultTrxnAVSResult, myResultTrxnCardAuthNumber, myResultTrxnCardRefNumber, myResultTrxnCardType, myResultTrxnIPResult, myResultTrxnIPCountry, myResultTrxnIPRegion, myResultTrxnIPCity
	Private myError, myErrorMessage

	'***********************************************************************
	'*** SET values to send to PsiGate                                   ***
	'***********************************************************************
	Public Property Let setGatewayURL(url)
		myGatewayURL = url
	End Property

	Public Property Let setStoreID(storeId)
		myStoreID = storeId
	End Property

	Public Property Let setPassphrase(Passphrase )
			myPassphrase = Passphrase
	End Property

	Public Property Let setPaymentType(PaymentType )
			myPaymentType = PaymentType
	End Property

	Public Property Let setCardAction(CardAction )
			myCardAction = CardAction
	End Property

	Public Property Let setSubtotal(Subtotal )
			mySubtotal = Subtotal
	End Property

	Public Property Let setTaxTotal1(TaxTotal1 )
		myTaxTotal1 = TaxTotal1
	End Property

	Public Property Let setTaxTotal2(TaxTotal2 )
		myTaxTotal2 = TaxTotal2
	End Property

	Public Property Let setTaxTotal3(TaxTotal3 )
		myTaxTotal3 = TaxTotal3
	End Property

	Public Property Let setTaxTotal4(TaxTotal4 )
		myTaxTotal4 = TaxTotal4
	End Property

	Public Property Let setTaxTotal5(TaxTotal5 )
		myTaxTotal5 = TaxTotal5
	End Property

	Public Property Let setShiptotal(Shiptotal )
		myShiptotal = Shiptotal
	End Property

	Public Property Let setCardNumber(CardNumber )
			myCardNumber = CardNumber
	End Property

	Public Property Let setCardExpMonth(CardExpMonth )
			myCardExpMonth = CardExpMonth
	End Property

	Public Property Let setCardExpYear(CardExpYear )
			myCardExpYear = CardExpYear
	End Property

	Public Property Let setCardIDCode(CardIDCode )
			myCardIDCode = CardIDCode
	End Property

	Public Property Let setCardIDNumber(CardIDNumber )
			myCardIDNumber = CardIDNumber
	End Property

	Public Property Let setTestResult(TestResult )
			myTestResult = TestResult
	End Property

	Public Property Let setOrderID(OrderID )
			myOrderID = OrderID
	End Property

	Public Property Let setUserID(UserID )
			myUserID = UserID
	End Property

	Public Property Let setBname(Bname )
			myBname = Bname
	End Property

	Public Property Let setBcompany(Bcompany )
			myBcompany = Bcompany
	End Property

	Public Property Let setBaddress1(Baddress1 )
			myBaddress1 = Baddress1
	End Property

	Public Property Let setBaddress2(Baddress2 )
			myBaddress2 = Baddress2
	End Property

	Public Property Let setBcity(Bcity )
			myBcity = Bcity
	End Property

	Public Property Let setBprovince(Bprovince )
			myBprovince = Bprovince
	End Property

	Public Property Let setBpostalcode(Bpostalcode)
		myBpostalcode = Bpostalcode
	End Property

	Public Property Let setBcountry(Bcountry)
		myBcountry = Bcountry
	End Property

	Public Property Let setSname(Sname)
		mySname = Sname
	End Property

	Public Property Let setScompany(Scompany)
		myScompany = Scompany
	End Property

	Public Property Let setSaddress1(Saddress1)
		mySaddress1 = Saddress1
	End Property

	Public Property Let setSaddress2(Saddress2)
		mySaddress2 = Saddress2
	End Property

	Public Property Let setScity(Scity)
		myScity = Scity
	End Property

	Public Property Let setSprovince(Sprovince)
		mySprovince = Sprovince
	End Property

	Public Property Let setSpostalcode(Spostalcode)
		mySpostalcode = Spostalcode
	End Property

	Public Property Let setScountry(Scountry)
		myScountry = Scountry
	End Property

	Public Property Let setPhone(Phone)
		myPhone = Phone
	End Property

	Public Property Let setFax(Fax)
		myFax = Fax
	End Property

	Public Property Let setEmail(Email)
		myEmail = Email
	End Property

	Public Property Let setComments(Comments)
		myComments = Comments
	End Property

	Public Property Let setCustomerIP(CustomerIP)
		myCustomerIP = CustomerIP
	End Property

	'***********************************************************************
	'*** GET values returned by PsiGate                                  ***
	'***********************************************************************
	Public Property Get getTrxnTransTime()
			getTrxnTransTime = myResultTrxnTransTime
	End Property

	Public Property Get getTrxnOrderID()
			getTrxnOrderID = myResultTrxnOrderID
	End Property

	Public Property Get getTrxnApproved()
			getTrxnApproved = myResultTrxnApproved
	End Property

	Public Property Get getTrxnReturnCode()
			getTrxnReturnCode = myResultTrxnReturnCode
	End Property

	Public Property Get getTrxnErrMsg()
			getTrxnErrMsg = myResultTrxnErrMsg
	End Property

	Public Property Get getTrxnTaxTotal()
			getTrxnTaxTotal = myResultTrxnTaxTotal
	End Property

	Public Property Get getTrxnShipTotal()
			getTrxnShipTotal = myResultTrxnShipTotal
	End Property

	Public Property Get getTrxnSubTotal()
			getTrxnSubTotal = myResultTrxnSubTotal
	End Property

	Public Property Get getTrxnFullTotal()
			getTrxnFullTotal = myResultTrxnFullTotal
	End Property

	Public Property Get getTrxnPaymentType()
			getTrxnPaymentType = myResultTrxnPaymentType
	End Property

	Public Property Get getTrxnCardNumber()
			getTrxnCardNumber = myResultTrxnCardNumber
	End Property

	Public Property Get getTrxnCardExpMonth()
			getTrxnCardExpMonth = myResultTrxnCardExpMonth
	End Property

	Public Property Get getTrxnCardExpYear()
			getTrxnCardExpYear = myResultTrxnCardExpYear
	End Property

	Public Property Get getTrxnTransRefNumber()
			getTrxnTransRefNumber = myResultTrxnTransRefNumber
	End Property

	Public Property Get getTrxnCardIDResult()
			getTrxnCardIDResult = myResultTrxnCardIDResult
	End Property

	Public Property Get getTrxnAVSResult()
			getTrxnAVSResult = myResultTrxnAVSResult
	End Property

	Public Property Get getTrxnCardAuthNumber()
			getTrxnCardAuthNumber = myResultTrxnCardAuthNumber
	End Property

	Public Property Get getTrxnCardRefNumber()
			getTrxnCardRefNumber = myResultTrxnCardRefNumber
	End Property

	Public Property Get getTrxnCardType()
			getTrxnCardType = myResultTrxnCardType
	End Property

	Public Property Get getTrxnIPResult()
			getTrxnIPResult = myResultTrxnIPResult
	End Property

	Public Property Get getTrxnIPCountry()
			getTrxnIPCountry = myResultTrxnIPCountry
	End Property

	Public Property Get getTrxnIPRegion()
			getTrxnIPRegion = myResultTrxnIPRegion
	End Property

	Public Property Get getTrxnIPCity()
			getTrxnIPCity = myResultTrxnIPCity
	End Property

	Private Property Get getError()
		If myError <> 0 Then
			'Internal Error
			getError = myError
		Else
			'PsiGate Error
			If getTrxnApproved() = "APPROVED" Then
				getError = PSIGATE_TRANSACTION_OK
			ElseIf getTrxnApproved() = "DECLINED" Then
				getError = PSIGATE_TRANSACTION_DECLINED
			Else
				getError = PSIGATE_TRANSACTION_ERROR
			End If
		End If
	End Property

	Private Property Get getErrorMessage()
		If myError <> 0 Then
			'Internal Error
			getErrorMessage = myErrorMessage
		Else
			'PsiGate Error
			getErrorMessage = getTrxnError
		End If
	End Property

	'***********************************************************************
	'*** Class Constructor                                               ***
	'***********************************************************************
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()	
	End Sub

	'***********************************************************************
	'*** Business Logic                                                  ***
	'***********************************************************************
	Public Sub addItem(Id, Desc, Qty, Price)
		Items = Items & _
			"<Item>" & _
				"<ItemID>"&Id&"</ItemID>" & _
				"<ItemDescription>"&Desc&"</ItemDescription>" & _
				"<ItemQty>"&Qty&"</ItemQty>" & _
				"<ItemPrice>"&Price&"</ItemPrice>" & _
			"</Item>"
	End Sub

	Public Function doPayment()
		xmlRequest = "<Order>" & _
			"<StoreID>"&Server.HTMLEncode(myStoreID)&"</StoreID>" & _
			"<Passphrase>"&Server.HTMLEncode(myPassphrase)&"</Passphrase>" & _
			"<Tax1>"&Server.HTMLEncode(myTaxTotal1)&"</Tax1>" & _
			"<Tax2>"&Server.HTMLEncode(myTaxTotal2)&"</Tax2>" & _
			"<Tax3>"&Server.HTMLEncode(myTaxTotal3)&"</Tax3>" & _
			"<Tax4>"&Server.HTMLEncode(myTaxTotal4)&"</Tax4>" & _
			"<Tax5>"&Server.HTMLEncode(myTaxTotal5)&"</Tax5>" & _
			"<ShippingTotal>"&Server.HTMLEncode(myShippingtotal)&"</ShippingTotal>" & _
			"<Subtotal>"&Server.HTMLEncode(mySubtotal)&"</Subtotal>" & _
			"<PaymentType>"&Server.HTMLEncode(myPaymentType)&"</PaymentType>" & _
			"<CardAction>"&Server.HTMLEncode(myCardAction)&"</CardAction>" & _
			"<CardNumber>"&Server.HTMLEncode(myCardNumber)&"</CardNumber>" & _
			"<CardExpMonth>"&Server.HTMLEncode(myCardExpMonth)&"</CardExpMonth>" & _
			"<CardExpYear>"&Server.HTMLEncode(myCardExpYear)&"</CardExpYear>" & _
			"<CardIDCode>"&Server.HTMLEncode(myCardIDCode)&"</CardIDCode>" & _
			"<CardIDNumber>"&Server.HTMLEncode(myCardIDNumber)&"</CardIDNumber>" & _
			"<TestResult>"&Server.HTMLEncode(myTestResult)&"</TestResult>" & _
			"<OrderID>"&Server.HTMLEncode(myOrderID)&"</OrderID>" & _
			"<UserID>"&Server.HTMLEncode(myUserID)&"</UserID>" & _
			"<Bname>"&Server.HTMLEncode(myBname)&"</Bname>" & _
			"<Bcompany>"&Server.HTMLEncode(myBcompany)&"</Bcompany>" & _
			"<Baddress1>"&Server.HTMLEncode(myBaddress1)&"</Baddress1>" & _
			"<Baddress2>"&Server.HTMLEncode(myBaddress2)&"</Baddress2>" & _
			"<Bcity>"&Server.HTMLEncode(myBcity)&"</Bcity>" & _
			"<Bprovince>"&Server.HTMLEncode(myBprovince)&"</Bprovince>" & _
			"<Bpostalcode>"&Server.HTMLEncode(myBpostalcode)&"</Bpostalcode>" & _
			"<Bcountry>"&Server.HTMLEncode(myBcountry)&"</Bcountry>" & _
			"<Sname>"&Server.HTMLEncode(mySname)&"</Sname>" & _
			"<Scompany>"&Server.HTMLEncode(myScompany)&"</Scompany>" & _
			"<Saddress1>"&Server.HTMLEncode(mySaddress1)&"</Saddress1>" & _
			"<Saddress2>"&Server.HTMLEncode(mySaddress2)&"</Saddress2>" & _
			"<Scity>"&Server.HTMLEncode(myScity)&"</Scity>" & _
			"<Sprovince>"&Server.HTMLEncode(mySprovince)&"</Sprovince>" & _
			"<Spostalcode>"&Server.HTMLEncode(mySpostalcode)&"</Spostalcode>" & _
			"<Scountry>"&Server.HTMLEncode(myScountry)&"</Scountry>" & _
			"<Phone>"&Server.HTMLEncode(myPhone)&"</Phone>" & _
			"<Email>"&Server.HTMLEncode(myEmail)&"</Email>" & _
			"<Comments>"&Server.HTMLEncode(myComments)&"</Comments>" & _
			"<CustomerIP>"&Server.HTMLEncode(myCustomerIP)&"</CustomerIP>"
			If Items <> "" Then xmlRequest = xmlRequest & Items
			xmlRequest = xmlRequest & "</Order>"
		
		' Use ASPTear to execute XML POST and write output into a string
		Set http = Server.CreateObject("SOFTWING.ASPtear")
		xmlResponse = http.Retrieve(myGatewayURL, 1, "orderXML=" & xmlRequest, "", "")

		' Check whether the ASPTear worked.
		If Err.Number = 0 Then
			' It worked, so setup an XML parser for the result.
			Set parser = Server.CreateObject("Microsoft.XMLDOM")
			' Set async
			parser.async = false
			' Parse the XML response
			parser.loadXML xmlResponse
			If parser.parseError.errorCode = 0 Then
				' Get the result into local variables.
				myResultTrxnTransTime = getData(parser, "TransTime")
				myResultTrxnOrderID = getData(parser, "OrderID")
				myResultTrxnApproved = getData(parser, "Approved")
				myResultTrxnReturnCode = getData(parser, "ReturnCode")
				myResultTrxnErrMsg = getData(parser, "ErrMsg")
				myResultTrxnTaxTotal = getData(parser, "TaxTotal")
				myResultTrxnShipTotal = getData(parser, "ShipTotal")
				myResultTrxnSubTotal = getData(parser, "SubTotal")
				myResultTrxnFullTotal = getData(parser, "FullTotal")
				myResultTrxnPaymentType = getData(parser, "PaymentType")
				myResultTrxnCardNumber = getData(parser, "CardNumber")
				myResultTrxnCardExpMonth = getData(parser, "CardExpMonth")
				myResultTrxnCardExpYear = getData(parser, "CardExpYear")
				myResultTrxnTransRefNumber = getData(parser, "TransRefNumber")
				myResultTrxnCardIDResult = getData(parser, "CardIDResult")
				myResultTrxnAVSResult = getData(parser, "AVSResult")
				myResultTrxnCardAuthNumber = getData(parser, "CardAuthNumber")
				myResultTrxnCardRefNumber = getData(parser, "CardRefNumber")
				myResultTrxnCardType = getData(parser, "CardType")
				myResultTrxnIPResult = getData(parser, "IPResult")
				myResultTrxnIPCountry = getData(parser, "IPCountry")
				myResultTrxnIPRegion = getData(parser, "IPRegion")
				myResultTrxnIPCity = getData(parser, "IPCity")
				myError = 0
				myErrorMessage = ""
			Else
				' An XML error occured. Return the error message and number.
				myError = parser.parseError.errorCode
				myErrorMessage = parser.parseError.reason
			End If
			' Clean up our XML parser
			Set parser = Nothing
		Else
			' A ASPTear Error occured. Return the error message and number.
			myError = Err.Number
			myErrorMessage = Err.Description
		End If
		Set http = Nothing

		'***********************************************************************
		'*** Optional commented-out Debug.  Dont mess with it.               ***
		'***********************************************************************
		' Response.Write xmlRequest
		' Response.Write xmlResponse
		' End Sub

    doPayment = getError()
	End Function

	Private Function getData(ByVal xmlSource, tag)
		getData = ""
		Dim Result
		For Each Result In xmlSource.documentElement.childNodes
			If tag = Result.nodeName AND Result.hasChildNodes Then
				getData = Result.childNodes(0).nodeValue
			End If
		Next
	End Function
End Class
%>


<%

sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst
	Do While Not Rs_Store.EOF
			select case Rs_store("Property")
			case "PsiGate_StoreId"
				PsiGate_StoreId = decrypt(rs_store("Value"))
                        case "Passphrase"
 				Passphrase = decrypt(rs_store("Value"))
		end select
		Rs_store.MoveNext
	Loop
Rs_store.Close

'StoreID ="teststore"
'Passphrase="psigate1234"

if Tax_Exempt then
	TaxExempt="1"
else
	 TaxExempt="0"
end if


If Auth_Capture then
	sChargeType = 0
else
	sChargeType = 1
end if




Dim xObj, strResult, strUrl, strErr
    	Set xObj = CreateObject("SOFTWING.ASPtear")

	'strUrl = "https://dev.psigate.com:7989/Messenger/XMLMessenger"     ' test server
    	strUrl = "https://secure.psigate.com:7934/Messenger/XMLMessenger" ' Production server
    	
    	Dim InputString


        InputString = "<Order>"
        InputString = + InputString + "<StoreID>"+PsiGate_StoreID+"</StoreID>"
        InputString = + InputString + "<Passphrase>"+Passphrase+"</Passphrase>"
        InputString = + InputString + "<Userid>"+cstr(cid)+"</Userid>"
        InputString = + InputString + "<Bname>"+First_name&" "&Last_name+"</Bname>"
        InputString = + InputString + "<Bcompany>"+Company+"</Bcompany>"
	    InputString = + InputString + "<Baddress1>"+Address1+"</Baddress1>"
	    InputString = + InputString + "<Baddress2>"+Address2+"</Baddress2>"
	    InputString = + InputString + "<Bcity>"+City+"</Bcity>"
	    InputString = + InputString + "<Bprovince>"+State+"</Bprovince>"
	    InputString = + InputString + "<Bpostalcode>"+zip+"</Bpostalcode>"
	    InputString = + InputString + "<Bcountry>"+Country+"</Bcountry>"
	    InputString = + InputString + "<Phone>"+Phone+"</Phone>"
	    InputString = + InputString + "<Email>"+ShipEmail+"</Email>"
	    InputString = + InputString + "<Sname>"+ShipFirstname +" "+ ShipLastname+"</Sname>"
        InputString = + InputString + "<Scompany>"+ShipCompany +"</Scompany>"
        InputString = + InputString + "<Saddress1>"+ShipAddress1 +"</Saddress1>"
        InputString = + InputString + "<Saddress2>"+ShipAddress2 +"</Saddress2>"
        InputString = + InputString + "<Scity>"+ShipCity +"</Scity>"
	    InputString = + InputString + "<Sprovince>"+ShipState +"</Sprovince>"
	    InputString = + InputString + "<Spostalcode>"+Shipzip +"</Spostalcode>"
	    InputString = + InputString + "<Scountry>"+ShipCountry +"</Scountry>"
	    InputString = + InputString + "<CustomerIP>"+Request.ServerVariables("REMOTE_ADDR")+"</CustomerIP>"
		InputString = + InputString + "<Subtotal>"+cstr(GGrand_Total)+"</Subtotal>"
        InputString = + InputString + "<PaymentType>"+"CC"+"</PaymentType>"
        InputString = + InputString + "<CardAction>"+cstr(sChargeType)+"</CardAction>"
        InputString = + InputString + "<CardNumber>"+cstr(CardNumber)+"</CardNumber>"
        InputString = + InputString + "<CardExpMonth>"+cstr(Request.Form("mm"))+"</CardExpMonth>"
        InputString = + InputString + "<CardExpYear>"+cstr(Request.Form("yy"))+"</CardExpYear>"
        InputString = + InputString + "</Order>"
        strResult = xObj.Retrieve(strUrl, 1, "orderXML=" + InputString, " ", "")
        strErr = Err.Description & vbCrLf & "Error number: " & Err.Number
	   fn_print_debug "InputString="&InputString

        ' Check whether the ASPTear worked.
		If Err.Number = 0 Then
	' It worked, so setup an XML parser for the result.
			Set parser = Server.CreateObject("Microsoft.XMLDOM")
			' Set async
			parser.async = false
			' Parse the XML response
			parser.loadXML strResult 
			If parser.parseError.errorCode = 0 Then
				' Get the result into local variables.
                           mystatus=parser.getElementsByTagName("Approved")(0).childNodes(0).nodeValue
                           if mystatus="APPROVED" then
                            myResultTrxnOrderID= parser.getElementsByTagName("TransRefNumber")(0).childNodes(0).nodeValue
                            myAuthID= parser.getElementsByTagName("CardAuthNumber")(0).childNodes(0).nodeValue
                            Verified_Ref = myResultTrxnOrderID
                            AuthNumber = myAuthID
                            else
                            fn_purchase_decline oid,"The transaction was rejected by the payment processor " & strResult
                            end if

             










				
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
			'Response.Write "The transaction was rejected by the payment processor:<BR>"& myErrorMessage
			fn_purchase_decline oid,"The transaction was rejected by the payment processor:<BR>"& myErrorMessage
			End If
		Set http = Nothing
 




if Auth_Capture then
	trans_type = 1
else
	trans_type = 0
end if


%>


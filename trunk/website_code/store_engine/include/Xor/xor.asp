<%
sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
rs_Store.open sql_real_time,conn_store,1,1

rs_Store.MoveFirst
Do While Not Rs_Store.EOF
	select case Rs_store("Property")
		case "Xor_user"
			Xor_User = decrypt(rs_store("Value"))
			'Xor_User="utjs"
		 case "Xor_password"
			  Xor_Password =decrypt(rs_Store("Value"))
  			  'Xor_Password ="tkp1ho"
		 case "Xor_currency"
			  Xor_Currency =decrypt(rs_Store("Value"))
  			  'Xor_Currency ="Ils"
		 case "Xor_terminalno"
			  Xor_terminalno =decrypt(rs_Store("Value"))
  			  'Xor_terminalno ="0962858" 
	end select
	Rs_store.MoveNext
Loop
Rs_store.Close

rs_store.open "select * from store_customers where record_type=0 and cid="&cid, conn_store, 1, 1
If not rs_Store.eof then
	billTo_firstName=rs_Store("First_name")
	billTo_lastName=rs_Store("Last_name")
	billTo_street1=rs_Store("Address1")

	billTo_city=rs_Store("City")
	billTo_state=rs_Store("State")
	billTo_postalCode=rs_Store("zip")
	billTo_country=rs_Store("Country")
	billTo_phoneNumber=rs_Store("Phone")
	billTo_email=rs_Store("fax")
'	purchaseTotals_currency="ILS"
End If
rs_store.close



        Dim XmlString

        XmlString = "<ashrait> _"
        XmlString = XmlString & "<request> "
        XmlString = XmlString & "<command>doDeal</command> "
        XmlString = XmlString & "<requestId>" + cstr(Oid) + "</requestId> "
        XmlString = XmlString & "<version>1000</version> "
        XmlString = XmlString & "<language>Eng</language> "
        XmlString = XmlString & "<doDeal>  "	 
        XmlString = XmlString & "<terminalNumber>" + Xor_terminalno + "</terminalNumber> "
        XmlString = XmlString & "<cardNo>" + CardNumber + "</cardNo> "
        if Use_CVV2 then
		XmlString = XmlString &	"<Cvv>"+Card_Code+"</Cvv>"
	end if
        XmlString = XmlString & "<cardExpiration>"&Request.Form("mm")&Request.Form("yy")&"</cardExpiration> "
	XmlString = XmlString & "<transactionType>Debit</transactionType> "
        XmlString = XmlString & "<creditType>RegularCredit</creditType> "
        XmlString = XmlString & "<currency>"&Xor_Currency&"</currency> "
        XmlString = XmlString & "<transactionCode>Phone</transactionCode> "
        XmlString = XmlString & "<total>"&formatnumber((GGrand_Total*100),0,0,0,0)&"</total> "
        XmlString = XmlString & "<validation>AutoComm</validation> "
        XmlString = XmlString & "<user>" + cstr(cid) + "</user> "
        XmlString = XmlString & "</doDeal> "
        XmlString = XmlString & "</request> "
        XmlString = XmlString & "</ashrait>"
        
		Dim objWinHttp
		Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1") 
		
		Const Option_SSLErrorIgnoreFlags = 4
		Const SslErrorFlag_Ignore_All = 13056   '0x3300
		objWinHttp.Option(Option_SSLErrorIgnoreFlags) = SslErrorFlag_Ignore_All

		Dim URL
		URL = "https://nprd.xor-t.com/relay/Relay"
		URL = URL & "?user="& Xor_User & "&password="& Xor_Password &"&int_in=" & server.urlencode(XmlString)

		objWinHttp.Open "GET", URL
		objWinHttp.Send 
		Response.Write "<textarea rows=15 cols=70>" & objWinHttp.ResponseText & "</textarea>"

%>

		<br>
		<br>

<%

		Dim objXML
		Dim objLst

		Set objXML = Server.CreateObject("Microsoft.XMLDOM") 
		objXML.async="false"
		objXML.loadXML(objWinHttp.ResponseText)

		If objXML.parseError.errorCode <> 0 Then
			' handle the error
		End If

		Set objLst = objXML.getElementsByTagName("*")

		For i = 0 to objLst.length-1

			If objLst.item(i).nodeName = "message" Then
				StrMsg = objLst.item(i).text
			End If

			If objLst.item(i).nodeName = "status" Then
				Status = objLst.item(i).text
			End If

			If objLst.item(i).nodeName = "requestId" Then
				ReqId = objLst.item(i).text
			End If

			If objLst.item(i).nodeName = "authNumber" Then
				authNumber = objLst.item(i).text
			End If
			
			If objLst.item(i).nodeName = "cvvStatus" Then
				cvCode = objLst.item(i).text
			End If

		Next

if Status="000" then

		AuthNumber = authNumber
		'avs_result	= ""
		Verified_Ref =  ReqId
		card_verif = cvCode
	

else ' REJECT or ERROR
        reasonCode = StrMsg
        fn_purchase_decline oid,"The transaction was rejected by the payment processor:<BR>"&reasonCode
end if

%>

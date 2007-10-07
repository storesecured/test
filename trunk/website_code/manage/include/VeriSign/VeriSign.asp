
<%
If sType = "Capture" Then
	v_type = "D"
else
	v_type = "C"
end if
sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"

rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst 	 
	Do While Not Rs_Store.EOF
		select case Rs_store("Property")
			case "v_User" 
				user = Rs_store("Value")
			case "v_Vendor"
				vendor = Rs_store("Value")
			case "v_Partner"
				partner = Rs_store("Value") 
			case "v_Password"
				password = Rs_store("Value")
		end select
		Rs_store.MoveNext
	Loop
Rs_store.Close

rs_store.open "select * from store_customers where record_type=0 and cid="&cid, conn_store, 1, 1
If not rs_Store.eof then
	first_name=rs_Store("First_name")
	last_name=rs_Store("Last_name")
	address1=rs_Store("Address1")
	address2=rs_Store("Address2")
	city=rs_Store("City")
	state=rs_store("state")
	zip=rs_Store("zip")
	country=rs_Store("Country")
	phone=rs_Store("Phone")
	fax=rs_Store("fax")
	ccid = rs_Store("CCID")
	if rs_Store("Tax_Exempt") then
		TaxExempt="Y"
	else
		 TaxExempt="N"
	end if
End If
rs_store.close

rs_store.open "Select * from Store_Purchases where Shopper_id ='"&Shopper_ID&"' AND oid = "&Oid&" and Store_id ="&Store_id,conn_store,1,1
If not rs_Store.eof then
	Shipping_Method_Price = Rs_store("Shipping_Method_Price")
	Tax = Rs_store("Tax")
	Cust_PO = Rs_store("Cust_PO")
	ShipFirstname=rs_Store("ShipFirstname")
	ShipLastname=rs_Store("ShipLastname")
	ShipAddress1=rs_Store("ShipAddress1")
	ShipAddress2=rs_Store("ShipAddress2")
	ShipCity=rs_Store("ShipCity")
	ShipState=rs_store("ShipState")
	Shipzip=rs_Store("Shipzip")
	ShipCountry=rs_Store("ShipCountry")
	ShipPhone=rs_Store("ShipPhone")
	ShipFax=rs_Store("ShipFax")
	ShipEmail=rs_Store("ShipEmail")
	ShipCompany=rs_Store("ShipCompany")
	Verified_Ref=rs_Store("Verified_Ref")
	Grand_Total=rs_Store("Grand_Total")
End If
rs_Store.Close

if GGrandTotal>Grand_Total then
   fn_error "You cannot capture or credit an amount greater than the original authorization amount."
end if
Set client = Server.CreateObject("PFProCOMControl.PFProCOMControl.1")


parmList = "TRXTYPE="&v_Type
'S=Sale, C=Credit, V=Void,A=Auth,D=Delayed Capture


parmList = parmList + "&ACCT="&Server.UrlEncode(CardNumber)
parmList = parmList + "&PWD="&Server.UrlEncode(password)
parmList = parmList + "&USER="&Server.UrlEncode(user)
parmList = parmList + "&VENDOR="&Server.UrlEncode(vendor)
parmList = parmList + "&PARTNER="&Server.UrlEncode(partner)
CardExp = left(CardExpiration,2) & right(CardExpiration,2)
parmList = parmList + "&EXPDATE="&Server.UrlEncode(CardExp)
parmList = parmList + "&AMT="&formatnumber(cdbl(GGrand_Total),2)
parmList = parmList + "&STREET="&Server.UrlEncode(Address1)
parmList = parmList + "&ZIP="&Server.UrlEncode(zip)
parmList = parmList + "&NAME="&Server.UrlEncode(ShipFirstname & " " & ShipLastname)
parmList = parmList + "&TENDER=C"
parmList = parmList + "&PONUM="&Server.UrlEncode(Cust_PO)
parmList = parmList + "&SHIPTOZIP="&Server.UrlEncode(Shipzip)
parmList = parmList + "&TAXAMT="&Server.UrlEncode(Tax)
parmList = parmList + "&TAXEXEMPT="&Server.UrlEncode(TaxExempt)
parmList = parmList + "&CITY="&Server.UrlEncode(ShipCity)
parmList = parmList + "&CUSTCODE="&Server.UrlEncode(Shopper_ID)
parmList = parmList + "&STATE="&Server.UrlEncode(ShipState)
parmList = parmList + "&RECURRING=N"
parmList = parmList + "&FREIGHTAMT="&Server.UrlEncode(Shipping_Method_Price)
parmList = parmList + "&ORIGID="&Server.UrlEncode(Verified_Ref)

if Use_CVV2 then
		parmList = parmList + "&CVV2="&Server.UrlEncode(CardCode)
end if

'for testing use payflow.verisign.com
Ctx1 = client.CreateContext("test-payflow.verisign.com", 443, 30, "", 0, "", "")
curString = client.SubmitTransaction(Ctx1, parmList, Len(parmList))
client.DestroyContext (Ctx1)

if curstring="" then
	Response.Redirect "admin_error.asp?message_id=40"
end if

done = 0
Send_Mail sReport_email,sReport_email,"Verisign",curstring
Do while Len(curString) <> 0
	if InStr(curString,"&") Then
		varString = Left(curString, InStr(curString , "&" ) -1)
	else
		varString = curString
	end if
	name = Left(varString, InStr(varString, "=" ) -1)
	value = Right(varString, Len(varString) - (Len(name)+1))
	select case name
		case "RESULT" 
			resultval = value
		case "RESPMSG"
			respMessage = value
		case "AUTHCODE"
			AuthNumber = value
		case "PNREF"
			Verified_Ref = value
		case "AVSADDR"
			avsaddr = value
		case "AVSZIP"
			avszip = value
		case "IAVS"
			iavs = value
		case "CVV2MATCH"
			card_code_verif = value
	end select
	if Len(curString) <> Len(varString) Then 
		curString = Right(curString, Len(curString) - (Len(varString)+1))
	else
		curString = ""
	end if
Loop

if avsaddr = "N" and avszip = "N" then
	avs_result = "N"
elseif avsaddr = "Y" and avszip = "Y" then
	avs_result = "Y"
elseif avsaddr = "N" and avszip = "Y" then
	avs_result = "Z"
elseif avsaddr = "Y" and avszip = "N" then
	avs_result = "A"
elseif iavs = "Y" then
	avs_result = "G"
elseif iavs = "X" or avszip = "X" or avsaddr = "X" then
	avs_result = "S"
end if

if card_code_verif = "Y" then
	card_verif = "M"
elseif card_code_verif = "N" then
	card_verif = "N"
elseif card_code_verif = "X" then
	card_verif = "X"
elseif card_code_verif = "" or not Use_CVV2 then
	card_verif = "P"
end if

If resultval <> 0 then
	Response.Redirect "admin_error.asp?message_id=41"&"&Message_Add="&Server.UrlEncode(respMessage)&" (Err Code = "&resultval&")"
End IF

if Auth_Capture then
	trans_type = 1
else
	trans_type = 0
end if

%>

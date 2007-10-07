
<%

sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
    
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst	 
	Do While Not Rs_Store.EOF
		select case Rs_store("Property")
			case "v_User" 
				user = decrypt(Rs_store("Value"))
			case "v_Vendor"
				vendor = decrypt(Rs_store("Value"))
			case "v_Partner"
				partner = decrypt(Rs_store("Value"))
			case "v_Password"
				password = decrypt(Rs_store("Value"))
		end select
		Rs_store.MoveNext
	Loop
Rs_store.Close

if Tax_Exempt then
	TaxExempt="Y"
else
	TaxExempt="N"
end if

Set client = Server.CreateObject("PFProCOMControl.PFProCOMControl.1")

If Auth_Capture then
	v_type = "S"
else
	v_type = "A"
end if

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

if Use_CVV2 then
		parmList = parmList + "&CVV2="&Server.UrlEncode(CardCode)
end if
if IssueStart<>"Null" then
		parmList = parmList + "&CARDSTART="&Server.UrlEncode(IssueStart)
end if
if IssueNumber<>"Null" then
		parmList = parmList + "&CARDISSUE="&Server.UrlEncode(IssueNumber)
end if

'for real use payflow.verisign.com
'for testing use test-payflow.verisign.com
Ctx1 = client.CreateContext("payflow.verisign.com", 443, 30, "", 0, "", "")
curString = client.SubmitTransaction(Ctx1, parmList, Len(parmList))
client.DestroyContext (Ctx1)

if curstring="" then
        fn_purchase_decline oid,"The transaction was rejected by the payment processor:<BR>Unable to communicate with payment processor."
end if

done = 0
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
        fn_purchase_decline oid,"The transaction was rejected by the payment processor:<BR>"&respMessage&" (Err Code = "&resultval&")"
End IF

if Auth_Capture then
	trans_type = 1
else
	trans_type = 0
end if

%>

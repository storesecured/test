<%
'pfpro test-payflow.verisign.com 443 '“TRXTYPE=C&TENDER=C&PARTNER=VeriSign&VENDOR=SuperMerchant&USER=SuperMerchant&PWD=x1y2z3&ACCT=5555 5555 5555 4444&EXPDATE=0308&AMT=123.00”30

response.write "in the refund block<br>"
cardnumber= request.form("CardNumber")
cardExp= request.form("exp")&"<br>"
refundDescription=request.form("refundDescription") 
GGrand_Total= formatnumber(cdbl(request.form("new_amount")),2)
card_ending = right(cardnumber,4)
store_id =   request.form("store_id")
card_ending =   request.form("card_ending")
user="blac6789"
partner="StoreSecured"
password="ankle237"
CardNumber = DeCrypt(CardNumber)

Set client = Server.CreateObject("PFProCOMControl.PFProCOMControl.1")

parmList = "TRXTYPE=C"
'S=Sale, C=Credit, V=Void,A=Auth,D=Delayed Capture


parmList = parmList + "&ACCT="&Server.UrlEncode(CardNumber)
parmList = parmList + "&PWD="&Server.UrlEncode(password)
parmList = parmList + "&USER="&Server.UrlEncode(user)
parmList = parmList + "&VENDOR="&Server.UrlEncode(vendor)
parmList = parmList + "&PARTNER="&Server.UrlEncode(partner)
parmList = parmList + "&EXPDATE="&Server.UrlEncode(CardExp)
parmList = parmList + "&AMT="&formatnumber(GGrand_Total,2)
parmList = parmList + "&ORIGID="&Orig_id
parmList = parmList + "&TENDER=C"

Ctx1 = client.CreateContext("payflow.verisign.com", 443, 30, "", 0, "", "")
curString = client.SubmitTransaction(Ctx1, parmList, Len(parmList))
client.DestroyContext (Ctx1)

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

If resultval <> 0 then
	Response.Redirect "admin_error.asp?message_id=41"&"&Message_Add="&Server.UrlEncode(respMessage)&" (Err Code = "&resultval&")"
End IF


sql_insert = "Insert into Sys_Payments (Store_Id,Amount,Transaction_ID, Card_Ending,Payment_Description,Payment_Type,Payment_Term) values ("&_
  		           Store_Id&",-"&GGrand_Total&",'"&Verified_Ref&"','"&Card_Ending&"','"&refundDescription&"','Custom',1)"
  		conn_store.Execute sql_insert


response.end

%>

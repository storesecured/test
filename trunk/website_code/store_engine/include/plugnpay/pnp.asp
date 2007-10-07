<%

sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst
Do While Not Rs_Store.EOF
	select case Rs_store("Property")
		case "publisher-name"
			publisher_name = decrypt(Rs_store("Value"))
	end select
	Rs_store.MoveNext
Loop
Rs_store.Close

if Tax_Exempt then
	TaxExempt="TRUE"
else
  TaxExempt="FALSE"
end if

GGrand_Total = FormatNumber(GGrand_Total,2)
if auth_capture then
	sTransType = "authpostauth"
else
	sTransType = "authonly"
end if
         Set PNPObj = CreateObject("pnpcom.main")

	 STR = "authtype="&sTransType
	 STR = STR 	& "&card-address1="&address1
	 STR = STR 	& "&card-address2="&address2
	 STR = STR 	& "&card-city="&city
	 STR = STR 	& "&card-state="&State
	 STR = STR 	& "&card-zip="&zip
	 STR = STR 	& "&card-country="&country
	 STR = STR 	& "&email="&ShipEmail
	 STR = STR 	& "&shipname="&Server.URLEncode(ShipFirstname&" "&ShipLastname)
	 STR = STR 	& "&address1="&ShipAddress1
	 STR = STR 	& "&address2="&ShipAddress2
	 STR = STR 	& "&city="&ShipCity
	 STR = STR 	& "&state="&ShipState
	 STR = STR 	& "&zip="&Shipzip
	 STR = STR 	& "&ipaddress="&Request.ServerVariables("REMOTE_ADDR")
	 STR = STR 	& "&orderID="&Oid&"-"&now()
	 STR = STR 	& "&tax="&Tax
	 STR = STR 	& "&shipping="&Shipping_Method_Price
	 STR = STR 	& "&publisher-name="&publisher_name
	 STR = STR 	& "&publisher-email="&Store_Email
    STR = STR 	& "&card-amount="&GGrand_Total
	 
if Payment_Method="eCheck" then
	BankABA = Request.Form("BankABA")
	BankAccount = Request.Form("BankAccount")
	acct_type = Request.Form("acct_type")
	if acct_type = "CHECKING" then
		sType = "checking"
	else
		sType = "savings"
	end if
	 STR = STR & "&accttype="&sType
	 STR = STR 	& "&routingnum="&BankABA
	 STR = STR 	& "&accountnum="&BankAccount
	 STR = STR 	& "&checknum=1000"
	 STR = STR 	& "&paymethod=onlinecheck"
	 STR = STR & "&card-name="&first_name&" "&last_name
else
	 STR = STR & "&card-name="&first_name&" "&last_name
	 STR = STR 	& "&card-number="&CardNumber
	 STR = STR 	& "&card-exp="&CardExpiration
	 STR = STR 	& "&paymethod=credit"
	 if Use_CVV2 then
		STR = STR & "&card-cvv="&CardCode
	 end if
end if

 Results = PNPObj.doTransaction("",STR,"","")

 Dim objResults
 Set objResults = Server.CreateObject("Scripting.Dictionary")

 myArray = split (Results,"&")
 for i = 0 to UBound(myArray)
   myArray(i) = replace(myArray(i),"+"," ")
   myArray(i) = replace(myArray(i),"%20"," ")
   myArray(i) = replace(myArray(i),"%21","!")
   myArray(i) = replace(myArray(i),"%23","#")
   myArray(i) = replace(myArray(i),"%24","$")
   myArray(i) = replace(myArray(i),"%25","%")
   myArray(i) = replace(myArray(i),"%26","&")
   myArray(i) = replace(myArray(i),"%27","'")
   myArray(i) = replace(myArray(i),"%28","(")
   myArray(i) = replace(myArray(i),"%29",")")
   myArray(i) = replace(myArray(i),"%2c",",")
   myArray(i) = replace(myArray(i),"%2d","-")
   myArray(i) = replace(myArray(i),"%2e",".")
   myArray(i) = replace(myArray(i),"%40","@")
   pos = inStr(1,myArray(i),"=")
   if (pos > 1) then
     myKey = Left(myArray(i),pos-1)
     myVal = Mid(myArray(i),pos+1)
     response.write myKey + " = " + myVal + "<br>"
     objResults.Item(myKey) = myVal
   End If
 Next
      
 if(objResults.Item("FinalStatus") = "success") then
   avs_result = objResults.Item("avs-code")
   card_verif = objResults.Item("cvvresp")
   Verified_Ref = objResults.Item("resp-code")
   AuthNumber = objResults.Item("auth-code")
 Else
   fn_purchase_decline oid,"The transaction was rejected by the payment processor:<BR>"&objResults.Item("MErrMsg")
 End If
 
if Auth_Capture then
	trans_type = 1
else
	trans_type = 0
end if

%>

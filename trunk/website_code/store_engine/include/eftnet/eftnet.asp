
<%

If Auth_Capture then
		stype = "01"
	else
		stype = "02"
	end if


sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst	 
Do While Not Rs_Store.EOF 
	select case Rs_store("Property")
		case "M_ID"
				 M_ID = decrypt(Rs_store("Value"))
		case "M_Key"
				 M_Key = decrypt(Rs_store("Value"))
	end select
	Rs_store.MoveNext
Loop
Rs_store.Close

if Tax_Exempt then
		TaxExempt="TRUE"
else
		 TaxExempt="FALSE"
end if

max_Inv = Oid & "-"&now()
sProductString = "Invoice #: " & oid


Set xObj = CreateObject("SOFTWING.ASPtear")
Post_String = ""
Post_String = Post_String &"M_id="&Server.UrlEncode(M_id)
Post_String = Post_String &"&M_Key="& Server.UrlEncode(M_Key)
Post_String = Post_String &"&T_code="& Server.UrlEncode(stype)
Post_String = Post_String &"&T_ordernum="& Server.UrlEncode(oid & "-"&now())
Post_String = Post_String &"&T_tax="&Server.UrlEncode(Tax)
Post_String = Post_String &"&ForceReload="& Server.UrlEncode(Now())
Post_String = Post_String &"&Store_Id="& Server.UrlEncode(Store_Id)
Post_String = Post_String &"&C_ship_zip="& Server.UrlEncode(ShipZip)
Post_String = Post_String &"&C_ship_state="& Server.UrlEncode(ShipState)
Post_String = Post_String &"&C_ship_city="& Server.UrlEncode(ShipCity)
Post_String = Post_String &"&C_ship_address="& Server.UrlEncode(ShipAddress1&ShipAddress2)
Post_String = Post_String &"&C_ship_name="& Server.UrlEncode(ShipFirstname & " " &ShipLastname)
Post_String = Post_String &"&C_state="& Server.UrlEncode(State)
Post_String = Post_String &"&T_code="& Server.UrlEncode(stype)
Post_String = Post_String &"&T_amt="&Server.UrlEncode(GGrand_Total)
Post_String = Post_String &"&C_address="& Server.UrlEncode(Address1&Address2)
Post_String = Post_String &"&C_city="& Server.UrlEncode(city)
Post_String = Post_String &"&C_country="& Server.UrlEncode(country)
Post_String = Post_String &"&C_zip="& Server.UrlEncode(zip)
Post_String = Post_String &"&T_customer_number="& Server.UrlEncode(cid)
Post_String = Post_String &"&C_email="&Server.UrlEncode(email)
Post_String = Post_String &"&T_shipping="& Server.UrlEncode(Shipping_Method_Price)

Post_String = Post_String &"&C_name="& Server.UrlEncode(first_name & " " &last_name)
Post_String = Post_String &"&C_cardnumber="& Server.UrlEncode(CardNumber)
Post_String = Post_String &"&C_exp="& Server.UrlEncode(Request.Form("mm")&Request.Form("yy"))
if Use_CVV2 then
	Post_String = Post_String &"&C_cvv="& Server.UrlEncode(CardCode)
end if


strResult = xObj.Retrieve("https://va.eftsecure.net/cgi-bin/eftBankcard.dll?transaction", 1, Post_String, "", "")

  sApproval = mid(strResult,2,1)
  sError = mid(strResult,9,32)
  auth_code = mid(strResult,3,6)
  card_code_response= mid(strResult,43,1)
  avs_code=mid(strResult,44,1)
  trans_id= mid(strResult,47,10)



If sApproval="A" then
	' Transaction accepted
Else
	' Transaction denied or error, redirect the user to try again ...
	fn_purchase_decline oid,"The transaction was rejected by the payment processor:<BR>"&sError&strResult
End IF
Verified_Ref=trans_id
AuthNumber= auth_code ' value returend from authorizenet
avs_result = avs_code
card_verif = card_code_response
if Auth_Capture then
	trans_type = 1
else
	trans_type = 0
end if
%>

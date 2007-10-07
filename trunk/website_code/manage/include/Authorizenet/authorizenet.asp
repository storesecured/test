

<%

If sType = "Capture" Then
	x_type = "PRIOR_AUTH_CAPTURE"
elseif sType = "Credit" then
	x_type = "CREDIT"
else
        x_type="VOID"
end if
sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst	 
Do While Not Rs_Store.EOF 
	select case Rs_store("Property")
		case "x_Login"
			x_Login = decrypt(Rs_store("Value"))
		case "x_tran_key"
			x_tran_key = decrypt(Rs_store("Value"))
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
		TaxExempt="TRUE"
	else
		 TaxExempt="FALSE"
	end if
End If
rs_store.close

rs_store.open "Select * from Store_Purchases where oid = "&Oid&" and Store_id ="&Store_id,conn_store,1,1
If not rs_Store.eof then
	Shipping_Method_Price = Rs_store("Shipping_Method_Price")
	Tax = Rs_store("Tax")
	Cust_PO = Rs_store("Cust_PO")
	ShipFirstname=rs_Store("ShipFirstname")
	ShipLastname=rs_Store("ShipLastname")
	ShipAddress1=rs_Store("ShipAddress1")
	ShipAddress2=rs_Store("ShipAddress2")
	ShipCity=rs_Store("ShipCity")
	ShipState=rs_Store("ShipState")
	Shipzip=rs_Store("Shipzip")
	ShipCountry=rs_Store("ShipCountry")
	ShipPhone=rs_Store("ShipPhone")
	ShipFax=rs_Store("ShipFax")
	ShipEmail=rs_Store("ShipEmail")
	ShipCompany=rs_Store("ShipCompany")
	Verified_Ref=rs_Store("Verified_Ref")
	CardNumber=rs_Store("CardNumber")
	CardExpiration=rs_Store("CardExpiration")
	AuthNumber=rs_Store("AuthNumber")
End If
rs_Store.Close

Set xObj = CreateObject("SOFTWING.ASPtear")
'xObj.TrustUnknownCA = True
'xObj.IgnoreInvalidCN = True
'xObj.ForceReload = True

Post_String = ""
Post_String = Post_String &"x_login="&Server.UrlEncode(x_Login)
Post_String = Post_String &"&x_tran_key="& Server.UrlEncode(x_tran_key)
Post_String = Post_String &"&x_version="& Server.UrlEncode("3.1")
Post_String = Post_String &"&x_delim_data=True"
Post_String = Post_String &"&x_delim_char=|"
Post_String = Post_String &"&x_encap_char="
Post_String = Post_String &"&x_ADC_URL=False"
Post_String = Post_String &"&x_relay_response=FALSE"
Post_String = Post_String &"&x_test_request="
Post_String = Post_String &"&x_trans_ID="& Server.UrlEncode(Verified_Ref)
Post_String = Post_String &"&x_tax_exempt="&TaxExempt
Post_String = Post_String &"&x_tax="&Server.UrlEncode(Tax)
Post_String = Post_String &"&ForceReload="& Server.UrlEncode(Now())
Post_String = Post_String &"&x_ship_to_country="& Server.UrlEncode(ShipCountry)
Post_String = Post_String &"&x_ship_to_zip="& Server.UrlEncode(ShipZip)
Post_String = Post_String &"&x_ship_to_state="& Server.UrlEncode(ShipState)
Post_String = Post_String &"&x_ship_to_city="& Server.UrlEncode(ShipCity)
Post_String = Post_String &"&x_ship_to_address="& Server.UrlEncode(ShipAddress1&ShipAddress2)
Post_String = Post_String &"&x_ship_to_first_Name="& Server.UrlEncode(ShipFirstname)
Post_String = Post_String &"&x_ship_to_last_Name="& Server.UrlEncode(ShipLastname)
Post_String = Post_String &"&x_first_name="& Server.UrlEncode(ShipFirstname)
Post_String = Post_String &"&x_last_name="& Server.UrlEncode(ShipLastname)
Post_String = Post_String &"&x_company="& Server.UrlEncode(ShipCompany)
Post_String = Post_String &"&x_state="& Server.UrlEncode(ShipState)
Post_String = Post_String &"&x_type="& Server.UrlEncode(x_Type)
Post_String = Post_String &"&x_amount="&Server.UrlEncode(GGrand_Total)
Post_String = Post_String &"&x_address="& Server.UrlEncode(ShipAddress1&ShipAddress2)
Post_String = Post_String &"&x_city="& Server.UrlEncode(ShipCity)
Post_String = Post_String &"&x_country="& Server.UrlEncode(ShipCountry)
Post_String = Post_String &"&x_zip="& Server.UrlEncode(ShipZip)
Post_String = Post_String &"&x_phone="& Server.UrlEncode(ShipPhone)
Post_String = Post_String &"&x_fax="& Server.UrlEncode(ShipFax)
Post_String = Post_String &"&x_cust_id="& Server.UrlEncode(ccid)
Post_String = Post_String &"&x_customer_ip="&Server.UrlEncode(Request.ServerVariables("REMOTE_ADDR"))
Post_String = Post_String &"&x_email="&Server.UrlEncode(ShipEmail)
'Post_String = Post_String &"&x_email_customer=TRUE"
Post_String = Post_String &"&x_freight="& Server.UrlEncode(Shipping_Method_Price)
Post_String = Post_String &"&x_duty=0"
Post_String = Post_String &"&x_invoice_num="&Oid
Post_String = Post_String &"&x_po_num="&Cust_PO

if Payment_Method="eCheck" then
	BankABA = Request.Form("BankABA")
	BankAccount = Request.Form("BankAccount")
	acct_type = Request.Form("acct_type")
	TaxID = Request.Form("TaxID")
	org_type = Request.Form("org_type")
	Name = first_name & " " & last_name
	Post_String = Post_String &"&x_method="& "ECHECK"
	Post_String = Post_String &"&x_echeck_type="& "WEB"
	Post_String = Post_String &"&x_bank_aba_code="& Server.UrlEncode(BankABA)
	Post_String = Post_String &"&x_bank_acct_num="& Server.UrlEncode(BankAccount)
	Post_String = Post_String &"&x_bank_acct_type="& Server.UrlEncode(acct_type)
	Post_String = Post_String &"&x_bank_name="& Server.UrlEncode(BankName)
	Post_String = Post_String &"&x_bank_acct_name="& Server.UrlEncode(Name)
	Post_String = Post_String &"&x_customer_organization_type="& Server.UrlEncode(org_type)
	if TaxID = "" then
		DrvState = Request.Form("DrvState")
		DrvNumber = Request.Form("DrvNumber")
		dobd = Request.Form("dobd")
		dobm = Request.Form("dobm")
		doby = Request.Form("doby")
	  dob = dobm & "/" & dobd & "/" & doby
		Post_String = Post_String &"&x_drivers_license_state="& Server.UrlEncode(DrvState)
		Post_String = Post_String &"&x_drivers_license_dob="&Server.UrlEncode(dob)
		Post_String = Post_String &"&x_drivers_license_num="& Server.UrlEncode(DrvNumber)
	
	else
		Post_String = Post_String &"&x_customer_tax_id="& Server.UrlEncode(CardExpiration)
	end if
else
	Post_String = Post_String &"&x_method="& "CC"
	Post_String = Post_String &"&x_card_num="& Server.UrlEncode(decrypt(CardNumber))
	Post_String = Post_String &"&x_exp_date="& Server.UrlEncode(CardExpiration)
	if Use_CVV2 then
		Post_String = Post_String &"&x_card_code="& Server.UrlEncode(CardCode)
	end if
end if

strResult = xObj.Retrieve("https://secure.authorize.net/gateway/transact.dll", 1, Post_String, "", "")
bErrors = HandleError()
If  bErrors or Instr(1, strResult, "|") = 0 Then
	response.redirect "error.asp?Message_id=101&Message_Add="&Server.UrlEncode("Error Occured while communicating with authorize.net servers ..")
else
	Returned_Var_Array = Split(strResult,"|")
	response_code = replace(Returned_Var_Array(0),"""","")
	response_subcode = replace(Returned_Var_Array(1),"""","")
	response_reason_code = replace(Returned_Var_Array(2),"""","")
	response_reason_text = replace(Returned_Var_Array(3),"""","")
	auth_code = replace(Returned_Var_Array(4),"""","")
	avs_code = replace(Returned_Var_Array(5),"""","")
	trans_id = replace(Returned_Var_Array(6),"""","")
	card_code_response = replace(Returned_Var_Array(38),"""","")
End If

If response_code="1" then
	' Transaction accepted
Else
	' Transaction denied or error, redirect the user to try again ...
	response.redirect "error.asp?Message_id=101&Message_Add="&Server.UrlEncode(response_reason_text)
End IF
Verified_Ref=trans_id
AuthNumber= auth_code ' value returend from authorizenet
avs_result = avs_code
card_verif = card_code_response

%>

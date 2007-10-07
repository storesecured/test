
<%

If Auth_Capture then
	x_type = "AUTH_CAPTURE"
else
	x_type = "AUTH_ONLY"
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

if Tax_Exempt then
   TaxExempt="TRUE"
else
	 TaxExempt="FALSE"
end if

if 1=0 then
	sql_select = "select Invoice_num from store_settings where store_id="&Store_Id
	rs_store.open sql_select,conn_store,1,1
		max_Inv = rs_Store("Invoice_num")
	rs_store.close
	max_Inv=max_Inv+1

	sql_auth = "update store_settings set Invoice_num="&max_Inv&" where Store_id ="&Store_id
	conn_store.Execute sql_auth
   	sql_auth = "update store_purchases set Invoice_Id="&max_Inv&" where Shopper_id ="&Shopper_ID&" AND oid = "&Oid&" and Store_id ="&Store_id&";"
	conn_store.Execute sql_auth
elseif store_id=625 then
	sProductString = "--"
	sql_tran_items = "select Item_Name,Quantity from store_transactions where Transaction_Processed=0 and shipto="&shipto&"  and Shopper_id ="&Shopper_ID&"	 AND Store_id ="&Store_id
	Set rs_store_tran = Server.CreateObject("ADODB.Recordset")
	rs_store_tran.open sql_tran_items, conn_store, 1, 1
	do while not rs_store_tran.eof
  		Item_Name = rs_store_tran("Item_Name")
  		Quantity = rs_store_tran("Quantity")
  		sProductString = sProductString & Quantity & ":"&Item_Name&"--"
  		rs_store_tran.movenext
	loop
  	rs_store_tran.close
  	set rs_store_tran = nothing
else
	sProductString = "Invoice #: " & oid
end if

max_Inv = Oid

Set xObj = CreateObject("SOFTWING.ASPtear")
'xObj.TrustUnknownCA = True
'xObj.IgnoreInvalidCN = True
'xObj.ForceReload = True
Post_String = ""
Post_String = Post_String &"x_login="&Server.UrlEncode(x_Login)
'Post_String = Post_String &"&x_password=im13test"
Post_String = Post_String &"&x_tran_key="& Server.UrlEncode(x_tran_key)
Post_String = Post_String &"&x_version="& Server.UrlEncode("3.1")
Post_String = Post_String &"&x_delim_data=True"
Post_String = Post_String &"&x_delim_char=|"
Post_String = Post_String &"&x_relay_response=FALSE"
Post_String = Post_String &"&x_test_request=FALSE"
Post_String = Post_String &"&x_tax_exempt="&TaxExempt
Post_String = Post_String &"&x_description=" & Server.UrlEncode(left(sProductString,255))
Post_String = Post_String &"&x_tax="&Server.UrlEncode(Tax)
Post_String = Post_String &"&ForceReload="& Server.UrlEncode(Now())
Post_String = Post_String &"&x_ship_to_country="& Server.UrlEncode(ShipCountry)
Post_String = Post_String &"&x_ship_to_zip="& Server.UrlEncode(ShipZip)
Post_String = Post_String &"&x_ship_to_state="& Server.UrlEncode(ShipState)
Post_String = Post_String &"&x_ship_to_city="& Server.UrlEncode(ShipCity)
Post_String = Post_String &"&x_ship_to_address="& Server.UrlEncode(ShipAddress1&ShipAddress2)
Post_String = Post_String &"&x_ship_to_first_Name="& Server.UrlEncode(ShipFirstname)
Post_String = Post_String &"&x_ship_to_last_Name="& Server.UrlEncode(ShipLastname)
Post_String = Post_String &"&x_first_name="& Server.UrlEncode(first_name)
Post_String = Post_String &"&x_last_name="& Server.UrlEncode(last_name)
Post_String = Post_String &"&x_company="& Server.UrlEncode(Company)
Post_String = Post_String &"&x_ship_to_company="& Server.UrlEncode(ShipCompany)
Post_String = Post_String &"&x_state="& Server.UrlEncode(State)
Post_String = Post_String &"&x_type="& Server.UrlEncode(x_Type)
Post_String = Post_String &"&x_amount="&Server.UrlEncode(GGrand_Total)
Post_String = Post_String &"&x_address="& Server.UrlEncode(Address1&Address2)
Post_String = Post_String &"&x_city="& Server.UrlEncode(city)
Post_String = Post_String &"&x_country="& Server.UrlEncode(country)
Post_String = Post_String &"&x_zip="& Server.UrlEncode(zip)
Post_String = Post_String &"&x_phone="& Server.UrlEncode(phone)
Post_String = Post_String &"&x_fax="& Server.UrlEncode(fax)
Post_String = Post_String &"&x_cust_id="& Server.UrlEncode(ccid)
Post_String = Post_String &"&x_customer_ip="&Server.UrlEncode(Request.ServerVariables("REMOTE_ADDR"))
Post_String = Post_String &"&x_email="&Server.UrlEncode(email)
Post_String = Post_String &"&x_freight="& Server.UrlEncode(Shipping_Method_Price)
Post_String = Post_String &"&x_duty=0"
Post_String = Post_String &"&x_invoice_num="&max_Inv
Post_String = Post_String &"&x_po_num="&Cust_PO
Post_String = Post_String &"&x_recurring_billing=NO"
Post_String = Post_String &"&x_duplicate_window=300"

if Payment_Method="eCheck" then
	BankABA = Request.Form("BankABA")
	BankAccount = Request.Form("BankAccount")
	BankName = Request.Form("BankName")
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
		Post_String = Post_String &"&x_customer_tax_id="& Server.UrlEncode(TaxID)
	end if
else
	Post_String = Post_String &"&x_method="& "CC"
	Post_String = Post_String &"&x_card_num="& Server.UrlEncode(CardNumber)
	Post_String = Post_String &"&x_exp_date="& Server.UrlEncode(CardExpiration)
	if Use_CVV2 then
		Post_String = Post_String &"&x_card_code="& Server.UrlEncode(CardCode)
	end if
end if
Post_String=replace(Post_String,"&#39;","")

on error resume next
if store_id=217 then
strResult = xObj.Retrieve("https://certification.authorize.net/gateway/transact.dll", 1, Post_String, "", "")
else
strResult = xObj.Retrieve("https://secure.authorize.net/gateway/transact.dll", 1, Post_String, "", "")
end if

if err.number<>0 then
	fn_error "There was an error processing your payment.  The payment has not been processed and the Authorize.net Gateway appears to be experiencing problems.  Please try again later."
end if
on error goto 0

if instr(strResult,"|") then
	Returned_Var_Array = Split(strResult,"|")
	response_code = replace(Returned_Var_Array(0),"""","")
	response_subcode = replace(Returned_Var_Array(1),"""","")
	response_reason_code = replace(Returned_Var_Array(2),"""","")
	response_reason_text = replace(Returned_Var_Array(3),"""","")
	auth_code = replace(Returned_Var_Array(4),"""","")
	avs_code = replace(Returned_Var_Array(5),"""","")
	trans_id = replace(Returned_Var_Array(6),"""","")
	card_code_response = replace(Returned_Var_Array(38),"""","")
else
	response.write strResult
    response.end
end if

If response_code="1" then
	' Transaction accepted
Else
	' Transaction denied or error, redirect the user to try again ...
	if response_reason_code=13 then
            sReason = "The transaction was rejected by the payment processor due to an invalid setup.  Your credit card has not been charged.<BR><HR><BR>Store Owner, please recheck your Authorize.net login and ensure your account is active with Authorize.net."
        elseif response_reason_code=103 then
            sReason = "The transaction was rejected by the payment processor due to an invalid setup.  Your credit card has not been charged.<BR><HR><BR>Store Owner, please recheck your Authorize.net login and transaction key combination."
        else
            sReason = "The transaction was rejected by the payment processor:<BR>"&response_reason_text&"<BR>"&response_reason_code&"-"&response_subcode&"-"&response_code
        end if
        fn_purchase_decline oid,sReason

End IF
Verified_Ref=trans_id
AuthNumber= auth_code ' value returned from authorizenet
avs_result = avs_code
card_verif = card_code_response
if Auth_Capture then
	trans_type = 1
else
	trans_type = 0
end if
%>

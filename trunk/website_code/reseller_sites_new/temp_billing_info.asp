<%
dim strinsert
'Retriving card type selected from reseller_payment_page
cardtype = trim(Request.Form("payment_method"))
'Retriving reseller id from session variable
resellerid = trim(Request.Form("hidSessionID"))

if trim(Request.QueryString("infoaction"))<> "save" then
	'Updating card type for that reseller
	str="update tblTemp set fld_card_type='"&cardtype&"' where fld_reseller_id="&resellerid&""
	conn.Execute str
end if

'Retriving the values from fields.
if trim(Request.QueryString("infoaction"))="save" then
	
	'Code here to send the page to process the billing in the manage folder
	
		
	cardno = trim(Request.Form("cardno"))
	cardcode = trim(Request.Form("cardcode"))
	strmonth = trim(Request.Form("month"))
	stryear = trim(Request.Form("year"))
	bankname = trim(Request.Form("bank"))
	routing = trim(Request.Form("routing"))
	account = trim(Request.Form("account"))
	accounttype = trim(Request.Form("acctype"))
	
	accounttype1 = trim(Request.Form("Orgtype1"))
	check = trim(Request.Form("check"))
	lic = trim(Request.Form("lic"))
	drilic = trim(Request.Form("drilic"))
	
	Day1 = trim(Request.Form("licexp"))
	month1 = trim(Request.Form("month"))
	year1 = trim(Request.Form("fromyy"))

	mail = trim(Request.Form("email"))
	
	flag = trim(Request.Form("hidflag"))
	
	
	'Code here to update the values into the database for that reseller depending on the card type
	if flag = "1" then
	strinsert = "update tbltemp set fld_credit_card_number="&cardno&",fld_card_expire_month="&strmonth&","&_
				" fld_card_expire_year="&stryear&", fld_card_code="&cardcode&" where fld_Reseller_Id="&resellerid&" "
	conn.Execute (strinsert)

	end if
	
	if flag = "2" then
	strinsert1 = "update tbltemp set fld_bank_name='"&bankname&"',fld_routing_number='"&routing&"', fld_account_number='"&account&"',"&_
				 " fld_account_type='"&accounttype&"',fld_org_type='"&accounttype1&"', fld_check_number='"&check&"',"&_
				 " fld_driver_licenses='"&lic&"',fld_driver_State='"&drilic&"',fld_licenses_exp_month='"&month1&"', "&_
				 " fld_licenses_exp_year='"&year1&"',fld_licenses_exp_day='"&Day1&"' where fld_Reseller_Id="&resellerid&""
	Response.Write "strinsert1"&strinsert1
	conn.Execute (strinsert1)
	end if
	
	if flag = "3" then
	strinsert2 = "update tbltemp set fld_mail_paypal='"&mail&"' where fld_Reseller_Id="&resellerid&""
	conn.Execute (strinsert2)
	end if
	
	
	
	
	'code here to redirect him to the billing thing
	
	Response.Redirect "reseller_billing_info.asp"
	Response.End
%>
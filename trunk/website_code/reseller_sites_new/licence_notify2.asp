<!--#include virtual="common/connection.asp"-->

<!--#include file="include/sub.asp"-->
<%

' read post from PayPal system and add 'cmd'
str = Request.Form

Txn_id = Request.Form("txn_id")
' post back to PayPal system to validate
str = str & "&cmd=_notify-validate"
' assign posted variables to local variables
' note: additional IPN variables also available -- see IPN documentation
Receiver_email = Request.Form("receiver_email")
Temp_Reseller_Id = Request.Form("item_number")
Payment_status = Request.Form("payment_status")
total = Request.Form("payment_gross")
Txn_id = Request.Form("txn_id")
Payer_email = Request.Form("payer_email")
Temp_Reseller_Id = Request.Form("custom")

isverif = false
' Check notification validation

'response.write objHttp.responseText
Set rs_store = Server.CreateObject("ADODB.Recordset")
'if (objHttp.status <> 200 ) then
'elseif (objHttp.responseText = "VERIFIED") then
 if (Payment_status = "Completed" or Payment_status="Pending") then
			'look up info in db

					if Receiver_email = sPaypal_email then

							'transaction looks good so update verified status
							if (Payment_status="Pending") then
									AuthNumber = "Pending"
							else
									AuthNumber = Txn_id
							end if
							if Payment_status <> "Pending" then
  								'code here to make an entry into the final table after the reseller has paid 	
  	
  								'Code here to retrieve all values from tbltemp
								sqlselect = "select fld_first_name,fld_last_name,fld_user_name,fld_password,fld_address,"&_
											" fld_city,fld_state,fld_zip_code,fld_phone,fld_fax,fld_mail,fld_mail_paypal,"&_
											" fld_website,fld_card_type,fld_credit_card_number,fld_card_expire_month,fld_card_expire_year,"&_
											" fld_card_code,fld_company_name,fld_license_fee,fld_bank_name,fld_routing_number,fld_account_number,"&_
											" fld_account_type,fld_account_type1,fld_check_number,fld_driver_licenses,"&_
											" fld_driver_licenses1,fld_licenses_exp_month,fld_licenses_exp_year from tbltemp"&_
											" where fld_reseller_id="&Temp_Reseller_Id&""
										
								set rsselect = conn_store.Execute(sqlselect)
								if not rsselect.EOF then
									first = trim(rsselect("fld_first_name"))
									last = trim(rsselect("fld_last_name"))
									username = trim(rsselect("fld_user_name"))
									password = trim(rsselect("fld_password"))
									address = trim(rsselect("fld_address"))
									city = trim(rsselect("fld_city"))
									state = trim(rsselect("fld_state"))
									zip = trim(rsselect("fld_zip_code"))
									phone = trim(rsselect("fld_phone"))
									fax = trim(rsselect("fld_fax"))
									mail = trim(rsselect("fld_mail"))
									mail1 = trim(rsselect("fld_mail_paypal"))
									reseller_site = trim(rsselect("fld_website"))
									cardtype = trim(rsselect("fld_card_type"))
									cardno = trim(rsselect("fld_credit_card_number"))
									cmonth = trim(rsselect("fld_card_expire_month"))
									cyear = trim(rsselect("fld_card_expire_year"))
									cardcode = trim(rsselect("fld_card_code"))
									company = trim(rsselect("fld_company_name"))
									bank = trim(rsselect("fld_bank_name"))
									routing = trim(rsselect("fld_routing_number"))
									accountno = trim(rsselect("fld_account_number"))
									accounttype = trim(rsselect("fld_account_type"))
									accounttype1 = trim(rsselect("fld_account_type1"))
									check = trim(rsselect("fld_check_number"))
									driver = trim(rsselect("fld_driver_licenses"))
									driver1 = trim(rsselect("fld_driver_licenses1"))
									dmonth = trim(rsselect("fld_licenses_exp_month"))
									dyear = trim(rsselect("fld_licenses_exp_year"))
								end if
	
								'Code here to insert values into tbl_reseller_master
								sqlinsert = "insert into tbl_reseller_master(fld_first_name,fld_last_name,"&_
											" fld_user_name,fld_password,fld_address,fld_city,fld_state,fld_zip_code,"&_
											" fld_phone,fld_fax,fld_mail,fld_mail1,fld_card_type,fld_credit_card_number,"&_
											" fld_card_expire_month,fld_card_expire_year,fld_card_code,fld_company_name,"&_
											" fld_bank_name,fld_routing_number,fld_account_number,"&_
											" fld_account_type,fld_account_type1,fld_check_number,fld_driver_licenses,"&_
											" fld_driver_licenses1,fld_licenses_exp_month,fld_licenses_exp_year,fld_website) values "&_
											" ('"&first&"','"&last&"','"&username&"','"&password&"','"&address&"','"&city&"',"&_
											" '"&state&"','"&zip&"','"&phone&"','"&fax&"','"&mail&"','"&mail1&"','"&cardtype&"','"&cardno&"',"&_
											" '"&cmonth&"','"&cyear&"','"&cardcode&"','"&company&"','"&bank&"','"&routing&"','"&accountno&"','"&accounttype&"',"&_
											" '"&accounttype1&"','"&check&"','"&driver&"','"&driver1&"','"&dmonth&"','"&dyear&"','"&reseller_site&"')"	
								conn1.Execute(sqlinsert)
	
								'Code here to create session for newly inserted reseller
								set rsNew = conn1.execute("select max(fld_reseller_id) as maxid from tbltemp")
									if not rsNew.eof then
									intresellerid = trim(rsNew(0))
									end if
									Session("resellerid") = intresellerid
	
									'code here to create the store of the reseller 
									 
									'reseller_site is the site name entered by the reseller
	
																			
									Response.Redirect "../code/reselleradmin/reseller_home.asp"
									Response.End
	
							end if
							
							Send_Mail Payer_email,Receiver_email,"Store Paid","Store " & ResellerID & " has paid " & total & " for easystorecreator reseller's site to be activated. (PayPal)" &Payment_status
						
		
				 end if
	end if
' check that Payment_status=Completed
' check that Txn_id has not been previously processed
' check that Receiver_email is an email address in your PayPal account
' process payment
'end if
'set objHttp = nothing

%>


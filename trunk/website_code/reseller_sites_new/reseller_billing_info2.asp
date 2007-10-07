

<%
 
if trim(Session("tempresellerid")) ="" then 
	Response.redirect "http://admin1.easystorecreator.com/error.asp?Message_id=94"
end if


%>


<%
title = "Free Website Builder eCommerce Merchant Account Free Online Store Builder"
description = "Free website builder allows easy ecommerce merchant account integration. Free online store builder expedites increased sales. Trial free website builder today."
keyword1="free website builder"
keyword2="ecommerce merchant account"
keyword3="free online store builder"
keyword4=""
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="reseller_billing_info"
includejs=1
%>

<!--#include virtual="common/crypt.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include virtual="common/cc_validation.asp"-->
<%
'ERROR CHECKING
on error goto 0


If Form_Error_Handler(Request.Form) <> "" then
   Error_Log = Form_Error_Handler(Request.Form)
   %><!--#include file="Include/Error_Template.asp"--><%

else

 processtype = "auth_capture"

   total = "00.00"
   
   if isNumeric(total) then
      total = formatnumber(total,2)
   else
      Response.redirect "error.asp?Message_id=100&Message_Add="&server.urlencode("Your transaction could not be processed due to an invalid total.")
   end if
   '**********************************Code added ny Sudha********************************
   set conn = server.CreateObject("adodb.connection")
   'strconn = "DRIVER=SQL Server;SERVER=afroz;UID=user;PWD=thinkmore;DATABASE=esc"
   strconn = "DRIVER=SQL Server;SERVER=198.87.87.59;UID=user;PWD=thinkmore;DATABASE=wizard"		
   conn.open strconn
   sqlselect = "select fld_first_name,fld_last_name,fld_user_name,fld_password,fld_address,"&_
				" fld_city,fld_state,fld_zip_code,fld_phone,fld_fax,fld_mail,fld_mail_paypal,"&_
				" fld_website,fld_card_type,fld_credit_card_number,fld_card_expire_month,fld_card_expire_year,"&_
				" fld_card_code,fld_company_name,fld_license_fee,fld_bank_name,fld_routing_number,fld_account_number,"&_
				" fld_account_type,fld_org_type,fld_check_number,fld_driver_licenses,"&_
				" fld_driver_State,fld_licenses_exp_month,fld_licenses_exp_year,fld_licenses_exp_day from tbltemp"&_
				" where fld_reseller_id="&Session("tempresellerid") &""
	
			
	set rsselect = conn.Execute(sqlselect)
	if not rsselect.EOF then
	
		first_name=	trim(rsselect("fld_first_name"))
		last_name = trim(rsselect("fld_last_name"))
		username = trim(rsselect("fld_user_name"))
		password = trim(rsselect("fld_password"))
		address = trim(rsselect("fld_address"))
		city = trim(rsselect("fld_city"))
		state = trim(rsselect("fld_state"))
		zip = trim(rsselect("fld_zip_code"))
		phone = trim(rsselect("fld_phone"))
		fax = trim(rsselect("fld_fax"))
		email = trim(rsselect("fld_mail"))
		reseller_site = trim(rsselect("fld_website"))
		
		
		payment_method = trim(rsselect("fld_card_type"))
		
		cc_num = trim(rsselect("fld_credit_card_number"))
		mm = trim(rsselect("fld_card_expire_month"))
		yy = trim(rsselect("fld_card_expire_year"))
		CardCode = trim(rsselect("fld_card_code"))
			company  = trim(rsselect("fld_company_name"))
			BankName = trim(rsselect("fld_bank_name"))
			BankABA= trim(rsselect("fld_routing_number"))
			BankAccount = trim(rsselect("fld_account_number"))
			acct_type = trim(rsselect("fld_account_type"))
			org_type = trim(rsselect("fld_org_type"))
			CheckSerial= trim(rsselect("fld_check_number"))
			DrvNumber  = trim(rsselect("fld_driver_licenses"))
			DrvState = trim(rsselect("fld_driver_State"))
			dobd = trim(rsselect("fld_licenses_exp_day"))
			dobm = trim(rsselect("fld_licenses_exp_month"))
			doby = trim(rsselect("fld_licenses_exp_year"))
			if payment_method <> "eCheck" then 
				
			end if
	end if
	'code here to reterive the values corresponding to the paypment metheod choose over here either payal or check
	'Code here to retrieve the values from tbltemp

	
			sqlget = ("getnew "&Session("tempresellerid")&" ")
			
			'Code here to retrieve the values fromtbltemp
			set rsget= conn.execute(sqlget)
			
			if not rsget.eof then 
				mode = trim(rsget("fld_new_mode"))
				nemail = trim(rsget("fld_new_mail"))
				npayable = trim(rsget("fld_new_pay"))
				naddress = trim(rsget("fld_new_address"))
				nstate = trim(rsget("fld_new_state"))
				ncity = trim(rsget("fld_new_city"))
				nzip = trim(rsget("fld_new_zip"))
				ncountry=trim(rsget("fld_new_country"))
			end if
	
  %>
     

  <%

		'Response.Write "reseller_site "&reseller_site 
         if payment_method="eCheck" then
           sRight = Right(BankAccount,4)
         else
           sRight = Right(cc_num,4)
         end if

         'code here to insert the reseller into the final table
         
         
         'Code here to insert values into tbl_reseller_master
			sqlinsert = "insert into tbl_reseller_master(fld_first_name,fld_last_name,"&_
						" fld_user_name,fld_password,fld_address,fld_city,fld_state,fld_zip_code,"&_
						" fld_phone,fld_fax,fld_mail,fld_card_type,fld_credit_card_number,"&_
						" fld_card_expire_month,fld_card_expire_year,fld_card_code,fld_company_name,"&_
						" fld_bank_name,fld_routing_number,fld_account_number,"&_
						" fld_account_type,fld_org_type,fld_check_number,fld_driver_licenses,"&_
						" fld_driver_State,fld_licenses_exp_month,fld_licenses_exp_year,fld_website,fld_licenses_exp_day) values "&_
						" ('"&first_name&"','"&last_name&"','"&username&"','"&password&"','"&address&"','"&city&"',"&_
						" '"&state&"','"&zip&"','"&phone&"','"&fax&"','"&email&"','"&payment_method&"','"&cc_num&"',"&_
						" '"&mm&"','"&yy&"','"&CardCode&"','"&company&"','"&BankName&"','"&BankABA&"','"&BankAccount&"','"&acct_type&"',"&_
						" '"&org_type&"','"&CheckSerial&"','"&DrvNumber&"','"&DrvState&"','"&dobm&"','"&doby&"','"&reseller_site&"','"&dobd&"')"	
			conn.Execute(sqlinsert)
	
		'Code here to create session for newly inserted reseller
		'set rsNew =  conn.execute("SELECT @@IDENTITY")
			set rsNew = conn.execute("select max(fld_reseller_id) as maxid from tbl_reseller_master")
			if not rsNew.eof then
				intresellerid = trim(rsNew(0))
			end if
			
		Session("resellerid") = intresellerid
		
		
		'*******************************************************************************************************
		
		'code here to insert the values into the tbl_esc_reseller_Payment_mode
			
		sqlnew = ("putnew "&intresellerid&","&mode&",'"&nemail&"', '"&npayable&"', '"&naddress&"', '"&nstate&"','"&ncity&"','"&nzip&"', "&ncountry&" ")
		conn.execute sqlnew
		
		'*******************************************************************************************************
		'************************************************** Ends Here*****************************************************

		
		currentdate=date()
		conn.execute("Put_reseller_rates "&intresellerid&",'"&currentdate&"'") 
        
        
        'code here to send the mail to at the reseller's account for the site information
        '****************************************************************************************************************
        Email_Text = "Congratulations, your site has been created and is now available at http://reseller.storesecured.com/?Reseller="&Session("resellerid")&""&_
					 "<br>You can login to manage your store at http://manage.storesecured.com  or from our homepage at "&_
					 "<br> or from our homepage at http://www.storesecured.com "&_
					 "<br>"&_
					 "Login Name : "&username&"<BR>Password: "&password

	'mail is sent here to the registred reseller
		 Set Mail = Server.CreateObject("Persits.MailSender")
		 Mail.From = sNoReply_email
		 Mail.AddAddress Email
		 Mail.Subject =  "Site Created"
		 Mail.Body = Email_Text
		 Mail.isHTML = True
		 Mail.Queue=True
		 Mail.Send
		Set Mail = Nothing 
        
       '****************************************************************************************************************
	
	'mail is sent here to the easystore creator
	Email_Text = "Reseller ID "&intresellerID&" has requested for the site - www."&reseller_site&".com"
	Set Mail = Server.CreateObject("Persits.MailSender")
		 Mail.From = sReport_email
		 Mail.Subject =  "Site Request Notification "
		 Mail.Body = Email_Text
		 Mail.Queue=True
		 Mail.Send
		Set Mail = Nothing 
	
		
		Response.Redirect "intermediate.asp"
		
end if
%>
<!--#include file="footer.asp"-->
<!--#include virtual="common/connection.asp"-->
<!--#include file="header.asp"-->
<%


Session("tempresellerid")= trim(Request.Form("hidSessionID"))

if trim(Request.Form("hidSessionID")) = "" and Session("tempresellerid")="" then 
	Response.redirect "error.asp?Message_id=94"
end if
'REPLACE 'SINGLE QUOTES
function fixquotes(str)
fixquotes = replace(str&"","'","''")
end function

function checkEncode(encodeVal)

	if encodeVal<>"" then
	checkEncode = server.HTMLEncode(encodeVal)
	else
	checkEncode = encodeVal
	end if
	
end function

%>
<!--#include virtual="common/crypt.asp"-->
<!--#include virtual="common/cc_validation.asp"-->
<%
'ERROR CHECKING
on error goto 0




 processtype = "auth_capture"

   total = trim(Request.Form("AMOUNT"))
   
   if isNumeric(total) then
      total = formatnumber(total,2)
   else
      Response.redirect "error.asp?Message_id=100&Message_Add="&server.urlencode("Your transaction could not be processed due to an invalid total.")
   end if
   
   
   
   '**********************************Code added ny Sudha********************************
   
   
   sqlselect = "select fld_user_name,fld_password,fld_mail,fld_website,fld_Type from tbltemp"&_
				" where fld_reseller_id="&Session("tempresellerid")&""
			
	set rsselect = conn_store.Execute(sqlselect)
	first_name=	 trim(Request.Form("First_name"))
	last_name =	 trim(Request.Form("Last_name"))
	address = 	 trim(Request.Form("address"))
	city = 	 trim(Request.Form("city"))
	state = 	 trim(Request.Form("state"))
	zip = 	 trim(Request.Form("Zip"))
	zip = trim(fixquotes(Request.Form("zip")))
	phone = 	 trim(Request.Form("Phone"))
	fax = 	 trim(Request.Form("Fax"))
	company = 	 trim(Request.Form("Company"))
	strcountry = trim(Request.Form("selcountry"))
	payment_method = trim(Request.Form("hidtype"))
	mm = trim(Request.Form("month"))
	yy = trim(Request.Form("year"))
	cc_num = trim(Request.Form("cardno"))
	cardcode = trim(Request.Form("cardcode"))
	if trim(cardcode) ="" then
		cardcode ="0"
	end if

	if not rsselect.EOF then
		username = trim(checkencode(rsselect("fld_user_name")))
		password = trim(checkencode(rsselect("fld_password")))
		email = trim(rsselect("fld_mail"))
		reseller_site = trim(rsselect("fld_website"))
		Type_of_Request = trim(rsselect("fld_Type"))

		'if yy < 10 then
			'yy = "0"&yy
		'end if
		'if mm < 10 then
		'	'mm = "0"&mm
		'end if
		
	'code here to retrive the credit card information


	if bankname<> "" then 
		bankname = replace(bankname,"'","''")
	end if
	
			CardCode = cardcode
			company  = trim(Request.Form("Company"))
			BankName =  trim(Request.Form("bank"))
			BankABA= trim(Request.Form("routing"))
			BankAccount = trim(Request.Form("account"))
			acct_type = trim(Request.Form("acctype"))
			org_type = trim(Request.Form("Orgtype1"))
			CheckSerial= trim(Request.Form("check"))
			DrvNumber  =  trim(Request.Form("lic"))
			DrvState = trim(Request.Form("drilic"))
			dobd = trim(Request.Form("licexp"))
			dobm = trim(Request.Form("month"))
			doby = trim(Request.Form("fromyy"))
			if payment_method <> "eCheck" then 
				'code here to check the validations for the cedit card
				if not IsCreditCard(payment_method,cc_num) then
					 Response.redirect "error.asp?Message_id=45"
				end if
				if 2000+cint(yy)=year(now) then
					 if cint(mm)<month(now) then
						Response.redirect "error.asp?Message_id=46"
					end if
				end if
		
				 	
		end if
		
	end if
	
  %>
  <!--#include file="process_cc.asp"-->
  <%	

		'Response.Write "reseller_site "&reseller_site 
         if payment_method="eCheck" then
           sRight = Right(BankAccount,4)
         else
           sRight = Right(cc_num,4)
         end if
		
        'code here to insert the reseller into the final table
        currentdate=date()
		first_name=	fixquotes(first_name)
		last_name = fixquotes(last_name) 
		username = fixquotes(username)
		password = fixquotes(password)
		address = fixquotes(address)
		city = fixquotes(city)
		state = fixquotes(state)
		company  = fixquotes(company)
		BankName = fixquotes(BankName)
		sRight = fixquotes(sRight)
		BankABA = fixquotes(BankABA)
		BankAccount = fixquotes(BankAccount)
		CheckSerial = fixquotes(CheckSerial)
		DrvNumber = fixquotes(DrvNumber)
		
		 'Code here to insert values into tbl_reseller_master
			sqlinsert = "insert into tbl_reseller_master(fld_first_name,fld_last_name,"&_
						" fld_user_name,fld_password,fld_address,fld_city,fld_state,fld_zip_code,"&_
						" fld_phone,fld_fax,fld_mail,fld_card_type,fld_credit_card_number,"&_
						" fld_card_expire_month,fld_card_expire_year,fld_card_code,fld_company_name,"&_
						" fld_bank_name,fld_routing_number,fld_account_number,"&_
						" fld_account_type,fld_org_type,fld_check_number,fld_driver_licenses,"&_
						" fld_driver_State,fld_licenses_exp_month,fld_licenses_exp_year,fld_website,fld_licenses_exp_day,fld_plan_transaction_date,fld_type,Auth_Number,AVS_Result,Card_Verif,fld_license_fee,fld_country) values "&_
						" ('"&first_name&"','"&last_name&"','"&username&"','"&password&"','"&address&"','"&city&"',"&_
						" '"&state&"','"&zip&"','"&phone&"','"&fax&"','"&email&"','"&payment_method&"','"&encrypt(cc_num)&"',"&_
						" '"&mm&"','"&yy&"','"&CardCode&"','"&company&"','"&BankName&"','"&encrypt(BankABA)&"','"&encrypt(BankAccount)&"','"&acct_type&"',"&_
						" '"&org_type&"','"&CheckSerial&"','"&DrvNumber&"','"&DrvState&"','"&dobm&"','"&doby&"','"&reseller_site&"','"&dobd&"','"&currentdate&"','"&Type_of_Request&"','"&card_verif&"','"&Verified_Ref&"','"&sRight&"','"&total&"','"&strcountry&"')"	
			conn_store.Execute(sqlinsert)
			'Response.Write "sqlinsert "&sqlinsert 
		'	Response.End
			
			
	
			
		'Code here to create session for newly inserted reseller
		set rsNew =  conn_store.execute("SELECT IDENT_CURRENT('tbl_reseller_master')")
		'set rsNew = conn_store.execute("select max(fld_reseller_id) as maxid from tbl_reseller_master")
			if not rsNew.eof then
				intresellerid = trim(rsNew(0))
			end if
			
		Session("resellerid") = intresellerid
		
		'redirecting here to the reseller's login
'		Response.Redirect "intermediate.asp"
		
		
		
		'*******************************************************************************************************
		'************************************************** Ends Here*****************************************************
		conn_store.execute("Put_reseller_rates "&intresellerid&",'"&currentdate&"'") 
        
        'code here to send the mail to at the reseller's account for the site information
        '****************************************************************************************************************
        Email_Text = "Congratulations, your site has been created and is now available to preview at http://reseller.storesecured.com/?Reseller="&Session("resellerid")&_
					 vbcrlf&vbcrlf&"You can login to manage your store at http://managereseller.storesecured.com  or from our homepage at "&_
					 " http://reseller.storesecured.com "&_
					 vbcrlf&vbcrlf&_
					 "Login Name : "&username&"<BR>Password: "&password&_
					 vbcrlf&vbcrlf&_
					 "Your new domain name should be ready in approximately 48 hours.  After the 48 hours have passed you can start sending visitors to your new website at www."&reseller_site

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
	Email_Text = "Reseller ID "&intresellerID&" has requested to "&Type_of_Request&" the site - www."&reseller_site&" "
	Set Mail = Server.CreateObject("Persits.MailSender")
		 Mail.From = sNoReply_email
		 Mail.AddAddress sReport_email
		 Mail.Subject =  "Site Request Notification "
		 Mail.Body = Email_Text
		 Mail.Queue=True
		 Mail.Send
		Set Mail = Nothing 
		
		'redirection here to the selection of payment mode	
		
		'redirecting here to the reseller's login
		Response.Redirect "reseller_payment_mode.asp"
		


%>

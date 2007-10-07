<!--#include file="header_noview.asp"-->
<!--#include file="emails.asp"-->
<%
If not CheckReferer then
	fn_redirect Switch_Name&"error.asp?message_id=1"
end if

'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	fn_redirect Switch_Name&"form_error.asp?Error_Log="&server.urlencode(Error_Log)
else
	Record_type =Request.Form("Record_type")
	If request.form("ExpressCheckout")<>"" and ExpressCheckout then
		User_ID = EXPRESS_CHECKOUT_CUSTOMER
		Randomize
		Password = "PASS_"&cstr(Int((10000) * Rnd + lowerbound))&month(now)&day(now)&year(now)&cstr(Int((10000) * Rnd + lowerbound))
	else
		User_ID = checkStringForQ(Request.Form("User_ID"))
		Password = checkStringForQ(Request.Form("Password"))
		if Request.Form("Password") <> Request.Form("Password_Confirm") then
			fn_redirect Switch_Name&"error.asp?Message_id=10"
		end if

	end if

	Zip = checkStringForQ(Request.Form("Zip"))
	State_UnitedStates = checkStringForQ(Request.Form("State_UnitedStates"))
	State_Opt = checkStringForQ(Request.Form("State_Opt"))
	State_Canada = checkStringForQ(Request.Form("State_Canada"))
	Registration_Date = now()
	Last_name = checkStringForQ(Request.Form("Last_name"))
	First_name = checkStringForQ(Request.Form("First_name"))
	Company = checkStringForQ(Request.Form("Company"))
	Address1 = checkStringForQ(Request.Form("Address1"))
	Address2 = checkStringForQ(Request.Form("Address2"))
	City = checkStringForQ(Request.Form("City")) 
	Country = checkStringForQ(Request.Form("Country"))
	if Country="United States" then
		State=State_UnitedStates
		if State="" then
          fn_error "Please select a state."
       end if
	elseif Country="Canada" then
		State=State_Canada
		if State="" then
			fn_error "Please select a province."
		end if
    else
       State=State_Opt
	end if
	
	Phone = checkStringForQ(Request.Form("Phone"))
	EMail = checkStringForQ(Request.Form("EMail"))

	'CHECK EMAIL VALIDITY
	if Instr(1, EMail, "@") = 0 or	Instr(1, EMail, ".") = 0 then
		fn_redirect Switch_Name&"error.asp?Message_Id=33"
	end if

	FAX = checkStringForQ(Request.Form("FAX"))
	'SET THE DATE WHEN THIS RECORD WAS LAST CHANGED
	if Request.Form("Spam") <>  "" then
		Spam = Request.Form("Spam")
	else 
		Spam = 0
	end if
	
        Tax_id = Request.Form("Tax_id")
	if Request.Form("tax_exempt") <>  "" then
		tax_exempt = Request.Form("tax_exempt")
	else
		tax_exempt = 0
	end if
	if tax_exempt<>0 and Tax_id="" then
        	fn_error "You must enter a tax id if you are tax exempt."
	end if
	
	if Request.Form("Is_Residential") <>  "" then
		Is_Residential = Request.Form("Is_Residential")
	else 
		Is_Residential = 0
	end if
	'THE CUSTOMER DOES NOT HAVE A BUDGET

    sql_create_customer = "exec wsp_customer_register "&Store_id&","&Shopper_Id&",'"&User_ID&"','"&Password&"','"&Last_name&"','"&First_name&"','"&Company&"','"&Address1&"','"&Address2&"','"&City&"','"&Zip&"','"&State&"','"&Country&"','"&Phone&"','"&EMail&"','"&FAX&"',"&Spam&","&StartCID&","&Tax_Exempt&","&Is_Residential
    fn_print_debug sql_create_customer
    session("sql")=sql_create_customer
    on error resume next
    conn_store.execute sql_create_customer
    if err.number<>0 then
    		fn_error err.description
    end if
    on error goto 0

    if cid=-1 then
     fn_error ccid
    end if
    'LOGIN THE CUSTOMER
    if AllowCookies=-1 then
	    %><!--#include file="cookie.asp"--><%
	    if request.form("SaveCookie")<>"" then
		    call saveUserCookie(Request.Form("User_id"), Request.Form("Password"))
	    else
		    call deleteUserCookie()
	    end if
    end if

    'IF EMAIL NOTIFICATION ENABLE, SENT EMAIL TO THE CUSTOMER
    if Registration_enable <> 0 and not (request.form("ExpressCheckout")<>"" and ExpressCheckout) then
		if Registration_user_psw <> 0 then
			If request.form("ExpressCheckout")<>"" and ExpressCheckout then
				user_psw_message = ""
			else
				user_psw_message = "Your login is : "&User_ID&vbcrlf&"Your password is : "&Password&vbcrlf&"You can login to the store again using your login and password above at our online store "&Site_Name
			end if
		else
			user_psw_message = ""
		end if	
		send_to = email&","&Replace(Registration_sent_to, " ", "")
        if InStr(Registration_body,"%") > -1 then
            Registration_body = Replace(Registration_body,"%LASTNAME%",Last_name)
            Registration_body = Replace(Registration_body,"%FIRSTNAME%",First_name)
            Registration_body = Replace(Registration_body,"%LOGIN%",User_ID)
            Registration_body = Replace(Registration_body,"%PASSWORD%",Password)
        end if
        if InStr(Registration_subject,"%") > -1 then
            Registration_subject = Replace(Registration_subject,"%LASTNAME%",Last_name)
            Registration_subject = Replace(Registration_subject,"%FIRSTNAME%",First_name)
            Registration_subject = Replace(Registration_subject,"%LOGIN%",User_ID)
            Registration_subject = Replace(Registration_subject,"%PASSWORD%",Password)
        end if
		Call Send_Mail_Html(Store_Email,send_to,Registration_subject,Registration_body&vbcrlf&user_psw_message)
	end if



	'IF USER WANTS TO USE THE BILLING ADDRESS ALSO FOR SHIPPING, ADD A NEW
	'RECORD IN STORE_CUSTOMERS TABLE ELSE REDIRECT HIM TO REGISTER SHIPPING
	'PAGE

	if Request.Form("Shipping_Same") <> "" then
		if request.form("Redirect") <> "" then
		    fn_redirect unencode(request.form("Redirect"))
		end if
		fn_redirect Switch_Name&"Register_Thank_You.asp"
	else
		fn_redirect Switch_Name&"Register_Shipping.asp?ExpressCheckout="&request.form("ExpressCheckout")&"&ReturnTo="&Server.Urlencode(unencode(request.form("Redirect")))
	end if

End If 

%>

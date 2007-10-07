<!--#include file="header_noview.asp"-->
<!--#include file="sub.asp"-->

<% 

'ERROR CHECKING

If not CheckReferer then
	fn_redirect Switch_Name&"error.asp?message_id=1"
end if

If Form_Error_Handler(Request.Form) <> "" and not Request.Form("Delete_Ship_Addr") <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	fn_redirect Switch_Name&"form_error.asp?Error_Log="&server.urlencode(Error_Log)
else

	'MODIFY BILLING ADDRESS
	if Request.Form("Form_Name") = "modify_customer" then
  	'RETRIEVE FORM DATA
		Zip= checkStringForQ(Request.Form("Zip"))
		State_UnitedStates = checkStringForQ(Request.Form("State_UnitedStates"))
    	State_Opt = checkStringForQ(Request.Form("State_Opt"))
    	State_Canada = checkStringForQ(Request.Form("State_Canada"))
    	Record_type = Request.Form("Record_type")

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
		Email = checkStringForQ(Request.Form("Email"))
		Fax = checkStringForQ(Request.Form("Fax"))
		LastAccess = Now ()
		if request.form("Is_Residential")<>"" then
		   Is_Residential=request.form("Is_Residential")
		else
		   Is_Residential=0
		end if

		sql_update_customer = "exec wsp_customer_update_addr "&store_id&","&Cid&","&Record_type&",'"&Last_name&"','"&First_name&"','"&Company&"','"&Address1&"','"&Address2&"','"&City&"','"&Zip&"','"&State&"','"&Country&"','"&Phone&"','"&EMail&"','"&FAX&"',"&Is_Residential&";"
        fn_print_debug sql_update_customer
        session("sql")=sql_create_customer
    		conn_store.Execute sql_update_customer
        
        if cint(record_type)=cint(-1) then
            sql="select max(record_type) as record_type from store_customers where store_id="&store_id&" and cid="&cid
            fn_print_debug sql
            rs_store.open sql
            record_type=rs_store("record_type")
            rs_store.close
        end if

		if Record_type = 0 then
			fn_redirect Switch_name&"Modify_my_Billing.asp"
		elseif Record_type = 1 then
            fn_redirect Switch_name&"Modify_my_shipping.asp"
        else
			fn_redirect Switch_name&"Modify_my_shipping.asp?ssadr="&Record_type
		end if
	end if

	'MODIFY LOGIN INFO
	if Request.Form("Form_Name") = "modify_account" then
		Record_type = Request.Form("Record_type")
		User_id = checkStringForQ(Request.Form("User_id"))
		Password = checkStringForQ(Request.Form("Password"))
		Password_Confirm = checkStringForQ(Request.Form("Password_Confirm"))
		if Request.Form("Password") <> Request.Form("Password_Confirm") then
			fn_redirect Switch_name&"error.asp?Message_id=10"
		end if
		'ONE CLICK SHOPPING SETTINGS
		if Request.Form("Spam") <>  "" then
			Spam = Request.Form("Spam")
		else 
			Spam = 0
		end if

		sql_update_customer = "Update store_customers Set User_ID = '"&User_id&"' ,Password = '"&Password&"',Spam= "&Spam&" Where Cid = "&Cid&" And Store_id = "&Store_id&" And Record_type = "&Record_type
		conn_store.Execute sql_update_customer

		if Record_type=0 then
			fn_redirect Switch_Name&"Modify_my_Account.asp"
		else
			fn_redirect Switch_Name&"Modify_my_Account.asp?ssadr="&Record_type
		end if
	end if
 
end if 

'DELETE SHIPPING ADDRESS
if Request.Form("Form_Name") = "Delete_Addr" then
    sql_delete = "exec wsp_customer_delete_record "&store_id&","&cid&","&Request.Form("Delete_Addr")&";"
	fn_print_debug sql_delete
	conn_store.Execute sql_delete
	fn_redirect Switch_Name&"Modify_my_shipping.asp?ssadr=1"
end if

%>

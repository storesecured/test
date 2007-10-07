<!--#include file="header_noview.asp"-->

<%'ERROR CHECKING

If not CheckReferer then
	fn_redirect Switch_Name&"error.asp?message_id=1"
end if

If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	fn_redirect Switch_Name&"form_error.asp?Error_Log="&server.urlencode(Error_Log)

else
	Record_type = Request.Form("Record_type")
	if Record_type="" then
	      Record_type=0
	end if
	User_ID = "NULL"
	Password = "NULL"
	Registration_Date = now()			
	Last_name = checkStringForQ(Request.Form("Last_name"))
	First_name = checkStringForQ(Request.Form("First_name"))
	Company = checkStringForQ(Request.Form("Company"))
	Address1 = checkStringForQ(Request.Form("Address1"))
	Address2 = checkStringForQ(Request.Form("Address2"))
	City = checkStringForQ(Request.Form("City"))
	Zip = checkStringForQ(Request.Form("Zip"))
	State_UnitedStates = checkStringForQ(Request.Form("State_UnitedStates"))
	State_Opt = checkStringForQ(Request.Form("State_Opt"))
	State_Canada = checkStringForQ(Request.Form("State_Canada"))
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

	FAX = checkStringForQ(Request.Form("FAX"))
	if Request.Form("Spam") <>  "" then
		Spam = Request.Form("Spam")
	else 
		Spam = 0
	end if
	if Request.Form("Is_Residential") <>  "" then
		Is_Residential = Request.Form("Is_Residential")
	else 
		Is_Residential = 0
	end if
		
	sql_update_customer = "exec wsp_customer_register_update "&Store_Id&","&Cid&","&Record_type&",'"&Last_name&"','"&First_name&"','"&Company&"','"&Address1&"','"&Address2&"','"&City&"','"&Zip&"','"&State&"','"&Country&"','"&Phone&"','"&EMail&"','"&FAX&"',"&Is_Residential&";"
	session("sql")=sql_update_customer
	fn_print_debug sql_update_customer
    conn_store.Execute sql_update_customer
	if request.form("Redirect") <> "" then
		 fn_redirect unencode(request.form("Redirect"))
	end if
	fn_redirect Switch_Name&"Register_Thank_You.asp"

End If

%>   
	
	
	
	

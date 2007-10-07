<!--#include file="global_settings.asp"-->

<%

' IF CALLING FROM AFFILIATES' PAGES, USE AFFILIATES_SESSION.ASP
' FOR SESSION MANAGEMENT
RUN_FROM_AFFILIATES = FALSE

	Session("AFFILIATE_SESSION")="FALSE"


' ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include virtual="common/Error_Template.asp"--><%
else

	' ADD A NEW AFFILIATE
	if request.form("ACTION")="ADD_AFFILIATE" then

		' RETRIEVING FORM DATA
		Code = replace(request.form("Code"),"'","''")
		Contact_Name = replace(request.form("Contact_Name"),"'","''")
		Email = replace(request.form("Email"),"'","''")
		URL = replace(request.form("URL"),"'","''")
		Email_Not = replace(request.form("Email_Not"),"'","''")
		Password = replace(request.form("Password"),"'","''")
		Password_Confirm = replace(request.form("Password_Confirm"),"'","''")
		Zip = checkStringForQ(Request.Form("Zip"))
		State = checkStringForQ(Request.Form("State"))
		Company = checkStringForQ(Request.Form("Company"))
		Address = checkStringForQ(Request.Form("Address"))
		City = checkStringForQ(Request.Form("City"))
		Country = checkStringForQ(Request.Form("Country"))
		Phone = checkStringForQ(Request.Form("Phone"))

		' CHECK PASSWORD AND PASSWORD CONFIRMATION
		if (Password<>Password_Confirm) then
			response.redirect "admin_error.asp?message_id=64"
		end if
		
		' CHECK IF ADDRESS STARTS WITH HTTP://
		if (len(URL)<7) then
			response.redirect "admin_error.asp?message_id=71"
		else
			if ucase((mid(URL,1,7)))<>"HTTP://" then
				response.redirect "admin_error.asp?message_id=71"
			end if
		end if
		
		' CHECK EMAIL VALIDITY
		if Instr(1, Email, "@") = 0 or  Instr(1, Email, ".") = 0 then
				response.redirect "admin_error.asp?message_id=72"
			end if

		' CHECK IF AFFILIATE CODE EXIST
		sql_check = "select Affiliate_ID from Store_Affiliates where Code='"&Code&"' and store_id="&store_id
		rs_Store.open sql_check,conn_store,1,1
		if not rs_store.eof then
			rs_store.close
			response.redirect "admin_error.asp?message_id=65"
		end if
		rs_store.close
	
		'INSERT INTO THE DATABASE
		sql_ins = "insert into Store_Affiliates (Store_ID, Code, Contact_Name, Email, URL, Password, Email_Notification,Approved,Zip,State,Company,Address,City,Country,Phone) values ("&store_id&", '"&Code&"', '"&Contact_Name&"', '"&Email&"', '"&URL&"', '"&Password&"', "&Email_Not&",0,'"&Zip&"','"&State&"','"&Company&"','"&Address&"','"&City&"','"&Country&"','"&Phone&"')"
		conn_store.execute sql_ins
	
		' REDIRECT TO AFFILIATES LIST WINDOW
		response.redirect "affiliates_manager.asp"
	
	end if


	if request.form("ACTION")="EDIT_AFFILIATE" then

		' RETRIEVING FORM DATA
		Code = replace(request.form("Code"),"'","''")
		Contact_Name = replace(request.form("Contact_Name"),"'","''")
		Email = replace(request.form("Email"),"'","''")
		URL = replace(request.form("URL"),"'","''")
		Email_Not = replace(request.form("Email_Not"),"'","''")
		Password = replace(request.form("Password"),"'","''")
		Password_Confirm = replace(request.form("Password_Confirm"),"'","''")
		AFF_ID = request.form("AFF_ID")
		Zip = checkStringForQ(Request.Form("Zip"))
		State = checkStringForQ(Request.Form("State"))
		Company = checkStringForQ(Request.Form("Company"))
		Address = checkStringForQ(Request.Form("Address"))
		City = checkStringForQ(Request.Form("City"))
		Country = checkStringForQ(Request.Form("Country"))
		Phone = checkStringForQ(Request.Form("Phone"))

		' CHECK PASSWORD AND PASSWORD CONFIRMATION
		if (Password<>Password_Confirm) then
			response.redirect "admin_error.asp?message_id=64"
		end if

		' CHECK IF ADDRESS STARTS WITH HTTP://
		if (len(URL)<7) then
			response.redirect "admin_error.asp?message_id=71"
		else
			if ucase((mid(URL,1,7)))<>"HTTP://" then
				response.redirect "admin_error.asp?message_id=71"
			end if
		end if
		
		' CHECK EMAIL VALIDITY
		if Instr(1, Email, "@") = 0 or  Instr(1, Email, ".") = 0 then
				response.redirect "admin_error.asp?message_id=72"
			end if

		' CHECK IF CODE EXISTS
		sql_check = "select Affiliate_ID from Store_Affiliates where Affiliate_ID<>"&AFF_ID&" and Code='"&Code&"' and store_id="&store_id
		rs_Store.open sql_check,conn_store,1,1
		if not rs_store.eof then
			rs_store.close
			response.redirect "admin_error.asp?message_id=65"
		end if
		rs_store.close
	
		' GET THE OLD CODE
		'sql_old_code = "select Code from Store_Affiliates where Affiliate_ID="&AFF_ID&" and store_id="&store_id
		''rs_Store.open sql_old_code,conn_store,1,1
		'old_code = rs_Store("Code")
		'rs_Store.close
		' REPLACE THE OLD CODE WITH THE NEW CODE IN STORE_SHOPPERS
		'sql_updatess = "update store_shoppers set CAME_FROM ='"&Code&"' where CAME_FROM='"&old_code&"'"
		'conn_store.execute sql_updatess
	
		'UPDATE THE DATABASE
		sql_ins = "UPDATE Store_Affiliates set Code = '"&Code&"', Contact_Name = '"&Contact_Name&"', Email = '"&Email&"', URL = '"&URL&"', Password = '"&Password&"', Email_Notification = "&Email_Not&", Zip='"&Zip&"',State='"&State&"',Company='"&Company&"',Address='"&Address&"',City='"&City&"',Country='"&Country&"',Phone='"&Phone&"' where store_id="&store_id&" and Affiliate_ID="&AFF_ID&" "
		conn_store.execute sql_ins
	
		' REDIRECT TO AFFILIATES LIST WINDOW
		if RUN_FROM_AFFILIATES then
			response.redirect "new_affiliate.asp"
		else
			response.redirect "affiliates_manager.asp"
		End If
	
	end if

End If 

%> 

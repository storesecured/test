<!--#include virtual="common/connection.asp"-->

<% 

' IF CALLING FROM AFFILIATES' PAGES, USE AFFILIATES_SESSION.ASP
' FOR SESSION MANAGEMENT
RUN_FROM_AFFILIATES = TRUE
if Session("AFFILIATE_SESSION")="TRUE" then
	%><!--#include file="affiliates_session.asp"--><%
else
	Response.Redirect "affiliates_login.asp"
end if

' ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include virtual="common/Error_Template.asp"--><%
else

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
		Zip = replace(Request.Form("Zip"),"'","''")
    	State = replace(Request.Form("State"),"'","''")
    	Company = replace(Request.Form("Company"),"'","''")
    	Address = replace(Request.Form("Address"),"'","''")
   	City = replace(Request.Form("City"),"'","''")
    	Country = replace(Request.Form("Country"),"'","''")
    	Phone = replace(Request.Form("Phone"),"'","''")

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
		if Instr(1, Email, "@") = 0 or	Instr(1, Email, ".") = 0 then
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
		sql_old_code = "select Code from Store_Affiliates where Affiliate_ID="&AFF_ID&" and store_id="&store_id
		rs_Store.open sql_old_code,conn_store,1,1
		old_code = rs_Store("Code")
		rs_Store.close
		' REPLACE THE OLD CODE WITH THE NEW CODE IN STORE_SHOPPERS
		sql_updatess = "update store_shoppers set CAME_FROM ='"&Code&"' where CAME_FROM='"&old_code&"'"
		conn_store.execute sql_updatess
	
		'UPDATE THE DATABASE
		sql_ins = "UPDATE Store_Affiliates set Code = '"&Code&"', Contact_Name = '"&Contact_Name&"', Email = '"&Email&"', URL = '"&URL&"', Password = '"&Password&"', Email_Notification = "&Email_Not&", Zip='"&Zip&"',State='"&State&"',Company='"&Company&"',Address='"&Address&"',City='"&City&"',Country='"&Country&"',Phone='"&Phone&"' where store_id="&store_id&" and Affiliate_ID="&AFF_ID&" "
		conn_store.execute sql_ins

		' REDIRECT TO AFFILIATES LIST WINDOW
		response.redirect "new_affiliate.asp"

	end if

End If 

Function Form_Error_Handler(My_Form)
	' This function gets a form collection and return an Error string , if not empty that mean that we
	' have errors

	' set the error to initial empty value :
	Error = ""

	For Each Key In My_Form
		New_Key = Key&"_C"
		if My_Form(New_Key) <> "" then
			' that is the signal that we need to do validation for this form element ...
			' first split with | as delimiter
			MyFormElementArray = split(My_Form(New_Key),"|",-1,1)
			'glosary of checks to perform :
			'position		value	check			error message ...
			'0				Re		required		Field required
			'0				Op						Field Optional
			'1 (type)		String	     String
			'1				Integer 	Integer 		value integer
			'2 min size (for text ..)			100					Field to long ...
			'3 max size (for text)
			'4 must have characters @,$,%
			'5 must NOT have characters @,$,%
			if MyFormElementArray(0) = "Re" AND My_Form(Key) = "" then
				Error = Error & "Field <b>"& Key & "</b> cannot be empty | "
			end if
			' check to see if the type is ok ...
			if MyFormElementArray(1) = "Integer" and not IsNumeric(My_Form(Key)) then
				Error = Error & "Field <b>"& Key & "</b> must be a integer | "
			end if
			if MyFormElementArray(1) = "date" and not IsDate(My_Form(Key)) then
				Error = Error & "Field <b>"& Key & "</b> must be a date | "
			end if

			' process the size check only if not null
			if MyFormElementArray(2) <> "" then
				' for strings only :
				if MyFormElementArray(1) = "String" then
					if len(My_Form(Key)) > int(MyFormElementArray(3)) or len(My_Form(Key)) < int(MyFormElementArray(2)) then
						Error = Error & "Field <b>"& Key & "</b> must be between "&MyFormElementArray(2)&" and "&MyFormElementArray(3)&" characters | "
					end if
				end if
				' for numbers - Integers only :
				if MyFormElementArray(1) = "Integer" then
					if len(My_Form(Key)) > int(MyFormElementArray(3)) or len(My_Form(Key)) < int(MyFormElementArray(2)) then
						Error = Error & "Field <b>"& Key & "</b> must be between "&MyFormElementArray(2)&" and "&MyFormElementArray(2)& "|"
					end if
				end if
			end if
			' end of process the size check only if not null
		end if
	Next

	Form_Error_Handler = Error

End Function


%> 

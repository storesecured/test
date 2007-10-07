<!--#include virtual="common/connection.asp"-->
<!--#include file="include\sub.asp"-->
<% 

' ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include file="Include/Error_Template.asp"--><% 
else

	' RETRIVING FORM DATA
	First_name = Replace(Request.Form("First_name"), "'", "''")
	Last_name = Replace(Request.Form("Last_name"), "'", "")
	Website = LCase(Replace(Website, " ", ""))
	Login = Replace(Request.Form("Login"), "'", "''")
	Password = Replace(Request.Form("Password"), "'", "''")
	password_confirm = Replace(Request.Form("password_confirm"), "'", "''")
	Email = Replace(Request.Form("Email"), "'", "''")

	if Request.Form("Password") <> Request.Form("Password_confirm") then
		response.redirect "admin_error.asp?Message_id=10"
	end if

	if isNumeric(request.form("Affiliate")) then
		Affiliate = request.form("Affiliate")
	else
		Affiliate = 0
	end if

	sql_insert = "Insert into Sys_Affiliates (First_name, Last_name, Website_URL, Username, Password, Email_address) VALUES ('"&_
		First_name & "','"&Last_name&"','"&Website&"','"&Login&"','"&Password&"','"&Email&"')"
	conn_store.Execute sql_insert

	sql_login = "select Affiliate_ID from Sys_Affiliates where Username = '"&Login&"' and Password = '"&Password&"'"
	rs_Store.open sql_login,conn_store,1,1
	Affiliate_ID = rs_Store("Affiliate_ID")
	rs_Store.close

	' SET THE SESSION VARIABLE AND STATE
	Session("Affiliate_ID") = Affiliate_ID
	Session("Is_Login") = True

	response.redirect "affiliatespage.asp"
end if
%>



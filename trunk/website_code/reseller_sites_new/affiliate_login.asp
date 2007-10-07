<!--#include virtual="common/connection.asp"-->
<!--#include file="include\sub.asp"-->
<% 

' ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include file="Include/Error_Template.asp"--><% 
else

	' RETRIVING FORM DATA
	Login = Replace(Request.Form("Login"), "'", "''")
	Password = Replace(Request.Form("Password"), "'", "''")

	sql_insert = "Select Affiliate_ID from Sys_Affiliates where Username = '"&Login&"' and Password='"&Password&"'"
	rs_Store.open sql_insert,conn_store,1,1
	if rs_Store.eof then
		response.redirect "http://www.easystorecreator.com/affiliates.asp"
	else
		Session("Affiliate_ID") = rs_Store("Affiliate_ID")
		Session("Is_Login") = True
	end if
	rs_Store.close

	response.redirect "http://refer.easystorecreator.com/affiliatespage.asp"
end if
%>



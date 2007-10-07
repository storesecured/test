<!--#include file="header.asp"-->
<%

if Request.Form("Form_Name") = "Log_Out_Yes" then
	Call LogOut_Customer
	response.write "You have been successfully logged out."
end if

if Request.Form("Form_Name") = "Log_Out_No" then
	response.write "The log out has been cancelled."
end if

%>
<!--#include file="footer.asp"-->

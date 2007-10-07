<%
	msel9=3
	msel10=2
	'code here to end the session of the reseller here
	session("ESCId") = ""
	
	Response.Redirect "ESC_login.asp"
	Response.End
%>
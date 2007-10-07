<%
	'code here to end the session of the reseller here
	session("ResellerID") = ""
	
	Response.Redirect "reseller_login.asp"
	Response.End
%>
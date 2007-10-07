<%

if Session("AFFILIATE_SESSION")="TRUE" then
	'LOGOUT AFFILIATE IF LOGED IN FROM MORE THAT 1000 MINUTES
	if DateDiff("n", Session("AFFILIATE_LOGIN_TIME"), Now()) > 1000 then 
		Session("AFFILIATE_SESSION")="FALSE"
		Response.Redirect "affiliates_login.asp"
	end if								
	Store_id = Session("AFFILIATE_STORE_ID")
else
	'AFFILIATE IS NOT LOGED IN,
	'REDIRECT AFFILIATE TO LOGIN WINDOW
	Session("Is_Login") = False
	Response.Redirect "affiliates_login.asp"
end if

%>

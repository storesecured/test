<!--#include file="include/header_noview.asp"-->

<%
Server.ScriptTimeout = 900
ReturnTo = fn_get_querystring("ReturnTo")
ReturnTo = unencode(ReturnTo)

'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	fn_redirect Switch_Name&"form_error.asp?Error_Log="&server.urlencode(Error_Log)
else
	'CHECK WHAT FORM DID WE CAME FROM

	'LOGOUT REQUEST
	if Request.Form("Form_Name") = "Log_Out_Yes" then
		Call LogOut_Customer
		fn_redirect Switch_Name&"Show_Big_Cart.asp"
	end if
	
	if Request.Form("Form_Name") = "Log_Out_No" then
		fn_redirect Switch_Name&"store.asp"
	end if

	'LOGIN REQUEST
	if Request.Form("Form_Name") = "Check_Out" then
	    if Request.Form("User_id") = "" or Request.Form("Password") = "" then
            fn_error "You must enter both a username and password."
	    end if

		'AUTHENTIFICATE USER
		sql_Auth = "exec wsp_customer_login "&Store_Id&","&Shopper_Id&",'"&checkStringForQ(Request.Form("User_id"))&"','"&checkStringForQ(Request.Form("Password"))&"',-1;"
		on error resume next
		conn_store.execute sql_Auth
		if err.number<>0 then
		    fn_error err.description
		end if
		
		if AllowCookies=-1 then
			%><!--#include file="include/cookie.asp"--><%
			if request.form("SaveCookie")<>"" then
				call saveUserCookie(Request.Form("User_id"), Request.Form("Password"))
			else
				call deleteUserCookie()
			end if
		end if
		
		if request.form("Redirect") <> "" then
			fn_redirect request.form("Redirect")
		else
		    fn_redirect Switch_Name&"Login_Thank_You.asp"
		end if
	end if
end if


%>

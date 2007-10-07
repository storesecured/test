<!--#include file="header_noview.asp"-->
<%
if Request.Form("Form_Name") = "Forgot" then
		sql_Auth = "select Cid, user_id, password from Store_Customers WITH (NOLOCK) where First_Name = '"&checkStringForQ(Request.Form("First_Name"))&"' AND Last_Name = '"&checkStringForQ(Request.Form("Last_Name"))&"' AND Email = '"&checkStringForQ(Request.Form("Email"))&"' AND Store_Customers.Store_id="&Store_id&" and user_id<>'--ExpressCheckout Customer--'"

		rs_store.open sql_Auth,conn_store,1,1
		if rs_store.BOF = True or rs_store.EOF = True Then
			'NAME AND EMAIL NOT FOUND
			fn_redirect Switch_Name&"error.asp?Message_id=9"
		else
			'NAME AND EMAIL FOUND, EMAIL LOGIN INFO TO THE USER
			user_id = Rs_store("user_id")
			password = Rs_store("password")

			Subject = Store_name&", Password Reminder"
			Body = "Here is the login information you requested."&chr(10)&"Login: "&user_id&chr(10)&"Password: "&Password
			Call Send_Mail(Store_Email,Request.Form("Email"),Subject,Body)
			'REDIRECT TO FORGOT THANK YOU PAGE
			fn_redirect Switch_Name&"forgot_thank_you.asp"
		end if
		rs_store.Close
end if
%>
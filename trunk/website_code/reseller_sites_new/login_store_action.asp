<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<%

'response.redirect "maint.htm"

'on error resume next
'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include file="Include/Error_Template.asp"--><%
else

	'RETRIEVE FOR DATA
	Store_User_Id = checkStringForQ(Request("Store_User_Id"))
	Store_Password = checkStringForQ(Request("Store_Password"))
	

	'CHECK USER ID AND PASSWORD
	sql_login = "select store_logins.Store_id, Login_Privilege, Access_Priviledge, server from store_logins inner join store_settings on store_logins.store_id =  store_settings.store_id where store_logins.Store_User_Id = '"&Store_User_Id&"' and Store_logins.Store_Password = '"&Store_Password&"'"
	rs_Store.open sql_login,conn_store,1,1
	if rs_store.bof = true then 
		'INVALID LOGIN
		Session("Is_Login") = false
		rs_Store.Close
		Response.Redirect "http://manage.storesecured.com/admin_error.asp?Message_id=7"
	else
	   server_id=rs_Store("Server")
      end if
End If %>

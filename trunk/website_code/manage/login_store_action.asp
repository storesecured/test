<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include virtual="common/crypt.asp"-->

<%


'on error resume next
'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include virtual="common/Error_Template.asp"--><%
else

	'RETRIEVE FOR DATA
	Store_User_Id = checkStringForQ(Request.Form("Store_User_Id"))
	Store_Password = checkStringForQ(Request.Form("Store_Password"))
	Store_Id = checkStringForQ(Request.Form("Store_Id"))
	if Store_User_Id="" then
	   'check querystring
           Store_User_Id=decrypt(fn_get_querystring("1"))
           Store_Password=decrypt(fn_get_querystring("2"))
	end if
	if not isNumeric(store_id) then
	       Store_Id=0
	end if

	'CHECK USER ID AND PASSWORD
	sql_login = "select store_logins.Store_id, Login_Privilege, Access_Priviledge, server from store_logins WITH (NOLOCK) inner join store_settings on store_logins.store_id =  store_settings.store_id where store_logins.Store_User_Id = '"&Store_User_Id&"' and Store_logins.Store_Password = '"&Store_Password&"' and Store_logins.Store_Id = "&Store_Id
	session("sql")=sql_login
   rs_Store.open sql_login,conn_store,1,1
	if rs_store.bof = true then 
		'INVALID LOGIN
		rs_Store.Close
		Response.Redirect "admin_error.asp?Message_id=7"
	else
	   server_id=rs_Store("Server")
      
		'VALID LOGIN, SETTING THE SESSION VARIABLES
		Store_Id	= Rs_store("Store_id")
		Session("store_id")=Store_Id
		if Store_Id=217 then
		   Session("Super_User") = 1
		end if
		Session("Login_Privilege") = Rs_store("Login_Privilege")
		Session("Access_Priviledge") = Rs_store("Access_Priviledge")
		Session("User_Name")=Store_User_Id
    rs_Store.Close
		sql_update = "Update Store_Settings WITH (ROWLOCK) set Last_Login = '"&Now()&"',Warning_Sent=0 where store_Id = "&Store_Id
		conn_store.Execute sql_update
		
		sql_update = "insert into store_access_log WITH (ROWLOCK) (store_id,clientip,user_name) values ("&Store_Id&",'"&Request.ServerVariables("REMOTE_ADDR")&"','"&Store_User_Id&"')"
                conn_store.Execute sql_update

		on error resume next
		if request.form("Save_Info")="TRUE" then

		      response.cookies("STORESECURED")("Store_User_Id") = Store_User_Id
		      response.cookies("STORESECURED")("Store_Password") = Store_Password
		      response.cookies("STORESECURED")("Store_Id") = Store_Id
		      response.cookies("STORESECURED")("Save_Info") = "checked"
		      response.cookies("STORESECURED").expires = DateAdd("d",360,Now())
		else
		      
                end if

	end if
    %>
    <script src="https://ssl.google-analytics.com/urchin.js" type="text/javascript">
    </script>
    <script type="text/javascript">
    _uacct = "UA-1343888-1";
    _udn="none";
    _ulink=1;
    urchinTracker();
    </script>
    <%

    sType = "Login"
    if not isNumeric(Store_Id) or Store_Id="" or isNull(Store_Id) then
       fn_redirect ("admin_error.asp?Message_id=7")
    else
        sQuerystring=fn_get_querystring("ReturnTo")
        if sQuerystring<>"" then
            fn_redirect(sQuerystring)
        else
            fn_redirect("admin_home.asp")
        end if 
    End If
end if
%>


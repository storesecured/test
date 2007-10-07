<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

if Session("Super_User") <> 1 then
     Response.redirect "noaccess.asp"
end if

if request("Store_Id") <> "" then

    sql_select = "select store_id from store_settings where store_id="&request("Store_Id")
    rs_Store.open sql_select,conn_store,1,1
    if rs_store.bof and rs_Store.eof then
        rs_Store.close
        fn_error("No matching store found")
    else
        Session("Store_id")  = request("Store_Id")
        Session("Login_Privilege") = 1
        Session("Is_Login") = False
        Session("Super_User") = 1

        Session("Preview")=""
        sType = "Login"
        on error goto 0
        rs_store.close

        fn_redirect("admin_home.asp")
    end if
end if

sFormAction = "login_as.asp"
sCommonName = "Store Id"
sFullTitle="Admin > Login As"
sTitle = "Login As"
thisRedirect = "login_as.asp"
sMenu = "admin"
createHead thisRedirect

%>
<tr bgcolor='#FFFFFF'>
					<td width="24%" height="23" class="inputname"><B>Store Id</B></td>
					<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Store_Id" value="" size="5" maxlength=200>
							</td>
					</tr>

<% createFoot thisRedirect,1 %>


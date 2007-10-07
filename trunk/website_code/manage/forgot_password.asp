<!--#include virtual="common/connection.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/sub.asp"-->
<% 


sFormAction = "forgot_password.asp"
sTitle = "Forgot Login Info"
thisRedirect = "forgot_password.asp"
sMenu = "none"
sTopic = "Login"
createHead thisRedirect

if request.form <> "" then
	Store_Email=request.form("Store_Email")

	sql_login = "select store_logins.Store_User_Id, store_logins.Store_Password, store_logins.Store_id from store_logins inner join store_settings on store_logins.store_id=store_settings.store_id where ((Store_Settings.Store_email='"&Store_Email&"' and Store_logins.store_email is Null) or (Store_logins.store_email='"&Store_Email&"'))"
        rs_Store.open sql_login,conn_store,1,1
	if rs_store.bof = true then
		response.write "<TR bgcolor=FFFFFF><TD>The email address "&Store_Email&" does not match any of our current stores.  The information cannot be retrieved.  Please try again with a different email address or contact "&sSupport_email&" for help finding your password.</td></tr>"
		response.write "<TR bgcolor=FFFFFF><TD><a href=forgot_password.asp class=link>Try Again</a></td></tr>"
	else
	       sText = "You requested the following login information from StoreSecured."
	       do while not rs_store.eof
		      sText = sText &vbcrlf&vbcrlf&"**************************************"&vbcrlf&vbcrlf&_
                              "Site #: "&rs_Store("Store_Id")&vbcrlf&vbcrlf&_
                              "Username: "&rs_Store("Store_User_Id")&vbcrlf&vbcrlf&_
		              "Password: "&rs_Store("Store_Password")&vbcrlf&vbcrlf&_
		              "Login to manage your store at http://manage.storesecured.com"
		      rs_store.movenext
                loop
                Send_Mail sNoReply_email,Store_Email,"StoreSecured Requested Info",sText
		response.write "<TR bgcolor=FFFFFF><TD>Your login information has been sent to "&Store_Email&"</td></tr>"
		response.write "<TR bgcolor=FFFFFF><TD><a href=login_store.asp class=link>Click here to login</a></td></tr>"


        end if
	rs_store.close
else
%>



		 <TR bgcolor=FFFFFF>
		 <td class="inputname"><B>Email</b></td>
			 <td class="inputvalue">
			 <input type="text" name="Store_Email" size="60" value="">
			 <% small_help "Email" %></td>
		 </tr>

	 <TR bgcolor=FFFFFF>

			 <td width="17%" align="center">&nbsp;</td>
			 <td width="84%" align="left" valign="top"><br>
			 <input type="submit" class="Buttons" value="Get Login Info" name="I1"></td>
			 <td width="84%" align="left" valign="top">&nbsp;</td>
		 </tr>
<% end if %>
<% createFoot thisRedirect,0 %>

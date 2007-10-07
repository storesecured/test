<!--#include file="include/header.asp"-->

<%


if No_Login then
	fn_redirect Switch_Name&"Register.asp?ExpressCheckout=True&ReturnTo="&server.urlencode(unencode(fn_get_querystring("ReturnTo")))
end if
%>
<form method="POST" action="<%=Switch_Name %>check_out_action.asp" name=login>
<% if "http://"&Request.ServerVariables("HTTP_HOST")&"/"=Switch_Name then %>
        <Input type="hidden" name="Shopper_Id" value="<%=Shopper_Id%>">
<% end if %>
<input type="Hidden" name="Form_Name" value="Check_Out">
<input type="Hidden" name="Redirect" value=<%= unencode(fn_get_querystring("ReturnTo")) %>>
<% if AllowCookies=-1 then %>
	<!--#include file="include/cookie.asp"-->
	<% if isUserCookie() then %>
		<% call getUserCookie(theUserName, theUserPassword) %>
		<% saveChecked = "checked" %>
	<% end if %>
<% end if %>
<% if fn_get_querystring("Protected") = "Yes" and cid > 0 then %>
<table border="0" width="100%" cellspacing=0 cellpadding=0 border=1>
<TR><TD valign=top colspan=2>
<B>Notice: This page has been protected by the store owner.  You do not currently have access, 
please contact the store owner for instructions on how to obtain access.
</b><BR><BR></td></tr></table>
<% else %>
<table border="0" width="100%" cellspacing=0 cellpadding=0>
<TR><TD valign=top width=250>
<B>New Users</b>
<% ReturnTo = fn_get_querystring("ReturnTo") %>
<BR><BR>
<a href="<%=Switch_Name %>Register.asp?ReturnTo=<%= server.urlencode(fn_get_querystring("ReturnTo")) %>" class='link'>Click here to Register</a>
<% If ExpressCheckout then %>
<BR><BR>
OR
<BR><BR>
<a href="<%=Switch_Name %>Register.asp?ExpressCheckout=True&ReturnTo=<%= server.urlencode(fn_get_querystring("ReturnTo")) %>" class='link'>Checkout without registering</a>
<% end if %>
</TD><TD width=30>&nbsp;</td><TD width=1 bgcolor=000000><img src=spacer.gif width=1></td><TD width=30>&nbsp;</td><TD>
<B>Returning Users</b>
<BR><BR>
<table>
<tr>
		<td width="24%" height="27" class='normal'>Login</td>
		<td width="76%" height="27" colspan="2" class='normal'>
			<input type="text" name="User_id" size="22" value="<%= theUserName %>">
			<input type="hidden" name="User_id_C" value="Re|String|0|50|||Login" ></td>
	</tr>
	<tr>
		<td width="24%" height="25" class='normal'>Password</td>
		<td width="76%" height="25" colspan="2">
			<input type="password" name="Password" size="22" value="<%= theUserPassword %>">
			<input type="hidden" name="Password_C" value="Re|String|0|50|||Password" ></td>
	</tr>
	<tr>
		<td width="24%" height="27">&nbsp;</td>
		<td width="38%" height="27" colspan="2">
				  <%= fn_create_action_button ("Button_image_Login", "Login", "Login") %>

			&nbsp;<a href="<%=Switch_Name %>forgot.asp" class='link'>Forgot</a></td>
	</tr>
	
	<% if AllowCookies=-1 then %>
		<tr>
			<td colspan="3" class='normal'>

				<input type="checkbox" name="SaveCookie" value="TRUE" <%= saveChecked %>>
				Save login and password for future access<br>&nbsp;
			</td>
		</tr>
	<% end if %>
</table>
</TD></TR>
	</table>
<% end if %>
</form>
<%
sFormName = "login"
sSubmitName = "Login" 
%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator("<%= sFormName %>");
 frmvalidator.addValidation("User_id","req","Please enter a login.");
 frmvalidator.addValidation("Password","req","Please enter a password.");
</script>
<!--#include file="include/footer.asp"-->

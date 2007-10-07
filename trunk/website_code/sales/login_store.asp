<%
quicktext="Current Customer Login"

%>
<!--#include file="header.asp"-->
		<h2>Current Customer Login</h2>


<form method="POST" action="http://manage.storesecured.com/Login_Store_Action.asp">
<table border="0" cellspacing="0" cellpadding="0" width="200">
<tr>
<td valign="top" class="bluetext" width="60">Site #</td>
<td width="15">&nbsp;</td>
<td width="125" valign="top"><input type="text" size="21" style="border: #889DCD 1px solid" name="Store_Id"></td>
</tr>

<tr>
<td valign="top" class="bluetext" width="60">Username</td>
<td width="15">&nbsp;</td>
<td width="125" valign="top"><input type="text" size="21" style="border: #889DCD 1px solid" name="Store_User_Id"></td>
</tr>

<tr>
<td colspan="3" height="3"></td>
</tr>

<tr>
<td valign="top" class="bluetext" width="60">Password</td>
<td width="15">&nbsp;</td>
<td width="125" valign="top"><input type="password" size="21" style="border: #889DCD 1px solid" name="Store_Password"></td>
</tr>

<tr>
<td colspan="3" height="4"></td>
</tr>

<tr>
<td></td>
<td></td>
<td><input type="image" src="images/login_inner.gif" name="Login">&nbsp;<input type="image" src="images/secure_login_inner.gif" name="Secure"></td>
</tr>
<Tr><td colspan=3><a href=http://manage.storesecured.com/forgot_password.asp>If you have forgotten your Username or Password please click here</a></td></tr>
</table>
</form>
<!--#include file="footer.asp"-->


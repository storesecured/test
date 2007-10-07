<!--#include file="include/header.asp"-->
<% sFormName="cancel_signup" %>
<form method="POST" action="<%= Site_Name %>cancel_signup_action.asp" name="<%= sFormName %>">
<table border="0" width="82%">
      <tr>
		<td class='normal' colspan=2><b>Newsletter Cancel</b></td>
	</tr>
	<tr>
		<td class='normal' colspan=2>Enter your email address below to be removed from our newsletter.</td>
	</tr>
	<tr>
		<td width="10%" class='normal' nowrap>Email Address</td>
		<td width="90%" class='normal'><input type="text" name="Email_Address" size="40" value="<%= fn_get_querystring("Email_Address")%>"></td>
	</tr>


	<tr>
		<td width="10%" class='normal'></td>
		<td width="90%" class='normal'>
			<%= fn_create_action_button ("Button_image_Continue", "Continue", "Continue") %>
			</td>
	</tr>

</table>

</form>

<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator("<%= sFormName %>");
 frmvalidator.addValidation("Email_Address","req","Please enter a email address.");
 frmvalidator.addValidation("Email_Address","email","The email address you enter must be a valid email.");
</script>

<!--#include file="include/footer.asp"-->

<!--#include file="include/header.asp"-->

<form method="POST" action="<%= Switch_Name %>include/forgot_action.asp" name="Forgot">
<input type="Hidden" name="Form_Name" value="Forgot">
<% small_text_size = text_size -1 %>

<table border="0" width="46%" height="113">
	<tr>
		<td width="100%" height="27" colspan="3" class='normal'>Enter the following information and we will email you your login and password</td>
	</tr>
    
	<tr>
		<td width="24%" height="27" class='normal' nowrap>First Name</td>
		<td width="76%" height="27" colspan="2" class='normal'>
			<input type="text" name="First_name" size="22"></td>
	</tr>
    
	<tr>
		<td width="24%" height="25" class='normal' nowrap>Last Name</td>
		<td width="76%" height="25" colspan="2" class='normal'>
			<input type="text" name="Last_name" size="22"></td>
	</tr>
    
	<tr>
		<td width="24%" height="25" class='normal' nowrap>Email</td>
		<td width="76%" height="25" colspan="2" class='normal'>
			<input type="text" name="EMail" size="22"></td>
	</tr>

	<tr>
		<td width="24%" height="27" class='normal'>&nbsp;</td>
		<td width="38%" height="27" class='normal'>
			<%= fn_create_action_button ("Button_image_Continue", "Remind_me", "Continue") %>
			</td>
		<td width="38%" height="27"></td>
	</tr>

</table>

</form>
<%
sFormName = "Forgot"
sSubmitName = "Remind_me"
%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator("<%= sFormName %>");
 frmvalidator.addValidation("First_name","req","Please enter a first name.");
 frmvalidator.addValidation("Last_name","req","Please enter a last name.");
 frmvalidator.addValidation("EMail","req","Please enter a email address.");
 frmvalidator.addValidation("EMail","email","Please enter a email address.");
</script>

<!--#include file="include/footer.asp"-->

<!--#include file="include/header.asp"-->

<form method="POST" action="Save_Cart_action.asp" name=Save_Cart>
<table border="0" width="82%">

	<tr>
		<td width="74%" class='normal'>Your Email</td>
		<td width="55%" class='normal'>
      <input type="text" name="Send_to_email" size="60" maxlength=100>
     <INPUT type="hidden"  name=Send_to_email_C value="Re|Email|0|100|||Email"><br><%= enter_email_address %> </font></td>
	</</td>
	</tr>
	
	<tr>
		<td width="74%" class='normal'><hr></td>
		<td width="55%" class='normal'><hr></td>
	 </tr>

	 <tr>
		<td width="74%" class='normal'>Send Wish List To</font></td>
		<td width="55%" class='normal'>
      <input type="text" name="cc_to_email" size="60" maxlength=100>
      <INPUT type="hidden"  name=cc_to_email_C value="Op|Email|0|100|||Send To"><br><%= enter_email_address %> </font></td>
	</tr>

	<tr>
		<td width="74%" class='normal'>With Message</td>
		<td width="55%" class='normal'>
			<textarea name="message_cart" cols="45" rows="6"></textarea></td>
	</tr>

	<tr>
		<td width="74%" class='normal'></td>	
		<td width="55%" class='normal'>
			<%= fn_create_action_button ("Button_image_SaveCart", "Save_Cart", "Save Cart") %>
			</td>
	</tr>

</table>

</form>
<%
sFormName = "Save_Cart"
sSubmitName = "Save_Cart"
%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator("<%= sFormName %>");
 frmvalidator.addValidation("Send_to_email","req","Please enter your email.");
 frmvalidator.addValidation("Send_to_email","email","Please enter a valid email.");
 frmvalidator.addValidation("cc_to_email","email","Please enter a valid email.");
 frmvalidator.addValidation("message_cart","req","Please enter a message.");

</script>

<!--#include file="include/footer.asp"-->

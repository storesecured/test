<!--#include file="include/header.asp"-->

<form method="POST" action="signup_action.asp" name=signup>
<table border="0" cellspacing=0 cellpadding=5>
     <tr>
		<td class='normal' colspan=2><b>Newsletter Signup</b></td>
	</tr>
	<tr>
		<td class='normal' colspan=2>Enter your email address below to signup for our newsletter.</td>
	</tr>
	<tr>
		<td class='normal'>First Name</td>
		<td class='normal'><input type="text" name="First_Name" size="20">
          <INPUT type="hidden"  name=First_Name_C value="Re|String|0|50|||First Name"></td>
	</tr>
	<tr>
		<td class='normal'>Last Name</td>
		<td class='normal'><input type="text" name="Last_Name" size="20">
          <INPUT type="hidden"  name=Last_Name_C value="Re|String|0|50|||Last Name"></td>
	</tr>
	<tr>
		<td class='normal'>Email Address</td>
		<td class='normal'><input type="text" name="Email_Address" size="20">
          <INPUT type="hidden"  name=Email_Address_C value="Re|Email|0|50|||Email Address"></td>
	</tr>

	<tr>
		<td class='normal'></td>
		<td class='normal'>
			<%= fn_create_action_button ("Button_image_Continue", "Continue", "Continue") %>
			</td>
	</tr>

</table>

</form>

<% 
sSubmitName = "Continue"
sFormName = "signup" 
%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator("<%= sFormName %>");
 frmvalidator.addValidation("First_Name","req","Please enter a first name.");
 frmvalidator.addValidation("Last_Name","req","Please enter a last name");
 frmvalidator.addValidation("Email_Address","req","Please enter a valid email.");
 frmvalidator.addValidation("Email_Address","email","Please enter a valid email.");
 </script>

<!--#include file="include/footer.asp"-->

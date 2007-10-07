<!--#include file="include/header.asp"-->
	<font class='normal'>
	<%= fn_display_address(0) %>
</font>
	<form method="POST" action="<%= Site_Name %>Contact_Action.asp" name="ContactUs">
	<table>
	<tr><td class='normal' nowrap>Name</td>
	<td class='normal'><input type=text name=fromname size=30></td></tr>
	<tr><td class='normal' nowrap>Reply Email</td>
	<td class='normal'><input type=text name=fromemail size=30></td></tr>
	<!--<tr><td class='normal' nowrap>To Email</td>
	<td class='normal'><%= Display_email %></td></tr>-->
	<tr><td colspan=2 class='normal'><textarea rows="8" name="message" cols="50"></textarea></td></tr>
	<tr><td colspan=2 class='normal'><%= fn_create_action_button ("Button_image_Continue", "Update", "Continue") %></td></tr>
	</table>
	</form>
</font>
<%
sSubmitName = "Update"
sFormName = "ContactUs"
%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator("<%= sFormName %>");
 frmvalidator.addValidation("fromname","req","Please enter your name.");
 frmvalidator.addValidation("fromemail","req","Please enter a reply email.");
 frmvalidator.addValidation("fromemail","email","Please enter a reply email.");
 frmvalidator.addValidation("message","req","Please enter your message.");

</script>

<!--#include file="include/footer.asp"-->

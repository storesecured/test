<!--#include file="include/header.asp"-->
<% if cid=0 then
   'if user is not logged in redirect to homepage
   fn_redirect Switch_Name
end if %>

<form method="POST" action="<%= Switch_Name %>include/log_out_action.asp" name=log_out_yes>
<input type="Hidden" name="Form_Name" value="Log_Out_Yes">

	<table border="0" align="center" cellspacing="0">
		<tr>
			<td class='big'"><b>Warning</b></td>
		</tr>
		
		<tr>
			<td width="100%" colspan="2" class='normal'>Are you sure you would like to log out?</td>
		</tr>

		<tr>
			<td width="50%" align="right" class='normal'>
				<%= fn_create_action_button ("Button_image_Yes", "Log_Out_Yes", "Yes") %>
			</td></form>
			<form method="POST" action="<%= Switch_Name %>include/log_out_action.asp" name=log_out_no>
			<input type="Hidden" name="Form_Name" value="Log_Out_No">
     <td width="50%" class='normal'>
				<%= fn_create_action_button ("Button_image_No", "Log_Out_No", "No") %>
			</td>
		</tr>
	</table>

</form>
<%
      sFormName1 = "log_out_yes"
      sFormName = "log_out_yes,log_out_no"
      sSubmitName = "Log_Out_Yes,Log_Out_No"

      %>
      <SCRIPT language="JavaScript">
       var frmvalidator  = new Validator("<%= sFormName1 %>");

      </script>


<!--#include file="include/footer.asp"-->

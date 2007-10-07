<!--#include file="include/header.asp"-->

<% ' RECORD_TYPA PATAM : 0=BILLING; 1=SHIPPING ADDRESS 1; 2=SHIPPING ADDRESS 2; ETC.

if Spam <> 0 then
	checked_spam = "checked" 
end if

%>  

<script language="JavaScript">

	function checkAddress()
	{
		selVal = "";
		for (i=0;i<document.forms['adrsel'].addrs.length;i++)
			if (document.forms['adrsel'].addrs[i].selected)
				selVal=document.forms['adrsel'].addrs[i].value;
		document.location = "Modify_my_Account.asp/ssadr/"+selVal;}

</script>
 
<% if not (User_id=EXPRESS_CHECKOUT_CUSTOMER and ExpressCheckout) then %>
 
<form method="POST" action="include/update_records_action.asp" name=modify_account>
<input type="Hidden" name="Form_Name" value="modify_account">
<input type="Hidden" name="Page_id" value="<%= Page_id %>">
<input type="Hidden" name="Record_Type" value="0">
	
	<table border="0" width="100%" height="222" cellspacing="0">
		<tr>
			<td height="171">
				<table border="0" width="400">
					
					<tr> 
						<TD width="400" colspan="2" class='big'> <b>Modify My Account</b></TD>
					</tr>
					
					<tr> 
						<td width="150" class='normal'>Login</td>
						<td width="154" class='normal'>
							<INPUT type="hidden"  name=User_Id_C value="Re|String|0|50|||Login">
							<input name="User_id" size="20" value="<%= User_id %>" maxlength=50></td>
					</tr>

					<tr>
						<td width="150" class='normal'>Password</td>
						<td width="154" class='normal'>
							<input type="Password" name="Password" value="<%= Password %>" size="20" maxlength=50>
							<INPUT name=Password_C type=hidden value="Re|String|0|50|||Password"></td>
					</tr>

					<tr>
						<td width="150" class='normal'>Confirm Password</td>
						<td width="154" class='normal'>
							<input type="Password" name="Password_confirm" size="20" maxlength=50>
							<INPUT name=Password_C type=hidden value="Re|String|0|50|||Confirm Password"></td>
					</tr>
          <% if Budget_left > 0 then %>
					<tr>
						<td width="150" class='normal'>Budget Left</td>
						<td width="154" class='normal'><%= Currency_Format_Function(Budget_left) %> &nbsp;
					<a class="link" href="budget_view_cust.asp?cid=<%=cid%>">History </a></td>

               </tr>
					<% end if %>
					<% if Enable_Rewards then %>
					<tr>
						<td width="150" class='normal'>Rewards Left</td>
						<td width="154" class='normal'><%= Currency_Format_Function(Reward_Left) %></td>
					</tr>
					<% end if %>
					<tr>
						<td width="150" class='normal'>Promotional Emails</td>
						<td width="154" class='normal'>
							<input type="checkbox" <%= checked_spam %> name="Spam" value="-1" ></td>
					</tr>

					<tr>
						<td width="154" class='normal'>

						  </td>
                                                <td width="150" class='normal'>

                                                        <%= fn_create_action_button ("Button_image_Update", "Modify_my_Account", "Update") %></td>
					</tr>
				</table>
			</form>
			<%
      sFormName = "modify_account"
      sSubmitName = "Modify_my_Account"
      %>
      <SCRIPT language="JavaScript">
       var frmvalidator  = new Validator("<%= sFormName %>");
       frmvalidator.addValidation("User_id","req","Please enter a login.");
       frmvalidator.addValidation("User_id","alnumhyphen","Login can only contain letters and numbers");
       frmvalidator.addValidation("Password","req","Please enter a password.");
       frmvalidator.addValidation("Password","alnumhyphen","Password can only contain letters and numbers");
       frmvalidator.addValidation("Password_confirm","req","Please confirm your password.");
       frmvalidator.setAddnlValidationFunction("DoCustomValidation");

        function DoCustomValidation()
        {
          var frm = document.forms["<%= sFormName %>"];
          if (document.forms["<%= sFormName %>"].Password.value == document.forms["<%= sFormName %>"].Password_confirm.value)
          {
            return true;
          }
          else
          {
            alert('The 2 passwords must match.');
            return false;
          }
        }
      </script>

		<a href="Modify_my_Billing.asp" class='link'>Change Billing Address</a><BR>
		<% if Show_shipping then %>
		<a href="Modify_my_Shipping.asp" class='link'>Change Shipping Address</a><BR>
		<% end if %>
</td>
</tr>
</table>
<% else %>
You did not create an account.
<% End If %>

<!--#include file="include/footer.asp"-->

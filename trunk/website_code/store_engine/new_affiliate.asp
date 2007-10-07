
<!--#include file="include/header.asp"-->

<% 

checked1="checked"
checked2=""

 %>
 


<form method="POST" action="<%= Site_Name %>Affiliates_Action.asp" name="new_affiliates">
	<input type="hidden" name="ACTION" value="ADD_AFFILIATE">

<table border="0" width="100%" cellpadding="0" cellspacing="0">

   


	<tr>
		<td width="100%" colspan="3" height="11">
			<table border="0" width="100%" height="1">
				<tr>
					<td width="1%" height="12" class='normal' nowrap>Login</td>

						<td height="12" class='normal'>

							<input type="hidden" name="Store_id" value="<%= Store_id %>">
							<INPUT type="text" name="Code" value="<%= Code %>">
							</td>

				</tr>
    
				<tr>
					<td width="1%" height="12" nowrap class='normal'>Contact Name</td>
					<td height="12" class='normal'>
						<input type="text" name="Contact_Name" size="20" value="<%= Contact_Name %>" maxlength="50">
						<INPUT type="hidden" name="Contact_Name_C" value="Re|String||||"></td>
			</tr>
	
				<tr>
					<td width="1%" height="12" nowrap class='normal'>Email</td>
					<td height="12" class='normal'>
						<input type="text" name="Email" size="20" value="<%= Email %>" maxlength="50">
						<INPUT type="hidden" name="Email_C" value="Re|String||||"></td>
				</tr>
	
				<tr>
					<td width="1%" height="12" nowrap class='normal'>URL</td>
					<td height="12" class='normal'>
						<% if aff_id<>"" then %>
							<input type="text" name="URL" size="20" value="<%= URL %>" maxlength="255">
						<% Else %>
							<input type="text" name="URL" size="20" value="http://" maxlength="255">
						<% End If %>
						<INPUT type="hidden" name="URL_C" value="Re|String||||"></td>
				</tr>
	
				<tr>
					<td width="1%" height="12" nowrap class='normal'>Email Notifications</td>
					<td height="12" class='normal'>
						<input type="radio" name="Email_Not" value="-1" <%= checked1 %>>Yes&nbsp;
						<input type="radio" name="Email_Not" value="0" <%= checked2 %>>No</td>
				</tr>

				<tr>
					<td width="1%" height="12" class='normal' nowrap>Password</td>
					<td height="12">
						<input type="password" name="Password" size="20" value="<%= Password %>" maxlength="50">
						<INPUT type="hidden" name="Password_C" value="Re|String||||"></td>
				</tr>
    
				<tr>
					<td width="1%" height="12" class='normal' nowrap>Confirm Password</td>
					<td height="12">
						<input type="password" name="Password_Confirm" size="20" value="<%= Password %>" maxlength="50"></td>
			</tr>
	
				<tr>
					<td width="1%" height="12" class='normal' nowrap>&nbsp;</td>
					<td height="12" class='normal'>
						<input type="submit" value="Add" name="Add_Affiliate">

					</td>
			</tr>
			</table>
		</td>
	</tr>
</table>

</form>


<!--#include file="include/footer.asp"-->

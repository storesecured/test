<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%

if request.querystring("Id") <> "" then 
	'IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	Login_Id = request.querystring("Id")
	if not isNumeric(Login_Id) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	Str_Select = "select * from Store_Logins where Login_Id = "&Login_Id&" and Store_Id ="&Store_Id
	rs_Store.open Str_Select,conn_store,1,1
		if rs_store.eof or rs_store.bof then
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		Store_User_id = rs_store("Store_User_id")
		Store_Password = rs_store("Store_Password")
		Store_First_Name = rs_store("Store_First_Name")
		Store_Last_Name = rs_store("Store_Last_Name")
		Store_Email = rs_store("Store_Email")
		Login_Privilege_This = rs_store("Login_Privilege")
		Access_Priviledge_This = rs_store("Access_Priviledge")
	rs_Store.close
	'if this is an admin check to see if it is the only one and if so don't allow changing login priviledge
	if checked1 = "selected" then
		Str_Select = "select count(Login_Privilege) as admins from Store_Logins where Login_Privilege = 1 and Store_Id ="&Store_Id
		rs_Store.open Str_Select,conn_store,1,1
		iAdminCnt = rs_Store("admins")
		rs_Store.close
	end if
end if

If request.querystring("Id") <> "" then
	 sTitle = "Edit Admin Login"
	 sFullTitle = "My Account > <a href=security.asp class=white>Admin Logins</a> > Edit - "&Store_User_id
	 sFormName = "Security_Edit"
Else
	 sTitle = "Add Admin Login"
	 sFullTitle = "My Account > <a href=security.asp class=white>Admin Logins</a> > Add"
	 sFormName = "Security_Add"
End If
sCancel="security.asp"
sCommonName="Admin Login"
sFormAction = "Store_Settings.asp"
sName = "Add_Store_Log_ins"
thisRedirect = "security_add.asp"
sSubmitName = "Add_Store_Log_ins"
sMenu = "account"
sQuestion_Path = "advanced/login_1.htm"
createHead thisRedirect

%>



<% If request.querystring("Id") <> "" then %>
	<input type="Hidden" name="Login_Id" value="<%= Request.querystring("Id") %>">
<% End If %>

				 
		


				<tr bgcolor='#FFFFFF'>
				<td width="24%" class="inputname"><b>Username</b></td>
				<td width="76%" class="inputvalue">
						<input type="text" name="Store_User_Id" size="60" value="<%= Store_User_id %>">
						<INPUT type="hidden"  name=Store_User_Id_C value="Re|String|0|20|||Username">
					<% small_help "Username" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
				<td width="24%" class="inputname"><b>Password</b></td>
				<td width="76%" class="inputvalue">
						<input type="password" name="Store_Password" size="60" value="<%= Store_Password %>">
						<INPUT type="hidden"  name=Store_Password_C value="Re|String|0|20|||Password">
					<% small_help "Password" %></td>
				</tr>
		
				<tr bgcolor='#FFFFFF'>
				<td width="24%" class="inputname"><b>Confirm Password</b></td>
				<td width="76%" class="inputvalue">
						<input type="password" name="Store_Password_Confirm" size="60" value="<%= Store_Password %>">
						<INPUT type="hidden"  name=Store_Password_Confirm_C value="Re|String|0|20|||Confirm Password">
					<% small_help "Confirm Password" %></td>
				</tr>
		
				<tr bgcolor='#FFFFFF'>
				<td width="24%" class="inputname"><b>First Name</b></td>
				<td width="76%" class="inputvalue">
						<input type="text" name="Store_First_Name" size="60" value="<%= Store_First_Name %>">
						<INPUT type="hidden"  name=Store_First_Name_C value="Re|String|0|100||First Name">
					<% small_help "First Name" %></td>
				</tr>
		
				<tr bgcolor='#FFFFFF'>
				<td width="24%" class="inputname"><b>Last Name</b></td>
				<td width="76%" class="inputvalue">
						<input type="text" name="Store_Last_Name" size="60" value="<%= Store_Last_Name %>">
						<INPUT type="hidden"  name=Store_Last_Name_C value="Re|String|0|100|||Last Name">
					<% small_help "Last Name" %></td>
				</tr>
		
				<tr bgcolor='#FFFFFF'>
				<td width="24%" class="inputname"><b>Email</b></td>
				<td width="76%" class="inputvalue">
						<input type="text" name="Store_Email" size="60" value="<%= Store_Email %>">
						<INPUT type="hidden"  name=Store_Email_C value="Re|String|0|100|@,.||Email">
					<% small_help "Email" %></td>
				</tr>
				<% if Login_Privilege_This <> 1 then %>
				<tr bgcolor='#FFFFFF'>
				<td width="24%" class="inputname"><b>Access</b></td>
				<td width="76%" class="inputvalue">
				<select size="4" name="Access_Priviledge" multiple>

					<option value="General"
					<% if instr(Access_Priviledge_This,"General") > 0 then
						response.write "selected"
					end if %>
					>General</option>
					<option value="Shipping"
					<% if instr(Access_Priviledge_This,"Shipping") > 0 then
						response.write "selected"
					end if %>
					>Shipping</option>
					<option value="Design"
					<% if instr(Access_Priviledge_This,"Design") > 0 then
						response.write "selected"
					end if %>
					>Design</option>
					<option value="Marketing"
					<% if instr(Access_Priviledge_This,"Marketing") > 0 then
						response.write "selected"
					end if %>
					>Marketing</option>
					<option value="Inventory"
					<% if instr(Access_Priviledge_This,"Inventory") > 0 then
						response.write "selected"
					end if %>
					>Inventory</option>
					<option value="Statistics"
					<% if instr(Access_Priviledge_This,"Statistics") > 0 then
						response.write "selected"
					end if %>
					>Statistics</option>
					<option value="Customers"
					<% if instr(Access_Priviledge_This,"Customers") > 0 then
						response.write "selected"
					end if %>
					>Customers</option>
					<option value="Orders"
					<% if instr(Access_Priviledge_This,"Orders") > 0 then
						response.write "selected"
					end if %>
					>Orders</option>
					<option value="My Account"
					<% if instr(Access_Priviledge_This,"My Account") > 0 then
						response.write "selected"
					end if %>
					>My Account</option>
					</select>
					<% small_help "Access Privilege" %></td>
				</tr>
				<% end if %>

<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Store_User_Id","req","Please enter your username.");
frmvalidator.addValidation("Store_Password","req","Please enter your password.");
frmvalidator.addValidation("Store_Password_Confirm","req","Please enter your password again.");
frmvalidator.addValidation("Store_First_Name","req","Please enter a first name.");
frmvalidator.addValidation("Store_Last_Name","req","Please enter a last name.");
frmvalidator.addValidation("Store_Email","email","Please enter a valid email");
frmvalidator.addValidation("Store_Email","req","Please enter a valid email");
frmvalidator.setAddnlValidationFunction("DoCustomValidation"); 

function DoCustomValidation()
{
  var frm = document.forms[0];
  if (document.forms[0].Store_Password.value == document.forms[0].Store_Password_Confirm.value)
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

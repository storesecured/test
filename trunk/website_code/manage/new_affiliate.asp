<!--#include virtual="common/connection.asp"-->
<!--#include file="pagedesign.asp"-->


<% 
sTextHelp="affiliates/affiliates.doc"

' IF CALLING FROM AFFILIATES' PAGES, USE AFFILIATES_SESSION.ASP
' FOR SESSION MANAGEMENT
RUN_FROM_AFFILIATES = FALSE

	Session("AFFILIATE_SESSION")="FALSE"
	%><!--#include file="include\sub.asp"--><%


checked1="checked"
checked2=""

if RUN_FROM_AFFILIATES = TRUE then
	aff_id = Session("AFFILIATE_AFFILIATE_ID")
else
	aff_id = request.querystring("id")
	if aff_id="" then
	aff_id = request.querystring("aff_id")
	end if
	if not isNumeric(aff_id) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
end if

if aff_id<>"" then
	'IF EDIT THE LOAD CURRENT VALUES FROM THE DATABASE
	sql_sel = "select * from store_affiliates where store_id="&store_id&" and affiliate_id="&aff_id
	rs_Store.open sql_sel,conn_store,1,1
	if rs_store.bof or rs_store.eof then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
		Code = rs_Store("Code")
		Contact_Name = rs_Store("Contact_Name")
		Email = rs_Store("Email")
		URL = rs_Store("URL")
		Password = rs_Store("Password")
		Zip = rs_Store("Zip")
		State = rs_Store("State")
		Company = rs_Store("Company")
		Address = rs_Store("Address")
		City = rs_Store("City")
		Country = rs_Store("Country")
		Phone = rs_Store("Phone")
		if cint(rs_Store("Email_Notification"))=-1 then
			checked1="checked"
		else
			checked2="checked"
		end if
	rs_Store.close
end if

sFormAction = "Affiliates_Action.asp"
if aff_id<>"" then
	sTitle = "Edit Affiliate - "&Code
	sFullTitle = "Marketing > <a href=affiliates_manager.asp class=white>Affiliates</a> > Edit - "&Code
Else
	sTitle = "Add Affiliate"
	sFullTitle = "Marketing > <a href=affiliates_manager.asp class=white>Affiliates</a> > Add"

End If
sCommonName="Affiliate"
sCancel="affiliates_manager.asp"
sFormName = "Add_Affiliate"
thisRedirect = "new_affiliate.asp"
sMenu="marketing"
createHead thisRedirect
 %>
 



<% if aff_id<>"" then %>
	<input type="hidden" name="ACTION" value="EDIT_AFFILIATE">
	<input type="hidden" name="AFF_ID" value="<%= aff_id %>">
<% Else %>
	<input type="hidden" name="ACTION" value="ADD_AFFILIATE">
<% End If %>





				<TR bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Code</b></td>
					<% if RUN_FROM_AFFILIATES then %>
						<td width="90%" class="inputvalue">
							<b><%= Code %></b>
							<INPUT type="hidden" name="Code" value="<%= Code %>">
							</td>
					<% Else %>
						<td width="90%" class="inputvalue">
							<% if aff_id<>"" then %>
							<%= Code %><input type="hidden" name="Code" size="60" value="<%= Code %>" maxlength="50">
							<% else %>
							<input type="text" name="Code" size="60" value="<%= Code %>" maxlength="50">
							<INPUT type="hidden" name="Code_C" value="Re|String|0|50|||Code">
							<% end if %>
                                                        <% small_help "Code" %></td>
					<% End If %>
				</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Contact Name</b></td>
					<td width="90%" class="inputvalue">
						<input type="text" name="Contact_Name" size="60" value="<%= Contact_Name %>" maxlength="50">
						<INPUT type="hidden" name="Contact_Name_C" value="Re|String|0|50|||Contact Name">
						<% small_help "Contact Name" %></td>
			</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Email</b></td>
					<td width="90%" class="inputvalue">
						<input type="text" name="Email" size="60" value="<%= Email %>" maxlength="50">
						<INPUT type="hidden" name="Email_C" value="Re|String|0|50|@,.||Email">
						<% small_help "Email" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>URL</b></td>
					<td width="90%" class="inputvalue">
						<% if aff_id<>"" then %>
							<input type="text" name="URL" size="60" value="<%= URL %>" maxlength="255">
						<% Else %>
							<input type="text" name="URL" size="60" value="http://" maxlength="255">
						<% End If %>
						<INPUT type="hidden" name="URL_C" value="Re|String|0|255|http://||URL">
						<% small_help "URL" %></td>
				</tr>
	
				<TR bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Email Notifications</b></td>
					<td width="90%" class="inputvalue">
						<input class="image" type="radio" name="Email_Not" value="-1" <%= checked1 %>>Yes&nbsp;
						<input class="image" type="radio" name="Email_Not" value="0" <%= checked2 %>>No
						<% small_help "Email Notifications" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Password</b></td>
					<td width="90%" class="inputvalue">
						<input type="password" name="Password" size="60" value="<%= Password %>" maxlength="50">
						<INPUT type="hidden" name="Password_C" value="Re|String|0|50|||Password">
						<% small_help "Password" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Confirm Password</b></td>
					<td width="90%" class="inputvalue">
						<input type="password" name="Password_Confirm" size="60" value="<%= Password %>" maxlength="50">
						<% small_help "Confirm Password" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
		<td width="10%" class="inputname"><b>Company</TD>
		<td width="90%" class="inputvalue"><INPUT  name=company size="60" value="<%= Company %>" maxlength=100>
		<INPUT type="hidden"  name=Company_C value="Op|String|0|100|||Company">
		<% small_help "Company" %></TD>
	</TR>

	<TR bgcolor='#FFFFFF'>
		<td width="10%" class="inputname"><b>Address</TD>
		<td width="90%" class="inputvalue"><INPUT  name=Address size="60" value="<%= address %>" maxlength=200>
		<INPUT type="hidden" name=Address_C value="Re|String|0|200|||Address">
		<% small_help "Address" %></TD>
	</TR>

	<TR bgcolor='#FFFFFF'>
		<td width="10%" class="inputname"><b>City</TD>
		<td width="90%" class="inputvalue"><INPUT  name=City size="60" value="<%= city %>" maxlength=200>
		<INPUT type="hidden"  name=City_C value="Re|String|0|200|||City">
		<% small_help "City" %></TD>
	</TR>
	<TR bgcolor='#FFFFFF'>
		<td width="10%" class="inputname"><b>State</TD>
		<td width="90%" class="inputvalue"><INPUT  name=State size="60" value="<%= state %>" maxlength=200>
		<INPUT type="hidden"  name=State_C value="Re|String|0|200|||State">
		<% small_help "State" %></TD>
	</TR>

	<TR bgcolor='#FFFFFF'>
		<td width="10%" class="inputname"><b>Country</TD>
		<td width="90%" class="inputvalue">
			<INPUT	name=Country size="60"	value="<%= country %>">
			<% small_help "Country" %></TD>
	</TR>

	<TR bgcolor='#FFFFFF'>
		<td width="10%" class="inputname"><b>Zip Code</TD>
		<td width="90%" class="inputvalue">
			<INPUT	name=Zip size="10" value="<%= Zip %>" maxlength=10>
			<INPUT type="hidden"  name=zip_C value="Re|String|0|14|||Zip Code">
			<% small_help "Zip Code" %></TD>
	</TR>

	<TR bgcolor='#FFFFFF'>
		<td width="10%" class="inputname"><b>Phone</TD>
		<td width="90%" class="inputvalue">
			<INPUT	name=phone size="60" value="<%= Phone %>" maxlength=50 onKeyPress="return goodchars(event,'0123456789()-')">
			<INPUT type="hidden"  name=phone_C value="Re|String|0|50|||Phone">
			<% small_help "Phone" %></TD>
	</TR>


<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Code","req","Please enter a code name for this affiliate.");
 frmvalidator.addValidation("Contact_Name","req","Please enter a contact name.");
 frmvalidator.addValidation("Email","req","Please enter a email address.");
 frmvalidator.addValidation("Email","email","Please enter a email address.");
 frmvalidator.addValidation("Password","req","Please enter a password.");
 frmvalidator.addValidation("Password_Confirm","req","Please enter the password again.");
 frmvalidator.addValidation("Address","req","Please enter a address.");
 frmvalidator.addValidation("City","req","Please enter a city.");
 frmvalidator.addValidation("State","req","Please enter a state.");
 frmvalidator.addValidation("Country","req","Please enter a country.");
 frmvalidator.addValidation("Zip","req","Please enter a zip code.");
 frmvalidator.addValidation("phone","req","Please enter a phone number.");
 frmvalidator.setAddnlValidationFunction("DoCustomValidation");

function DoCustomValidation()
{
  var frm = document.forms[0];
  if (document.forms[0].Password.value == document.forms[0].Password_Confirm.value)
  {
    return true;
  }
  else
  {
    alert('The 2 passwords entered did not match.');
    return false;
  }
}


</script>

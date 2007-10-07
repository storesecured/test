<!--#include virtual="common/connection.asp"-->

<% 

' IF CALLING FROM AFFILIATES' PAGES, USE AFFILIATES_SESSION.ASP
' FOR SESSION MANAGEMENT
RUN_FROM_AFFILIATES = TRUE
if Session("AFFILIATE_SESSION")="TRUE" then
	%><!--#include file="affiliates_session.asp"--><%
else
	Response.Redirect "affiliates_login.asp"
end if

checked1="checked"
checked2=""

aff_id = Session("AFFILIATE_AFFILIATE_ID")

'IF EDIT THE LOAD CURRENT VALUES FROM THE DATABASE
sql_sel = "select * from store_affiliates where store_id="&store_id&" and affiliate_id="&aff_id
rs_Store.open sql_sel,conn_store,1,1
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

%>
	<!--#include file="affiliates_header.asp"-->
<form method="POST" action="Affiliates_Action.asp">
	<input type="hidden" name="ACTION" value="EDIT_AFFILIATE">
	<input type="hidden" name="AFF_ID" value="<%= aff_id %>">




	<tr>
		<td width="100%" colspan="3" height="11">
			<table border="0" width="100%" height="1">
				<tr>
					<td width="1%" height="12" nowrap><font size="2" face="Arial">Login</font></td>

						<td height="12">
							<font size="2" face="Arial"><b><%= Code %></b>
							<INPUT type="hidden" name="Code" value="<%= Code %>">
							</font></td>

				</tr>
    
				<tr>
					<td width="1%" height="12" nowrap><font size="2" face="Arial">Contact Name</font></td>
					<td height="12"><font size="2" face="Arial">
						<input type="text" name="Contact_Name" size="20" value="<%= Contact_Name %>" maxlength="50">
						<INPUT type="hidden" name="Contact_Name_C" value="Re|String||||"></font></td>
			</tr>
	
				<tr>
					<td width="1%" height="12" nowrap><font size="2" face="Arial">Email</font></td>
					<td height="12"><font size="2" face="Arial">
						<input type="text" name="Email" size="20" value="<%= Email %>" maxlength="50">
						<INPUT type="hidden" name="Email_C" value="Re|String||||"></font></td>
				</tr>
	
				<tr>
					<td width="1%" height="12" nowrap><font size="2" face="Arial">URL</font></td>
					<td height="12"><font size="2" face="Arial">

							<input type="text" name="URL" size="20" value="<%= URL %>" maxlength="255">

						<INPUT type="hidden" name="URL_C" value="Re|String||||"></font></td>
				</tr>
	
				<tr>
					<td width="1%" height="12" nowrap><font size="2" face="Arial">Email Notifications</font></td>
					<td height="12"><font size="2" face="Arial">
						<input type="radio" name="Email_Not" value="-1" <%= checked1 %>>Yes&nbsp;
						<input type="radio" name="Email_Not" value="0" <%= checked2 %>>No</font></td>
				</tr>

				<tr>
					<td width="1%" height="12" nowrap><font size="2" face="Arial">Password</font></td>
					<td height="12"><font size="2" face="Arial">
						<input type="password" name="Password" size="20" value="<%= Password %>" maxlength="50">
						<INPUT type="hidden" name="Password_C" value="Re|String||||"></font></td>
				</tr>

				<tr>
					<td width="1%" height="12" nowrap><font size="2" face="Arial">Confirm Password</font></td>
					<td height="12"><font size="2" face="Arial">
						<input type="password" name="Password_Confirm" size="20" value="<%= Password %>" maxlength="50"></font></td>
			</tr>
			 <TR>
		<TD class='normal'><font size="2" face="Arial">Company</TD>
		<TD colspan="2" class='normal'><INPUT	name=company size="20" value="<%= Company %>">
		<INPUT type="hidden"  name=Company_C value="Op|String|0|100|||Company"></TD>
	</TR>

	<TR>
		<TD class='normal'><font size="2" face="Arial">Address</TD>
		<TD colspan="2" class='normal'><INPUT	name=Address size="20" value="<%= address %>">
		<INPUT type="hidden" name=Address_C value="Re|String|0|200|||Address"></TD>
	</TR>

	<TR>
		<TD class='normal'><font size="2" face="Arial">City</TD>
		<TD colspan="2" class='normal'><INPUT	name=City size="20" value="<%= city %>">
		<INPUT type="hidden"  name=City_C value="Re|String|0|200|||City"></TD>
	</TR>
	<TR>
		<TD class='normal'><font size="2" face="Arial">State</TD>
		<TD colspan="2" class='normal'><INPUT	name=State size="20" value="<%= state %>">
		<INPUT type="hidden"  name=State_C value="Re|String|0|200|||State"></TD>
	</TR>

	<TR>
		<TD class='normal'><font size="2" face="Arial">Country</TD>
		<TD colspan="2" class='normal'>
			<INPUT	name=Country size="20"  value="<%= country %>"></TD>
	</TR>

	<TR>
		<TD class='normal'><font size="2" face="Arial">Zip Code</TD>
		<TD colspan="2" class='normal'>
			<INPUT	name=Zip size="10" value="<%= Zip %>">
			<INPUT type="hidden"  name=zip_C value="Re|String|0|14|||Zip Code"></TD>
	</TR>
	
	<TR>
		<TD class='normal'><font size="2" face="Arial">Phone</TD>
		<TD colspan="2" class='normal'>
			<INPUT	name=phone size="20" value="<%= Phone %>">
			<INPUT type="hidden"  name=phone_C value="Re|String|0|50|||Phone"></TD>
	</TR>

			 <tr>
					<td colspan=2 height="12" nowrap><input type="submit" value="Save" name="Edit_Affiliate"></td>
			</tr>
			</table>
		</td>
	</tr>
</form>

<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<% 

op=Request.QueryString("op")
if op="edit" then
	newsletter_id = Request.QueryString("Id")
	if newsletter_id = "" then
	newsletter_id = Request.QueryString("newsletter_id")
	end if
	if not isNumeric(newsletter_id) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	' IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlSelNewsletter="select * from store_newsletter where newsletter_id=" & newsletter_id & " and store_id=" & store_id
	rs_Store.open sqlSelNewsletter,conn_store,1,1
	if rs_Store.bof or rs_Store.eof then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	Email_Address=rs_store("Email_Address")
     First_Name=rs_store("First_Name")
     Last_Name=rs_store("Last_Name")
	rs_Store.close
end if

sFullTitle ="Marketing > Newsletter > <a href=newsletter_manager.asp class=white>Subscribers</a> > "
if op<>"edit" then
      sTitle ="Add Newsletter Subscriber"
      sFullTitle = sFullTitle&"Add"
else
      sTitle = "Edit Newsletter Subscriber - "&Email_Address
      sFullTitle = sFullTitle&"Edit - "&Email_Address
end if
sCommonName = "Newsletter Subscriber"
sCancel="newsletter_manager.asp"
sFormAction = "Newsletter_action.asp"
sSubmitName = "Store_Activation_Update"
thisRedirect = "new_newsletter.asp"
sFormName = "Create_Newsletter"
sMenu = "marketing"
sQuestion_Path = "marketing/newsletter.htm"
createHead thisRedirect
%>

<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Newsletter_Id" value="<%=newsletter_id%>">
	



			
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Email</b></td>
				<td width="60%" class="inputvalue"><input type="text" name="Email_Address" size="60" value="<%= Email_Address %>">
					<INPUT type="hidden"  name=Email_Address_C value="Re|Email|0|255|||Email Address">
			 <% small_help "Email" %></td>

			</tr>
               <tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>First Name</b></td>
				<td width="60%" class="inputvalue"><input type="text" name="First_Name" size="60" value="<%= First_Name %>">
					<INPUT type="hidden"  name=First_Name_C value="Re|String|0|50|||First Name">
			 <% small_help "First Name" %></td>

			</tr>
              <tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Last Name</b></td>
				<td width="60%" class="inputvalue"><input type="text" name="Last_Name" size="60" value="<%= Last_Name %>">
					<INPUT type="hidden"  name=Last_Name_C value="Re|String|0|50|||Last Name">
			 <% small_help "Last Name" %></td>

			</tr>

<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Email_Address","email","Please enter a email address.");
 frmvalidator.addValidation("Email_Address","req","Please enter a email address.");
 frmvalidator.addValidation("First_Name","req","Please enter a first name.");
 frmvalidator.addValidation("Last_Name","req","Please enter a last name.");

</script>

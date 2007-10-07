<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%
on error goto 0
if request.querystring("Email") <> "" then
	'IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	Email = lcase(request.querystring("Email"))
	Store_Domain = lcase(replace(Store_Domain,"www.",""))
	if not isnull(Store_Domain2) then
   Store_Domain2 = lcase(replace(store_domain2,"www.",""))
   end if
   if (Store_Domain <> "" and instr(Email,"@"&Store_Domain)) > 0 or (Store_Domain2 <> "" and instr(Email,"@"&Store_Domain2)>0) then
    set Users=server.CreateObject("MailServerX.Users")

    i=Users.IndexOf(Email)
    if i > -1 then
       set User=Users.Items(i)
       Email=User.UserName
       Password=User.Password
       ForwardAddress=User.ForwardAddress
       KeepCopies=User.KeepCopies
       if KeepCopies then
          Keep_checked="checked"
       else
          Keep_checked=""
       end if
       set User=Nothing
    end if
    set Users=Nothing
  else
    Email=""
  end if
end if

If Email <> "" then
	 sTitle = "Edit Email"
	 sFormName = "Email_Edit"
Else
	 sTitle = "Add Email"
	 sFormName = "Email_Add"
End If

sFormAction = "email_action.asp"
sName = "Add_Store_Log_ins"
thisRedirect = "email_add.asp"
sSubmitName = "Add_Store_Log_ins"
sMenu = "general"
createHead thisRedirect

Email_Domain=lcase(replace(Store_Domain,"www.",""))
'Email_Domain="easystorecreator.com"

if Email_Domain <> "" then %>






				 <tr>
					<td width="100%" colspan="2" height="15">
						<input type="button" class="Buttons" OnClick=JavaScript:self.location="newemail.asp" value="Back to Emails" name="Back_To_Store_Emails"></td>
				</tr>
		
        <tr>
				<td width="24%" class="inputname"><b>Email</b></td>
				<td width="76%" class="inputvalue">
		    <% If Email <> "" then %>
        	<input type="Hidden" name="Email" value="<%= Email %>">
        	<%= Email %>

        <% else %>

						<input onKeyPress="return goodchars(event,'0123456789_abcdefghijklmnopqrstuvwxyz-')" type="text" name="Email" size="30" value="<%= Email %>">@<%= replace(Store_Domain,"www.","") %>
        <% End If %>
        <INPUT type="hidden"  name=Email_C value="Re|String|0|50|||Email">
					<% small_help "Email" %></td></tr>

		
				<tr>
				<td width="24%" class="inputname"><b>Password</b></td>
				<td width="76%" class="inputvalue">
						<input type="password" name="Password" size="30" value="<%= Password %>">
						<INPUT type="hidden"  name=Password_C value="Re|String|0|50|||Password">
					<% small_help "Password" %></td>
				</tr>
		
				<tr>
				<td width="24%" class="inputname"><b>Confirm Password</b></td>
				<td width="76%" class="inputvalue">
						<input type="password" name="Password_Confirm" size="30" value="<%= Password %>">
						<INPUT type="hidden"  name=Password_Confirm_C value="Re|String|0|50|||Confirm Password">
					<% small_help "Confirm Password" %></td>
				</tr>
				
				<tr>
				<td width="24%" class="inputname"><b>Forward Address</b></td>
				<td width="76%" class="inputvalue">
						<input type="text" name="ForwardAddress" size="50" value="<%= ForwardAddress %>">
						<INPUT type="hidden"  name=ForwardAddress_C value="Op|String|0|50|||Forward Address">
					<% small_help "Forward" %></td>
				</tr>
				
				<tr>
				<td width="24%" class="inputname"><b>Keep Copies</b></td>
				<td width="76%" class="inputvalue">
						<input class=image type="checkbox" name="KeepCopies" <%=Keep_checked %>>
					<% small_help "Keep Copies" %></td>
				</tr>




<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
frmvalidator.addValidation("Email","req","Please enter your email.");
frmvalidator.addValidation("Password","req","Please enter your password.");
frmvalidator.addValidation("Password_Confirm","req","Please enter your password again.");
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
    alert('The 2 passwords must match.');
    return false;
  }
}

</script>

<% else
    response.write "You cannot setup email addresses until you <a href=domain.asp class=link>choose a domain name</a>."
    createFoot thisRedirect, 0
end if

%>

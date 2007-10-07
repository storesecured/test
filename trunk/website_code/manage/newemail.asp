<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

if request.querystring("Delete") then
   Delete_Email = request.querystring("Delete")
   Store_Domain = lcase(replace(Store_Domain,"www.",""))
   Store_Domain2 = lcase(replace(Store_Domain2,"www.",""))

   if (Store_Domain <> "" and instr(Delete_Email,"@"&Store_Domain)) > 0 or (Store_Domain2 <> "" and instr(Delete_Email,"@"&Store_Domain2)) then
    set Users=server.CreateObject("MailServerX.Users")
    Users.Delete(Delete_Email)
    set Users=Nothing
  end if
end if

sFormAction = "newemail.asp"
sName = "Email"
sFormName = "email"
sTitle = "Email Accounts"
sSubmitName = "Email"
thisRedirect = "newemail.asp"
sTopic="Email Accounts"
sMenu = "general"
createHead thisRedirect
if Service_Type=0 then %>
	<tr>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		PEARL Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>

	</td></tr>

<% else
Email_Domain=lcase(replace(Store_Domain,"www.",""))
'Email_Domain="easystorecreator.com"

if Email_Domain <> "" then %>

<tr>
		<td width="100%" colspan="3" height="21">
			<input type="button" class="Buttons" OnClick=JavaScript:self.location="email_add.asp" value="Add Email" name="Add"></td>
	 </tr>
<%

set Users=server.CreateObject("MailServerX.Users")
set thisCount=0

for i=0 to Users.Count-1
    if instr(Users.Items(i).UserName,"@"&Email_Domain) > 0 then
       sEmail = Users.Items(i).UserName
       sForward = Users.Items(i).ForwardAddress
       sKeep = Users.Items(i).KeepCopies
       sLastLogin = Users.Items(i).LastLoginTime
       if thisCount=0 then %>
          <tr>
      		<td width="100%" colspan="3" height="74">
      			<table width="100%" cellspacing="0" cellpadding=2 class="list">

      				<tr bgcolor=#DDDDDD>
      					<td width="20%" align="center"><b>Email</b></td>
      					<td width="20%" align="center"><b>Forward</b></td>
      					<td width="20%" align="center"><b>Keep Copies</b></td>
      					<td width="20%" align="center"><b>Last Login</b></td>
      					<td width="20%">&nbsp;</td>
      					<td width="20%">&nbsp;</td>
      				</tr>
       <% end if %>
       <tr><td><%= sEmail %></td><td><%= sForward %></td><td><%=sKeep %></td><td><%= sLastLogin %></td><td><a class="link" href="email_add.asp?Email=<%=sEmail %>">Edit</a></td><td><a class="link" href=newemail.asp?Delete=<%=sEmail %>>Delete</a></td></tr>
       <% thisCount=thisCount+1
    end if
next
if thisCount=0 then
   response.write "<tr><td>There are no email addresses for your store.</td></tr>"
else
   response.write "</table></td></tr>"
end if
Set Users=Nothing
else
    response.write "You cannot setup email addresses until you choose a domain name.  <a href=domain.asp class=link>Click here to choose a domain name.</a>"
end if

end if 

createFoot thisRedirect, 0 %>

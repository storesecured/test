<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->
<!--#include file="help/message_customers.asp"-->

<%

Server.ScriptTimeout =500

if Session("Super_User") <> 1 then
     Response.redirect "noaccess.asp"
end if

if request.form("Message_Customers") <> "" then
   set myEmails = server.createobject("scripting.dictionary")
if request.form("From")<>"" then
From = request.form("From")
else
From = sNoReply_email
end if
	
	StartBody = request.form("Body")
	StartSubject = request.form("Subject")

if request.form("To")=1 then
ToEmail = lcase(request.form("ToMerchant"))
			Body = StartBody
			Subject = StartSubject

               Send_Mail_Html From,ToEmail,Subject,Body
else
'From = sNoReply_email


		sql_customers = "select distinct store_email from Store_settings WITH (NOLOCK) where Newsletter_Receive<>0"

		set myfields=server.createobject("scripting.dictionary")
		Call DataGetrows(conn_store,sql_customers,mydata,myfields,noRecords)
		if noRecords = 0 then
		 FOR rowcounter= 0 TO myfields("rowcount")
			ToEmail = lcase(mydata(myfields("email_address"),rowcounter))
			Body = StartBody
			Subject = StartSubject

               Send_Mail_Html From,ToEmail,Subject,Body
		Next
		end if
end if
set myEmails=Nothing
end if
on error resume next
sTextHelp="newsletter/newsletter_send.doc"

sFormAction = "message_merchants.asp"
sName = "Message_Customers"
sFormName = "customers"
sSubmitName = "Message_Customers"
thisRedirect = "message_merchants.asp"
sTopic="Message_Customers"
sMenu="marketing"
addPicker=1
sTitle = "Send Newsletter"
sFullTitle = "Marketing > <a href=newsletter_manager.asp class=white>Newsletter</a> > Send"
sQuestion_Path = "marketing/mail_merge.htm"
createHead thisRedirect
if Service_Type < 1	then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		PEARL Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>

<% else %>

		<tr bgcolor='#FFFFFF'><td class="inputname"><b>From</b></td><td class="inputvalue">
      <input type=text name=From value='<%= sNoReply_email %>' size=38>
                		<% small_help "From" %></td></tr>
                <tr bgcolor='#FFFFFF'><td class="inputname"><b>To</b></td><td class="inputvalue">
                 <input class="image" type="radio" name="To" value="0" checked onClick="this.form.ToMerchant.disabled = this.checked"><b>All Merchants&nbsp;&nbsp;<input class="image" type="radio" name="To" value="1" onClick="this.form.ToMerchant.disabled = !this.checked">This Email</b><br>
                	 <input type=text name=ToMerchant value='' size=38 disabled=true>
                        	<% small_help "To" %></td></tr>
		<tr bgcolor='#FFFFFF'><td class="inputname"><b>Subject</b></td><td class="inputvalue"><input type=text name=Subject value='' size=38>
		<% small_help "Subject" %></td></tr>
		<tr bgcolor='#FFFFFF'><td class="inputname" colspan=2><B>Body</b><BR>
		<% on error goto 0 %>
		<%= create_editor ("Body","","[""First Name"",""%FIRSTNAME%""],[""Last Name"",""%LASTNAME%""],[""Login"",""%LOGIN%""],[""Password"",""%PASSWORD%""],[""City"",""%CITY%""],[""Last Visit"",""%LASTVISIT%""],[""Budget Left"",""%BUDGETLEFT%""],[""Reward Left"",""%REWARDLEFT%""],[""Order Total"",""%ORDERSTOTAL%""],[""First Visit"",""%FIRSTVISIT%""]") %>
		<% small_help "Body" %></td></tr>
		<tr bgcolor='#FFFFFF'><td colspan=3 align=center><input class=buttons type=submit name=Message_Customers value="Send Newsletter">
                <input type="button" OnClick=JavaScript:self.location="newsletter_manager.asp" class="Buttons" value="Cancel" name="Create_new_Page_link">
			</td></tr>
		<tr bgcolor='#FFFFFF'><td colspan=3 class=instructions>Please note that depending on the size of your customer list this may take a while, please only hit send once.</td></tr>
<% end if %>

<% createFoot thisRedirect, 0


%>



<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
response.redirect "support_request.asp"

sTitle = "Support Request"
thisRedirect = "support.asp"
sMenu="none"
sFormAction = "support.asp"

createHead thisRedirect



if request.form("Update") <> "" then
	sMessage = "Support Request from Store Id: "&Store_Id&vbcrlf&request.form("fromname")&vbcrlf&request.form("phone")&vbcrlf&Service_Type&vbcrlf&"==================================="&vbcrlf&vbcrlf&request.form("message")
	Send_Mail request.form("fromemail"),"support@easystorecreator.com",request.form("subject"),sMessage
	%>
	<tr><td>Your support request has been sent.  Average response time is 30-60 minutes during regular business hours 8:00 AM - 8:00 PM PST.</td></tr>
<% else %>
   <TR><TD colspan=2>Before requesting support you might want to check the <a href=help.asp class=link>Help Section</a> to see if you question is already answered.</td></tr>
   <tr><td colspan=2><BR>Please fill out the form below for prompt, efficient support or email support@easystorecreator.com.<BR><BR>
	When requesting support please note your store id
	of <%= Store_Id %> which will help us quickly locate your account.  Please ensure that all contact information
	provided is correct.<BR></td></tr>
   <tr><td class=inputname>Contact Email</td>
	<td class=inputvalue><input type=text name=fromemail size=30 value='<%= Store_Email %>'><% small_help "Email" %></td></tr>
	<tr><td class=inputname>Contact Name</td>
	<td class=inputvalue><input type=text name=fromname size=30 value=''><% small_help "Name" %></td></tr>
	<tr><td class=inputname>Contact Phone</td>
	<td class=inputvalue><input type=text name=phone size=30 value='<%= Store_Phone %>'><% small_help "Phone" %></td></tr>
	<tr><td class=inputname>Subject</td>
	<td class=inputvalue><input type=text name=subject size=30 value='Question Store: <%= Store_Id %>'><% small_help "Subject" %></td></tr>

	<tr><td class=inputname>Question or Request Details</td><td class=inputvalue><textarea rows="15" name="message" cols="55"></textarea><% small_help "Details" %></td></tr>
	<tr><td colspan=2 align=center><input type="submit" name=Update value="Send"></td></tr>

<% end if %>

<% createFoot thisRedirect, 0 %>



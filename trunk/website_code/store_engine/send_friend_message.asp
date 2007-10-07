<!--#include file="include/header.asp"-->
<table>
<%

if request.form("With_message") <> "" then
	Email_Address = request.form("Email_Address")
	Send_To = request.form("Send_To")
	With_Message = request.form("With_Message")
	Subject = request.form("Subject")

	Call Send_Mail(Email_Address,Send_To,Subject,With_Message)
  Call Send_Mail(Email_Address,Store_Email,"Send Friend",Send_To)
	
	sMessage = "Your message has been successfully sent.<BR><BR>"
%>
<tr><td colspan=2><%= sMessage %></td></tr>
<%
else

%>
	<form method=post action=<%= Site_Name %>send_friend_message.asp><input type=hidden name=Item_Id value='<%= Item_Id %>'>

	<tr><td>Your Email</td><td><input type=text size=35 name=Email_Address></td></tr>
	<tr><td nowrap>Friends Email</td><td><input type=text size=35 name=Send_To></td></tr>
	<tr><td>Friends Email</td><td><input type=text size=35 name=Send_To></td></tr>
	<tr><td>Friends Email</td><td><input type=text size=35 name=Send_To></td></tr>
	<tr><td>Friends Email</td><td><input type=text size=35 name=Send_To></td></tr>
	<tr><td>Friends Email</td><td><input type=text size=35 name=Send_To></td></tr>
	<tr><td>Friends Email</td><td><input type=text size=35 name=Send_To></td></tr>
	<tr><td>Friends Email</td><td><input type=text size=35 name=Send_To></td></tr>
	<tr><td>Friends Email</td><td><input type=text size=35 name=Send_To></td></tr>
	<tr><td>Friends Email</td><td><input type=text size=35 name=Send_To></td></tr>
	<tr><td>Friends Email</td><td><input type=text size=35 name=Send_To></td></tr>
	<tr><td>Subject</td><td><input type=text size=35 name=Subject></td></tr>
	<tr><td>Message</td><td><textarea rows=10 cols=40 name=With_message>Type your message here</textarea></td></tr>
	<tr><td colspan=2 align=center><%= fn_create_action_button ("Button_image_Continue", "Continue", "Send Now") %></td></tr>
<% end if %>
</table>
<!--#include file="include/footer.asp"-->

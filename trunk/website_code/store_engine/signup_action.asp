<!--#include file="include/header.asp"-->
<!--#include file="include/emails.asp"-->
<% 
If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	fn_redirect Switch_Name&"form_error.asp?Error_Log="&server.urlencode(Error_Log)
else
'SEND EMAIL TO THE CUSTOMER
    Email_Address = checkStringForQ(Request.Form("Email_Address"))
	if Email_Address="" then
		 fn_error "You did not enter a valid email address."
	end if
	First_Name = checkStringForQ(Request.Form("First_Name"))
	Last_Name = checkStringForQ(Request.Form("Last_Name"))
		
	on error resume next
    sql_execute = "exec wsp_newsletter_add "&Store_id&",'"&Email_Address&"','"&First_name&"','"&Last_Name&"';"
	conn_store.execute sql_execute
	if err.number<>0 then
	    fn_error err.description
	end if
	on error goto 0

	sURL = Site_Name&"cancel_signup.asp?Email_Address="&Server.urlEncode(Email_Address)
	email_message = "You are now subscribed to our newsletter.	To be removed from the newsletter at anytime go to "&sURL
	email_message = Newsletter_notification & vbcrlf & vbcrlf & email_message
	Call Send_Mail_Html(Store_Email,Email_Address,Newsletter_notification_subject,email_message)
	
%>
		<table border="0" width="100%" cellspacing="0" cellpadding="0">	
			<tr>
			<td width="100%" class='normal'>
				<p>You have been successfully subscribed to our newsletter.</p>
				<p>An email message with a link to unsubscribe has been sent to <%= Email_Address %> </p>
				</td>
			</tr>

		</table>
		
<%	
	
    if request.form("Redirect") <> "" then
       fn_redirect request.form("Redirect")
    end if
%>

<% end if %>
<!--#include file="include/footer.asp"-->


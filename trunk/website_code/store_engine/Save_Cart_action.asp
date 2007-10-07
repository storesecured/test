<!--#include file="include/header.asp"-->

<% 

'SEND EMAIL TO THE CUSTOMER
If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	fn_redirect "form_error.asp?Error_Log="&server.urlencode(Error_Log)
else
   sURL = Switch_Name&"include/form_action.asp?Shopper_Id_Retrieve="&Shopper_ID
   Send_to_email = Request.Form("Send_to_email")
   Email_Subject = Store_name&" wish list"
   email_message = "Your cart has been saved as cart #"&Shopper_ID&chr(10)&"To retrieve your cart please use the following URL: "&sURL&chr(10)&"Sincerely, "&Chr(10)&Store_Name
   cc_to_email = Request.Form("cc_to_email")
   message_cart = Request.Form("message_cart")&chr(10)&email_message
   Call Send_Mail(Store_Email,Send_to_email,Email_Subject,message_cart)
   'SEND EMAIL TO A FRIEND SPECIFIED BY THE CUSTOMER
   if cc_to_email<>"" then
      Call Send_Mail(Send_to_email,cc_to_email,Email_Subject,message_cart)
   end if
%>
  
<table border="0" width="100%" cellspacing="0" cellpadding="0">	
	<tr>
		<td width="100%" class='normal'><b>Your cart was saved</b>
			<p>&nbsp;Cart ID <b> <%= Shopper_Id %></b></p>
			<p>An email message with a link to your cart was sent to: <%= Send_to_email %> &nbsp;<%= cc_to_email %></p>
			<p><%= sURL %></td>
	</tr>
</table>
<% end if %>
<!--#include file="include/footer.asp"-->

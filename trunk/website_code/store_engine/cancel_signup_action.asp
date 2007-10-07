<!--#include file="include/header.asp"-->

<% 

'SEND EMAIL TO THE CUSTOMER

Email_Address = Request.Form("Email_Address")
email_message = "Your email has been successfully removed from our newsletter list."

sql_delete = "exec wsp_newsletter_remove "&Store_Id&",'"&Email_Address&"';"
fn_print_debug sql_delete
conn_store.Execute sql_delete

%>

<table border="0" width="100%" cellspacing="0" cellpadding="0">	
	<tr>
		<td width="100%" class='normal'>
			<p>You have successfully unsubscribed from our newsletter.</p>
			</td>
	</tr>
</table>

<!--#include file="include/footer.asp"-->

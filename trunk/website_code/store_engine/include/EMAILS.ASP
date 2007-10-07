
<% 

sql_store_emails = "select * from store_emails WITH (NOLOCK) where Store_id = "&Store_id 
fn_print_debug sql_store_emails
rs_Store.open sql_store_emails,conn_store,1,1
if rs_Store.Bof = False then
	order_notification_enable= Rs_store("order_notification_enable")
	order_notification_to_admin_enable= Rs_store("order_notification_to_admin_enable")
	order_notification_subject = Rs_store("order_notification_subject")
	order_notification_body= Rs_store("order_notification_body")
	order_notification_invoice= Rs_store("order_notification_invoice")
	order_notification_sent_to= Rs_store("order_notification_sent_to")
	Registration_enable=Rs_store("Registration_enable")
	Registration_user_psw=Rs_store("Registration_user_psw")
	Registration_subject=Rs_store("Registration_subject")
	Registration_body=Rs_store("Registration_body")
	Registration_sent_to=Rs_store("Registration_sent_to")
	Newsletter_notification_subject=Rs_store("Newsletter_notification_subject")
	Newsletter_notification=Rs_store("Newsletter_notification")
	Gift_buyer_email_body=Rs_store("Gift_buyer_email_body")
	Gift_email_Subject=Rs_store("Gift_email_Subject")
	Pin_email_subject=Rs_store("Pin_email_subject")
	pin_buyer_email_body=Rs_store("pin_buyer_email_body")
end if
Rs_store.Close

%>

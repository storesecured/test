<!--#include virtual="store_engine/include/store_settings.asp"-->
<!--#include virtual="store_engine/include/emails.asp"-->
<!--#include virtual="store_engine/include/Send_Invoice_By_Mail.asp"-->

<%

oid=fn_get_querystring("oid")

if isNumeric(oid) then
if order_notification_to_admin_enable <> 0 or order_notification_enable = -1 then
  'CHECK IF TO SEND ORDER NOTIFICATIONS EMAILS TO MERCHANT, CUSTOMER
  sql_select_orders = "select distinct oid,ccid,Shipping_Method_Price,Shipping_Method_Name, Tax, Coupon_id, Coupon_Amount,Grand_Total, Recurring_Total, Recurring_Days, Recurring_Fee, Recurring_Tax, Purchase_Date, Invoice_Id, shipto,payment_method,shipemail,shipfirstname,shiplastname	from store_purchases WITH (NOLOCK) where (oid="&oid&" or masteroid="&oid&") and store_id="&store_id
  
  set mailfields=server.createobject("scripting.dictionary")
  Call DataGetrows(conn_store,sql_select_orders,maildata,mailfields,noRecordsmail)

  if noRecordsmail = 0 then
  FOR rowcountermail= 0 TO mailfields("rowcount")
  	loid = maildata(mailfields("oid"),rowcountermail)
  	ccid = maildata(mailfields("ccid"),rowcountermail)
  	session("ccid")=ccid
  	Shipping_Method_Price = maildata(mailfields("shipping_method_price"),rowcountermail)
  	Shipping_Method_Name = maildata(mailfields("shipping_method_name"),rowcountermail)
  	Tax = maildata(mailfields("tax"),rowcountermail)
  	Coupon_id = maildata(mailfields("coupon_id"),rowcountermail)
  	Coupon_Amount = maildata(mailfields("coupon_amount"),rowcountermail)
  	Grand_Total = maildata(mailfields("grand_total"),rowcountermail)
  	Payment_Method = maildata(mailfields("payment_method"),rowcountermail)
  	Recurring_Total = maildata(mailfields("recurring_total"),rowcountermail)
  	Recurring_Days = maildata(mailfields("recurring_days"),rowcountermail)
  	Recurring_FeeT = maildata(mailfields("recurring_fee"),rowcountermail)
  	Recurring_Tax = maildata(mailfields("recurring_tax"),rowcountermail)
  	Purchase_Date = maildata(mailfields("purchase_date"),rowcountermail)
  	Invoice_Id = maildata(mailfields("invoice_id"),rowcountermail)
  	fship_to = maildata(mailfields("shipto"),rowcountermail)
   email = maildata(mailfields("shipemail"),rowcountermail)
   first_name = maildata(mailfields("shipfirstname"),rowcountermail)
   last_name = maildata(mailfields("shiplastname"),rowcountermail)

  	if order_notification_to_admin_enable <> 0 then
  		'SEND INDEPENDENT EMAIL TO STORE ADMIN
  		Subject = "Easystorecreator.com new order notification: "&loid&" for "&store_name
  		Body = "Store Owner"&chr(10)&"You have received a new order #"&loid&" at "&FormatDateTime(Now(),2)&chr(10)&" To process and view the details of this order please click here http://manage10.easystorecreator.com/order_details.asp?oid="&loid&"&cid="&cid&chr(10)
  		Call Send_Mail(Store_Email,Store_Email,Subject,Body)
  		send_to = email&","&store_email&","&Replace(order_notification_sent_to, " ", "")
      response.write "<BR>admin email notification sent to "&send_to
          else
  		send_to = email&","&order_notification_sent_to

  	end if
  	if order_notification_enable = -1 then
  		'IF SET TO SEND MAIL TO STORE ADMIN ALSO

                order_notification_body_replaced=order_notification_body
  		if InStr(order_notification_body,"%") > -1 then

  				order_notification_body_replaced = Replace(order_notification_body_replaced,"%LASTNAME%",Last_name)
  				order_notification_body_replaced = Replace(order_notification_body_replaced,"%FIRSTNAME%",First_name)
  				order_notification_body_replaced = Replace(order_notification_body_replaced,"%LOGIN%",User_Id)
  				order_notification_body_replaced = Replace(order_notification_body_replaced,"%PASSWORD%",Password)
                                order_notification_body_replaced = Replace(order_notification_body_replaced,"%INVOICE_TEXT%",Send_Text_Invoice_By_Mail(loid)	)

  			end if
  		if order_notification_invoice = -1 then
  			'FORMAT THE INVOICE TO BE SENT TO THE CUSTOMER
  			add_invoice = add_invoice_T1
  			Call Send_Invoice_By_Mail(loid,Store_Email,send_to,order_notification_subject,order_notification_body_replaced&chr(10)&add_invoice)
  		else
  			add_invoice = add_invoice_T2
  			Call Send_Mail_Html(Store_Email,send_to,order_notification_subject,order_notification_body_replaced&chr(10)&add_invoice)
  		end if
  		response.write "<BR>general email notification sent to "&send_to
  	end if
  Next
  End if
end if
response.write "<BR>Notification(s) sent for order "&oid
response.write "<BR>Close this window to exit"
end if
%>


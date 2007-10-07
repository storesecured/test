<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include file="include\google\googleglobal.asp"-->
<% 

'RETRIEVE FORM DATA

if request.form = "" then
	response.redirect "admin_home.asp"
end if
Tracking_ID = Replace(Request.Form("Tracking_ID"), "'", "''")
ShippedPr = Request.Form("ShippedPr") 
Shippment_notes = Replace(Request.Form("Shippment_notes"), "'", "''")
Shipping_Company = request.form("Shipping_Company") 
Shipped_Date = Now()
Cid = Request.Form("Cid")
Oid = Request.Form("Oid")

sql_select_type = "Select Transaction_Type,Processor_id from Store_Purchases where oid = "&Oid&" and Store_id ="&Store_id
rs_Store.open sql_select_type,conn_store,1,1
     Transaction_Type = Rs_store("Transaction_Type")
     Processor_id = Rs_store("Processor_id")
rs_Store.close

if cid="" then
   cid=0
end if

'INSERT SHIPPING DETAILS
Real_Time_Processor = Processor_id
if Real_Time_Processor="" then
	Real_Time_Processor=0  
end if
if Request.Form("capture") = 1 then
     Real_Time_Processor = Request.form("Processor_id")
     GGrand_Total = Request.Form("new_amount")
     sType="Capture"
	if Real_Time_Processor = 1 then
		%><!--#include file="include\PlugnPay\pnp.asp"--><%
	elseif Real_Time_Processor = 2 then
		%><!--#include file="include\Authorizenet\Authorizenet.asp"--><%
	elseif Real_Time_Processor = 7 then
	  %><!--#include file="include\Echo\Echo.asp"--><%
       elseif Real_Time_Processor = 9 then
	    %><!--#include file="include\Verisign\verisign.asp"--><%
       end if
	sql_update_purchases = "update store_purchases set Transaction_Type='2' where (oid = "&oid&" or masteroid="&oid&") AND Store_id ="&Store_id
  conn_store.Execute sql_update_purchases
end if
'response.write(Real_Time_Processor)
'response.end
 if Real_Time_Processor = 38 then
		%><!--#include file="include\google\google.asp"--><%
		end if

if Real_Time_Processor = 38 then

Carrier=Shipping_Company
            TrackingNumber= Tracking_ID

	    MyOrder.DeliverOrder GoogleOrderNumber, Carrier, TrackingNumber, True
	    MyOrder.SendOrderCommand

 end if

sql_insert = "Insert into Store_Purchases_Shippments (Store_id,oid,ShippedPr,Tracking_ID,Shippment_notes,Shipped_Date,Shipping_Company) values ("&Store_id&","&oid&",'"&ShippedPr&"','"&Tracking_ID&"','"&Shippment_notes&"','"&Shipped_Date&"','"&Shipping_Company&"')"
session("sql") = sql_insert
Conn_store.execute sql_insert

'UPDATE STORE_PURCHASES TABLE WITH THE GENERAL STATUS OF SHIPPING
sql_update = "update store_purchases set ShippedPr = '"&ShippedPr&"' where Oid = "&Oid&" and Store_id= "&Store_id
Conn_store.execute sql_update
if Request.Form("send_shipment_email") = 1 then
	' send_email to customer regards to that shipping :
	Shipping_Email_To = Request.Form("Shipping_Email_To")
	Shipping_Email_From = Request.Form("Shipping_Email_From")
	Shipping_Email_Subject = Request.Form("Shipping_Email_Subject")
	Shipping_Email_First = Request.Form("first_name")
	Shipping_Email_Last = Request.Form("last_name")
	Shipping_Email_Username = Request.Form("username")
     Shipping_Email_Password = Request.Form("password")
     html_notifications=Request.Form("html_notifications")

	Shipping_Email_Body = Request.Form("Shipping_Email_Body")
	Shipping_Email_Body = replace(Shipping_Email_Body,"%ORDER%",oid)
	Shipping_Email_Body = replace(Shipping_Email_Body,"%TRACKING%",Tracking_ID)
	Shipping_Email_Body = replace(Shipping_Email_Body,"%SHIPCOMPANY%",Shipping_Company)

	Shipping_Email_Body = Replace(Shipping_email_body,"%FIRSTNAME%",Shipping_Email_First)
     Shipping_email_body = Replace(Shipping_email_body,"%LASTNAME%",Shipping_Email_Last)
     Shipping_email_body = Replace(Shipping_email_body,"%LOGIN%",Shipping_Email_Username)
     Shipping_email_body = Replace(Shipping_email_body,"%PASSWORD%",Shipping_Email_Password)


	Shipping_Email_Subject = replace(Shipping_Email_Subject,"%ORDER%",oid)
	Shipping_Email_Subject = replace(Shipping_Email_Subject,"%TRACKING%",Tracking_ID)
	Shipping_Email_Subject = replace(Shipping_Email_Subject,"%SHIPCOMPANY%",Shipping_Company)

     Shipping_Email_Subject = Replace(Shipping_Email_Subject,"%FIRSTNAME%",Shipping_Email_First)
     Shipping_Email_Subject = Replace(Shipping_Email_Subject,"%LASTNAME%",Shipping_Email_Last)
     Shipping_Email_Subject = Replace(Shipping_Email_Subject,"%LOGIN%",Shipping_Email_Username)
     Shipping_Email_Subject = Replace(Shipping_Email_Subject,"%PASSWORD%",Shipping_Email_Password)

     if html_notifications=0 then
        	Call Send_Mail(Shipping_Email_From,Shipping_Email_To,Shipping_Email_Subject,Shipping_Email_Body)
     else
	       Call Send_Mail_Html(Shipping_Email_From,Shipping_Email_To,Shipping_Email_Subject,Shipping_Email_Body)
     end if
end if


Response.Redirect "order_details.asp?oid="&oid&"&cid="&cid

%>

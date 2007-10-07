<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include file="include/gateways_functions.asp"-->
<%

server.scripttimeout=4800

on error resume next
'RETRIEVE FORM DATA
'ShippedPr = Request.Form("ShippedPr")
Shippment_notes = Replace(Request.Form("Shippment_notes"), "'", "''")
Shipped_Date = Now()
Orders = Request.Form("Orders")


'INSERT SHIPPING DETAILS
sArray = Split(Orders,",")
for each Oid in sArray
Oid=Trim(Oid)

shippingcompany=request.form("Shipping_Company")

myTracking_ID=request.form("Tracking_ID" & Oid)
myTracking_ID = Replace(myTracking_ID, "'", "''")
mycapture=request.form("capture" & Oid)

'response.write(myTracking_ID & " " & mycapture & "<br>")


        sql_customers = "select Shiplastname,shipfirstname,user_id,password,Grand_Total, Processor_Id, ShipEmail,store_purchases.cid from store_purchases inner join store_customers on store_purchases.store_id=store_customers.store_id and store_purchases.cid=store_customers.cid where oid = "&Oid&" and store_customers.record_type=0 and store_purchases.Store_Id="&Store_Id
  rs_Store.open sql_customers,conn_store,1,1
	if not rs_Store.bof and not rs_Store.eof then
      Real_Time_Processor = rs_Store("Processor_Id")
      'response.write(Real_Time_Processor)
  	  cid = rs_Store("cid")
  	  email = Rs_store("ShipEmail")
  	  GGrand_Total = rs_Store("Grand_Total")
  	  last_name = rs_Store("Shiplastname")
  	  first_name = rs_Store("shipfirstname")
  	  username = rs_Store("user_id")
  	  password = rs_Store("password")
      Rs_store.Close
         'response.end
         'capture=""

      ShippedPr="yes"
      sql_insert = "Insert into Store_Purchases_Shippments (Store_id,oid,ShippedPr,Shipped_Date,Tracking_ID,Shipping_Company,Shippment_notes) values ("&Store_id&","&oid&",'"&ShippedPr&"','"&Shipped_Date&"','" & myTracking_ID & "','" & shippingcompany & "','"&Shippment_notes&"')"
      Conn_store.execute sql_insert

      'UPDATE STORE_PURCHASES TABLE WITH THE GENERAL STATUS OF SHIPPING
      sql_update = "update store_purchases set ShippedPr = '"&ShippedPr&"' where Oid = "&Oid&" and Store_id= "&Store_id
      Conn_store.execute sql_update
      if Request.Form("send_shipment_email") = 1 then
      	  ' send_email to customer regards to that shipping :
      	  Shipping_Email_To = email
      	  html_notifications=Request.Form("html_notifications")
		Shipping_Email_From = Request.Form("Shipping_Email_From")
      	Shipping_Email_Subject = Request.Form("Shipping_Email_Subject")
      	Shipping_Email_Body = Request.Form("Shipping_Email_Body")
      	Shipping_email_body = Replace(Shipping_email_body,"%FIRSTNAME%",first_name)
        	Shipping_email_body = Replace(Shipping_email_body,"%LASTNAME%",last_name)
        	Shipping_email_body = Replace(Shipping_email_body,"%LOGIN%",username)
        	Shipping_email_body = Replace(Shipping_email_body,"%PASSWORD%",password)
        	Shipping_Email_Body = replace(Shipping_Email_Body,"%ORDER%",oid)
      	Shipping_Email_Body = replace(Shipping_Email_Body,"%TRACKING%",myTracking_ID)
      	Shipping_Email_Body = replace(Shipping_Email_Body,"%SHIPCOMPANY%",shippingcompany)

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
  else
      Rs_store.Close
  end if
   if mycapture<>"" then
   sType = "Capture"
   trans_type = "2"
   if Real_Time_Processor = 1 then
      		  %><!--#include file="include\PlugnPay\pnp.asp"--><%
               elseif Real_Time_Processor = 2 then
      		  %><!--#include file="include\Authorizenet\Authorizenet.asp"-->
                        <%
      	  elseif Real_Time_Processor = 7 then
      		 %><!--#include file="include\Echo\Echo.asp"--><%
      		  elseif Real_Time_Processor = 10 then
	  %><!--#include file="include\linkpoint\linkpoint.asp"--><%
	  elseif Real_Time_Processor = 36 then
	  %><!--#include file="include\paypalpro\paypalpro.asp"--><%
      	  end if
         sql_update_purchases = "update store_purchases set Transaction_Type='2' where (oid = "&Oid&" or masteroid="&Oid&") AND Store_id ="&Store_id
      	  conn_store.Execute sql_update_purchases
                 end if
Next


Response.Redirect "orders.asp"

%>


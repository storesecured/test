<!--#include virtual="common/connection.asp"-->

<!--#include file="include/sub.asp"-->
<%
on error resume next

' read post from PayPal system and add 'cmd'
str = Request.Form
				
Txn_id = Request.Form("txn_id")

' post back to PayPal system to validate
str = str & "&cmd=_notify-validate"
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
'objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false
'objHttp.Send str

' assign posted variables to local variables
' note: additional IPN variables also available -- see IPN documentation
Receiver_email = Request.Form("receiver_email")
Store_Id = Request.Form("item_number")
Payment_status = Request.Form("payment_status")
total = Request.Form("payment_gross")
Txn_id = Request.Form("txn_id")
Payer_email = Request.Form("payer_email")
Store_id=Request.Form("custom")
level = Request.Querystring("Level")
Service = Request.querystring("Service")
Term = Request.Querystring("Term")
Term_Name = request.querystring("Term_Name")
sType=request.querystring("Type")
Grand_Total = request.querystring("Grand_Total")
Reseller_total = trim(request.form("hidResellerAmt"))
if Reseller_Total = "" then
   Reseller_Total=0
end if
Reseller_today = trim(request.form("hidResellerAmt2"))
if Reseller_Today = "" then
   Reseller_Today = 0
end if
HidEscAmount = trim(Request.Form("ESCAmount"))
ResellerRate = trim(Request.Form("ResellerRate"))
address_city=request.form("address_city")
address_country=request.form("address_country")
address_state=request.form("address_state")
address_street=request.form("address_street")
address_zip=request.form("address_zip")
first_name=request.form("first_name")
last_name=request.form("last_name")
payer_business_name=request.form("payer_business_name")
payer_email=request.form("payer_email")
item_name=request.form("item_name")
txn_type=request.form("txn_type")
subscr_id=request.form("subscr_id")
isverif = false
' Check notification validation


'response.write objHttp.responseText
Set rs_store = Server.CreateObject("ADODB.Recordset")
'if (objHttp.status <> 200 ) then
'elseif (objHttp.responseText = "VERIFIED") then
 if (Payment_status = "Completed" or Payment_status="Pending" or Payment_status="Refunded" or txn_type="subscr_signup") then
	'look up info in db
	if Receiver_email = "sales@easystorecreator.com" or Receiver_email = "upgrade@easystorecreator.com" or Receiver_email = "paypal@easystorecreator.com" or receiver_email=sPaypal_email then
		
		'transaction looks good so update verified status
		if (Payment_status="Pending") then
			AuthNumber = "Pending"
		else
			AuthNumber = Txn_id
		end if

		if txn_type="subscr_signup" then
			sql_select = "select * from Sys_Billing where Store_Id="&Store_Id
			rs_store.open sql_select, conn_store, 1, 1
			if rs_store.eof then
				newrec="1"
			else
				newrec="0"
				curr_amount=rs_store("amount")
				curr_term=rs_store("term")
			end if
			rs_store.close

			sql_update = "update store_settings set Trial_Version=0,Service_Type="&level&" where Store_id ="&Store_id
			conn_store.Execute sql_update

			if newrec = "1" then
				sql_insert = "Insert into Sys_Billing (Store_ID,Amount,Term, Payment_Method,Email,payment_term, reseller_portion,first_name,last_name,address,city,state,zip,country,company,notes) Values ("&Store_Id&","&Grand_Total&",'"&Term_Name&"','PayPal','"&Payer_email&"',"&Term&","&Reseller_total&",'"&first_name&"','"&last_name&"','"&address_street&"','"&address_city&"','"&address_state&"','"&address_zip&"','"&address_country&"','"&payer_business_name&"','"&notes&"')"
			else
				sql_insert = "Update Sys_Billing set Amount="&Grand_Total&",Reseller_Portion="&reseller_total&_
					",Term='"&Term_Name&"',Payment_Term="&Term&", Payment_Method='PayPal',Sys_Created='"&Now()&"',"&_
					" first_name='"&first_name&"',last_name='"&last_name&"', address='"&address_street&_
					"',city='"&address_city&"', state='"&address_state&"', zip='"&address_zip&"',country='"&_
					address_country&"', company='"&payer_business_name&"',email='"&payer_email&"',notes='"&subscr_id&_
					"' where Store_Id = "&Store_ID
			end if
			conn_store.Execute sql_insert
			response.end
			
		elseif txn_type="subscr_payment" then
			if curr_amount<>Grand_Total then
				'there appears to be an extra payment not matching, check this
				'Send_Mail Payer_email,Receiver_email,"Store Paid Problem","Store " & Store_Id & vbcrlf & request.form
				'send mail to client as well about error and how we will look into it
				'Send_Mail Receiver_email,Payer_email,"Store "&Store_Id&" Paypal Payment Issue","Dear Store Owner,"&vbcrlf&vbcrlf&"We have received what appears to be an extra payment for your store from Paypal.  According to our records this payment does not match your existing subscription and may be from a previous uncancelled subscription.  An email alert has been sent to our staff who will review this issue for you right away and cancel the subscription, refund the payment or credit it as necessary within the next 24 hours."
				
			end if
		end if

						
		if Payment_status <> "Pending" then
			if Payment_status="Refunded" then
				Service=Service&" Refund"
				sType="Custom"
			end if
  			sql_insert = "Insert into Sys_Payments (Store_Id,Amount,Auth_Number,Transaction_ID, Card_Ending,Payment_Description,Payment_Type,Payment_Term) values ("&_
  				Store_Id&","&total&",'"&AuthNumber&"','"&subscr_id&"','Paypal','"&Term_Name&" "&Service&"','"&sType&"',"&Term&")"
  			conn_store.Execute sql_insert
  			sql_update="update store_settings set overdue_payment=0 where store_id="&Store_Id
                        conn_store.Execute sql_update

					
			sql_select="select store_id from store_settings where store_id="&Store_Id
            rs_store.open sql_select, conn_store, 1, 1
			if rs_store.eof then
				Send_Mail Payer_email,Receiver_email,"Store " & Store_ID&" paypal no store",str
				Send_Mail Receiver_email,Payer_email,"Store "&Store_Id&" Paypal Payment Issue","Dear Store Owner,"&vbcrlf&vbcrlf&"We have received what appears to be an extra payment for your store from Paypal.  According to our records this payment should not of been sent as your store no longer exists.  It may be from a previous uncancelled subscription.  An email alert has been sent to our staff who will review this issue for you right away and cancel the subscription, refund the payment or credit it as necessary within the next 24 hours."
				 
            end if
    		rs_store.close
		end if
		'*********************************************************************************************************
		'code here to make an entry for the customer in the tbl_reseller_master
		service = term_name&" "&service	
		if SiteResellerID  <> "" then 
			sql_reseller = "Put_Reseller_Customer '"&SiteResellerID&"','"&Store_Id&"','"&term&"','"&Reseller_today&"','"&term_name&"','"&now()&"','"&service&"'"
			conn_store.Execute sql_reseller
		end if	

		'*********************************************************************************************************

		'Send_Mail Payer_email,Receiver_email,"Store Paid","Store " & Store_Id & " has paid " & total & " for easystorecreator "& service & " " & term_name & " service to be activated. (PayPal)" &Payment_status
	else
		Send_Mail Payer_email,Receiver_email,"Store " & Store_ID&" paypal error",str & "<BR>Service="&Service&"<BR>Term="&Term
	end if	
		
end if
' check that Payment_status=Completed
' check that Txn_id has not been previously processed
' check that Receiver_email is an email address in your PayPal account
' process payment
'end if
'set objHttp = nothing

%>


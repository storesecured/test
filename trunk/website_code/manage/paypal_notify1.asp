<!--#include virtual="common/connection.asp"-->

<!--#include file="include/sub.asp"-->
<%

' read post from PayPal system and add 'cmd'
str = Request.Form

Txn_id = Request.Form("txn_id")

' post back to PayPal system to validate
'str = str & "&cmd=_notify-validate"
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
Reseller_total = trim(request.form("hidResellerAmt"))

isverif = false
' Check notification validation

Set rs_store = Server.CreateObject("ADODB.Recordset")
if (Payment_status = "Completed" or Payment_status="Pending") then
			'look up info in db

		if Receiver_email = "sales@easystorecreator.com" or Receiver_email = "upgrade@easystorecreator.com" or Receiver_email = "paypal@easystorecreator.com" or Receiver_email = sPaypal_email then
				'transaction looks good so update verified status
				if (Payment_status="Pending") then
						AuthNumber = "Pending"
				else
						AuthNumber = Txn_id
				end if

				if Payment_status = "Completed" then
				  sql_insert = "Insert into Sys_Payments (Store_Id,Amount,Auth_Number,Transaction_ID, Card_Ending,Payment_Description,Payment_Type) values ("&_
						 Store_Id&","&total&",'"&AuthNumber&"','"&Verified_Ref&"','Paypal','"&Service&"','"&sType&"')"
				  conn_store.Execute sql_insert

				  sql_update = "Update store_settings set Custom_Amount=0,overdue_payment=0 where Store_Id="&Store_Id
				  conn_store.Execute sql_update
				  
				end if
				Send_Mail Payer_email,Receiver_email,"Store Paid","Store " & Store_Id & " has paid " & total & " for easystorecreator custom service. (PayPal)"&Custom_Description&vbcrlf&Payment_status
		end if
		
end if

%>


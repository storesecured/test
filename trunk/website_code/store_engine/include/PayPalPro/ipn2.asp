<!--#include virtual="include/header_noview.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include virtual="include/sub.asp"-->

<%

' dim some variables
Dim Item_name, Item_number, Payment_status, Payment_amount
Dim Txn_id, Receiver_email, Payer_email
Dim objHttp, str


'begin IPN handling
' read post from PayPal system and add 'cmd'
str = Request.Form & "&cmd=_notify-validate"


' post back to PayPal system to validate
set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
' set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
' set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
'objHttp.open "POST", "https://www.sandbox.paypal.com/cgi-bin/webscr", false
objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send str

' assign posted variables to local variables
item_name = Request.Form("item_name")
item_number = Request.Form("item_number")
payment_status = Request.Form("payment_status")
txn_id = Request.Form("txn_id")
parent_txn_id = Request.Form("parent_txn_id")
receiver_email = Request.Form("receiver_email")
payer_email = Request.Form("payer_email")
reason_code = Request.Form("reason_code")
business = Request.Form("business")
quantity = Request.Form("quantity")
invoice = Request.Form("invoice")
custom = Request.Form("custom")
tax = Request.Form("tax")
option_name1 = Request.Form("option_name1")
option_selection1 = Request.Form("option_selection1")
option_name2 = Request.Form("option_name2")
option_selection2 = Request.Form("option_selection2")
num_cart_items = Request.Form("num_cart_items")
pending_reason = Request.Form("pending_reason")
payment_date = Request.Form("payment_date")
mc_gross = Request.Form("mc_gross")
mc_fee = Request.Form("mc_fee")
mc_currency = Request.Form("mc_currency")
settle_amount = Request.Form("settle_amount")
settle_currency = Request.Form("settle_currency")
exchange_rate = Request.Form("exchange_rate")
txn_type = Request.Form("txn_type")
first_name = Request.Form("first_name")
last_name = Request.Form("last_name")
payer_business_name = Request.Form("payer_business_name")
address_name = Request.Form("address_name")
address_street = Request.Form("address_street")
address_city = Request.Form("address_city")
address_state = Request.Form("address_state")
address_zip = Request.Form("address_zip")
address_country = Request.Form("address_country")
address_status = Request.Form("address_status")
payer_email = Request.Form("payer_email")
payer_id = Request.Form("payer_id")
payer_status = Request.Form("payer_status")
payment_type = Request.Form("payment_type")
notify_version = Request.Form("notify_version")
verify_sign = Request.Form("verify_sign")



  



' Check notification validation
if (objHttp.status <> 200 ) then
' HTTP error handling

elseif (objHttp.responseText = "VERIFIED") then
' check that Payment_status=Completed
' check that Txn_id has not been previously processed
' check that Receiver_email is your Primary PayPal email
' check that Payment_amount/Payment_currency are correct
' process payment


'******************my IPN code**********
'response.write(payment_type & "<br>" & payment_status & "<br>" & txn_id)
mystr2="payment_type: " &  payment_type & ", payment_status: " & payment_status & ",txn_id : " & txn_id
  '********log to file*****************
  Sub WriteToFile(strFilePath, strData, iLineNumber)
   Dim objFSO, objFile, arrLines
   Dim strAllFile, x
   Set objFSO=Server.CreateObject("Scripting.FileSystemObject")
   strAllFile=""
   If objFSO.FileExists(strFilePath) Then
      Set objFile=objFSO.OpenTextFile(strFilePath)
      If Not(objFile.AtEndOfStream) Then
         strAllFile = objFile.ReadAll
      End If
      objFile.Close
   End If
   arrLines = Split(strAllFile, VBCrLf)
   Set objFile=objFSO.CreateTextFile(strFilePath)
   For x=0 To UBound(arrLines)
      If (iLineNumber-1)=x Then objFile.WriteLine(strData)
      objFile.WriteLine(arrLines(x))
   Next
   If iLineNumber>=UBound(arrLines) Then objFile.WriteLine(strData)
   objFile.Close
   Set objFile=Nothing
   Set objFSO=Nothing
  End Sub



  '************************************
	dim  myverify
	if parent_txn_id<>"" then
		mytrx= parent_txn_id
	else
		mytrx= txn_id
	end if

	if payment_status="Refunded" then
     	sql_update_purchases = "update store_purchases set Transaction_Type='3' where Store_id ="&Store_id&" and verified_ref='"&mytrx&"'"
		conn_store.Execute sql_update_purchases
	else
		if payment_type="instant" then
			if payment_status="Completed" or payment_status="Processed"  or payment_status="Canceled-Reversal" or payment_status="Partially-Refunded" or (payment_status="Pending" and pending_reason="authorization") then
	  			myverify=1
	  		else
	  			myverify=0
	  		end if
	   	else
	     	if payment_status="Completed" or payment_status="Processed"  or payment_status="Canceled-Reversal" or payment_status="Partially-Refunded" then
	  			myverify=1
	  		else
	  			myverify=0
	   		end if
		end if

		sql_update = "exec wsp_ipn_update "&store_id &","& myverify & ",'" & mytrx & "','" & str & "','" & payment_type & "','" & payment_status & "','" & pending_reason & "';"
		conn_store.Execute sql_update
    end if

'***************************************




elseif (objHttp.responseText = "INVALID") then
' log for manual investigation
' add code to handle the INVALID scenario

else
' error
end if
set objHttp = nothing
%>


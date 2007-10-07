<%
Server.ScriptTimeout = 180

sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst
Do While Not Rs_Store.EOF
	select case Rs_store("Property")
		case "merchant_echo_id"
			merchant_echo_id = decrypt(Rs_store("Value"))
		case "merchant_echo_pin"
			merchant_echo_pin = decrypt(Rs_store("Value"))
	end select
	Rs_store.MoveNext
Loop
Rs_store.Close


rs_store.open "select * from store_customers where record_type=0 and cid="&cid, conn_store, 1, 1
If not rs_Store.eof then
	first_name=rs_Store("First_name")
	last_name=rs_Store("Last_name")
	address1=rs_Store("Address1")
	address2=rs_Store("Address2")
	city=rs_Store("City")
	state=rs_store("state")
	zip=rs_Store("zip")
	country=rs_Store("Country")
	phone=rs_Store("Phone")
	fax=rs_Store("fax")
	ccid = rs_Store("CCID")
	if rs_Store("Tax_Exempt") then
		TaxExempt="TRUE"
	else
		 TaxExempt="FALSE"
	end if
End If
rs_store.close

rs_store.open "Select * from Store_Purchases where oid = "&Oid&" and Store_id ="&Store_id,conn_store,1,1
If not rs_Store.eof then
	Shipping_Method_Price = Rs_store("Shipping_Method_Price")
	Tax = Rs_store("Tax")
	Cust_PO = Rs_store("Cust_PO")
	ShipFirstname=rs_Store("ShipFirstname")
	ShipLastname=rs_Store("ShipLastname")
	ShipAddress1=rs_Store("ShipAddress1")
	ShipAddress2=rs_Store("ShipAddress2")
	ShipCity=rs_Store("ShipCity")
	ShipState=rs_store("ShipState")
	Shipzip=rs_Store("Shipzip")
	ShipCountry=rs_Store("ShipCountry")
	ShipPhone=rs_Store("ShipPhone")
	ShipFax=rs_Store("ShipFax")
	ShipEmail=rs_Store("ShipEmail")
	ShipCompany=rs_Store("ShipCompany")
	Verified_Ref=rs_Store("Verified_Ref")
	CardNumber=rs_Store("CardNumber")
	CardExpiration=rs_Store("CardExpiration")
	AuthNumber=rs_Store("AuthNumber")
End If
rs_Store.Close

if GGrandTotal>Grand_Total then
   fn_error "You cannot capture or credit an amount greater than the original authorization amount."
end if

CheckSerial = Request.Form("CheckSerial")
GGrand_Total = FormatNumber(GGrand_Total,2)
Set Echo = Server.CreateObject("ECHOCom.Echo")
Randomize ' for the counter
'******************* Required Fields *******************************
Echo.EchoServer="https://wwws.echo-inc.com/scripts/INR200.EXE"
Echo.counter=Round((100000 * Rnd))   ' This randomizing is just for demo purposes.
Echo.order_type="S"
Echo.merchant_echo_id=merchant_echo_id
Echo.merchant_pin=merchant_echo_pin
Echo.billing_ip_address=Request.ServerVariables("REMOTE_ADDR")
Echo.merchant_email=Store_Email
Echo.grand_total=FormatNumber(GGrand_Total,2)
Echo.purchase_order_number=Cust_PO
Echo.sales_tax=Tax
Echo.merchant_trace_nbr=Oid

Echo.debug="N"


if Payment_Method="eCheck" then
	if auth_capture then
		sTransType = "DD"
	else
		sTransType = "DV"
	end if

	BankABA = Request.Form("BankABA")
	BankAccount = Request.Form("BankAccount")
	acct_type = Request.Form("acct_type")
	org_type = Request.Form("org_type")
	DrvState = Request.Form("DrvState")
	DrvNumber = Request.Form("DrvNumber")
	dobd = Request.Form("dobd")
	dobm = Request.Form("dobm")
	doby = Request.Form("doby")
	transdate = Year(now()) & Month(now()) & Day(now()) & Hour(now()) & Minute(now()) & Second(now())
	if org_type = "I" then
		sType = "P"
	else
		sType = "B"
	end if
	if acct_type = "CHECKING" then
		sType = sType & "C"
	else
		sType = sType & "S"
	end if
	Echo.ec_id_type = "DL"
	Echo.transaction_type=sTransType
	Echo.ec_account=BankAccount
	Echo.ec_account_type =sType
	Echo.ec_address1=address1
	Echo.ec_address2=address2
	Echo.ec_bank_name=BankName
	Echo.ec_city=city
	Echo.ec_email=ShipEmail
	Echo.ec_first_name=first_name
	'Echo.ec_id_country="US"
	Echo.ec_id_exp_mm=dobm
	Echo.ec_id_exp_dd=dobd
	Echo.ec_id_exp_yy=doby
	Echo.ec_id_number=DrvNumber
	Echo.ec_id_state=DrvState

	Echo.ec_last_name=last_name
	Echo.ec_payee=Store_name
	Echo.ec_payment_type="WEB"
	Echo.ec_rt=BankABA
	Echo.ec_serial_number=CheckSerial
	Echo.ec_state=State
	Echo.ec_transaction_dt=transdate
	Echo.ec_zip=zip

else

	if sType="Capture" then
		sTypeTrans = "DS"
	else
		sTypeTrans = "CR"
	end if

	Echo.transaction_type=sTypeTrans
	Echo.billing_first_name =first_name
	Echo.billing_last_name=last_name
	Echo.billing_company_name=ShipCompany
	Echo.billing_phone=phone
	Echo.billing_address1=address1
	Echo.billing_address2=address2
	Echo.billing_city=city
	Echo.billing_state=State
	Echo.billing_zip=zip
	Echo.billing_email=ShipEmail
	Echo.billing_country=country
	Echo.billing_fax=fax
	Echo.cc_number=decrypt(CardNumber)
	Echo.original_amount=FormatNumber(GGrand_Total,2)
	'Echo.original_tran_date_mm
	'Echo.original_tran_date_yyyy
	'Echo.original_tran_date_dd
	'Echo.original_reference
	'Echo.order_number
	sArray = split(CardExpiration,"/")
	Echo.ccexp_year=sArray(1)
	Echo.ccexp_month=sArray(0)
	Echo.cnp_recurring="N"
     Echo.order_number=Verified_Ref
     Echo.authorization =AuthNumber
end if

If Echo.Submit Then
Else
	response.redirect "error.asp?Message_id=101&Message_Add="&Server.UrlEncode(Echo.echotype2 & "--" & transdate)
end If

on error goto 0

AuthNumber=Echo.authorization

sText = Echo.Echotype3
iIndex = InStrRev(sText,"<avs_result>")
if iIndex > 0 then
	avs_result = Mid(sText,iIndex+12,1)
end if

iIndex = InStrRev(sText,"<security_result>")
if iIndex > 0 then
	card_verif = Mid(sText,iIndex+17,1)
end if

iIndex = InStrRev(sText,"<order_number>")
if iIndex > 0 then
	Verified_Ref = Mid(sText,iIndex+14,14)
end if

'avs_result=xmlDoc.getElementsByTagName("avs_result").item(0).text
'card_verif=xmlDoc.getElementsByTagName("security_result").item(0).text



%>


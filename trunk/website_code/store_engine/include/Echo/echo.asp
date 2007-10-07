<%

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


if Tax_Exempt then
	TaxExempt="TRUE"
else
	TaxExempt="FALSE"
end if

CheckSerial = Request.Form("CheckSerial")
GGrand_Total = FormatNumber(GGrand_Total,2)
Set Echo = Server.CreateObject("ECHOCom.Echo")
Randomize ' for the counter
'******************* Required Fields *******************************
Echo.EchoServer="https://wwws.echo-inc.com/scripts/INR200.EXE"
Echo.counter=Oid
Echo.order_type="S"
Echo.merchant_echo_id=merchant_echo_id
Echo.merchant_pin=merchant_echo_pin
Echo.billing_ip_address=Request.ServerVariables("REMOTE_ADDR")
Echo.merchant_email=Store_Email
Echo.grand_total=GGrand_Total
Echo.purchase_order_number=Cust_PO
Echo.sales_tax=formatnumber(Tax,2)
Echo.merchant_trace_nbr=Oid

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

'Echo.EchoDebug="F"


if Payment_Method="eCheck" then
	if auth_capture then
		sTransType = "DD"
	else
		sTransType = "DV"
	end if

	BankABA = Request.Form("BankABA")
	BankAccount = Request.Form("BankAccount")
	BankName = Request.Form("BankName")
	acct_type = Request.Form("acct_type")
	org_type = Request.Form("org_type")
	DrvState = Request.Form("DrvState")
	DrvNumber = Request.Form("DrvNumber")
	dobd = Request.Form("dobd")
	dobm = Request.Form("dobm")
	doby = Request.Form("doby")
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
	Echo.ec_id_country="US"
	Echo.ec_id_exp_mm=dobm
	Echo.ec_id_exp_dd=dobd
	Echo.ec_id_exp_yy=doby
	Echo.ec_id_number=DrvNumber
	Echo.ec_id_state=DrvState
	Echo.billing_phone=Replace(phone,"-","")

	Echo.ec_last_name=last_name
	Echo.ec_payee=Store_name
	Echo.ec_payment_type="WEB"
	Echo.ec_rt=BankABA
	Echo.ec_serial_number=CheckSerial
	Echo.ec_state=State
	Echo.ec_zip=zip

else
	if auth_capture then
		sType = "EV"
	else
		sType = "AV"
	end if

	Echo.transaction_type=sType

	Echo.cc_number=CardNumber
	Echo.ccexp_year=Request.Form("yy")
	Echo.ccexp_month=Request.Form("mm")
	Echo.cnp_recurring="N"
	if Use_CVV2 then
		Echo.cnp_security=CardCode
	end if
end if

If Echo.Submit Then
Else
	fn_purchase_decline oid,"The transaction was rejected by the payment processor:<BR>"&Echo.echotype2 & "--" & transdate
end If

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
	Verified_Ref = Mid(sText,iIndex+14,16)
end if

'avs_result=xmlDoc.getElementsByTagName("avs_result").item(0).text
'card_verif=xmlDoc.getElementsByTagName("security_result").item(0).text

if Auth_Capture then
	trans_type = 1
else
	trans_type = 0
end if


%>


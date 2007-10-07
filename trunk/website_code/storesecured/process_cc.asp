<%
sAuthorized=0

if payment_method<>"eCheck" then
    Set xObj = CreateObject("SOFTWING.ASPtear")
    Post_String = ""
    Post_String = Post_String &"x_login="&Server.UrlEncode("esystrcrt77834RR")
    Post_String = Post_String &"&x_tran_key="& Server.UrlEncode("yVSpJwJL69RDD9vs")
    Post_String = Post_String &"&x_version="& Server.UrlEncode("3.1")
    Post_String = Post_String &"&x_delim_data=True"
    Post_String = Post_String &"&x_delim_char=|"
    Post_String = Post_String &"&x_encap_char="
    Post_String = Post_String &"&x_ADC_URL=False"
    Post_String = Post_String &"&x_relay_response=FALSE"
    Post_String = Post_String &"&x_test_request="
    Post_String = Post_String &"&x_trans_ID="& Server.UrlEncode(oid & "-"&now())
    Post_String = Post_String &"&x_tax_exempt=1"
    Post_String = Post_String &"&x_description=" & Server.UrlEncode(left(sDescription,255))
    Post_String = Post_String &"&x_tax=0"
    Post_String = Post_String &"&ForceReload="& Server.UrlEncode(Now())
    Post_String = Post_String &"&x_ship_to_country="& Server.UrlEncode(country)
    Post_String = Post_String &"&x_ship_to_zip="& Server.UrlEncode(zip)
    Post_String = Post_String &"&x_ship_to_state="& Server.UrlEncode(state)
    Post_String = Post_String &"&x_ship_to_city="& Server.UrlEncode(city)
    Post_String = Post_String &"&x_ship_to_address="& Server.UrlEncode(address)
    Post_String = Post_String &"&x_ship_to_first_Name="& Server.UrlEncode(first_name)
    Post_String = Post_String &"&x_ship_to_last_Name="& Server.UrlEncode(last_name)
    Post_String = Post_String &"&x_first_name="& Server.UrlEncode(first_name)
    Post_String = Post_String &"&x_last_name="& Server.UrlEncode(last_name)
    Post_String = Post_String &"&x_company="& Server.UrlEncode(company)
    Post_String = Post_String &"&x_ship_to_company="& Server.UrlEncode(company)
    Post_String = Post_String &"&x_state="& Server.UrlEncode(state)
    Post_String = Post_String &"&x_type="&server.urlencode("AUTH_CAPTURE")
    Post_String = Post_String &"&x_amount="&Server.UrlEncode(total)
    Post_String = Post_String &"&x_address="& Server.UrlEncode(address)
    Post_String = Post_String &"&x_city="& Server.UrlEncode(city)
    Post_String = Post_String &"&x_country="& Server.UrlEncode(country)
    Post_String = Post_String &"&x_zip="& Server.UrlEncode(zip)
    Post_String = Post_String &"&x_phone="& Server.UrlEncode(phone)
    Post_String = Post_String &"&x_fax="& Server.UrlEncode(fax)
    Post_String = Post_String &"&x_cust_id="& Server.UrlEncode(Store_Id)
    Post_String = Post_String &"&x_customer_ip="&Server.UrlEncode(Request.ServerVariables("REMOTE_ADDR"))
    Post_String = Post_String &"&x_email="&Server.UrlEncode(email)
    Post_String = Post_String &"&x_freight=0"
    Post_String = Post_String &"&x_duty=0"
    Post_String = Post_String &"&x_invoice_num="&server.urlencode(Store_Id&"-"&now())
    Post_String = Post_String &"&x_po_num="&server.urlencode(sPONum)
    if sRecurring = "Y" then
       Post_String = Post_String &"&x_recurring_billing=YES"
    end if
    Post_String = Post_String &"&x_method=CC"
    Post_String = Post_String &"&x_card_num="& Server.UrlEncode(cc_num)
    Post_String = Post_String &"&x_exp_date="& Server.UrlEncode(mm&yy)
    Post_String = Post_String &"&x_card_code="& Server.UrlEncode(CardCode)
    
    
    'strResult = xObj.Retrieve("https://secure.authorize.net/gateway/transact.dll", 1, Post_String, "", "")
    'Returned_Var_Array = Split(strResult,"|")
    '	response_code = replace(Returned_Var_Array(0),"""","")
    '	response_subcode = replace(Returned_Var_Array(1),"""","")
    '	response_reason_code = replace(Returned_Var_Array(2),"""","")
    '	response_reason_text = replace(Returned_Var_Array(3),"""","")
    '	AuthNumber = replace(Returned_Var_Array(4),"""","")
    '	avs_result = replace(Returned_Var_Array(5),"""","")
    '	Verified_Ref = replace(Returned_Var_Array(6),"""","")
    '	card_verif = replace(Returned_Var_Array(38),"""","")
     'Set xObj = nothing

     If response_code="1" then
        sAuthorized=1
        trans_type = 1
     end if

end if

if sAuthorized=0 then
  Set Echo = Server.CreateObject("ECHOCom.Echo")
  Randomize ' for the counter
  '******************* Required Fields *******************************
  Echo.EchoServer="https://wwws.echo-inc.com/scripts/INR200.EXE"
  if counter = "" then
  	counter = DatePart("n", now())
  end if
  Echo.counter=counter
  Echo.order_type="S"
  Echo.merchant_echo_id="123>4682520"
  Echo.merchant_pin="54121678"
  'Echo.merchant_echo_id="760>5802579"
  'Echo.merchant_pin="24342434"
  Echo.billing_ip_address=Request.ServerVariables("REMOTE_ADDR")
  Echo.merchant_email="sales@easystorecreator.com"
  Echo.grand_total=total
  Echo.sales_tax=0
  Echo.merchant_trace_nbr=Oid
  
  if payment_method<>"eCheck" then
  
    Echo.billing_first_name =first_name
    Echo.billing_last_name=last_name
    Echo.billing_company_name=company
    Echo.billing_phone=phone
    Echo.billing_address1=address
    Echo.billing_address2=""
    Echo.billing_city=city
    Echo.billing_state=state
    Echo.billing_zip=zip
    Echo.billing_email=email
    Echo.billing_country=country
    Echo.billing_fax=fax
    if processtype = "auth_capture" then
  		sTransType = "EV"
    else
  		sTransType = "AV"
    end if
    Echo.cc_number=cc_num
    Echo.ccexp_year=yy
    Echo.ccexp_month=mm
    Echo.cnp_recurring=sRecurring
    Echo.cnp_security=CardCode
  else
  	if processtype = "auth_capture" then
  		sTransType = "DD"
  	else
  		sTransType = "DV"
  	end if
  	if org_type = "I" then
  		sType = "P"
  	else
  		'sType = "B"
  		sType = "P"
  	end if
  	if acct_type = "CHECKING" then
  		sType = sType & "C"
  	else
  		sType = sType & "S"
  	end if
  	Echo.ec_id_type = "DL"
  	Echo.ec_account=BankAccount
  	Echo.ec_account_type =sType
  	Echo.ec_address1=address
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
  
  	Echo.ec_last_name=last_name
  	Echo.ec_payee="Easystorecreator"
  	Echo.ec_payment_type="WEB"
  	Echo.ec_rt=BankABA
  	Echo.ec_serial_number=CheckSerial
  	Echo.ec_state=State
  	Echo.ec_zip=zip
  
  
  end if
  'Echo.EchoDebug="F"
  
  Echo.transaction_type=sTransType
  Echo.billing_phone=Replace(phone,"-","")
  
  Echo.product_description=sDescription
  Echo.purchase_order_number=sPONum
  
  If Echo.Submit Then
  Else
  	 response.redirect "error.asp?Message_id=101&Message_Add="&Server.UrlEncode("An error has occured, your transaction could not be processed, please check your information and try again.<BR>"&Echo.echotype2 & "--" & transdate)
  	 response.write Echo.echotype2
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
  	Verified_Ref = Mid(sText,iIndex+14,14)
  end if
  
  if Auth_Capture then
  	trans_type = 1
  else
  	trans_type = 0
  end if
end if


%>

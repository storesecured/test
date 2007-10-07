<%
sAuthorized=0

if payment_method<>"eCheck" then
    sDescription="Reseller"
    Set client = Server.CreateObject("PFProCOMControl.PFProCOMControl.1")
    parmList = "TRXTYPE=S"
    parmList = parmList + "&ACCT="&cc_num
    'parmList = parmList + "&ACCT=4111111111111111"
    parmList = parmList + "&PWD=ankle237"
    parmList = parmList + "&USER=blac6789"
    parmList = parmList + "&VENDOR="
    parmList = parmList + "&PARTNER=StoreSecured"
    parmList = parmList + "&EXPDATE="&replace(mm&yy," ","")
    parmList = parmList + "&AMT="&formatnumber(cdbl(total),2)
    parmList = parmList + "&STREET="&replace(Address,"&#39;","")
    parmList = parmList + "&EMAIL="&replace(email,"&#39;","")
    parmList = parmList + "&ZIP="&replace(zip,"&#39;","")
    parmList = parmList + "&FIRSTNAME="&replace(first_name,"&#39;","")
    parmList = parmList + "&LASTNAME="&replace(Last_name,"&#39;","")
    parmList = parmList + "&COMPANYNAME="&replace(Company,"&#39;","")
    parmList = parmList + "&TENDER=C"
    parmList = parmList + "&PONUM="&Store_Id
    parmList = parmList + "&SHIPTOZIP="&replace(zip,"&#39;","")
    parmList = parmList + "&TAXAMT=0"
    parmList = parmList + "&TAXEXEMPT=Y"
    parmList = parmList + "&CITY="&replace(City,"&#39;","")
    parmList = parmList + "&CUSTCODE="&Store_Id
    parmList = parmList + "&STATE="&replace(State,"&#39;","")
    parmList = parmList + "&RECURRING=Y"
    parmList = parmList + "&FREIGHTAMT=0"
    parmList = parmList + "&COMMENT1="&left(sDescription,255)
    if Session("Custom_Description")<>"" then
       parmList = parmList + "&COMMENT2="&Custom_Description
    end if
    if CardCode<>"" then
       parmList = parmList + "&CVV2="&CardCode
    end if

    'for real use payflow.verisign.com
    'for testing use test-payflow.verisign.com
    
    Ctx1 = client.CreateContext("payflow.verisign.com", 443, 30, "", 0, "", "")
    curString = client.SubmitTransaction(Ctx1, parmList, Len(parmList))
    client.DestroyContext (Ctx1)

    Do while Len(curString) <> 0
      	if InStr(curString,"&") Then
      		varString = Left(curString, InStr(curString , "&" ) -1)
      	else
      		varString = curString
      	end if
      	name = Left(varString, InStr(varString, "=" ) -1)
      	value = Right(varString, Len(varString) - (Len(name)+1))
      	select case name
      		case "RESULT" 
      			resultval = value
      		case "RESPMSG"
      			respMessage = value
      		case "AUTHCODE"
      			AuthNumber = value
      		case "PNREF"
      			Verified_Ref = value
      		case "AVSADDR"
      			avsaddr = value
      		case "AVSZIP"
      			avszip = value
      		case "IAVS"
      			iavs = value
      		case "CVV2MATCH"
      			card_code_verif = value
      	end select
      	if Len(curString) <> Len(varString) Then 
      		curString = Right(curString, Len(curString) - (Len(varString)+1))
      	else
      		curString = ""
      	end if
      Loop
      
      if avsaddr = "N" and avszip = "N" then
      	avs_result = "N"
      elseif avsaddr = "Y" and avszip = "Y" then
      	avs_result = "Y"
      elseif avsaddr = "N" and avszip = "Y" then
      	avs_result = "Z"
      elseif avsaddr = "Y" and avszip = "N" then
      	avs_result = "A"
      elseif iavs = "Y" then
      	avs_result = "G"
      elseif iavs = "X" or avszip = "X" or avsaddr = "X" then
      	avs_result = "S"
      end if
      
      if card_code_verif = "Y" then
      	card_verif = "M"
      elseif card_code_verif = "N" then
      	card_verif = "N"
      elseif card_code_verif = "X" then
      	card_verif = "X"
      elseif card_code_verif = "" or not Use_CVV2 then
      	card_verif = "P"
      end if
      
      If resultval <> 0 then
      	response.redirect "error.asp?Message_id=101&Message_Add="&Server.UrlEncode("An error has occured, your transaction could not be processed, please view the error below and if applicable try again.<BR><BR>"&respMessage)

      else
        sAuthorized=1
        trans_type = 1
      End IF

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
  'Echo.merchant_echo_id="123>4682520"
  'Echo.merchant_pin="54121678"
  Echo.merchant_echo_id="760>5802579"
  Echo.merchant_pin="24342434"
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

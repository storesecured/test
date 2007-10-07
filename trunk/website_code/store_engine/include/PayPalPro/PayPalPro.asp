<!--#include file="callerservice.asp"-->
<%
on error resume next
set rs_Store =  server.createobject("ADODB.Recordset")
sql_real_time = "exec wsp_real_time_property "&Store_Id&",36;"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst
Do While Not Rs_Store.EOF
    select case Rs_store("Property")
        case "PayPal_Pro_Api_Username"
	        PayPal_Pro_API_username = decrypt(Rs_store("Value"))
        case "PayPal_Pro_Password"
	        PayPal_pro_Password = decrypt(Rs_store("Value"))
        case "PayPal_Pro_Currency"
	        PayPal_Pro_Currency = decrypt(Rs_store("Value"))
        case "PayPal_Pro_Signature"
	        PayPal_Pro_Signature = decrypt(Rs_store("Value"))
    end select
    Rs_store.MoveNext
Loop
Rs_store.Close

'************************************
if PayPal_Pro_Signature<>"" then
'Authorize or Sale
If Auth_Capture then
    sPaymentType="Sale"
    trans_type = 1
else
    sPaymentType="Authorization"
    trans_type = 0
end if

if Payment_Method ="American Express" then
    sPaymentMethod="Amex"
else
    sPaymentMethod=Payment_Method
end if

currentpath = "http://" & Request.ServerVariables("SERVER_NAME")

'PayPal_Pro_notifyurl=  currentpath & "/include/paypalpro/paypal_ipn.asp"
PayPal_Pro_notifyurl =  currentpath & "/include/paypalpro/ipn2.asp"

country_code = fn_country_code(country)
ship_country_code = fn_country_code(ShipCountry)

nvpstr	=	"USER="&server.URLEncode(PayPal_Pro_API_username)&_
    "&PWD="&server.URLEncode(PayPal_pro_Password)&_
    "&SIGNATURE="&server.URLEncode(PayPal_Pro_Signature)&_
    "&CURRENCYCODE=" &server.URLEncode(PayPal_Pro_Currency)&_
    "&VERSION="&server.URLEncode(2.3)&_
    "&BUTTONSOURCE="&server.URLEncode("EasyStoreCreator_Cart_DP_US")&_
    "&METHOD=doDirectPayment"&_
    "&PAYMENTACTION=" &server.URLEncode(sPaymentType) & _
    "&IPADDRESS="&server.URLEncode(Request.ServerVariables("REMOTE_ADDR"))&_
    "&AMT="&server.URLEncode(GGrand_Total) &_
    "&CREDITCARDTYPE="&server.URLEncode(sPaymentMethod) &_
    "&ACCT="&server.URLEncode(request.form("CardNumber")) & _
    "&EXPDATE=" & server.URLEncode(request.form("mm")&"20"& request.form("yy")) &_
    "&CVV2=" & server.URLEncode(request.form("CardCode")) &_
    "&FIRSTNAME=" & server.URLEncode(first_name) &_
    "&LASTNAME=" & server.URLEncode(last_Name) &_
    "&STREET=" & server.URLEncode(address1 )&_
    "&CITY=" & server.URLEncode(city) &_
    "&STATE=" & server.URLEncode(state) &_
    "&ZIP=" &server.URLEncode(zip) &_
    "&COUNTRYCODE="&server.URLEncode(country_code) &_
    "&INVNUM="&server.URLEncode(oid&now())&_
    "&EMAIL="&server.URLEncode(email)&_
    "&STREET2="&server.URLEncode(address2)&_
    "&PHONENUM="&server.URLEncode(phone)&_
    "&SHIPTONAME="&server.URLEncode(ShipFirstname&" "&ShipLastname)&_
    "&SHIPTOSTREET="&server.URLEncode(ShipAddress1)&_
    "&SHIPTOCITY="&server.URLEncode(ShipCity)&_
    "&SHIPTOCOUNTRY="&server.URLEncode(ship_country_code)&_
    "&SHIPTOSTREET2="&server.URLEncode(ShipAddress2)&_
    "&SHIPTOSTATE="&server.URLEncode(ShipState)&_
    "&SHIPTOPHONENUM="&server.URLEncode(ShipPhone)&_
    "&SHIPTOZIP="&server.URLEncode(Shipzip)&_
    "&NOTIFYURL=" &Server.URLEncode(PayPal_Pro_notifyurl)
    
'call sub_write_log("Request="&nvpstr)

Set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
'objHttp.open "POST", "https://api-3t.sandbox.paypal.com/nvp", False
objHttp.open "POST", "https://api-3t.paypal.com/nvp", False
WinHttpRequestOption_SslErrorIgnoreFlags=4
objHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
objHttp.Send nvpstr
objHttp.WaitForResponse(240)
sResponse=objHttp.responseText

Set  nvpResponseCollection =deformatNVP(sResponse)
'call sub_write_log("Response="&sResponse)
sErrorMessage=nvpResponseCollection("L_LONGMESSAGE0")
sAck=nvpResponseCollection("ACK")
Set objHttp = Nothing 

if sAck<>"Success" then
    fn_error sErrorMessage		
end if
avs_result=nvpResponseCollection("AVSCODE")
card_verif = nvpResponseCollection("CVV2MATCH")
Verified_Ref = nvpResponseCollection("TRANSACTIONID")
else
  'instantiate te object
		set PayPalDirect = Server.CreateObject("IBizPayPalAsp.DirectPayment")
		
	   	Key_Folder = fn_get_sites_folder(Store_Id,"Key") 

		'merchant information
		PayPalDirect.Username = PayPal_Pro_API_username
		PayPalDirect.Password = PayPal_pro_Password
		PayPalDirect.SSLCertStoreType = 4
		PayPalDirect.SSLCertStore = Key_Folder&"cert_key_pem.txt" 
		PayPaldirect.SSLCertStorePassword = ""
		PayPalDirect.SSLCertSubject = "*"

		if err.number<>0 then
            if err.number="-2146808016" or err.number="20272" then
               response.redirect Switch_Name&"error.asp?Message_Id=101&Message_Add="&Server.urlencode("The Paypal Pro certificate could not be found.<BR><HR><BR>Store Owner please ensure your certificate is uploaded properly.")
            else
               response.redirect Switch_Name&"error.asp?Message_Id=101&Message_Add="&Server.urlencode(err.description &err.number)
            end if
      end if

		'order information
		PayPalDirect.OrderTotal =GGrand_Total
		PayPalDirect.OrderDescription = "Order "&oid

		'card details
		if Payment_Method = "Visa" then
			cType = 0
		elseif Payment_Method="Mastercard" then
			cType = 1
		elseif Payment_Method ="Discover" then
			cType = 2
		elseif Payment_Method ="American Express" then
			cType = 3

		end if

		PayPalDirect.CardType = cType
		PayPalDirect.CardNumber =request.form("CardNumber")


		PayPalDirect.CardExpMonth = cint(request.form("mm"))
		PayPalDirect.CardExpYear = cint("20" + request.form("yy"))
		PayPalDirect.CardCVV =request.form("CardCode")
		
		'Payer Information
		PayPalDirect.PayerFirstName = first_name
		PayPalDirect.PayerLastName = last_Name
		PayPalDirect.PayerStreet1 = address1
		PayPalDirect.PayerCity = city
		PayPalDirect.PayerState = State
		
		country_code = fn_country_code(country)

      PayPalDirect.PayerCountryCode = country_code
		PayPalDirect.PayerPostalCode =Zip
		PayPalDirect.PayerIPAddress = Request.ServerVariables("REMOTE_ADDR")

		PayPalDirect.URL = "https://api.paypal.com/2.0/"
	'	PayPalDirect.URL = "https://api.sandbox.paypal.com/2.0/"
      if PayPal_Pro_Currency <>"" then
         PayPalDirect.CurrencyCode=PayPal_Pro_Currency
      end if

		'Authorize or Sale
		If Auth_Capture then
			PayPalDirect.Sale()
			trans_type = 1
		else
			PayPalDirect.Authorize()
			trans_type = 0
		end if

		if err.Number <> 0 then
		   sDescription=err.description
			if err.Number = 20302 then
				sMessage= "Connection timed out"
         elseif err.Number=380 then
            sMessage="The server this site is hosted on does not support Paypal Pro.<BR><HR><BR>Store Owner you wish to use Paypal Pro please submit a support request asking that your store be moved to a Paypal Pro compatible server."
         elseif err.number="20528" and instr(sDescription,"[10752]")>0 then
            sMessage="Transaction refused"
         elseif err.number="20528" and instr(sDescription,"[10748]")>0 then
            sMessage="The card code value is missing.<BR><HR><BR>Store Owner please ensure this is enabled for capture in your store as Paypal Pro requires it."
         elseif err.number="20528" and instr(sDescription,"[10002]")>0 then
            sMessage="The Paypal Pro username and password are invalid or not activated.<BR><HR><BR>Store owner please double check that this information is entered correctly and that the Paypal Pro account is fully activated.<BR><BR>Note: If the PayPal Pro button on the view cart page works but the final checkout does not, your username and password are correct but your account is not yet fully activated with PayPal for direct checkout."
         else
            sMessage="Please try again later.<BR><BR>"&err.description
			end if
			response.redirect Switch_Name&"error.asp?Message_id=101&Message_Add="&Server.UrlEncode("The transaction was rejected by the payment processor:<BR><BR>"&sMessage)

		end if


      avs_result = PayPalDirect.ResponseAVS
    	Verified_Ref = PayPalDirect.ResponseTransactionId
    	card_verif = PayPalDirect.ResponseCVV


		PayPalDirect.Reset()
		set PayPalDirect =nothing
end if
%>

<!--#include virtual="include/header_noview.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!-- #include virtual="include/paypalpro/CallerService.asp" -->
<!--#include virtual="common/common_functions.asp"-->
<%
dim sPaymentType

Call sub_check_minimum()
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
set rs_Store = nothing

'*********authentication method*************
if PayPal_Pro_Signature<>"" then

'----------------------------------------------------------------------
' Define the PayPal URL.  This is the URL that the buyer is
' first sent to to authorize payment with their paypal account
' change the URL depending if you are testing on the sandbox
' or going to the live PayPal site
' For the sandbox, the URL is
' https://www.sandbox.paypal.com/webscr&cmd=_express-checkout&token=
' For the live site, the URL is
' https://www.sandbox.paypal.com/webscr&cmd=_express-checkout&token=
'------------------------------------------------------------------------
	On Error Resume Next
	'if Store_Id=33019 or Store_Id= 32084 then
	'PAYPAL_URL	= "https://www.sandbox.paypal.com/cgi-bin/webscr?cmd=_express-checkout&token="
        ' else
	PAYPAL_URL	= "https://www.paypal.com/cgi-bin/webscr?cmd=_express-checkout&token="
'	end if
'-----------------------------------------------------------------------------
' An express checkout transaction starts with a token, that
' identifies to PayPal your transaction
' In this example, when the script sees a token, the script
' knows that the buyer has already authorized payment through
' paypal.  If no token was found, the action is to send the buyer
' to PayPal to first authorize payment
'--------------------------------------------------------------------------



Auth_Capture  = cbool(request.form("Auth_Capture"))

If Auth_Capture then
    sPaymentType="Sale"
else
    sPaymentType="Authorization"
end if



'---------------------------------------------------------------------------
' The servername and serverport tells PayPal where the buyer
' should be directed back to after authorizing payment.
' In this case, its the local webserver that is running this script
' Using the servername and serverport, the return URL is the first
' portion of the URL that buyers will return to after authorizing payment
'----------------------------------------------------------------------------
		url = GetURL()
		paymentAmount=cstr(formatNumber(request.form("total_amount")))
		currencyCodeType=PayPal_Pro_Currency
		paymentType=sPaymentType 
'---------------------------------------------------------------------------
' The returnURL is the location where buyers return when a
' payment has been succesfully authorized.
' The cancelURL is the location buyers are sent to when they hit the
' cancel button during authorization of payment during the PayPal flow
'---------------------------------------------------------------------------
	currentpath = "http://" & Request.ServerVariables("SERVER_NAME")
  	 '   ExpressCheckout1.ReturnURL = currentpath & "/include/paypalpro/paypal_action.asp?status=Paid"
  	  '  ExpressCheckout1.CancelURL =  currentpath & "/show_big_cart.asp"
		
		
		returnURL	= currentpath & "/include/paypalpro/paypal_action.asp?currencyCodeType=" &  currencyCodeType & _
					"&paymentAmount=" & paymentAmount & _ 
					"&paymentType=" &paymentType 
					cancelURL	= url & "/show_big_cart.asp"

'---------------------------------------------------------------------------
' Construct the parameter string that describes the PayPal payment
' the varialbes were set in the web form, and the resulting string
' is stored in nvpstr
'---------------------------------------------------------------------------
		nvpstr		= "&" & Server.URLEncode("AMT")&"=" & Server.URLEncode(paymentAmount) & _
					"&" &Server.URLEncode("PAYMENTACTION")&"=" & Server.URLEncode(sPaymentType) & _
					"&"&server.URLEncode("RETURNURL")&"=" & Server.URLEncode(returnURL) & _
					"&" &Server.URLEncode("CANCELURL")&"=" &Server.URLEncode(cancelURL) & _ 
					"&"&server.UrlEncode("CURRENCYCODE")&"=" & Server.URLEncode(currencyCodeType)

'--------------------------------------------------------------------------- 
' Make the call to PayPal to set the Express Checkout token
' If the API call succeded, then redirect the buyer to PayPal
' to begin to authorize payment.  If an error occured, show the
' resulting errors
'---------------------------------------------------------------------------
		Set resArray=hash_call("SetExpressCheckout",nvpstr)
		Set SESSION("nvpResArray")=resArray
		ack = UCase(resArray("ACK"))
                sErrorMessage=resArray("L_LONGMESSAGE0")

		If ack="SUCCESS" Then
				' Redirect to paypal.com here
				token = resArray("TOKEN")
				payPalURL = PAYPAL_URL & "&cmd=_express-checkout&token=" & token
				ReDirectURL(payPalURL)
		Else  
				fn_error sErrorMessage

		End If

		
		
else

SandboxAPIUserName =PayPal_Pro_API_username
SandboxAPIPassword = PayPal_pro_Password

Key_Folder = fn_get_sites_folder(Store_Id,"Key") 
APICertificatePath = Key_Folder&"cert_key_pem.txt"
Dim ExpressCheckout1
Set ExpressCheckout1 = Server.CreateObject("iBizPayPalASP.ExpressCheckout")
'ExpressCheckout1.URL = "https://api.sandbox.paypal.com/2.0/" 'Test server URL
ExpressCheckout1.URL = "https://api.paypal.com/2.0/"

ExpressCheckout1.Username = SandboxAPIUserName
ExpressCheckout1.Password = SandboxAPIPassword
ExpressCheckout1.SSLCertStoreType = 4 'PEM
ExpressCheckout1.SSLCertStorePassword = ""
on error resume next
ExpressCheckout1.SSLCertStore = APICertificatePath

ExpressCheckout1.SSLCertSubject = "*"
if err.number<>0 then
    if err.number="-2146808016" or err.number="20272" then
       fn_error "The Paypal Pro certificate could not be found.<BR><HR><BR>Store Owner please ensure your certificate is uploaded properly."
    else
       fn_error err.description &err.number
    end if
end if

if PayPal_Pro_Currency <>"" then
    ExpressCheckout1.CurrencyCode=PayPal_Pro_Currency
end if
Auth_Capture  = cbool(request.form("Auth_Capture"))
if auth_capture then
	ExpressCheckout1.PaymentAction = 0	 'Authorize and Deposit				
else
	ExpressCheckout1.PaymentAction = 1 'Authorize Only
end if

call  StartPayment

Private Sub StartPayment

  	ExpressCheckout1.OrderTotal = cstr(formatNumber(request.form("total_amount")))
	if Email="" then
		Email="unknown@unknown.com"
	end if
	ExpressCheckout1.BuyerEmail =  Email
 	currentpath = "http://" & Request.ServerVariables("SERVER_NAME")
  	ExpressCheckout1.ReturnURL = currentpath & "/include/paypalpro/paypal_action.asp?status=Paid"
  	ExpressCheckout1.CancelURL =  currentpath & "/show_big_cart.asp"
  	on error resume next
    ExpressCheckout1.SetCheckout
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

  	responsetoken = ExpressCheckout1.ResponseToken
	if ExpressCheckout1.Ack = "Success" then
		fn_redirect ("https://www.paypal.com/cgi-bin/webscr?cmd=_express-checkout&token=" + responsetoken)
	else
        fn_error err.description
	end if

End Sub

end if








%>

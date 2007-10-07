<!--#include virtual="store_engine/include/header_noview.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include virtual="store_engine/include/sub.asp"-->
<!-- #include virtual="include/paypalpro/CallerService.asp" -->

<%


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
if PayPal_Pro_Signature<>"" then


 API_USERNAME	= PayPal_Pro_API_username
 API_PASSWORD	= PayPal_pro_Password
 API_SIGNATURE	= PayPal_Pro_Signature

		'---------------------------------------------------------------------------
' At this point, the buyer has completed in authorizing payment
' at PayPal.  The script will now call PayPal with the details
' of the authorization, incuding any shipping information of the
' buyer.  Remember, the authorization is not a completed transaction
' at this state - the buyer still needs an additional step to finalize
' the transaction
'---------------------------------------------------------------------------
		token=Request.Querystring("TOKEN")

		SESSION("token") = Request.Querystring("TOKEN")
		SESSION("currencyCodeType") = Request.Querystring("currencyCodeType")
		SESSION("paymentAmount") = Request.Querystring("paymentAmount")
		SESSION("PaymentType")= Request.Querystring("PaymentType")
		SESSION("PayerID")= Request.Querystring("PayerID")
                paymentAmount= SESSION("paymentAmount")

 '---------------------------------------------------------------------------
 'Build a second API request to PayPal, using the token as the
    'ID to get the details on the payment authorization
'---------------------------------------------------------------------------
		nvpstr="&TOKEN="&Request.Querystring("TOKEN")
		
'---------------------------------------------------------------------------
' Make the API call and store the results in an array.  If the
    'call was a success, show the authorization details, and provide
   ' an action to complete the payment.  If failed, show the error
'---------------------------------------------------------------------------
		Set resArray=hash_call("GetExpressCheckoutDetails",nvpstr)
		ack = UCase(resArray("ACK"))
		Set SESSION("nvpResArray")=resArray
		

		
	If UCase(ack)="SUCCESS" Then
	

	
	
	
	
	
	
	
	first_name = resArray("FIRSTNAME")
last_name =resArray("LASTNAME")
Address1 =resArray("SHIPTOSTREET")
Address2 = resArray("SHIPTOSTREET2")
City = resArray("SHIPTOCITY")
zip = resArray("SHIPTOZIP")
State = resArray("SHIPTOSTATE")
Country = resArray("SHIPTOCOUNTRYNAME")
Email = resArray("EMAIL")
phone = resArray("SHIPTOPHONENUM")
Is_Residential = 1
spam=0

	'STORE INTO DATABASE
		User_ID = EXPRESS_CHECKOUT_CUSTOMER
		Randomize
		Password = "PASS_"&cstr(Int((10000) * Rnd + lowerbound))&month(now)&day(now)&year(now)&cstr(Int((10000) * Rnd + lowerbound))
  
  sql_create_customer = "exec wsp_customer_register "&Store_id&","&Shopper_Id&",'"&User_ID&"','"&Password&"','"&Last_name&"','"&first_name&"','"&Company&"','"&Address1&"','"&Address2&"','"&City&"','"&Zip&"','"&State&"','"&Country&"','"&Phone&"','"&EMail&"','"&FAX&"',"&Spam&","&StartCID&",0,"&Is_Residential&",'';"
  conn_store.execute sql_create_customer

	fn_redirect "http://" & Request.ServerVariables("SERVER_NAME")  &"/before_payment.asp?returnFrom=1001&token="&token
		
	Else
		fn_error UCase(resArray("L_LONGMESSAGE0"))
		          
	
	End If	
	
else

  	SandboxAPIUserName =PayPal_Pro_API_username
			SandboxAPIPassword = PayPal_pro_Password

	      Key_Folder = fn_get_sites_folder(Store_Id,"Key") 
			APICertificatePath = Key_Folder&"cert_key_pem.txt"


			Dim ExpressCheckout1
			Set ExpressCheckout1 = Server.CreateObject("iBizPayPalASP.ExpressCheckout")
		'	ExpressCheckout1.URL = "https://api.sandbox.paypal.com/2.0/" 'Test server URL
	         	ExpressCheckout1.URL = "https://api.paypal.com/2.0/"
			ExpressCheckout1.Username = SandboxAPIUserName
			ExpressCheckout1.Password = SandboxAPIPassword 
			ExpressCheckout1.SSLCertStoreType = 4 'PEM
			ExpressCheckout1.SSLCertStorePassword = ""
			ExpressCheckout1.SSLCertStore = APICertificatePath
			ExpressCheckout1.SSLCertSubject = "*"

			expresscheckout1.getcheckoutdetails(request.querystring("token"))
			token = request.querystring("token")
			session("payerId") =  expresscheckout1.PayerId
			session("token") = token


first_name = checkstringforQ(expresscheckout1.PayerFirstName)
last_name =checkstringforQ(expresscheckout1.PayerLastName)
Address1 = expresscheckout1.PayerStreet1
Address2 = expresscheckout1.PayerStreet2
City = expresscheckout1.PayerCity
zip = expresscheckout1.PayerPostalCode
State = expresscheckout1.PayerState
Country = fn_country_name(expresscheckout1.PayerCountryCode) 'function to get the country name
Email = expresscheckout1.PayerEmail
phone = expresscheckout1.ContactPhone
Is_Residential = 1
spam=0

	'STORE INTO DATABASE
		User_ID = EXPRESS_CHECKOUT_CUSTOMER
		Randomize
		Password = "PASS_"&cstr(Int((10000) * Rnd + lowerbound))&month(now)&day(now)&year(now)&cstr(Int((10000) * Rnd + lowerbound))
  
  sql_create_customer = "exec wsp_customer_register "&Store_id&","&Shopper_Id&",'"&User_ID&"','"&Password&"','"&Last_name&"','"&first_name&"','"&Company&"','"&Address1&"','"&Address2&"','"&City&"','"&Zip&"','"&State&"','"&Country&"','"&Phone&"','"&EMail&"','"&FAX&"',"&Spam&","&StartCID&",0,"&Is_Residential&",'';"
  conn_store.execute sql_create_customer


	response.redirect "http://" & Request.ServerVariables("SERVER_NAME")  &"/before_payment.asp?returnFrom=1001&token="&token
	
end if

%><!--#include virtual="store_engine/include/footer.asp"-->

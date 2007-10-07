<!--#include virtual="include/header_noview.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include file="CallerService.asp" -->


<%
set rs_Store =  server.createobject("ADODB.Recordset")

srv = request.servervariables("server_name")


 API_USERNAME	= PayPal_Pro_API_username
 API_PASSWORD	= PayPal_pro_Password
 API_SIGNATURE	= PayPal_Pro_Signature

currencyCodeType=PayPal_Pro_Currency

if PayPal_Pro_Signature<>"" then

currentpath = "http://" & Request.ServerVariables("SERVER_NAME")

'PayPal_Pro_notifyurl=  currentpath & "/include/paypalpro/paypal_ipn.asp"
PayPal_Pro_notifyurl =  currentpath & "/include/paypalpro/ipn2.asp"


	           token= session("token")
			   OrderTotal = request.form("total_amount")
		       PayerId = session("payerId")
			If Auth_Capture then
               sPaymentType="Sale"
            else
               sPaymentType="Authorization"
            end if
			
                        on error resume next
			order =  request.form("Order")
			nvpstr			=	"&" & Server.URLEncode("TOKEN") & "=" & Server.URLEncode(token)  & "&" &_
						Server.URLEncode("PAYERID")&"=" &Server.URLEncode(payerID) & "&" &_
						Server.URLEncode("PAYMENTACTION")&"=" & Server.URLEncode(sPaymentType) & "&" &_
						Server.URLEncode("AMT") &"=" & Server.URLEncode(OrderTotal )&"&"&_
						Server.URLEncode("BUTTONSOURCE") &"=" & Server.URLEncode("EasyStoreCreator_Cart_EC_US")&"&"&_
						Server.URLEncode("CURRENCYCODE")& "=" &Server.URLEncode(PayPal_Pro_Currency)&"&"&_
						Server.URLEncode("NOTIFYURL")& "=" &Server.URLEncode(PayPal_Pro_notifyurl)

'-------------------------------------------------------------------------------------------
' Make the call to PayPal to finalize payment
' If an error occured, show the resulting errors
'-------------------------------------------------------------------------------------------
	Set resArray=hash_call("DoExpressCheckoutPayment",nvpstr)
	ack = UCase(resArray("ACK"))
'-------------------------------------------------------------------------------------------
' Display the API request and API response back to the browser using APIError.asp.
' If the response from PayPal was a success, display the response parameters
' If the response was an error, display the errors received
'-------------------------------------------------------------------------------------------	
	If ack=UCase("Success") Then
		message			= "Thank you for your payment!"
		 Verified_Ref = resArray("TRANSACTIONID")
		 mypayment_type= resArray("PAYMENTTYPE")
	Else       
		 fn_error UCase(resArray("L_LONGMESSAGE0"))
	End If
	
	if Auth_Capture then
        	trans_type = 1
        else
        	trans_type = 0
        end if









                if err.number<> 0 then
			fn_error "Error occurred while processing your payment please try again later<BR><BR>"&err.description
		else
			'sql_update = "update store_purchases set AuthNumber='"&session("payerId")&"',Processor_id=36,transaction_type="&trans_type& ",Verified_Ref = '"&Verified_Ref&"' where (oid = "& order &" or masteroid="& order &") AND Store_id ="&Store_id
			if mypayment_type="instant" then
                        sql_update = "update store_purchases set Verified =1,AuthNumber='"&session("payerId")&"',Processor_id=36,transaction_type="&trans_type& ",Verified_Ref = '"&Verified_Ref&"' where (oid = "& order &" or masteroid="& order &") AND Store_id ="&Store_id
                        else
                        sql_update = "update store_purchases set Verified =0,AuthNumber='"&session("payerId")&"',Processor_id=36,transaction_type="&trans_type& ",Verified_Ref = '"&Verified_Ref&"' where (oid = "& order &" or masteroid="& order &") AND Store_id ="&Store_id
                        end if
		        conn_store.execute sql_update
                        fn_redirect "http://"&srv&"/include/check_out_payment_action.asp?Return_From=PayPalExpress&Shopper_Id="&Shopper_Id
		end if
		
		else
		
		  	srv = request.servervariables("server_name")



	sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
    
		 rs_Store.open sql_real_time,conn_store,1,1
		Do While Not Rs_Store.EOF
					select case Rs_store("Property")
						case "PayPal_Pro_Api_Username"
							PayPal_Pro_API_username = decrypt(Rs_store("Value"))
						case "PayPal_Pro_Password"
							PayPal_pro_Password = decrypt(Rs_store("Value"))
						case "PayPal_Pro_Currency"
      				     PayPal_Pro_Currency = decrypt(Rs_store("Value"))
                  case "Certificate_file_name"
							Certificate_file_name = decrypt(rs_store("Value"))
					end select
					Rs_store.MoveNext
			Loop
			Rs_store.Close

	      Key_Folder = fn_get_sites_folder(Store_Id,"Key") 
         Set ExpressCheckout1 = Server.CreateObject("iBizPayPalASP.ExpressCheckout")
			ExpressCheckout1.URL = "https://api.paypal.com/2.0/"
		'	ExpressCheckout1.URL = "https://api.sandbox.paypal.com/2.0/" 'Test server URL
			ExpressCheckout1.Username = PayPal_Pro_API_username
			ExpressCheckout1.Password = PayPal_pro_Password
			ExpressCheckout1.SSLCertStoreType = 4  'PEM
			ExpressCheckout1.SSLCertStorePassword = ""
			ExpressCheckout1.SSLCertStore = Key_Folder&"cert_key_pem.txt"
			ExpressCheckout1.SSLCertSubject = "*"

			ExpressCheckout1.token= session("token")
			ExpressCheckout1.OrderTotal = request.form("total_amount")
			ExpressCheckout1.PayerId = session("payerId")
			if request.form("Auth_Capture") = "True" then
				ExpressCheckout1.PaymentAction = 0 'aSale
			else
				ExpressCheckout1.PaymentAction = 1 'aAuthorization 
			end if
			
                        on error resume next
			order =  request.form("Order")
			ExpressCheckout1.checkoutPayment




                if err.number<> 0 then
			response.redirect "http://"&srv&"/error.asp?Message_Id=101&Message_Add="&Server.urlencode("Error occurred while processing your payment please try again later<BR><BR>"&err.description)
		else
			sql_update = "update store_purchases set Verified =1,AuthNumber='"&session("payerId")&"',Processor_id=36 where (oid = "& order &" or masteroid="& order &") AND Store_id ="&Store_id
		        conn_store.execute sql_update
                        response.redirect "http://"&srv&"/include/check_out_payment_action.asp?Return_From=PayPalExpress&Shopper_Id="&sShopperId
		end if
		

		
		
		
		
		
		
		end if
%>

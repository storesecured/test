

<%

sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst	 
Do While Not Rs_Store.EOF 
	select case Rs_store("Property")
		case "x_Login"
			x_Login = decrypt(Rs_store("Value"))
		case "x_tran_key"
			x_tran_key = decrypt(Rs_store("Value"))
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





'apiUsername="mblack6789_api1.yahoo.com"
'apiPassword="QXZKVJQJW37LL4RW"
'apiSignature="AcOdyTTj-hhKLwyaf8su9uxekBKlAUVjyRq92f4Uhb23Abbz00pxHUS1"
'currencyCode="USD"

'apiUsername	= "bassel_gado_api1.yahoo.com"
'apiPassword	= "EBJSQG3FDTW8KEAH"
'apiSignature	= "AGrpXtwFJ7TMeqHT6t2t-5rWd7u1Afcz3IcSFAbsB4nALeUtir-cH6DV"


 if PayPal_Pro_Signature<>"" then
sSecurityString = "USER="&server.URLEncode(PayPal_Pro_API_username)&_
    "&PWD="&server.URLEncode(PayPal_pro_Password)&_
    "&SIGNATURE="&server.URLEncode(PayPal_Pro_Signature)&_
    "&CURRENCYCODE="&server.URLEncode(PayPal_Pro_Currency)&_
    "&VERSION="&server.urlencode(2.3)

rs_store.open "Select verified_ref,grand_total from Store_Purchases WITH (NOLOCK) where oid = "&Oid&" and Store_id ="&Store_id,conn_store,1,1
If not rs_Store.eof then
    Verified_Ref=rs_Store("Verified_Ref")
	Grand_Total=rs_Store("Grand_Total")
End If
rs_Store.Close

If sType = "Capture" Then
	sTypeString="&METHOD=DoCapture&AUTHORIZATIONID="&Verified_Ref&"&AMT="&GGrand_Total&"&COMPLETETYPE=Complete"
elseif sType = "Credit" then
	if cdbl(GGrand_Total)=cdbl(Grand_Total) then
	    sTypeString="&METHOD=RefundTransaction&TRANSACTIONID="&Verified_Ref&"&REFUNDTYPE=Full&NOTE="
	else
	    sTypeString="&METHOD=RefundTransaction&TRANSACTIONID="&Verified_Ref&"&REFUNDTYPE=Partial&AMT="&GGrand_Total&"&NOTE="
        trans_type=4
    end if
else
    sTypeString="&METHOD=DoVoid&AUTHORIZATIONID="&Verified_Ref
end if

Set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
'objHttp.open "POST", "https://api-3t.sandbox.paypal.com/nvp", False
objHttp.open "POST", "https://api-3t.paypal.com/nvp", False
WinHttpRequestOption_SslErrorIgnoreFlags=4
objHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
objHttp.Send sSecurityString&sTypeString
call sub_write_log("Request="&sSecurityString&sTypeString)

'Set SESSION("nvpReqArray")= deformatNVP(nvpStr)
sResponse=objHttp.responseText

Set  nvpResponseCollection =deformatNVP(sResponse)
call sub_write_log("Response="&sResponse)
sErrorMessage=nvpResponseCollection("L_LONGMESSAGE0")
sAck=nvpResponseCollection("ACK")
Set  hash_call =nvpResponseCollection
Set objHttp = Nothing

if sErrorMessage<>"" and sAck="Failure" then
    response.Redirect "error.asp?Message_Id=102&Message_Add="&server.URLEncode(sErrorMessage)
end if

'----------------------------------------------------------------------------------
' Purpose: It will convert nvp string to Collection object.
' Inputs:  NVP string.
' Returns: NVP Collection object deformated from NVP string.
'----------------------------------------------------------------------------------

else
 response.Redirect "error.asp?Message_Id=101&message_add="&server.URLEncode("You must upgrade to the new release for paypalpro to be able to use this feature")
end if
%>

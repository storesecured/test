<%

sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst
Do While Not Rs_Store.EOF
   select case Rs_store("Property")
		case "Cybersource_ID"
			Cybersource_ID = decrypt(Rs_store("Value"))
		case "Cybersource_Currency"
			Cybersource_Currency =decrypt(rs_Store("Value"))
	end select
	Rs_store.MoveNext
Loop
Rs_store.Close

Key_Folder = fn_get_sites_folder(Store_Id,"Key")  

Dim oMerchantConfig
set oMerchantConfig = Server.CreateObject( "CyberSourceWS.MerchantConfig" )

oMerchantConfig.MerchantID = Cybersource_Id
oMerchantConfig.KeysDirectory = Key_Folder
oMerchantConfig.SendToProduction = "1"
oMerchantConfig.TargetAPIVersion = "1.10"


Dim oRequest
set oRequest = Server.CreateObject( "CyberSourceWS.Hashtable" )

oRequest( "merchantReferenceCode" ) = Oid
oRequest.Value( "ccAuthService_run" ) = "true"
If Auth_Capture then
	oRequest.Value( "ccCaptureService_run" ) = "true"
else
	oRequest.Value( "ccCaptureService_run" ) = "false"
end if


oRequest.Value( "billTo_firstName" ) = first_name
oRequest.Value( "billTo_lastName" ) = last_name
oRequest( "billTo_street1" ) = address1
oRequest( "billTo_city" ) =city
oRequest( "billTo_state" ) = state
oRequest( "billTo_postalCode" ) = zip
oRequest( "billTo_country" ) =country
oRequest( "billTo_email" ) = email
oRequest( "billTo_phoneNumber" ) = phone
oRequest( "card_accountNumber" ) = Request.Form("CardNumber")
oRequest( "card_cvIndicator") = 1
oRequest( "card_cvNumber" ) = Request.Form("CardCode")
oRequest( "card_expirationMonth" ) = Request.Form("mm")
oRequest( "card_expirationYear" ) = Request.Form("yy")
oRequest( "purchaseTotals_currency" ) = Cybersource_Currency
oRequest( "purchaseTotals_grandTotalAmount" ) =  formatnumber(GGrand_Total,2,0,0,0)


dim oClient
set oClient = Server.CreateObject( "CyberSourceWS.Client" )
dim varReply, nStatus, strErrorInfo
nStatus = oClient.RunTransaction( oMerchantConfig, Nothing, Nothing, oRequest, varReply, strErrorInfo )

if nStatus=0 then


	dim decision
	decision = UCase( varReply( "decision" ) )
	if decision = "ACCEPT" then
		AuthNumber = varReply("ccAuthReply_authorizationCode")
		avs_result	= varReply("	ccAuthReply_avsCode")
		Verified_Ref =  varReply("requestID")
		card_verif = varReply("ccAuthReply_cvCode")

	elseif decision = "REVIEW" then

	else ' REJECT or ERROR
			dim reasonCode
				reasonCode = varReply( "reasonCode" )
				
				'if decision = "REJECT" then
					if reasonCode=101 then
						reasonCode="Missing field="&varReply( "missingField_0" )
					elseif reasonCode=102 then
						reasonCode="Invalid field="&varReply( "invalidField_0" )
					elseif reasonCode=104 then
						reasonCode="Duplicate transaction"
               elseif reasonCode=150 or reasonCode=236 then
						reasonCode="Payment gateway failure please wait and try again later"
               elseif reasonCode=151 or reasonCode=152 or reasonCode=250 then
						reasonCode="Payment gateway timeout"
               elseif reasonCode=201 then
						reasonCode="Transaction cannot be processed online please call for approval"
               elseif reasonCode=202 then
						reasonCode="Credit card has expired"
               elseif reasonCode=203 or reasonCode=233 then
						reasonCode="General decline"
               elseif reasonCode=204 then
						reasonCode="Insufficient funds"
               elseif reasonCode=205 then
						reasonCode="Card reported lost or stolen"
               elseif reasonCode=207 then
						reasonCode="Issuing bank unavailable"
               elseif reasonCode=208 then
						reasonCode="Inactive card"
					elseif reasonCode=210 then
						reasonCode="Card has reached its limit"
					elseif reasonCode=211 then
						reasonCode="Invalid card code"
					elseif reasonCode=221 then
						reasonCode="Card listed in negative file"
					elseif reasonCode=231 then
						reasonCode="Invalid card number"
					elseif reasonCode=232 then
						reasonCode="Card type not accepted"
					elseif reasonCode=234 then
						reasonCode="Error in cybersource setup"
					elseif reasonCode=520 then
						reasonCode="Declined due to fraud scoring system"
               end if
               fn_purchase_decline oid,"The transaction was rejected by the payment processor:<BR>"&reasonCode
	end if
else
    fn_purchase_decline oid,"The transaction was rejected by the payment processor:<BR>"&strErrorInfo
end if

if Auth_Capture then
	trans_type = 1
else
	trans_type = 0
end if

%>


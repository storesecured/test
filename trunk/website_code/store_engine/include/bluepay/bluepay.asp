<!--#include file="bluepay_include.asp"-->

<%



sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst
Do While Not Rs_Store.EOF 
	select case Rs_store("Property")
		case "Account_ID"
			Account_ID = decrypt(rs_store("Value"))	
		 case "Secret_Key"
			 Secret_Key =decrypt(rs_Store("Value"))
	end select
	Rs_store.MoveNext
Loop
Rs_store.Close

' First we have to set up the Bluepay transaction:

sURL="http://www.test.com/include/bluepay/bluepayresp.asp"

bp_setup Account_ID,Secret_Key,"LIVE", sURL, sURL, sURL, "https://secure.bluepay.com/interfaces/bp10emu"

' Set the customer required information -- this is necessary for a SALE or AUTH.
sPayment_Method=request.form("Payment_Method")

if sPayment_Method="eCheck" then
        
        bp_set_cust_required_ach Request.Form("BankABA"), Request.Form("BankAccount"), "", _
                     request.form("CardName") , address1, city, state, zip
else
        bp_set_cust_required CardNumber, CardCode, CardExpiration, _
                     request.form("CardName") , address1, city, state, zip
end if

' Set Optional Information:		   
bp_set_orderid(oid)
bp_set_addr2(ShipAddress2)
'bp_set_comment(request.form("CC_comment"))
bp_set_phone(phone)
bp_set_email(email)

If Auth_Capture or sPayment_Method="eCheck" then
	bp_easy_sale(GGrand_Total)
else
        bp_easy_auth(GGrand_Total)
end if

bp_process()
sUrlArray=split(BP_RESPONSE,"?")
For Each querystring_group In split(sUrlArray(1),"&")
    sQuerystringParts=split(querystring_group,"=")
    sQueryName=sQuerystringParts(0)
    sQueryValue=sQuerystringParts(1)
    select case sQueryName
           case "Result"
                Result = sQueryValue
           case "RRNO"
                RRNO = sQueryValue
           case "AVS"
                AVS = sQueryValue
           case "CVV2"
                CVV2 = sQueryValue
           case "AUTHCODE"
                AUTH_CODE = sQueryValue
           case "MESSAGE"
                MESSAGE = unencode(sQueryValue)
           case "MISSING"
                MISSING = unencode(sQueryValue)
  end select
next

if Result="APPROVED" then
    Verified_Ref=RRNO
    AuthNumber= Auth_Code
    avs_result = Avs
    card_verif = CVV2
else
        if Result = "DECLINED" then
                ErrMsg = MESSAGE
                sReason = "The transaction was rejected by the payment processor:<BR>"& ErrMsg
        elseif Result = "ERROR" then
                ErrMsg = MESSAGE
                sReason = "The transaction was rejected by the payment processor:<BR>"& ErrMsg
        elseif Result = "MISSING" then
                sMissing = Missing
                sResponse = "Transaction cannot be processed without a valid:<BR>"&sMissing
        else
                sReason = "The transaction was rejected by the payment processor:<BR>An unknown error occurred."
        end if
        fn_purchase_decline oid,sReason
end if

if Auth_Capture then
	trans_type = 1
else
	trans_type = 0
end if


%>

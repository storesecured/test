<%

 sql_real_time = "Select * from store_external where Store_id = "&Store_id
	 rs_Store.open sql_real_time,conn_store,1,1

preCharge_MerchantID = rs_Store("Precharge_MerchantID")
preCharge_Security1 = rs_Store("Precharge_Security1")
preCharge_Security2 = rs_Store("Precharge_Security2")

rs_store.close 

rs_store.open "select * from store_customers where record_type=0 and cid="&cid, conn_store, 1, 1

If not rs_store.eof then 
	Ecom_BillTo_Postal_Name_First =rs_store("first_name")
	Ecom_BillTo_Postal_Name_Last =rs_store("Last_name")
	Ecom_BillTo_Postal_PostalCode =rs_store("Zip")
	Ecom_BillTo_Postal_CountryCode =rs_store("Country")
	Ecom_BillTo_Telecom_Phone_Number = rs_store("Phone")
	Ecom_BillTo_Online_Email =rs_store("Email")
	Ecom_BillTo_Postal_Street_Line1 = rs_store("Address1") 
	Ecom_BillTo_Postal_Street_Line2 = rs_store("Address2") 
	Ecom_BillTo_Postal_City = rs_store("City")
	Ecom_BillTo_Postal_StateProv = rs_store("State")
	Ecom_BillTo_Online_IP =Request.ServerVariables("REMOTE_ADDR") 
	Invoice_Number = oid 
	Client_ID =cid 
	Ecom_Transaction_Amount = formatnumber(GGrand_Total,2) 
End If
	rs_store.close 

Set xObj = CreateObject("SOFTWING.ASPtear")
Post_String = ""
Post_String = Post_String &"Merchant_ID="&Server.UrlEncode(preCharge_MerchantID)
Post_String = Post_String &"&Security1="& Server.UrlEncode(preCharge_Security1)
Post_String = Post_String &"&Security2="& Server.UrlEncode(preCharge_Security2)
Post_String = Post_String &"&Ecom_BillTo_Postal_Name_First="&Server.UrlEncode(Ecom_BillTo_Postal_Name_First)
Post_String = Post_String &"&Ecom_BillTo_Postal_Name_Last=" & Server.UrlEncode(Ecom_BillTo_Postal_Name_Last)
Post_String = Post_String &"&Ecom_BillTo_Postal_PostalCode="&Server.UrlEncode(Ecom_BillTo_Postal_PostalCode)
Post_String = Post_String &"&Ecom_BillTo_Postal_CountryCode="& Server.UrlEncode(Ecom_BillTo_Postal_CountryCode)
Post_String = Post_String &"&Ecom_BillTo_Telecom_Phone_Number="& Server.UrlEncode(Ecom_BillTo_Telecom_Phone_Number)
Post_String = Post_String &"&Ecom_BillTo_Online_Email="& Server.UrlEncode(Ecom_BillTo_Online_Email)
Post_String = Post_String &"&Ecom_BillTo_Postal_Street_Line1="& Server.UrlEncode(Ecom_BillTo_Postal_Street_Line1)
Post_String = Post_String &"&Ecom_BillTo_Postal_Street_Line2="& Server.UrlEncode(Ecom_BillTo_Postal_Street_Line2)
Post_String = Post_String &"&Ecom_BillTo_Postal_City="& Server.UrlEncode(Ecom_BillTo_Postal_City)
Post_String = Post_String &"&Ecom_BillTo_Postal_StateProv="& Server.UrlEncode(Ecom_BillTo_Postal_StateProv)
Post_String = Post_String &"&Ecom_BillTo_Online_IP="& Server.UrlEncode(Ecom_BillTo_Online_IP)
Post_String = Post_String &"&Invoice_Number="& Server.UrlEncode(Invoice_Number)
Post_String = Post_String &"&Client_ID="& Server.UrlEncode(Client_ID)
Post_String = Post_String &"&Ecom_Transaction_Amount="& Server.UrlEncode(Ecom_Transaction_Amount)
Post_String = Post_String &"&Ecom_Payment_Card_Number="& Server.UrlEncode(CardNumber)
Post_String = Post_String &"&Ecom_Payment_Card_ExpDate_Month="& Server.UrlEncode(Request.Form("mm"))
Post_String = Post_String &"&Ecom_Payment_Card_ExpDate_Year="& Server.UrlEncode( "20"& Request.Form("yy"))
'Post_String = Post_String &"&test="& Server.UrlEncode("1")
'response.write Post_String
'response.end

strResult = xObj.Retrieve("https://api.precharge.net/charge", 1, Post_String, "", "")


 Returned_Var_Array = Split(strResult,",")
 response_code = replace(Returned_Var_Array(0),"response=","")
 response_score = replace(Returned_Var_Array(1),"score=","")
 response_transaction = replace(Returned_Var_Array(2),"transaction=","")

'Response.write "    response_code - " &response_code
'Response.write "    response_score - " &response_score
'Response.write "    response_transaction - " &response_transaction

'Response Code 1 - Transaction Completed
If response_code = "1" then
	response_score = replace(Returned_Var_Array(1),"score=","")
	sql_pre = "update store_purchases set Precharge_Response=1, Precharge_Score = "&response_score&", Precharge_Transaction = '"&response_transaction&"'  where oid = "&oid&" and store_id="&store_id
	 rs_Store.open sql_pre,conn_store,1,1
	 rs_Store.close
End If

'Response Code 2 - Transaction Rejected
If response_code = "2" then
	sql_pre = "update store_purchases set Precharge_Response=2, Precharge_Score = "&response_score&", Precharge_Transaction = '"&response_transaction&"'  where oid = "&oid&" and store_id="&store_id
	rs_Store.open sql_pre,conn_store,1,1
	rs_Store.close
	StrMsg = "Transaction declined by Precharge.   Fraud Score="&response_score
	fn_purchase_decline oid,"The following (Precharge) error occured.<BR>"&StrMsg
End If

'Response Code 3 - Transaction Error
If response_code = "3" then

	response_error = replace(Returned_Var_Array(1),"error=","")
	select case response_error
		case "101"
			StrMsg = "Invalid Request Method"
		case "102"
			StrMsg = "Invalid Request URL"
		case "103"
			StrMsg = "Invalid Security Code(s)"
		case "104"
			StrMsg = "Merchant Status Not Verified"
		case "105"
			StrMsg = "Merchant Feed Disabled"
		case "106"
			StrMsg = "Invalid Request Type"
		case "107"
			StrMsg = "Missing IP Address"
		case "108"
			StrMsg = "Invalid IP Address Syntax"
		case "109"
			StrMsg = "Missing First Name"
		case "110"
			StrMsg = "Invalid First Name"
		case "111"
			StrMsg = "Missing Last Name"
		case "112"
			StrMsg = "Invalid Last Name"
		case "113"
			StrMsg = "Invalid Address 1"
		case "114"
			StrMsg = "Invalid Address 2"
		case "115"
			StrMsg = "Invalid City"
		case "116"
			StrMsg = "Invalid State"
		case "117"
			StrMsg = "Invalid Country"
		case "118"
			StrMsg = "Missing Postal Code"
		case "119"
			StrMsg = "Invalid Postal Code"
		case "120"
			StrMsg = "Missing Phone Number"
		case "121"
			StrMsg = "Invalid Phone Number"
		case "122"
			StrMsg = "Missing Expiration Month"
		case "123"
			StrMsg = "Invalid Expiration Month"
		case "124"
			StrMsg = "Missing Expiration Year"
		case "125"
			StrMsg = "Invalid Expiration Year"
		case "126"
			StrMsg = "Expired Credit Card"
		case "127"
			StrMsg = "Missing Credit Card Number"
		case "128"
			StrMsg = "Invalid Credit Card Number"
		case "129"
			StrMsg = "Missing Email Address"
		case "130"
			StrMsg = "Invalid Email Syntax"
		case "131"
			StrMsg = "Duplicate Transaction"
		case "998"
			StrMsg = "Unknown Error"
		case "999"
			StrMsg = "Service Unavailable"
	end select

	sql_pre = "update store_purchases set Precharge_Response=3, Precharge_Transaction = '"&StrMsg&"'  where oid = "&oid&" and store_id="&store_id
	rs_Store.open sql_pre,conn_store,1,1
	rs_Store.close

	StrMsg = "Transaction Error : " & StrMsg
        fn_purchase_decline oid,"The following error occured.<BR>"&StrMsg
End If
%>

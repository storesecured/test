<%
if Request.Querystring("Shopper_id")<>"" then
    Shopper_ID = Request.Querystring("Shopper_id")
    Session("Shopper_ID")=Shopper_ID
end if
Payment_status = Request.Form("payment_status")

%>
<!--#include virtual="include/check_out_payment_action.asp"-->
<%

function fn_updateThirdParty()
	on error resume next
    fn_print_debug "in update third party"
    ' read post from PayPal system and add 'cmd'
    str = Request.Form
    fn_print_debug "str="&str

    ServerName = Request.ServerVariables("server_name")

    OrderID = Request.Form("item_number")
    Txn_id = Request.Form("txn_id")
    Payment_status = Request.Form("payment_status")

    ' post back to PayPal system to validate
    str = str & "&cmd=_notify-validate"
    set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
    objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false
    objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    objHttp.Send str

    fn_print_debug "response="&objHttp.responseText
    fn_print_debug "stats="&objHttp.status
    fn_print_debug "str="&request.form

    ' assign posted variables to local variables
    ' note: additional IPN variables also available -- see IPN documentation
    Item_name = Request.Form("item_name")
    Receiver_email = Request.Form("receiver_email")
    OrderID = Request.Form("item_number")
    Invoice = Request.Form("invoice")
    Payment_status = Request.Form("payment_status")
    payment_type = Request.Form("payment_type")
    pending_reason = Request.Form("pending_reason")
    Payment_gross = Request.Form("mc_gross")
    Txn_id = Request.Form("txn_id")
    Payer_email = Request.Form("payer_email")
    Shopper_Id = request.querystring("Shopper_Id")
    Session("Shopper_Id")=Shopper_Id
    isverif = false
    address_street = checkstringforQ(Request.Form("address_street"))
    address_city = checkstringforQ(Request.Form("address_city"))
    address_state = checkstringforQ(Request.Form("address_state"))
    address_zip = checkstringforQ(Request.Form("address_zip"))
    address_country = checkstringforQ(Request.Form("address_country"))

    ' Check notification validation
    
    

    Set rs_store = Server.CreateObject("ADODB.Recordset")
    if (objHttp.status <> 200 ) then
        fn_print_debug "in paypal error"

    elseif (objHttp.responseText = "VERIFIED") then
        fn_print_debug "in paypal verifed"
	    if (Payment_status="Pending") then
	    	    sReason = request.form("pending_reason")
                AuthNumber = "Pending "&sReason
                Verified=0
            elseif (Payment_status="Refunded") then
        	  	     response.end
				AuthNumber = "Refunded"
            	Verified=0
		  elseif (Payment_status="Denied") then
        	  	AuthNumber = "Denied"
            	Verified=0
            elseif (Payment_status="Reversed") then
        	  	AuthNumber = "Reversed"
            	Verified=0
		  elseif (Payment_status="Completed") then
                AuthNumber = Txn_id
                Verified=1
            elseif (Payment_status="") then
            	 'response is probably from subscription, not needed
                response.end
            else
		  	AuthNumber = Payment_status
            	Verified=0
		  end if
		   sql_insert ="insert into Store_Ipn(store_id,OID,ipn_post,payment_type,payment_status,pending_reason) values(" & Store_id & "," & OrderID & ",'" & str & "','" & payment_type & "','" & Payment_status & "','" & pending_reason & "')"
                   conn_store.execute sql_insert
            sql_update = "exec wsp_purchase_verify "&Store_id&","&OrderID&",4,"&verified&",'"&AuthNumber&"','"&AuthNumber&"','','',0;"
            session("sql")=sql_update
            conn_store.execute(sql_update)
		  if address_street<>"" then
		  	'sql_update = "update store_purchases set shipaddress1='"&address_street&"',shipcity='"&address_city&"',shipstate='"&address_state&"',shipcountry='"&address_country&"',shipzip='"&address_zip&"' where store_id="&store_id&" and (masteroid="&OrderID&" or oid="&OrderID&")"
            	'session("sql")=sql_update
		 	'conn_store.execute(sql_update)
		  end if
        ' check that Payment_status=Completed
        ' check that Txn_id has not been previously processed
        ' check that Receiver_email is an email address in your PayPal account
        ' process payment
    end if

    set objHttp = nothing
	fn_updateThirdParty=1
end function

%>


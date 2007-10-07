<%
if Request.Querystring("Shopper_id")<>"" then
    Shopper_ID = Request.Querystring("Shopper_id")
    Session("Shopper_ID")=Shopper_ID
end if


%>
<!--#include virtual="include/header_noview.asp"-->
<%

	on error resume next
    fn_print_debug "in update third party"
    ' read post from PayPal system and add 'cmd'
    str = Request.Form
    sub_write_log str
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
    Payment_gross = Request.Form("mc_gross")
    Txn_id = Request.Form("txn_id")
    Payer_email = Request.Form("payer_email")
    Shopper_Id = request.querystring("Shopper_Id")
    Session("Shopper_Id")=Shopper_Id
    isverif = false
    address_name = Request.Form("address_name")
    sArrayName = split(address_name," ")
    address_first = checkstringforQ(sArrayName(0))
    address_last = checkstringforQ(sArrayName(1))
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
                AuthNumber = "Pending"
                Verified=0
            elseif (Payment_status="Refunded") then
        	  	AuthNumber = "Refunded"
            	Verified=0
		  elseif (Payment_status="Denied") then
        	  	AuthNumber = "Denied"
            	Verified=0
		  elseif (Payment_status="Completed") then
                AuthNumber = Txn_id
                Verified=1		  	
		  end if
            sql_update = "exec wsp_purchase_verify "&Store_id&","&oid&",4,"&verified&",'"&AuthNumber&"','"&AuthNumber&"','','',0;"
            session("sql")=sql_update
		  conn_store.execute(sql_update)
		  if address_street<>"" then
		  	sql_update = "update store_purchases set ShipFirstname='"&address_first&"', ShipLastName='"&address_last&"', shipaddress1='"&address_street&"',shipcity='"&address_city&"',shipstate='"&address_state&"',shipcountry='"&address_country&"',shipzip='"&address_zip&"' where store_id="&store_id&" and (masteroid="&oid&" or oid="&oid&")"
            	session("sql")=sql_update
		 	conn_store.execute(sql_update)
		  end if
		  

        ' check that Payment_status=Completed
        ' check that Txn_id has not been previously processed
        ' check that Receiver_email is an email address in your PayPal account
        ' process payment
    end if
    set objHttp = nothing



%>
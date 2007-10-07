<!--#include file="header_noview.asp"-->
<!--#include virtual="common/CreditCardFraudDetection.class"-->
<!--#include file="emails.asp"-->
<!--#include virtual="common/cc_validation.asp"-->
<!--#include file="Send_Invoice_By_Mail.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include file="sub.asp"-->
<!--#include file="check_out_payment_include.asp"-->
<%

Server.ScriptTimeout = 360
on error goto 0

if request.Form("Local_Post")="1" then
    ThirdPartyPost=0
else
    ThirdPartyPost=1
end if

Return_From = fn_get_querystring("Return_From")
Processor = Real_Time_Processor
sUpdateRealtime=0
trans_type=0
iDontRedirect=0

if Return_From="Admin" then
   sql_select = "update store_shoppingcart set cart_processed=0 where shopper_id="&Shopper_Id&";"
   fn_print_debug sql_select
   conn_store.Execute sql_select
end if

CardName = checkStringForQ(Request.Form("CardName"))
CardNumber = Request.Form("CardNumber")
if CardNumber = "" then
    CardNumber = decrypt(Request.Form("Card_Ending"))
end if
CardExpiration = Request.Form("mm")&"/"&Request.Form("yy")
CardCode=Request.Form("CardCode")
CardNumberInsert = CardNumber
IssueNumber = Request.Form("Issue_No")
IssueStart = Request.Form("issue_date")
if IssueStart="" then
	IssueStart="Null"
else
	IssueStart="'"&IssueStart&"'"
end if
if IssueNumber="" then
	IssueNumber="Null"
end if


Budget_Left=0
Reward_Left=0
fship_to=0
Ship_Location_Id=0

iProcessedPayment=0
if fn_get_querystring("Return_From")="Admin" then
	iProcessedPayment=1
	fn_print_debug "iProcessedPayment="&iProcessedPayment
elseif ThirdPartyPost=1 then
	on error resume next
    iProcessedPayment = fn_updateThirdParty()
    on error goto 0
end if


sql_select = "exec wsp_purchase_select "&Store_id&","&Oid&";"

fn_print_debug sql_select
session("sql")=sql_select
rs_store.open sql_select,conn_store,1,1
If not rs_Store.eof then
    lock_timestamp = Rs_store("lock_timestamp")
    purchase_completed = Rs_store("purchase_completed")
    payment_method = Rs_store("payment_method")
    Verified = Rs_store("Verified")
    if payment_method = "" then
        fn_purchase_decline oid,"You have not chosen a valid payment method."
    end if
    if purchase_completed <> 0 then     
        fn_redirect Switch_Name&"Recipiet.asp"
    end if
    if ThirdPartyPost<>1 then
    	   fn_purchase_lock oid,lock_timestamp,purchase_completed,Verified
    end if
    
    first_name=rs_Store("First_name")
    last_name=rs_Store("Last_name")
    address1=rs_Store("Address1")
    address2=rs_Store("Address2")
    city=rs_Store("City")
    state=rs_Store("State")
    zip=rs_Store("zip")
    country=rs_Store("Country")
    phone=rs_Store("Phone")
    fax=rs_Store("fax")
    email=rs_Store("email")
    ccid = rs_Store("CCID")
    cid = rs_Store("cid")
    Company = rs_Store("Company")
    Tax_Exempt = rs_Store("Tax_Exempt")
    Shipping_Method_Price = Rs_store("Shipping_Method_Price")
    Tax = Rs_store("Tax")
    ShipFirstname=rs_Store("ShipFirstname")
    ShipLastname=rs_Store("ShipLastname")
    ShipAddress1=rs_Store("ShipAddress1")
    ShipAddress2=rs_Store("ShipAddress2")
    ShipCity=rs_Store("ShipCity")
    ShipState=rs_Store("ShipState")
    Shipzip=rs_Store("Shipzip")
    ShipCountry=rs_Store("ShipCountry")
    ShipPhone=rs_Store("ShipPhone")
    ShipFax=rs_Store("ShipFax")
    ShipEmail=rs_Store("ShipEmail")
    ShipCompany=rs_Store("ShipCompany")
    shipto =rs_store("shipto")
    Cust_PO = Rs_store("Cust_PO")
    User_Id=rs_Store("User_Id")
    Password=rs_Store("Password")
    Company=rs_Store("Company")
    Orders_Total=rs_Store("Orders_Total")
    Orders_Dept=rs_Store("Orders_Dept")
    Tax = Rs_store("Tax")
    Coupon_id = Rs_store("Coupon_id")
    Coupon_Amount = Rs_store("Coupon_Amount")
    Giftcert_id = Rs_store("Giftcert_id")
    Giftcert_Amount = Rs_store("Giftcert_Amount")
    Rewards = Rs_store("Rewards")
    Grand_Total = Rs_store("Grand_Total")
    GGrand_Total = Rs_store("GGrand_Total")
    Purchase_Date = Rs_store("Purchase_Date")
    Invoice_Id = Rs_store("Invoice_Id")
    fship_to = Rs_store("ShipTo")
    ship_location_id = Rs_store("ship_location_id")
End If
rs_Store.Close

if ShipEmail="" then
  ShipEmail=Email
end if
if Orders_Total="" then
    Orders_Total=0
end if
if isNull(Orders_Dept) then
  Orders_Dept=""
end if
sql_update=""

GGrand_Total = max(cdbl(GGrand_Total),0)
GGrand_Total=formatnumber(GGrand_Total,2,0,0,0)

RewardsAdd=0

if GGrand_Total =< 0 then
    Is_Verified = "Yes"
end if

sPayments="Visa,Mastercard,American Express,Discover,Diners Club,eCheck,Debit Card,Solo,Switch,Maestro"

if Verified=0 and Real_Time_Processor<>4 and iProcessedPayment=0 and GGrand_Total > 0 and Is_In_Collection(sPayments,Payment_Method,",") then
    'IF CC PAYMENT, GET CC INFORMATION
    sql_cc="exec wsp_purchase_cc "&store_id&","&oid&",'"&CardName&"','"&EnCrypt(CardNumberInsert)&"','"&CardExpiration&"',"&IssueNumber&","&IssueStart&";"
    fn_print_debug sql_cc
    conn_store.execute sql_cc 
    if Payment_Method = "Visa" or Payment_Method = "American Express" or Payment_Method = "Mastercard" or Payment_Method = "Discover"  then

		    'MATHEMATICAL CHECK THE CC NUMBER
		    if not IsCreditCard(Payment_Method,CardNumber) then
		        fn_purchase_decline oid,"This is not a valid credit card number for "&Payment_Method
            end if
		    if 2000+cint(Request.Form("yy"))=year(now) then
			    if cint(Request.Form("mm"))<month(now) then
				    fn_purchase_decline oid,"The credit card has expired."
			    end if
		    end if
    end if
    
    'IF STORE IS USING A REAL TIME CC PROCESSOR, THEN INCLUDE CODE FOR
    'COMUNICATE WITH THE PAYMENT GATEWAY
    if (Real_Time_Processor = 1 or Real_Time_Processor = 2 or Real_Time_Processor = 5 or Real_Time_Processor=7 or Real_Time_Processor=9 or Real_Time_Processor =10 or Real_Time_Processor =14 or Real_Time_Processor =15 or Real_Time_Processor =21 or Real_Time_Processor =26 or Real_Time_Processor =28 or Real_Time_Processor =29 or Real_Time_Processor =31 or Real_Time_Processor =36) and Return_From <> "PayPal" and Return_From <> "Admin" then
		call sub_check_quantity

	    ' CHECK IF PRECHARGE IS ENABLED
	    sql_select = "SELECT Precharge_Enable,maxmind_enable,Precharge_MerchantID,Precharge_Security1,Precharge_Security2,maxmind_License,maxmind_reject FROM Store_external WITH (NOLOCK) WHERE Store_id="&Store_id
	    fn_print_debug sql_select
	    rs_Store.open sql_select,conn_store,1,1
	    If not rs_Store.eof then
		    preCharge_Enable = rs_Store("Precharge_Enable")
		    preCharge_MerchantID = rs_Store("Precharge_MerchantID")
		    preCharge_Security1 = rs_Store("Precharge_Security1")
		    preCharge_Security2 = rs_Store("Precharge_Security2")
		    maxmind_Enable = rs_Store("maxmind_Enable")
		    maxmind_License = rs_Store("maxmind_License")
		    maxmind_reject = rs_Store("maxmind_reject")
	    End If
	    rs_Store.Close

	    if maxmind_enable then
		     %><!--#include file="maxmind/maxmind.asp"--><%																																										'	 call create_form_post_precharge	 																																										'	 call create_form_post_precharge
	    end if
	    If preCharge_Enable then
		     %><!--#include file="precharge/precharge.asp"--><%																																										'	 call create_form_post_precharge
	    End If

	    buf = Grand_Total
	    Grand_Total = GGrand_Total
	    
	    if Real_Time_Processor = 1 then
		    %><!--#include file="plugnpay\pnp.asp"--><%
	    elseif Real_Time_Processor = 2 then
		    %><!--#include file="Authorizenet\Authorizenet.asp"--><%
	    elseif Real_Time_Processor = 5 then
		    %><!--#include file="psigate\psigate.asp"--><%
	    elseif Real_Time_Processor = 7 then
		    %><!--#include file="echo\echo.asp"--><%
	    elseif Real_Time_Processor = 9 then
		    %><!--#include file="verisign\verisign.asp"--><%
	    elseif Real_Time_Processor = 10 then
		    %><!--#include file="linkpoint\linkpoint2.asp"--><% 
	    elseif Real_Time_Processor = 14 then
		    %><!--#include file="bluepay\bluepay.asp"--><%
	    elseif Real_Time_Processor = 15 then
		    %><!--#include file="electronic\electronic.asp"--><%
	    elseif Real_Time_Processor = 26 then
		    %><!--#include file="eftnet\eftnet.asp"--><%
	    elseif Real_Time_Processor = 28 then
		    %><!--#include file="cybersource\cybersource.asp"--><%
	    elseif Real_Time_Processor = 29 then
		    %><!--#include file="xor\xor.asp"--><%
	    elseif Real_Time_Processor = 31 then
		    %><!--#include file="propay\propay.asp"--><%
	    elseif Real_Time_Processor = 36 then
		    %><!--#include file="PayPalPro\PayPalPro.asp"--><%
	    end if
	    Is_Verified = "Yes"
	    Verified=1
	    sql_realtime = "exec wsp_purchase_verify_local "&Store_Id&","&Oid&","&Verified&",'"&AuthNumber&"','"&Verified_Ref&"',"&Real_Time_Processor&",'"&avs_result&"','"&card_verif&"','"&trans_type&"';"&vbcrlf
	    fn_print_debug sql_realtime
	    session("sql")=sql_realtime
	    conn_store.execute sql_realtime
		
	    Grand_Total = buf
	    'PAYMENT GATEWAY CHECKED AND APROVED THE PAYMENT,
	    'SET THE ORDER AS VERIFIED
    Else
        
    end if
end if

Cust_PO=""
Payment_Method_lower=lcase(Payment_Method)
if Payment_Method_lower = "charge my account"  then
    'UPDATE CUSTOMERS TABLE TO REFLECT THE NEW BUDGET
    Is_Verified = "Yes"
    sql_update = sql_update&"exec wsp_budget_update "&store_id&","&cid&","&oid&","&GGrand_Total&";"&vbcrlf
Elseif Payment_Method_lower = "purchase order" then
    Is_Verified="No"
    Cust_PO=request.form("Cust_PO")
Elseif Payment_Method_lower = "call me" OR Payment_Method_lower = "check" OR Payment_Method_lower = "money order" OR Payment_Method_lower = "cash on delivery (cod)" OR Payment_Method_lower = "fax order in" then
    Is_Verified = "No"
end if

If Enable_Rewards then
    RewardsAdd = cdbl((GGrand_Total) * cdbl(Rewards_Percent) / 100)
else
    RewardsAdd = 0
end if

if Is_Verified <> "Yes" then
    ' double check table to make sure not already verified
    if Verified then
	    Is_Verified = "Yes"
    end if
end if
if Is_Verified = "Yes" or Verified then
    Verified = 1
else
    Verified = 0
end if
			
'COMMIT THE SHOPPING TRANSACTION
sql_select_other_order = "select oid, shipto,ship_location_id from store_purchases WITH (NOLOCK) where store_id="&store_id&" AND (oid="&oid&" or masteroid="&oid&")" 
fn_print_debug sql_select_other_order
set myfieldspurch=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select_other_order,mydatapurch,myfieldspurch,noRecordsPurch)

if noRecordsPurch = 0 then
    FOR rowcounter= 0 TO myfieldspurch("rowcount")
    	fn_print_debug "1st rowcounter="&myfieldspurch("rowcount")
        loid = mydatapurch(myfieldspurch("oid"),rowcounter)
        lshipto = mydatapurch(myfieldspurch("shipto"),rowcounter)
        lship_location_id = mydatapurch(myfieldspurch("ship_location_id"),rowcounter)
        sql_select = "exec wsp_trans_final "&store_Id&","&Shopper_Id&","&lshipto&","&lship_location_id&";"&vbcrlf
        set myfieldstrans=server.createobject("scripting.dictionary")
        Call DataGetrows(conn_store,sql_select,mydatatrans,myfieldstrans,noRecordstrans)
        if noRecordstrans = 0 then
            FOR rowcountertrans= 0 TO myfieldstrans("rowcount")
                department_id = mydatatrans(myfieldstrans("department_id"),rowcountertrans)
                gift_id = mydatatrans(myfieldstrans("gift_id"),rowcountertrans)
                item_pin = mydatatrans(myfieldstrans("item_pin"),rowcountertrans)
                quantity_control = mydatatrans(myfieldstrans("quantity_control"),rowcountertrans)
                item_sku = mydatatrans(myfieldstrans("item_sku"),rowcountertrans)
                item_name = mydatatrans(myfieldstrans("item_name"),rowcountertrans)
                item_id = mydatatrans(myfieldstrans("item_id"),rowcountertrans)
                quantity = mydatatrans(myfieldstrans("quantity"),rowcountertrans)
                if Is_In_Collection(Orders_Dept,department_id,",") then
                    if Orders_Dept="" then
                        Orders_Dept=department_id
                    else
                        Orders_Dept=Orders_Dept&","&department_id
                    end if
                end if
                if item_pin<>0 then
                    sql_update = sql_update & fn_assign_pins(item_sku,quantity) 
                end if
                if not isNull(gift_id) and gift_id<>0 then
                    u_d_1 = mydatatrans(myfieldstrans("u_d_1"),rowcountertrans)
                    u_d_2 = mydatatrans(myfieldstrans("u_d_2"),rowcountertrans)
                    gift_amount = mydatatrans(myfieldstrans("gift_amount"),rowcountertrans)
                    'make gift certificate
                    sql_update = sql_update & fn_assign_gift(gift_id,gift_amount,u_d_1,u_d_2,quantity)
                end if
                if quantity_control<>0 then
                    quantity_control_number = mydatatrans(myfieldstrans("quantity_control_number"),rowcountertrans)
                    quantity_in_stock = mydatatrans(myfieldstrans("quantity_in_stock"),rowcountertrans)
                    sql_update = sql_update & fn_update_stock(item_id,item_sku,item_name,quantity,quantity_in_stock,quantity_control_number)
                end if
            next
        end if
    
    next
end if

if Giftcert_Id<>"" and Giftcert_Amount>0 then
    sql_update = sql_update & "exec wsp_giftcert_redeem "&store_id&","&giftcert_amount&",'"&giftcert_id&"';"&vbcrlf
end if

sql_update = sql_update&"exec wsp_purchase_update "&Store_Id&","&Oid&","&shopper_Id&","&Verified&",'"&cust_po&"';"&vbcrlf
sql_update = sql_update&"exec wsp_customer_update "&Store_Id&","&Oid&","&Cid&","&GGrand_Total&",'"&Orders_Dept&"',"&RewardsAdd-Rewards&";"&vbcrlf


on error goto 0
fn_print_debug sql_update
session("sql")=sql_update
conn_store.Execute sql_update

call affiliate_notify

if order_notification_to_admin_enable or order_notification_enable then
    FOR rowcountersend= 0 TO myfieldspurch("rowcount")
        loid = mydatapurch(myfieldspurch("oid"),rowcountersend)
        lshipto = mydatapurch(myfieldspurch("shipto"),rowcountersend)
        lship_location_id = mydatapurch(myfieldspurch("ship_location_id"),rowcountersend)
        call send_order_notifications(loid)
    next
end if

if (Real_Time_Processor=17 and iProcessedPayment=1) or Real_Time_Processor=38 then
    'dont do anything, worldpay and 2checkout wont allow a redirect it causes issues
else
	fn_redirect Switch_Name&"Recipiet.asp"
end if
%>

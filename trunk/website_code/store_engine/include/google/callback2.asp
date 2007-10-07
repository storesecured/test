<%
'******************************************************************************
' Copyright (C) 2006 Google Inc.
'  
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'      http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'******************************************************************************
on error resume next
%>

<!--#include virtual="common/connection.asp"-->
<!--#include virtual="common/common_functions.asp"-->


<!-- #include virtual="include/before_payment_action_include.asp" -->



<%
'server.scripttimeout = 300
  sServerPath=server.MapPath("/")

  sub_write_log "sServerPath="&sServerPath
sServerPathArray=split(sServerPath,"\")
sSiteFolder=sServerPathArray(3)
sub_write_log "sSiteFolder="&sSiteFolder
sGroupFolder=sServerPathArray(2)
sub_write_log "sGroupFolder="&sGroupFolder
Store_id=replace(sSiteFolder,"files_","")
sub_write_log "store_id="&store_id
sql_select = "exec wsp_settings_select "&store_id&";"
sub_write_log sql_select
rs_store.open sql_select,conn_store,1,1
if not rs_store.eof then
 Affiliate_amount=rs_Store("Affiliate_amount")
	Affiliate_cookie=rs_Store("Affiliate_cookie")
	Affiliate_payout=rs_Store("Affiliate_payout")
	Affiliate_type=rs_Store("Affiliate_type")
	AllowCookies =rs_Store("AllowCookies")
	Auth_Capture=rs_Store("Auth_Capture")
    Auto_RMA_Days=rs_Store("Auto_RMA_Days")
	Cart_thumbnails= rs_Store("Cart_thumbnails")
    Content_Width = rs_Store("Content_Width")
	Continue_Shopping = rs_Store("Continue_Shopping")
	Countries_Selected=rs_store("Countries_Selected")
    Coupon= rs_Store("Coupon")
    Department_Layout=rs_store("Department_Layout")
	dept_display=rs_Store("dept_display")
    dept_rows=rs_Store("dept_rows")
	Detail_NextPrev = rs_Store("Detail_NextPrev")
	Dict_Accessory = rs_Store("Dict_Accessory")
    Email_Me= rs_Store("Email_Me")
    Enable_affiliates=rs_Store("Enable_affiliates")
	Enable_Ip_Tracking=rs_Store("Enable_Ip_Tracking")
	Enable_Rewards=rs_Store("Enable_Rewards")
	Enable_RMA=rs_Store("Enable_RMA")
	ExpressCheckout =rs_Store("ExpressCheckout")
	Gift= rs_Store("Gift")
	Gift_Message=rs_Store("Gift_Message")
	Gift_Service=rs_Store("Gift_Service")
	Gift_Wrapping_surcharge=rs_Store("Gift_Wrapping_surcharge")
	GoogleCheckout= rs_Store("GoogleCheckout")
	sHandling_Fee=rs_Store("Handling_Fee")
	Handling_Weight=rs_Store("Handling_Weight")
	sHandling_Weight=rs_Store("Handling_Weight")
	Html_Notifications=rs_store("Html_Notifications")
    Hide_Empty_Depts=rs_Store("Hide_Empty_Depts")
	Hide_outofStock_Items=rs_store("Hide_outofstock_items")
    Hide_Quantity=rs_Store("Hide_Quantity")
	Hide_Retail_Price=rs_Store("Hide_Retail_Price")
    Inventory_Reduce=rs_Store("Inventory_Reduce")
	Invoice_Num=rs_Store("Invoice_Num")
	item_display=rs_Store("item_display")
	Item_F_display=rs_Store("Item_F_display")
	Item_F_Layout=rs_store("Item_F_Layout")
	item_f_rows=rs_Store("item_f_rows")
	Item_L_Layout=rs_store("Item_L_Layout")
	item_rows=rs_Store("item_rows")
	Item_S_Layout=rs_store("Item_S_Layout")
	Minimum_Amount=rs_Store("Minimum_Amount")
	No_Login= rs_Store("No_Login")
    Overdue_Payment=rs_Store("Overdue_Payment")
	Paypal_Express=rs_Store("Paypal_Express")
	Real_Time_Processor=rs_Store("Real_Time_Processor")
	Reload_Attr=rs_Store("Reload_Attr")
	Rewards_Minimum=rs_Store("Rewards_Minimum")
	Rewards_Percent=rs_Store("Rewards_Percent")
	Save_Cart=rs_Store("Save_Cart")
	Screen_affiliates=rs_Store("Screen_affiliates")
	Secure_Name= rs_Store("Secure_Name")
    Service_Type=rs_Store("Service_Type")
	Ship_Multi= rs_Store("Ship_Multi")
	Shipping_Class= rs_Store("Shipping_Classes")
	Shipping=rs_store("Shipping")
	Shopping_Cart=rs_store("Shopping_Cart")
	Show_Coupon=rs_Store("Show_Coupon")
    Show_Jump=rs_Store("Show_Jump")
	Show_Residential=rs_Store("Show_Residential")
	Show_SecureLogo=rs_Store("Show_SecureLogo")
	Show_Shipping= rs_Store("Show_Shipping")
	Show_Special_Offers=rs_Store("Show_Special_Offers")
	Show_Tax_Exempt=rs_Store("Show_Tax_Exempt")
	Show_TopNav=rs_Store("Show_TopNav")
	Site_name= rs_Store("Site_name")
    Site_name_Orig= rs_Store("Site_name")
    Special_Notes=rs_store("Special_Notes")
    StartCID=rs_Store("StartCID")
	StartOID=rs_Store("StartOID")
	StartTID=rs_Store("StartTID")
	Store_active=rs_Store("Store_active")
	Store_address1= rs_Store("Store_address1")
	Store_address2= rs_Store("Store_address2")
	Store_City= rs_Store("Store_City")
	Store_Company= rs_Store("Store_Company")
	Store_Country= rs_Store("Store_Country")
	Store_currency= rs_Store("Store_currency")
	Store_domain= rs_Store("Store_domain")
    Store_domain1= rs_Store("Store_domain")
    Store_domain2= rs_Store("Store_domain2")
    Store_Email= rs_Store("Store_Email")
	Store_Fax= rs_Store("Store_Fax")
	Store_Homepage = rs_Store("Store_Homepage")
    Store_name= rs_Store("Store_name")
	Store_Phone= rs_Store("Store_Phone")
	Store_Public=rs_Store("Store_Public")
    Store_State= rs_Store("Store_State")
	Store_Zip= rs_Store("Store_Zip")
	Subdept_location=rs_store("Subdept_location")
    Top_Depts= rs_Store("Top_Depts")
	Trial_Version= rs_Store("Trial_Version")
	Unverified_Reduce=rs_Store("Unverified_Reduce")
	USE_AIRBORNE=rs_Store("USE_AIRBORNE")
	USE_CANADA=rs_Store("USE_CANADA")
	USE_CONWAY=rs_Store("USE_CONWAY")
	Use_CVV2=rs_Store("Use_CVV2")
	USE_DHL=rs_Store("USE_DHL")
	Use_Domain_Name= rs_Store("Use_Domain_Name")
    USE_FEDEX=rs_Store("USE_FEDEX")
	USE_UPS=rs_Store("USE_UPS")
	USE_USPS=rs_Store("USE_USPS")
	User_Defined_Fields = rs_Store("User_Defined_Fields")
	User_Defined_Fields_2 = rs_Store("User_Defined_Fields_2")
	User_Defined_Fields_3 = rs_Store("User_Defined_Fields_3")
	User_Defined_Fields_4 = rs_Store("User_Defined_Fields_4")
	User_Defined_Fields_5 = rs_Store("User_Defined_Fields_5")
	When_Adding=rs_Store("When_Adding")
'sub_write_log "Auth_Capture=" & Auth_Capture
end if   
rs_store.close
%>
 <!--#include virtual="include/emails.asp"-->
<!--#include virtual="include/Send_Invoice_By_Mail.asp"-->
<!-- #include virtual="include/check_out_payment_include.asp" -->
<!-- #include virtual="include/Sub.asp" -->
  <!-- #include file="googleglobal.asp" -->
<%
sub_write_log "in callback2.asp"
dim Shopper_Id
Shopper_Id=request.QueryString("Shopper_id")
'const EXPRESS_CHECKOUT_CUSTOMER = "--ExpressCheckOut Customer--"
dim cid
Dim ResponseXml, biData, nIndex



biData = Request.BinaryRead(Request.TotalBytes)

For nIndex = 1 To LenB(biData)
	ResponseXml = ResponseXml & Chr(AscB(MidB(biData,nIndex,1)))
Next

LogMessage ResponseXml
'sub_write_log ResponseXml



Shipping_Class=replace(Shipping_Class," ","")
'sub_write_log "class before callback2: " & Shipping_Class

Dim MyOrder
Set MyOrder = New Order
 dim myOID
Dim ResponseNode, RootTag
Set ResponseNode = GetRootNode(ResponseXml)
RootTag = ResponseNode.Tagname
LogMessage RootTag
 sub_write_log RootTag
Select Case RootTag
	Case "new-order-notification"
	       ProcessNewOrderNotification ResponseXml
	Case "order-state-change-notification"
		ProcessOrderStateChangeNotification ResponseXml
	Case "risk-information-notification"
		ProcessRiskInformationNotification ResponseXml
	Case "charge-amount-notification"
		ProcessChargeAmountNotification ResponseXml
	Case "chargeback-amount-notification"
		ProcessChargebackAmountNotification ResponseXml
	Case "refund-amount-notification"
		ProcessRefundAmountNotification ResponseXml
	Case "authorization-amount-notification"
		ProcessAuthorizationAmountNotification ResponseXml
	Case "merchant-calculation-callback"
		ProcessMerchantCalculationCallback ResponseXml
	Case Else
End Select
' Handle a new-order-notifiation message
Sub ProcessNewOrderNotification(ResponseXml)
     sub_write_log "in process new order"
        sub_write_log ResponseXml

	Dim MyNewOrder
	Set MyNewOrder = New NewOrderNotification
	MyNewOrder.ParseNotification ResponseXml
	sub_write_log "after parse notification"
' ********** How to access NewOrderNotification variables **********
	' ***************************** BEGIN ******************************

	'DoSomethingWith MyNewOrder.TimeStamp
	Verified_Ref = MyNewOrder.GoogleOrderNumber
	sub_write_log Verified_Ref
	Shipping_Method_name = MyNewOrder.ShippingName
	prices = MyNewOrder.ShippingCost
	'DoSomethingWith MyNewOrder.BuyerId
	BuyerId=MyNewOrder.BuyerId
	'DoSomethingWith MyNewOrder.MerCalcSuccessful
	'DoSomethingWith MyNewOrder.FulfillmentOrderState
	'DoSomethingWith MyNewOrder.FinancialOrderState
	Adjustment= MyNewOrder.AdjustmentTotal
	sub_write_log "Adjustment :" & Adjustment
	Tax = MyNewOrder.TaxTotal
	Grand_Total = MyNewOrder.OrderTotal
	sub_write_log "Grand_Total :" & Grand_Total
	Spam = MyNewOrder.MarketingEmailAllowed
	sub_write_log "Spam :" & Spam
	if Spam="false" then
	   Spam=0
	else
	    Spam=1
	end if

        sub_write_log "before private data"

	mysum=Tax+prices
        sub_write_log "mysum :" & mysum

        If IsObject(MyNewOrder.MerchantPrivateData) Then
		Dim MyPrivateData

		on error resume next
		Set MyPrivateData = MyNewOrder.MerchantPrivateData
		Shopper_id= GetElementText(MyPrivateData, "session-id")
		myOID=GetElementText(MyPrivateData, "cart-id")
		sub_write_log "Order-id:" & myOID
		 sub_write_log "after private data"
		' Or get the raw <merchant-private-data> XML
		'DoSomethingWith MyPrivateData.xml
		Set MyPrivateData = Nothing
	End If
        sub_write_log "after private data"

	Dim BillingAddress
	Set BillingAddress = MyNewOrder.BuyerBillingAddress
	'User_Id = EXPRESS_CHECKOUT_CUSTOMER

	sub_write_log "lookup customer username"
				User_Id = "--GoogleCheckout " & BuyerId & "--"
                                sql_select = "select cid,ccid from store_customers WITH (NOLOCK) where store_id="&store_id&" and User_ID='"&User_Id&"'"
    sub_write_log sql_select
    rs_store.open sql_select
    if not rs_store.eof then
       cid = rs_store("cid")
       ccid = rs_store("ccid")
       sql_update="update Store_Customers set Spam=" & Spam & " where Store_ID=" & store_id & " and Cid=" & cid & " and CCid=" & ccid
       conn_store.execute sql_update
       rs_store.close
       else
      Randomize
	Password = "PASS_"&cstr(Int((10000) * Rnd + lowerbound))&month(now)&day(now)&year(now)&cstr(Int((10000) * Rnd + lowerbound))

        sub_write_log "after username"

	Contact_name = BillingAddress.ContactName
	Company = BillingAddress.CompanyName
	Address1 = BillingAddress.Address1
	Address2 = BillingAddress.Address2
	City = BillingAddress.City
	State = BillingAddress.Region
	Zip = BillingAddress.PostalCode

	Country = BillingAddress.CountryCode
        sub_write_log Country
        sql_sel_country= "exec wsp_country_name "&Country&";"
        sub_write_log sql_sel_country
        Set rs_Store_country = Server.CreateObject("ADODB.Recordset")
        rs_Store_country.open sql_sel_country, conn_store, 1, 1
        Country=rs_Store_country("country")
        rs_Store_country.close

	Email = BillingAddress.Email
	Phone = BillingAddress.Phone
	Fax = BillingAddress.Fax
	Set BillingAddress = Nothing
	
	sub_write_log "after billing Contact_name="&Contact_name

	ContactNameArray = split(Contact_name," ")
	First_name=ContactNameArray(0)
	if ubound(ContactNameArray)>1 then
           Last_Name=ContactNameArray(2)
        else
          if ubound(ContactNameArray)>0 then
           Last_Name=ContactNameArray(1)
	   else
	    Last_Name=""
	    end if
        end if
        sub_write_log "after contact name split"

	Tax_Exempt=0
	Is_Residential=1
	
	sub_write_log "before sql create"
	
	sql_create_customer = "exec wsp_customer_register "&Store_id&",'"&Shopper_Id&_
	    "','"&User_ID&"','"&Password&"','"&Last_name&"','"&First_name&"','"&Company&_
	    "','"&Address1&"','"&Address2&"','"&City&"','"&Zip&"','"&State&"','"&Country&_
	    "','"&Phone&"','"&EMail&"','"&FAX&"',"&Spam&","&StartCID&","&Tax_Exempt&","&Is_Residential
        sub_write_log sql_create_customer
        conn_store.execute sql_create_customer

      sql_select = "select cid,ccid from store_customers WITH (NOLOCK) where store_id="&store_id&" and password='"&password&"' and record_type=0"
    sub_write_log sql_select
    rs_store.open sql_select
    if not rs_store.eof then
       cid = rs_store("cid")
       ccid = rs_store("ccid")
    end if
    rs_store.close
    end if

    Is_Residential=1

    sub_write_log "before shipping"
	Dim ShippingAddress
	sub_write_log "after dim"
	Set ShippingAddress = MyNewOrder.BuyerShippingAddress
	'DoSomethingWith ShippingAddress.ID
	sub_write_log "after set shipping"
	ShipContactName = ShippingAddress.ContactName
	sub_write_log "after contact"

	ShipCompany = ShippingAddress.CompanyName
	sub_write_log "after company"

	ShipAddress1 = ShippingAddress.Address1
	sub_write_log "after address"

	ShipAddress2 = ShippingAddress.Address2
	sub_write_log "after address 2"

	ShipCity = ShippingAddress.City
	sub_write_log "after city"
	ShipState = ShippingAddress.Region
	ShipZip = ShippingAddress.PostalCode
	ShipCountry = ShippingAddress.CountryCode
	sub_write_log ShipCountry
        sql_sel_country= "exec wsp_country_name "&ShipCountry&";"
        sub_write_log sql_sel_country
        Set rs_Store_country = Server.CreateObject("ADODB.Recordset")
        rs_Store_country.open sql_sel_country, conn_store, 1, 1
        ShipCountry=rs_Store_country("country")
        rs_Store_country.close
	ShipEmail = ShippingAddress.Email
	ShipPhone = ShippingAddress.Phone
	sub_write_log "after phone"

	ShipFax = ShippingAddress.Fax
	Set ShippingAddress = Nothing
	sub_write_log "after shipping"
	
	ShipContactNameArray = split(ShipContactName," ")
	ShipFirstname=ShipContactNameArray(0)
	if ubound(ShipContactNameArray)>0 then
	   ShipLastname=ShipContactNameArray(1)
	else
            ShipLastname=""
        end if
        Record_type = 1

	sql_update_customer = "exec wsp_customer_register_update "&Store_Id&","&Cid&","&Record_type&_
	    ",'"&ShipLastname&"','"&ShipFirstname&"','"&ShipCompany&"','"&ShipAddress1&"','"&ShipAddress2&"','"&ShipCity&_
	    "','"&ShipZip&"','"&ShipState&"','"&ShipCountry&"','"&ShipPhone&"','"&ShipEMail&"','"&ShipFAX&"',"&Is_Residential&";"
	sub_write_log sql_update_customer
	conn_store.execute sql_update_customer

     sql_login = "exec wsp_customer_login "&Shopper_id&",'"&User_Id&"','"&Password&"';"
     conn_store.execute sql_login

     
	Coupon_Amount = 0
	Coupon_Id = "" 

	If MyNewOrder.CouponIndex >= 0 Then
		Dim MyCoupon
		For Each MyCoupon In MyNewOrder.CouponArr
                         Set rs_coup = Server.CreateObject("ADODB.Recordset")
			    Coupon_Id = MyCoupon.Code
			    applied_amount=MyCoupon.AppliedAmount
			    CalculatedAmount=MyCoupon.CalculatedAmount
                            sub_write_log "applied_amount : " & applied_amount & " CalculatedAmount" & CalculatedAmount
		          sql_coupon  = "exec wsp_coupon_valid "&Store_id&","&Shopper_Id&",'"&Coupon_Id&"',"&cid&";"
                                        sub_write_log sql_coupon
                                        rs_coup = conn_store.execute(sql_coupon)
                                      	on error resume next
                                        Coupon_Type = rs_coup("Coupon_Type")

    if err.number<>0 then

    else
    Coupon_Amount = rs_coup("Coupon_Amount")
    Discounted_items_ids=rs_coup("Discounted_items_ids")
    Is_Exclusion=rs_coup("Is_Exclusion")
    Coupon_total=cdbl(rs_coup("Total"))

    sub_write_log "Coupon_Amount="&Coupon_Amount

    if Discounted_items_ids = "" then
        coupon_price = cdbl(Coupon_total)
    else
        sql_coupon = "exec wsp_coupon_amount "&Store_id&","&Shopper_Id&","&cint(is_exclusion)&",'"&discounted_items_ids&"';"
        sub_write_log sql_coupon
        set rs_store = conn_store.execute(sql_coupon)
        coupon_price = rs_store("Coupon_total")
    end if

    sub_write_log "coupon_price="&coupon_price

    if isNull(coupon_price) then
         else
        coupon_price=cdbl(coupon_price)



    if Coupon_Type = 0 then
	    'PERCENT FROM TOTAL
	    Coupon_Discount = (Coupon_Amount/100)*coupon_price
    Else
	    'FLAT TOTAL

	    Coupon_Discount = Coupon_Amount



    End if

   sql_update = "exec wsp_coupon_apply "&Store_Id&","&Shopper_Id&",'"&Coupon_Id&"',"&applied_amount&";"
    sub_write_log sql_update
    conn_store.execute sql_update
    end if
    end if
			'DoSomethingWith MyCoupon.Message
		Next

	End If
    Coupon_Amount=applied_amount
    Giftcert_Amounts = 0
    giftcert_id = ""
	If MyNewOrder.GiftCertIndex >= 0 Then
		Dim MyGiftCert
		For Each MyGiftCert In MyNewOrder.GiftCertArr
                Set rs_coup = Server.CreateObject("ADODB.Recordset")
			    giftcert_id = MyGiftCert.Code
			    applied_amount=MyGiftCert.AppliedAmount
			    CalculatedAmount=MyGiftCert.CalculatedAmount
                            sub_write_log "applied_amount : " & applied_amount & " CalculatedAmount" & CalculatedAmount

			sql_coupon = "exec wsp_giftcert_valid "&Store_id&","&Shopper_Id&",'"&giftcert_id&"';"
                                       sub_write_log sql_coupon
                                        rs_coup = conn_store.execute(sql_coupon)
	                                on error resume next
                                        Current_Amount = rs_coup("Current_Amount")
                                        if err.number<>0 then
                                        else
                                        Giftcert_Amounts=Current_Amount
                                	on error goto 0
                                   	set rs_coup = Nothing

                                         Giftcert_Amounts= applied_amount
                                  	sql_update = "exec wsp_giftcert_apply "&Store_id&","&Shopper_Id&",'"&Giftcert_id&"',"&applied_amount&";"
                                        sub_write_log sql_update
                                        conn_store.execute sql_update
                                       end if
		Next

	End If

	shipAddr = 1
    Verified = 0
    Reward_Amounts = 0
    Payment_Method = "Google Checkout"
    Recurring_Total = 0
    Recurring_Days = 0
    Recurring_Tax = 0
    Recurring_Grand_Total = 0
    Cust_PO = ""
    custom_fields = ""
    Is_Gifts = 0
    Gift_Wrappings = 0
    Gift_Messages = ""
    Gift_Wrapping_amount = 0
    if Coupon_Amount="" then
    		Coupon_Amount=0
    	end if
    ShipResidential = 0
    ship_location_id=-1
    
    sql_select = "select sum(Quantity*wholesale_price) as wholesale_price,sum(Quantity*sale_price) as total from Store_Transactions where store_id="&store_id&" and shopper_id='"&shopper_id&"' and OID=" & myOID
    sub_write_log sql_select
	conn_store.execute sql_select
   rs_store.open sql_select
    wholesale_price=rs_store("wholesale_price")
    Total=rs_store("total")
    rs_store.close

     masteroid=1


    sql_update = "exec wsp_purchase_fill "&masteroid&","&myOID&","&CCID&_
	    ","&shipAddr&","&Store_id&_
	    ",'"&Shopper_ID&"',"&Cid&",'"&ShipLastname&_
	    "','"&ShipFirstname&"','"&ShipCompany&"','"&ShipAddress1&_
	    "','"&ShipAddress2&"','"&ShipCity&"','"&ShipZip&_
	    "','"&ShipState&"','"&ShipCountry&"','"&ShipPhone&_
	    "','"&ShipFax&"','"&ShipEMail&"',"&Tax&","&prices&_
	    ","&Total&","&wholesale_price&","&Grand_Total&_
            ","&Reward_Amounts& "," & Giftcert_Amounts & "," & Coupon_Amount & ",'"&Payment_Method&_
	    "','"&Shipping_Method_name&"','"&cust_notess&_
	    "','"&custom_fields&"','"&Cust_PO&"',"&Is_Gifts&_
	    ",'"&Gift_Messages&"',"&Gift_Wrappings&_
	    ","&Gift_Wrapping_amount&","&Recurring_Total&_
	    ","&Recurring_Days&","&Recurring_Tax&_
	    ","&Recurring_Grand_Total&","&ship_location_id&_
	    ","&ShipResidential&";"


    sub_write_log sql_update
    conn_store.execute sql_update


    sql_update="update Store_Purchases set Verified_Ref='"& Verified_Ref & "',Processor_id=38 where store_id="&store_id&" and OID=" & myOID
    sub_write_log sql_update
    conn_store.execute sql_update

    'sql_update="update Store_Transactions set transaction_processed=1 where OID=" & myOID & " and store_id="&store_id & " and Shopper_ID='" & Shopper_Id & "'"
    'sub_write_log sql_update
    'conn_store.execute sql_update
    

    


    sql_update="update store_shoppers set OID="&myOID&",Cid="&cid&" where Store_ID="&store_id&" and shopper_id='"&shopper_id&"';"
    sub_write_log sql_update
    conn_store.execute sql_update

    	' ***************************** END ******************************

SendAck

   sub_write_log Secure_Site&"include/check_out_payment_action.asp?Shopper_id="&Shopper_id


End Sub
' Handle a charge-amount-notification message
Sub ProcessChargeAmountNotification(ResponseXml)
        LogMessage ResponseXml
        sub_write_log "in process charge amount"
        sub_write_log ResponseXml
	Dim MyChargeAmountNot
	Set MyChargeAmountNot = New ChargeAmountNotification
	MyChargeAmountNot.ParseNotification ResponseXml

	' ********** How to access ChargeAmountNotification variables **********
	' ****************************** BEGIN *********************************
	'DoSomethingWith MyChargeAmountNot.TimeStamp
	'DoSomethingWith MyChargeAmountNot.GoogleOrderNumber
	'DoSomethingWith MyChargeAmountNot.LatestChargeAmount
	'DoSomethingWith MyChargeAmountNot.TotalChargeAmount
	 Totalchargeamount=MyChargeAmountNot.TotalChargeAmount
         sub_write_log "Total charge amount =" & Totalchargeamount
	' ****************************** END *********************************

	SendAck
End Sub

' Handle a chargeback-amount-notification message
Sub ProcessChargebackAmountNotification(ResponseXml)
        LogMessage ResponseXml
        sub_write_log "in process chargeback amount"
        sub_write_log ResponseXml
	Dim MyChargebackAmountNot
	Set MyChargebackAmountNot = New ChargebackAmountNotification
	MyChargebackAmountNot.ParseNotification ResponseXml

	' ******* How to access ChargebackAmountNotification variables ********
	' ****************************** BEGIN ********************************
	'DoSomethingWith MyChargebackAmountNot.TimeStamp
	'DoSomethingWith MyChargebackAmountNot.GoogleOrderNumber
	'DoSomethingWith MyChargebackAmountNot.LatestChargebackAmount
	'DoSomethingWith MyChargebackAmountNot.TotalChargebackAmount
	' ****************************** END *********************************

	SendAck
End Sub

' Handle a refund-amount-notification message
Sub ProcessRefundAmountNotification(ResponseXml)
        LogMessage ResponseXml
        sub_write_log "in process Refund amount"
        sub_write_log ResponseXml
	Dim MyRefundAmountNot
	Set MyRefundAmountNot = New RefundAmountNotification
	MyRefundAmountNot.ParseNotification ResponseXml

	' ********* How to access RefundAmountNotification variables **********
	' ****************************** BEGIN ********************************
	'DoSomethingWith MyRefundAmountNot.TimeStamp
	'DoSomethingWith MyRefundAmountNot.GoogleOrderNumber
	'DoSomethingWith MyRefundAmountNot.LatestRefundAmount
	'DoSomethingWith MyRefundAmountNot.TotalRefundAmount
	' ****************************** END ********************************

	SendAck
End Sub
Sub ProcessAuthorizationAmountNotification(ResponseXml)
        LogMessage ResponseXml
        sub_write_log "in process authorization amount"
        sub_write_log ResponseXml

	Dim MyAuthAmountNot
	Set MyAuthAmountNot = New AuthorizationAmountNotification
	MyAuthAmountNot.ParseNotification ResponseXml

	' ****** How to access AuthorizationAmountNotification variables *******
	' ****************************** BEGIN *********************************
	'DoSomethingWith MyAuthAmountNot.TimeStamp
	'DoSomethingWith MyAuthAmountNot.GoogleOrderNumber
	'DoSomethingWith MyAuthAmountNot.AvsResponse
	'DoSomethingWith MyAuthAmountNot.CvnResponse
	'DoSomethingWith MyAuthAmountNot.AuthorizationAmount
	'DoSomethingWith MyAuthAmountNot.AuthorizationExpirationDate
	Totalauthorizationamount=MyAuthAmountNot.AuthorizationAmount
         sub_write_log "Total Authorization amount =" & Totalauthorizationamount
	' ******************************* END *********************************

	SendAck
End Sub

Sub ProcessMerchantCalculationCallback(ResponseXml)

    sub_write_log "in process callback"
    sub_write_log Store_id
    sub_write_log ResponseXml
    Dim MyMerCalcCallback
	Set MyMerCalcCallback = New MerchantCalculationCallback
	MyMerCalcCallback.ParseNotification ResponseXml
         Dim MyItem
	For Each MyItem In MyMerCalcCallback.ItemArr
		sub_write_log "item name: " & MyItem.Name
		sub_write_log "item description: " & MyItem.Description
		sub_write_log "item Quantity : " &MyItem.Quantity
		sub_write_log "unit price: " &MyItem.UnitPrice
		sub_write_log "MerchantItemId: " &MyItem.MerchantItemId
		sub_write_log "tax :" & MyItem.TaxTableSelector

	Next


		If IsObject(MyMerCalcCallback.MerchantPrivateData) Then
		Dim MyPrivateData
		Set MyPrivateData = MyMerCalcCallback.MerchantPrivateData
		' Retrieve <session-id>
                Shopper_id=GetElementText(MyPrivateData, "session-id")
                  sub_write_log "Shopper_id :" & Shopper_id
                 sub_write_log "after private data"
		Set MyPrivateData = Nothing

	End If
	BuyerID=MyMerCalcCallback.BuyerId
	sub_write_log "BuyerID : " & BuyerID

	If MyMerCalcCallback.MerchantCodeIndex >= 0 Then
		'Dim MerchantCode
		'For Each MerchantCode In MyMerCalcCallback.MerchantCodeArr
		'	DoSomethingWith MerchantCode
		'Next
	End If


	If MyMerCalcCallback.ShippingMethodIndex >= 0 Then
		Dim ShippingMethod
		For Each ShippingMethod In MyMerCalcCallback.ShippingMethodArr
			'DoSomethingWith ShippingMethod
		Next
	End If

	' ****************************** END **********************************

    
    
    
    
      ' Construct <merchant-calculation-results>
	Dim MyMerCalcResults, MyResult, MyCoupon, MyGiftCert, MerchantCodeType
	dim Giftcert_no,coupon_no
	Set MyMerCalcResults = New MerchantCalculationResults

    Dim arrDualArray(30,2)
    dim shippingArray
    Redim shippingArray(100)

    myflag=0
    SADDR=1
    location_id=0
   'SADDR=-1
    'location_id=1
    sql_select = "select Ship_Location_Id from Store_Transactions where Store_ID="&Store_id&" and Shopper_ID='"&Shopper_id&"' and Transaction_Processed=0"
    sub_write_log sql_select
    rs_store.open sql_select, conn_store, 1, 1
    ship_location_id=rs_store("Ship_Location_Id")
    rs_store.close
    sql_sel_items = "exec wsp_trans_totals_shipping "& Store_id &",'"&Shopper_id&"',"&SADDR&","&location_id&";"
    sub_write_log sql_sel_items
     rs_store.open sql_sel_items, conn_store, 1, 1
    Total_Quantity=FormatNumber(rs_store("TQuantity"),2)
    sub_write_log Total_Quantity
    if Total_Quantity=0 then
        rs_store.close 
      myflag=1
     sub_write_log "free shipping in callback"



 else
    Total_Weight=FormatNumber((rs_store("TWeight")+Handling_Weight),2,,,0)
    Total_Price=FormatNumber(rs_store("TPrice"),2,,,0)
    SFee=FormatNumber(rs_Store("SFee"),2,,,0)
    sum_handling=FormatNumber((rs_store("THandling")+sHandling_fee),2,,,0)
    rs_store.close

   end if

  '  sub_write_log ship_location_id
   ' myclassArr=split(ShippingMethod,"")
'sSiteFolder=myclassArr(0)

    '*************

	' Loop through each anonymous address from the callback
	For Each AnonymousAddress In MyMerCalcCallback.AnonymousAddressArr
	    sub_write_log "in anonymous addr loop"
	    iArrayIndex=0
        sub_write_log "in anonymousaddress loop"

		' Construct and add <result> into <merchant-calculation-results>

		' Loop through each merchant-calculated-shipping method
		If MyMerCalcCallback.ShippingMethodIndex >= 0 Then
                     sub_write_log " there are merchant calculated shipping methods"
                    'DoSomethingWith AnonymousAddress.ID
			'ContactName = AnonymousAddress.ContactName
			'CompanyName = AnonymousAddress.CompanyName
			'Address1 = AnonymousAddress.Address1
			'Address2 = AnonymousAddress.Address2
			'City = AnonymousAddress.City
			State = AnonymousAddress.Region
			Zip = AnonymousAddress.PostalCode
                         dim myZrr
                         myZrr=Split(Zip,"-")
                         myZip=myZrr(0)
			'myZip=replace(Zip,"-","")
                        sub_write_log "Zip=" & myZip
                        Zip=myZip
			Country = AnonymousAddress.CountryCode
			sql_sel_country= "exec wsp_country_id '"&Country&"';"
                        sub_write_log sql_sel_country
                       Set rs_Store_country = Server.CreateObject("ADODB.Recordset")
                        rs_Store_country.open sql_sel_country, conn_store, 1, 1
                        country_id=rs_Store_country("country_id")
                         rs_Store_country.close
			D_dest_data_country_id=203
			'Email = AnonymousAddress.Email
			'Phone = AnonymousAddress.Phone
			'Fax = AnonymousAddress.Fax
			 sub_write_log "shipping class " & Shipping_Class
			 if myflag=0 then
                         sql_shipping = "exec wsp_shipping_display "&Store_Id&","&Total_Weight&","&Total_Price&","&Ship_Location_Id&","&country_id&",'"&Zip&"','"&Shipping_Class&"';"
            sub_write_log sql_shipping
            set shippingfields=server.createobject("scripting.dictionary")
            Call DataGetrows(conn_store,sql_shipping,shippingdata,shippingfields,noRecords)
	      sub_write_log "after call DataGetrows"
                if noRecords=0 then
	            FOR shippingrowcounter= 0 TO shippingfields("rowcount")
	                sub_write_log "in shippingrowcounter loop"
	                this_Shipping_id=shippingdata(shippingfields("Shipping_method_id"),shippingrowcounter)
	                this_Shipping_Class=shippingdata(shippingfields("shippers_class"),shippingrowcounter)
	                this_shipping_method_name=shippingdata(shippingfields("shipping_method_name"),shippingrowcounter)
	                this_base_fee=shippingdata(shippingfields("base_fee"),shippingrowcounter)
	                this_weight_fee=shippingdata(shippingfields("weight_fee"),shippingrowcounter)
	                this_countries=shippingdata(shippingfields("countries"),shippingrowcounter)
	                'check countries again because sql proc cant correctly check 100% of time
	                sub_write_log "this_countries="&this_countries
	                sub_write_log "D_dest_data_country_id="&D_dest_data_country_id
	                pos=InStr(this_countries,"216")
                        If Is_In_Collection(this_countries,D_dest_data_country_id,",") or pos<>0 then
	                    sub_write_log "this_Shipping_Class="&this_Shipping_Class
                        select case this_Shipping_Class
	                        case 1 ship_price=this_base_fee
	                        case 2 ship_price=this_base_fee+(this_weight_fee*total_weight)
	                        case 3 ship_price=this_base_fee+SFee
	                        case 4 ship_price=this_base_fee
	                        case 5 ship_price=this_base_fee*(Total_Price/100)
	                        case 6 ship_price=this_base_fee
	                    end select
	                    ship_price = ship_price+sum_Handling
	                    sub_write_log "ship_price="&ship_price
	                    shippingArray(shippingrowcounter)=this_shipping_method_name
    if shippingrowcounter>0 then
    found=0
    for k=0 to shippingrowcounter-1
    if lcase(this_shipping_method_name)= lcase(shippingArray(k)) then
    found=1
    end if
    next
    if found=0 then
    arrDualArray(iArrayIndex,0) = this_shipping_method_name
				        arrDualArray(iArrayIndex,1) = ship_price
				        arrDualArray(iArrayIndex,2) = this_shipping_method_name
				        iArrayIndex=iArrayIndex+1
    else

    arrDualArray(iArrayIndex,0) = this_Shipping_id&"-"&this_shipping_method_name
				        arrDualArray(iArrayIndex,1) = ship_price
				        arrDualArray(iArrayIndex,2) = this_Shipping_id&"-"&this_shipping_method_name
				        iArrayIndex=iArrayIndex+1
    end if
    else
    arrDualArray(iArrayIndex,0) = this_shipping_method_name
				        arrDualArray(iArrayIndex,1) = ship_price
				        arrDualArray(iArrayIndex,2) = this_shipping_method_name
				        iArrayIndex=iArrayIndex+1
    end if



				        arrDualArray(iArrayIndex,0) = this_shipping_method_name
				        arrDualArray(iArrayIndex,1) = ship_price
				        arrDualArray(iArrayIndex,2) = this_shipping_method_name
				        iArrayIndex=iArrayIndex+1
	                end if
	            next
	        end if
	        set shippingfields=nothing
	        sub_write_log "after shippingfields=nothing"

	        if err.number<>0 then
	           sub_write_log err.description
	        end if
	           end if
	        %>
         <%
	        For Each ShippingMethod In MyMerCalcCallback.ShippingMethodArr
	            sub_write_log "in shippingmethod for loop"
				sFound=0
			    Set MyResult = New Result
			    sub_write_log "after new result"
                         if myflag=1 then
			    MyResult.ShippingName = "Free Shipping"
                            MyResult.ShippingRate = formatnumber(0,2)
                            MyResult.Shippable = True
			 else
                        	ShippingName = ShippingMethod
                sub_write_log "before loop arrDualArray"
                %>
                <%
                For row = 0 To UBound(arrDualArray) - 1
                    sub_write_log "in arrDualArray loop"
                    sThisName = arrDualArray(row,2)
                    if ShippingName=sThisName then
                        sPrice = arrDualArray(row,1)
                        MyResult.ShippingName = sThisName
				        MyResult.ShippingRate = formatnumber(sPrice,2)
				        MyResult.Shippable = True
				        sFound=1

                        exit for
			        end if
			    next
			    %>
			    <%
			    sub_write_log "after for loop"

			    if sFound=0 then
			        MyResult.ShippingName = ShippingMethod
			        MyResult.ShippingRate = formatNumber(0,2)
			        MyResult.Shippable = False
			    end if
                            end if
			    sub_write_log "after sfound=0"
			    'else


				' Get the shipping name, rate and whether it's shippable

				' Anonymous-address ID
				MyResult.AddressId = AnonymousAddress.ID

				' Calculate merchnat-calculated tax
				%>
				<%
				If MyMerCalcCallback.Tax = "true" Then
                    sub_write_log "in calculate tax"
                    this_City = AnonymousAddress.City
                    this_State = AnonymousAddress.Region
                    this_Zip = AnonymousAddress.PostalCode
                    this_Country = AnonymousAddress.CountryCode
                    shipto=-1
                    location_id=-1
                    discount_untax=0
                    ship_taxable=0
                    iTaxAmount = fn_Calc_Tax(0,this_Zip,this_State,this_Country,shipto,location_id,discount_untax,ship_taxable,"single_total")
                    sub_write_log "fn_Calc_Tax(0,"&this_Zip&","&this_State&","&this_Country&","&shipto&","&location_id&","&discount_untax&","&ship_taxable&",'single_total')"
                    MyResult.TotalTax = formatNumber(iTaxAmount,2)
                    sub_write_log "iTaxAmount="&iTaxAmount
				End If
				%>
				<%
				sub_write_log "after tax"
				sub_write_log "lookup customer username"
				User_Id = "--GoogleCheckout " & BuyerID & "--"
                                sql_select = "select cid,ccid from store_customers WITH (NOLOCK) where store_id="&store_id&" and User_ID='"&User_Id&"'"
    sub_write_log sql_select
    Set rs_storeu = Server.CreateObject("ADODB.Recordset")
    rs_storeu.open sql_select, conn_store, 1, 1
    if not rs_storeu.eof then
       cid = rs_storeu("cid")
       ccid = rs_storeu("ccid")
       else
       cid=0
    end if
    rs_storeu.close
   



				' Validate coupons and gift certificates
							If MyMerCalcCallback.MerchantCodeIndex >= 0 Then

							Giftcert_no=0
                                                        coupon_no=0
					For Each MerchantCode In MyMerCalcCallback.MerchantCodeArr
                                  Set rs_coup = Server.CreateObject("ADODB.Recordset")
					' Determine whether the code is a Coupon or GiftCert
				        Coupon_id = checkStringForQ(MerchantCode)
				        sub_write_log "Coupon_id :" & Coupon_id
                                        'fn_print_debug "gcrt="&instr(Coupon_Id,"GCRT_"&Store_Id)
	                                if instr(Coupon_Id,"GCRT_"&Store_Id)=1 then
                                        Giftcert_id=Coupon_id
                                        Coupon_id=""
		                        Giftcert_no=Giftcert_no+1
		                        Set MyGiftCert = New GiftCertResult
					MyGiftCert.Code = MerchantCode
		                        if Giftcert_no>1 then
                                         MyGiftCert.Valid = False
					MyGiftCert.Amount = 0
					MyGiftCert.Message = "Sorry we can't accept more than one Gift certificate per order."
		                        else

		                        Set MyGiftCert = New GiftCertResult
					MyGiftCert.Code = MerchantCode
		                        sql_coupon = "exec wsp_giftcert_valid "&Store_id&","&Shopper_Id&",'"&Giftcert_id&"';"
                                        sub_write_log sql_coupon
                                        rs_coup = conn_store.execute(sql_coupon)
	                                on error resume next
                                        Current_Amount = rs_coup("Current_Amount")
                                        if err.number<>0 then
                                        MyGiftCert.Valid = False
					MyGiftCert.Amount = 0
					MyGiftCert.Message = "Invalid gift certificate."

                                        else
                                        Giftcert_Amounts=Current_Amount
                                	on error goto 0
                                   	set rs_coup = Nothing
                                  	'sql_update = "exec wsp_giftcert_apply "&Store_id&","&Shopper_Id&",'"&Giftcert_id&"',"&Current_Amount&";"
                                        'sub_write_log sql_update
                                        'conn_store.execute sql_update
                                        MyGiftCert.Valid = True
					MyGiftCert.Amount = Current_Amount
					MyGiftCert.Message = "You gift certificate has been applied successfully."
                                        end if
                                        end if
						MyResult.AddGiftCertResult MyGiftCert
						Set MyGiftCert = Nothing
                                        else
                                        Set MyCoupon = New CouponResult
                                         Coupon_id=MerchantCode
                                         MyCoupon.Code = MerchantCode
                                         coupon_no=coupon_no+1
                                         if coupon_no>1 then
                                         MyCoupon.Valid = False
	                                 MyCoupon.Amount = 0
	                                 MyCoupon.Message = "Sorry we can't accept more than one coupon code per order."

                                         else
                                         sql_coupon  = "exec wsp_coupon_valid "&Store_id&","&Shopper_Id&",'"&coupon_id&"',"&cid&";"
                                        sub_write_log sql_coupon
                                        rs_coup = conn_store.execute(sql_coupon)
                                      	on error resume next
                                        Coupon_Type = rs_coup("Coupon_Type")

    if err.number<>0 then
        MyCoupon.Valid = False
	MyCoupon.Amount = 0
	MyCoupon.Message = "Your coupon code is invalid."

    else
    Coupon_Amount = rs_coup("Coupon_Amount")
    Discounted_items_ids=rs_coup("Discounted_items_ids")
    Is_Exclusion=rs_coup("Is_Exclusion")
    Coupon_total=cdbl(rs_coup("Total"))

    sub_write_log "Coupon_Amount="&Coupon_Amount

    if Discounted_items_ids = "" then
        coupon_price = cdbl(Coupon_total)
    else
        sql_coupon = "exec wsp_coupon_amount "&Store_id&","&Shopper_Id&","&cint(is_exclusion)&",'"&discounted_items_ids&"';"
        sub_write_log sql_coupon
        set rs_store = conn_store.execute(sql_coupon)
        coupon_price = rs_store("Coupon_total")
    end if

    sub_write_log "coupon_price="&coupon_price

    if isNull(coupon_price) then
        MyCoupon.Valid = False
	MyCoupon.Amount = 0
	MyCoupon.Message = "The coupon code you have entered is not valid for any of the items you have purchased."
    else
        coupon_price=cdbl(coupon_price)

    

    if Coupon_Type = 0 then
	    'PERCENT FROM TOTAL
	    Coupon_Discount = (Coupon_Amount/100)*coupon_price
    Else
	    'FLAT TOTAL
	    Coupon_Discount = Coupon_Amount
    End if
    
  '  sql_update = "exec wsp_coupon_apply "&Store_Id&","&Shopper_Id&",'"&coupon_id&"',"&Coupon_Discount&";"
   ' sub_write_log sql_update
   ' conn_store.execute sql_update
						MyCoupon.Valid = True
						MyCoupon.Amount = Coupon_Discount
						MyCoupon.Message = "Your coupon has been applied successfully."
						end if
						end if
						end if
						MyResult.AddCouponResult MyCoupon
						Set MyCoupon = Nothing
                                        end if

						%>
						<%
					Next
				End If

                                	sub_write_log "after coupons and gift certificates"

				' Add this <result> to <merchant-calculation-results>
				MyMerCalcResults.AddResult MyResult
				Set MyResult = Nothing
			Next
			%>
			<%
			sub_write_log "outside of first next"


		' No merchant-calculated-shipping methods
		 sub_write_log "No merchant-calculated-shipping methods"
		Else
		    %>
		    <%
            sub_write_log "in else"
			Set MyResult = New Result

			' Specify whether its shippable
			MyResult.Shippable = True

			' Anonymous-address ID
			MyResult.AddressId = AnonymousAddress.ID

			' Calculate merchnat-calculated tax
			If MyMerCalcCallback.Tax = "true" Then
				MyResult.TotalTax = 5.00
			End If

			' Validate coupons and gift certificates
								If MyMerCalcCallback.MerchantCodeIndex >= 0 Then

							Giftcert_no=0
                                                        coupon_no=0
					For Each MerchantCode In MyMerCalcCallback.MerchantCodeArr
                                  Set rs_coup = Server.CreateObject("ADODB.Recordset")
					' Determine whether the code is a Coupon or GiftCert
				        Coupon_id = checkStringForQ(MerchantCode)
				        sub_write_log "Coupon_id :" & Coupon_id
                                        'fn_print_debug "gcrt="&instr(Coupon_Id,"GCRT_"&Store_Id)
	                                if instr(Coupon_Id,"GCRT_"&Store_Id)=1 then
                                        Giftcert_id=Coupon_id
                                        Coupon_id=""
		                        Giftcert_no=Giftcert_no+1
		                        Set MyGiftCert = New GiftCertResult
					MyGiftCert.Code = MerchantCode
		                        if Giftcert_no>1 then
                                         MyGiftCert.Valid = False
					MyGiftCert.Amount = 0
					MyGiftCert.Message = "Sorry we can't accept more than one Gift certificate per order."
		                        else

		                        Set MyGiftCert = New GiftCertResult
					MyGiftCert.Code = MerchantCode
		                        sql_coupon = "exec wsp_giftcert_valid "&Store_id&","&Shopper_Id&",'"&Giftcert_id&"';"
                                        sub_write_log sql_coupon
                                        rs_coup = conn_store.execute(sql_coupon)
	                                on error resume next
                                        Current_Amount = rs_coup("Current_Amount")
                                        if err.number<>0 then
                                        MyGiftCert.Valid = False
					MyGiftCert.Amount = 0
					MyGiftCert.Message = "Invalid gift certificate."

                                        else
                                        Giftcert_Amounts=Current_Amount
                                	on error goto 0
                                   	set rs_coup = Nothing
                                  	sql_update = "exec wsp_giftcert_apply "&Store_id&","&Shopper_Id&",'"&Giftcert_id&"',"&Current_Amount&";"
                                        sub_write_log sql_update
                                        conn_store.execute sql_update
                                        MyGiftCert.Valid = True
					MyGiftCert.Amount = Current_Amount
					MyGiftCert.Message = "You gift certificate has been applied successfully."
                                        end if
                                        end if
						MyResult.AddGiftCertResult MyGiftCert
						Set MyGiftCert = Nothing
                                        else
                                        Set MyCoupon = New CouponResult
                                         Coupon_id=MerchantCode
                                         MyCoupon.Code = MerchantCode
                                         coupon_no=coupon_no+1
                                         if coupon_no>1 then
                                         MyCoupon.Valid = False
	                                 MyCoupon.Amount = 0
	                                 MyCoupon.Message = "Sorry we can't accept more than one coupon code per order."

                                         else
                                         sql_coupon  = "exec wsp_coupon_valid "&Store_id&","&Shopper_Id&",'"&coupon_id&"',0;"
                                        sub_write_log sql_coupon
                                        rs_coup = conn_store.execute(sql_coupon)
                                      	on error resume next
                                        Coupon_Type = rs_coup("Coupon_Type")

    if err.number<>0 then
        MyCoupon.Valid = False
	MyCoupon.Amount = 0
	MyCoupon.Message = "Your coupon code is invalid."

    else
    Coupon_Amount = rs_coup("Coupon_Amount")
    Discounted_items_ids=rs_coup("Discounted_items_ids")
    Is_Exclusion=rs_coup("Is_Exclusion")
    Coupon_total=cdbl(rs_coup("Total"))
    
    sub_write_log "Coupon_Amount="&Coupon_Amount
    
    if Discounted_items_ids = "" then
        coupon_price = cdbl(Coupon_total)
    else
        sql_coupon = "exec wsp_coupon_amount "&Store_id&","&Shopper_Id&","&cint(is_exclusion)&",'"&discounted_items_ids&"';"
        sub_write_log sql_coupon
        set rs_store = conn_store.execute(sql_coupon)
        coupon_price = rs_store("Coupon_total")
    end if
    
    sub_write_log "coupon_price="&coupon_price
  
    if isNull(coupon_price) then
        MyCoupon.Valid = False
	MyCoupon.Amount = 0
	MyCoupon.Message = "The coupon code you have entered is not valid for any of the items you have purchased."
    else
        coupon_price=cdbl(coupon_price)

    

    if Coupon_Type = 0 then
	    'PERCENT FROM TOTAL
	    Coupon_Discount = (Coupon_Amount/100)*coupon_price
    Else
	    'FLAT TOTAL
	    Coupon_Discount = Coupon_Amount
    End if
    
    'sql_update = "exec wsp_coupon_apply "&Store_Id&","&Shopper_Id&",'"&coupon_id&"',"&Coupon_Discount&";"
    'sub_write_log sql_update
    'conn_store.execute sql_update
						MyCoupon.Valid = True
						MyCoupon.Amount = Coupon_Discount
						MyCoupon.Message = "Your coupon has been applied successfully."
						end if
						end if
						end if
						MyResult.AddCouponResult MyCoupon
						Set MyCoupon = Nothing
                                        end if

						%>
						<%
					Next
				End If



			' Add this <result> to <merchant-calculation-results>
			MyMerCalcResults.AddResult MyResult
			Set MyResult = Nothing
			%>
			<%
		End If
		sub_write_log "outside of end if"
	Next
	sub_write_log "outside of second next"

    sub_write_log "before get results"
	' Respond to Google with <merchant-calculation-results>
	Dim MyMerCalcResultsXml
	MyMerCalcResultsXml = MyMerCalcResults.GetXml
	sub_write_log MyMerCalcResultsXml
      
        response.Write MyMerCalcResultsXml
	LogMessage MyMerCalcResultsXml

	Set MyMerCalcCallback = Nothing
	Set MyMerCalcResults = Nothing





End Sub
'**********************
Sub ProcessOrderStateChangeNotification(ResponseXml)
        sub_write_log "in process order state change"
        sub_write_log ResponseXml

	Dim MyOrderState
	Set MyOrderState = New OrderStateChangeNotification
	MyOrderState.ParseNotification ResponseXml

	' ******** How to access OrderStateChangeNotification variables ********
	' ****************************** BEGIN *********************************



	' ****************************** END *********************************

	Dim GoogleOrderNumber, Amount, Reason, Comment
	Dim MerchantOrderNumber, Carrier, TrackingNumber, Mesage

	GoogleOrderNumber = MyOrderState.GoogleOrderNumber
        sub_write_log "GoogleOrderNumber=" & GoogleOrderNumber
	' Possible list of Order Processing commands

	' Financial Commands:

	' MyOrder.ChargeOrder GoogleOrderNumber, Amount
	' MyOrder.SendOrderCommand

	' MyOrder.RefundOrder GoogleOrderNumber, Reason, Amount, Comment
	' MyOrder.SendOrderCommand

	' MyOrder.CancelOrder GoogleOrderNumber, Reason, Comment
	' MyOrder.SendOrderCommand

	' MyOrder.AuthorizeOrder GoogleOrderNumber
	' MyOrder.SendOrderCommand

	' Fulfillment Commands:

	' MyOrder.ProcessOrder GoogleOrderNumber
	' MyOrder.SendOrderCommand

	' MyOrder.AddMerchantOrderNumber GoogleOrderNumber, MerchantOrderNumber
	' MyOrder.SendOrderCommand

	' MyOrder.DeliverOrder GoogleOrderNumber, Carrier, TrackingNumber, True
	' MyOrder.SendOrderCommand

	' MyOrder.AddTrackingData GoogleOrderNumber, Carrier, TrackingNumber
	' MyOrder.SendOrderCommand

	' MyOrder.SendBuyerMessage GoogleOrderNumber, Mesage, True
	' MyOrder.SendOrderCommand

	' Archiving Commands:

	' MyOrder.ArchiveOrder GoogleOrderNumber
	' MyOrder.SendOrderCommand

	' MyOrder.UnarchiveOrder GoogleOrderNumber
	' MyOrder.SendOrderCommand

        Google_FinancialStatus=MyOrderState.NewFinancialOrderState
        Google_FulfillmentStatus=MyOrderState.NewFulfillmentOrderState
        sql_update="update Store_Purchases set Google_FinancialStatus='" & Google_FinancialStatus & "',Google_FulfillmentStatus='" & Google_FulfillmentStatus & "' where Verified_Ref='"& GoogleOrderNumber & "' and store_id="&store_id
        sub_write_log sql_update
         conn_store.execute sql_update

	Select Case MyOrderState.NewFinancialOrderState
		Case "REVIEWING"

		Case "CHARGEABLE"

                 sub_write_log "Auth_Capture=" & Auth_Capture
                 sql_select="select OID,Cid,Shopper_ID from Store_Purchases where store_id="&store_id&" and Verified_Ref='"&GoogleOrderNumber&"'"
                 sub_write_log sql_select
                 Set rs_store1 = Server.CreateObject("ADODB.Recordset")
                 rs_store1.open sql_select,conn_store,1,1
    if not rs_store1.eof then
    OID=rs_Store1("OID")
    cid=rs_Store1("Cid")
    Shopper_ID=rs_Store1("Shopper_ID")
      end if
      rs_store1.close
         sql_update2="update Store_ShoppingCart set Cart_Processed=1 where Store_ID=" & store_id & " and Shopper_ID='" & Shopper_ID & "'"
         sub_write_log sql_update2
         conn_store.Execute sql_update2







      Budget_Left=0
Reward_Left=0

      sql_select = "exec wsp_purchase_select "&Store_id&","&OID&";"

sub_write_log sql_select
rs_Store.open sql_select,conn_store,1,1
If not rs_Store.eof then
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
end if


if Orders_Total="" then
    Orders_Total=0
end if
if isNull(Orders_Dept) then
  Orders_Dept=""
end if


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

GGrand_Total = cdbl(GGrand_Total)
GGrand_Total=formatnumber(GGrand_Total,2,0,0,0)

RewardsAdd=0

if GGrand_Total =< 0 then
    Is_Verified = "Yes"
end if

buf = Grand_Total
	    Grand_Total = GGrand_Total


      If Enable_Rewards then
    RewardsAdd = cdbl((GGrand_Total) * cdbl(Rewards_Percent) / 100)
       else
    RewardsAdd = 0
     end if

sub_write_log "after RewardsAdd"
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
rs_store.close
'COMMIT THE SHOPPING TRANSACTION
sql_select_other_order = "select oid,shipto,ship_location_id from store_purchases WITH (NOLOCK) where store_id=" &Store_id&" AND (oid="&OID&" or masteroid="&OID&")"
sub_write_log sql_select_other_order
set myfieldspurch=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select_other_order,mydatapurch,myfieldspurch,noRecordsPurch)
if noRecordsPurch = 0 then
    FOR rowcounter= 0 TO myfieldspurch("rowcount")
        loid = mydatapurch(myfieldspurch("oid"),rowcounter)
        lshipto = mydatapurch(myfieldspurch("shipto"),rowcounter)
        lship_location_id = mydatapurch(myfieldspurch("ship_location_id"),rowcounter)
        sub_write_log "Shopper_ID : " & Shopper_ID & " lshipto:" & lshipto & "lship_location_id:" & lship_location_id
        Store_Id=Store_id

        sql_select = "exec wsp_trans_final "&Store_id&","&Shopper_ID&","&lshipto&","&lship_location_id&";"&vbcrlf
        sub_write_log sql_select
        set myfieldstrans=server.createobject("scripting.dictionary")
        Call DataGetrows(conn_store,sql_select,mydatatrans,myfieldstrans,noRecordstrans)
        if noRecordstrans = 0 then
            FOR rowcountertrans= 0 TO myfieldstrans("rowcount")
                sub_write_log " inside wsp_trans_final"
                department_id = mydatatrans(myfieldstrans("department_id"),rowcountertrans)
                gift_id = mydatatrans(myfieldstrans("gift_id"),rowcountertrans)
                item_pin = mydatatrans(myfieldstrans("item_pin"),rowcountertrans)
                quantity_control = mydatatrans(myfieldstrans("quantity_control"),rowcountertrans)
                sub_write_log " quantity_control :" & quantity_control
                item_sku = mydatatrans(myfieldstrans("item_sku"),rowcountertrans)
                item_name = mydatatrans(myfieldstrans("item_name"),rowcountertrans)
                item_id = mydatatrans(myfieldstrans("item_id"),rowcountertrans)
                quantity = mydatatrans(myfieldstrans("quantity"),rowcountertrans)
                sub_write_log "department_id " & department_id & "gift_id " & gift_id & "item_pin " & item_pin & " quantity_control " & quantity_control & " item_id " & item_id & "quantity " & quantity
                shopper_id=Shopper_ID
                oid=OID
                if Is_In_Collection(Orders_Dept,department_id,",") then
                    if Orders_Dept="" then
                        Orders_Dept=department_id
                    else
                        Orders_Dept=Orders_Dept&","&department_id
                    end if
                end if
                if item_pin<>0 then
                sub_write_log "make PIN"

                      sql_pin=""

        sql_new = "select top "&quantity&" pin,other_info from Store_pin WITH (NOLOCK) where store_id="&Store_id & " AND item_sku='" & item_sku & "' AND (pin_used is null or pin_used =0)"
        sub_write_log sql_new
        set myfields=server.createobject("scripting.dictionary")
        Call DataGetrows(conn_store,sql_new,mydata,myfields,noRecords)
                if noRecords = 0 then
	        FOR rowcounterp= 0 TO myfields("rowcount")
                sPin = mydata(myfields("pin"),rowcounterp)
                sOther= mydata(myfields("other_info"),rowcounterp)
                pininfo = sPin & "-" & sOther
                sql_pin = sql_pin & "exec wsp_pin_redeem "&Store_id&","&cid&","&OID&",'"&item_sku&"','"&sPin&"';"
                sub_write_log sql_pin
                Next
            sPins=myfields("rowcount")+1
        else
            sPins=0
        end if
        if sPins<quantity then
            pininfo="Store is out of pins for sku: "&item_sku&" please contact store owner."
    body = "Dear store owner, the StoreSecured system has sent you this email to alert you that the item with sku:"&item_sku&" has run out of pins.  We were unable to auto deliver "&sPinQuantity&" pin(s) for order id "&oid&" to email "&shipemail&"."
    call Send_Mail_Html(sNoReply_email, store_email, "Pins Empty in your store", body)
        sub_write_log ":pins not enough"
        end if

        if pininfo<>"" then
            send_from = Session("Store_Email")
            if send_from="" or isnull(send_from) then
                send_from = Store_Email
            end if

            send_cont = pin_buyer_email_body
            send_cont = send_cont&vbcrlf&"The pin code(s) you have purchased are below: "&pininfo&vbcrlf
            sub_write_log "before pin send email"
            call Send_Mail_Html(send_from, shipemail, pin_email_Subject, send_cont)
        end if

                    sql_update = sql_update & sql_pin
                    sub_write_log sql_update
                end if
                if not isNull(gift_id) then
                    u_d_1 = mydatatrans(myfieldstrans("u_d_1"),rowcountertrans)
                    u_d_2 = mydatatrans(myfieldstrans("u_d_2"),rowcountertrans)
                    gift_amount = mydatatrans(myfieldstrans("gift_amount"),rowcountertrans)
                    sub_write_log "shipemail:" & shipemail
                    sub_write_log "make gift certificate"
                    'make gift certificate
                     sql_gift=""
    for iQuantity=1 to formatnumber(quantity,0)
        Randomize
        GIFT_CODE = 1
		GIFT_CODE = "GCRT_"&Store_id&Oid&GIFT_CODE&cstr(Int((10000) * Rnd + lowerbound))
		GIFT_CODE = ucase(GIFT_CODE)
		GIFT_CODE = clearSpaces(GIFT_CODE)
		
                sql_gift = sql_gift & "exec wsp_giftcert_insert "&Store_Id&","&Shopper_ID&","&cid&",'"&shipemail&"',"&gift_id&","&OID&",'"&gift_code&"',"&gift_amount&",'"&u_d_1&"','"&u_d_2&"',"&verified&";"
            sub_write_log "make gift certificate"
            if verified then
		    from_mail = Session("Store_Email")
            if from_mail="" or isnull(from_mail) then
                from_mail = Store_Email
            end if

            send_cont = Gift_buyer_email_body
            if instr(send_cont,"%GIFT_CODE%")>0 then
                send_cont=replace(send_cont,"%GIFT_CODE%",gift_code)
            else
                send_cont = send_cont&vbcrlf&"Your gift certificate code is: "&gift_code
            end if
            send_cont=replace(send_cont,"%AMOUNT%",Currency_Format_Function(gift_amount))

            send_cont = send_cont&vbcrlf&u_d_2

            call Send_Mail_Html(from_mail, u_d_1, Gift_email_Subject, send_cont)
            send_cont="Here is a copy of the message sent to your gift certificate recipient.<BR>"&send_cont
            call Send_Mail_Html(from_mail, shipemail, Gift_email_Subject, send_cont)
		end if
    next

                    '*********************
                    sql_update = sql_update & sql_gift
                    sub_write_log sql_update
                end if
                if quantity_control<>0 then
                    sub_write_log "stock check"
                    quantity_control_number = mydatatrans(myfieldstrans("quantity_control_number"),rowcountertrans)
                    quantity_in_stock = mydatatrans(myfieldstrans("quantity_in_stock"),rowcountertrans)
                    sql_stock=""
    if (Verified AND Inventory_Reduce <>0) OR (not(verified) AND Unverified_Reduce <>0) then
         sub_write_log "update stock"
        sql_stock=sql_stock&"exec wsp_item_qty_update "&store_id&","&item_id&","&oid&","&quantity&";"
        sub_write_log sql_stock
        if Quantity_in_Stock>0 and (Quantity_in_Stock-Quantity <= Quantity_Control_Number or Quantity_in_Stock-Quantity <= 0) then
            Stock_Updated=-1
            subject="Out of stock items in your store"
            body = "Dear store owner, the StoreSecured system has sent you this email to alert you that the item with sku:"&item_sku&", and name:"&item_name&", has reached it's minimum sell quantity or quantity in stock has reached 0."
            call Send_Mail(Store_Email,Store_Email,subject,body)
        end if
    end if

                    sql_update = sql_update & sql_stock
                    sub_write_log sql_update
                end if
            next
        end if

    next
end if
if (Verified AND Inventory_Reduce <>0) OR (not(verified) AND Unverified_Reduce <>0) then
	sql_update=sql_update&"exec wsp_purchase_stock_updated "&Store_id&","&OID&";"
	sub_write_log sql_update
end if

if Coupon_Id<>"" then
    sql_update = sql_update & "exec wsp_coupon_qty_update "&Store_id&",'"&Coupon_id&"';"&vbcrlf
    sub_write_log sql_update
end if
if Giftcert_Id<>"" and Giftcert_Amount>0 then
    sql_update = sql_update & "exec wsp_giftcert_redeem "&Store_id&","&giftcert_amount&",'"&giftcert_id&"';"&vbcrlf
    sub_write_log sql_update
end if

sql_update = sql_update&"exec wsp_purchase_update "&Store_id&","&OID&","&Shopper_Id&","&Verified&",'"&cust_po&"';"&vbcrlf
sql_update = sql_update&"exec wsp_customer_update "&Store_id&","&OID&","&cid&","&GGrand_Total&",'"&Orders_Dept&"',"&RewardsAdd-Rewards&";"&vbcrlf


on error goto 0
sub_write_log sql_update
session("sql")=sql_update
conn_store.Execute sql_update

'***call affiliate_notify***
sql_select = "select SA.Email from Store_Affiliates SA WITH (NOLOCK), Store_Purchases SP WITH (NOLOCK) where SA.store_id="&store_id&" and SA.Code=SP.CAME_FROM AND SP.oid="&oid&" and SA.Email_Notification<>0"
	sub_write_log sql_select
	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

	if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		Email_Addr = mydata(myfields("email"),rowcounter)
		Email_Subject = "New order was placed at " &Store_Name
		Email_Body = "New order was placed at "&Store_Name&vbcrlf
		Email_Body = Email_Body&"Order is is "&oid&vbcrlf
		Email_Body = Email_Body&"Grand Total "&formatcurrency(grand_total)&vbcrlf&vbcrlf
		Email_Body = Email_Body&Store_Name&"You are being alerted of this because the order was placed via a link on your site."
		Send_Mail Store_email,Email_Addr,Email_Subject,Email_Body
	Next
	End if

	set myfields = Nothing



'****************************************end of checkout_out_payment_simulation****************
                 sql_update3="update Store_Transactions set Transaction_Processed=1 where Store_ID=" & store_id & " and Shopper_ID='" & Shopper_ID & "'"
         sub_write_log sql_update3
         conn_store.Execute sql_update3


                 sql_update = "update store_Purchases set Verified=1,Purchase_Completed=1,Transaction_Type=0 where store_id="&store_id&" and Verified_Ref='"&GoogleOrderNumber&"'"

                 sub_write_log sql_update
                 conn_store.execute sql_update




			' MyOrder.ProcessOrder GoogleOrderNumber
			' MyOrder.SendOrderCommand

			' MyOrder.AuthorizeOrder GoogleOrderNumber
			' MyOrder.SendOrderCommand



			' MyOrder.CancelOrder GoogleOrderNumber, Reason, Comment
			' MyOrder.SendOrderCommand
        
		Case "CHARGING"


		Case "CHARGED"
		sql_update = "update store_Purchases set Verified=1 where store_id="&store_id&" and Verified_Ref='"&GoogleOrderNumber&"'"
		conn_store.execute sql_update
                sub_write_log "Order charged"
			' MyOrder.AddTrackingData GoogleOrderNumber, Carrier, TrackingNumber
			' MyOrder.SendOrderCommand

			' MyOrder.DeliverOrder GoogleOrderNumber, Carrier, TrackingNumber, True
			' MyOrder.SendOrderCommand

			' MyOrder.RefundOrder GoogleOrderNumber, Reason, Amount, Comment
			' MyOrder.SendOrderCommand

		Case "PAYMENT_DECLINED"
			' MyOrder.CancelOrder GoogleOrderNumber, Reason, Comment
			' MyOrder.SendOrderCommand

		Case "CANCELLED"
                 sub_write_log "Order Canceled"
		Case "CANCELLED_BY_GOOGLE"
			' MyOrder.SendBuyerMessage GoogleOrderNumber, Mesage, true
			' MyOrder.SendOrderCommand

		Case Else
	End Select

	Select Case MyOrderState.NewFulfillmentOrderState
		Case "NEW"

		Case "PROCESSING"

		Case "DELIVERED"
			' MyOrder.ArchiveOrder GoogleOrderNumber
			' MyOrder.SendOrderCommand

		Case "WILL_NOT_DELIVER"

		Case Else
	End Select


	SendAck
End Sub

' Handle a risk-information-notification message
Sub ProcessRiskInformationNotification(ResponseXml)
        sub_write_log "in process risk info"
        sub_write_log ResponseXml

	Dim MyRiskInfo
	Set MyRiskInfo = New RiskInformationNotification
	MyRiskInfo.ParseNotification ResponseXml

	' ******** How to access RiskInformationNotification variables *********
	' ****************************** BEGIN *********************************

	'DoSomethingWith MyRiskInfo.TimeStamp
	Verified_Ref = MyRiskInfo.GoogleOrderNumber
	'DoSomethingWith MyRiskInfo.IpAddress
	'DoSomethingWith MyRiskInfo.EligibleForProtection
	Avs_result = MyRiskInfo.AvsResponse
	Card_Code_verification = MyRiskInfo.CvnResponse
	CardNumber = EnCrypt(MyRiskInfo.PartialCC)
	sub_write_log "CardNumber : " & CardNumber
	'DoSomethingWith MyRiskInfo.BuyerAccountAge

        sql_update = "update store_purchases set avs_result='"&AVS_Result&"',Card_Code_verification='"&Card_Code_verification&"',CardNumber='"&CardNumber&"' where store_id="&store_id&" and Verified_Ref='"&verified_ref&"'"
	sub_write_log sql_update
        conn_store.execute sql_update

        'Dim BillingAddress
	'Set BillingAddress = MyRiskInfo.BillingAddress
	'DoSomethingWith BillingAddress.ID
	'DoSomethingWith BillingAddress.ContactName
	'DoSomethingWith BillingAddress.CompanyName
	'DoSomethingWith BillingAddress.Address1
	'DoSomethingWith BillingAddress.Address2
	'DoSomethingWith BillingAddress.City
	'DoSomethingWith BillingAddress.Region
	'DoSomethingWith BillingAddress.PostalCode
	'DoSomethingWith BillingAddress.CountryCode
	'DoSomethingWith BillingAddress.Email
	'DoSomethingWith BillingAddress.Phone
	'DoSomethingWith BillingAddress.Fax
	'Set BillingAddress = Nothing

	' ****************************** END *********************************

	SendAck
End Sub
'**********************
  ' Acknowledge the notification by responding to Google Checkout 
' with <notification-acknowledgment>.
Public Sub SendAck()
	Dim XmlAcknowledgment
	XmlAcknowledgment = "<?xml version=""1.0"" encoding=""UTF-8""?><notification-acknowledgment xmlns=""http://checkout.google.com/schema/2"" />"

	LogMessage XmlAcknowledgment
	Response.Write XmlAcknowledgment

End Sub

Set MyOrder = Nothing



%>

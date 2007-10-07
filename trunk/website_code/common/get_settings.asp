<%
sql_select = "exec wsp_settings_select "&store_id&";"
rs_store.open sql_select,conn_store,1,1
if not rs_store.eof then
    Additional_Storage = rs_Store("Additional_Storage")
    Affiliate_amount=rs_Store("Affiliate_amount")
	Affiliate_cookie=rs_Store("Affiliate_cookie")
	Affiliate_id=rs_Store("Affiliate_id")
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
    Custom_Amount= rs_Store("Custom_Amount")
    Custom_Description= rs_Store("Custom_Description")
    Department_Layout=rs_store("Department_Layout")
	dept_display=rs_Store("dept_display")
    dept_rows=rs_Store("dept_rows")
	Detail_NextPrev = rs_Store("Detail_NextPrev")
	Dict_Accessory = rs_Store("Dict_Accessory")
    Email_Me= rs_Store("Email_Me")
    Enable_affiliates=rs_Store("Enable_affiliates")
    Enable_Coupon = rs_store("Enable_Coupon")
	Enable_Ip_Tracking=rs_Store("Enable_Ip_Tracking")
	Enable_Rewards=rs_Store("Enable_Rewards")
	Enable_RMA=rs_Store("Enable_RMA")
	ExpressCheckout =rs_Store("ExpressCheckout")
	Gift= rs_Store("Gift")
	Gift_Message=rs_Store("Gift_Message")
	Gift_Service=rs_Store("Gift_Service")
	Gift_Wrapping_surcharge=rs_Store("Gift_Wrapping_surcharge")
	GoogleCheckout= rs_Store("GoogleCheckout")
	GoogleCheckout_ButtonStyle = rs_store("GoogleCheckout_ButtonStyle")
	Handling_Fee=rs_Store("Handling_Fee")
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
	Newsletter_receive=rs_store("Newsletter_receive")
	No_Ecommerce= rs_Store("No_Ecommerce")
	No_Login= rs_Store("No_Login")
    Overdue_Payment=rs_Store("Overdue_Payment")
	Paypal_Express=rs_Store("Paypal_Express")
	Real_Time_Processor=rs_Store("Real_Time_Processor")
	Reload_Attr=rs_Store("Reload_Attr")
	Reseller_ID=rs_store("Reseller_ID")
	SiteResellerID=rs_store("Reseller_ID")
	Rewards_Minimum=rs_Store("Rewards_Minimum")
	Rewards_Percent=rs_Store("Rewards_Percent")
	Save_Cart=rs_Store("Save_Cart")
	Screen_affiliates=rs_Store("Screen_affiliates")
	Secure_Name= rs_Store("Secure_Name")
    Service_Fee=rs_Store("Service_Fee")
    Service_Type=rs_Store("Service_Type")
	Ship_Multi= rs_Store("Ship_Multi")
	Shipping_Class= rs_Store("Shipping_Classes")
	Shipping_Classes=rs_store("Shipping_Classes")
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
    Special_Notes=rs_store("Special_Notes")
    StartCID=rs_Store("StartCID")
	StartOID=rs_Store("StartOID")
	StartTID=rs_Store("StartTID")
	Store_active=rs_Store("Store_active")
	Store_address1= rs_Store("Store_address1")
	Store_address2= rs_Store("Store_address2")
	Store_Cancel = rs_Store("Store_Cancel")
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
    Store_Id = rs_Store("Store_Id")
    Store_name= rs_Store("Store_name")
	Store_Phone= rs_Store("Store_Phone")
	Store_Public=rs_Store("Store_Public")
    Store_State= rs_Store("Store_State")
	Store_Trial = rs_Store("Store_Trial")
	Store_Zip= rs_Store("Store_Zip")
	Subdept_location=rs_store("Subdept_location")
	Sys_Created= rs_Store("Sys_Created")
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

    Show_SubDept = rs_Store("subdept_location")
    Show_ItemNav = rs_Store("detail_nextprev")
    Store_address= Store_address1
end if   
rs_store.close
sLineBreak=""

Site_name = fn_url_rewrite(Site_Name)
Secure_Name = fn_url_rewrite(Secure_Name)
Site_name_Orig = Site_name

Secure_Site="https://"&Secure_Name&"/"
if Use_Domain_Name then
    Site_name="http://"&Store_domain&"/"
else
    Site_name="http://"&Site_name&"/"
end if
if show_shipping="" then
   show_shipping=True
end if
if isNull(Shipping) or Shipping="" then
    Shipping = "Shipping"
end if
if isNull(Shopping_Cart) or Shopping_Cart="" then
    Shopping_Cart = "Shopping Cart"
end if
if isNull(Special_notes) or Special_notes="" then
    Special_notes = "Special Notes"
end if
if isNull(Coupon) or Coupon="" then
    Coupon = "Coupon/Gift Certificate"
end if
if isNull(Email_Me) or Email_Me="" then
    Email_Me = "Email me special offers"
end if
if isNull(Dict_Accessory) or Dict_Accessory="" then
    Dict_Accessory = "Accessories"
end if
if isNull(Top_Depts) then
    Top_Depts="Top"
end if
if item_f_rows = "" or isNull(item_f_rows) then
    item_f_rows=10
end if
if not isnumeric(sHandling_Fee) then
    sHandling_Fee = 0
end if
Shipping_Class=replace(Shipping_Class," ","")


%>

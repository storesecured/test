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
%>
<!-- #include virtual="include/header_noview.asp" -->
<!-- #include file="googleglobal.asp" -->
<!--#include virtual="include/cart_display.asp"-->

<%
server.scripttimeout = 300
'response.write(Session("Shopper_Id") & " " & Store_Id)
'response.end
isFull=fn_Is_Cart_Full()
if isFull Then
else
fn_error "Sorry your order can't be proceed because you cart is empty."
end if
Call sub_check_minimum()
sub_write_log "class before : " & Shipping_Class
on error goto 0
Dim MyCart
dim zipArray
dim stateArray
dim shippingArray
Set MyCart = New Cart

' Add items
Dim MyItem

iShip_Location_id=""

sql_update = sql_update&"exec wsp_purchase_create "&Store_Id&","&Shopper_Id&",0,1;"
conn_store.execute sql_update

  sql_select_max = "select max(OID) as max_id from store_purchases where store_id="&store_id &" and Shopper_ID="& Shopper_Id
	sub_write_log sql_select_max
	rs_Store.open sql_select_max,conn_store,1,1
	rs_Store.MoveFirst

	OID = rs_Store("max_id")
	sub_write_log "Order_ID=" & OID
	rs_Store.Close

sql_select = "select * from Store_Transactions where store_id="&store_id&" and shopper_id='"&shopper_id&"' and Transaction_Processed=0"
sub_write_log sql_select
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
FOR rowcounter= 0 TO myfields("rowcount")
    Set MyItem = New Item
    MyItem.Name = mydata(myfields("item_name"),rowcounter)
    MyItem.Description = ""
    MyItem.Quantity = mydata(myfields("quantity"),rowcounter)
    MyItem.UnitPrice = mydata(myfields("sale_price"),rowcounter)
    if iShip_Location_id<>"" and iShip_Location_id<>mydata(myfields("ship_location_id"),rowcounter) then
        fn_error "Google Checkout cannot be used for items shipping from multiple locations at this time."
    end if
    iShip_Location_id=mydata(myfields("ship_location_id"),rowcounter)
    MyCart.AddItem MyItem
    Set MyItem = Nothing
next

sub_write_log "URL="&Switch_Name&"include/google/callback2.asp?Shopper_Id="&Shopper_Id

' Set optional parameters
MyCart.CartExpiration = "2007-12-31T23:59:59"
'MyCart.MerchantPrivateData = "<session-id>"&shopper_id&"</session-id><cart-id>"&shopper_id&"</cart-id>"
'MyCart.MerchantCalculationsUrl = Secure_site&"include/google/callback.asp?Shopper_Id="&Shopper_Id
MyCart.MerchantCalculationsUrl = Secure_site&"include/google/callback2.asp?Shopper_Id="&Shopper_Id
MyCart.MerchantCalculatedTax = True

if Show_Coupon=1 then
MyCart.AcceptMerchantCoupons = True
MyCart.AcceptGiftCertificates = True
else
MyCart.AcceptMerchantCoupons = False
MyCart.AcceptGiftCertificates = False
end if

MyCart.EditCartUrl = Secure_site&"show_big_cart.asp?Shopper_Id="&Shopper_Id
MyCart.ContinueShoppingUrl = Secure_site&"recipiet.asp?Shopper_id="&Shopper_id
MyCart.RequestBuyerPhoneNumber = True
MyCart.RequestInitialAuthDetails = True
MyCart.PlatformID="137736699110011"
' Add a merchant-calculated-shipping method


sql_sel_items = "wsp_trans_totals_shipping "&Store_Id&",'"&Shopper_ID&"',1,"&iShip_Location_id&";"
sub_write_log sql_sel_items
rs_Store.open sql_sel_items, conn_store, 1, 1
Total_Quantity=FormatNumber(rs_store("tquantity"),2)
sub_write_log "Total_Quantity " & Total_Quantity
 if Total_Quantity=0 then
 Set MyShipping = New MerchantCalculatedShipping
    Set MyShippingRestrictions = New ShippingFilters
        rs_store.close 
        MyShipping.Name = "Free Shipping"
        MyShipping.Price = 0
        MyShippingRestrictions.AllowedWorldArea = True
        MyShippingRestrictions.AllowUsPoBox = False
        sub_write_log "Free shipping"
        MyCart.AddShipping MyShipping
        Set MyShipping = Nothing
        Set MyShippingRestrictions = Nothing

 else

Total_Weight=FormatNumber((rs_store("tweight")+Handling_Weight),2,,,0)
Total_Price=FormatNumber(rs_store("tprice"),2,,,0)
SFee=FormatNumber(rs_store("sfee"),2,,,0)
sum_handling=FormatNumber((rs_store("thandling")+sHandling_fee),2,,,0)


' Add shipping methods
' Note: You cannot mix merchant-calculated-shipping methods with other shipping
' methods (flat-rate-shipping or pickup) in the same cart.
Dim MyShipping, MyShippingRestrictions, MyAddressFilters
Redim shippingArray(100)

sql_select = "exec wsp_shipping_display "&store_id&","&Total_Weight&","&Total_Price&","&iShip_Location_id&",203,-1,'"&Shipping_Class&"';"
sub_write_log sql_select
set shippingfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,shippingdata,shippingfields,noRecords)
if noRecords<>0 then

sub_write_log " No shipping methods available"
 fn_error "Google Checkout can't be used when there is no shipping methods available"

 end if

FOR shippingrowcounter= 0 TO shippingfields("rowcount")
    sub_write_log " shipping methods available"
     kount=kount+1
    this_Shipping_id=shippingdata(shippingfields("Shipping_method_id"),shippingrowcounter)
    this_Shipping_Class=shippingdata(shippingfields("shippers_class"),shippingrowcounter)
    this_shipping_method_name=shippingdata(shippingfields("shipping_method_name"),shippingrowcounter)
    this_base_fee=shippingdata(shippingfields("base_fee"),shippingrowcounter)
    this_weight_fee=shippingdata(shippingfields("weight_fee"),shippingrowcounter)
    this_zip_start=shippingdata(shippingfields("zip_start"),shippingrowcounter)
    this_zip_end=shippingdata(shippingfields("zip_end"),shippingrowcounter)
    this_countries=shippingdata(shippingfields("countries"),shippingrowcounter)


    select case this_Shipping_Class
        case 1 ship_price=this_base_fee
        case 2 ship_price=this_base_fee+(this_weight_fee*total_weight)
        case 3 ship_price=this_base_fee+SFee
        case 4 ship_price=this_base_fee
        case 5 ship_price=this_base_fee*(Total_Price/100)
        case 6 ship_price=this_base_fee
    end select
    ship_price = ship_price+sum_Handling
    

      Set MyShipping = New MerchantCalculatedShipping
    Set MyShippingRestrictions = New ShippingFilters

    shippingArray(shippingrowcounter)=this_shipping_method_name
    if shippingrowcounter>0 then
    found=0
    for k=0 to shippingrowcounter-1
    if lcase(this_shipping_method_name)= lcase(shippingArray(k)) then
    found=1
    end if
    next
    if found=0 then
    MyShipping.Name = this_shipping_method_name
    else
    MyShipping.Name = this_Shipping_id&"-"&this_shipping_method_name
    end if
    else
    MyShipping.Name = this_shipping_method_name
    end if


    MyShipping.Price = ship_price
    pos=InStr(this_countries,"216")
    if pos<>0 then
    MyShippingRestrictions.AllowedWorldArea = True
    MyShippingRestrictions.AllowUsPoBox = False
    else
    sArray = Split(this_countries,",")
    for each country in sArray
 country_id=Trim(country)
 sql_sel_country= "exec wsp_country_code_by_ID "&country_id&";"
sub_write_log sql_sel_country
Set rs_Store_country = Server.CreateObject("ADODB.Recordset")
rs_Store_country.open sql_sel_country, conn_store, 1, 1
country_code=rs_Store_country("country_code")
rs_Store_country.close
if country_id=216 then
MyShippingRestrictions.AllowedWorldArea = True
MyShippingRestrictions.AllowUsPoBox = False
exit for

else
 call MyShippingRestrictions.AddAllowedPostalArea(country_code,"")
end if

  Next
  end if
  pos=InStr(this_countries,"216")
    if pos=0 then
    if this_zip_start=Null or this_zip_end=Null or this_zip_start="" or this_zip_end="" then
    sub_write_log "no zips"

    else
        if isNumeric(this_zip_start) and isNumeric(this_zip_end) and this_zip_start<this_zip_end then
           foundflag=0
           dim newzip_start
           kount = Len(this_zip_start)
           zeropos=InStrRev(this_zip_start,"0")
           if zeropos=kount then
           if Mid(this_zip_end,kount,1)="9" then
           foundflag=1

           newzip_start=""
           FOR k= 1 TO kount
           Mychar = Mid(this_zip_start,k,1)
           if Mychar="0" then
           Mychar2 = Mid(this_zip_end,k,1)
           if Mychar2="9" then
           newzip_start=newzip_start & "*"
           else
           newzip_start=newzip_start & Mychar
           end if
           else
           newzip_start=newzip_start & Mychar
           end if
           Next
           end if
           end if
          if foundflag=1 then
          MyShippingRestrictions.AllowedZipPatterns = Array(newzip_start)
          else
          sZips=""
            iIndex=0
            Redim zipArray(100000)
            FOR zip= this_zip_start TO this_zip_end
                while len(zip)<5
                    zip="0"&zip        
                wend
                zipArray(iIndex)=zip
                iIndex=iIndex+1
            next
            Redim preserve zipArray(iIndex-1)
    ' sub_write_log zipArray
 MyShippingRestrictions.AllowedZipPatterns = zipArray
        erase zipArray
        end if
        end if
    end if
    end if
   'Set MyShippingRestrictions = New ShippingFilters
'MyShippingRestrictions.AllowedStates = Array("CA", "AZ", "OR", "WA")
 MyShipping.AddShippingRestrictions MyShippingRestrictions


    MyCart.AddShipping MyShipping
    Set MyShipping = Nothing
     Set MyShippingRestrictions = Nothing
next
end if

' Define default tax rules
Dim MyDefaultTaxRule

sql_select = "Select * from Store_State_Tax where store_id="&store_id
sub_write_log sql_select
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
FOR rowcounter= 0 TO myfields("rowcount")
    Set MyDefaultTaxRule = New DefaultTaxRule
    MyDefaultTaxRule.ShippingTaxed = mydata(myfields("tax_shipping"),rowcounter)
    MyDefaultTaxRule.Rate = mydata(myfields("taxrate"),rowcounter)/100
    sStates=replace(mydata(myfields("state"),rowcounter)," ","")
    iIndex=0
    Redim stateArray(100)
    for each this_state in split(sStates,",")
        stateArray(iIndex)=this_state
        iIndex=iIndex+1
    next
    Redim preserve stateArray(iIndex-1)
    MyDefaultTaxRule.States = stateArray
    erase stateArray
    MyCart.AddDefaultTaxRule(MyDefaultTaxRule)
    Set MyDefaultTaxRule = Nothing
next

sql_select = "Select * from Store_Country_Tax where store_id="&store_id&" and (country like '%All Countries%' or country like '%United States%')"
sub_write_log sql_select
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
FOR rowcounter= 0 TO myfields("rowcount")
    Set MyDefaultTaxRule = New DefaultTaxRule
    MyDefaultTaxRule.ShippingTaxed = mydata(myfields("tax_shipping"),rowcounter)
    MyDefaultTaxRule.Rate = mydata(myfields("taxrate"),rowcounter)/100
    MyDefaultTaxRule.CountryArea = "ALL"
    MyCart.AddDefaultTaxRule(MyDefaultTaxRule)
    Set MyDefaultTaxRule = Nothing
next

sql_select = "Select * from Store_Zips where store_id="&store_id
sub_write_log sql_select
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
FOR rowcounter= 0 TO myfields("rowcount")
    Set MyDefaultTaxRule = New DefaultTaxRule
    this_zip_start = mydata(myfields("zip_start"),rowcounter)
    this_zip_end = mydata(myfields("zip_end"),rowcounter)
    MyDefaultTaxRule.ShippingTaxed = mydata(myfields("tax_shipping"),rowcounter)
    MyDefaultTaxRule.Rate = mydata(myfields("taxrate"),rowcounter)/100
    if this_zip_start=Null or this_zip_end=Null then
    else
        if isNumeric(this_zip_start) and isNumeric(this_zip_end) and this_zip_start<this_zip_end then
            foundflag=0

           kount = Len(this_zip_start)
           zeropos=InStrRev(this_zip_start,"0")
           if zeropos=kount then
           if Mid(this_zip_end,kount,1)="9" then
           foundflag=1

           newzip_start=""
           FOR k= 1 TO kount
           Mychar = Mid(this_zip_start,k,1)
           if Mychar="0" then
           Mychar2 = Mid(this_zip_end,k,1)
           if Mychar2="9" then
           newzip_start=newzip_start & "*"
           else
           newzip_start=newzip_start & Mychar
           end if
           else
           newzip_start=newzip_start & Mychar
           end if
           Next
           end if
           end if
          if foundflag=1 then
          MyDefaultTaxRule.ZipPatterns = Array(newzip_start)
          else
            sZips=""
            iIndex=0
            iArray=this_zip_end-this_zip_start
            if iArray>500 then
            	iArray=500
            end if
            redim zipArray(iArray)
            FOR zip= this_zip_start TO this_zip_end
                while len(zip)<5
                    zip="0"&zip
                wend
        
                 zipArray(iIndex)=zip
               if iIndex=500 then
               	exit for
               end if
                iIndex=iIndex+1
            next
            'Redim preserve zipArray(iIndex-1)

            MyDefaultTaxRule.ZipPatterns = zipArray
            erase zipArray
        end if
    end if
   end if 


    MyCart.AddDefaultTaxRule(MyDefaultTaxRule)
    Set MyDefaultTaxRule = Nothing
next


' Display <checkout-shopping-cart> XML
'fn_print_debug MyCart.GetXml

' Diagnose <checkout-shopping-cart> XML
' MyCart.DiagnoseXml

' Post Cart to Google Checkout 


MyCart.MerchantPrivateData = "<session-id>"&shopper_id&"</session-id><cart-id>"&OID&"</cart-id>"
MyCart.PostCartToGoogle

Set MyCart = Nothing 

%>

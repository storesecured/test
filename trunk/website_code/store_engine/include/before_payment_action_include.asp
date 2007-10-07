<%
function fn_checkCustomFields (sAppend)
    Custom_Fields=""
    sError=""
    sql_select = "select custom_field_name,custom_field_type, field_required from store_custom_fields WITH (NOLOCK) where Store_id ="&Store_id&" order by view_order,custom_field_name"
    fn_print_debug sql_select
    set myfields=server.createobject("scripting.dictionary")
    Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
    if noRecords = 0 then
        Custom_Fields = ""
        FOR rowcounter= 0 TO myfields("rowcount")
            Custom_Field_Name = mydata(myfields("custom_field_name"),rowcounter)
            Custom_Field_Type = mydata(myfields("custom_field_type"),rowcounter)
            Field_Required = mydata(myfields("field_required"),rowcounter)
            Custom_Field_Value=request.form(Custom_Field_Name&sAppend)
            if Field_Required and Custom_Field_Value="" then
                if Custom_Field_Type=4 then
                    sError=sError& "<LI>You must check the box labeled "&Custom_Field_Name&" to proceed."
                else
                    sError=sError& "<LI>"&Custom_Field_Name&" is required."
                end if
            elseif Field_Required and Custom_Field_Type=3 then
		  	if instr(lcase(Custom_Field_Value),"select")>0 then
		  		sError=sError& "<LI>"&Custom_Field_Name&" is required."
            	end if
		  end if
            Custom_Fields = Custom_Fields & Custom_Field_Name & "="&Custom_Field_Value&"<BR>"
        Next
        if sError<>"" then
            fn_error "<UL>"&sError&"</UL>"
        end if
        Custom_Fields = checkStringforQ(Custom_Fields)
    end if
    fn_checkCustomFields=Custom_Fields
end function

'CODE FOR CREATE AN ORDER
'FIRST ORDER WILL BE THE MASTER ONE, ALL OTHER (TO OTHER SHIPPING
'ADDRESSES) WILL BE LINKED TO THE MASTER
sub createOrder (ship_location_id,shipAddr)
    'GET SHIPPING INFO
	sql_select_cust = "exec wsp_customer_lookup "&store_id&","&cid&","&shipAddr&";"
    fn_print_debug sql_select_cust
    session("sql")=sql_select_cust
    rs_Store.open sql_select_cust,conn_store,1,1
	rs_Store.MoveFirst

	Tax_Exempt=rs_Store("Tax_Exempt")
	ShipCompany=checkStringForQ(rs_Store("Company"))
	ShipFirstname=rs_Store("First_name")
	ShipLastname=checkStringForQ(rs_Store("Last_name"))
	ShipAddress1=checkStringForQ(rs_Store("Address1"))
	ShipAddress2=checkStringForQ(rs_Store("Address2"))
	ShipCity=checkStringForQ(rs_Store("City"))
	Shipzip=checkStringForQ(rs_Store("zip"))
	ShipState=checkStringForQ(rs_Store("State"))
	ShipCountry=checkStringForQ(rs_Store("Country"))
	ShipPhone=checkStringForQ(rs_Store("Phone"))
	ShipFax=checkStringForQ(rs_Store("Fax"))
	ShipEMail=checkStringForQ(rs_Store("EMail"))
	ShipResidential=cint(rs_Store("Is_Residential"))
	rs_Store.Close

	'SET ORDER VARIABLED
	Purchase_Date = now()
	Verified = 0
	
	sql_totals = "exec wsp_cart_totals "&store_id&",'"&shopper_id&"',"&shipAddr&","&ship_location_id&";"
    fn_print_debug sql_totals
    set rs_store = conn_store.execute(sql_totals)
    Total=rs_store("Total")
    Recurring_total=rs_store("Recurring_Total")
    Recurring_Days=rs_store("Recurring_days")
    Weight=rs_store("Weight")
    Wholesale_Price_Total=rs_store("Wholesale_Price_Total")
    rs_store.close
    
    sAppendString="_"&ship_location_id&"_"&shipAddr

    sShippingPrice=Request.Form("Shipping_Method_name_Price"&sAppendString)
    if sShippingPrice = "" and Show_shipping then
        fn_error "Please choose a shipping method from the dropdown on the previous page before continuing."
    else
        Shipping_Method_name_Price_array =	Split(sShippingPrice,"|")
	    if sShippingPrice="" then
		    fNo =1
	    end if
	    if fNo <> 1 then
	       shipping_method_name = Shipping_Method_name_Price_array(0)
	       prices = Shipping_Method_name_Price_array(1)
	    end if
    end if
    shipping_method_name=left(checkstringforQ(shipping_method_name),50)
  
    if Request.Form("Payment_Method") = "Free" and cint(prices)>0 then
        fn_error "You cannot select Free as the payment method if your shipping total is over $0"
    end if
    if not isnumeric(Is_Gifts) or isNull(Is_Gifts) or Is_Gifts="" then
      Is_Gifts=0
	end if
	if not isnumeric(Gift_Wrappings) or isNull(Gift_Wrappings) or Gift_Wrappings="" then
      Gift_Wrappings=0
	end if
	if Gift_Wrappings="-1" then
	    Gift_Wrapping_amount=Gift_Wrapping_surcharge
	else
	    Gift_Wrapping_amount=0
	end if
    
    sNewTotal=This_Total
    if coupon_amount>0 then
        if cdbl(coupon_amount)>cdbl(sNewTotal) then
            coupon_amount=sNewTotal
        end if
        sNewTotal=sNewTotal-coupon_amount
    end if
    if Reward_Amounts>0 then
        if cdbl(Reward_Amounts)>cdbl(sNewTotal) then
            Reward_Amounts=sNewTotal
        end if
        sNewTotal=sNewTotal-Reward_Amounts
    end if
    discount_nontax=coupon_amount+Reward_Amounts
    Tax = fn_Calc_Tax(Tax_Exempt,shipzip, shipstate,shipcountry,shipAddr,ship_location_id,discount_nontax,Prices,"single_total")
	if Recurring_Grand_Total>0 then
	    Recurring_Tax = fn_Calc_Tax(Tax_Exempt,shipzip, shipstate,shipcountry,shipAddr,ship_location_id,0,0,"recurring_total")
    else
        Recurring_Tax=0
    end if
    Recurring_Grand_Total = Recurring_Total + Recurring_Tax
    
    This_Total = FormatNumber(This_Total,2,0,0,0)
    Prices = FormatNumber(Prices,2,0,0,0)
    coupon_amount = FormatNumber(coupon_amount,2,0,0,0)
    Tax = FormatNumber(Tax,2,0,0,0)
    Recurring_Tax = FormatNumber(Recurring_Tax,2,0,0,0)
    
    sNewTotal=cdbl(sNewTotal)+cdbl(Prices)+cdbl(Tax)+cdbl(Gift_Wrapping_amount)
    if giftcert_amount>0 then
        if cdbl(giftcert_amount)>cdbl(sNewTotal) then
            giftcert_amount=sNewTotal
        end if
        sNewTotal=sNewTotal-giftcert_amount
    end if

    Grand_Total = CDbl(Total)
    Grand_Total = Grand_Total + CDbl(Tax)
    Grand_Total = Grand_Total + CDbl(prices)
    Grand_Total = Grand_Total - CDbl(coupon_amount)
    Grand_Total = Grand_Total - CDbl(giftcert_amount)
    Grand_Total = Grand_Total - CDbl(Reward_Amounts)
    Grand_Total = Grand_Total + CDbl(Gift_Wrapping_amount)
    
    'IF GRAND_TOTAL IS NULL OR LESS THEN ZERO, SET TO DEFAULT VALUE (0)
	if IsNull(Grand_Total) OR Grand_Total <= 0 then
		Grand_Total = 0
        Payment_Method="Free"  
	end if
    'GET A NEW ORDER ID
	sql_select_max = "select max(OID)+1 as max_id from store_purchases where store_id="&store_id
	fn_print_debug sql_select_max
	rs_Store.open sql_select_max,conn_store,1,1
	rs_Store.MoveFirst

	moid = rs_Store("max_id")
	rs_Store.Close
	if IsNull(moid) then
		moid=1
	end if

	if moid<StartOID then
		moid = StartOID
	end if
	
	'INSERT ORDER INTO THE DATABASE
	if (oid=-1) then
		oid = moid
		masteroid="Null"
    else
		sub_oid = moid
		Coupon_Id=""
		coupon_amount=0
		giftcert_id=""
		giftcert_amount=0
		Reward_Amounts=0
		
		masteroid=oid
		moid=sub_oid
	end if
	sql_update = "exec wsp_purchase_insert "&CCID&","&moid&_
	    ","&shipAddr&","&Verified&","&Store_id&_
	    ",'"&Shopper_ID&"',"&Cid&",'"&ShipLastname&_
	    "','"&ShipFirstname&"','"&ShipCompany&"','"&ShipAddress1&_
	    "','"&ShipAddress2&"','"&ShipCity&"','"&ShipZip&_
	    "','"&ShipState&"','"&ShipCountry&"','"&ShipPhone&_
	    "','"&ShipFax&"','"&ShipEMail&"',"&Tax&","&prices&_
	    ","&Total&","&Wholesale_Price_Total&","&Grand_Total&_
	    ",'"&Coupon_Id&"',"&coupon_amount&",'"&giftcert_id&_
	    "',"&giftcert_amount&","&Reward_Amounts&",'"&Payment_Method&_
	    "','"&Shipping_Method_name&"','"&cust_notess&_
	    "','"&custom_fields&"','"&Cust_PO&"',"&Is_Gifts&_
	    ",'"&Gift_Messages&"',"&Gift_Wrappings&_
	    ","&Gift_Wrapping_amount&","&Recurring_Total&_
	    ","&Recurring_Days&","&Recurring_Tax&_
	    ","&Recurring_Grand_Total&","&ship_location_id&_
	    ","&ShipResidential&","&masteroid&";"
		    
	fn_print_debug sql_update
	session("sql")=sql_update
	conn_store.Execute sql_update
end sub
 
' ================================================================
' RETURNS TOTAL TAX AMOUNT
function fn_Calc_Tax(Tax_Exempt,shipzip,shipstate,shipcountry,shipto,location_id,discount_untax,ship_taxable,sType)
    if Tax_Exempt <> 0 then
        fn_Calc_Tax = 0
        exit function
    end if
  
    Tax_zipcode=left(shipzip,5)
    if not isNumeric(tax_zipcode) or instr(tax_zipcode,",")>0 or instr(tax_zipcode,"-")>0 or instr(tax_zipcode,".")>0 then
        'this does not appear to be a US zipcode so set it to something numeric
        tax_zipcode=-1
    end if
    ' SELECT THE TAX RATE
    sql_tax_rate = "exec wsp_taxable_rates "&store_id&","&tax_zipcode&",'"&shipstate&"','"&shipcountry&"';"
    set taxfields=server.createobject("scripting.dictionary")
    fn_print_debug sql_tax_rate
    Call DataGetrows(conn_store,sql_tax_rate,taxdata,taxfields,noRecords)
    
    sCalc_Tax=0
    Discount_Tax=0
    Ship_Tax=0
    
    if noRecords=0 then
        sql_tax="exec wsp_taxable_items "&store_id&",'"&Shopper_Id&"',"&shipto&","&location_id&";"
        set taxfields1=server.createobject("scripting.dictionary")
        fn_print_debug sql_tax
        Call DataGetrows(conn_store,sql_tax,taxdata1,taxfields1,noRecords)
        
        FOR taxrowcounter= 0 TO taxfields("rowcount")
		    Department_Ids = taxdata(taxfields("department_ids"),taxrowcounter)
            TaxRate = taxdata(taxfields("taxrate"),taxrowcounter)/100
            Tax_Shipping = taxdata(taxfields("tax_shipping"),taxrowcounter)
            Collection = Cstr(Department_Ids)
            FOR taxrowcounter1= 0 TO taxfields1("rowcount")
                one_item = Cstr(taxdata1(taxfields1("department_id"),taxrowcounter1))
                if Is_In_Collection(Collection,"-1",",") or Is_In_Collection(Collection,one_item,",") then
                    sCalc_Tax = sCalc_Tax + (taxdata1(taxfields1(sType),taxrowcounter1)*TaxRate)
                end if
            Next

	        if discount_untax>0 then
	        	fn_print_debug "discount_untax="&discount_untax
	            Discount_Tax=Discount_Tax+(discount_untax*TaxRate)
	        end if
	        if ship_taxable>0 and Tax_Shipping then
		   	fn_print_debug "Tax_Shipping="&Tax_Shipping
                Ship_Tax = Ship_Tax+(ship_taxable*TaxRate)
		    end if
        Next
    end if
    set taxfields = Nothing
	set taxfields1 = Nothing
	
	if Discount_Tax>sCalc_Tax then
	    Discount_Tax=sCalc_Tax
	end if
	sCalc_Tax=sCalc_Tax-Discount_Tax+Ship_Tax

    if sCalc_Tax < 0 then
		sCalc_Tax = 0
	end if
	
    fn_calc_Tax=sCalc_Tax
End function

' ================================================================ 
' COMPUTE PROMOTION DISCOUNT
%>

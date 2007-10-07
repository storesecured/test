<%
function fn_Calc_Promotion_Discount()
    sql_promo_date_total_check = "exec wsp_promotion_get "&Store_id&","&Shopper_id&",'"&Groups&"';"
	fn_print_debug sql_promo_date_total_check
	sql_apply_promo=""
	set promofields=server.createobject("scripting.dictionary")
    Call DataGetrows(conn_store,sql_promo_date_total_check,promodata,promofields,noPromoRecords)
	if noPromoRecords=0 then
	    FOR promorowcounter= 0 TO promofields("rowcount")
	    	fn_print_debug "in promotion rowcounter"
	        free_items_ids = promodata(promofields("free_items_ids"),promorowcounter)
		    Discounted_items_ids = promodata(promofields("discounted_items_ids"),promorowcounter)
		    Discounted_items_Amount = promodata(promofields("discounted_items_amount"),promorowcounter)
		    Is_Exclusion = cint(promodata(promofields("is_exclusion"),promorowcounter))
		    Promotion_id = promodata(promofields("promotion_id"),promorowcounter)
		    Discount = (100-Discounted_items_Amount)/100
		    sql_apply_promo = sql_apply_promo&"exec wsp_promotion_apply "&Store_id&","&oid&","&Shopper_Id&","&promotion_id&","&Is_Exclusion&","&Discount&",'"&Discounted_items_ids&"','"&free_items_ids&"';"
		next    
	end if
	set promofields = Nothing
    fn_Calc_Promotion_Discount=sql_apply_promo
End function

function fn_Calc_Gift_Certificate(Coupon_id)
	Set rs_coup = Server.CreateObject("ADODB.Recordset")
	sql_coupon = "exec wsp_giftcert_valid "&store_id&","&Shopper_id&",'"&coupon_id&"';"
    fn_print_debug sql_coupon
    rs_coup = conn_store.execute(sql_coupon)
	on error resume next
    Current_Amount = rs_coup("Current_Amount")
    if err.number<>0 then
        fn_error "Invalid gift certificate"
    end if
    fn_Calc_Gift_Certificate=Current_Amount
	on error goto 0
	set rs_coup = Nothing
	sql_update = "exec wsp_giftcert_apply "&Store_Id&","&Shopper_Id&",'"&coupon_id&"',"&fn_Calc_Gift_Certificate&";"
    fn_print_debug sql_update
    conn_store.execute sql_update
end function

' ================================================================ 
' COMPUTE THE AMOUNT OF DISCOUNT RESULTED FROM A COUPON
Function fn_Calc_Coupon_Discount(Coupon_id)
	sql_coupon = "exec wsp_coupon_valid "&store_id&","&Shopper_id&",'"&coupon_id&"',"&cid&";"
    fn_print_debug sql_coupon
    rs_coup = conn_store.execute(sql_coupon)
	on error resume next
    Coupon_Type = rs_coup("Coupon_Type")
    if err.number<>0 then
        fn_error "Coupon code entered is not valid or reached it maximum usage limit."
    end if
    Coupon_Amount = rs_coup("Coupon_Amount")
    Discounted_items_ids=rs_coup("Discounted_items_ids")
    Is_Exclusion=rs_coup("Is_Exclusion")
    Coupon_total=cdbl(rs_coup("Total"))
    
    fn_print_debug "Coupon_Amount="&Coupon_Amount
    
    if Discounted_items_ids = "" then
        coupon_price = cdbl(Coupon_total)
    else
        sql_coupon = "exec wsp_coupon_amount "&Store_Id&","&Shopper_id&","&cint(is_exclusion)&",'"&discounted_items_ids&"';"
        fn_print_debug sql_coupon
        set rs_store = conn_store.execute(sql_coupon)
        coupon_price = rs_store("Coupon_total")
    end if
    
    fn_print_debug "coupon_price="&coupon_price
  
    if isNull(coupon_price) then
        fn_error "The coupon code you have entered is not valid for any of the items you have purchased."
    else
        coupon_price=cdbl(coupon_price)
    end if
    

    if Coupon_Type = 0 then
	    'PERCENT FROM TOTAL
	    fn_Calc_Coupon_Discount = (Coupon_Amount/100)*coupon_price
    Else
	    'FLAT TOTAL
	    fn_Calc_Coupon_Discount = Coupon_Amount
    End if
    
    sql_update = "exec wsp_coupon_apply "&Store_Id&","&Shopper_Id&",'"&coupon_id&"',"&fn_Calc_Coupon_Discount&";"
    fn_print_debug sql_update
    conn_store.execute sql_update
End Function 

'DISPLAYS A DROPDOWN MENU CONTAINING THE ADDRESS BOOK
function fn_displayAddressBook(bookName,myShipto)
	sql_select_addrs = "exec wsp_customer_lookup_no "&store_id&","&cid&",0;"
	fn_print_debug sql_select_addrs
	set myfieldsaddr=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select_addrs,mydataaddr,myfieldsaddr,noRecordsaddr)
     
	sAddrBookText = ("<select name='"&bookName&"' onchange=""JavaScript:changeAddressBookSelection();"">"&vbcrlf)
	if noRecordsaddr = 0 then
	    FOR rowcounteraddr= 0 TO myfieldsaddr("rowcount")
	        sFirstName=mydataaddr(myfieldsaddr("first_name"),rowcounteraddr)
	        sLastName=mydataaddr(myfieldsaddr("last_name"),rowcounteraddr)
		    sAddr1=mydataaddr(myfieldsaddr("address1"),rowcounteraddr)
		    sRecordType=mydataaddr(myfieldsaddr("record_type"),rowcounteraddr)
		    sAddrBookText = sAddrBookText & ("<option value='"&sRecordType&"'")
		    if cstr(myShipto)=cstr(sRecordType) then
			    sAddrBookText = sAddrBookText & ("selected")
		    End If
		    sAddrBookText = sAddrBookText & (">"&sFirstName&_
		        "&nbsp;"&sLastName&"&nbsp;-&nbsp;"&sAddr1&"</option>"&vbcrlf)
        Next
	    sAddrBookText = sAddrBookText & ("<option value='New'>Add New</option>"&vbcrlf)
    End if
    set myfieldsaddr = Nothing

    sAddrBookText = sAddrBookText & ("</select>"&vbcrlf)
    fn_displayAddressBook=sAddrBookText
	        
end function

function fn_get_custom_fields (sAppendString)
    sql_select = "select custom_field_name,custom_field_type,custom_field_values,field_required from store_custom_fields WITH (NOLOCK) where Store_id ="&Store_id&" order by view_order,custom_field_name"
    fn_print_debug sql_select
    set myfields=server.createobject("scripting.dictionary")
    Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
    if noRecords = 0 then
         FOR rowcounter= 0 TO myfields("rowcount")
            Custom_Field_Name = mydata(myfields("custom_field_name"),rowcounter)
            Custom_Field_Name2 = Custom_Field_Name&sAppendString
            Custom_Field_Type = mydata(myfields("custom_field_type"),rowcounter)
            Custom_Field_Values = mydata(myfields("custom_field_values"),rowcounter)
            Field_Required = mydata(myfields("field_required"),rowcounter)
            if Custom_Field_Name <> "" then
                sCustFldText = sCustFldText & ("<tr><td class='normal' colspan=4><b>"&_
                    Custom_Field_Name&"</b>&nbsp;</td><td class='normal'>")
                if Custom_Field_Type = 1 then
                    sCustFldText = sCustFldText & ("<input type=text name='"&Custom_Field_Name2&"' size='"&Custom_Field_Values&"'></textarea>")
                elseif Custom_Field_Type = 2 then
                    sCustFldText = sCustFldText & ("<textarea name='"&Custom_Field_Name2&"' rows='"&Custom_Field_Values&"' wrap='ON' cols='35'></textarea>")
                elseif Custom_Field_Type = 3 then
                    sCustFldText = sCustFldText & ("<select name='"&Custom_Field_Name2&"'>")
                    myValueArray = split(Custom_Field_Values,",")
                    for each sValue in myValueArray
                        sCustFldText = sCustFldText & ("<option value='"&sValue&"'>"&sValue&"</option>")
                    next
                    sCustFldText = sCustFldText & ("</select>")
                elseif Custom_Field_Type = 4 then
                    sCustFldText = sCustFldText & ("<input type=checkbox name='"&Custom_Field_Name2&"' value='-1'>")
                end if
                if Field_Required then
                    if Custom_Field_Type=4 then
                        sCustFldText = sCustFldText & ("<input type=hidden name='"&Custom_Field_Name2&"_C' value='Re|String|0|500|-1||"&Custom_Field_Name&"'>")
                    else
                        sCustFldText = sCustFldText & ("<input type=hidden name='"&Custom_Field_Name2&"_C' value='Re|String|0|500|||"&Custom_Field_Name&"'>")
                    end if
                end if
                sCustFldText = sCustFldText & ("</td></tr>")
            end if
        next
    end if
    fn_get_custom_fields=sCustFldText
end function

function fn_other_options(sAppendString)
    
    If Gift_Service = -1 then
        sPaymentText = sPaymentText & ("<input type=hidden name='Is_Gift"&sAppendString&"' value=-1>"&_
            "<tr><td class='normal' colspan=4>Would you like it wrapped?</td>"&_
	        "<td width='50%' class='normal'>"&_
	        "<input type='radio' value='-1' name='Gift_Wrapping"&sAppendString&"'>Yes ("&_
	        Currency_Format_Function(Gift_Wrapping_surcharge)&" Extra)<br>"&_
		    "<input type='radio' value='0' name='Gift_Wrapping"&sAppendString&"' checked>No</td></tr>")
    end if
    If Gift_Message = -1 then
        if Gift = "" or isNull(Gift) then
             Gift = "Gift Message"
        end if
        sPaymentText = sPaymentText & ("<tr><td class='normal' colspan=4>"&Gift&"</font></td><td width='50%'>"&_
            "<textarea name='Gift_Message"&sAppendString&"' rows='3' wrap='ON' cols='20'></textarea></td></tr>")			
    end if	
    sPaymentText = sPaymentText & (fn_get_custom_fields(sAppendString))
    sPaymentText = sPaymentText & ("<tr><td class='normal' colspan=4><b>"&_
        Dict_Special_Notes&"</font></b></td><td class='normal'>"&_
        "<textarea name='Cust_Notes"&sAppendString&"' rows='3' wrap='ON' cols='35'></textarea></td></tr>"&_
        "<tr><td colspan='5'><hr></td></tr><tr>")
end function

function fn_disp_payment()
	Total = fn_Get_Price_Total ()
	sql_Payment_Methods =  "SELECT Payment_Method_id,Payment_name FROM Store_Payment_Methods WITH (NOLOCK) WHERE Accept = 1 AND Store_id="&Store_id
	set paymentfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_Payment_Methods,paymentdata,paymentfields,noRecords)
	
	sPaymentBox = "<select name='Payment_Method' size='1'><option selected value=''>Select Payment"
	if noRecords = 0 then
		FOR paymentrowcounter= 0 TO paymentfields("rowcount")
			'CHECK IF CUSTOMER HAS A SUFFICIENT BUDGET FOR THE PAYMENT METHOD
				Payment_name = paymentdata(paymentfields("payment_name"),paymentrowcounter)
            
				if (lcase(Payment_name) = "charge my account" and cdbl(Budget_left) < cdbl(Total)) then
				elseif (lcase(Payment_name) = "free" and Total > 0) then
				elseif (lcase(Payment_name) = "echeck" and Real_Time_Processor <> 2 and Real_Time_Processor <> 7 and Real_Time_Processor <> 10 and Real_Time_Processor <> 6 and Real_Time_Processor <> 1 and Real_Time_Processor <> 13 and Real_Time_Processor <> 14) then
				elseif (lcase(Payment_name)="debit card" and Real_Time_Processor <>22 and Real_Time_Processor <> 17 and Real_Time_Processor <> 19 and Real_Time_Processor <> 0 and Real_Time_Processor <> 4) then
				else
					if lcase(Payment_name) = "Free" then
						 sPaymentBox = sPaymentBox & "<option selected value='"&Payment_name&"'>"&Payment_name
					else
						 sPaymentBox = sPaymentBox & "<option value='"&Payment_name&"'>"&Payment_name
					end if
				end if
			Next
	end if
	
	sPaymentBox = sPaymentBox & ("<INPUT type='hidden'  name=Payment_Method_C value='Re|String|0|100|||Payment Method'></select>")
	fn_disp_payment = sPaymentBox
	set paymentfields = Nothing
End function

function fn_Shipping_Options_Multi (sel_name, SADDR, location_id)

    sql_sel_items = "exec wsp_trans_totals_shipping "&Store_Id&","&Shopper_ID&","&SADDR&","&location_id&";"
    'fn_print_debug sql_sel_items

    rs_Store.open sql_sel_items, conn_store, 1, 1
    Total_Quantity=FormatNumber(rs_store("tquantity"),2)
    if Total_Quantity=0 then
        rs_store.close
        haveNormalItem = false
        sShippingText = ("<select name='"&sel_name&"' size='1'>"&_
		    "<option value='Free "&Dict_Shipping&"|0'>Free "&Dict_Shipping&" ("&Currency_Format_Function(0)&_
		    ")</option></select>")
    else
        haveNormalItem = true
        Total_Weight=FormatNumber((rs_store("tweight")+Handling_Weight),2,,,0)
        Total_Price=FormatNumber(rs_store("tprice"),2,,,0)
        SFee=FormatNumber(rs_store("sfee"),2,,,0)
        sum_handling=FormatNumber((rs_store("thandling")+sHandling_fee),2,,,0)
        rs_store.close

	    sShippingText = ("<select name='"&sel_name&"' size='1'>")

	    D_dest_data_zip = ""
  	    D_dest_data_country = ""
  	    Set rs_loc = Server.CreateObject("ADODB.Recordset")
  	    sql_sel = "select cu.Country, cu.Zip, cu.state, cu.is_residential, co.country_id,co.country_code from store_customers cu WITH (NOLOCK) inner join sys_countries co WITH (NOLOCK) on cu.country=co.country where Cid="&CID&" and Record_type= "&SADDR&" and store_id="&store_id
	   fn_print_debug sql_sel
        rs_loc.open sql_sel, conn_store, 1, 1
  	    if not (rs_loc.eof) then
  		    D_dest_data_zip = rs_loc("Zip")
  		    D_dest_data_country = rs_loc("Country")
  		    D_dest_data_country_id = rs_loc("Country_id")
  		    dest_country_code = rs_loc("Country_code")
  		    dest_state = rs_loc("state")
  		    is_residential = rs_loc("is_residential")
  	      else
  	      	fn_error "You cannot proceed without a valid shipping address."
		 end if
  	    rs_loc.close
   
	    Dim arrDualArray(30,2)
	    iArrayIndex=0
    	
	    sFullZip=D_dest_data_zip
        D_dest_data_zip=left(D_dest_data_zip,5)
        
        if isNumeric(D_dest_data_zip) then
            dest_zip_int=D_dest_data_zip
        else
            dest_zip_int=0
        end if
        
        if instr(Shipping_Class,"7") > 0 Then
	        sql_select = "Select * from store_real_time_settings WITH (NOLOCK) where store_id="&Store_Id
	        fn_print_debug sql_select
	        rs_Store.open sql_select, conn_store, 1, 1
            Allowed_Countries = rs_Store("Countries")
            fn_print_debug "Allowed_Countries="&Allowed_Countries
            Restrict_Options = rs_store("Restrict_Options")
            if Restrict_Options<>"" then
                Restrict_Options = replace(Restrict_Options,", ",",")
            end if
            sShippingOptions = ""
        
		
		if USE_USPS then
            USPS=1
        else
            USPS=0
        end if
       
	    if USE_AIRBORNE then
            AIRBORNE=1
        else
            AIRBORNE=0
        end if
       
	    if USE_DHL then
            DHL=1
            DHL_Service=rs_Store("DHL_Service")
        else
            DHL=0
        end if

		if USE_FEDEX then
            FedEx_Pack=rs_Store("FedEx_Pack")
            FedEx_Ground=rs_Store("FedEx_Ground")
            FEDEX=1
        else
            FEDEX=0
        end if
        
		if USE_CANADA then
           Canada_Info=rs_Store("Canada_Login")
           CANADA=1
        else
           CANADA=0
        end if
        
		if USE_CONWAY then
            Conway_Info=rs_Store("Conway_Login")&","&rs_Store("Conway_Password")
            CONWAY=1
        else
            CONWAY=0
        end if
		UPS=0


        if USE_UPS then
            UPS_Info = rs_Store("UPS_User")&","&rs_Store("UPS_Password")&","&rs_Store("UPS_AccessLicense")
            UPS_Pickup=rs_Store("UPS_Pickup")
            UPS_Pack=rs_Store("UPS_Pack")
            UPS=1	
        else
            UPS=0
        end if

		Max_Weight=rs_Store("Max_Weight")
        rs_store.close
        
		'====================== BUILDING QUERYSTRING FOR OLD SERVER
		
        sQuerystring_oldServer = "&USE_UPS="&UPS&"&USE_USPS=0&USE_DHL=0&USE_CANADA=0&USE_CONWAY=0&USE_AIRBORNE=0&USE_FEDEX=0&UPS_Info="&UPS_Info&_
            "&UPS_Pack="&server.URLEncode(UPS_Pack)&"&UPS_Pickup="&server.URLEncode(UPS_Pickup)&_
            "&Conway_Info="&server.URLEncode(Conway_Info)&"&Canada_Info="&server.URLEncode(Canada_Info)&_
            "&Fedex_Pack="&server.URLEncode(Fedex_Pack)&"&Fedex_Ground="&server.URLEncode(Fedex_Ground)&_
            "&DHL_Service="&server.URLEncode(DHL_Service)&"&Max_Weight="&server.URLEncode(Max_Weight)
			
		'====================== BUILDGING QUERYSTRING FOR NEW SERVER	
		UPS=0
        sQuerystring_newServer = "&USE_UPS="&UPS&"&USE_USPS="&USPS&"&USE_DHL="&DHL&"&USE_CANADA="&CANADA&_
            "&USE_CONWAY="&CONWAY&"&USE_AIRBORNE="&AIRBORNE&"&USE_FEDEX="&FEDEX&"&UPS_Info="&UPS_Info&_
            "&UPS_Pack="&server.URLEncode(UPS_Pack)&"&UPS_Pickup="&server.URLEncode(UPS_Pickup)&_
            "&Conway_Info="&server.URLEncode(Conway_Info)&"&Canada_Info="&server.URLEncode(Canada_Info)&_
            "&Fedex_Pack="&server.URLEncode(Fedex_Pack)&"&Fedex_Ground="&server.URLEncode(Fedex_Ground)&_
            "&DHL_Service="&server.URLEncode(DHL_Service)&"&Max_Weight="&server.URLEncode(Max_Weight)

			
			if (instr(Allowed_Countries,"All Countries") >0 or instr(Allowed_Countries,D_dest_data_country)>0) then
                fn_print_debug "in allowed countries if "&D_dest_data_country
			 sDetailString=fn_create_realtime_string(Total_Weight,sFullZip,dest_country_code,dest_state,ship_location_id)
                fn_print_debug sDetailString
                fn_print_debug sQuerystring
                Set xObj = CreateObject("SOFTWING.ASPtear")
                Response.ContentType = "text/html"
				
				fn_print_debug "sUrl="&sUrl_oldServer
				fn_print_debug "sUrl="&sUrl_newServer
				
				sUrl_oldServer="http://shippingserver1.storesecured.com/realtime_rates1.aspx"
				sUrl_newServer="http://shippingserver2.storesecured.com/realtime_rates2.aspx"
				
				strResult_newServer = xObj.Retrieve(sUrl_newServer, 1, sDetailString&sQuerystring_newServer, "", "")
				
				if USE_UPS then
					strResult_oldServer = xObj.Retrieve(sUrl_oldServer, 1, sDetailString&sQuerystring_oldServer, "", "")				
					strResult = strResult_oldServer & Right(strResult_newServer,len(strResult_newServer)-6)
				else
					strResult = strResult_newServer
				end if

				'response.Write("Server_old = " & strResult_oldServer & "<br><br>")
				'response.Write("Server_new = " & strResult_newServer & "<br><br>")
				'response.Write("Combined Result = " & strResult & "<br><br>")
				'response.Write("Old Server Query = " & sQuerystring_oldServer & "<br><br>")
				'response.Write("New Server Query = " & sQuerystring_newServer & "<br><br>")
				                
                if err.number<>0 then
                    response.Write "error="&err.number&"description="&err.description
                end if
                on error goto 0
                
                fn_print_debug "retrieved="&strResult
                Set xObj = Nothing
                dest_country=D_dest_data_country
                Restrict_Options=fn_append_list(Restrict_Options,",","UPS Service name is not available.")

                For Each sVariableString in split(strResult,"&")
                    sStringArray = split(sVariableString,"=")
                    sVariableName = sStringArray(0)
                    sVariableValue = sStringArray(1)
                    if sVariableName="Rates" then
                        Rates=sVariableValue
                        fn_print_debug "rates="&Rates
                        For Each Service In split(Rates,"|")
                            if Service <> "" then
                                sArray=split(Service,",")
                                if ubound(sArray)>=1 then
						  	cShipCompany=sArray(0)
						  	cSingleService=sArray(1)
                                	if cShipCompany="UPS" and cSingleService="Service name is not available." then
						  		cSingleService="Worldwide Saver"
						  	end if
                                	cTotalCharge=sArray(2)

							 fn_print_debug "Restrict_Options="&Restrict_Options
				                if isNull(Restrict_Options) or (fn_Is_Instr_Collection(Restrict_Options,cShipCompany & " " & cSingleService,",")=False and Is_In_Collection(Restrict_Options,cSingleService,",")=False) then
                                    if cSingleService="Media" or cSingleService="BPM" then
                                       cSingleService = cSingleService & " (Restrictions Apply)"
                                    end if                          
                                    if cShipCompany<>"" then
                                        cShipCompany=cShipCompany&" "
                                    end if
                                    iAmount=cdbl(cTotalCharge)+cdbl(sum_Handling)
                                    fn_print_debug "cTotalCharge="&cTotalCharge
                                    fn_print_debug "sum_Handling="&sum_Handling


				                    arrDualArray(iArrayIndex,0) = cShipCompany& cSingleService
					                arrDualArray(iArrayIndex,1) = iAmount
					                arrDualArray(iArrayIndex,2) = trim(cShipCompany)
					                iArrayIndex=iArrayIndex+1
				                end if
				     	end if
                            end if
                            if sErrorString <> "" then
		                       if instr(sErrorString,"Memory")>0 then
		                          Send_Mail "melanie@easystorecreator.com","8588374707@mobile.att.net","Shipping error",fn_get_querystring("Error")
                                end if
	                        end if
	                    Next
	                elseif sVariableName="Error" then
	                	fn_print_debug "Error="&sVariableValue
	                    sErrorString=sVariableValue
	                end if
	            Next
	        end if
        end if
        
        if iArrayIndex=0 then
        	realtime_rates_exist=0
        else
          realtime_rates_exist=1
        end if
        
        sql_shipping = "exec wsp_shipping_display "&Store_Id&","&Total_Weight&","&Total_Price&","&Ship_Location_Id&","&D_dest_data_country_id&",'"&dest_zip_int&"','"&Shipping_Class&"';"
        fn_print_debug sql_shipping
        set shippingfields=server.createobject("scripting.dictionary")
        Call DataGetrows(conn_store,sql_shipping,shippingdata,shippingfields,noRecords)
	    if noRecords=0 then
	        FOR shippingrowcounter= 0 TO shippingfields("rowcount")
	            this_Shipping_Class=shippingdata(shippingfields("shippers_class"),shippingrowcounter)
	            this_shipping_method_name=shippingdata(shippingfields("shipping_method_name"),shippingrowcounter)
	            this_base_fee=shippingdata(shippingfields("base_fee"),shippingrowcounter)
	            this_weight_fee=shippingdata(shippingfields("weight_fee"),shippingrowcounter)
	            this_countries=shippingdata(shippingfields("countries"),shippingrowcounter)
	            this_backup=shippingdata(shippingfields("realtime_backup"),shippingrowcounter)
	            'check countries again because sql proc cant correctly check 100% of time
	            if this_backup=0 or (this_backup<>0 and realtime_rates_exist=0) then
			  If Is_In_Collection(this_countries,D_dest_data_country_id,",") or this_countries="216" then
	                select case this_Shipping_Class
	                    case 1 ship_price=this_base_fee
	                    case 2 ship_price=this_base_fee+(this_weight_fee*total_weight)
	                    case 3 ship_price=this_base_fee+SFee
	                    case 4 ship_price=this_base_fee
	                    case 5 ship_price=this_base_fee*(Total_Price/100)
	                    case 6 ship_price=this_base_fee
	                end select
	                ship_price = ship_price+sum_Handling
	                
				    arrDualArray(iArrayIndex,0) = this_shipping_method_name
				    arrDualArray(iArrayIndex,1) = ship_price
				    arrDualArray(iArrayIndex,2) = this_shipping_method_name
				    iArrayIndex=iArrayIndex+1
	            end if
	            end if
	        next
	    end if
	    set shippingfields=nothing

        sql_shipping = "exec wsp_insurance_display "&Store_Id&","&Total_Weight&","&Total_Price&","&Ship_Location_Id&","&D_dest_data_country_id&",'"&dest_zip_int&"';"
        fn_print_debug sql_shipping
        set shippingfields=server.createobject("scripting.dictionary")
        Call DataGetrows(conn_store,sql_shipping,shippingdata,shippingfields,noRecords)
	    if noRecords=0 then
	        FOR shippingrowcounter= 0 TO shippingfields("rowcount")
	            this_insurance_Class=shippingdata(shippingfields("insurance_class"),shippingrowcounter)
	            this_insurance_method_name=shippingdata(shippingfields("insurance_method_name"),shippingrowcounter)
	            this_base_fee=shippingdata(shippingfields("base_fee"),shippingrowcounter)
	            select case this_insurance_Class
                    case 1 ins_price=this_base_fee
                    case 4 ins_price=this_base_fee
                    case 5 ins_price=this_base_fee*(Total_Price/100)
                end select
                For row = 0 To UBound(arrDualArray) - 1
                    if arrDualArray(row,2) = this_insurance_method_name then
			            'if insurance method name matched then add ins price to ship price
			            arrDualArray(row,1) = arrDualArray(row,1)+ins_price
			        end if    
			    next
	        next
	    end if
	    set shippingfields=nothing
    	
        call sub_DualSorter(arrDualArray, 1)
        sValue=""
        iShippingOptions=0
        
        For i = LBound(arrDualArray) to UBound(arrDualArray)
            if arrDualArray(i, 0) <> "" then
                iShippingOptions=iShippingOptions+1
                sShippingName=arrDualArray(i, 0)
                sShippingPrice=arrDualArray(i, 1)
                if sShippingPrice<0 then
                	sShippingPrice=0
                end if
                sValue=sValue& "<option value="""&sShippingName &"|"&sShippingPrice&""">"
                sValue=sValue& sShippingName &" ("&Currency_Format_Function(sShippingPrice)&")</option>"
            end if
        Next

        if iShippingOptions=0 then
            'no shipping options available give customer an error
            call sub_display_shipping_error(sErrorString,D_dest_data_country,Store_Zip,Store_Country,total_weight,sFullZip,Total_Price,Shipping_Class,Location_Id)
        elseif iShippingOptions<>1 then
            sValue="<option value=''>Please Select</option>"&sValue
        end if
        sShippingText = sShippingText& (sValue&"</select>")
    end if
    fn_Shipping_Options_Multi=sShippingText
    
end function

sub sub_display_shipping_error (sShippingErrorMessage,D_dest_data_country,Store_Zip,Store_Country,total_weight,sFullZip,Total_Price,Shipping_Class,Location_Id)
    'we should of stopped on an error above if there are really no realtime rates and not checking for that causes this to happen immediately before going to get realtime rates
    if Location_Id=0 then
        Location_Name="Default"
    else
        sql_select = "select location_name from store_ship_location WITH (NOLOCK) where store_id="&Store_Id&" and ship_location_id="&location_Id
        fn_print_debug sql_select
        session("sql")=sql_select
        rs_store.open sql_select,conn_store,1,1
        if not rs_Store.eof then
           Location_Name=rs_Store("Location_Name")
        else
            Location_Name="Does Not Exist"
        end if
        rs_store.close
    end if
    sErrorExtra="<BR><BR><a href='modify_my_shipping.asp' class='link'>Click here to modify your "&Dict_Shipping&" information</a><BR>"&_
            "<a href='show_big_cart.asp' class='link'>Click here to return to the shopping cart view</a><BR>"
       
    if instr(sShippingErrorMessage,"Missing value for ZipOrigination")>0 then
      sErrorText="Realtime shipping rates cannot be calculated as there is a missing postal code in the store setup.<BR><HR><BR>Store Owner please ensure you have entered a valid zip code in your store information."
    elseif instr(sShippingErrorMessage,"Please enter a valid ZIP Code for the sender")>0 then
      sErrorText="Realtime shipping rates cannot be calculated as there is an invalid postal code in the store setup.<BR><HR><BR>Store Owner please ensure you have entered a valid zip code in your store information."
    elseif instr(sShippingErrorMessage,"Please enter a valid ZIP Code for the recipient")>0 then
      sErrorText="Realtime shipping rates cannot be calculated as the destination information is invalid.<BR><BR>The ship to zip code you have entered is invalid.  Please correct it and try again.<BR><HR><BR>"&sShippingErrorMessage
    elseif instr(sShippingErrorMessage,"Invalid Zip/Postal Code")>0 or instr(sShippingErrorMessage,"Invalid postal format")>0 or instr(sShippingErrorMessage,"Could not get Service Commitement")>0 then
      sErrorText="Realtime shipping rates cannot be calculated as there is either an invalid origination or destination zip code for the country selected.<BR>Destination Country="&D_dest_data_country&"<BR>Destination Postal Code="&sFullZip&"<BR>Origination Country="&Store_Country&"<BR>Origination Postal Code="&Store_Zip&"<BR><HR><BR>"&sShippingErrorMessage
    elseif instr(sShippingErrorMessage,"The requested service is unavailable between the selected locations")>0 or instr(sShippingErrorMessage,"No Services found for origin and destination addresses")>0 then
      sErrorText="Realtime shipping rates cannot be calculated as there is no service available for the requested orignation to destination with the weight entered.  If you believe there should be service please ensure that the destination and origination are entered correctly.<BR>Destination Country="&D_dest_data_country&"<BR>Destination Postal Code="&sFullZip&"<BR>Origination Country="&Store_Country&"<BR>Origination Postal Code="&Store_Zip&"<BR><HR><BR>"&sShippingErrorMessage
    elseif instr(sShippingErrorMessage,"Invalid or unsupported origin country code")>0 or instr(sShippingErrorMessage,"This measurement system is not valid for the selected country")>0 or instr(sShippingErrorMessage,"Fedex.com Rates is unable to process your request")>0 then
      sErrorText="Realtime shipping rates cannot be calculated as realtime shipping is not supported by this shipping company when the package originates from "&Store_Zip&"," &Store_Country&"&Weight="&(total_weight)&"<BR><HR><BR>"&sShippingErrorMessage
    elseif instr(sShippingErrorMessage,"CONWAY:The remote server returned an error: (401) Unauthorized")>0 then
      sErrorText="Realtime shipping rates cannot be calculated for Conway at this time.<BR><HR><BR>Store Owner please recheck your Conway Login and Password."
    elseif instr(sShippingErrorMessage,"Canada Post :Merchant CPC Id not found on server")>0 then
      sErrorText="Realtime shipping rates cannot be calculated for Canada Post at this time.<BR><HR><BR>Store Owner please recheck your Canada Post merchant ID and ensure it is correct and that you are setup for the live server."
    elseif instr(sShippingErrorMessage,"UPS:Invalid Access License number")>0 then
      sErrorText="Realtime shipping rates cannot be calculated for UPS at this time.<BR><HR><BR>Store Owner please recheck your UPS Access License number and ensure it is correctly entered."
    elseif instr(sShippingErrorMessage,"USPS:Please enter the package weight")>0 then
      sErrorText="Realtime shipping rates cannot be calculated for USPS at this time.<BR><HR><BR>Store Owner please ensure that your items have the correct weight assigned, USPS cannot provide quotes for packages weighing 0 lbs as this is not a valid weight."
    elseif sShippingErrorMessage<>"" then
      sErrorText=sShippingErrorMessage&"<BR>Destination Country="&D_dest_data_country&"<BR>Destination Postal Code="&sFullZip&"<BR>Origination Country="&Store_Country&"<BR>Origination Postal Code="&Store_Zip&"<BR>Location="&Location_Name
    elseif instr(Shipping_Class,"7") = 0 then
        sErrorText="There are no "&Dict_Shipping&" methods available for your order, please ensure that you have entered a proper country and postal code.<BR>Total Price="&Total_Price&"<BR>Total Weight="&Total_Weight&"<BR>Destination Country="&D_dest_data_country&"<BR>Destination Postal Code="&sFullZip&"<BR>Ship From Location="&Location_Name&"<BR><HR><BR>Store Owner, please ensure if you wish for this customer to checkout that you have created an appropriate "&Dict_Shipping&" method for the criteria above.  If there are no valid "&Dict_Shipping&" methods the customer will be unable to continue."
    end if
    if sErrorText<>"" then
        fn_error sErrorText&sErrorExtra
    end if
end sub

Sub sub_DualSorter( byRef arrArray, DimensionToSort )
    Dim row, j, StartingKeyValue, StartingOtherValue, _
        NewStartingKey, NewStartingOther, _
        swap_pos, OtherDimension
    Const column = 1
    
    ' Ensure that the user has picked a valid DimensionToSort
    If DimensionToSort = 1 then
		OtherDimension = 0
	ElseIf DimensionToSort = 0 then
		OtherDimension = 1
	Else
	    'Shoot, invalid value of DimensionToSort
	    Response.Write "Invalid dimension for DimensionToSort: " & _
	                   "must be value of 1 or 0."
	    Response.End
	End If
    
    For row = 0 To UBound( arrArray, column ) - 1
    'Start outer loop.

        'Take a snapshot of the first element
        'in the array because if there is a 
        'smaller value elsewhere in the array 
        'we'll need to do a swap.
        StartingKeyValue = arrArray ( row, DimensionToSort )
        StartingOtherValue = arrArray ( row, OtherDimension )
        
        ' Default the Starting values to the First Record
        NewStartingKey = arrArray ( row, DimensionToSort )
        NewStartingOther = arrArray ( row, OtherDimension )
        
        swap_pos = row
		
        For j = row + 1 to UBound( arrArray, column )
        'Start inner loop.
            If arrArray ( j, DimensionToSort ) < NewStartingKey Then
            'This is now the lowest number - 
            'remember it's position.
                swap_pos = j
                NewStartingKey = arrArray ( j, DimensionToSort )
                NewStartingOther = arrArray ( j, OtherDimension )
            End If
        Next
		
        If swap_pos <> row Then
        'If we get here then we are about to do a swap
        'within the array.
            arrArray ( swap_pos, DimensionToSort ) = StartingKeyValue
            arrArray ( swap_pos, OtherDimension ) = StartingOtherValue
            
            arrArray ( row, DimensionToSort ) = NewStartingKey
            arrArray ( row, OtherDimension ) = NewStartingOther
            
        End If	
    Next
End Sub

' ================================================================ 

function fn_create_realtime_string(sWt,DestZip,DestCty,DestST,ship_location_id)
    if ship_location_id = 0 then
        OrgZip = Store_Zip
        OrgCty = fn_country_code(Store_Country)
        OrgST = Store_State
    else
        sql_select = "select location_state, location_zip, location_country from store_ship_location WITH (NOLOCK) where Ship_Location_Id="&ship_location_id&" and store_id="&Store_Id
        fn_print_debug sql_select
        session("sql")=sql_select
	   rs_Store.open sql_select,conn_store,1,1
        if not rs_store.eof and not rs_store.bof then
            OrgZip = rs_store("location_zip")
            OrgCty = fn_country_code(rs_store("location_country"))
            OrgST = rs_store("location_state")
        end if
        rs_store.close
    end if

    if not isNumeric(sWt) or (sWt<.1) then
        'set min weights for carriers
        sWt=.1
    end if
    if Use_Conway then
        sWt=formatnumber(sWt,0)
    end if
    sInfo = "&OrgZip="&OrgZip&"&OrgST="&OrgST&"&OrgCty="&OrgCty&"&DestZip="&DestZip&_
        "&DestST="&DestST&"&DestCty="&DestCty&"&sWt="&sWt&"&Res="&DestRes
    fn_create_realtime_string=sInfo
end function

%>

<%

function getConfOptID(theItem, theConfig, theName)
	Set rs_storeloc = Server.CreateObject("ADODB.Recordset")
	sql_sel = "select det_id from Store_Items_Configs_Dets where store_id="&store_id&" and item_id="&theItem&" and config_id="&theConfig&" and name='"&theName&"'"
	rs_storeloc.open sql_sel, conn_store, 1, 1
	if not rs_storeloc.eof then
		getConfOptID = rs_storeloc("det_id")
	else
		getConfOptID = 0
	end if
	rs_storeloc.close
	Set rs_storeloc = nothing
end function

function getConfOptName(theItem, theConfig, theID)
	Set rs_storeloc = Server.CreateObject("ADODB.Recordset")
	sql_sel = "select Name from Store_Items_Configs_Dets where store_id="&store_id&" and item_id="&theItem&" and config_id="&theConfig&" and Det_ID="&theID&""
	rs_storeloc.open sql_sel, conn_store, 1, 1
	if not rs_storeloc.eof then
		getConfOptName = rs_storeloc("Name")
	else
		getConfOptName = ""
	end if
	rs_storeloc.close
	Set rs_storeloc = nothing
end function

sub getItemPrice (thePrice, theWeight, theSku, theIName, theItem, theQuantity, straceStock)

on error resume next
	Set rs_store_loc = Server.CreateObject("ADODB.Recordset")
	sql_sel_item = "select Quantity_in_stock, Quantity_Control, Item_Name, Item_SKU, Item_Weight, Use_Price_By_Matrix, Retail_Price, Retail_Price_special_Discount, Special_start_date, Special_end_date from store_items where item_id="&theItem&" and store_id="&store_id
   rs_store_loc.open sql_sel_item, conn_store, 1, 1
	if rs_store_loc.eof then
		thePrice = 0
		theWeight = 0
		theSku = ""
		theIName = ""
		rs_store.close
		exit sub
	end if
	Use_Price_By_Matrix = rs_store_loc("Use_Price_By_Matrix")
	RPrice = cdbl(rs_store_loc("Retail_Price"))
        RPrice = get_group_price(store_id,groups,theItem,RPrice)
        
	theWeight = rs_store_loc("Item_Weight")
	theSku = rs_store_loc("Item_SKU")
	theIName = rs_store_loc("Item_Name")
	exQ = rs_store_loc("Quantity_in_stock")

	if Use_Price_By_Matrix <> 0 then
		rs_store_loc.close
		'CALCULATE THE PRICE BASED ON THE MATRIX
		sql_select="select Item_Price from Store_items_Price_Matrix where Item_Id = "&theItem&" and store_id = "&store_id&" and ( (Matrix_low <= "&theQuantity&" and Matrix_high >= "&theQuantity&") or (Matrix_low <= "&theQuantity&" and Matrix_high = -1) ) and store_id = "&store_id&""
		
      rs_store_loc.open sql_select,conn_store,1,1
		if rs_store_loc.bof = false then
			thePrice = theQuantity * rs_Store_loc("Item_Price")
		else
			thePrice = theQuantity * cdbl(RPrice)
		end if
	else
		If  rs_store_loc("Retail_Price_special_Discount")>0 AND rs_store_loc("Special_start_date") <= Now() AND rs_store_loc("Special_end_date") >= Now() Then
			thePrice = theQuantity * cdbl(rs_store_loc("Retail_Price"))*(cdbl(100)-cdbl(rs_store_loc("Retail_Price_special_Discount")))/100
	    Else
			thePrice = theQuantity * cdbl(rs_store_loc("Retail_Price"))
		end if
	end if


	if straceStock then
		If rs_store_loc("Quantity_Control") <> 0 then

			if cint(exQ) < cint(theQuantity) then
            sp = 0
			else
				sp = (exQ - (exQ mod theQuantity))/theQuantity
			end if

			if stock_corection = -1 then
				stock_corection = sp
			else
				if sp<stock_corection then
					stock_corection = sp
				end if
			end if
		end if
	end if

	rs_store_loc.close
	set rs_store_loc = nothing

end sub

function get_group_price(store_id,groups,item_id,retail_price)
    get_group_price=retail_price
    sql_select = "select Group_Price,customer_group from store_items_price_group where store_id="&Store_Id&" and Item_Id="&Item_Id&" order by group_price"
    Set rs_storeprice = Server.CreateObject("ADODB.Recordset")

    rs_storeprice.open sql_select, conn_store, 1, 1
    do while not rs_storeprice.eof
       sCustomerGroup = rs_storeprice("customer_group")
       for each iGroup in split(groups,",")
           if iGroup = sCustomerGroup then
              get_group_price=rs_storeprice("Group_Price")
              rs_storeprice.close
              exit function
           end if
      next
      rs_storeprice.movenext
   loop
   rs_storeprice.close
   set rs_storeprice=nothing
end function

%>

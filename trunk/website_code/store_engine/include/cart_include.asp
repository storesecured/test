<%
if request.Form<>"" then
    'ERROR CHECKING
    If Form_Error_Handler(Request.Form) <> "" then 
	    Error_Log = Form_Error_Handler(Request.Form)
	    fn_error Error_Log
    else
    		fn_print_debug "request.form="&request.form
        if fn_get_querystring("Return_To") <> "" then
	        Return_To = fn_get_querystring("Return_To")
	        fLocal=instr(Return_To,Switch_Name)
	        if Return_To <> "http:" and (fLocal=0 or (fLocal>0 and instr(Return_To,"items/")>0)) then
	        else
		        Return_To = Switch_Name&"items/list.htm"
	        end if
        elseif request.form("Return_To")<>"" then
            Return_To = request.form("Return_To")
        elseif request.ServerVariables("HTTP_REFERER")<>"" and request.form("Return_Back")<>"0" then
            Return_To = request.ServerVariables("HTTP_REFERER")
            sArray = split(Return_To,"#")
    		  Return_To = sArray(0)
		  if instr(Return_To,"?")=0 and request.form("Item_Id")<>"" then
		  	Return_To=Return_To&"#"&request.form("Item_Id")
		  end if
        else
	        Return_To = Switch_Name&"items/list.htm"
        end if
        fn_print_debug "Return_to="&Return_To
        
        if Request.Form("Continue_Shopping") <> "" or Request.Form("Continue_Shopping.x") <> "" then
            Line_id = 1
            Call fn_update_cart
	        if not isnull(Continue_Shopping) and Continue_Shopping <> "" then
	            fn_redirect Continue_Shopping
	        else
		        fn_redirect Return_To
	        end if
        elseif Request.Form("Save_Cart") <> "" or Request.Form("Save_Cart.x") <> "" then
	        fn_redirect Switch_Name&"Save_Cart.asp"
        elseif Request.Form("Retrieve_Cart") <> "" or Request.Form("Retrieve_Cart.x") <> "" then
	        fn_redirect Switch_Name&"Retrieve_Cart.asp"
        elseif Request.Form("Check_Out") <> "" or Request.Form("Check_Out.x") <> "" then
	        Line_id = 1
	        call fn_update_cart
	        call sub_check_quantity
	        fn_redirect Secure_Site&"Before_Payment.asp?Shopper_Id="&Shopper_Id
	        'fn_redirect Switch_Name&"Before_Payment.asp?Shopper_Id="&Shopper_Id
        elseif Request.Form("Re_Calculate") <> "" or Request.Form("Re_Calculate.x") <> "" then
	        call fn_update_cart
        else 
            sub_add_to_cart	     
        end if
    end if
end if
fn_redirect Switch_name&"show_big_cart.asp?Return_To="&server.URLEncode(Return_To)

Function fn_update_cart
    fn_print_debug "in fn_update_cart"
    'UPDATE THE CART REDIRECTION FLAG
    Redirect_To_Cart = Request.Form("Redirect_To_Cart")
    Redirect_To_Cart_Old = Request.Form("Redirect_To_Cart_Old")
    if isNumeric(Redirect_To_Cart) and isNumeric(Redirect_To_Cart_Old) then
    		if cint(Redirect_To_Cart) <> cint(Redirect_To_Cart_Old) then
        		sql_update = "exec wsp_cart_redirect "&Store_Id&","&Shopper_ID&","&Redirect_To_Cart&";"
    		end if
	end if
    'GENERAL HANDLER FOR FORM LOOPS
    'USE THE LINE_ID PARAM TO ITERATE VIA WHILE LOOP
    Line_id = 1
	items_to_delete = ""
	Do While Request.Form("Item_Id_"&Line_id) <> ""
        Shopping_cart_id = Request.Form("Shopping_cart_id_"&Line_id)
        Quantity = Request.Form("Quantity_"&Line_id)
        if NOT IsNumeric(Quantity) then
			'SHOPPING CART FIELD MUST BE NUMERIC
			fn_redirect Switch_Name&"error.asp?Message_Id=27"
		end if
		If Request.Form("Remove_From_Cart_"&Line_id) <> "" OR Quantity <=0 OR Quantity = "" Then
			if items_to_delete = "" then
			    items_to_delete = Shopping_cart_id
			else
			    items_to_delete = ItemS_to_delete&","&Shopping_cart_id
			end if
		End if
		If Quantity > 0 then
			Item_Attribute_ids = ""
			Quantity_Old=Request.Form("Quantity_Old_"&Line_id)
			'CHECK FOR USED DEFINABLE FIELDS
		    U_d_1 = checkStringForQ(Request.Form("U_d_1_"&Line_id))
		    U_d_2 = checkStringForQ(Request.Form("U_d_2_"&Line_id))
		    U_d_3 = checkStringForQ(Request.Form("U_d_3_"&Line_id))
		    U_d_4 = checkStringForQ(Request.Form("U_d_4_"&Line_id))
		    U_d_5 = checkStringForQ(Request.Form("U_d_5_"&Line_id))
		    'CHECK FOR USED DEFINABLE FIELDS
		    U_d_1_old = checkStringForQ(Request.Form("U_d_1_old_"&Line_id))
		    U_d_2_old = checkStringForQ(Request.Form("U_d_2_old_"&Line_id))
		    U_d_3_old = checkStringForQ(Request.Form("U_d_3_old_"&Line_id))
		    U_d_4_old = checkStringForQ(Request.Form("U_d_4_old_"&Line_id))
		    U_d_5_old = checkStringForQ(Request.Form("U_d_5_old_"&Line_id))
			
			'only do update if something changed
			session("sql")="Quantity="&quantity
			if ((isNumeric(Quantity) and isNumeric(Quantity_Old)) and Quantity<>Quantity_Old) or U_d_1<>U_d_1_old or U_d_2<>U_d_2_old or U_d_3<>U_d_3_old or U_d_4<>U_d_4_old or U_d_5<>U_d_5_old then
			    Item_Id = Request.Form("Item_Id_"&Line_id)
			    if not isNumeric(item_id) then
					fn_error "The item id must be numeric.  Please contact the store owner to correct this issue."
			    end if
			    Item_Attribute_ids = Request.Form("Item_Attribute_ids_"&Line_id)
			    if Item_Attribute_ids="" then
				    Item_Attribute_ids = -1
			    end if

			    'CHECK FOR USED DEFINABLE FIELDS
			    U_d_1 = checkStringForQ(Request.Form("U_d_1_"&Line_id))
			    U_d_2 = checkStringForQ(Request.Form("U_d_2_"&Line_id))
			    U_d_3 = checkStringForQ(Request.Form("U_d_3_"&Line_id))
			    U_d_4 = checkStringForQ(Request.Form("U_d_4_"&Line_id))
			    U_d_5 = checkStringForQ(Request.Form("U_d_5_"&Line_id))
			    Cust_Price = checkStringForQ(Request.Form("Cust_Price_"&Line_id))
			    if Cust_Price="" then
                    Cust_Price=-1
                elseif Cust_Price<>"-1" then
                    Cust_Price=request.Form("sale_price_"&Line_id)
                end if
                sql_update = sql_update&"exec wsp_cart_update "&store_id&","&Shopper_id&","&Shopping_cart_id&","&Item_Id&","&Quantity&",'"&Groups&"','"&Item_Attribute_ids&"','"&u_d_1&"','"&u_d_2&"','"&u_d_3&"','"&u_d_4&"','"&u_d_5&"',"&Cust_Price&";"
            end if
        end if
        Line_id = Line_id + 1
    Loop
    if items_to_delete <> "" and items_to_delete<>"-1" Then
		sql_update = sql_update & "exec wsp_cart_delete "&Store_Id&","&Shopper_Id&",'"&items_to_delete&"','"&Groups&"';"                  
	End if
    if sql_update<>"" then
        fn_print_debug sql_update
        on error resume next
		Conn_Store.Execute sql_update
		if err.number<>0 then
			fn_error err.description
		end if
		on error goto 0
    end if
End Function

Sub sub_add_to_cart()
    'RETRIEVE FOR DATA
    if request.form("num_items") = "" then
	    i = 0
	    num_items = 0
    else
	    i = 1
	    num_items = request.form("num_items")
    end if

    Do While (cint(i) <= cint(num_items))
	    if i = 0 then
		    i = ""
	    end if
	    quantity = Request.Form("quantity"&i)

	    if not isNumeric(i) then
		    i = 0
	    end if
	    
	    if NOT IsNumeric(Quantity) then
			'SHOPPING CART FIELD MUST BE NUMERIC
			fn_redirect Switch_Name&"error.asp?Message_Id=27"
		end if

	    if quantity > 0 then
		    if i = 0 then
			    i = ""
		    end if
		    fn_print_debug "Request.form="&Request.form
		    Item_name = checkStringForQ(Request.Form("Item_name"&i))
		    Base_name = Item_name
		    Item_Sku = checkStringForQ(Request.Form("Item_Sku"&i))
		    Item_Weight = Request.Form("Item_Weight"&i)
		    Item_Id = Request.Form("Item_Id"&i)
		    U_d_1 = checkStringForQ(Request.Form("U_d_1"&i))
		    U_d_2 = checkStringForQ(Request.Form("U_d_2"&i))
		    U_d_3 = checkStringForQ(Request.Form("U_d_3"&i))
		    U_d_4 = checkStringForQ(Request.Form("U_d_4"&i))
		    U_d_5 = checkStringForQ(Request.Form("U_d_5"&i))
		    Cust_Price = request.form("Cust_Price"&i)
		    if not isNumeric(Cust_Price) or isnull(Cust_Price) or Cust_Price="" then
			    Cust_Price=-1
		    end if
			
		    'PARAMERER TO DETERMINE THE AMOUNT OF DISCOUNT RECEIVED BY DYNAMIC PRICING
		    matrix_discout=1
		    iAttrItem=0
			 
		    if Not IsNumeric(quantity) then
			    fn_redirect Switch_Name&"error.asp?Message_Id=27"
		    end if
            
		    if  Request.Form("Count_Class") = "" then
			    Count_Class = 0
		    else
			    Count_Class = Request.Form("Count_Class")
		    end if
			
			Item_Attribute_ids = ""
		    Item_Attribute_Class_value = ""
		    Item_Attribute_Price_Add = 0
		    Item_Attribute_Weight_Add = 0
		    Item_Attribute_skus = ""
		    Item_Attribute_hiddens = ""
		    if not(isNumeric(Count_Class)) then
		    		Count_Class=0
		    end if
		    if Count_Class > 0 Then
			    'WE HAVE ATTRIBUTES FOR THAT ITEM
			    'ALSO CALCULATE THE TOTAL WEIGHT AND PRICE ADDITION
			    'THE STRING WE ARE GETTING IS ID|CLASS|VALUE|PRICE_ADD|WEIGHT_ADD, I.E. : 2|RED|23|12

			    
			    For Counter = 1 To Count_Class
			        fn_print_debug "in counter"
				    Str = "Count_Class_"&Counter
				    sRequestValue=Request.Form(Str)
				    fn_print_debug "request value='"&sRequestValue&"'"    
				    if sRequestValue<>"" then
				        Item_Attribute_Array = Split (checkStringForQ(sRequestValue),"|")
				        Attribute_Id = Item_Attribute_Array(0)
				        fn_print_debug "Attribute_Id='"&Attribute_Id&"'"    
				        if isNumeric(Attribute_Id) and Attribute_Id<>"" and not isNull(Attribute_id) then
					        if Item_Attribute_ids="" then
					            Item_Attribute_ids = Attribute_Id
					        else
					            Item_Attribute_ids = Item_Attribute_ids&","&Attribute_Id
					        end if
					        fn_print_debug "Item_Attribute_ids='"&Item_Attribute_ids&"'"    
					    end if
				    end if
				next
				if Item_Attribute_ids<>"" then
				    sql_item = "select a.*,i.item_name,i.item_sku as iitem_sku from store_items_attributes a WITH (NOLOCK) left join store_items i WITH (NOLOCK) on i.store_id=a.store_id and i.item_id=a.iitem_id where a.store_id="&store_id&" and attribute_id in ("&Item_Attribute_ids&") order by attribute_class"
		            fn_print_debug sql_item
		            set attrfields=server.createobject("scripting.dictionary")
                    Call DataGetrows(conn_store,sql_item,attrdata,attrfields,noRecords)
                    FOR attrrowcounter= 0 TO attrfields("rowcount")
                        use_item=attrdata(attrfields("use_item"),attrrowcounter)
                        if use_item=0 then
                            attribute_class = attrdata(attrfields("attribute_class"),attrrowcounter)
                            attribute_value = attrdata(attrfields("attribute_value"),attrrowcounter)
                            Item_name = Item_name&"<BR>&nbsp;&nbsp;&nbsp;"&checkStringForQ(attribute_class&"="&attribute_value)
		                    attribute_sku = attrdata(attrfields("attribute_sku"),attrrowcounter)    
                            attribute_hidden = attrdata(attrfields("attribute_hidden"),attrrowcounter)    
                            Item_Attribute_Price_Add = Item_Attribute_Price_Add + attrdata(attrfields("attribute_price_difference"),attrrowcounter)    
                            Item_Attribute_Weight_Add = Item_Attribute_Weight_Add + attrdata(attrfields("attribute_weight_difference"),attrrowcounter)    
                            if Item_Attribute_hiddens = "" then
                                Item_Attribute_hiddens = attribute_hidden
                            else
                                Item_Attribute_hiddens = Item_Attribute_hiddens & "," & attribute_hidden
                            end if
                            if Item_Attribute_skus = "" then
                                Item_Attribute_skus = attribute_sku
                            else
                                Item_Attribute_skus = Item_Attribute_skus & "," & attribute_sku
                            end if

                        else
                            attribute_class = attrdata(attrfields("attribute_class"),attrrowcounter)
                            attribute_value = attrdata(attrfields("attribute_value"),attrrowcounter)
                            Item_name = Item_name&"<BR>&nbsp;&nbsp;&nbsp;"&checkStringForQ(attribute_class&"=*"&attribute_value)
		                  IItem_Id = attrdata(attrfields("iitem_id"),attrrowcounter)
                            IItem_Quantity = attrdata(attrfields("iitem_quantity"),attrrowcounter)*quantity
                            IItem_Name = attrdata(attrfields("item_name"),attrrowcounter)
                            IItem_Sku = attrdata(attrfields("iitem_sku"),attrrowcounter)
                            iAttrItem=1
                            sql_update=sql_update&"exec wsp_cart_add "&Store_Id&","&Shopper_Id&","&IItem_Id&","&cid&","&IItem_Quantity&",'"&iitem_name&" (for *"&Base_name&")','"&iitem_sku&"','"&u_d_1&"','"&u_d_2&"','"&u_d_3&"','"&u_d_4&"','"&u_d_5&"',"&Cust_Price&",'"&Groups&"','','','',0,0,0;"
                        end if 
                    next
		            set attrfields=nothing
		        end if
		    end if    

		    sql_update=sql_update&"exec wsp_cart_add "&Store_Id&","&Shopper_Id&","&Item_Id&","&cid&","&quantity&",'"&Item_Name&"','"&Item_Sku&"','"&u_d_1&"','"&u_d_2&"','"&u_d_3&"','"&u_d_4&"','"&u_d_5&"',"&Cust_Price&",'"&Groups&"','"&Item_Attribute_ids&"','"&Item_Attribute_skus&"','"&Item_Attribute_hiddens&"',"&Item_Attribute_Price_Add&","&Item_Attribute_Weight_Add&","&iAttrItem&";"
	    end if
	    if not isNumeric(i) then
		    i = 0
	    end if
	    i = i + 1
    loop
    if sql_update<>"" then

    		fn_print_debug sql_update

		session("sql")=sql_update
		on error resume next
		Conn_Store.Execute sql_update
		if err.number<>0 then
			fn_error err.description
		end if
		on error goto 0
    end if
    sArray = split(Return_To,"#")
    ReturnTo = sArray(0)&"#"&Item_Id

    if Redirect_To_Cart <> 0 then
        ReferTo = server.urlencode(ReturnTo)
    else
	    fn_redirect ReturnTo
    end if

End Sub

%>

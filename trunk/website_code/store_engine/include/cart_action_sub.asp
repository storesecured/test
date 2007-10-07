
<%

'THE SUB UPDATES STORE_SHOPPINGCART TABLE, ACCORDING TO THE REQUESTED
'OPERATION
Sub Cart_Action(Action,Shopper_id,Item_Id,Item_Name,Item_Sku,Quantity,Sale_Price,Retail_Price,Weight,Cid,Cust_price,Item_Attribute_ids,Item_Attribute_skus,Item_Attribute_hiddens,Cart_name,Recurring_fee,Recurring_days,Ship_LOcation_Id,Waive_Shipping,Item_Handling,Item_Attribute_Price_Add)
			if cid="" or isNull(cid) then
			   cid=0
			end if
			if Waive_Shipping then
			   Waive_Shipping=1 
			else
			   Waive_Shipping=0
			end if
			
			if Item_Handling="" then
			   Item_Handling=0
			end if

   'ACTION = INSERT / UPDATE / DELETE
    err.number=0

    sql_select = "exec wsp_cart_action '"&Action&"',"&Shopper_ID&","&Item_id&",'"&Item_Name&"','"&_
        Item_Sku&"',"&Quantity&","&Sale_Price&","&Retail_Price&","&Weight&","&Cid&","&Cust_Price&",'"&_
        Item_Attribute_ids&"','"&Item_Attribute_skus&"','"&Item_Attribute_hiddens&"','"&_
        Cart_name&"',"&Store_id&", "&Wholesale_Price&",'"&U_d_1&"','"&U_d_2&"','"&_
        U_d_3&"','"&U_d_4&"','"&U_d_5&"',"&Recurring_Fee&","&Recurring_days&","&_
        Ship_Location_Id&","&Waive_Shipping&","&Item_Handling&","&Item_Attribute_Price_Add
    fn_print_debug sql_select
    on error resume next
    conn_store.execute sql_select
    if err.number<>0 then
	    fn_print_debug "Error="&err.number
	    fn_error err.description
    end if
    on error goto 0

End Sub 


%>

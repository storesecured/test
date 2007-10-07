<%


sub create_form_post_2checkout ()

    sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
    rs_Store.open sql_real_time,conn_store,1,1
    rs_Store.MoveFirst
    Do While Not Rs_Store.EOF
        select case Rs_store("Property")
        case "seller_id"
            seller_id = decrypt(Rs_store("Value"))
        end select
        Rs_store.MoveNext
    Loop
    Rs_store.Close
    on error resume next
    if not isNumeric(seller_id) then
        fn_error "The 2Checkout Seller Id must be numeric."
    end if
    if isNumeric(seller_id) and seller_id < 200000 then
        response.write "<form method='POST' action='https://www.2checkout.com/cgi-bin/sbuyers/cartpurchase.2c' name='Payment' onSubmit=""return checkFields();""><input type=hidden name=sid value="&seller_id&">"
    else
        response.write "<form method='POST' action='https://www2.2checkout.com/2co/buyer/purchase' name='Payment' onSubmit=""return checkFields();""><input type=hidden name=sid value="&seller_id&">"
    end if
end sub

sub create_form_content_2checkout ()
    Return_Addr = Switch_Name&"include/2checkout/2checkout.asp"
 %>

	<input type="hidden" name="fixed" value="Y">
	<input type="hidden" name="x_receipt_link_url" value="<%=Return_Addr %>">
	<% if store_id=1 then %>
	<input type="hidden" name="demo" value="Y">
	<% end if %>
    	<input type="hidden" name="cart_order_id" value="<%= oid %>">
	<input type="hidden" name="Shopper_id" value="<%= shopper_id %>">
	<input type="hidden" name="total" value="<%= formatnumber(GGrand_Total,2) %>">
	<% rs_store.open "select * from store_customers where record_type=0 and cid="&cid&" and Store_Id="&Store_Id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
		<INPUT type="hidden" NAME="card_holder_name" value="<%= rs_store("First_Name") & " " & rs_store("Last_Name") %>">
		<INPUT type="hidden" NAME="street_address" value="<%= rs_store("Address1") %>">
		<INPUT type="hidden" NAME="city" value="<%= rs_store("City") %>">
		<INPUT type="hidden" NAME="zip" value="<%= rs_store("Zip") %>">
		<INPUT type="hidden" NAME="country" value="<%= rs_store("Country") %>">
		<INPUT type="hidden" NAME="phone " value="<%= rs_store("Phone") %>">
		<INPUT type="hidden" NAME="email" value="<%= rs_store("Email") %>">
		  <INPUT type="hidden" NAME="state" value="<%= rs_store("State") %>">
	  <% End If
	rs_store.close %>

	<INPUT type="hidden" NAME="Store_Id" value="<%= Store_Id %>">
	<% rs_store.open "Select * from Store_Purchases where Shopper_id ="&Shopper_ID&" AND oid = "&Oid&" and Store_id ="&Store_id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
		<INPUT type="hidden" NAME="ship_name" value="<%= rs_store("ShipFirstName") & " " & rs_store("ShipLastName") %>">
		<INPUT type="hidden" NAME="ship_street_address" value="<%= rs_store("ShipAddress1") %>">
		<INPUT type="hidden" NAME="ship_city" value="<%= rs_store("ShipCity") %>">
		<INPUT type="hidden" NAME="ship_zip" value="<%= rs_store("ShipZip") %>">
		<INPUT type="hidden" NAME="ship_country" value="<%= rs_store("ShipCountry") %>">
		<INPUT type="hidden" NAME="ship_state" value="<%= rs_store("ShipState") %>">
	<% End If
	rs_store.close 
	Set rs_store_tran = Server.CreateObject("ADODB.Recordset")
	sql_select_other_order = "select oid, shipto,Shipping_Method_Price, Tax from store_purchases where (masteroid="&oid&" or oid="&oid&") and store_id="&store_id

	rs_Store.open sql_select_other_order,conn_store,1,1
	Shipping_Method_Price = rs_Store("Shipping_Method_Price")
	Tax = rs_Store("Tax")
	iItem_no=1
	
        sql_tran_items = "select s.Item_Id,i.Item_Name,i.Item_Sku,s.Quantity,s.Sale_Price,s.Item_Weight,s.Cust_price,s.Item_Attribute_ids,s.Item_Attribute_skus,s.Item_Attribute_hiddens,s.Store_id,s.U_d_1, s.U_d_2, s.U_d_3, s.U_d_4 from store_ShoppingCart s WITH (NOLOCK) inner join store_items i WITH (NOLOCK) on s.item_id=i.item_id and s.store_id=i.store_id where Cart_Processed<>-1 and Shopper_id ="&Shopper_ID&"	AND s.Store_id ="&Store_id
        session("sql") = sql_tran_items
                rs_store_tran.open sql_tran_items, conn_store, 1, 1
		do while not rs_store_tran.eof
			 Item_Name = rs_store_tran("Item_Name")
			 Item_Sku = rs_store_tran("Item_Sku")
			 Quantity = rs_store_tran("Quantity")
			 Sale_Price = rs_store_tran("Sale_Price") %>
			 <input type='hidden' name='id_type' value='1'>
                        <input type='hidden' name='c_prod_<%= iItem_no %>' value='<%= Item_Sku %>,<%= Quantity %>'>
                        <input type='hidden' name='c_name_<%= iItem_no %>' value='<%= Item_Name %>'>
                        <input type='hidden' name='c_description_<%= iItem_no %>' value='<%= Item_Name %>'>
                        <input type='hidden' name='c_price_<%= iItem_no %>' value='<%= Sale_Price %>'>

			 <% rs_store_tran.movenext
			 iItem_no=iItem_no+1
		loop
		rs_store_tran.close
		rs_Store.movenext
        rs_Store.Close
                %>


				
<% end sub %>

<%
sub create_form_post_internet ()
    sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
    rs_Store.open sql_real_time,conn_store,1,1
    rs_Store.MoveFirst
    Do While Not Rs_Store.EOF
     select case Rs_store("Property")
	     case "MerchantNumber"
		     MerchantNumber = decrypt(Rs_store("Value"))
      end select
     Rs_store.MoveNext
    Loop
    Rs_store.Close
    response.write "<form method='POST' action='https://secure.internetsecure.com/process.cgi' name='Payment' onSubmit=""return checkFields();""><input type=hidden name=MerchantNumber value="&MerchantNumber&">"

end sub

sub create_form_content_internet ()
 %>
	<Font color=red>After entering your payment information at our processor's site <B>CLICK CONTINUE</b> to return to our store to finish your order processing and print a receipt.</font>
	<% Return_Addr = Switch_Name&"include/internetsecure/internetsecure.asp?Shopper_Id="&Shopper_Id

	Set rs_store_tran = Server.CreateObject("ADODB.Recordset")
	sql_select_other_order = "select oid, shipto,Shipping_Method_Price, Tax, coupon_amount from store_purchases where (masteroid="&oid&" or oid="&oid&") and store_id="&store_id

	rs_Store.open sql_select_other_order,conn_store,1,1
	Shipping_Method_Price = rs_Store("Shipping_Method_Price")
	Tax = rs_Store("Tax")
	Coupon_Amount = rs_Store("Coupon_Amount")
	sProductString = ""
	do while not rs_Store.eof
		loid = rs_Store("oid")
		lshipto = rs_Store("shipto")
		sql_tran_items = "select Item_Id,Item_Name,Item_Sku,Quantity,Sale_Price,Item_Weight,Cust_price,Item_Attribute_ids,Item_Attribute_skus,Item_Attribute_hiddens,Store_id,U_d_1, U_d_2, U_d_3, U_d_4 from store_transactions where transaction_Processed=0 and shipto="&lshipto&"  and Shopper_id ="&Shopper_ID&"	AND Store_id ="&Store_id
		session("sql")=sql_tran_items
		rs_store_tran.open sql_tran_items, conn_store, 1, 1
		do while not rs_store_tran.eof
			 Item_Id = rs_store_tran("Item_Id")
			 Item_Name = rs_store_tran("Item_Name")
			 Item_Sku = rs_store_tran("Item_Sku")
			 Quantity = rs_store_tran("Quantity")
			 Sale_Price = rs_store_tran("Sale_Price")
			 sProductString = sProductString & Sale_Price & "::"&Quantity&"::"&Item_Sku&"::"&Item_Name&"::|"
			 rs_store_tran.movenext
		loop
		rs_store_tran.close
		rs_Store.movenext
	loop
	rs_Store.Close
	sProductString = sProductstring & Tax & "::1::TAX::Tax|"&Shipping_Method_Price&"::1::SHIPPING::Shipping and Handling" 
   if Coupon_Amount<>0 then
      sProductString = sProductstring & "|-" & Coupon_Amount & "::1::COUPON::Coupon"
   end if
   %>
	<input type=hidden name="Products" value="<%= sProductString %>">
	<input type="hidden" name="language" value="English">
	<input type="hidden" name="ReturnCGI" value="<%= Return_Addr %>">
	<% rs_store.open "select * from store_customers where record_type=0 and cid="&cid&" and Store_Id="&Store_Id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
		<INPUT type="hidden" NAME="xxxName" value="<% response.write rs_store("First_name") & " " & rs_store("Last_name") %>">
		<INPUT type="hidden" NAME="xxxCompany" value="<%= rs_store("Company") %>">
		<INPUT type="hidden" NAME="xxxAddress" value="<%= rs_store("Address1") %>">
		<INPUT type="hidden" NAME="xxxCity" value="<%= rs_store("City") %>">
			<INPUT type="hidden" NAME="xxxProvince" value="<%= rs_store("State") %>">
<% Country = rs_store("Country") %>

		<INPUT type="hidden" NAME="xxxPostal" value="<%= rs_store("Zip") %>">
		<INPUT type="hidden" NAME="xxxPhone" value="<%= rs_store("Phone") %>">
		<INPUT type="hidden" NAME="xxxEmail" value="<%= rs_store("Email") %>">
		<INPUT type="hidden" NAME="xxxCCType" value="<%= Payment_Method %>">
		<% Zip = rs_Store("Zip") %>
	<% End If %>
	<% rs_store.close %>
	<% if Country = "Canada" then
		code = "CA"
		elseif Country = "United States" then
		code = "US"
		elseif Country = "Australia" then
		code = "AU"
		elseif Country = "Aruba" then
		code = "AW"
		elseif Country = "Belgium" then
		code = "BE"
		elseif Country = "Brazil" then
		code = "BR"
		elseif Country = "Cyprus" then
		code = "CY"
		elseif Country = "Germany" then
		code = "DE"
		elseif Country = "Spain" then
		code = "ES"
		elseif Country = "France" then
		code = "FR"
		elseif Country = "United Kingdom" then
		code = "GB"
		elseif Country = "Hong Kong" then
		code = "HK"
		elseif Country = "Japan" then
		code = "JP"
		elseif Country = "South Korea" then
		code = "KR"
		elseif Country = "Netherlands" then
		code = "NL"
		elseif Country = "Norway" then
		code = "NO"
		elseif Country = "New Zealand" then
		code = "NZ"
		elseif Country = "Singapore" then
		code = "SG"
		elseif Country = "Venezuela" then
		code = "VE"
		else
		code = "99"
		end if %>

	<INPUT type="hidden" NAME="xxxCountry" value="<%= code %>">
<% end sub %>

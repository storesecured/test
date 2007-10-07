<%


sub create_form_post_worldpay ()

    sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
    rs_Store.open sql_real_time,conn_store,1,1
    Do While Not Rs_Store.EOF
     select case Rs_store("Property")
	     case "WorldPay_InstId"
		     WorldPay_InstId = decrypt(Rs_store("Value"))
	     case "WorldPay_Currency"
		     WorldPay_Currency = decrypt(Rs_store("Value"))
      end select
     Rs_store.MoveNext
    Loop
    Rs_store.Close

    response.write "<form method='POST' action='https://select.worldpay.com/wcc/purchase' name='Payment' onSubmit=""return checkFields();""><input type=hidden name=instId value="&WorldPay_InstId&"><input type=hidden name='currency' value="&WorldPay_Currency&">"
end sub

sub create_form_content_worldpay ()
 %>
<input type="hidden" name="fixContact">
<input type="hidden" name="hideContact">
<input type=hidden name="cartId" value="<%= oid %>">

<% rs_store.open "Select * from Store_Purchases where Shopper_id ="&Shopper_ID&" AND oid = "&Oid&" and Store_id ="&Store_id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
	<% if rs_Store("recurring_fee") <> 0 then %>
	 <input type="hidden" name="futurePayType" value="regular">
	 <input type="hidden" name="startDelayUnit" value="1">
	 <input type="hidden" name="startDelayMult" value="<%= rs_store("recurring_days") %>">
	 <input type="hidden" name="intervalUnit" value="1">
	 <input type="hidden" name="intervalMult" value="<%= rs_store("recurring_days") %>">
	 <input type="hidden" name="normalAmount" value="<%= formatnumber(rs_store("recurring_total"),2) %>">
   <input type="hidden" name="option" value="0">
   <% end if %>
	<% end if %>
	<% rs_store.close %>
<input type=hidden name="amount" value="<%= formatnumber(GGrand_Total,2,,,0) %>">

  <% if Store_id = 1 then %>
  <input type="hidden" name="testMode" value="100">
  <% else %>
	<input type="hidden" name="testMode" value="0">
	<% end if %>
	
	<input type="hidden" name="cart_order_id" value="<%= oid %>">
	<input type="hidden" name="total" value="<%= formatnumber(GGrand_Total,2,0) %>">
	<% rs_store.open "select * from store_customers where record_type=0 and cid="&cid&" and Store_Id="&Store_Id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
		<input type=hidden name="name" value="<%= rs_store("First_Name") & " " & rs_store("Last_Name") %>">
		<input type=hidden name="address" value="<%= rs_store("Address1") %>">
		<input type=hidden name="postcode" value="<%= rs_store("Zip") %>">
		<input type=hidden name="tel" value="<%= rs_store("Phone") %>">
		<input type=hidden name="email" value="<%= rs_store("Email") %>">
		<input type=hidden name="fax" value="<%= rs_store("Fax") %>">
		<%
Country = rs_store("Country")
end if
rs_store.close

country_code=fn_country_code(country)
%>
	<input type=hidden name=country value="<%= country_code %>">
	<INPUT type="hidden" NAME="M_Store_Id" value="<%= Store_Id %>">
	<INPUT type="hidden" NAME="M_Shopper_Id" value="<%= Shopper_Id %>">


				
<% end sub %>

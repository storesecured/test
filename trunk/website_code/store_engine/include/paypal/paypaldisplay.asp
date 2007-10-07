<%

sub create_form_post_paypal ()

    sql_real_time = "exec wsp_real_time_property "&Store_Id&",4;"
    rs_Store.open sql_real_time,conn_store,1,1
    rs_Store.MoveFirst
    Do While Not Rs_Store.EOF
	    select case Rs_store("Property")
		     case "Pay_To"
			     PP_PAY_TO = decrypt(Rs_store("Value"))
		     case "Pay_Currency"
			     PP_Currency = decrypt(Rs_store("Value"))
	      end select
	     Rs_store.MoveNext
     Loop
	 
	 if PP_PAY_TO="" then
	    PP_PAY_TO = Store_Email
	 end if
	 response.write "<form method='POST' action='https://www.paypal.com/cgi-bin/webscr' name='Payment' onSubmit=""return checkFields();""><input type='hidden' name='business' value='"&PP_PAY_TO&"'>"
	 if PP_Currency = "" then
	     PP_Currency = "USD"
	 end if
    response.write "<input type='hidden' name='currency_code' value='"&PP_Currency&"'>"
    Rs_store.Close

end sub

sub create_form_content_paypal ()
 %>
      <% rs_store.open "Select * from Store_Purchases where Shopper_id ="&Shopper_ID&" AND oid = "&Oid&" and Store_id ="&Store_id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
	<% if rs_Store("recurring_fee") <> 0 then %>
      <input type="hidden" name="a1" value="<%= formatnumber(GGrand_Total,2) %>">
    	<input type="hidden" name="p1" value="<%= rs_store("recurring_days") %>">
    	<input type="hidden" name="t1" value="D">
    	<input type="hidden" name="a3" value="<%= formatnumber(rs_store("recurring_total"),2) %>">
    	<input type="hidden" name="p3" value="<%= rs_store("recurring_days") %>">
    	<input type="hidden" name="t3" value="D">
    	<input type="hidden" name="src" value="1">
	  <input type="hidden" name="sra" value="1">
      <input type="hidden" name="cmd" value="_xclick-subscriptions">
     <% else %>
		 <input type="hidden" name="cmd" value="_ext-enter">
      <input type="hidden" name="amount" value="<%= formatnumber(GGrand_Total,2) %>">
     <% end if %>
   <% End If %>
	<% rs_store.close %>


				<input type="hidden" name="redirect_cmd" value="_xclick">
				<input type="hidden" name="pal" value="UZNSQBQEC6NSW">
				<input type="hidden" name="mrb" value="R-3NR90971Y1052394E">
				<input type="hidden" name="bn" value="EasyStoreCreator.EasyStoreCreator">
				<input type="hidden" name="item_name" value="<%= store_name %> - Receipt <%= oid %>">
				<input type="hidden" name="item_number" value="<%= oid %>">
				<% Notify_Addr = Switch_Name&"include/paypal/paypal.asp?Shopper_Id="&Shopper_id %>
				<% Return_Addr = Switch_Name&"recipiet.asp" %>
				<% Cancel_Addr = Switch_Name&"error.asp?Message_id=81" %>
				<input type="hidden" name="return" value="<%= Return_Addr %>">
				<input type="hidden" name="notify_url" value="<%= Notify_Addr %>">
				<input type="hidden" name="cancel_return" value="<%= Cancel_Addr %>">
				<input type="hidden" name="no_note" value="1">
				<input type="hidden" name="no_shipping" value="0">

				<input type="hidden" name="custom" value="<%= Store_id %>">
				<% rs_store.open "select * from store_customers where record_type=0 and cid="&cid&" and Store_Id="&Store_Id, conn_store, 1, 1 %>
	         <% If not rs_Store.eof then 
            sphone=rs_store("Phone")
            sphone=replace(sPhone,"-","")
            sphone=replace(sPhone,"(","")
            sphone=replace(sPhone,")","")
            sphone=replace(sPhone,".","")
            sphone=replace(sPhone," ","")
            sphone=replace(sPhone,"/","")
            sphone=replace(sPhone,"*","")
            sphone=replace(sPhone,"+","")
            %>
					<INPUT type="hidden" NAME="first_name" value="<%= rs_store("First_name") %>">
					<INPUT type="hidden" NAME="last_name" value="<%= rs_store("Last_name") %>">
					<INPUT type="hidden" NAME="address1" value="<%= rs_store("Address1") %>">
					<INPUT type="hidden" NAME="city" value="<%= rs_store("City") %>">
					<INPUT type="hidden" NAME="zip" value="<%= rs_store("Zip") %>">
					<INPUT type="hidden" NAME="state" value="<%= rs_store("State") %>">
					<INPUT type="hidden" NAME="country" value="<%= fn_country_code(rs_store("Country")) %>">
					<INPUT type="hidden" NAME="night_phone_a" value="<%= left(sphone,3) %>">
					<INPUT type="hidden" NAME="night_phone_b" value="<%= mid(sphone,4,3) %>">
					<INPUT type="hidden" NAME="night_phone_c" value="<%= right(sphone,4) %>">
					<INPUT type="hidden" NAME="email" value="<%= rs_store("Email") %>">
				<% End If %>
				<% rs_store.close %>

<% end sub %>

<%

sub create_form_post_nochex ()

    sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
    rs_Store.open sql_real_time,conn_store,1,1
    rs_Store.MoveFirst
    Do While Not Rs_Store.EOF
    select case Rs_store("Property")
     case "nochex_email"
		     nochex_email = decrypt(Rs_store("Value"))
      end select
     Rs_store.MoveNext
    Loop
    Rs_store.Close
    'nochex_email="test1@nochex.com"
    response.write "<form method='POST' action='https://secure.nochex.com' name='Payment' onSubmit=""return checkFields();""><INPUT type='hidden' name='email' value='"&nochex_email&"'>"
    'cart page is http://help.nochex.com/?Action=Q&ID=148
end sub

sub create_form_content_nochex ()
 %>

	<% Notify_Addr = Switch_Name&"include/nochex/nochex.asp?Shopper_id="&Shopper_Id %>
	<% Return_Addr = Notify_Addr %>

	<INPUT type = "hidden" NAME="callback_url" value="<%=Notify_Addr %>">
	<INPUT type = "hidden" NAME="success_url" value="<%= Return_Addr %>">
	<INPUT type = "hidden" NAME="description" value="<%= store_name %> - Receipt <%= oid %>">
	<INPUT type = "hidden" NAME="amount" value="<%= formatnumber(GGrand_Total,2) %>">
	<INPUT type = "hidden" NAME="order_id" value="<%= oid %>">
	
        <% if 1=0 then %>
        <INPUT type = "hidden" NAME="test_transaction" value="100">
        <% end if %>

	<% rs_store.open "select * from store_customers where record_type=0 and cid="&cid&" and Store_Id="&Store_Id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
		<INPUT type="hidden" NAME="billing_fullname" value="<%= rs_store("First_name") %>&nbsp;<%= rs_store("Last_name") %>">
		<INPUT type="hidden" NAME="billing_address" value="<%= rs_store("Address1") %>">
		<INPUT type="hidden" NAME="billing_postcode" value="<%= rs_store("Zip") %>">
		<INPUT type="hidden" NAME="customer_phone_number" value="<%= rs_store("Phone") %>">
		<INPUT type="hidden" NAME="email_address" value="<%= rs_store("Email") %>">
        <INPUT type="hidden" NAME="hide_billing_details" value="1">

	<% End If %>
	<% rs_store.close %>
	
		<% rs_store.open "select * from store_customers where record_type=1 and cid="&cid&" and Store_Id="&Store_Id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
		<INPUT type="hidden" NAME="delivery_fullname" value="<%= rs_store("First_name") %>&nbsp;<%= rs_store("Last_name") %>">
		<INPUT type="hidden" NAME="delivery_address" value="<%= rs_store("Address1") %>">
		<INPUT type="hidden" NAME="delivery_postcode" value="<%= rs_store("Zip") %>">

	<% End If %>
	<% rs_store.close %>
<% end sub %>

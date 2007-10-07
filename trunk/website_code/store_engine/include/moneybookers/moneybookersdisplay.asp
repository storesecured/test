<%

sub create_form_post_moneybookers ()

    sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
    rs_Store.open sql_real_time,conn_store,1,1
    rs_Store.MoveFirst
    Do While Not Rs_Store.EOF
     select case Rs_store("Property")
	     case "mb_pay_to_email"
		     pay_to_email = decrypt(Rs_store("Value"))
	     case "mb_pay_currency"
		     pay_currency = decrypt(Rs_store("Value"))
      end select
     Rs_store.MoveNext
    Loop
    Rs_store.Close

    response.write "<form method='POST' action='https://www.moneybookers.com/app/payment.pl' name='Payment' onSubmit=""return checkFields();""> <INPUT type='hidden' NAME='currency' value='"&pay_currency&"'><INPUT type='hidden' NAME='pay_to_email' value='"&pay_to_email&"'>"

end sub

sub create_form_content_moneybookers ()
 %>

	<% Return_Addr = Site_Name&"include/moneybookers/moneybookers.asp?Shopper_Id="&shopper_id %>
	<% Other_Addr = Return_Addr %>

 <INPUT type="hidden" NAME="transaction_id" value="<%=oid %>">
 <INPUT type="hidden" NAME="return_url" value='<%= Other_Addr %>'>
 <INPUT type="hidden" NAME="cancel_url" value='<%=Return_Addr %>'>
 <INPUT type="hidden" NAME="status_url"  value='<%=Return_Addr %>'>
 <INPUT type="hidden" NAME="amount2_description" value="Total">
 <INPUT type="hidden" NAME="amount3_description" value="Shipping">
 <INPUT type="hidden" NAME="amount4_description" value="Taxes">
 <INPUT type="hidden" NAME="amount" value="<%= formatnumber(GGrand_Total,2) %>">

 <INPUT type="hidden" NAME="detail1_description" value="<%= store_name %>">
 <INPUT type="hidden" NAME="detail1_text" value="Receipt <%= oid %>">
 <INPUT type="hidden" NAME="merchant_fields" value="Store_Id">
 <INPUT type="hidden" NAME="Store_Id" value="<%= Store_Id %>">

	<% rs_store.open "select * from store_customers where record_type=0 and cid="&cid&" and Store_Id="&Store_Id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
		<INPUT type="hidden" NAME="firstname" value="<%= rs_store("First_name") %>">
		<INPUT type="hidden" NAME="lastname" value="<%= rs_store("Last_name") %>">
		<INPUT type="hidden" NAME="address" value="<%= rs_store("Address1") %>">
		<INPUT type="hidden" NAME="city" value="<%= rs_store("City") %>">
		<INPUT type="hidden" NAME="postal Code" value="<%= rs_store("Zip") %>">
		<INPUT type="hidden" NAME="country" value="<%= rs_store("Country") %>">
		<INPUT type="hidden" NAME="pay_from_email" value="<%= rs_store("Email") %>">
		<INPUT type="hidden" NAME="state" value="<%= rs_store("State") %>">

	<% End If %>
	<% rs_store.close %>

	<% rs_store.open "Select * from Store_Purchases where Shopper_id ="&Shopper_ID&" AND oid = "&Oid&" and Store_id ="&Store_id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
		<INPUT type="hidden" NAME="amount2" value="<%= formatnumber(Rs_store("Total"),2) %>">
		<INPUT type="hidden" NAME="amount3" value="<%= formatnumber(Rs_store("Shipping_Method_Price"),2) %>">
		<INPUT type="hidden" NAME="amount4" value="<%= formatnumber(Rs_store("Tax"),2) %>">
	<% End If %>
	<% rs_store.close %>

<% end sub %>

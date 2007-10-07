<%

sub create_form_post_pri ()

    sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
    rs_Store.open sql_real_time,conn_store,1,1
    rs_Store.MoveFirst
    Do While Not Rs_Store.EOF
    select case Rs_store("Property")
	    case "MerchantID"
		    id = decrypt(Rs_store("Value"))
	    case "MerchantKey"
		    key = decrypt(Rs_store("Value"))
     end select
    Rs_store.MoveNext
    Loop
    Rs_store.Close
    response.write "<form method='POST' action='https://webservices.primerchants.com/billing/TransactionCentral/processCCOnline.asp' name='Payment' onSubmit=""return checkFields();""><INPUT type='hidden' name='MerchantID' value='"&id&"'><INPUT type='hidden' name='RegKey' value='"&key&"'>"

end sub

sub create_form_content_pri ()

	Return_Addr = Switch_Name&"include/pri/pri.asp" %>
	<input type="hidden" name="Amount" value="<%= formatnumber(GGrand_Total,2) %>">
	<input type="hidden" name="REFID" value="<%= oid %>">
	<input type="hidden" name="CCRURL" value="<%= Return_Addr %>">
	<input type="hidden" name="USER1" value="<%= Store_Id %>">
	<input type="hidden" name="USER2" value="<%= Shopper_Id %>" />

	<% if Payment_Method = "eCheck" then %>
	
	<% else %>
	<TR>
					<TD width="176" class='normal'>Name on Card</TD>
					<TD width="219" class='normal'><INPUT type="text" name="NameonAccount" size="28" maxlength=30><INPUT name=CardName_C type=hidden value="Re|String||||Name on Card" size="20"></TD>
						 
				</TR>
				<TR>
					<TD width="176" class='normal'>Card Number</font></TD>
					<TD width="219" class='normal'><INPUT type="text" name="AccountNo" size="28" maxlength=17><INPUT name=CC_NUM_C type=hidden value="Re|Integer||||Card Number" size="20"></TD>
				</TR>
				<% if Use_CVV2 then%>
				 <TR>
						<TD width="176" class='normal'>Card Code</TD>
						<TD width="219" class='normal'><table><tr><td>
                  <INPUT type="text" name="CVV2" size="4" onKeyPress="return goodchars(event,'0123456789')">
						</td><td><a href="JavaScript:goCardCode();" class=small><img src=images/mini_cvv2.gif border=0></a></td><td>
						<a href="JavaScript:goCardCode();" class=link>What's This?</a></td></tr></table></TD>
					</TR>
            
				<% end if %>
				<TR>
		<TD width="176" class='normal'>Expiration Date MM/YY</TD>
		<TD width="219" class='normal'>
			<select name="CCMonth" size="1">
		<option selected value="01">January</option>
		<option value="02">February</option>
		<option value="03">March</option>
		<option value="04">April</option>
		<option value="05">May</option>
		<option value="06">June</option>
		<option value="07">July</option>
		<option value="08">August</option>
		<option value="09">September</option>
		<option value="10">October</option>
		<option value="11">November</option>
		<option value="12">December</option>
		</select>
			<% Current_year = year(now()) %>
			<select name="CCYear" size="1">
				<% for go_year =	Current_year to Current_year + 10 %>
									<% go_year_two_digit = right(go_year,2) %> 
									<option value="<%= go_year_two_digit %>"><%= go_year %>
								<% next %>
			</select></TD>
	</TR>
	<% end if %>

				<TR>
					<TD width="176" class='normal'>&nbsp;</TD>
					<TD width="219" class='normal'>&nbsp;</TD>
				</TR>


	<% rs_store.open "select * from store_customers where record_type=0 and cid="&cid&" and store_id="&Store_Id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
		<INPUT type="hidden" NAME="AVSADDR" value="<%= left(rs_store("Address1"),30) %>">
		<INPUT type="hidden" NAME="AVSZIP" value="<%= left(rs_store("Zip"),10) %>">
		<INPUT type="hidden" NAME="Email" value="<%= left(rs_store("Email"),100) %>">
	<% End If %>
	<% rs_store.close %>



<% end sub %>

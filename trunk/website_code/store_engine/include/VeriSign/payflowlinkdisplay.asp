<%
sub create_form_post_payflowlink ()

    sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
    rs_Store.open sql_real_time,conn_store,1,1
    rs_Store.MoveFirst
    Do While Not Rs_Store.EOF
      select case Rs_store("Property")
        case "pl_Login"
	        LOGIN = decrypt(Rs_store("Value"))
        case "pl_Partner"
	        PARTNER = decrypt(Rs_store("Value"))
    end select
    Rs_store.MoveNext
    Loop
    Rs_store.Close
    response.write "<form method='POST' action='https://payments.verisign.com/payflowlink' name='Payment' onSubmit=""return checkFields();""><input type=hidden name=LOGIN value="&LOGIN&"><input type=hidden name=PARTNER value="&PARTNER&"><input type=hidden name=USER1 value="&Shopper_ID&"><input type=hidden name=USER2 value="&Store_id&">"

end sub

sub create_form_content_payflowlink ()
 %>
		 <Font color=red>After entering your payment information you will be redirected to our processor's site <B>CLICK CONTINUE</b> to return to our store to finish your order processing and print a receipt.</font>
		 <TR>
					<TD width="176" class='normal'>&nbsp;</font></TD>
					<TD width="219" class='normal'>&nbsp;</font></TD>
				</TR>
				<TR>
					<TD width="176" class='normal'>Name on Card</TD>
					<TD width="219" class='normal'><INPUT type="text" name="CardName" size="28"><INPUT name=CardName_C type=hidden value="Re|String||||Name on Card" size="20"></TD>
				</TR>
				<TR>
					<TD width="176" class='normal'>Card Number</font></TD>
					<TD width="219" class='normal'><INPUT type="text" name="CARDNUM" size="28"><INPUT name=CARDNUM_C type=hidden value="Re|Integer||||Card Number" size="20"></TD>
										
										  </TR>
				<% if Use_CVV2 then %>
					<TR>
						<TD width="176" class='normal'>Card Code</TD>
						<TD width="219" class='normal'><table><tr><td>
                  <INPUT type="text" name="CSC" size="4" onKeyPress="return goodchars(event,'0123456789')">
						</td><td><a href="JavaScript:goCardCode();" class=small><img src=images/images_themes/mini_cvv2.gif border=0></a></td><td>
						<a href="JavaScript:goCardCode();" class=link>What's This?</a></td></tr></table></TD>
					</TR>

					<% end if %>
				<TR>
					<TD width="176" class='normal'>Expiration Date MM/YY</TD>
					<TD width="219" class='normal'><INPUT type="text" name="EXPDATE" size="5"><INPUT name=EXPDATE_C type=hidden value="Re|String||||Expiration Date" size="20">
						</TD>
				</TR>

				<TR>
					<TD width="176" class='normal'>&nbsp;</TD>
					<TD width="219" class='normal'>&nbsp;</TD>
				</TR>

				<% rs_store.open "select * from store_customers where record_type=0 and cid="&cid&" and Store_Id="&Store_Id, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
					<INPUT type="hidden" NAME="NAME" value="<%= rs_store("First_name") %> <%= rs_store("Last_name") %>">
					<INPUT type="hidden" NAME="ADDRESS" value="<%= rs_store("Address1") %>">
					<INPUT type="hidden" NAME="CITY" value="<%= rs_store("City") %>">
					<INPUT type="hidden" NAME="ZIP" value="<%= rs_store("Zip") %>">
					<INPUT type="hidden" NAME="COUNTRY" value="<%= rs_store("Country") %>">
					<INPUT type="hidden" NAME="PHONE" value="<%= rs_store("Phone") %>">
					<INPUT type="hidden" NAME="FAX" value="<%= rs_store("fax") %>">
					<INPUT type="hidden" NAME="EMAIL" value="<%= rs_store("Email") %>">
					<INPUT type="hidden" NAME="STATE" value="<%= rs_store("State") %>">
						  <% End If %>
				<% rs_store.close %>


				<% rs_store.open "Select * from Store_Purchases where Shopper_id ="&Shopper_ID&" AND oid = "&Oid&" and Store_id ="&Store_id, conn_store, 1, 1 %>
				<% If not rs_Store.eof then %>
					<INPUT type="hidden" NAME="NAMETOSHIP" value="<%= rs_store("ShipFirstname") %> <%= rs_store("ShipLastname") %>">
					<INPUT type="hidden" NAME="ADDRESSTOSHIP" value="<%= rs_store("ShipAddress1") %>">
					<INPUT type="hidden" NAME="CITYTOSHIP" value="<%= rs_store("ShipCity") %>">
					<INPUT type="hidden" NAME="ZIPTOSHIP" value="<%= rs_store("ShipZip") %>">
					<INPUT type="hidden" NAME="COUNTRYTOSHIP" value="<%= rs_store("ShipCountry") %>">
					<INPUT type="hidden" NAME="PHONETOSHIP" value="<%= rs_store("ShipPhone") %>">
					<INPUT type="hidden" NAME="FAXTOSHIP" value="<%= rs_store("Shipfax") %>">
					<INPUT type="hidden" NAME="EMAILTOSHIP" value="<%= rs_store("ShipEmail") %>">
					<INPUT type="hidden" NAME="STATETOSHIP" value="<%= rs_store("ShipState") %>">
               <INPUT type="hidden" NAME="TAX" value="<%= formatnumber(rs_Store("Tax"),2) %>">
               <INPUT type="hidden" NAME="SHIPAMOUNT" value="<%= formatnumber(rs_store("Shipping_Method_Price"),2) %>">

				<% End If %>
				<% rs_store.close %>

				<INPUT type="hidden" NAME="CUSTID" value="<%= cid %>">
				<INPUT type="hidden" NAME="INVOICE" value="<%= oid %>">
				<INPUT type="hidden" NAME="METHOD" value="CC">
				<INPUT type="hidden" NAME="PONUM" value="<%= Cust_PO %>">
				<INPUT type="hidden" NAME="ORDERFORM" value="False">
				<INPUT type="hidden" NAME="SHOWCONFIRM" value="False">
				<INPUT type="hidden" NAME="AMOUNT" value="<%= formatnumber(GGrand_Total,2) %>">
				<% If Auth_Capture then
					v_type = "S"
				else
					v_type = "A"
				end if %>
				<INPUT type="hidden" NAME="TYPE" value="<%= v_type %>">
				
<% end sub %>



<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include virtual="common/crypt.asp"-->


<%
sTitle = "Overdue Payment"
thisRedirect = "overdue.asp"
sMenu="account"
createHead thisRedirect
 %>



		<TR bgcolor='#FFFFFF'><td>
		<% if Custom_Description <>"Manual Payment" then %>
		   Payment in the amount of $<%= Custom_Amount %> is due for <%= Custom_Description %>

      <% else %>
                <% sql_select = "select card_number,payment_method from Sys_Billing where Store_Id = "&Store_Id
                rs_store.open sql_select, conn_store, 1, 1
                if not rs_store.eof then
                        Card_Number=right(decrypt(rs_store("Card_Number")),4)
                        Payment_Method=rs_store("Payment_Method")
                        if Payment_Method="paypal" then
                                sTextMessage=" with PayPal"
                        else
                                sTextMessage=" to the "&Payment_Method&" ending in "&Card_Number
                        end if
                end if
                %>
		The automatic payment <%=sTextMessage %> was unsuccessful, your payment is now <%= Overdue_Payment %> days overdue.

		<% end if %>

		<% if Custom_Amount < 999 then %>

		<BR><BR><a href=custom.asp class=link><B>Pay Now</b></a>

		<% end if %>

		<a href=cancel_store.asp class=link><BR><BR>Cancel Store</a>


		<a href=support_list.asp class=link><BR><BR>Contact support</a>

               </td></tr>


<% createFoot thisRedirect, 0 %>



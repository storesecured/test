<!--#include file="cart_display.asp"-->
<%
sub create_receipt(iPrintable)

sql_sel_ggtotal = "select sum(grand_total) as Total,sum(total)-sum(coupon_amount) as sub_total,max(shopper_id) as shopper_id,sum(cast(purchase_completed as int)) as purchase_completed from store_purchases WITH (NOLOCK) where store_id="&Store_Id&" AND (masteroid="&oid&" or oid="&oid&")"
fn_print_debug sql_sel_ggtotal
rs_Store.open sql_sel_ggtotal ,conn_store,1,1
rs_Store.movefirst
	if not rs_store.eof then
		if Rs_store("Total")<>"" then
			GGrand_Total = Rs_store("Total") 
			TTotal = Rs_store("Sub_Total")
			purchase_completed=rs_store("purchase_completed")
			shopper_id=rs_store("shopper_id")
		end if
	end if
rs_Store.Close

if purchase_completed=0 then
	 'purchase isnt done yet, wait and try again
	 iWait=fn_wait(10)
	 iTry=fn_get_querystring("Reload")
	 if iTry="" then
	 	iTry=0
	 end if
	 iTry=iTry+1
	 if iTry>=5 then
	 	fn_error "An error has occurred.  Your order was not completed."
	 end if
	 fn_redirect Switch_Name&"recipiet.asp?Reload="&iTry
end if

Session("TTotal") = TTotal
Session("GGrand_Total") = GGrand_Total

If Payment_Method = "Fax" then %>
<!-- start default fax instructions -->
	<p align="center" class=big>Facsimile Order Form</font><br>
	<font class='normal'>Please print this page and fax it to the number below.</font></p>
	
	<table border="1" cellPadding="3" cellSpacing="0" width="100%">
		<tr> 
			<td class='normal'>To <b><%= Store_Company %></b></td>
			<td class='normal'>From <b><%= first_name %> &nbsp;  <%= last_name %></b></td>
		</tr> 
	
		<tr>
			<td class='normal'>Fax <b><%= Store_Fax %></b></td>
			<td class='normal'>Total Pages _________</td>
		</tr>
		
		<tr>
			<td class='normal'>Order ID <b><%= oid %></b></td>
			<td class='normal'>Date <b><%= Now() %></b></td>
		</tr>
		<% if not isNull(Invoice_Id) then %>
		<tr>
			 <td><B>Invoice ID </B> <%= Invoice_ID %></td>
		</tr>
		<% end if %>

	</table>

	<br>
	<hr align="left" width="100%">
	</p>
<!-- end default fax instructions -->
<% End If %>

<% if iPrintable <>1 then %>
<table width="100%">
<tr><td><center><a href=print_receipt.asp?Soid=<%= Oid%> target=_blank class=link>Printable Receipt</a></center><BR>
</td></tr></table>
<% end if %>

 
<%
response.Write fn_display_invoice (oid,shopper_id,1,0)

end sub

%>

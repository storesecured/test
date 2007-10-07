<!--#include file="include/header.asp"-->
<!--#include file="include/receipt_include.asp"-->

<%
'GTE THE SPECIFIC OID FROM REQUEST QUERY STRING
oid = fn_get_querystring("Soid")
'LOAD ORDER DETAILS FROM STORE PURCHASES TABLE

if not isNumeric(oid) or oid="" then
	fn_redirect Switch_Name&"error.asp?message_id=1"
end if
create_receipt 0

' ORDER NOTES
			sql_ordernotes = "select sys_created, notes from store_purchases_notes WITH (NOLOCK) where  store_id=" &store_id&" and order_id = "&oid&" and status = 0 order by sys_created desc"
			rs_Store.open sql_ordernotes,conn_store,1,1
			if not rs_Store.eof then
			%>
			<hr width="600" align="left">
			<B>Notes</b>
                        <table cellspacing="0" cellpadding="0" border="1"  width="100%" >
				<% while not rs_Store.eof %>
				<tr>
					<td width="30%"><%=FormatDateTime(rs_Store("sys_created"))%></td>
					<td width="70%">&nbsp;&nbsp;<%=rs_Store("notes")%></td>
				</tr>

			<% 
			rs_Store.movenext
			wend
			%>
						</table>
			<% end if
			rs_store.close


if Show_shipping then

	'RETRIEVE SHIPPING DETAILS
	sql_select_shippments = "Select * from Store_Purchases_shippments WITH (NOLOCK) where oid = "&oid&" and Store_id ="&Store_id&" order by sys_created desc"
        rs_Store.open sql_select_shippments,conn_store,1,1
	if rs_Store.eof then

	else %>
		<HR>
                <table border="0" width="32%">
	<tr>
		<td width="100%" class='normal'><a name="Shipping Information">Shipping Information</a></td>
	</tr>
		</table>

		<table border="1" cellspacing="0" cellpadding=4 width=100%>
	<tr>
		<td class='normal'><b>Shipped</b></td>
		<td class='normal'><b>Company</b></td>
		<td class='normal'><b>Date</b></td>
		<td class='normal'><b>Tracking ID</b></td>
		<td class='normal'><b>Notes</b></td>
	</tr>

	<% Do While not rs_store.eof %>
		<tr>
		<td align="middle" class='normal'><%= Rs_store("ShippedPr") %></td>
			<td class='normal'><%= Rs_store("Shipping_Company") %></td>
			<td align="middle" class='normal'><%= Rs_store("Shipped_Date") %></td>
			<td align="middle" class='normal'><%= Rs_store("Tracking_ID") %></td>
			<td align="middle" class='normal'><%= Rs_store("Shippment_notes")%></td>
		</tr>
		<% rs_store.MoveNext %>
	<% loop %>
   </table>
	<% end if
	rs_store.close
	end if  'SHOW/HIDE SHIPPING %>



<!--#include file="include/footer.asp"-->

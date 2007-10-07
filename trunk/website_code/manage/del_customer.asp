<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%	
'GET THE CUSTOMER ID
CCid = Request.QueryString("CCid")
'CHECK IF THE CUSTOMER HAS ORDERS
sql_check ="select distinct Store_Purchases.cid, Store_Purchases.oid From Store_Purchases INNER JOIN Store_Transactions ON Store_Purchases.OID = Store_Transactions.OID Where Store_Purchases.ccid = "&ccid&" and Store_Purchases.store_id = "&store_id&" and store_transactions.Store_id="&Store_Id
rs_Store.open sql_check,conn_store,1,1
if rs_Store.eof = false then
	' CUSTOMER HAS VALID ORDERS, WARN THE ADMIN
	Error_code = 1
else
	Error_code = 0
end if

if	Error_code = 0 then
	sql_delete = "delete from Store_customers WHERE CCID="&CCID&" AND store_id = "&store_id
	conn_store.Execute sql_delete
	sTitle = "Delete Success"
else
	sTitle = "Delete Denied"
end if


sFormName = "del_customer"
thisRedirect = "del_customer.asp"
sMenu="reports"
createHead thisRedirect

%>


	<tr>
		<td width="100%" colspan="3" height="74">
			<table border="0" width="95%" height="25">
				<% If Error_code = 1 then %>
					<tr>
					<td width="100%" height="1">Customer cannot be deleted because there are orders relating to that customer.
							<a class="link" href="orders.asp?View_Cid=<%= CCID %>">Click here to view orders</a></td>
					</tr>
				<% Else  %>	  
					<tr>
						<td width="100%" height="1">Customer <%= CCID %> deleted sucessfully<br><br><a class="link" href=my_customer_base.asp>Click here to return to customers.</a></td>
					</tr>
				<% End If %> 
		</table>
		</td>
	</tr>
<% createFoot thisRedirect, 0%>

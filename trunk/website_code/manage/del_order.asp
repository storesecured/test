<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%	
'GET THE ORDER ID AND DELETE IT
Oid = Request.QueryString("Oid")
sql_delete = "delete from Store_purchases WHERE OID="&OID&" AND store_id = "&store_id
conn_store.Execute sql_delete
'ALSO DELETE THE RELATED RECORDS FROM STORE_TRANSACTIONS TABLE
sql_delete = "delete from Store_transactions WHERE OID="&OID&" AND store_id = "&store_id
conn_store.Execute sql_delete

sFormAction = ""
sName = ""
sTitle = "Delete Successfull"
thisRedirect = "del_order.asp"
createHead thisRedirect
%>

				<tr>
					<td width="100%" height="1">Order <%= oid %> deleted successfully.<br><br><a class="link" href=orders.asp>Click here to return to orders.</a></td>
				</tr>
<% createFoot thisRedirect, 0%>

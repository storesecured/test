<!--#include file="Global_Settings.asp"-->
<HTML>
<title>StoreSecured Order Add Notes - #<%=request("oID") %></title>
</HEAD>
<body>
<script language="javascript">
function addNotes()
{ 
   document.frmAddNote.action="addOrderNotes.asp?action=add"
   document.frmAddNote.submit()	
   opener.window.location="order_details.asp?Id=<%= request.querystring("oid") %>"

}

</script>
<%
orderId = request("oID")
storeId = store_id

if trim(request("action")) = "add" then

orderId = request.form("hidOrderId")

orderNotes = trim(Request.Form("ordernotes"))
orderNotes=replace(replace(orderNotes,"<OBJ_TEXTAREA_START","<TEXTAREA"),"<OBJ_TEXTAREA_END","</TEXTAREA")
orderNotes = Replace(orderNotes,"'","''")

Show_NotesStatus = Request.Form("Show_NotesStatus")
if Show_NotesStatus = "" then
Show_NotesStatus = 0
else
Show_NotesStatus = 1
end if

on error goto 0
sql_insert = " insert into store_purchases_notes (Store_ID, Order_id, Notes,Status ) values "&_
			 " ("&store_id&","&orderid&",'"&ordernotes&"',"&Show_NotesStatus&")"
conn_store.Execute sql_insert
response.write "Note added"
end if
%>
<form name="frmAddNote" action="addOrderNotes.asp" method="post">
<input type="hidden" name="hidOrderId" value="<%=orderId%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td>			
			<table border=0 cellspacing=1 cellpadding=3 width="100%" > 
			<tr>

			<td align="left" height="25" valign=top>
			<font face="Arial" size="2">	
			<B>&nbsp;Add Notes to Order Id <%=orderid%></b>
			</font>
			</td>				
			</tr>	
			<tr BGCOLOR="#ffffff">
			<td align="left">
			<TEXTAREA cols=60 rows=12 name="ordernotes" style="background:#FFFFCC"></TEXTAREA>
			</td>
			</tr>

			<tr BGCOLOR="#ffffff">
			<td align="left">
			<input type=checkbox value="-1" checked name="Show_NotesStatus"> Internal (Only to Store Admininstrator)
			</td>
			</tr>

			
			</table>
		</td>
	</tr> 
</table>
	
	<tr height="0">
	<td width="1%"></td>
	<td align="left">
	<input type=button onclick="javascript:addNotes()" value="Save" id=button2 name=button2>	
	<input type=button onclick="window.close()" value="Close" id=button2 name=button3>
	</td>
	</tr>	

</td>
</tr>
</table>
</form>
</body>
</html>

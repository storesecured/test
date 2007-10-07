<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<% 
BID = "-1"
use_date = ""
BID = request.querystring("ID")
if BID = "" then
   BID = request.querystring("BID")
end if

if BID<>"" then
	if not isNumeric(BID) then
	   Response.Redirect "admin_error.asp?message_id=1"
	end if
	sql_sel = "select * from store_pin where store_id="&store_id&" and itemdown_id="&bid
	rs_store.open sql_sel, conn_store, 1, 1
	     if rs_store.bof or rs_store.eof then
		Response.Redirect "admin_error.asp?message_id=1"
	     end if
		 itemdown_id= rs_store("itemdown_id")
		 item_sku= rs_store("item_sku")
		 pin= rs_store("pin")
		 other_info= rs_store("other_info")
		 pin_used = rs_store("pin_used")
		 use_date = rs_store("use_date")
		 oid = rs_store("oid")
		 cid = rs_store("cid")
	rs_store.close

	if pin_used then
		Enabledchecked = "checked"
	else
		Enabledchecked = ""
	end if
end if

If BID="" then
   sTitle = "Add Pin"
   sFullTitle = "Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=edit_pin1.asp class=white>Pins</a> > Add"
Else
        sTitle = "Edit Pin - "&pin
        sFullTitle = "Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=edit_pin1.asp class=white>Pins</a> > Edit - "&pin

End If

sCommonName="Pin"
sCancel="edit_pin1.asp"
sFormAction = "pin_action.asp"
thisRedirect = "pin_add.asp"
sMenu = "inventory"
sQuestion_Path = "inventory/edit_pins.htm"
createHead thisRedirect
%>


<% If BID="" then %>
	<input type="hidden" name="Action" value="Add">
<% Else %>
	<input type="hidden" name="Action" value="Edit">
	<input type="hidden" name="BID" value="<%= BID %>">
<% End If %>
				

				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Pin Used</b></td>
					<td width="90%" class="inputvalue"><input class="image" type="checkbox" name="pin_used" <%= Enabledchecked %> size="30">
					<% small_help "pin_used" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Other Info</b></td>
					<td width="90%" class="inputvalue"><input type="text" maxlength=100 name="other_info" value="<%= other_info %>" size="30">
					<INPUT type="hidden"  name="other_info_C" value="Re|String|0|100|||Other Info">
					<% small_help "other_info" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Pin</b></td>
					<td width="90%" class="inputvalue"><input type="text" maxlength=20 name="pin" value="<%= pin %>" size="30">
					<INPUT type="hidden"  name="pin_C" value="Re|String|0|20|||Pin">
					<% small_help "pin" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Item Sku</b></td>
					<td width="90%" class="inputvalue"><input type="text" maxlength=50 name="item_sku" value="<%= item_sku %>" size="30">
					<INPUT type="hidden"  name="item_sku_C" value="Re|String|0|50|||Item Sku">
					<% small_help "item_sku" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Order ID</b></td>
					<td width="90%" class="inputvalue"><input type="text" maxlength=50 name="oid" value="<%= oid %>" size="30">
					<INPUT type="hidden"  name="oid_C" value="Op|Integer|||||Order ID">
					<% small_help "oid" %></td>
				</tr>

				
<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("item_sku","req","Please enter a SKU for this Pin.");

</script>

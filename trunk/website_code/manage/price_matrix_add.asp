<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%

Item_Id = Request.QueryString("Item_Id")

sql_select="SELECT Item_Name,Use_Price_By_Matrix FROM Store_Items WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" ORDER BY Item_Name;"
rs_store.open sql_select,conn_store,1,1 
Item_Name = Rs_store("Item_Name")
Use_Price_By_Matrix = Rs_store("Use_Price_By_Matrix")
rs_store.close

'ADD THE PRICE MATRIX
if Request.Form("Form_Name") = "Add_Price_Matrix" then
	'ERROR CHECKING
	If Form_Error_Handler(Request.Form) <> "" then 
		Error_Log = Form_Error_Handler(Request.Form)
		%><!--#include file="Include/Error_Template.asp"--><%
		response.end
	else
		'INSERT THE PRICE MATRIX RECORD
		Matrix_Low = Request.Form("Matrix_Low")
		Matrix_High = Request.Form("Matrix_High")
		Item_Price = Request.Form("Item_Price")
		Item_Id = Request.Form("Item_Id")
		Item_Sku = "Null"

		sql_ins="insert into Store_items_Price_Matrix (Store_ID,Item_Id,Matrix_low,Matrix_high,Item_Price) values ("&Store_ID&","&Item_Id&","&Matrix_low&","&Matrix_high&","&Item_Price&")"

		conn_store.Execute sql_ins
		response.redirect "Price_Matrix_List.asp?Item_Id="&Item_Id
	end if
end if

sInstructions="Matrix of Prices by Quantity Ordered<br>Add a quantity range and price<br>For example, buy 4 - 10 items and price is 3.75<BR>Quantity Between = 4 - 10, Item Price = 3.75<BR>Use -1 as the high to specify any quantity"

sFormAction = "price_matrix_add.asp"
sName = "Add_Class_1"
sFormName = "Add_Price_Matrix"
sTitle = "Add Price Matrix - "&Item_Name
sCancel="price_matrix_list.asp?Item_Id="&Item_Id&"&"&sAddString
sFullTitle ="Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > <a href="&sCancel&" class=white>Price Matrix</a> > Add - "&Item_Name
sCommonName="Price Matrix"
sSubmitName = "Add_Price_Matrix"
thisRedirect = "price_matrix_add.asp"
sMenu="inventory"
createHead thisRedirect

if Service_Type < 7  then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.
	</td></tr>
   <% createFoot thisRedirect, 0%>
<% else %>
<% if Use_Price_By_Matrix=0 then %>
<TR bgcolor='red'><TD colspan=3>
<font class=white>Quantity discount is currently not enabled for this item.  Any price matrix's created will not be applied until quantity discount is enabled for this item.  Go to the Advanced Edit Page, Options tab to enable Qty Discounts.</font>
</td></tr>
<% end if %>

				<input type="hidden" name="Item_Id" value="<%= Item_Id %>" size="30">	

				<tr bgcolor='#FFFFFF'>
				<td class="inputname">Quantity Between</td>
					<td class="inputvalue">
						<input type="text" name="Matrix_Low" size="10" onKeyPress="return goodchars(event,'0123456789.')">
						<input type="hidden" name="Matrix_Low_C" value="Re|Integer||||Quantity Low">
						and
						<input type="text" name="Matrix_High" size="10" onKeyPress="return goodchars(event,'0123456789.-')">
						<input type="hidden" name="Matrix_High_C" value="Re|Integer||||Quantity High">
						<% small_help "Quantity Between" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td class="inputname">Item Price</td>
					<td class="inputvalue">
						<%= Store_Currency %><input type="text" name="Item_Price" size="10" onKeyPress="return goodchars(event,'0123456789.')">
						<input type="hidden" name="Item_Price_C" value="Re|Integer||||Item Price">
						<% small_help "Item Price" %></td>
				</tr>
	 


<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Matrix_Low","req","Please enter a low quantity.");
 frmvalidator.addValidation("Matrix_High","req","Please enter a high quantity.");
 frmvalidator.addValidation("Item_Price","req","Please enter a item price.");

</script>
<% end if %>

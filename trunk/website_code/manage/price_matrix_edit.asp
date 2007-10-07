<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%


if id = "" then
	Item_Id = Request.QueryString("Item_Id")
	id= Request.QueryString("Id")
	if id="" or isnull(id)  then
		id= Request.form("Id")
	end if

end if

sql_select="SELECT Item_Name,Use_Price_By_Matrix FROM Store_Items WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" ORDER BY Item_Name;"
rs_store.open sql_select,conn_store,1,1 
Item_Name = Rs_store("Item_Name")
Use_Price_By_Matrix = Rs_store("Use_Price_By_Matrix")
rs_store.close

		sql_select="SELECT * FROM Store_items_Price_Matrix WHERE Store_id="&Store_id&" AND Item_Id = "&Item_Id & " and Price_Matrix_id="&id
	'	Response.Write "<BR>sql_select="& sql_select
		rs_store.open sql_select,conn_store,1,1 
		rs_store.MoveFirst
		
		Matrix_low=rs_store("Matrix_low")
		Matrix_high=rs_store("Matrix_high")
		item_price = rs_store("Item_Price")
		id=rs_store("Price_Matrix_id")
		'Response.Write "<BR><BR> test=" & id
		rs_store.close
		'Response.End 
'================
	'Response.Write "<BR><BR>form name= " & Request.Form("Form_Name")
	'ADD THE PRICE MATRIX
	
	if Request.Form("Form_Name") = "Edit_Price_Matrix" then
		'ERROR CHECKING
		Response.Write " error in here one "
		'Response.End 
		If Form_Error_Handler(Request.Form) <> "" then 
			Response.Write " error in here  two "
			  
			Error_Log = Form_Error_Handler(Request.Form)
			%><!--#include file="Include/Error_Template.asp"--><%
			'response.end
		else
		Response.Write " error in here  theree"
		
		'INSERT THE PRICE MATRIX RECORD
			Matrix_Low = Request.Form("Matrix_Low")
			Matrix_High = Request.Form("Matrix_High")
			Item_Price = Request.Form("Item_Price")
			Item_Id = Request.Form("Item_Id")

		

			sql_upd="update Store_items_Price_Matrix set Matrix_low=" &Matrix_low& ",Matrix_high=" &Matrix_high& ",Item_Price=" &Item_Price & " where store_id=" &store_id & " and Item_Id=" & Item_Id & " and  Price_Matrix_id=" & id
			session("sql") = sql_upd
			conn_store.execute sql_upd
			response.redirect "Price_Matrix_List.asp?Item_Id="&Item_Id&"&op=edit&id="&id
		End if
	end if
	
'====================

	sFormAction = "price_matrix_edit.asp"
	sName = "Edit_Class_1"
	sFullTitle ="Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > <a href="&sCancel&" class=white>Price Matrix</a> > Edit - "&Item_Name
        sFormName = "Edit_Price_Matrix"
	sTitle = "Edit Price Matrix - "&Item_Name
	sSubmitName = "Edit_Price_Matrix"
	thisRedirect = "price_matrix_edit.asp"
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
<TR bgcolor='red'><TD>
<font class=white>Quantity discount is currently not enabled for this item.  Any price matrix's created will not be applied until quantity discount is enabled for this item.  Go to the Advanced Edit Page, Options tab to enable Qty Discounts.</font>
</td></tr>
<% end if %>


	<tr bgcolor='#FFFFFF'> 
		<td width="100%" colspan="3" height="23">
			<input OnClick=JavaScript:self.location="Price_Matrix_List.asp?Item_Id=<%=Item_Id%>&id=<%=Item_Id %>" name="Price_Matrix_List" type="button" class="Buttons" value="Back to Matrix List">
			<input OnClick=JavaScript:self.location="item_edit.asp?Item_Id=<%=Item_Id%>" name="Price_Matrix_List" type="button" class="Buttons" value="Back to Item">
			<br> Matrix of Prices by Quantity Ordered<br>
			Edit a quantity range and price<br>
			For example, buy 4 - 10 items and price is 3.75<BR>
			Quantity Between = 4 - 10, Item Price = 3.75<BR>
			Use -1 as the high to specify any quantity</td>
	</tr>
    

				<input type="hidden" name="Item_Id" value="<%= Item_Id %>" size="30" >	
				<input type="hidden" name="id" value="<%=id%>" size="30" >	

				<tr bgcolor='#FFFFFF'>
				<td class="inputname">Quantity Between</td>
					<td class="inputvalue">
						<input type="text" name="Matrix_Low" size="10" Value="<%=Matrix_low%>"  onKeyPress="return goodchars(event,'0123456789.')">
						<input type="hidden" name="Matrix_Low_C" value="Re|Integer||||Quantity Low">
						and
						<input type="text" name="Matrix_High" size="10" Value="<%=Matrix_high%>"  onKeyPress="return goodchars(event,'0123456789.-')">
						<input type="hidden" name="Matrix_High_C" value="Re|Integer||||Quantity High">
						<% small_help "Quantity Between" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td class="inputname">Item Price</td>
					<td class="inputvalue">
						<%= Store_Currency %><input type="text" name="Item_Price" size="10" Value="<%=item_price%>"  onKeyPress="return goodchars(event,'0123456789.')">
						<input type="hidden" name="Item_Price_C" value="Re|Integer||||Item Price">
						<% small_help "Item Price" %></td>
				</tr>
	 
<% createFoot thisRedirect, 1%>


<% end if %>

<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

Item_Id = Request.QueryString("Item_Id")
if isNumeric(Item_Id) then
else
	Response.Redirect "admin_error.asp?message_id=1"
end if
sql_select="SELECT Item_Name,Use_Price_By_Matrix FROM Store_Items WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" ORDER BY Item_Name;"
rs_store.open sql_select,conn_store,1,1 
Item_Name = Rs_store("Item_Name")
Use_Price_By_Matrix = Rs_store("Use_Price_By_Matrix")
rs_store.close

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_items_Price_Matrix"
myStructure("TableWhere") = "Item_Id="&item_Id
myStructure("ColumnList") = "Price_Matrix_id,matrix_low,matrix_high,item_price"
myStructure("DefaultSort") = "item_price"
myStructure("PrimaryKey") = "Price_Matrix_id"
myStructure("Level") = 7
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Menu") = "inventory"
myStructure("FileName") = "price_matrix_list.asp?Item_Id="&Item_Id
myStructure("FormAction") = "price_matrix_list.asp?Item_Id="&Item_Id
myStructure("Title") = "Item Price Matrix - "&Item_Name
myStructure("FullTitle") = "Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > Price Matrix - "&Item_Name
myStructure("CommonName") = "Price Matrix"
myStructure("NewRecord") = "price_matrix_add.asp?Item_Id="&Item_Id
myStructure("EditRecord") = "price_matrix_edit.asp?Item_Id="&Item_Id
myStructure("Heading:Price_Matrix_id") = "PK"
myStructure("Heading:matrix_low") = "Low"
myStructure("Heading:matrix_high") = "High"
myStructure("Heading:item_price") = "Price"
myStructure("Format:matrix_low") = "INT"
myStructure("Format:matrix_high") = "INT"
myStructure("Format:item_price") = "CURR"
myStructure("Length:matrix_low") = "0"
myStructure("Length:matrix_high") = "0"


%>
<!--#include file="head_view.asp"-->

<TR bgcolor='#FFFFFF'><TD>
<input type="button" class="Buttons" value="Back to Item" name="Item_Attributes" OnClick=JavaScript:self.location="item_edit.asp?Item_Id=<%= Item_Id %>">
</td></tr>
<% if Use_Price_By_Matrix=0 then %>
<TR bgcolor='red'><TD>
<font class=white>Quantity discount is currently not enabled for this item.  Any price matrix's created will not be applied until quantity discount is enabled for this item.  Go to the Advanced Edit Page, Options tab to enable Qty Discounts.</font>
</td></tr>
<% end if %>
<!--#include file="list_view.asp"-->


<% createFoot thisRedirect, 0%>

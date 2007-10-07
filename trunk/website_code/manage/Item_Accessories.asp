<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<% 
Item_Id=request.querystring("Item_Id")

sql_select="SELECT Item_Name FROM Store_Items WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" ORDER BY Item_Name;"
rs_store.open sql_select,conn_store,1,1 
Item_Name = Rs_store("Item_Name")
rs_store.close

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "wsv_item_accessories"
myStructure("DeleteTable")="store_items_accessories"
myStructure("TableWhere") = "Item_Id="&Item_Id
myStructure("ColumnList") = "accessory_line_id,accessory_item_id,accessory_item_name"
myStructure("DefaultSort") = "accessory_item_name"
myStructure("PrimaryKey") = "accessory_line_id"
myStructure("Level") = 5
myStructure("EditAllowed") = 0
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Menu") = "inventory"
myStructure("FileName") = "item_accessories.asp?Item_Id="&Item_Id
myStructure("FormAction") = "item_accessories.asp?Item_Id="&Item_Id
myStructure("Title") = "Item Accessories - "&Item_Name
myStructure("FullTitle") = "Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > Accessories - "&Item_Name
myStructure("CommonName") = "Accessory"
myStructure("NewRecord") = "Item_accessories_add.asp?Item_Id="&Item_Id
myStructure("Heading:accessory_line_id") = "PK"
myStructure("Heading:accessory_item_name") = "Name"
myStructure("Heading:accessory_item_id") = "Id"
myStructure("Format:accessory_item_name") = "STRING"
myStructure("Format:accessory_item_id") = "STRING"
myStructure("Link:accessory_item_id") = "item_edit.asp?Id=THISFIELD"


%>
<!--#include file="head_view.asp"-->
 <TR bgcolor='#FFFFFF'>
		<td width="100%" colspan="3">
			<input type="button" class="Buttons" value="Back to Item" name="Back_To_Item" OnClick=JavaScript:self.location="Item_Edit.asp?Id=<%= request.querystring("Item_Id") %>">
	</tr>
<!--#include file="list_view.asp"-->



<% createFoot thisRedirect, 0%>

<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

Item_Id = Request.QueryString("Item_Id")
if len(Item_Id)  = 0 then
	response.Redirect("admin_error.asp?message_id=105")
end if
		if isNumeric(Item_Id) then
		else
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		
                sql_select="SELECT Item_Name FROM Store_Items WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" ORDER BY Item_Name;"
                rs_store.open sql_select,conn_store,1,1
                Item_Name = Rs_store("Item_Name")
                rs_store.close

		set myStructure=server.createobject("scripting.dictionary")
		myStructure("TableName") = "Store_items_Price_group"
		myStructure("TableWhere") = "Item_Id="&Item_Id
		myStructure("ColumnList") = "Price_group_id,customer_group,group_price"
		myStructure("DefaultSort") = "customer_group"
		myStructure("PrimaryKey") = "Price_group_id"
		myStructure("Level") = 7
		myStructure("EditAllowed") = 1
		myStructure("AddAllowed") = 1
		myStructure("DeleteAllowed") = 1
		myStructure("Menu") = "inventory"
		myStructure("FileName") = "price_group_list.asp?Item_Id="&Item_Id
		myStructure("FormAction") = "price_group_list.asp?Item_Id="&Item_Id
		myStructure("Title") = "Item Price Group - "&Item_Name
		myStructure("FullTitle") = "Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > Price Group - "&Item_Name
                myStructure("CommonName") = "Price Group"
		myStructure("NewRecord") = "price_group_add.asp?Item_Id="&Item_Id
		myStructure("EditRecord") ="price_group_add.asp?Item_Id="&Item_Id
		myStructure("Heading:Price_group_id") = "PK"
		myStructure("Heading:customer_group") = "Customer Group"
		myStructure("Heading:group_price") = "Group Price"

		myStructure("Format:customer_group") = "SQL"
		myStructure("Format:group_price") = "CURR"
		myStructure("Sql:customer_group") = "select group_name as customer_group from Store_Customers_Groups where group_id = THISFIELD and store_id="&Store_Id
		myStructure("Length:group_price") = "0"

		
	
		
		
		


		
		
%>
<!--#include file="head_view.asp"-->
 <TR bgcolor='#FFFFFF'>
		<td width="100%" colspan="3">
			<input type="button" class="Buttons" value="Back to Item" name="Back_To_Item" OnClick=JavaScript:self.location="Item_Edit.asp?Id=<%= Item_Id %>">
	</tr>
<!--#include file="list_view.asp"-->
<%
	%>

<% createFoot thisRedirect, 0%>

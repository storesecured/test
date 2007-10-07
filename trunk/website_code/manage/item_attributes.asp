<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

Item_Id = Request.QueryString("Item_Id")
if isNumeric(Item_Id) then
else
	Response.Redirect "admin_error.asp?message_id=1"
end if

sql_select="SELECT Item_Name FROM Store_Items WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" ORDER BY Item_Name;"
rs_store.open sql_select,conn_store,1,1 
Item_Name = Rs_store("Item_Name")
rs_store.close

sFlashHelp="attributes/attributes.htm"
sMediaHelp="attributes/attributes.wmv"
sZipHelp="attributes/attributes.zip"
sArticleHelp="AddingAttributes.htm"

sInstructions="Attributes are choices your customer can make for different options for an item.  Attributes will be shown as a pull down menu.  Common uses for attributes include a choice of sizes and/or colors."



set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_items_attributes"
myStructure("TableWhere") = "Item_Id="&request.querystring("Item_Id")
myStructure("ColumnList") = "attribute_id,attribute_class,attribute_value,attribute_price_difference,attribute_weight_difference"
myStructure("DefaultSort") = "attribute_class"
myStructure("PrimaryKey") = "attribute_id"
myStructure("Level") = 5
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Menu") = "inventory"
myStructure("FileName") = "item_attributes.asp?Item_Id="&request.querystring("Item_Id")
myStructure("FormAction") = "item_attributes.asp?Item_Id="&request.querystring("Item_Id")
myStructure("Title") = "Item Attributes - "&Item_Name
myStructure("FullTitle") = "Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > Attributes - "&Item_Name
myStructure("CommonName") = "Attribute"
myStructure("NewRecord") = "Item_Add_Attribute_Value.asp?Item_Id="&request.querystring("Item_Id")
myStructure("Heading:attribute_id") = "PK"
myStructure("Heading:attribute_class") = "Class"
myStructure("Heading:attribute_value") = "Value"
myStructure("Heading:attribute_price_difference") = "Price Diff"
myStructure("Heading:attribute_weight_difference") = "Weight Diff"
myStructure("Format:attribute_class") = "STRING"
myStructure("Format:attribute_value") = "STRING"
myStructure("Format:attribute_price_difference") = "CURR"
myStructure("Format:attribute_weight_difference") = "INT"
myStructure("Length:attribute_weight_difference") = "2"


%>
	        
<!--#include file="head_view.asp"-->
<TR bgcolor='#FFFFFF'><TD>
<input type="button" class="Buttons" value="Back to Item" name="Item_Attributes" OnClick=JavaScript:self.location="item_edit.asp?Item_Id=<%= Item_Id %>">
<input type="button" class="Buttons" value="Item Configs" name="Item_Attributes" OnClick=JavaScript:self.location="item_configs.asp?Item_Id=<%= Item_Id %>">
</td></tr>
<!--#include file="list_view.asp"-->


<% createFoot thisRedirect, 0%>

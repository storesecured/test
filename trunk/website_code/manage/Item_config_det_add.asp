<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

Item_Id = Request.QueryString("Item_Id")
sql_select="SELECT Item_Name FROM Store_Items WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" ORDER BY Item_Name;" 
rs_store.open sql_select,conn_store,1,1 
Item_Name = Rs_store("Item_Name")
rs_store.close

Config_ID = Request.QueryString("Config_Id")
sql_select="SELECT Config_Name FROM Store_Items_Configs WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" and config_id="&Config_ID
rs_store.open sql_select,conn_store,1,1 
Config_Name = Rs_store("Config_Name")
rs_store.close

op=Request.QueryString("Op")

if op="edit" then
	sqlConfig="select Name from store_items_configs_dets where config_id="&Config_ID&" and store_id=" & store_id&" and Det_ID="&request.queryString("Class")
	rs_Store.open sqlConfig,conn_store,1,1	
	Option_Name = rs_Store("Name") 
	rs_store.close
end if

sCancel="item_configs.asp?Item_Id="&Item_Id&"&"&sAddString
sFullTitle="Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > <a href="&sCancel&" class=white>Configurations</a> > "
sFormAction = "Inventory_Action.asp"
sFormName = "Add_Dets"
if op="edit" then
	sTitle = "Edit Configuration Option - "&Item_Name
	sFullTitle=sFullTitle&"Edit Option - "&Item_Name
else
	sTitle = "Add Configuration Option - "&Item_Name
	sFullTitle=sFullTitle&"Add Option - "&Item_Name
end if
sCommonName="Configuration Option"
sMenu="inventory"
thisRedirect = "item_config_det_add.asp"
createHead thisRedirect



Set rs_store2 = Server.CreateObject("ADODB.Recordset")
Set rs_store3 = Server.CreateObject("ADODB.Recordset")
if Service_Type < 5  then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.
	</td></tr>
   <% createFoot thisRedirect, 0%>
<% else 
sql_attribute_group =  "SELECT Attribute_class FROM store_items_attributes WHERE (((Store_id)="&Store_id&") AND ((Item_Id)="&Item_Id&")) GROUP BY Attribute_class ORDER BY store_items_attributes.Attribute_class;"
rs_store.open sql_attribute_group,conn_store,1,1

%>
<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Old_Option_Name" value="<%= Option_Name %>">
<input type="hidden" name="Config_ID" value="<%= Config_ID %>">
<input type="hidden" name="Item_Id" value="<%=Item_Id%>">

	
    
	<tr bgcolor='#FFFFFF'>
					<td class="inputname" width="20%">Name</td>
					<td class="inputvalue" width="80%">
						<input type="text" name="Option_Name" value="<%= Option_Name %>" size="60" maxlength="50">
						<INPUT type="hidden" name="Option_Name_C" value="Re|String||||">
                              <% small_help "Name" %></td>
				</tr>
			
				<% sql_sel_attribs = "select distinct Attribute_class from store_items_attributes where store_id="&store_id&" and item_id="&item_id %>
				<% rs_store3.open sql_sel_attribs, conn_store, 1, 1 %>
				<% do while not rs_store3.eof %>
					<tr bgcolor='#FFFFFF'>
						<td class="inputname" width="20%">Values: <b><%= rs_store3("Attribute_class") %></b></td>
						<td class="inputvalue" width="80%">
							<% If op="edit" then %>
								<% selVals = getAttributesValuesID(item_id, config_id, Option_Name, rs_store3("Attribute_class")) %>
							<% Else %>
								<% selVals = "" %>
							<% End If %>
							<select name="<%= rs_store3("Attribute_class") %>" multiple>
							<% sql_sel_a = "select Attribute_id, Attribute_value from store_items_attributes where store_id="&store_id&" and item_id="&Item_id&" and Attribute_class='"&rs_store3("Attribute_class")&"'" %>
							<%= sql_sel_a %>
							<% rs_store2.open sql_sel_a, conn_store, 1, 1 %>
							<% do while not rs_store2.eof %>
								<option 
								<% if instr(selVals," "&rs_store2("Attribute_id")&", ")>0 then %>
									selected
								<% End If %>
								value="<%= rs_store2("Attribute_id") %>"><%= rs_store2("Attribute_value") %></option>
								<% rs_store2.movenext %>
							<% loop %>
							<% rs_store2.close %>
							</select>
							<INPUT type="hidden" name="<%= rs_store3("Attribute_class") %>_C" value="Re|String||||">
						<% small_help "Values" %></td>
					</tr>
					<% rs_store3.movenext %>
				<% loop %>
				<% rs_store3.close %>


<%

function getAttributesValuesID(theItem, theConfig, theOption, theClass)
	sql_sel_c = "select Attributes_Values from Store_Items_Configs_Dets where item_id="&theItem&" and config_id="&theConfig&" and store_id="&store_id&" and Attributes_Class='"&theClass&"' and Name='"&theOption&"'"
	Set rs_store_loc = Server.CreateObject("ADODB.Recordset")
	rs_store_loc.open sql_sel_c, conn_store, 1, 1
	if rs_store_loc.eof then
		sql_sel_a = "select Attribute_ID from store_items_attributes where store_id="&store_id&" and item_id="&theItem&" and Attribute_class='"&theClass&"'"
	else
		sql_sel_a = "select Attribute_ID from store_items_attributes where store_id="&store_id&" and item_id="&theItem&" and Attribute_class='"&theClass&"' and Attribute_id in ("&rs_store_loc("Attributes_Values")&")"
	end if
	rs_store_loc.close
	
	rs_store_loc.open sql_sel_a, conn_store, 1, 1
	do while not rs_store_loc.eof
		rezA = rezA&", "&rs_store_loc("Attribute_ID")
		rs_store_loc.movenext
	loop
	rs_store_loc.close
	rezA = rezA&", "
	getAttributesValuesID = rezA
	Set rs_store_loc = nothing
end function

Set rs_store2 = nothing
Set rs_store3 = nothing

createFoot thisRedirect, 1
end if
%>


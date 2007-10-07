<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

Item_Id = Request.QueryString("Item_Id")
if not isNumeric(Item_Id) then
	Response.Redirect "admin_error.asp?message_id=1"
end if
sql_select="SELECT Item_Name FROM Store_Items WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" ORDER BY Item_Name;"
rs_store.open sql_select,conn_store,1,1 
	if rs_store.bof or rs_store.eof then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	Item_Name = Rs_store("Item_Name")
rs_store.close
sql_configs = "SELECT config_name, config_id, config_desc from store_items_configs where store_id="&store_id&" and item_id="&item_id

sFormAction = ""
sFormName = ""
sTitle = "Item Configurations"
sFullTitle="Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > Configurations"
thisRedirect = "item_configs.asp"
sMenu = "inventory"
createHead thisRedirect

Set rs_store2 = Server.CreateObject("ADODB.Recordset")
Set rs_store3 = Server.CreateObject("ADODB.Recordset")
%>
	<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="23">
			<input type="button" class="Buttons" value="Back to Item" name="Back_To_Item" OnClick=JavaScript:self.location="Item_Edit.asp?Item_Id=<%= Item_Id %>">
			<input type="button" class="Buttons" value="Add Configuration" name="Add" OnClick=JavaScript:self.location="Item_config_add.asp?Item_Id=<%= Item_Id %>">
			<input type="button" class="Buttons" value="Item Attributes" name="Add" OnClick=JavaScript:self.location="Item_attributes.asp?Item_Id=<%= Item_Id %>"></td>
    </tr>
    
	<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3">

		<table border="0" width="100%" cellpadding="0">
				

				<tr bgcolor='#FFFFFF'><td colspan=2>Item configurations are groupings of attributes.  For instance if you sell computers you could package the
				attributes into different models.</td></tr>
				<% rs_store.open sql_configs,conn_store,1,1 %>
				<% if Not Rs_store.Bof then %>
					<% rs_store.MoveFirst %>
					<% Do While not rs_store.eof	%>
						<tr bgcolor="#DDDDDD"> 
						<td width="50%"><br><b><%= Rs_store("Config_Name") %></b><br>&nbsp;</td>
							<td width="50%"><%= Rs_store("Config_Desc") %></td>
							<td width="1%" nowrap><a class="link" href="item_config_det_add.asp?Item_ID=<%= Item_ID %>&Config_ID=<%= rs_store("Config_ID") %>">Add Options</a></td>
						<td width="1%" nowrap><a class="link" href="Item_config_add.asp?Item_Id=<%= Item_Id %>&Config_ID=<%= rs_store("Config_ID") %>&Op=edit">Edit</a></td>
							<td width="1%" nowrap><a class="link" href="Inventory_Action.asp?Delete_Config=True&Item_Id=<%= Item_Id %>&Config_ID=<%= rs_store("Config_ID") %>">Delete</a></td>
						</tr>
						<% sql_sel_options = "select distinct name from store_items_configs_dets where store_id="&store_id&" and item_id="&item_id&" and config_id="&rs_store("config_id")&" " %>
						<% rs_store2.open sql_sel_options, conn_store, 1, 1 %>
						<% if not rs_store2.eof then %>
							<tr bgcolor='#FFFFFF'>
								<td colspan="5">
									<table border=0>
										<% str_class=1
											Do while not rs_store2.eof %>
											   <% if str_class=1 then
										  str_class=0
									       else
										   str_class=1
									       end if %>
												<tr class="<%= str_class %>">
												<td width="1%" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;</td>
												<td width="30%">Option <b><%= rs_store2("Name") %></b></td>
												<td width="67%">
													<table border="0">
														<% sql_sel_attribs = "select distinct Attribute_class from store_items_attributes where store_id="&store_id&" and item_id="&item_id %>
														<% rs_store3.open sql_sel_attribs, conn_store, 1, 1 %>
														<% do while not rs_store3.eof %>
															<tr bgcolor='#FFFFFF'>
																<td width="1%" nowrap>Attribute: </td>
																<td width="1%" nowrap><b><%= rs_store3("Attribute_class") %></b></td>
																<td width="1%" nowrap>Values: </td>
																<td width="97%"><b><%= checkAttributesValues(item_id, rs_store("Config_ID"), rs_store2("Name"), rs_store3("Attribute_class")) %></b></td>
															</tr>
															<% rs_store3.movenext %>
														<% loop %>
														<% rs_store3.close %>
													</table>
												</td>
												<td width="1%" nowrap><a class="link" href="item_config_det_add.asp?Item_ID=<%= Item_ID %>&Config_ID=<%= rs_store("Config_ID") %>&Class=<%= getClassID(item_id, rs_store("Config_ID"), rs_store2("Name")) %>&op=edit">Edit</a>
												<td width="1%" nowrap><a class="link" href="inventory_action.asp?Item_ID=<%= Item_ID %>&Config_ID=<%= rs_store("Config_ID") %>&Class=<%= getClassID(item_id, rs_store("Config_ID"), rs_store2("Name")) %>&Delete_Option=True">Delete</a>
											</tr>


											<% rs_store2.movenext %>
										<% loop %>
									</table>
								</td>
							</tr>
						<% End If %>
						<% rs_store2.close %>
						<% rs_store.movenext %>
					<% loop %>
					<% rs_store.close  %>
				<% End if %>
			</table>
		</td>
	</tr>


<%
Set rs_store2 = nothing
Set rs_store3 = nothing

function checkAttributesValues(theItem, theConfig, theOption, theClass)
	sql_sel_c = "select Attributes_Values from Store_Items_Configs_Dets where item_id="&theItem&" and config_id="&theConfig&" and store_id="&store_id&" and Attributes_Class='"&theClass&"' and Name='"&theOption&"'"
	Set rs_store_loc = Server.CreateObject("ADODB.Recordset")
	rs_store_loc.open sql_sel_c, conn_store, 1, 1
	if rs_store_loc.eof then
		sql_sel_a = "select Attribute_value from store_items_attributes where store_id="&store_id&" and item_id="&theItem&" and Attribute_class='"&theClass&"'"
	else
		sql_sel_a = "select Attribute_value from store_items_attributes where store_id="&store_id&" and item_id="&theItem&" and Attribute_class='"&theClass&"' and Attribute_id in ("&rs_store_loc("Attributes_Values")&")"
	end if
	rs_store_loc.close
	
	rs_store_loc.open sql_sel_a, conn_store, 1, 1
	rezA = ""
	do while not rs_store_loc.eof
		if rezA<>"" then
			rezA = rezA&" / "&rs_store_loc("Attribute_value")
		else
			rezA = rs_store_loc("Attribute_value")
		end if
		rs_store_loc.movenext
	loop
	rs_store_loc.close
	checkAttributesValues = rezA
	Set rs_store_loc = nothing
end function

function getClassID(theItem, theConfig, theClass)
	Set rs_store_loc = Server.CreateObject("ADODB.Recordset")
	sql_sel = "select det_id from Store_Items_Configs_Dets where store_id="&store_id&" and item_id="&theItem&" and config_id="&theConfig&" and Name='"&theClass&"'"
	rs_store_loc.open sql_sel, conn_store, 1, 1
	getClassID = rs_store_loc("det_id")
	rs_store_loc.close
	Set rs_store_loc = nothing
end function

createFoot thisRedirect, 0
%>


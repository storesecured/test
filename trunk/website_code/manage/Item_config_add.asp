<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->

<%
on error goto 0
config_id = Request.QueryString("config_id")
Item_Id = Request.QueryString("Item_Id")
if not isNumeric(Item_Id) or not isNumeric(config_id) then
	Response.Redirect "admin_error.asp?message_id=1"
end if
op=Request.QueryString("Op")
if op="edit" then
	' IF EDIT, THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlConfig="select * from store_items_configs where config_id=" & config_id & " and store_id=" & store_id
	rs_Store.open sqlConfig,conn_store,1,1
	if rs_store.bof or rs_store.eof then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	Config_name = rs_store("Config_name")
	Config_desc = rs_store("Config_desc")
	rs_store.close
end if

sql_select="SELECT Item_Name FROM Store_Items WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" ORDER BY Item_Name;"
rs_store.open sql_select,conn_store,1,1
if rs_store.bof or rs_store.eof then
	Response.Redirect "admin_error.asp?message_id=1"
end if
Item_Name = Rs_store("Item_Name")
rs_store.close

sCancel="item_configs.asp?Item_Id="&Item_Id&"&"&sAddString
sFullTitle="Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > <a href="&sCancel&" class=white>Configurations</a> > "
sFormAction = "Inventory_Action.asp"
sFormName = "Add_Config"
if op="edit" then
	sTitle = "Edit Configuration - "&Item_Name
	sFullTitle=sFullTitle&"Edit - "&Item_Name
else
	sTitle = "Add Configuration - " &item_name
	sFullTitle=sFullTitle&"Add - "&Item_Name
end if
sCommonName="Configuration"
addPicker=1
sSubmitName = "Update"
thisRedirect = "item_config_add.asp"
sMenu = "inventory"
createHead thisRedirect
if Service_Type < 5  then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		SILVER Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
   <% createFoot thisRedirect, 0%>
<% else 

%>


<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Config_Id" value="<%=config_id%>">
<input type="Hidden" name="Item_Id" value="<%= Item_Id %>">


	<tr bgcolor='#FFFFFF'>
					<td class="inputname" width="20%">Name</td>
					<td class="inputvalue" width="80%">
						<input type="text" name="Config_Name" value="<%= Config_Name %>" size="60" maxlength="50">
						<INPUT type="hidden" name="Config_Name_C" value="Re|String|0|50|||Name">
			      <% small_help "Name" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td class="inputname" width="100%" colspan=2>Description<BR>
						<%= create_editor ("Config_Desc",Config_Desc,"") %>
						<INPUT type="hidden" name="Config_Desc_C" value="Op|String|||||Description">
			      <% small_help "Description" %></td>
				</tr>
				

				

<% createFoot thisRedirect, 1%>
<% end if %>

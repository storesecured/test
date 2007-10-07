<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/department_list.asp"-->

<%

F_NAME = ""
F_SKU = ""
F_DATE = ""
F_KEYWORD = ""
Sub_Department_Id = -1
F_LIVE = 3
LIVE_LIVE = ""
LIVE_NOT = ""
LIVE_ALL = "checked"

select case request.form("SortBy")
	case "2"
		sort2="checked"
		SQL_SORT = " ORDER BY Item_SKU "
	case "3"
		sort3="checked"
		SQL_SORT = " ORDER BY Item_ID "
	case else
		sort1="checked"
		SQL_SORT = " ORDER BY Item_Name "
end select

if Request.Form("Display_Items")="" then
else
	F_NAME = request.form("F_NAME")
	F_SKU = request.form("F_SKU")
	F_DATE = request.form("F_DATE")
	F_KEYWORD = request.form("F_KEYWORD")

	if request.form("F_LIVE")<>"" then
		select case request.form("F_LIVE")
			case "1"
				F_LIVE = 1
				LIVE_LIVE = "checked"
				LIVE_NOT = ""
				LIVE_ALL = ""
			case "2"
				F_LIVE = 2
				LIVE_LIVE = ""
				LIVE_NOT = "checked"
				LIVE_ALL = ""
			case "3"
				F_LIVE = 3
				LIVE_LIVE = ""
				LIVE_NOT = ""
				LIVE_ALL = "checked"
		end select
	else
		F_LIVE = 3
		LIVE_LIVE = ""
		LIVE_NOT = ""
		LIVE_ALL = "checked"
	end if
	if Request.Form("Sub_Department_Id") <> "" then
		Sub_Department_Id = Request.Form("Sub_Department_Id") 
	else
		Sub_Department_Id = Request.QueryString("Sub_Department_Id")	
	end if
	
	if not isNumeric(Sub_Department_Id) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	
	If instr(" "&Sub_Department_Id&","," -1,")>0 then
		Sub_Department_Id = "-1"
	end if
	if F_NAME<>"" then
		COND_NAME = " AND Item_Name LIKE '"&F_NAME&"%' "
	end if
	if F_SKU<>"" then
		COND_SKU = " AND Item_Sku LIKE '"&F_SKU&"%' "
	end if
	if F_DATE<>"" then
		if isdate(F_DATE) then
			COND_DATE = " AND Sys_Modified >= #"&F_DATE&"# "
			if Store_Database="Ms_Sql" then
				COND_DATE = Replace(COND_DATE, "#", "'")
			end if
		end if
	end if
	if F_KEYWORD<>"" then
		COND_KEYWORD =" AND (Description_S LIKE '%"&F_KEYWORD&"%' OR Description_L LIKE '%"&F_KEYWORD&"%' OR Item_Remarks LIKE '%"&F_KEYWORD&"%' )"
	end if
	select case F_LIVE
		case 1
			COND_LIVE = " AND Show=-1 "
		case 2
			COND_LIVE = " AND Show<>-1 "
	end select
	if Sub_Department_Id<>"" then
		COND_DEPT = " AND Department_Id = "&Sub_Department_Id
	end if
		
	sql_select_items="SELECT Item_Id, Item_Name, Show, Item_Sku FROM sv_items_dept_combine WHERE Store_id="&Store_id&COND_NAME & COND_SKU & COND_DATE & COND_KEYWORD & COND_LIVE & COND_DEPT&SQL_SORT
    
end if

if Request.QueryString("op")<>"" then
	op=Request.QueryString("op")
else
	if Request.Form("op")<>"" then
		op=Request.form("op")
	end if
end if
if op="edit" then
	'IF EDIT THEN LOAD CURRENT ACCESSORIES FROM THE DATABASE
	if Request.QueryString("Accessory_Line_Id") then
		Accessory_Line_Id = Request.QueryString("Accessory_Line_Id")
	else
		if Request.Form("Accessory_Line_Id") then
			Accessory_Line_Id = Request.form("Accessory_Line_Id")
		end if
	end if
	
	if not isNumeric(Accessory_Line_Id) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	
	sqlAccessory="select * from Store_items_accessories where Accessory_Line_Id=" & Accessory_Line_Id & " and store_id=" & store_id
	rs_Store.open sqlAccessory,conn_store,1,1
	AccessoryItemId=rs_store("Accessory_Item_Id")
	rs_store.close
end if

if Request.QueryString("Item_Id")<>"" then
	Item_Id = Request.QueryString("Item_Id")
else
	if Request.form("Item_Id") then
		Item_Id = Request.form("Item_Id")
	end if
end if

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

sCancel="Item_accessories.asp?Item_Id="&Item_Id&"&"&sAddString
sFullTitle ="Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > <a href="&sCancel&" class=white>Accessories</a> > "

sFormAction = "Item_Accessories_add.asp?Item_Id="&Item_Id
sFormName = "Add_Accessories"
if op="edit" then
	sTitle ="Edit Accessory - "&Item_Name
	sFullTitle=sFullTitle&"Edit - "&item_name&" - "&AttributeClass&" "&AttributeValue

else
	sTitle ="Add Accessory - "&Item_Name
	sFullTitle=sFullTitle&"Add - "&item_name&" - "&AttributeClass&" "&AttributeValue

end if
sCommonName="Accessory"
sSubmitName = "Add_Accessories"
thisRedirect = "item_accessories_add.asp"
sMenu="inventory"
createHead thisRedirect
if Service_Type < 5 then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		SILVER Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0%>
<% else %>



<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Accessory_Line_Id" value="<%= Accessory_Line_Id %>">
<input type="Hidden" name="Item_Id" value="<%= Item_Id %>">



	 


							<TR bgcolor='#FFFFFF'>
								<td colspan="3"><b>Select items you want to view</b></td>
							</tr>
					
							<TR bgcolor='#FFFFFF'>
								<td class="inputname">From Departments</td>
								<td class="inputvalue">
									<%= create_dept_list ("Sub_Department_ID",Sub_Department_ID,1,"") %>
                                                                        <% small_help "From Departments" %></td>
							</tr> 
					
							<TR bgcolor='#FFFFFF'>
								<td class="inputname">Product Name like</td>
								<td class="inputvalue"><input type="text" name="F_NAME" size="30" value="<%= F_NAME %>">
								<INPUT type="hidden" name="F_NAME_C" value="Op|String|||||Product Name">
								<% small_help "Product Name like" %></td>
							</tr>

							<TR bgcolor='#FFFFFF'>
								<td class="inputname">Product SKU like</td>
								<td class="inputvalue"><input type="text" name="F_SKU" size="30" value="<%= F_SKU %>">
								<INPUT type="hidden" name="F_SKU_C" value="Op|String|||||Product SKU">
								<% small_help "Product SKU like" %></td>
							</tr>
				
							<TR bgcolor='#FFFFFF'>
								<td class="inputname">Live status like</td>
								<td class="inputvalue"><input class="image" type="radio" name="F_LIVE" <%= LIVE_LIVE %> value="1">Live&nbsp;<input class="image" type="radio" name="F_LIVE" <%= LIVE_NOT %> value="2">Not Live&nbsp;<input class="image" type="radio" name="F_LIVE" <%= LIVE_ALL %> value="3">All
								<% small_help "Live status like" %></td>
							</tr>

							<TR bgcolor='#FFFFFF'>
								<td class="inputname">Last modified after</td>
								<td class="inputvalue"><input type="text" name="F_DATE" size="30" value="<%= F_DATE %>">
								<INPUT type="hidden" name="F_DATE_C" value="Op|String|||||Modified After">
								<% small_help "Last modified after" %></td>
							</tr>
			
							<TR bgcolor='#FFFFFF'>
								<td class="inputname">Containing keyword</td>
								<td class="inputvalue"><input type="text" name="F_KEYWORD" size="30" value="<%= F_KEYWORD %>">
								<INPUT type="hidden" name="F_KEYWORD_C" value="Op|String|||||Keyword">
								<% small_help "Containing keyword" %></td>
							</tr>
			
							<TR bgcolor='#FFFFFF'>
						<td colspan="3" width="109%" height="17" align="center">
									<input type="submit" class="Buttons" value="Display Items" name="Display_Items"><br></td>
							</tr>

			</form>

				<form method="POST" action="Inventory_Action.asp">
				<input type="hidden" name="op" value="<%=op%>">
				<input type="hidden" name="Accessory_Line_Id" value="<%= Accessory_Line_Id %>">
				<input type="Hidden" name="Form_Name" value="Add_Accessories">
				<input type="Hidden" name="Item_Id" value="<%= Item_Id %>">
 
				

				<TR bgcolor='#FFFFFF'>
				<td class="inputname">Select Items</td><td class="inputvalue">
					<select size="7" 
							<% if op="" then %>
								multiple
							<% end if %> 
							name="Items_Id">
						<% 
						set myfields=server.createobject("scripting.dictionary")
						Call DataGetrows(conn_store,sql_select_items,mydata,myfields,noRecords)
						if noRecords = 0 then
						FOR rowcounter= 0 TO myfields("rowcount")

								if mydata(myfields("show"),rowcounter) = -1 then
									show = "on" 
								else 
									show = "off"
								end if 
								if op="edit" then
									if AccessoryItemId=mydata(myfields("item_id"),rowcounter) then 
										selected = "selected"
									else
										selected = ""
									end if 
								else 
									selected=""
								end if 
								response.write "<option "&selected&" value='"&mydata(myfields("item_id"),rowcounter)&"'>["&Show&"] "&mydata(myfields("item_name"),rowcounter)&"</option>"
						Next
						end if %>
						</select><% small_help "Select Items" %></td>
				</tr>
			
		
<% 
createFoot thisRedirect, 1
end if 
set myfields=Nothing
%>

<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
'on error resume next
if request.form("ATTR_LINE_ID")<>"" then
	ATTR_LINE_ID = request.form("ATTR_LINE_ID")

	if clng(ATTR_LINE_ID)>=0 then
		A_NAME = checkstringforq(request.form("A_NAME_"&ATTR_LINE_ID))
		A_VALUE = checkstringforq(request.form("A_VALUE_"&ATTR_LINE_ID))
		A_ORDER = request.form("A_ORDER_"&ATTR_LINE_ID)
		if request.form("A_DEFAULT_"&ATTR_LINE_ID)="" then
			A_DEFAULT = 0
		else
			A_DEFAULT = -1
		end if
		A_USE_ITEM = request.form("A_USE_ITEM_"&ATTR_LINE_ID)
		A_IITEM_ID = request.form("A_IITEM_ID_"&ATTR_LINE_ID)
		A_IQUANTITY = request.form("A_IQUANTITY_"&ATTR_LINE_ID)
		A_USE_ITEM = request.form("A_USE_ITEM_"&ATTR_LINE_ID)
		A_PRICE_DIF = request.form("A_PRICE_DIF_"&ATTR_LINE_ID)
		A_WEI_DIF = request.form("A_WEI_DIF_"&ATTR_LINE_ID)
		A_SKU = request.form("A_SKU_"&ATTR_LINE_ID)
		A_HIDD = request.form("A_HIDD_"&ATTR_LINE_ID)

		if A_NAME<>"" and A_VALUE<>"" and isNumeric(A_ORDER) and isNumeric(A_IITEM_ID) and isNumeric(A_PRICE_DIF) and isNumeric(A_WEI_DIF) then
			sql_del = "delete from store_items_attributes where store_id="&store_id&" and Attribute_class='"&A_NAME&"' and Attribute_value='"&A_VALUE&"' and item_id in ("&request.form("DELETE_IDS")&")"
			conn_store.execute sql_del
		   
			for each itemID in split(request.form("DELETE_IDS"),",")
				sql_ins = "insert into store_items_attributes (Store_ID, Item_Id, Attribute_class, Attribute_value, Attribute_Price_difference, Attribute_Weight_difference, [Default], Attribute_sku, Attribute_hidden, Use_Item, IItem_ID, IItem_Quantity, Display_Order) values ("&store_id&", "&itemID&", '"&A_NAME&"', '"&A_VALUE&"', "&A_PRICE_DIF&", "&A_WEI_DIF&", "&A_DEFAULT&", '"&A_SKU&"', '"&A_HIDD&"', "&A_USE_ITEM&", "&A_IITEM_ID&", "&A_IQUANTITY&", "&A_ORDER&")"
				conn_store.execute sql_ins
			next
		else
			Error_Log = "Invalid values"
			%><!--#include file="Include/Error_Template.asp"--><%
			response.end
		end if
	else
		A_NAME = checkstringforQ(request.form("A_NAME_"&(-clng(ATTR_LINE_ID))))
		A_VALUE = checkstringforQ(request.form("A_VALUE_"&(-clng(ATTR_LINE_ID))))
		sql_del = "delete from store_items_attributes where store_id="&store_id&" and Attribute_class='"&A_NAME&"' and Attribute_value='"&A_VALUE&"' and item_id in ("&request.form("DELETE_IDS")&")"
		conn_store.execute sql_del
	end if
end if

sFormAction = "item_mass_edit_attr.asp"
sTitle = "Mass Edit Attributes"
sSubmitName = "MASS_ADD"
sName = "item_mass_edit_attr"
thisRedirect = "item_mass_edit_attr.asp"
sMenu = "inventory"
createHead thisRedirect

if Service_Type < 5 then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		SILVER Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>

<% else %>

<script language="JavaScript">
	function goSubmit(theVal){
		document.forms["<%= sName %>"].ATTR_LINE_ID.value = theVal;
		document.forms["<%= sName %>"].submit();}
</script>

<input type="hidden" name="Display_Items" value="Display_Items">
<input type="hidden" name="F_NAME" value="<%= request.form("F_NAME") %>">
<input type="hidden" name="F_SKU" value="<%= request.form("F_SKU") %>">
<input type="hidden" name="F_DATE" value="<%= request.form("F_DATE") %>">
<input type="hidden" name="F_KEYWORD" value="<%= request.form("F_KEYWORD") %>">
<input type="hidden" name="Sub_Department_Id" value="<%= Request.Form("Sub_Department_Id") %>">
<input type="hidden" name="F_LIVE" value="<%= request.form("F_LIVE") %>">
<input type="hidden" name="DELETE_IDS" value="<%= request.form("DELETE_IDS") %>">
<input type="hidden" name="ATTR_LINE_ID" value="">


	<TR bgcolor='#FFFFFF'>
		<td colspan="3"><input type="button" class="Buttons" value="Back to Items" name="Back_To_Item_List" OnClick='JavaScript:self.location="edit_Items.asp?<%= sAddString %>"'></td>
	</tr>
	
	<TR bgcolor='#FFFFFF'>
		<td colspan="3"><p><font face="Arial" size="2" color="red"><b>Use this page to setup the same attributes for multiple items.</b><BR><BR>
                </td>
	</tr>

	<TR bgcolor='#FFFFFF'>
		<td width="100%" colspan="3">
			<table border="0" width="100%" cellpadding="0">
				<tr bgcolor="#AAAAAA">
					<td><b>Name</b></td>
					<td><b>Value</b></td>
					<td><b>Order</b></td>
					<td><b>Default</b></td>
					
					<td><b>Use item</b></td>
					<td><b>Item ID</b></td>
					<td><b>Quantity</b></td>
					
					<td><b>Don't use item</b></td>
					<td><b>Price Diff</b></td>
					<td><b>Weight Diff</b></td>
					<td><b>SKU</b></td>
					<td><b>Hidden Value</b></td>
					
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr bgcolor="#CCCCCC">
					<td colspan="14"><b><i>Existing Attributes</i></b></td>
				</tr>
				
				<% 
				sql_select_att = "select distinct Attribute_class, Attribute_value, attribute_price_difference,attribute_weight_difference,[default],attribute_sku,attribute_hidden,use_item,iitem_id,iitem_quantity,display_order from store_items_attributes where store_id="&store_id&" and item_id in ("&request.form("DELETE_IDS")&") order by Attribute_class, Attribute_value"
				set myfields=server.createobject("scripting.dictionary")
				Call DataGetrows(conn_store,sql_select_att,mydata,myfields,noRecords)

				cline = 1
				if noRecords = 0 then
				FOR rowcounter= 0 TO myfields("rowcount") 
				    sUse_Item=mydata(myfields("use_item"),rowcounter)
                                    sDefault=mydata(myfields("default"),rowcounter)
                                    
                                    if sUse_Item=0 then
                                       sDontUseChecked="checked"
                                       sUseChecked=""
                                    else
                                       sDontUseChecked=""
                                       sUseChecked="checked"
                                    end if
                                    
                                    if sDefault=0 then
                                       sDefaultChecked=""
                                    else
                                       sDefaultChecked="checked"
                                    end if
                                    %>

				
					<input name="A_NAME_<%= cline %>" type="hidden" value="<%= mydata(myfields("attribute_class"),rowcounter) %>">
					<input name="A_VALUE_<%= cline %>" type="hidden" value="<%= mydata(myfields("attribute_value"),rowcounter) %>">
					
					<TR bgcolor='#FFFFFF'>
						<td class="inputname"><%= mydata(myfields("attribute_class"),rowcounter) %></td>
						<td class="inputname"><%= mydata(myfields("attribute_value"),rowcounter) %></td>
						<td class="inputname"><input type=text size=3 name=A_ORDER_<%= cline %> value="<%= mydata(myfields("display_order"),rowcounter) %>"></td>

                  <td class="inputvalue"><input class="image" name="A_DEFAULT_<%= cline %>" type="checkbox" value="-1" <%=sDefaultChecked%>></td>
						
						<td class="inputvalue"><input class="image" name="A_USE_ITEM_<%= cline %>" type="radio" value="-1" <%= sUseChecked %>></td>
						<td class="inputvalue"><input name="A_IITEM_ID_<%= cline %>" type="text" size="3" value="<%= mydata(myfields("iitem_id"),rowcounter) %>"></td>
						<td class="inputvalue"><input name="A_IQUANTITY_<%= cline %>" type="text" size="3" value="<%= mydata(myfields("iitem_quantity"),rowcounter) %>"></td>
					
						<td class="inputvalue"><input class="image" name="A_USE_ITEM_<%= cline %>" type="radio" value="0"  <%= sDontUseChecked %>></td>
						<td class="inputvalue"><input name="A_PRICE_DIF_<%= cline %>" type="text" size="3" value="<%= mydata(myfields("attribute_price_difference"),rowcounter) %>"></td>
						<td class="inputvalue"><input name="A_WEI_DIF_<%= cline %>" type="text" size="3" value="<%= mydata(myfields("attribute_weight_difference"),rowcounter) %>"></td>
						<td class="inputvalue"><input name="A_SKU_<%= cline %>" type="text" size="3" value="<%= mydata(myfields("attribute_sku"),rowcounter) %>"></td>
						<td class="inputvalue"><input name="A_HIDD_<%= cline %>" type="text" size="3" value="<%= mydata(myfields("attribute_hidden"),rowcounter) %>"></td>
						
						<td><input type="button" name="MASS_ADD" value="Add Update" onClick="JavaScript:goSubmit('<%= cline %>');"></td>
						<td><input type="button" name="MASS_ADD" value="Delete" onClick="JavaScript:goSubmit('-<%= cline %>');"></td>
					</tr>
					<% cline = cline + 1
				Next
				end if 
				set myfields= Nothing
				%>
				
				<tr bgcolor="#CCCCCC">
					<td colspan="14"><b><i>New Attribute</i></b></td>
				</tr>
				
				<TR bgcolor='#FFFFFF'>
					<td class="inputname"><input name="A_NAME_0" type="text" size="3" value=""></td>
					<td class="inputname"><input name="A_VALUE_0" type="text" size="3" value=""></td>
					<td class="inputname"><input type=text size=3 name=A_ORDER_0 value=""></td>
               <td class="inputvalue"><input class="image" name="A_DEFAULT_0" type="checkbox" value="-1"></td>
					
					<td class="inputvalue"><input class="image" name="A_USE_ITEM_0" type="radio" value="-1"></td>
					<td class="inputvalue"><input name="A_IITEM_ID_0" type="text" size="3" value=""></td>
					<td class="inputvalue"><input name="A_IQUANTITY_0" type="text" size="3" value=""></td>
					
					<td class="inputvalue"><input class="image" name="A_USE_ITEM_0" type="radio" value="0" checked></td>
					<td class="inputvalue"><input name="A_PRICE_DIF_0" type="text" size="3" value=""></td>
					<td class="inputvalue"><input name="A_WEI_DIF_0" type="text" size="3" value=""></td>
					<td class="inputvalue"><input name="A_SKU_0" type="text" size="3" value=""></td>
					<td class="inputvalue"><input name="A_HIDD_0" type="text" size="3" value=""></td>
					
					<td><input type="button" name="MASS_ADD" value="Add" onClick="JavaScript:goSubmit('0');"></td>
					<td>&nbsp;</td>
				</tr>
				
		</table>
		</td>
	</tr>
	<% end if %>
<% createFoot thisRedirect, 0%>

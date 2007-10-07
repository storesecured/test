<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->
<%
sql_select = "Select Item_F_Layout_Id, Item_F_Layout, Item_S_Layout_Id, Item_S_Layout, Item_L_Layout_Id, Item_L_Layout from Store_Settings where Store_Id="&Store_Id
rs_Store.open sql_select,conn_store,1,1
if not rs_Store.bof and not rs_Store.eof then
	Item_S_Layout_Id = rs_Store("Item_S_Layout_Id")
	Item_S_Layout = rs_Store("Item_S_Layout")
	Item_L_Layout_Id = rs_Store("Item_L_Layout_Id")
	Item_L_Layout = rs_Store("Item_L_Layout")
	Item_F_Layout_Id = rs_Store("Item_F_Layout_Id")
	Item_F_Layout = rs_Store("Item_F_Layout")
end if
rs_Store.close

if Item_S_Layout_Id = 6 then
	sitemchecked6 = "checked"
elseif Item_S_Layout_Id = 7 then
	sitemchecked7 = "checked"
elseif Item_S_Layout_Id = 8 then
	sitemchecked8 = "checked"
elseif Item_S_Layout_Id = 9 then
	sitemchecked9 = "checked"
elseif Item_S_Layout_Id = 10 then
	sitemchecked10 = "checked"
elseif Item_S_Layout_Id = 11 then
	sitemchecked11 = "checked"
elseif Item_S_Layout_Id = 12 then
	sitemchecked12 = "checked"
end if

if Item_L_Layout_Id = 6 then
	litemchecked6 = "checked"
elseif Item_L_Layout_Id = 7 then
	litemchecked7 = "checked"
elseif Item_L_Layout_Id = 8 then
	litemchecked8 = "checked"
elseif Item_L_Layout_Id = 9 then
	litemchecked9 = "checked"
elseif Item_L_Layout_Id = 10 then
	litemchecked10 = "checked"
elseif Item_L_Layout_Id = 11 then
	litemchecked11 = "checked"
elseif Item_L_Layout_Id = 12 then
	litemchecked12 = "checked"
end if

if Item_F_Layout_Id = 6 then
	fitemchecked6 = "checked"
elseif Item_F_Layout_Id = 7 then
	fitemchecked7 = "checked"
elseif Item_F_Layout_Id = 8 then
	fitemchecked8 = "checked"
elseif Item_F_Layout_Id = 9 then
	fitemchecked9 = "checked"
elseif Item_F_Layout_Id = 10 then
	fitemchecked10 = "checked"
elseif Item_F_Layout_Id = 11 then
	fitemchecked11 = "checked"
elseif Item_F_Layout_Id = 12 then
	fitemchecked12 = "checked"
end if

dim sItemTemplate(12)
dim sItemTemplatef(12)
sql_select = "select * from Sys_Inventory_Template where Template_Type=2"
rs_Store.open sql_select,conn_store,1,1

do while not rs_Store.eof
	sTemplateStart = rs_Store("Template")
	sTemplate_Id=rs_Store("Template_Id")
	sTemplate=sTemplateStart
     sTemplate = fn_make_template (sTemplate,0)
     sItemTemplate(sTemplate_Id) = sTemplate
	sTemplatef=sTemplateStart
	sTemplatef = fn_make_template (sTemplatef,1)
	sItemTemplatef(sTemplate_Id) = sTemplatef
	rs_Store.movenext
loop
rs_Store.close

function fn_make_template (sTemplateModify,bSearch)
	sTemplateModify = Replace(sTemplateModify,"OBJ_IMAGE_OBJ","<IMG Src=images/spacer.gif height=40 width=40 border=1>")
	sTemplateModify = Replace(sTemplateModify,"OBJ_NAME_OBJ","Item Name")
	sTemplateModify = Replace(sTemplateModify,"OBJ_PRICE","9.99")
	sTemplateModify = Replace(sTemplateModify,"OBJ_FINAL_PRICE_OBJ","9.99")
	sTemplateModify = Replace(sTemplateModify,"OBJ_SPECIAL_PRICE_OBJ","Our Price 8.99")
	sTemplateModify = Replace(sTemplateModify,"OBJ_MATRIX_OBJ","Buy 2 and up for 7.99")
	sTemplateModify = Replace(sTemplateModify,"OBJ_DESCRIPTION_OBJ","Item description will go here, yours might be longer or shorter.")
	sTemplateModify = Replace(sTemplateModify,"OBJ_STOCK_OBJ","In Stock (23)")
	if bSearch=0 then
		sTemplateModify = Replace(sTemplateModify,"OBJ_ATTRIBUTE","Select Size <select name=size><option value=1>Red</option><option value=2>Blue</option></select>")
          sTemplateModify = Replace(sTemplateModify,"OBJ_UD_OBJ","User Defined Fields")
          sTemplateModify = Replace(sTemplateModify,"OBJ_ACCESSORY_OBJ","<HR>Item Accessories<BR>List of accessories if any here.")
          sTemplateModify = Replace(sTemplateModify,"OBJ_QTY_OBJ","<input type='text' value='1' size=3>")
		sTemplateModify = Replace(sTemplateModify,"OBJ_ORDER_OBJ","<input type='submit' value='Add to Cart'>")
	else
		sTemplateModify = Replace(sTemplateModify,"OBJ_ATTRIBUTE","")
          sTemplateModify = Replace(sTemplateModify,"OBJ_UD_OBJ","")
          sTemplateModify = Replace(sTemplateModify,"OBJ_ACCESSORY_OBJ","")
          sTemplateModify = Replace(sTemplateModify,"OBJ_QTY_OBJ","")
		sTemplateModify = Replace(sTemplateModify,"OBJ_ORDER_OBJ","")
	end if

	fn_make_template=sTemplateModify
end function

sNeedTabs=1
addPicker=1
sFormAction = "Store_Settings.asp"
sName = "Store_Item_Layout"
sFormName = "Store_Item_Layout"
sTitle = "Item Layout"
sFullTitle = "Inventory > <a href=edit_items.asp class=white>Items</a> > Layout"
sCommonName="Item Layout"
sCancel="edit_items.asp"
sSubmitName = "Store_Item_Layout"
thisRedirect = "item_layout.asp"
sTopic="Store_Item_Layout"
sMenu = "inventory"
sQuestion_Path = "design/item_layout.htm"
createHead thisRedirect
if Service_Type < 5	then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		SILVER Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0%>

<% else %>

<tr bgcolor='#FFFFFF'><td width="724" align=center valign=top height=35>
<table border=0 cellspacing=0 cellpadding=0 width=724>
			 
	<!-- TAB MENU STARTS HERE -->

		<TR bgcolor='#FFFFFF'>
		<td align="center" valign=top height=45 width='100%'><br>
		<script type="text/javascript" language="JavaScript1.2" src="include/tabs-xp.js"></script>
		<script language="javascript1.2">
		var bselectedItem   = 0;
		var bmenuItems =
		[
		["Small Layout", "content1",,,"Small Layout","Small Layout"],
		["Detail Layout", "content2",,,"Detail Layout","Detail Layout"],
		["Featured Layout", "content3",,,"Featured Layout","Featured Layout"],

		];

		apy_tabsInit();
		</script>
		</td>
		</tr>
		
		<TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='25'>
		
		
		<!-- CONTENT 1 MAIN -->
			<div id="content1" style="visibility: hidden;" class=tpage>
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
                                      <TR bgcolor='#FFFFFF'><TD colspan=3>The small item layout is used for displaying
                                      multiple items at the same time in a summary view. </td></tr>





					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_S_Layout_Id" type="radio" value="6" <%=sItemChecked6 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(6) %>
							<% small_help "6" %></td>

						</tr>

					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_S_Layout_Id" type="radio" value="7" <%=sItemChecked7 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(7) %>
							<% small_help "7" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_S_Layout_Id" type="radio" value="8" <%=sItemChecked8 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(8) %>
							<% small_help "8" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_S_Layout_Id" type="radio" value="9" <%=sItemChecked9 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(9) %>
							<% small_help "9" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_S_Layout_Id" type="radio" value="10" <%=sItemChecked10 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(10) %>
							<% small_help "10" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_S_Layout_Id" type="radio" value="12" <%=sItemChecked12 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(12) %>
							<% small_help "10" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_S_Layout_Id" type="radio" value="11" <%=sItemChecked11 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><B>Custom Layout:</b>
							<BR>OBJ_ID_OBJ - Item Id
                                                        <BR>OBJ_SKU_OBJ - Item Sku
										<BR>OBJ_IMAGES_OBJ - Item Small Image
										<BR>OBJ_IMAGES_URL_OBJ - Item Small Image URL Only
										<BR>OBJ_IMAGEL_OBJ - Item Large Image
										<BR>OBJ_IMAGEL_URL_OBJ - Item Large Image URL Only
										<BR>OBJ_NAME_OBJ - Item Name
										<BR>OBJ_NAME_NOLINK_OBJ - Item Name without link to item detail
										<BR>OBJ_DETAIL_LINK_OBJ - URL to item detail page
										<BR>OBJ_DEPT_LINK_OBJ - URL to department page
										<BR>OBJ_DEPT_NAME_OBJ - Department Name
										<BR>OBJ_ATTRIBUTE - Select boxes for all item attributes (if any)
										<BR>OBJ_PRICE - Regular retail price
										<BR>OBJ_FINAL_PRICE_OBJ - Final price to charge shoppers
										<BR>OBJ_SPECIAL_PRICE_OBJ - Special price (if any)
										<BR>OBJ_MATRIX_OBJ - Quantity discount pricing details (if any)
										<BR>OBJ_UD_OBJ - User defined fields (if any)
										<BR>OBJ_DESCRIPTIONS_OBJ - Item small description
										<BR>OBJ_DESCRIPTIONL_OBJ - Item large description
										<BR>OBJ_QTY_OBJ - Order quantity input box
										<BR>OBJ_ORDER_OBJ - Add to Cart button
										<BR>OBJ_STOCK_OBJ - Number of items in stock (only shown in qty control is enabled)
										<BR>OBJ_WEIGHT_OBJ - Weight
										<BR>OBJ_HANDLING_FEE_OBJ - Handling Fee
										<BR>OBJ_RECURRING_FEE_OBJ - Recurring Fee
										<BR>OBJ_RECURRING_DAYS_OBJ - Recurring Days
										<BR>OBJ_SPECIAL_PRICE_DISCOUNT%_OBJ - Special Price Discount %
										<BR>OBJ_SPECIAL_PRICE_START_DATE_OBJ - Special Price Start Date
                                                                                <BR>OBJ_SPECIAL_PRICE_END_DATE_OBJ - Special PRice End Date
                                                                                <BR>OBJ_EXT_FIELD1_OBJ - Extended Field 1
                                                  				<BR>OBJ_EXT_FIELD2_OBJ - Extended Field 2
										<BR>OBJ_EXT_FIELD3_OBJ - Extended Field 3
										<BR>OBJ_EXT_FIELD4_OBJ - Extended Field 4
										<BR>OBJ_EXT_FIELD5_OBJ - Extended Field 5
										<BR>
	                  <%= create_editor ("Item_S_Layout",Item_S_Layout,"[""Item Id"",""OBJ_ID_OBJ""],[""Item Sku"",""OBJ_SKU_OBJ""],[""Small Image"",""OBJ_IMAGES_OBJ""],[""Large Image"",""OBJ_IMAGEL_OBJ""],[""Item Name"",""OBJ_NAME_OBJ""],[""Item Name no link"",""OBJ_NAME_NOLINK_OBJ""],[""Item Attributes"",""OBJ_ATTRIBUTE""],[""Retail Price"",""OBJ_PRICE""],[""Final Price"",""OBJ_FINAL_PRICE_OBJ""],[""Special Price"",""OBJ_SPECIAL_PRICE_OBJ""],[""Quantity Discount"",""OBJ_MATRIX_OBJ""],[""User Defined Fields"",""OBJ_UD_OBJ""],[""Small Description"",""OBJ_DESCRIPTIONS_OBJ""],[""Large Description"",""OBJ_DESCRIPTIONL_OBJ""],[""Quantity Box"",""OBJ_QTY_OBJ""],[""Add to Cart"",""OBJ_ORDER_OBJ""],[""Qty in Stock"",""OBJ_QTY_OBJ""],[""Weight"",""OBJ_WEIGHT_OBJ""],[""Handling Fee"",""OBJ_HANDLING_FEE_OBJ""],[""Recurring Fee"",""OBJ_RECURRING_FEE_OBJ""],[""Recurring Days"",""OBJ_RECURRING_DAYS_OBJ""],[""Discount %"",""OBJ_SPECIAL_PRICE_DISCOUNT%_OBJ""],[""Special Start Date"",""OBJ_SPECIAL_PRICE_START_DATE_OBJ""],[""Special End Date"",""OBJ_SPECIAL_PRICE_END_DATE_OBJ""],[""Extended Field 1"",""OBJ_EXT_FIELD1_OBJ""],[""Extended Field 2"",""OBJ_EXT_FIELD2_OBJ""],[""Extended Field 3"",""OBJ_EXT_FIELD3_OBJ""],[""Extended Field 4"",""OBJ_EXT_FIELD4_OBJ""],[""Extended Field 5"",""OBJ_EXT_FIELD5_OBJ""]") %>
	                  <input type="hidden" name="Item_S_Layout_C" value="Op|String|0|2000|||Small Layout">
			 <% small_help "Custom Small Layout" %></td>
						</tr>
  </table>
			</div>
			<!-- CONTENT 1 ENDS HERE -->
		<!-- CONTENT 2 MAIN -->
			<div id="content2" style="visibility: hidden;" class=tpage>
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>


                                      <TR bgcolor='#FFFFFF'><TD colspan=3>The detail item layout is used for the detailed
                                      display of items when a single item is displayed on a single page.</td></tr>



					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_L_Layout_Id" type="radio" value="6" <%=lItemChecked6 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(6) %>
							<% small_help "6" %></td>

						</tr>

					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_L_Layout_Id" type="radio" value="7" <%=lItemChecked7 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(7) %>
							<% small_help "7" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_L_Layout_Id" type="radio" value="8" <%=lItemChecked8 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(8) %>
							<% small_help "8" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_L_Layout_Id" type="radio" value="9" <%=lItemChecked9 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(9) %>
							<% small_help "9" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_L_Layout_Id" type="radio" value="10" <%=lItemChecked10 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(10) %>
							<% small_help "10" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_L_Layout_Id" type="radio" value="12" <%=lItemChecked12 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplate(12) %>
							<% small_help "10" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_L_Layout_Id" type="radio" value="11" <%=lItemChecked11 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><B>Custom Layout:</b>
							<BR>OBJ_ID_OBJ - Item Id
                                                        <BR>OBJ_SKU_OBJ - Item Sku
                                                  <BR>OBJ_IMAGES_OBJ - Item Small Image
										<BR>OBJ_IMAGES_URL_OBJ - Item Small Image URL Only
										<BR>OBJ_IMAGEL_OBJ - Item Large Image
										<BR>OBJ_IMAGEL_URL_OBJ - Item Large Image URL Only
										<BR>OBJ_NAME_OBJ - Item Name
										<BR>OBJ_DETAIL_LINK_OBJ - URL to item detail page
										<BR>OBJ_DEPT_LINK_OBJ - URL to department page
										<BR>OBJ_DEPT_NAME_OBJ - Department Name
										<BR>OBJ_ATTRIBUTE - Select boxes for all item attributes (if any)
										<BR>OBJ_PRICE - Regular retail price
										<BR>OBJ_FINAL_PRICE_OBJ - Final price to charge shoppers
										<BR>OBJ_SPECIAL_PRICE_OBJ - Special price (if any)
										<BR>OBJ_MATRIX_OBJ - Quantity discount pricing details (if any)
										<BR>OBJ_UD_OBJ - User defined fields (if any)
										<BR>OBJ_DESCRIPTIONS_OBJ - Item small description
										<BR>OBJ_DESCRIPTIONL_OBJ - Item large description
										<BR>OBJ_QTY_OBJ - Order quantity input box
										<BR>OBJ_ORDER_OBJ - Add to Cart button
										<BR>OBJ_STOCK_OBJ - Number of items in stock (only shown in qty control is enabled)
										<BR>OBJ_ACCESSORY_OBJ - Item accessories
										<BR>OBJ_WEIGHT_OBJ - Weight
										<BR>OBJ_HANDLING_FEE_OBJ - Handling Fee
										<BR>OBJ_RECURRING_FEE_OBJ - Recurring Fee
										<BR>OBJ_RECURRING_DAYS_OBJ - Recurring Days
										<BR>OBJ_SPECIAL_PRICE_DISCOUNT%_OBJ - Special Price Discount %
										<BR>OBJ_SPECIAL_PRICE_START_DATE_OBJ - Special Price Start Date
                                                                                <BR>OBJ_SPECIAL_PRICE_END_DATE_OBJ - Special PRice End Date
                                                                                <BR>OBJ_EXT_FIELD1_OBJ - Extended Field 1
                                                  				<BR>OBJ_EXT_FIELD2_OBJ - Extended Field 2
										<BR>OBJ_EXT_FIELD3_OBJ - Extended Field 3
										<BR>OBJ_EXT_FIELD4_OBJ - Extended Field 4
										<BR>OBJ_EXT_FIELD5_OBJ - Extended Field 5

	                  <BR>OBJ_FRIEND_OBJ - Send to a friend text link
	                  <BR>OBJ_FRIEND_URL_OBJ - Send to a friend url
	                  <BR>
	                  <%= create_editor ("Item_L_Layout",Item_L_Layout,"[""Item Id"",""OBJ_ID_OBJ""],[""Item Sku"",""OBJ_SKU_OBJ""],[""Small Image"",""OBJ_IMAGES_OBJ""],[""Large Image"",""OBJ_IMAGEL_OBJ""],[""Item Name"",""OBJ_NAME_OBJ""],[""Item Name no link"",""OBJ_NAME_NOLINK_OBJ""],[""Item Attributes"",""OBJ_ATTRIBUTE""],[""Retail Price"",""OBJ_PRICE""],[""Final Price"",""OBJ_FINAL_PRICE_OBJ""],[""Special Price"",""OBJ_SPECIAL_PRICE_OBJ""],[""Quantity Discount"",""OBJ_MATRIX_OBJ""],[""User Defined Fields"",""OBJ_UD_OBJ""],[""Small Description"",""OBJ_DESCRIPTIONS_OBJ""],[""Large Description"",""OBJ_DESCRIPTIONL_OBJ""],[""Quantity Box"",""OBJ_QTY_OBJ""],[""Add to Cart"",""OBJ_ORDER_OBJ""],[""Qty in Stock"",""OBJ_QTY_OBJ""],[""Weight"",""OBJ_WEIGHT_OBJ""],[""Handling Fee"",""OBJ_HANDLING_FEE_OBJ""],[""Recurring Fee"",""OBJ_RECURRING_FEE_OBJ""],[""Recurring Days"",""OBJ_RECURRING_DAYS_OBJ""],[""Discount %"",""OBJ_SPECIAL_PRICE_DISCOUNT%_OBJ""],[""Special Start Date"",""OBJ_SPECIAL_PRICE_START_DATE_OBJ""],[""Special End Date"",""OBJ_SPECIAL_PRICE_END_DATE_OBJ""],[""Extended Field 1"",""OBJ_EXT_FIELD1_OBJ""],[""Extended Field 2"",""OBJ_EXT_FIELD2_OBJ""],[""Extended Field 3"",""OBJ_EXT_FIELD3_OBJ""],[""Extended Field 4"",""OBJ_EXT_FIELD4_OBJ""],[""Extended Field 5"",""OBJ_EXT_FIELD5_OBJ""],[""Send to Friend"",""OBJ_FRIEND_OBJ""]") %>
	                  <input type="hidden" name="Item_L_Layout_C" value="Op|String|||||Large Layout">
			 <% small_help "Custom Large Layout" %></td>
						</tr>
  </table>
			</div>
			<!-- CONTENT 2 ENDS HERE -->
			<!-- CONTENT 3 MAIN -->
			<div id="content3" style="visibility: hidden;" class=tpage>
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
                             <TR bgcolor='#FFFFFF'><TD colspan=3>The featured item layout is used for items 
					    shown in search results, the homepage, and as accessories</td></tr>




					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_F_Layout_Id" type="radio" value="6" <%=fItemChecked6 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplatef(6) %>
							<% small_help "6" %></td>

						</tr>

					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_F_Layout_Id" type="radio" value="7" <%=fItemChecked7 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplatef(7) %>
							<% small_help "7" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_F_Layout_Id" type="radio" value="8" <%=fItemChecked8 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplatef(8) %>
							<% small_help "8" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_F_Layout_Id" type="radio" value="9" <%=fItemChecked9 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplatef(9) %>
							<% small_help "9" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_F_Layout_Id" type="radio" value="10" <%=fItemChecked10 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplatef(10) %>
							<% small_help "10" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_F_Layout_Id" type="radio" value="12" <%=fItemChecked12 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><%= sItemTemplatef(12) %>
							<% small_help "10" %></td>

						</tr>
					<TR bgcolor='#FFFFFF'>
							<td width="3%" class="inputvalue"><input class="image" name="Item_F_Layout_Id" type="radio" value="11" <%=fItemChecked11 %>>&nbsp;</td>
							<td width="93%" class="inputvalue"><B>Custom Layout:</b>
							<BR>OBJ_ID_OBJ - Item Id
                                                        <BR>OBJ_SKU_OBJ - Item Sku
										<BR>OBJ_IMAGES_OBJ - Item Small Image
										<BR>OBJ_IMAGES_URL_OBJ - Item Small Image URL Only
										<BR>OBJ_IMAGEL_OBJ - Item Large Image
										<BR>OBJ_IMAGEL_URL_OBJ - Item Large Image URL Only
										<BR>OBJ_NAME_OBJ - Item Name
										<BR>OBJ_DETAIL_LINK_OBJ - URL to item detail page
										<BR>OBJ_DEPT_LINK_OBJ - URL to department page
										<BR>OBJ_DEPT_NAME_OBJ - Department Name
										<BR>OBJ_PRICE - Regular retail price
										<BR>OBJ_FINAL_PRICE_OBJ - Final price to charge shoppers
										<BR>OBJ_SPECIAL_PRICE_OBJ - Special price (if any)
										<BR>OBJ_MATRIX_OBJ - Quantity discount pricing details (if any)
										<BR>OBJ_DESCRIPTIONS_OBJ - Item small description
										<BR>OBJ_STOCK_OBJ - Number of items in stock (only shown in qty control is enabled)
										<BR>OBJ_WEIGHT_OBJ - Weight
										<BR>OBJ_HANDLING_FEE_OBJ - Handling Fee
										<BR>OBJ_RECURRING_FEE_OBJ - Recurring Fee
										<BR>OBJ_RECURRING_DAYS_OBJ - Recurring Days
										<BR>OBJ_SPECIAL_PRICE_DISCOUNT%_OBJ - Special Price Discount %
										<BR>OBJ_SPECIAL_PRICE_START_DATE_OBJ - Special Price Start Date
                                                  <BR>OBJ_SPECIAL_PRICE_END_DATE_OBJ - Special PRice End Date
                                                  <BR>OBJ_SPECIAL_PRICE_END_DATE_OBJ - Special PRice End Date
                                                  <BR>OBJ_EXT_FIELD1_OBJ
                                                  <BR>
	                  <%= create_editor ("Item_F_Layout",Item_F_Layout,"[""Item Id"",""OBJ_ID_OBJ""],[""Item Sku"",""OBJ_SKU_OBJ""],[""Small Image"",""OBJ_IMAGES_OBJ""],[""Large Image"",""OBJ_IMAGEL_OBJ""],[""Item Name"",""OBJ_NAME_OBJ""],[""Item Name no link"",""OBJ_NAME_NOLINK_OBJ""],[""Retail Price"",""OBJ_PRICE""],[""Final Price"",""OBJ_FINAL_PRICE_OBJ""],[""Special Price"",""OBJ_SPECIAL_PRICE_OBJ""],[""Quantity Discount"",""OBJ_MATRIX_OBJ""],[""Small Description"",""OBJ_DESCRIPTIONS_OBJ""],[""Qty in Stock"",""OBJ_QTY_OBJ""],[""Weight"",""OBJ_WEIGHT_OBJ""],[""Handling Fee"",""OBJ_HANDLING_FEE_OBJ""],[""Recurring Fee"",""OBJ_RECURRING_FEE_OBJ""],[""Recurring Days"",""OBJ_RECURRING_DAYS_OBJ""],[""Discount %"",""OBJ_SPECIAL_PRICE_DISCOUNT%_OBJ""],[""Special Start Date"",""OBJ_SPECIAL_PRICE_START_DATE_OBJ""],[""Special End Date"",""OBJ_SPECIAL_PRICE_END_DATE_OBJ""],[""Extended Field 1"",""OBJ_EXT_FIELD1_OBJ""]") %>
	                  <input type="hidden" name="Item_F_Layout_C" value="Op|String|0|2000|||Featured Layout">
			 <% small_help "Custom Featured Layout" %></td>
						</tr>
  </table>
			</div>
			<!-- CONTENT 3 ENDS HERE -->
		</td>
		</tr>	
			 
</table></td></tr>					
<tr bgcolor=#FFFFFF>
<td colspan='4' class='tpage'>		
<% createFoot thisRedirect,1 %>
</td>
</tr>

<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

</script>
<% end if %>

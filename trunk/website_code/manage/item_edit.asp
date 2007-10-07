<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/department_list.asp"-->
<!--#include file="include/location_list.asp"-->
<!--#include file="editor_include.asp"-->
<!--#include virtual="common/common_functions.asp"-->
<!--#include file="help/item_edit.asp"-->

<% 
Calendar=1
Switch_Name=Site_Name
'LOAD CURRENT VALUES FROM THE DATABASE
End_date =	DateAdd("m", +1, Now())
Start_date = Now()

Item_Id = Request.QueryString("Id")
if item_id = "" then
   Item_Id = Request.QueryString("Item_Id")
   if Item_Id = "" then
      item_id=0
   end if
end if

if isNumeric(Item_Id) and Item_Id<>0 then
	sql_select="SELECT * FROM Store_Items WITH (NOLOCK) WHERE Item_Id = "&Item_Id&" and Store_id="&Store_id
	rs_store.open sql_select,conn_store,1,1
	rs_store.MoveFirst
	Sub_department_Id=rs_store("Sub_department_Id")

	Item_Page_Name = checkStringForQBack(Rs_store("Item_Page_Name"))
	Item_Name = checkStringForQBack(Rs_store("Item_Name"))
	Item_Id = Rs_store("Item_Id")
	Item_Sku = Rs_store("Item_Sku")
	Show_homepage = Rs_store("Show_homepage")

	Retail_Price = Rs_store("Retail_Price")
	Wholesale_Price = Rs_store("Wholesale_Price")
	ImageS_Path = Rs_store("ImageS_Path")
	ImageL_Path = Rs_store("ImageL_Path")

	Taxable = Rs_store("Taxable")
	Use_Price_By_Matrix = Rs_store("Use_Price_By_Matrix")

	Description_S = Rs_store("Description_S")
	Description_L = Rs_store("Description_L")
	Item_Weight = Rs_store("Item_Weight")
	Item_Handling = Rs_store("Item_Handling")
	Quantity_in_stock = Rs_store("Quantity_in_stock")
	Quantity_Control = Rs_store("Quantity_Control")

	Quantity_Control_Number = Rs_store("Quantity_Control_Number")
	Quantity_Minimum = Rs_store("Quantity_Minimum")
	Retail_Price_special_Discount = Rs_store("Retail_Price_special_Discount")
	Special_start_date = Rs_store("Special_start_date")
	Special_end_date = Rs_store("Special_end_date")

	Shipping_Fee = Rs_store("Shipping_Fee")
	Waive_Shipping = Rs_store("Waive_Shipping")
	Item_pin=Rs_store("Item_pin")
	Ship_Location_Id = Rs_store("Ship_Location_Id")
	File_Location = Rs_store("File_Location")
	Hide_Stock = rs_store("Hide_Stock")
    Fractional = rs_store("Fractional")

	Cust_Price = Rs_store("Enable_Cust_Price")
	Hide_Price = Rs_store("Hide_Price")
	Show = Rs_store("Show")
	U_d_1_name = rs_store("U_d_1_name")
	U_d_1 = Rs_store("U_d_1")
	U_d_2_name = rs_store("U_d_2_name")
	U_d_2 = Rs_store("U_d_2")
	U_d_3_name = rs_store("U_d_3_name")
	U_d_3 = Rs_store("U_d_3")
	U_d_4_name = rs_store("U_d_4_name")
	U_d_4 = Rs_store("U_d_4")
	U_d_5_name = rs_store("U_d_5_name")
	U_d_5 = Rs_store("U_d_5")

	M_d_1 = Rs_store("M_d_1")
	M_d_2 = Rs_store("M_d_2")
	M_d_3 = Rs_store("M_d_3")
	M_d_4 = Rs_store("M_d_4")
	M_d_5 = Rs_store("M_d_5")
	
	Item_Remarks = Rs_store("Item_Remarks")
	Meta_Keywords = rs_Store("Meta_Keywords")
	Meta_Description = rs_Store("Meta_Description")
	Meta_Title = rs_Store("Meta_Title")
	Recurring_days = rs_Store("Recurring_days")
	Recurring_fee = rs_Store("Recurring_fee")
	Brand = rs_Store("Brand")
	Condition = rs_Store("Condition")
	Product_Type = rs_Store("Product_Type")
	View_Order = rs_Store("View_Order")

	Custom_Link=rs_Store("Custom_Link")

	Attributes_Exist=rs_Store("Attributes_Exist")
	Accessories_Exist=rs_store("Accessories_Exist")
	Configs_Exist=rs_Store("Configs_Exist")
	rs_store.Close
	
	sDept_ids=fn_get_dept_ids (item_id)
else
    Item_Handling=0
    Item_Weight=0
    Quantity_in_Stock=0
    Quantity_Minimum=1
    View_Order=0
    Quantity_Control=0
    Wholesale_Price=0
    Shipping_Fee=0
    Quantity_Control_Number=0
    Recurring_days=0
    Recurring_fee=0
    Brand="NA"
    Condition="New"
    Product_Type="Unknown"
    Show_homepage=0
    Use_Price_By_Matrix=0
    Taxable=-1
    Fractional=0
    Department_Id=0
    Sub_Department_Id=0 
    Retail_Price=0
    Item_Weight=0
    Retail_Price_special_Discount=0
    Shipping_Fee=0
    Show=-1
    Special_start_date = Start_Date
	Special_end_date = End_Date
	Ship_Location_Id=0
	sDept_ids="0"
end if

if sDept_ids="" then
	sDept_ids="0"
end if
if Taxable= - 1 then
	checked_Taxable = "checked"
end if
if Show_homepage= -1 then
	checked_homepage = "checked"
end if
if Use_Price_By_Matrix= - 1 then 
	checked_Use_Price_By_Matrix = "checked"
end if
if Quantity_Control= -1 then 
	checked_Quantity_Control = "checked"
end if
if Not IsDate(Special_start_date) then
	Special_start_date = Start_Date
end if
if Not IsDate(Special_end_date) then 
	Special_end_date = End_Date
end if
if Waive_Shipping= 0 then
	checked_Waive_Shipping = ""
else
	checked_Waive_Shipping = "checked"
end if
if Hide_Stock=-1 then
	 checked_Hide_Stock = "checked"
end if
if Cust_Price= 0 then
	checked_Cust = ""
else
	checked_Cust = "checked"
end if
if Hide_Price= - 1 then
	checked_Hide = "checked"
end if
if Show= - 1 then
	checked_Show = "checked"
end if
if U_d_1= - 1 then 
	checked_U_d_1 = "checked"
end if
if U_d_2= - 1 then
	checked_U_d_2 = "checked"
end if
if U_d_3= - 1 then
	checked_U_d_3 = "checked"
end if
if U_d_4= - 1 then
	checked_U_d_4 = "checked"
end if
if U_d_5= - 1 then
	checked_U_d_5 = "checked"
end if
if Fractional=-1 then
	Fractional_Checked="checked"
end if
if Item_pin=-1 then
	Item_pin_Checked = "checked"
end if

if Sub_Department_Id = 0 then
	Full_Name = ""
end if

addPicker = 1
sFormAction = "Inventory_action.asp"
sName = "Store_Activation"
sFormName = "Item_Edit"
sTrialAdd=1
if Item_Id=0 then
    sTitle = "Advanced Add Item"
    sFullTitle = "Inventory > <a href=edit_items.asp class=white>Items</a> > Advanced Add"
    if Service_Type < 10 then
       Items = 0
       sql_select = "Select Count(Item_Id) as Num_Items from Store_Items WITH (NOLOCK) where Store_Id=" & Store_Id
       rs_Store.open sql_select,conn_store,1,1
       if rs_store.bof = false then
	  Items = rs_Store("Num_Items")
       end if
       if Items >= 10 and (Service_Type = 1 or Service_Type = 0) then
		sTrialAdd = 0
		sLimit = 10
        elseif Items >= 50 and Service_Type = 3 then
		sTrialAdd = 0
		sLimit = 50
		  elseif Items >= 100 and Service_Type = 5 then
		sTrialAdd = 0
		sLimit = 100
		  elseif Items >= 500 and Service_Type = 7 then
		sTrialAdd = 0
		sLimit = 500
		elseif Items >= 1000 and Service_Type = 9 then
		sTrialAdd = 0
		sLimit = 1000
		  end if
    end if
    rs_store.close
else
    sTitle = "Advanced Edit Item - "&Item_Name
    sFullTitle = "Inventory > <a href=edit_items.asp class=white>Items</a> > Advanced Edit - "&Item_Name
end if

sNeedTabs=1
sCommonName="Item"
sCancel="edit_items.asp"
sSubmitName = "Item_Edit"
thisRedirect = "item_edit.asp"
sMenu="inventory"
sQuestion_Path = "inventory/add_item.htm"
createHead thisRedirect

if sTrialAdd = 0 then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		You may not add any more items, your level of service is limited to <%= sLimit %>, your store currently has <%= Items %> items.
		<BR><BR>
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0 %>
<% else %>


	<tr bgcolor='#FFFFFF'>
		<td width="100%" height="22">
		<input type="hidden" value="<%= Item_Id %>" name="Item_Id">
		<input type="hidden" value="<%= Sub_Department_Id %>" name="Sub_Department_Id">

		<input type="button" class="Buttons" value="Item Search" name="Back_To_Item_List" OnClick='JavaScript:self.location="edit_Items.asp?<%= sAddString %>"'>
                <input type="button" class="Buttons" value="Add Department" name="Add_Department" OnClick='JavaScript:self.location="store_dept_add.asp"'>
                <% if Item_Id<>0 then %>
                <input class=buttons type="button" value="Preview Item" name="Editor" OnClick="javascript:goPreview('<%= fn_item_url("Preview",item_page_name) %>')">
			        	<BR><input type="button" class="Buttons" value="Add Item" name="Add" OnClick='JavaScript:self.location="Item_edit.asp?<%= sAddString %>"'>
			<input type="button" class="Buttons" value="Copy Item" name="Copy" OnClick='JavaScript:self.location="Item_copy.asp?Sub_Department_Id=<%= Sub_Department_Id %>&Item_Id=<%= Item_Id%>&Item_Name=<%= server.urlencode(Item_Name) %>&<%= sAddString %>"'>
                        <input type="button" class="Buttons" value="Basic Item Edit" name="Add_Department" OnClick='JavaScript:self.location="item_basic_edit.asp?Item_Id=<%= Item_Id%>"'>
                <% end if %>

		</td>
	 </tr>



	<tr bgcolor='#FFFFFF'><td width="724" align=center valign=top height=45>
	<table border=0 cellspacing=0 cellpadding=0 width=724>

	<!-- TAB MENU STARTS HERE -->

		<tr bgcolor='#FFFFFF'>
		<td align="center" valign=top height=65 width='100%' colspan='7'>
		<script type="text/javascript" language="JavaScript1.2" src="include/tabs-xp.js"></script>
		<script language="javascript1.2">
		var bmenuItems =
		[
		["-"],
		["Main", "content1",,,"Main","Main"],
		["Options",  "content2",,,"Options","Options",],
		["Shipping", "content3",,,"Shipping","Shipping",],
		["Inventory", "content4",,, "Inventory","Inventory",],
		["Search Engines",  "content5",,, "Search Engines","Search Engines",],
		["Recurring",  "content6",,,"Recurring","Recurring",],
		["Google Base",  "content13",,,"Google Base","Google Base",],
		["-"],
		["-"],
		["$Specials ",  "content7",,,"Specials","Specials"],
		["Softgoods",  "content8",,,"Softgoods","Softgoods"],
		["User Fields",  "content9",,,"User Fields","User Fields"],
		["Extended Fields",  "content10",,,"Extended Fields","Extended Fields"],
		["3rd Party",  "content11",,,"3rd Party","3rd Party"],
		["Advanced",  "content12",,,"Advanced","Advanced",],
		];

		apy_tabsInit();
		</script>
		</td>
		</tr>
		<%
		
		strUpgradeMsg = "This feature is not available at your current level of service.<br><br>"
		strGold = "GOLD"
		strBronze = "BRONZE"
		strMsg = " Service or higher is required. <a href='billing.asp' class=link>Click here to upgrade now.</a>"
		%>

		
		<!-- TAB MENU ENDS HERE -->
		
		<TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='25'>
			<!-- CONTENT 1 MAIN -->
			<div id="content1" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				<tr bgcolor='#FFFFFF'>
						<td width="30%" class="inputname">Department</td>
						<td width="70%" class="inputvalue">
						<%= create_dept_list ("Department_Id",sDept_ids,5,"") %>
					<input type="hidden" name="Department_Id_C" value="Re|String|0|250|||Department">
                    <input type="hidden" name="Old_Department_Id" value="<%= sDept_ids %>">

					<% small_help "Store Department" %></td>
				</tr>
							 
				<tr bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">SKU</td>
				<td width="70%" class="inputvalue">
							<input name="Item_Sku" Value="<%= Item_Sku %>" size="60">
							<input type="hidden" name="Item_Sku_C" value="Re|String|0|200|||SKU">
							<% small_help "SKU" %></td>
				</tr>
					  
					<tr bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Name</td>
				<td width="70%" class="inputvalue">
							<input name="Item_Name" value="<%= Item_Name %>" size="60" maxlength=250>
							<input type="hidden" name="Item_Name_C" value="Re|String|0|250|||Name">
							<% small_help "Name" %></td>
				</tr>
				
				<tr bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Page Filename</td>
				<td width="70%" class="inputvalue" nowrap>
							<input name="Item_Page_Name" value="<%= Item_Page_Name %>" maxlength=100 size="60" onKeyPress="return goodchars(event,'0123456789-abcdefghijklmnopqrstuvwxyz_')">-detail.htm
							<input type="hidden" name="Item_Page_Name_C" value="Op|String|0|100||@,$,%, ,',&,.,/,(,),`,;,#,!,?,^,®,™,†,©,”,“,½,¾,é,…,à,è,’|Page Filename">
							<% small_help "item_Filename" %></td>
				</tr>
					  
					<tr bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Retail Price</td>
				<td width="70%" class="inputvalue">
							<%= Store_Currency%><input name="Retail_Price" value="<%= Retail_Price %>" size="10" onKeyPress="return goodchars(event,'0123456789.-')" maxlength=10>
							<INPUT type="hidden"  name="Retail_Price_C" value="Re|Integer|||||Retail Price">
							<% small_help "Retail Price" %></td>
				</tr>
		
		                <% if Service_Type => 5 then %>
				<tr bgcolor='#FFFFFF'><td width="30%" class="inputname" colspan=2>Short Description<BR>
						<%= create_editor ("Description_S",Description_S,"") %>
						<input type="hidden" name="Description_S_C" value="Op|String|0|4000|||Short Description">
						<% small_help "Short Description" %></td>
			</tr>
			<% else %>
			 <input type=hidden name="Description_S" value="">

			<% end if %>
	  
				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Small Image</td>
			<td width="70%" class="inputvalue">
						<input type="text" name="ImageS_Path" value="<%= ImageS_Path %>" size="60" >
						<input type="hidden" name="ImageS_Path_C" value="Op|String|0|100|||Small Image">
						<font size="1" color="#000080"><a class="link" href="JavaScript:goImagePicker('ImageS_Path');"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
						<a class="link" href="JavaScript:goFileUploader('ImageS_Path');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
					<% small_help "Small Image" %></td>
			</tr>
	  
				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname" colspan=2>Large Description<BR>

						<%= create_editor ("Description_L",Description_L,"") %>
						<% small_help "Large Description" %></td>
			</tr>
	  
				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Large Image</td>
			<td width="70%" class="inputvalue">
						<input type="text" name="ImageL_Path" value="<%= ImageL_Path %>" size="60" >
						<input type="hidden" name="ImageL_Path_C" value="Op|String|0|100|||Large Image">
						<font size="1" color="#000080"><a class="link" href="JavaScript:goImagePicker('ImageL_Path');"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
				<a class="link" href="JavaScript:goFileUploader('ImageL_Path');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>

				<% small_help "Large Image" %></td>
			</tr>
			
			</table>
			</div>
			<!-- CONTENT 1 MAIN -->
			
			<!-- CONTENT 2 -->
			<div id="content2" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
			
			
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Homepage</td>
			<td width="70%" class="inputvalue">
						<input class="image" type="checkbox" <%= checked_homepage %> name="Show_Homepage" value="-1" >
						<% small_help "Homepage" %></td>

			</tr>
			<% if Service_Type => 5 then %>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Taxable</td>
			<td width="70%" class="inputvalue">
						<input class="image" type="checkbox" <%= checked_Taxable %> name="Taxable" value="-1" >
						<% small_help "Taxable" %></td>
			</tr>
	  
				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Visible</td>
			<td width="70%" class="inputvalue">
						<input class="image" type="checkbox" <%= checked_Show %> name="Show" value="-1" >
						<% small_help "Visible" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Hide Price</td>
			<td width="70%" class="inputvalue">
						<input class="image" type="checkbox" <%= checked_Hide %> name="Hide_Price" value="-1" >
						<% small_help "Hide Price" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Custom Pricing</td>
			<td width="70%" class="inputvalue">
						<input class="image" type="checkbox" <%= checked_Cust %> name="Cust_Price" value="-1" >
						<% small_help "Cust Price" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">View Order</td>
			<td width="70%" class="inputvalue">
						<input	name="View_Order" value="<%= View_Order %>" size="10" onKeyPress="return goodchars(event,'0123456789.-')">
						<INPUT type="hidden" name="View_Order_C" value="Re|Integer|||||View_Order">
						<% small_help "View Order" %></td>
			</tr>
			<% else %>
			<input type="hidden" name="Taxable" value="1" >
			<input type="hidden" name="Show" value="1" >
			<input type="hidden" name="Hide_Price" value="0" >
			<input type="hidden" name="Cust_Price" value="0" >
			<input type="hidden" name="View_Order" value="0" >
			<% end if %>
				
				
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Cost</td>
			<td width="70%" class="inputvalue">
						<%= Store_Currency%><input	name="Wholesale_Price" value="<%= Wholesale_Price %>" size="10" onKeyPress="return goodchars(event,'0123456789.-')" maxlength=10>
						<INPUT type="hidden" name="Wholesale_Price_C" value="Op|Integer|||||Cost Price">
						<% small_help "Wholesale Price" %></td>
			</tr>

			  <% if Service_Type < 7 then %>
			<input type="hidden" name="Use_Price_By_Matrix" value="0" >
		

		<% elseif Item_id<>0 then
			
			if Attributes_Exist then 
				sCountAttr="Yes"
			else
				sCountAttr="None"
			end if
			if Configs_Exist then
				sCountConfig="Yes"
			else
				sCountConfig="None"
			end if
			if Accessories_Exist then
				sCountAccessory="Yes"
			else
				sCountAccessory="None"
			end if
			'sql_select = "select count(price_group_id) as count from store_items_price_group where store_id="&Store_Id&" and Item_Id="&Item_Id
			'rs_store.open sql_select,conn_store,1,1
			'sCountPriceGroup=rs_Store("count")
			'rs_store.close
			'sql_select = "select count(price_matrix_id) as count from store_items_price_matrix where store_id="&Store_Id&" and Item_Id="&Item_Id
			'rs_store.open sql_select,conn_store,1,1
			'sCountPriceMatrix=rs_Store("count")
			'rs_store.close
		         %>
				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Qty Discount</td>
			<td width="70%" class="inputvalue">
						<input class="image" type="checkbox" <%= checked_Use_Price_By_Matrix %> name="Use_Price_By_Matrix" value="-1" >
						
						
						<input type="button" class="Buttons" value="Price Matrix (<%= sCountPriceMatrix %>)" name="Price_Matrix" OnClick='JavaScript:self.location="Price_Matrix_list.asp?Item_Id=<%= Item_Id %>&Sub_Department_Id=<%= Sub_Department_Id %>&<%= sAddString %>"'>
						<% small_help "Qty Discount" %></td>
			</tr>
			<% end if %>	
				
			<tr bgcolor='#FFFFFF'>
			<td class="inputname" colspan=2>Item Remarks Not seen by customers
			
						<input readonly type=text name=remLenRem size=3 class=char maxlength=3 value="<%= 200-len(Item_Remarks) %>" class=image><font size=1><I>characters left</i></font>
						<BR>
						<textarea name="Item_Remarks" cols="83" rows="2"><%= Item_Remarks %></textarea>
						<input type="hidden" name="Item_Remarks_C" value="Op|String|0|200|||Item Remarks">
						<% small_help "Item Remarks" %></td>
			</tr>
			
			</table>
			</div>
			<!-- CONTENT 2 -->
			
			<!-- CONTENT 3 -->
			<div id="content3" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
			
				 <% if instr(Shipping_Classes,"2") or instr(Shipping_Classes,"6") or instr(Shipping_Classes,"7") then %>
			  <tr bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">Shipping Weight</td>
			<td width="70%" class="inputvalue">
						<input	name="Item_Weight" value="<%= Item_Weight %>" size="8" onKeyPress="return goodchars(event,'0123456789.-')">
						pounds <INPUT type="hidden"  name="Item_Weight_C" value="Re|Integer|||||Weight">
				<% small_help "Weight" %></td>
			</tr>

			<% else %>
				<input type=hidden name="Item_Weight" value="<%= Item_Weight %>">
			<% end if %>
			<tr bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">Handling Fee</td>
			<td width="70%" class="inputvalue">
						<%= Store_Currency%><input	name="Item_Handling" value="<%= Item_Handling %>" size="8" onKeyPress="return goodchars(event,'0123456789.-')">
						<INPUT type="hidden"  name="Item_Handling_C" value="Re|Integer|||||Handling">
				<% small_help "Handling" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">Waive Shipping and Handling</td>
			<td width="70%" class="inputvalue">
						<input class="image" type="checkbox" <%= checked_Waive_Shipping %> name="Waive_Shipping" value="-1" >

				<% small_help "Waive Shipping" %></td>
			</tr>

	

							<tr bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">Shipping From</td>
			<td width="70%" class="inputvalue">
						<%= create_location_list ("Ship_Location_Id",Ship_Location_Id,1) %>
									
				<% small_help "Shipfrom" %></td>
			</tr>
			


		
			<% if instr(Shipping_Classes,"3") > 0 then %>

					  <tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Per Item Shipping Fee</td>
			<td width="70%" class="inputvalue">
						<%= Store_Currency%><input name="Shipping_Fee" value="<%= Shipping_Fee %>" size="10" onKeyPress="return goodchars(event,'0123456789.-')" maxlength=10>
						<INPUT type="hidden"  name="Shipping_Fee_C" value="Re|Integer|||||Shipping Fee">
				<% small_help "Shipping Fee" %></td>
			</tr>
			<% else %>
				<input type=hidden name="Shipping_Fee" value="<%= Shipping_Fee %>">
			<% end if %>

			</table>
			</div>
			<!-- CONTENT 3 -->
			
			<% if Service_Type < 5 then %>
			<input type=hidden name="Quantity_in_stock" Value="<%= Quantity_in_stock %>" size="8">
			<input type="hidden" name="Quantity_Control" value="0" >
			<input type="hidden" name="Hide_Stock" value="0" >
			<input type="hidden" name="Quantity_Control_Number" value="<%= Quantity_Control_Number %>" size="4">
			<input type="hidden" name="Quantity_Minimum" value="<%= Quantity_Minimum %>" size="4">
			<input type="hidden" name="Fractional" value="0">
			<input type=hidden	name="Meta_Keywords" value="<%= Meta_Keywords %>" size="30">
			<input type=hidden	name="Meta_Description" value="<%= Meta_Description %>" size="30">
			<input type=hidden	name="Meta_Title" value="<%= Meta_Title %>" size="30">
				<!--	 CONTENT 4 -->
				<div id="content4" style="visibility: hidden;" class="tabPage">
				<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
				<tr bgcolor='#FFFFFF'>
				<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strGold%><%=strMsg%></td>
				</tr>
				</table>
				</div>
				<!--	 CONTENT 4 -->
				
				<!--	 CONTENT 5 -->
				<div id="content5" style="visibility: hidden;" class="tabPage">
				<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
				<tr bgcolor='#FFFFFF'>
				<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strGold%><%=strMsg%></td>
				</tr>
				</table>
				</div>
				<!--	 CONTENT 5 -->
			<% else %>
			
				<!-- CONTENT 4 -->
				<div id="content4" style="visibility: hidden;" class="tabPage">
				<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
					<tr bgcolor='#FFFFFF'>
					<td width="30%"><b>Inventory Management</b></td>
			<td width="75%" colspan=2>&nbsp;</td>
			</tr>
			 <tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Stock</td>
			<td width="70%" class="inputvalue">
						<input	name="Quantity_in_stock" Value="<%= Quantity_in_stock %>" size="8" onKeyPress="return goodchars(event,'0123456789.-')">
						<INPUT type="hidden"  name="Quantity_in_stock_C" value="Re|Integer|||||Stock">
						<% small_help "Stock" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Qty Control</td>
			<td width="70%" class="inputvalue">
						<input class="image" type="checkbox" <%= checked_Quantity_Control %> name="Quantity_Control" value="-1" >
						<% small_help "Qty Control" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Control #</td>
			<td width="70%" class="inputvalue">
						<input name="Quantity_Control_Number" value="<%= Quantity_Control_Number %>" size="4" onKeyPress="return goodchars(event,'0123456789.-')">
						<INPUT type="hidden"  name="Quantity_Control_Number_C" value="Re|Integer|||||Quantity Control">
						<% small_help "Control #" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Hide Stock</td>
			<td width="70%" class="inputvalue">
						<input class="image" type="checkbox" <%= checked_Hide_Stock %> name="Hide_Stock" value="-1" >
						<% small_help "Hide Stock" %></td>
			</tr>

			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Min Quantity</td>
			<td width="70%" class="inputvalue">
							 <input name="Quantity_Minimum" value="<%= Quantity_Minimum %>" size="4" onKeyPress="return goodchars(event,'0123456789.-')">
						<INPUT type="hidden"  name="Quantity_Minimum_C" value="Re|Integer|||||Min Quantity">
						<% small_help "Min Quantity" %></td>
			</tr>
			 <tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Fractional</td>
			<td width="70%" class="inputvalue">
						<input class="image" name="Fractional" type="checkbox" <%= Fractional_checked%> value="-1">
						<% small_help "Fractional" %></td>
			</tr>
				</table>
				</div>
				<!-- CONTENT 4 -->
			
			
				<!-- CONTENT 5 -->
				<div id="content5" style="visibility: hidden;" class="tabPage">
				<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
					<tr bgcolor='#FFFFFF'>
					<td width="30%"><b>Search Engine Tags</b></td>
			<td width="75%" colspan=2>Help get your items listed in search engines</td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Title</td>
			<td width="70%" class="inputvalue">
						<input type=text	name="Meta_Title" value="<%= Meta_Title %>" maxlength=100 size=60>
						<input type="hidden" name="Meta_Title_C" value="Op|String|0|100|||Search Engine Title">
						<% small_help "Meta_Title" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td class="inputname" colspan=2>Keywords
			
						<input readonly type=text name=remLenMKey size=3 class=char maxlength=3 value="<%= 250-len(Meta_Keywords) %>" class=image><font size=1><I>characters left</i></font>
            <BR><textarea	name="Meta_Keywords" cols=83 rows=3 onKeyDown="textCounter(this.form.Meta_Keywords,this.form.remLenMKey,250);" onKeyUp="textCounter(this.form.Meta_Keywords,this.form.remLenMKey,250);"><%= Meta_Keywords %></textarea>
						<input type="hidden" name="Meta_Keywords_C" value="Op|String|0|250|||Search Engine Keywords">
						<% small_help "Meta_Keywords" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td class="inputname" colspan=2>Description
			
						<input readonly type=text name=remLenMDes size=3 class=char maxlength=3 value="<%= 500-len(Meta_Description) %>" class=image><font size=1><I>characters left</i></font>
            <BR><textarea	name="Meta_Description" cols=83 rows=5 onKeyDown="textCounter(this.form.Meta_Description,this.form.remLenMDes,500);" onKeyUp="textCounter(this.form.Meta_Description,this.form.remLenMDes,500);"><%= Meta_Description %></textarea>
						<input type="hidden" name="Meta_Description_C" value="Op|String|0|500|||Search Engine Description">
						<% small_help "Meta_Description" %></td>
			</tr>
				</table>
				</div>
				<!-- CONTENT 5 -->
			
			<% end if %>
			
			
			<% if Service_Type < 5 then %>
				<input type=hidden name="Brand" value="<%= Brand %>">
				<input type=hidden name="Condition" value="<%= Condition %>">
				<input type=hidden name="Product_Type" value="<%= Product_Type %>">
				<input type=hidden name="Recurring_days" value="<%= Recurring_days %>">
				<input type=hidden name="Recurring_fee" value="<%= Recurring_fee %>">
				<input type=hidden name="Retail_Price_special_Discount" value="<%= Retail_Price_special_Discount %>">
				<input type=hidden name="Special_Start_Date" value="<%= FormatDateTime(Special_start_date,2) %>">
				<input type=hidden name="Special_End_Date" value="<%=  FormatDateTime(Special_end_date,2) %>">
				<input type=hidden name="File_Location" value="<%= File_Location %>"">
				<input type="hidden" name="Item_pin" value="0">
				<input type=hidden name="U_d_1" value="0" >
				<input type=hidden name="U_d_1_name" value="<%= U_d_1_name %>">
				<input type=hidden name="U_d_2" value="0" >
				<input type=hidden name="U_d_2_name" value="<%= U_d_2_name %>">
				<input type=hidden name="U_d_3" value="0" >
				<input type=hidden name="U_d_3_name" value="<%= U_d_3_name %>">
				<input type=hidden name="U_d_4" value="0" >
				<input type=hidden name="U_d_4_name" value="<%= U_d_4_name %>">
				<input type=hidden name="U_d_5" value="0" >
				<input type=hidden name="U_d_5_name" value="<%= U_d_5_name %>">


				<input type=hidden name="M_d_1" value="<%= M_d_1 %>" size="30">
				<input type=hidden name="M_d_2" value="<%= M_d_2 %>" size="30">
				<input type=hidden name="M_d_3" value="<%= M_d_3 %>" size="30">
				<input type=hidden name="M_d_4" value="<%= M_d_4 %>" size="30">
				<input type=hidden name="M_d_5" value="<%= M_d_5 %>" size="30">
				
				
				<div id="content6" style="visibility: hidden;" class="tabPage">
				<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
				<tr bgcolor='#FFFFFF'>
				<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strGold%><%=strMsg%></td>
				</tr>
				</table>
				</div>
				
				<div id="content7" style="visibility: hidden;" class="tabPage">
				<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
				<tr bgcolor='#FFFFFF'>
				<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strGold%><%=strMsg%></td>
				</tr>
				</table>
				</div>
				
				<div id="content8" style="visibility: hidden;" class="tabPage">
				<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
				<tr bgcolor='#FFFFFF'>
				<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strGold%><%=strMsg%></td>
				</tr>
				</table>
				</div>
				
				<div id="content9" style="visibility: hidden;" class="tabPage">
				<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
				<tr bgcolor='#FFFFFF'>
				<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strGold%><%=strMsg%></td>
				</tr>
				</table>
				</div>
				
				<div id="content10" style="visibility: hidden;" class="tabPage">
				<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
				<tr bgcolor='#FFFFFF'>
				<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strGold%><%=strMsg%></td>
				</tr>
				</table>
				</div>

			<% else %>
			<div id="content6" style="visibility: hidden;" class="tabPage">
				<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
			<tr bgcolor='#FFFFFF'>
			<td width="30%"><b>Recurring</b></td>
			<td width="75%" colspan=2>Note that recurring billing transactions are only handled automatically by Paypal and Worldpay.  For all other gateways recurring billing will be shown to the user but the merchant is responsible for actual recurring charges at the correct time.</font></td>
			</tr>
	
				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Recurring Fee</td>
			<td width="70%" class="inputvalue">
						<%= Store_Currency%><input name="Recurring_fee" value="<%= Recurring_fee %>" size="5" value="0" onKeyPress="return goodchars(event,'0123456789.-')">
						<INPUT type="hidden"  name="Recurring_fee_C" value="Op|Integer|||||Recurring Fee">
						<% small_help "Recurring_Fee" %></td>
			</tr>

				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Recurring Interval</td>
			<td width="70%" class="inputvalue">
						<input name="Recurring_days" value="<%= Recurring_days %>" size="5" onKeyPress="return goodchars(event,'0123456789')">days
						<INPUT type="hidden"  name="Recurring_days_C" value="Op|Integer|||||Recurring Days">
						<% small_help "Recurring_Days" %></td>
			</tr>
			
			</table>
			</div>
			<div id="content13" style="visibility: hidden;" class="tabPage">
				<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
			<tr bgcolor='#FFFFFF'>
			<td width="30%"><b>Google Base Fields</b></td>
			<td width="75%" colspan=2>The following fields are required to be sent as part of the Google Base feed.</td>
			</tr>
	
				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Condition</td>
			<td width="70%" class="inputvalue">
				<select name="Condition">
				<option value="<%= Condition %>" selected><%= Condition %></option>
				<option value="New">New</option>
				<option value="Used">Used</option>
				<option value="Refurbished">Refurbished</option></select>
						<% small_help "Condition" %></td>
			</tr>

				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Brand</td>
			<td width="70%" class="inputvalue">
						<input name="Brand" value="<%= Brand %>" size="60" maxlength=100>
						<INPUT type="hidden"  name="Brand_C" value="Op|String|0|100|||Brand">
						<% small_help "Brand" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Product Type</td>
			<td width="70%" class="inputvalue">
						<input name="Product_Type" value="<%= Product_Type %>" size="60" maxlength=20>
						<INPUT type="hidden"  name="Product_Type_C" value="Op|String|0|20|||Product Type">
						<% small_help "Product_Type" %></td>
			</tr>
			
			</table>
			</div>
			
			<div id="content7" style="visibility: hidden;" class="tabPage">
				<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				<tr bgcolor='#FFFFFF'>
			<td width="30%"><b>Specials</b></td>
			<td width="75%" colspan=2>Not effective when quantity discounts are enabled.</font></td>
			</tr>

				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Special Price</td>
			<td width="70%" class="inputvalue">Discount
						<input name="Retail_Price_special_Discount" value="<%= Retail_Price_special_Discount %>" size="5" onKeyPress="return goodchars(event,'0123456789.-')">%
						<INPUT type="hidden"  name="Retail_Price_special_Discount_C" value="Op|Integer|||||Discount">
						<% small_help "Special Price" %></td>
			</tr>
	     <SCRIPT LANGUAGE="JavaScript" ID="jscal1">
      var now = new Date();
      var cal1 = new CalendarPopup("testdiv1");
      cal1.showNavigationDropdowns();
      </SCRIPT>
				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Dates</td>
			<td width="70%" class="inputvalue">Between
				<input name="Special_Start_Date" value="<%= FormatDateTime(Special_start_date,2) %>" size="10" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')">
						<A HREF="#" onClick="cal1.select(document.forms[0].Special_Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Special_Start_Date.value=='')?document.forms[0].Special_Start_Date.value:null); return false;" TITLE="Start Date" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>
			     <INPUT type="hidden"  name="Special_Start_Date_C" value="Re|date|||||Special Start Date"> and
						<input name="Special_End_Date" value="<%=  FormatDateTime(Special_end_date,2) %>" size="10" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')">
						<A HREF="#" onClick="cal1.select(document.forms[0].Special_End_Date,'anchor2','M/d/yyyy',(document.forms[0].Special_End_Date.value=='')?document.forms[0].Special_End_Date.value:null); return false;" TITLE="End Date" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>
           <INPUT type="hidden"  name="Special_end_Date_C" value="Re|date|||||Special End Date">
						<% small_help "Dates" %></td>
			</tr>
		</table>
		</div>
		
		<div id="content8" style="visibility: hidden;" class="tabPage">
				<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				<tr bgcolor='#FFFFFF'>
			<td width="30%"><b>Electronic Software Delivery</b></td>
			<td width="75%" colspan=2>Fill this out if your item is a software file, ie MP3, Game, etc.</td>
			</tr>
	  
				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Filename</td>
			<td width="70%" class="inputvalue">
						<input type="text" name="File_Location" value="<%= File_Location %>" size="60" maxlength=250>
						<input type="hidden" name="File_Location_C" value="Op|String|0|250|||Filename">
						<a class="link" href="javascript:goFilePicker('File_Location')"><img border="0" src="images/image.gif" width="23" height="22" alt="File Picker"></a>
						<% small_help "Filename" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Pin Delivery</td>
			<td width="70%" class="inputvalue">
						<input class="image" name="Item_pin" type="checkbox" <%= Item_pin_checked %> value="-1">
						<% small_help "Item_pin" %></td>
			</tr>
		</table>
		</div>
		
		<div id="content9" style="visibility: hidden;" class="tabPage">
				<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				
			<tr bgcolor='#FFFFFF'>
			<td width="30%"><b>User Definable Fields</b></td>
			<td width="75%" colspan=2>Custom named fields, provides a textbox for the user to fill in when purchasing item.</td>
			</tr>

			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">User Field 1</td>
			<td width="70%" class="inputvalue">Use
						<input class="image" type="checkbox" <%= checked_U_d_1 %> name="U_d_1" value="-1" >
						Name
						<input	name="U_d_1_name" value="<%= U_d_1_name %>" size="60" maxlength=250>
						<input type="hidden" name="U_d_1_C" value="Op|String|0|250|||User Field 1">
						<% small_help "User Field" %></td>
			</tr>

			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">User Field 2</td>
			<td width="70%" class="inputvalue">Use
						<input class="image" type="checkbox" <%= checked_U_d_2 %> name="U_d_2" value="-1" > 
						Name
						<input	name="U_d_2_name" value="<%= U_d_2_name %>" size="60" maxlength=250>
						<input type="hidden" name="U_d_2_C" value="Op|String|0|250|||User Field 2">
						<% small_help "User Field" %></td>
			</tr>
		
				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">User Field 3</td>
			<td width="70%" class="inputvalue">Use
						<input class="image" type="checkbox" <%= checked_U_d_3 %> name="U_d_3" value="-1" > 
						Name
						<input	name="U_d_3_name" value="<%= U_d_3_name %>" size="60" maxlength=250>
						<input type="hidden" name="U_d_3_C" value="Op|String|0|250|||User Field 3">
						<% small_help "User Field" %></td>
			</tr>

				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">User Field 4</td>
			<td width="70%" class="inputvalue">Use
						<input class="image" type="checkbox" <%= checked_U_d_4 %> name="U_d_4" value="-1" >
						Name
						<input	name="U_d_4_name" value="<%= U_d_4_name %>" size="60" maxlength=250>
						<input type="hidden" name="U_d_4_C" value="Op|String|0|250|||User Field 4">
						<% small_help "User Field" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">User Field 5</td>
			<td width="70%" class="inputvalue">Use
						<input class="image" type="checkbox" <%= checked_U_d_5 %> name="U_d_5" value="-1" >
						Name
						<input	name="U_d_5_name" value="<%= U_d_5_name %>" size="60" maxlength=250>
						<input type="hidden" name="U_d_5_C" value="Op|String|0|250|||User Field 5">
						<% small_help "User Field" %></td>
			</tr>
			</table>
			</div>


			<!-- -----------------------------------------Extended Fields Start --------------------------------------------------------------------------------- -->
			<div id="content10" style="visibility: hidden;" class="tabPage">
				<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
			<tr bgcolor='#FFFFFF'>
				<td colspan=2>Extended fields, are displayed in addition to other item details</td>
			</tr>

			<tr bgcolor='#FFFFFF'>
				<td width="100%" class="inputname">Extended Field 1<BR>
					<%= create_editor_button ("M_d_1",M_d_1,"") %>
					<% small_help "Extended Field 1" %>
				</td>
			</tr>
			
			<tr bgcolor='#FFFFFF'>
				<td width="100%" class="inputname">Extended Field 2<BR>
					<%= create_editor_button ("M_d_2",M_d_2,"") %>
                                        <% small_help "Extended Field 2" %>
				</td>
			</tr>
			

			<tr bgcolor='#FFFFFF'>
				<td width="100%" class="inputname">Extended Field 3<BR>
					<%= create_editor_button ("M_d_3",M_d_3,"") %>
                                        <% small_help "Extended Field 3" %>
				</td>
			</tr>
			

			<tr bgcolor='#FFFFFF'>
				<td width="100%" class="inputname">Extended Field 4<BR>
					<%= create_editor_button ("M_d_4",M_d_4,"") %>
                                        <% small_help "Extended Field 4" %>
				</td>
			</tr>
			
			<tr bgcolor='#FFFFFF'>
				<td width="100%" class="inputname">Extended Field 5<BR>
					<%= create_editor_button ("M_d_5",M_d_5,"") %>
                                        <% small_help "Extended Field 5" %>
				</td>
			</tr>
			</table>
			</div>
			<!-- -----------------------------------------Extended Fields End --------------------------------------------------------------------------------- -->
			<% end if %>
			
			
			<% if Service_Type > 0 then %>
			<div id="content11" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
			<tr bgcolor='#FFFFFF'>
			<td width="30%"><b>3rd Party HTML</b></td>
			<td width="75%" colspan=2>Code which can be placed on 3rd party sites to add this item to the shoppers cart</td>
			</tr>
			<tr bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">Custom Link</td>
					<td width="70%" class="inputvalue"><input type="text" name="Custom_Link" value="<%= Custom_Link %>" size=60>&nbsp;
					<input type="hidden" name="Custom_Link_C" value="Op|String|0|255|||Custom Link"><% small_help "Custom Link" %></td>
			</tr>
				<tr bgcolor='#FFFFFF'>
			<td class="inputname" colspan=2>Cart Code for 3rd Party Sites<BR>
			<input type="button" class="Buttons" value="View 3rd Party Code" name="Add_Department" OnClick='JavaScript:window.open("<%= Site_Name %>include/third_party_code.asp?Item_Id=<%= Item_Id%>")'><% small_help "User Field" %></td>
			</tr>
			</table>
			</div>
			<% else %>
			<div id="content11" style="visibility: hidden;" class="tabPage">
				<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
				<tr bgcolor='#FFFFFF'>
				<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strBronze%><%=strMsg%></td>
				</tr>
				</table>
			</div>
			<% end if %>
			
			
			<!-- CONTENT 12 -->
			<div id="content12" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
			<tr bgcolor='#FFFFFF'>
				<td width="100%" colspan='3'>
				<% if Item_Id<>0 then %>

					<br>
					<input type="button" class="Buttons" value="Attributes (<%= sCountAttr %>)" name="Attributes" OnClick='JavaScript:self.location="Item_Attributes.asp?Item_Id=<%= Item_Id %>&Sub_Department_Id=<%= Sub_Department_Id %>&<%= sAddString %>"'>
					<input type="button" class="Buttons" value="Configurations (<%= sCountConfig %>)" name="Configurations" OnClick='JavaScript:self.location="Item_Configs.asp?Item_Id=<%= Item_Id %>&<%= sAddString %>"'>
					<input type="button" class="Buttons" value="Accessories (<%= sCountAccessory %>)" name="Accessories" OnClick='JavaScript:self.location="Item_Accessories.asp?Item_Id=<%= Item_Id %>&Sub_Department_Id=<%= Sub_Department_Id %>&<%= sAddString %>"'>
					<input type="button" class="Buttons" value="Price Group (<%= sCountPriceGroup %>)" name="Price_Group" OnClick=JavaScript:self.location="Price_Group_list.asp?Item_Id=<%= Item_Id %>&Sub_Department_Id=<%= Sub_Department_Id %>&<%= sAddString %>">
               <input type="button" class="Buttons" value="Price Matrix (<%= sCountPriceMatrix %>)" name="Price_Matrix" OnClick='JavaScript:self.location="Price_Matrix_list.asp?Item_Id=<%= Item_Id %>&Sub_Department_Id=<%= Sub_Department_Id %>&<%= sAddString %>"'>
					<br>
				<% else %>
				The Advanced functionality buttons are not shown until after the item is saved intially.
                                <% end if %>
				</td>
				</tr>
			</table>
			</div>
			<!-- CONTENT 12 -->
		</td>
		</tr></table></td></tr>				
		
<tr bgcolor='#FFFFFF'>
<td colspan='4' class='tpage'>
<% createFoot thisRedirect, 1
function checkStringForQBack(theString)
	tmpResult = theString
	tmpResult = replace(tmpResult,"&#8243;","&Prime;")
	tmpResult = replace(tmpResult,"&#8242;","&prime;")
	checkStringForQBack = tmpResult
end function
%>

</td>
</tr>

<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 
</script>
<% end if %>




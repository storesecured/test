<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/department_list.asp"-->
<!--#include file="include/location_list.asp"-->
<!--#include file="editor_include.asp"-->
<!--#include virtual="common/common_functions.asp"-->
<!--#include file="help/item_basic_edit.asp"-->


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

	Item_Name = checkStringForQBack(Rs_store("Item_Name"))
	Item_Page_Name = checkStringForQBack(Rs_store("Item_Page_Name"))
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
    Show_homepage=0
    Use_Price_By_Matrix=0
    Taxable=-1
    Fractional=0
    Department_Id=0
    Retail_Price=0
    Item_Weight=0
    Retail_Price_special_Discount=0
    Shipping_Fee=0
    Show=-1
    Special_start_date = Start_Date
	Special_end_date = End_Date
	Ship_Location_Id=0
	sDept_ids=""
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

addPicker = 1
sFormAction = "Inventory_action.asp"
sName = "Item_Basic_Edit"
sFormName = "Item_Basic_Edit"
sTrialAdd=1
if Item_Id=0 then
    sTitle = "Add Item"
    sFullTitle = "Inventory > <a href=edit_items.asp class=white>Items</a> > Add"
    sql_select = "Select Count(department_id) as Num_Depts from Store_Dept WITH (NOLOCK) where Store_Id=" & Store_Id&" and department_id<>0"

     rs_Store.open sql_select,conn_store,1,1
     if rs_store.bof = false then
	  Depts = rs_Store("Num_Depts")
     end if

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
    Depts=1
    sTitle = "Edit Item - "&Item_Name
    sFullTitle = "Inventory > <a href=edit_items.asp class=white>Items</a> > Edit - "&Item_Name
end if

sNeedTabs=1
sSubmitName = "Item_Edit"
sCommonName="Item"
sCancel="edit_items.asp"
thisRedirect = "item_edit.asp"
sMenu="inventory"
createHead thisRedirect

if sTrialAdd = 0 then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		You may not add any more items, your level of service is limited to <%= sLimit %>, your store currently has <%= Items %> items.
		<BR><BR>
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0 %>
<% elseif Depts=0 then %>
        <tr bgcolor='#FFFFFF'>
	<td colspan=2>
                You must create at least 1 department before adding inventory items.
		<BR><BR>
                <input type="button" class="Buttons" value="Add Department" name="Add_Department" OnClick='JavaScript:self.location="store_dept_basic.asp"'>


	</td></tr>
	<% createFoot thisRedirect, 0 %>
<% else %>
	<tr bgcolor='#FFFFFF'>
		<td width="100%" height="22" colspan=3>
		<input type="hidden" value="<%= Item_Id %>" name="Item_Id">
		
		<input type="button" class="Buttons" value="Item Search" name="Back_To_Item_List" OnClick='JavaScript:self.location="edit_Items.asp?<%= sAddString %>"'>
                <input type="button" class="Buttons" value="Add Department" name="Add_Department" OnClick='JavaScript:self.location="store_dept_add.asp"'>
                <% if Item_Id<>0 then %>
                <input class=buttons type="button" value="Preview Item" name="Editor" OnClick="javascript:goPreview('<%= fn_item_url("Preview",item_page_name) %>')">
			 <BR><input type="button" class="Buttons" value="Add Item" name="Add" OnClick='JavaScript:self.location="Item_basic_edit.asp?<%= sAddString %>"'>
			<input type="button" class="Buttons" value="Copy Item" name="Copy" OnClick='JavaScript:self.location="Item_copy.asp?Sub_Department_Id=<%= Sub_Department_Id %>&Item_Id=<%= Item_Id%>&Item_Name=<%= server.urlencode(Item_Name) %>&<%= sAddString %>"'>
                        <input type="button" class="Buttons" value="Advanced Item Edit" name="Add_Department" OnClick='JavaScript:self.location="item_edit.asp?Item_Id=<%= Item_Id%>"'>
                <% end if %>
		</td>
	 </tr>

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
							<input name="Item_Name" value="<%= Item_Name %>" size="60">
							<input type="hidden" name="Item_Name_C" value="Re|String|0|250|||Name">
							<% small_help "Name" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Page Filename</td>
				<td width="70%" class="inputvalue" nowrap>
							<input name="Item_Page_Name" value="<%= Item_Page_Name %>" maxlength=100 size="60" onKeyPress="return goodchars(event,'0123456789-abcdefghijklmnopqrstuvwxyz_')">-detail.htm
							<input type="hidden" name="Item_Page_Name_C" value="Op|String|0|100||@,$,%, ,',&,.,/,(,),`,;,#,!,?,^,®,™,†,©,”,“,½,¾,é,…,à,è,’|Page Filename">
							<% small_help "Item_Filename" %></td>
				</tr>

					<tr bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Retail Price</td>
				<td width="70%" class="inputvalue">
							<%= Store_Currency%><input name="Retail_Price" value="<%= Retail_Price %>" size="10" onKeyPress="return goodchars(event,'0123456789.-')" maxlength=10>
							<INPUT type="hidden"  name="Retail_Price_C" value="Re|Integer|||||Retail Price">
							<% small_help "Retail Price" %></td>
				</tr>
		
		               <tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname" colspan=2>Large Description<BR>

						<%= create_editor ("Description_L",Description_L,"") %>
						<% small_help "Large Description" %></td>
			</tr>
	  
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
			<td width="30%" class="inputname">Large Image</td>
			<td width="70%" class="inputvalue">
						<input type="text" name="ImageL_Path" value="<%= ImageL_Path %>" size="60" >
						<input type="hidden" name="ImageL_Path_C" value="Op|String|0|100|||Large Image">
						<font size="1" color="#000080"><a class="link" href="JavaScript:goImagePicker('ImageL_Path');"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
				<a class="link" href="JavaScript:goFileUploader('ImageL_Path');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>

				<% small_help "Large Image" %></td>
			</tr>
	  
				
		
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
frmvalidator.addValidation("Item_Name","req","Please enter a item name.");
frmvalidator.addValidation("Item_Sku","req","Please enter a item sku.");
frmvalidator.addValidation("Retail_Price","req","Please enter a retail price.");

</script>
<% end if %>




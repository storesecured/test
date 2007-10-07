<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/department_list.asp"-->
<!--#include file="include/location_list.asp"-->
<!--#include virtual="common/common_functions.asp"-->
<!--#include file="help/item_quick_edit.asp"-->
<%

'SELECT ITEM TO QUICK EDIT
Calendar=1
sFormAction = "item_quick_edit_action.asp"
sTitle = "Quick Edit Items"
sSubmitName = "Quick_Edit_Update"
thisRedirect = "item_quick_edit.asp"
sFormName = "Quick_Edit_Update"
addPicker =1
sMenu = "inventory"
createHead thisRedirect
if Service_Type < 7 then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.
	</td></tr>

<% createFoot thisRedirect, 0%>
<% else

	select_str="SELECT * FROM Store_Items WHERE Store_id="&Store_id&" AND  item_id in ("&request.form("DELETE_IDS")&")" 

	set itemfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,select_str,itemdata,itemfields,noRecords)

%>

<input type="hidden" name="Display_Items" value="Display_Items">
<input type="hidden" name="F_NAME" value="<%= request.form("F_NAME") %>">
<input type="hidden" name="F_SKU" value="<%= request.form("F_SKU") %>">
<input type="hidden" name="F_DATE" value="<%= request.form("F_DATE") %>">
<input type="hidden" name="F_KEYWORD" value="<%= request.form("F_KEYWORD") %>">
<input type="hidden" name="Sub_Department_Id" value="<%= Request.Form("Sub_Department_Id") %>">
<input type="hidden" name="F_LIVE" value="<%= request.form("F_LIVE") %>">
<input type="hidden" name="DELETE_IDS" value="<%= request.form("DELETE_IDS") %>">

				<% if noRecords=0 then %>
				<% FOR itemrowcounter= 0 TO itemfields("rowcount")
					Item_ID = itemdata(itemfields("item_id"),itemrowcounter) %>
					<TR bgcolor='#FFFFFF'>
						<td colspan="3"><b><br>Item #<%= Item_ID %> (<%= itemdata(itemfields("item_sku"),itemrowcounter) %> - <%= itemdata(itemfields("item_name"),itemrowcounter) %>)<br>&nbsp;</b></td>
					</tr>

					<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Departments</td>
				<td width="80%" class="inputvalue">
					<% sDept_Id=fn_get_dept_ids (Item_ID) %>
					<%= create_dept_list ("Department_ID_"&Item_ID,sDept_Id,5,"")
					%>
					<input type="hidden" name="Old_Department_Id_<%= Item_ID %>" value="<%= sDept_Id %>">
					
						<% small_help "Department" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">SKU</td>
					<td width="80%" class="inputvalue">
								<input name="Item_Sku_<%= Item_ID %>" size="30" value="<%= itemdata(itemfields("item_sku"),itemrowcounter) %>">
								<input type="hidden" name="Item_Sku_<%= Item_ID %>_C" value="Re|String|0|200|||SKU">
								<% small_help "SKU" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Name</td>
					<td width="80%" class="inputvalue">
								<input name="Item_Name_<%= Item_ID %>" size="50" value="<%= itemdata(itemfields("item_name"),itemrowcounter) %>">
								<input type="hidden" name="Item_Name_<%= Item_ID %>_C" value="Re|String|0|250|||Name">
								<% small_help "Name" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Page Filename</td>
					<td width="80%" class="inputvalue">
								<input name="Item_Page_Name_<%= Item_ID %>" size="50" value="<%= itemdata(itemfields("item_page_name"),itemrowcounter) %>">
								<input type="hidden" name="Item_Page_Name_<%= Item_ID %>_C" value="Op|String|0|100||@,$,%, ,',&,.,/,(,),`,;,#,!,?,^,®,™,†,©,”,“,½,¾,é,…,à,è,’|Filename">
								<% small_help "Filename" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Retail Price</td>
					<td width="80%" class="inputvalue">
								<%= Store_Currency%><input name="Retail_Price_<%= Item_ID %>" size="10" value="<%= itemdata(itemfields("retail_price"),itemrowcounter) %>" onKeyPress="return goodchars(event,'-0123456789.')" >
								<INPUT type="hidden"  name="Retail_Price_<%= Item_ID %>_C" value="Re|Integer|||||Retail Price">
								<% small_help "Retail Price" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Cost</td>
					<td width="80%" class="inputvalue">
						<%= Store_Currency%><input	name="Wholesale_Price_<%= Item_ID %>" value="<%= itemdata(itemfields("wholesale_price"),itemrowcounter) %>" size="10" onKeyPress="return goodchars(event,'-0123456789.')" maxlength=10>
						<INPUT type="hidden" name="Wholesale_Price_<%= Item_ID %>_C" value="Op|Integer|||||Cost Price">
								<% small_help "Wholesale Price" %></td>
				</tr>




				 <% if Service_Type < 7 then %>
					<input type="hidden" name="Use_Price_By_Matrix_<%= Item_ID %>" value="0" >
				<% else %>
					<TR bgcolor='#FFFFFF'>
						<td width="20%" class="inputname">Qty Discount</td>
						<td width="80%" class="inputvalue">
									<input class="image" type="checkbox" name="Use_Price_By_Matrix_<%= Item_ID %>" value="-1" 
									<% If cint(itemdata(itemfields("use_price_by_matrix"),itemrowcounter))=-1 then %>
									checked
									<% End If %>
									>
									<% small_help "Qty Discount" %></td>
					</tr>
					<% end if %>



					<TR bgcolor='#FFFFFF'>
						<td width="20%" class="inputname">Homepage</td>
						<td width="80%" class="inputvalue">
									<input class="image" type="checkbox"  name="Show_Homepage_<%= Item_ID %>" value="-1" 
									<% If itemdata(itemfields("show_homepage"),itemrowcounter)=-1 then %>
											checked
										<% End If %>
										>
									<% small_help "Homepage" %>
                           </td>
					</tr>
				

				<% if Service_Type => 5 then %>
					<TR bgcolor='#FFFFFF'>
						<td width="20%" class="inputname">Taxable</td>
						<td width="80%" class="inputvalue">
									<input class="image" type="checkbox" name="Taxable_<%= Item_ID %>" value="-1"
										<% If itemdata(itemfields("taxable"),itemrowcounter)=-1 then %>
											checked
										<% End If %>
										><% small_help "Taxable" %></td>
					</tr>
	  
					<TR bgcolor='#FFFFFF'>
						<td width="20%" class="inputname">Visible</td>
						<td width="80%" class="inputvalue">
									<input class="image" type="checkbox" name="Show_<%= Item_ID %>" value="-1"
										<% If itemdata(itemfields("show"),itemrowcounter)=-1 then %>
											checked
										<% End If %>
										><% small_help "Visible" %></td>
					</tr>
					<TR bgcolor='#FFFFFF'>
						<td width="20%" class="inputname">Hide Price</td>
						<td width="80%" class="inputvalue">
									<input class="image" type="checkbox"  name="Hide_Price_<%= Item_ID %>" value="-1" 
									<% If itemdata(itemfields("hide_price"),itemrowcounter)=-1 then %>
											checked
										<% End If %>
										>
									<% small_help "Hide Price" %></td>
					</tr>
					<TR bgcolor='#FFFFFF'>
						<td width="20%" class="inputname">Custom Pricing</td>
						<td width="80%" class="inputvalue">
									<input class="image" type="checkbox"  name="Cust_Price_<%= Item_ID %>" value="-1" 
									<% If itemdata(itemfields("enable_cust_price"),itemrowcounter)=-1 then %>
											checked
										<% End If %>
										>
									<% small_help "Cust Price" %></td>
					</tr>
					<TR bgcolor='#FFFFFF'>
						<td width="20%" class="inputname">View Order</td>
						<td width="80%" class="inputvalue">
							<input	name="View_Order_<%= Item_ID %>"" value="<%= itemdata(itemfields("view_order"),itemrowcounter) %>" size="10" onKeyPress="return goodchars(event,'-0123456789.')">
							<INPUT type="hidden" name="View_Order_<%= Item_ID %>_C" value="Re|Integer|||||View_Order">
							<% small_help "View Order" %></td>
					</tr>
				<% else %>
					<input type="hidden" name="Taxable_<%= Item_ID %>" value="1" >
					<input type="hidden" name="Show_<%= Item_ID %>" value="1" >
					<input type="hidden" name="Hide_Price_<%= Item_ID %>" value="0" >
					<input type="hidden" name="Cust_Price_<%= Item_ID %>" value="0" >
					<input type="hidden" name="View_Order_<%= Item_ID %>" value="0" >
			<% end if %>

				<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Short Description </td>
					<td width="80%" class="inputvalue">
					<% editor_link "Description_S_"&Item_ID, "4000" %>
					  <input readonly type=text name=remLenDesS_<%= Item_ID %> size=3 class=char maxlength=3 value="<%= 4000-len(itemdata(itemfields("description_s"),itemrowcounter)) %>" class=image><font size=1><I>characters left</i></font>
					<BR>
								<textarea cols="55" name="Description_S_<%= Item_ID %>" rows="4" onfocus="textCounter(this.form.Description_S_<%= Item_ID %>,this.form.remLenDesS_<%= Item_ID %>,4000);" onKeyDown="textCounter(this.form.Description_S_<%= Item_ID %>,this.form.remLenDesS_<%= Item_ID %>,4000);" onKeyUp="textCounter(this.form.Description_S_<%= Item_ID %>,this.form.remLenDesS_<%= Item_ID %>,4000);"><%= itemdata(itemfields("description_s"),itemrowcounter) %></textarea>
								<input type="hidden" name="Description_S_<%= Item_ID %>_C" value="Op|String|0|4000|||Short Description">
								<% small_help "Short Description" %></td>
				</tr>
	  
				<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Small Image</td>
					<td width="80%" class="inputvalue">
								<input type="text" name="ImageS_Path_<%= Item_ID %>" value="<%= itemdata(itemfields("images_path"),itemrowcounter) %>" size="20" >
								<input type="hidden" name="ImageS_Path_C_<%= Item_ID %>" value="Op|String|0|100|||Small Image">
								<font size="1" color="#000080"><a class="link" href="JavaScript:goImagePicker('ImageS_Path_<%= Item_ID %>');"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
								<a class="link" href="JavaScript:goFileUploader('ImageS_Path_<%= Item_ID %>');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
							<% small_help "Small Image" %></td>
				</tr>



				<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Large Description</td>
					<td width="80%" class="inputvalue">
					<% editor_link "Description_L_"&Item_ID, "8000" %>
				  <BR>
								<textarea cols="55" name="Description_L_<%= Item_ID %>"	rows="6"><%= itemdata(itemfields("description_l"),itemrowcounter) %></textarea>
								<% small_help "Large Description" %></td>
				</tr>
	  
				<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Large Image</td>
					<td width="80%" class="inputvalue">
						<input type="text" name="ImageL_Path_<%= Item_ID %>" value="<%= itemdata(itemfields("imagel_path"),itemrowcounter) %>" size="20" >
						<input type="hidden" name="ImageL_Path_C_<%= Item_ID %>" value="Op|String|0|100|||Large Image">
						<font size="1" color="#000080"><a class="link" href="JavaScript:goImagePicker('ImageL_Path_<%= Item_ID %>');"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
						<a class="link" href="JavaScript:goFileUploader('ImageL_Path_<%= Item_ID %>');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
						<% small_help "Large Image" %></td>
				</tr>



		<% if instr(Shipping_Classes,"2") or instr(Shipping_Classes,"6") or instr(Shipping_Classes,"7") then %>
			  <TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Shipping Weight</td>
					<td width="80%" class="inputvalue">
						<input	name="Item_Weight_<%= Item_ID %>" value="<%=  itemdata(itemfields("item_weight"),itemrowcounter) %>" size="8" onKeyPress="return goodchars(event,'-0123456789.')">
						<INPUT type="hidden"  name="Item_Weight_<%= Item_ID %>_C" value="Re|Integer|||||Weight">
						<% small_help "Weight" %></td>
			</tr>

		<% else %>
				<input type=hidden name="Item_Weight_<%= Item_ID %>" value="<%=  itemdata(itemfields("item_weight"),itemrowcounter)  %>">
		<% end if %>


			<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Handling Fee</td>
					<td width="80%" class="inputvalue">
						<input	name="Item_Handling_<%= Item_ID %>" value="<%=itemdata(itemfields("item_handling"),itemrowcounter) %>" size="8" onKeyPress="return goodchars(event,'-0123456789.')">
						<INPUT type="hidden"  name="Item_Handling_<%= Item_ID %>_C" value="Re|Integer|||||Handling">
					<% small_help "Handling" %></td>
			</tr>


			<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Waive Shipping</td>
					<td width="80%" class="inputvalue">
					<input class="image" type="checkbox" name="Waive_Shipping_<%= Item_ID %>" value="-1"
							<% If cint(itemdata(itemfields("waive_shipping"),itemrowcounter))=-1 then %>
								checked
							<% End If %>
							><% small_help "Waive Shipping" %>
						</td>
			</tr>


				<% if Service_Type => 7 then %>

							
							<TR bgcolor='#FFFFFF'>
								<td width="20%" class="inputname">Shipping From</td>
								<td width="80%" class="inputvalue">
									<%= create_location_list ("Ship_Location_Id_"&Item_ID,itemdata(itemfields("ship_location_id"),itemrowcounter),1) %>
										<% small_help "Shipfrom" %>
									</td>
							</tr>
				

			<% else %>
			     <input type="hidden" name="Ship_Location_Id_<%= Item_ID %>" value="0">
			<% end if %>
			
			<% if instr(Shipping_Classes,"3") > 0 then %>

			  <TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Shipping Fee</td>
					<td width="80%" class="inputvalue">
								<%= Store_Currency%><input name="Shipping_Fee_<%= Item_ID %>" value="<%= itemdata(itemfields("shipping_fee"),itemrowcounter) %>" size="10" onKeyPress="return goodchars(event,'-0123456789.')" maxlength=10>
								<INPUT type="hidden"  name="Shipping_Fee_<%= Item_ID %>_C" value="Re|Integer|||||Shipping Fee">
						<% small_help "Shipping Fee" %></td>
			</tr>
			<% else %>
				<input type=hidden name="Shipping_Fee_<%= Item_ID %>" value="<%= itemdata(itemfields("shipping_fee"),itemrowcounter) %>">
			<% end if %>




		<% if Service_Type < 5 then %>
				<input type=hidden name="Quantity_in_stock_<%= Item_ID %>" Value="<%= itemdata(itemfields("quantity_in_stock"),itemrowcounter) %>" size="8">
				<input type="hidden" name="Quantity_Control_<%= Item_ID %>" value="0" >
				<input type="hidden" name="Hide_Stock_<%= Item_ID %>" value="0" >
				<input type="hidden" name="Quantity_Control_Number_<%= Item_ID %>" value="<%= itemdata(itemfields("quantity_control_number"),itemrowcounter) %>" size="4">
				<input type="hidden" name="Quantity_Minimum_<%= Item_ID %>" value="<%=itemdata(itemfields("quantity_minimum"),itemrowcounter)  %>" size="4">
				<input type="hidden" name="Fractional_<%= Item_ID %>" value="0">
				<input type=hidden	name="Meta_Keywords_<%= Item_ID %>" value="<%= itemdata(itemfields("meta_keywords"),itemrowcounter)  %>" size="30">
				<input type=hidden	name="Meta_Description_<%= Item_ID %>" value="<%= itemdata(itemfields("meta_description"),itemrowcounter) %>" size="30">
			  <input type=hidden	name="Meta_Title_<%= Item_ID %>" value="<%=itemdata(itemfields("meta_title"),itemrowcounter)  %>" size="30">

		<% else %>
			<TR bgcolor='#FFFFFF'>
					<td width="30%"><b>Inventory Management</b></td>
					<td width="75%" colspan=2>&nbsp;</td>
			</tr>
			 <TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Stock</td>
				<td width="80%" class="inputvalue">
						<input	name="Quantity_in_stock_<%= Item_ID %>" Value="<%=itemdata(itemfields("quantity_in_stock"),itemrowcounter)  %>" size="8" onKeyPress="return goodchars(event,'-0123456789.')">
						<INPUT type="hidden"  name="Quantity_in_stock_<%= Item_ID %>_C" value="Re|Integer|||||Stock">
						<% small_help "Stock" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Qty Control</td>
				<td width="80%" class="inputvalue">
					<input class="image" type="checkbox"  name="Quantity_Control_<%= Item_ID %>" value="-1"
					<% If cint(itemdata(itemfields("quantity_control"),itemrowcounter))=-1 then %>
								checked
							<% End If %>
							>
						<% small_help "Qty Control" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Control #</td>
				<td width="80%" class="inputvalue">
							<input name="Quantity_Control_Number_<%= Item_ID %>" value="<%= itemdata(itemfields("quantity_control_number"),itemrowcounter)  %>" size="4" onKeyPress="return goodchars(event,'-0123456789.')">
							<INPUT type="hidden"  name="Quantity_Control_Number_<%= Item_ID %>_C" value="Re|Integer|||||Quantity Control">
							<% small_help "Control #" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Hide Stock</td>
				<td width="80%" class="inputvalue">
					<input class="image" type="checkbox"  name="Hide_Stock_<%= Item_ID %>" value="-1" 
					<% If cint(itemdata(itemfields("hide_stock"),itemrowcounter))=-1 then %>
							checked
						<% End If %>
						>
					<% small_help "Control #" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Min Quantity</td>
				<td width="80%" class="inputvalue">
								 <input name="Quantity_Minimum_<%= Item_ID %>" value="<%= itemdata(itemfields("quantity_minimum"),itemrowcounter) %>" size="4" onKeyPress="return goodchars(event,'-0123456789.')">
							<INPUT type="hidden"  name="Quantity_Minimum_<%= Item_ID %>_C" value="Re|Integer|||||Min Quantity">
							<% small_help "Min Quantity" %></td>
			</tr>
			

			 <TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Fractional</td>
				<td width="80%" class="inputvalue">
							<input class="image" name="Fractional_<%= Item_ID %>" type="checkbox" value="-1"
								<% If cint(itemdata(itemfields("fractional"),itemrowcounter))=-1 then %>
									checked
								<% End If %>
								><% small_help "Fractional" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>Search Engine Tags</b></td>
				<td width="75%" colspan=2>Help get your items listed in search engines</td>
			</tr>
			
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Title</td>
				<td width="80%" class="inputvalue">
						<input type=text	name="Meta_Title_<%= Item_ID %>" value="<%= itemdata(itemfields("meta_title"),itemrowcounter) %>" maxlength=100 size=30>
						<input type="hidden" name="Meta_Title_<%= Item_ID %>_C" value="Op|String|0|100|||Meta Title">
						<% small_help "Meta_Title" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Keywords</td>
				<td width="80%" class="inputvalue">
					<input readonly type=text name=remLenMKey_<%= Item_ID %> size=3 class=char maxlength=3 value="<%= 250-len(itemdata(itemfields("meta_keywords"),itemrowcounter)) %>" class=image><font size=1><I>characters left</i></font>
					<textarea	 name="Meta_Keywords_<%= Item_ID %>" cols=55 rows=3 onKeyDown="textCounter(this.form.Meta_Keywords_<%= Item_ID %>,this.form.remLenMKey_<%= Item_ID %>,250);" onKeyUp="textCounter(this.form.Meta_Keywords_<%= Item_ID %>,this.form.remLenMKey_<%= Item_ID %>,250);"><%= itemdata(itemfields("meta_keywords"),itemrowcounter) %></textarea>
							<input type="hidden" name="Meta_Keywords_<%= Item_ID %>_C" value="Op|String|0|250|||Meta Keywords">
							<% small_help "Meta_Keywords" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Description</td>
				<td width="80%" class="inputvalue">
					<input readonly type=text name=remLenMDes_<%= Item_ID %> size=3 class=char maxlength=3 value="<%= 500-len(itemdata(itemfields("meta_description"),itemrowcounter)) %>" class=image><font size=1><I>characters left</i></font>
					<textarea	name="Meta_Description_<%= Item_ID %>" cols=55 rows=5 onKeyDown="textCounter(this.form.Meta_Description_<%= Item_ID %>,this.form.remLenMDes_<%= Item_ID %>,500);" onKeyUp="textCounter(this.form.Meta_Description_<%= Item_ID %>,this.form.remLenMDes_<%= Item_ID %>,500);"><%= itemdata(itemfields("meta_description"),itemrowcounter) %></textarea>
					<input type="hidden" name="Meta_Description_<%= Item_ID %>_C" value="Op|String|0|500|||Meta Description">
				<% small_help "Meta_Description" %></td>
			</tr>
			<% end if %>

	  
			<% if Service_Type < 5 then %>
				<input type=hidden name="Brand_<%= Item_ID %>" value="<%=itemdata(itemfields("brand"),itemrowcounter)  %>">
				<input type=hidden name="Condition_<%= Item_ID %>" value="<%=itemdata(itemfields("condition"),itemrowcounter)  %>">
				<input type=hidden name="Product_Type_<%= Item_ID %>" value="<%=itemdata(itemfields("product_type"),itemrowcounter)  %>">
				<input type=hidden name="Recurring_days_<%= Item_ID %>" value="<%=itemdata(itemfields("recurring_days"),itemrowcounter)  %>">
				<input type=hidden name="Recurring_fee_<%= Item_ID %>" value="<%= itemdata(itemfields("recurring_fee"),itemrowcounter) %>">
				<input type=hidden name="Retail_Price_special_Discount_<%= Item_ID %>" value="<%= itemdata(itemfields("retail_price_special_discount"),itemrowcounter) %>">
				<input type=hidden name="Special_Start_Date_<%= Item_ID %>" value="<%= FormatDateTime(itemdata(itemfields("special_start_date"),itemrowcounter),2) %>">
				<input type=hidden name="Special_End_Date_<%= Item_ID %>" value="<%=  FormatDateTime(itemdata(itemfields("special_end_date"),itemrowcounter),2) %>">
				<input type=hidden name="File_Location_<%= Item_ID %>" value="<%= itemdata(itemfields("file_location"),itemrowcounter) %>">
				<input type=hidden name="U_d_1_<%= Item_ID %>" value="0" >
				<input type=hidden name="U_d_1_name_<%= Item_ID %>" value="<%= itemdata(itemfields("u_d_1_name"),itemrowcounter) %>">
				<input type=hidden name="U_d_2_<%= Item_ID %>" value="0" >
				<input type=hidden name="U_d_2_name_<%= Item_ID %>" value="<%= itemdata(itemfields("u_d_2_name"),itemrowcounter) %>">
				<input type=hidden name="U_d_3_<%= Item_ID %>" value="0" >
				<input type=hidden name="U_d_3_name_<%= Item_ID %>" value="<%=  itemdata(itemfields("u_d_3_name"),itemrowcounter) %>">
				<input type=hidden name="U_d_4_<%= Item_ID %>" value="0" >
				<input type=hidden name="U_d_4_name_<%= Item_ID %>" value="<%=  itemdata(itemfields("u_d_4_name"),itemrowcounter) %>">
				<input type=hidden name="U_d_5_<%= Item_ID %>" value="0" >
				<input type=hidden name="U_d_5_name_<%= Item_ID %>" value="<%=  itemdata(itemfields("u_d_5_name"),itemrowcounter) %>">


				<input type=hidden name="M_d_1_<%= Item_ID %>" value="<%= itemdata(itemfields("m_d_1"),itemrowcounter) %>">
				<input type=hidden name="M_d_2_<%= Item_ID %>" value="<%= itemdata(itemfields("m_d_2"),itemrowcounter)  %>">
				<input type=hidden name="M_d_3_<%= Item_ID %>" value="<%= itemdata(itemfields("m_d_3"),itemrowcounter)  %>">
				<input type=hidden name="M_d_4_<%= Item_ID %>" value="<%= itemdata(itemfields("m_d_4"),itemrowcounter)  %>">
				<input type=hidden name="M_d_5_<%= Item_ID %>" value="<%= itemdata(itemfields("m_d_5"),itemrowcounter)  %>">

			<% else %>
			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>Recurring</b></td>
				<td width="75%" colspan=2>Note that some processors do not support recurring billing</font></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Recurring Fee</td>
				<td width="80%" class="inputvalue">
							<%= Store_Currency%><input name="Recurring_fee_<%= Item_ID %>" value="<%= itemdata(itemfields("recurring_fee"),itemrowcounter)  %>" size="5" value="0" onKeyPress="return goodchars(event,'-0123456789.')">
							<INPUT type="hidden"  name="Recurring_fee_<%= Item_ID %>_C" value="Op|Integer|||||Recurring Fee">
							<% small_help "Recurring_Fee" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Recurring Interval</td>
				<td width="80%" class="inputvalue">
							<input name="Recurring_days_<%= Item_ID %>" value="<%= itemdata(itemfields("recurring_days"),itemrowcounter) %>" size="5" onKeyPress="return goodchars(event,'0123456789')">days
							<INPUT type="hidden"  name="Recurring_days_<%= Item_ID %>_C" value="Op|Integer|||||Recurring Days">
							<% small_help "Recurring_Days" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>Google Base</b></td>
				<td width="75%" colspan=2></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Condition</td>
			<td width="70%" class="inputvalue">
				<select name="Condition_<%= Item_ID %>">
				<option value="<%= itemdata(itemfields("condition"),itemrowcounter) %>" selected><%= itemdata(itemfields("condition"),itemrowcounter) %></option>
				<option value="New">New</option>
				<option value="Used">Used</option>
				<option value="Refurbished">Refurbished</option></select>
						<% small_help "Condition" %></td>
			</tr>

				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Brand</td>
			<td width="70%" class="inputvalue">
						<input name="Brand_<%= Item_ID %>" value="<%= itemdata(itemfields("brand"),itemrowcounter) %>" size="60" maxlength=100>
						<INPUT type="hidden"  name="Brand_<%= Item_ID %>_C" value="Op|String|0|100|||Brand">
						<% small_help "Brand" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Product Type</td>
			<td width="70%" class="inputvalue">
						<input name="Product_Type_<%= Item_ID %>" value="<%= itemdata(itemfields("product_type"),itemrowcounter) %>" size="60" maxlength=20>
						<INPUT type="hidden"  name="Product_Type_<%= Item_ID %>_C" value="Op|String|0|20|||Product Type">
						<% small_help "Product_Type" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>Specials</b></td>
				<td width="75%" colspan=2>Not effective when quantity discounts are enabled.</font></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Special Price</td>
				<td width="80%" class="inputvalue">Discount
						<input name="Retail_Price_special_Discount_<%= Item_ID %>" value="<%= itemdata(itemfields("retail_price_special_discount"),itemrowcounter) %>" size="5" onKeyPress="return goodchars(event,'-0123456789.')">%
							<INPUT type="hidden"  name="Retail_Price_special_Discount_<%= Item_ID %>_C" value="Op|Integer|||||Discount">
							<% small_help "Special Price" %></td>
			</tr>

		 <SCRIPT LANGUAGE="JavaScript" ID="jscal1">
		      var now = new Date();
		      var cal1 = new CalendarPopup("testdiv1");
		      cal1.showNavigationDropdowns();
	      </SCRIPT>
	
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Dates</td>
				<td width="80%" class="inputvalue">Between
					<input name="Special_Start_Date_<%= Item_ID %>" value="<%= FormatDateTime(itemdata(itemfields("special_start_date"),itemrowcounter),2) %>" size="10" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')">
							<A HREF="#" onClick="cal1.select(document.forms[0].Special_Start_Date_<%= Item_ID %>,'anchor1_<%= Item_ID%>','M/d/yyyy',(document.forms[0].Special_Start_Date_<%= Item_ID %>.value=='')?document.forms[0].Special_Start_Date_<%= Item_ID %>.value:null); return false;" TITLE="Start Date" NAME="anchor1_<%= Item_ID%>" ID="anchor1_<%= Item_ID%>"><img src=images/calendar.gif border=0></A>
					 <INPUT type="hidden"  name="Special_Start_Date_<%= Item_ID %>_C" value="Re|date|||||Special Start Date"> and
							<input name="Special_End_Date_<%= Item_ID %>" value="<%=  FormatDateTime(itemdata(itemfields("special_end_date"),itemrowcounter),2) %>" size="10" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')">
							<A HREF="#" onClick="cal1.select(document.forms[0].Special_End_Date_<%= Item_ID %>,'anchor2_<%= Item_ID%>','M/d/yyyy',(document.forms[0].Special_End_Date_<%= Item_ID %>.value=='')?document.forms[0].Special_End_Date_<%= Item_ID %>.value:null); return false;" TITLE="End Date" NAME="anchor2_<%= Item_ID%>" ID="anchor2_<%= Item_ID%>"><img src=images/calendar.gif border=0></A>
				   <INPUT type="hidden"  name="Special_end_Date_<%= Item_ID %>_C" value="Re|date|||||Special End Date">
							<% small_help "Dates" %></td>
			</tr>

			
			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>Electronic Software Delivery</b></td>
				<td width="75%" colspan=2>Fill this out if your item is a software file, ie MP3, Game, etc.</td>
			</tr>
	  
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Filename</td>
				<td width="80%" class="inputvalue">
					<input type="text" name="File_Location_<%= Item_ID %>" value="<%= itemdata(itemfields("file_location"),itemrowcounter) %>" size="30" maxlength=250>
					<input type="hidden" name="File_Location_<%= Item_ID %>_C" value="Op|String|0|250|||Filename">
						<a class="link" href="javascript:goFilePicker('File_Location_<%= Item_ID %>')"><img border="0" src="images/image.gif" width="23" height="22" alt="File Picker"></a>
						<% small_help "Filename" %></td>
			</tr>


			<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Pin Delivery</td>
				<td width="70%" class="inputvalue">
					<input class="image" name="Item_pin_<%= Item_ID %>" type="checkbox" value="-1"
							<% If cint(itemdata(itemfields("item_pin"),itemrowcounter))=-1 then %>
								checked
							<% End If %>
							>
							<% small_help "Item_pin" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Item Remarks<BR><font size="1">Not seen by customers</font></td>
				<td width="80%" class="inputvalue">
							<input readonly type=text name=remLenRem_<%= Item_ID %> size=3 class=char maxlength=3 value="<%= 200-len(itemdata(itemfields("item_remarks"),itemrowcounter)) %>" class=image><font size=1><I>characters left</i></font>
				<textarea name="Item_Remarks_<%= Item_ID %>" cols="55" rows="2"><%= itemdata(itemfields("item_remarks"),itemrowcounter) %></textarea>
							<input type="hidden" name="Item_Remarks_<%= Item_ID %>_C" value="Op|String|0|200|||Item Remarks">
							<% small_help "Item Remarks" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>User Definable Fields</b></td>
				<td width="75%" colspan=2>Custom named fields, provides a textbox for the user to fill in when purchasing item.</td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">User Field 1</td>
				<td width="80%" class="inputvalue">Use
							<input class="image" type="checkbox"  name="U_d_1_<%= Item_ID %>" value="-1" 
							<% If cint(itemdata(itemfields("u_d_1"),itemrowcounter))=-1 then %>
									checked
								<% End If %>
								>
							
							Name
							<input	name="U_d_1_name_<%= Item_ID %>" value="<%= itemdata(itemfields("u_d_1_name"),itemrowcounter) %>" size="30" maxlength=250>
							<input type="hidden" name="U_d_1_<%= Item_ID %>_C" value="Op|String|0|250|||User Field 1">
							<% small_help "User Field" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">User Field 2</td>
				<td width="80%" class="inputvalue">Use
							<input class="image" type="checkbox"  name="U_d_2_<%= Item_ID %>" value="-1" 
							<% If cint(itemdata(itemfields("u_d_2"),itemrowcounter))=-1 then %>
									checked
								<% End If %>
								>
							

							Name
							<input	name="U_d_2_name_<%= Item_ID %>" value="<%= itemdata(itemfields("u_d_2_name"),itemrowcounter) %>" size="30" maxlength=250>
							<input type="hidden" name="U_d_2_<%= Item_ID %>_C" value="Op|String|0|250|||User Field 2">
							<% small_help "User Field" %></td>
			</tr>
		
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">User Field 3</td>
				<td width="80%" class="inputvalue">Use
							<input class="image" type="checkbox"  name="U_d_3_<%= Item_ID %>" value="-1" 
								<% If cint(itemdata(itemfields("u_d_3"),itemrowcounter))=-1 then %>
									checked
								<% End If %>
								>
							Name
							<input	name="U_d_3_name_<%= Item_ID %>" value="<%= itemdata(itemfields("u_d_3_name"),itemrowcounter) %>" size="30" maxlength=250>
							<input type="hidden" name="U_d_3_<%= Item_ID %>_C" value="Op|String|0|250|||User Field 3">
							<% small_help "User Field" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">User Field 4</td>
				<td width="80%" class="inputvalue">Use
							<input class="image" type="checkbox"  name="U_d_4_<%= Item_ID %>" value="-1" 
							<% If cint(itemdata(itemfields("u_d_4"),itemrowcounter))=-1 then %>
									checked
								<% End If %>
								>
							Name
							<input	name="U_d_4_name_<%= Item_ID %>" value="<%= itemdata(itemfields("u_d_4_name"),itemrowcounter) %>" size="30" maxlength=250>
							<input type="hidden" name="U_d_4_<%= Item_ID %>_C" value="Op|String|0|250|||User Field 4">
							<% small_help "User Field" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">User Field 5</td>
				<td width="80%" class="inputvalue">Use
							<input class="image" type="checkbox"  name="U_d_5_<%= Item_ID %>" value="-1" 
							<% If cint(itemdata(itemfields("u_d_5"),itemrowcounter))=-1 then %>
									checked
								<% End If %>
								>
							Name
							<input	name="U_d_5_name_<%= Item_ID %>" value="<%= itemdata(itemfields("u_d_5_name"),itemrowcounter) %>" size="30" maxlength=250>
							<input type="hidden" name="U_d_5_<%= Item_ID %>_C" value="Op|String|0|250|||User Field 5">
							<% small_help "User Field" %></td>
			</tr>



			<TR bgcolor='#FFFFFF'>
				<td width="20%"><b>Extended Fields</b></td>
				<td width="75%" colspan=2>Extended fields, are displayed in addition to other item details</td>
			</tr>
               <TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Extended Field 1</td>
					<td width="80%" class="inputvalue">
					<% editor_link "M_d_1_"&Item_ID, "8000" %>
				  <BR>
								<textarea cols="55" name="M_d_1_<%= Item_ID %>"	rows="6"><%= itemdata(itemfields("m_d_1"),itemrowcounter) %></textarea>
								<% small_help "Extended Field 1" %></td>
				</tr>
                <TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Extended Field 2</td>
					<td width="80%" class="inputvalue">
					<% editor_link "M_d_2_"&Item_ID, "8000" %>
				  <BR>
								<textarea cols="55" name="M_d_2_<%= Item_ID %>"	rows="6"><%= itemdata(itemfields("m_d_2"),itemrowcounter) %></textarea>
								<% small_help "Extended Field 2" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Extended Field 3</td>
					<td width="80%" class="inputvalue">
					<% editor_link "M_d_3_"&Item_ID, "8000" %>
				  <BR>
								<textarea cols="55" name="M_d_3_<%= Item_ID %>"	rows="6"><%= itemdata(itemfields("m_d_3"),itemrowcounter) %></textarea>
								<% small_help "Extended Field 3" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Extended Field 4</td>
					<td width="80%" class="inputvalue">
					<% editor_link "M_d_4_"&Item_ID, "8000" %>
				  <BR>
								<textarea cols="55" name="M_d_4_<%= Item_ID %>"	rows="6"><%= itemdata(itemfields("m_d_4"),itemrowcounter) %></textarea>
								<% small_help "Extended Field 4" %></td>
				</tr>

                    <TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Extended Field 5</td>
					<td width="80%" class="inputvalue">
					<% editor_link "M_d_5_"&Item_ID, "8000" %>
				  <BR>
								<textarea cols="55" name="M_d_5_<%= Item_ID %>"	rows="6"><%= itemdata(itemfields("m_d_5"),itemrowcounter) %></textarea>
								<% small_help "Extended Field 5" %></td>
				</tr>

			<!-- -----------------------------------------Extended Fields End --------------------------------------------------------------------------------- -->
			

				 <% end if %>


			<% if Service_Type > 0 then %>
			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>3rd Party HTML</b></td>
				<td width="75%" colspan=2>Code which can be placed on 3rd party sites to add this item to the shoppers cart</td>
			</tr>
			<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Custom Link</td>
					<td width="80%" class="inputvalue"><input type="text" name="Custom_Link_<%= Item_ID %>" value="<%=itemdata(itemfields("custom_link"),itemrowcounter)  %>" size=52>&nbsp;
					<input type="hidden" name="Custom_Link_<%= Item_ID %>_C" value="Op|String|0|255|||Custom Link"><% small_help "Custom Link" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Cart Code for 3rd Party Sites</td>
				<td width="80%" class="inputvalue"><input type="button" class="Buttons" value="3rd Party Code All Items" name="Add_Department" OnClick='JavaScript:window.open("<%= Site_Name %>/include/third_party_code.asp?Item_Id=<%= request.form("DELETE_IDS") %>")'>
				<input type="button" class="Buttons" value="3rd Party Code This Item" name="Add_Department" OnClick='JavaScript:window.open("<%= Site_Name %>/include/third_party_code.asp?Item_Id=<%= Item_Id %>")'>
				<% small_help "User Field" %></td>
			</tr>
			<% end if %>

				<% Next %>
				<% end if %>


<% createFoot thisRedirect, 1%>
<% end if 
%>
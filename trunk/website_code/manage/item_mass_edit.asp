<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/department_list.asp"-->
<!--#include file="include/location_list.asp"-->
<!--#include file="help/item_mass_edit.asp"-->

<%

calendar=1
sFormAction = "item_mass_edit_action.asp"
sTitle = "Mass Edit Items"
sSubmitName = "Mass_Edit_Update"
sFormName = "Mass_Edit_Update"
thisRedirect = "item_mass_edit.asp"
addPicker=1
sMenu = "inventory"
createHead thisRedirect


Special_start_date=FormatDateTime(now(),2)
Special_end_date=Dateadd("m",1,Special_start_date)

if Service_Type < 7 then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.
	</td></tr>

<% createFoot thisRedirect, 0%>
<% else %>

<input type="hidden" name="Display_Items" value="Display_Items">
<input type="hidden" name="F_NAME" value="<%= request.form("F_NAME") %>">
<input type="hidden" name="F_SKU" value="<%= request.form("F_SKU") %>">
<input type="hidden" name="F_DATE" value="<%= request.form("F_DATE") %>">
<input type="hidden" name="F_KEYWORD" value="<%= request.form("F_KEYWORD") %>">
<input type="hidden" name="Sub_Department_Id" value="<%= Request.Form("Sub_Department_Id") %>">
<input type="hidden" name="F_LIVE" value="<%= request.form("F_LIVE") %>">
<input type="hidden" name="DELETE_IDS" value="<%= request.form("DELETE_IDS") %>">


		<TR bgcolor='#FFFFFF'>
			<td width="20%" class="inputname">Departments</td>
			<td width="80%" class="inputvalue">
				<%= create_dept_list ("Department_Id","",5,"") %>
						
				<input type="hidden" name="Department_Id_C" value="Op|String|0|250|||Departments">
					<% small_help "Department" %></td>
		</tr>
			 
		<TR bgcolor='#FFFFFF'>
			<td width="20%" class="inputname">Retail Price</td>
			<td width="80%" class="inputvalue">
						<input name="Retail_Price" size="10">
						<INPUT type="hidden"  name="Retail_Price_C" value="Op|Integer|||||Retail Price">
						<% small_help "Retail Price" %></td>
		</tr>

	  	<TR bgcolor='#FFFFFF'>
			<td width="20%" class="inputname">Cost</td>
			<td width="80%" class="inputvalue">
						<input name="Wholesale_Price" size="10">
						<INPUT type="hidden" name="Wholesale_Price_C" value="Op|Integer|||||Cost Price">
						<% small_help "Wholesale Price" %></td>
		</tr>

		<TR bgcolor='#FFFFFF'>
			<td width="20%" class="inputname">Qty Discount</td>
			<td width="80%" class="inputvalue">
				<input class="image" type="radio" name="Use_Price_By_Matrix" value="-1">Yes&nbsp;
				<input class="image" type="radio" name="Use_Price_By_Matrix" value="0">No&nbsp;
				<input class="image" type="radio" name="Use_Price_By_Matrix" value="" checked>don't update&nbsp;
						<% small_help "Qty Discount" %></td>
		</tr>

		<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Home Page</td>
				<td width="80%" class="inputvalue">
					<input class="image" type="radio" name="Show_Homepage" value="-1">Yes&nbsp;
					<input class="image" type="radio" name="Show_Homepage" value="0">No&nbsp;
					<input class="image" type="radio" name="Show_Homepage" value="" checked>don't update&nbsp;
							<% small_help "Home Page" %></td>
		</tr>

		<TR bgcolor='#FFFFFF'>
			<td width="20%" class="inputname">Taxable</td>
			<td width="80%" class="inputvalue">
						<input class="image" type="radio" name="Taxable" value="-1">Yes&nbsp;
						<input class="image" type="radio" name="Taxable" value="0">No&nbsp;
						<input class="image" type="radio" name="Taxable" value="" checked>don't update&nbsp;
						<% small_help "Taxable" %></td>
		</tr>

		<TR bgcolor='#FFFFFF'>
			<td width="20%" class="inputname">Visible</td>
			<td width="80%" class="inputvalue">
						<input class="image" type="radio" name="Show" value="-1">Yes&nbsp;
						<input class="image" type="radio" name="Show" value="0">No&nbsp;
						<input class="image" type="radio" name="Show" value="" checked>don't update&nbsp;
						<% small_help "Visible" %></td>
		</tr>

		<TR bgcolor='#FFFFFF'>
			<td width="20%" class="inputname">Hide Price</td>
			<td width="80%" class="inputvalue">
						<input class="image" type="radio" name="Hide_Price" value="-1">Yes&nbsp;
						<input class="image" type="radio" name="Hide_Price" value="0">No&nbsp;
						<input class="image" type="radio" name="Hide_Price" value="" checked>don't update&nbsp;
						<% small_help "Hide Price" %></td>
		</tr>

		<TR bgcolor='#FFFFFF'>
			<td width="20%" class="inputname">Custom Pricing</td>
			<td width="80%" class="inputvalue">
						<input class="image" type="radio" name="Cust_Price" value="-1">Yes&nbsp;
						<input class="image" type="radio" name="Cust_Price" value="0">No&nbsp;
						<input class="image" type="radio" name="Cust_Price" value="" checked>don't update&nbsp;
						<% small_help "Custom Pricing" %></td>
		</tr>
				<tr bgcolor='#FFFFFF'><td width="30%" class="inputname" colspan=2>Short Description<BR>
						<textarea rows='12' name='Description_S' cols='83' id='Description_S'></textarea>
						<input type="hidden" name="Description_S_C" value="Op|String|0|4000|||Short Description">
						<% small_help "Short Description" %></td>
			</tr>
  
		<TR bgcolor='#FFFFFF'>
			<td width="20%" class="inputname">Small Image</td>
			<td width="80%" class="inputvalue">
						<input type="text" name="ImageS_Path" size="20" >
						<input type="hidden" name="ImageS_Path_C" value="Op|String|0|100|||Small Image">
						<font size="1" color="#000080"><a class="link" href="JavaScript:goImagePicker('ImageS_Path');"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
						<a class="link" href="JavaScript:goFileUploader('ImageS_Path');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
					<% small_help "Small Image" %></td>
		</tr>
	  
				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname" colspan=2>Large Description<BR>
						<textarea rows='12' name='Description_L' cols='83' id='Description_L'></textarea>
						<% small_help "Large Description" %></td>
			</tr>
	  
		<TR bgcolor='#FFFFFF'>
			<td width="20%" class="inputname">Large Image</td>
			<td width="80%" class="inputvalue">
						<input type="text" name="ImageL_Path" size="20" >
						<input type="hidden" name="ImageL_Path_C" value="Op|String|0|100|||Large Image">
						<font size="1" color="#000080"><a class="link" href="JavaScript:goImagePicker('ImageL_Path');"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
				<a class="link" href="JavaScript:goFileUploader('ImageL_Path');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>

				<% small_help "Large Image" %></td>
		</tr>

		 <% if instr(Shipping_Classes,"2") or instr(Shipping_Classes,"6") or instr(Shipping_Classes,"7") then %>
		  <TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Shipping Weight</td>
				<td width="80%" class="inputvalue">
						<input	name="Item_Weight" size="8" onKeyPress="return goodchars(event,'-0123456789.')">
						<INPUT type="hidden"  name="Item_Weight_C" value="Op|Integer|||||Weight">
				<% small_help "Weight" %></td>
			</tr>

			<% end if %>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Handling Fee</td>
				<td width="80%" class="inputvalue">
							<input	name="Item_Handling" size="8" onKeyPress="return goodchars(event,'-0123456789.')">
							<INPUT type="hidden"  name="Item_Handling_C" value="Op|Integer|||||Handling">
					<% small_help "Handling" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Waive Shipping</td>
				<td width="80%" class="inputvalue">
						<input class="image" type="radio" name="Waive_Shipping" value="-1">Yes&nbsp;
						<input class="image" type="radio" name="Waive_Shipping" value="0">No&nbsp;
						<input class="image" type="radio" name="Waive_Shipping" value="" checked>don't update&nbsp;
					<% small_help "Waive Shipping" %></td>
			</tr>

			 <% if Service_Type => 7 then %>
					<TR bgcolor='#FFFFFF'>
						<td width="20%" class="inputname">Shipping From</td>
						<td width="80%" class="inputvalue">
						<%= create_location_list ("Ship_Location_Id","",5) %>
						
						<% small_help "Shipfrom" %></td>
					</tr>
			<% end if %>

			<% if instr(Shipping_Classes,"3") > 0 then %>
		  <TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Shipping Fee</td>
				<td width="80%" class="inputvalue">
						<%= Store_Currency%><input name="Shipping_Fee"  size="10" onKeyPress="return goodchars(event,'-0123456789.')" maxlength=10>
						<INPUT type="hidden"  name="Shipping_Fee_C" value="Op|Integer|||||Shipping Fee">
					<% small_help "Shipping Fee" %></td>
			</tr>
			<% else %>
				<input type=hidden name="Shipping_Fee" >
			<% end if %>
			
		<% if Service_Type => 5 then %>
			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>Inventory Management</b></td>
				<td width="75%" colspan=2>&nbsp;</td>
			</tr>
			 <TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Stock</td>
				<td width="80%" class="inputvalue">
						<input	name="Quantity_in_stock"  size="8" onKeyPress="return goodchars(event,'-0123456789.')">
						<INPUT type="hidden"  name="Quantity_in_stock_C" value="Op|Integer|||||Stock">
						<% small_help "Stock" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Qty Control</td>
				<td width="80%" class="inputvalue">
						<input class="image" type="radio" name="Quantity_Control" value="-1">Yes&nbsp;
						<input class="image" type="radio" name="Quantity_Control" value="0">No&nbsp;
						<input class="image" type="radio" name="Quantity_Control" value="" checked>don't update&nbsp;
						<% small_help "Qty Control" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Control #</td>
				<td width="80%" class="inputvalue">
						<input name="Quantity_Control_Number"  size="4" onKeyPress="return goodchars(event,'-0123456789.')">
						<INPUT type="hidden"  name="Quantity_Control_Number_C" value="Op|Integer|||||Quantity Control">
						<% small_help "Control #" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Hide Stock</td>
				<td width="80%" class="inputvalue">
						<input class="image" type="radio" name="Hide_Stock" value="-1">Yes&nbsp;
						<input class="image" type="radio" name="Hide_Stock" value="0">No&nbsp;
						<input class="image" type="radio" name="Hide_Stock" value="" checked>don't update&nbsp;
						<% small_help "Control #" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Min Quantity</td>
				<td width="80%" class="inputvalue">
							 <input name="Quantity_Minimum"  size="4" onKeyPress="return goodchars(event,'-0123456789.')">
						<INPUT type="hidden"  name="Quantity_Minimum_C" value="Op|Integer|||||Min Quantity">
						<% small_help "Min Quantity" %></td>
			</tr>
			
			<TR bgcolor='#FFFFFF'>
					<td width="20%" class="inputname">Fractional</td>
					<td width="80%" class="inputvalue">
						<input class="image" name="Fractional" type="radio" value="-1">Yes&nbsp;
						<input class="image" name="Fractional" type="radio" value="0">No&nbsp;
						<input class="image" name="Fractional" type="radio" value="" checked>Don't Update&nbsp;
					<% small_help "Fractional" %></td>
				</tr>
			  <TR bgcolor='#FFFFFF'>


			<TR bgcolor='#FFFFFF'>
					<td width="30%"><b>Search Engine Tags</b></td>
					<td width="75%" colspan=2>Help get your items listed in search engines</td>
			</tr>
		
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Title</td>
				<td width="80%" class="inputvalue">
						<input type=text	name="Meta_Title"  maxlength=100 size=30>
						<input type="hidden" name="Meta_Title_C" value="Op|String|0|100|||Meta Title">
						<% small_help "Meta_Title" %></td>
			</tr>
	
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Keywords</td>
				<td width="80%" class="inputvalue">
						<input readonly type=text name=remLenMKey size=3 class=char maxlength=3 value="250" class=image><font size=1><I>characters left</i></font>
			        <textarea	name="Meta_Keywords" cols=55 rows=3 onKeyDown="textCounter(this.form.Meta_Keywords,this.form.remLenMKey,250);" onKeyUp="textCounter(this.form.Meta_Keywords,this.form.remLenMKey,250);"></textarea>
						<input type="hidden" name="Meta_Keywords_C" value="Op|String|0|250|||Meta Keywords">
						<% small_help "Meta_Keywords" %></td>
			</tr>
			
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Description</td>
				<td width="80%" class="inputvalue">
						<input readonly type=text name=remLenMDes size=3 class=char maxlength=3 value="500" class=image><font size=1><I>characters left</i></font>
			        <textarea	name="Meta_Description" cols=55 rows=5 onKeyDown="textCounter(this.form.Meta_Description,this.form.remLenMDes,500);" onKeyUp="textCounter(this.form.Meta_Description,this.form.remLenMDes,500);"></textarea>
						<input type="hidden" name="Meta_Description_C" value="Op|String|0|500|||Meta Description">
						<% small_help "Meta_Description" %></td>
			</tr>
			<% end if %>


			<% if Service_Type < 5 then %>
			
			<% else %>
			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>Recurring</b></td>
				<td width="75%" colspan=2>Note that some processors do not support recurring billing</font></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Recurring Fee</td>
				<td width="80%" class="inputvalue">
						<%= Store_Currency%><input name="Recurring_fee" size="5" value="0" onKeyPress="return goodchars(event,'-0123456789.')">
						<INPUT type="hidden"  name="Recurring_fee_C" value="Op|Integer|||||Recurring Fee">
						<% small_help "Recurring_Fee" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Recurring Interval</td>
				<td width="80%" class="inputvalue">
						<input name="Recurring_days"  size="5" onKeyPress="return goodchars(event,'0123456789')">days
						<INPUT type="hidden"  name="Recurring_days_C" value="Op|Integer|||||Recurring Days">
						<% small_help "Recurring_Days" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>Google Base</b></td>
				<td width="75%" colspan=2></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Condition</td>
			<td width="70%" class="inputvalue">
				<select name="Condition">
				<option value="" selected>Dont Update</option>
				<option value="New">New</option>
				<option value="Used">Used</option>
				<option value="Refurbished">Refurbished</option></select>
						<% small_help "Condition" %></td>
			</tr>

				<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Brand</td>
			<td width="70%" class="inputvalue">
						<input name="Brand" value="" size="60" maxlength=100>
						<INPUT type="hidden"  name="Brand_C" value="Op|String|0|100|||Brand">
						<% small_help "Brand" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Product Type</td>
			<td width="70%" class="inputvalue">
						<input name="Product_Type" value="" size="60" maxlength=20>
						<INPUT type="hidden"  name="Product_Type_C" value="Op|String|0|20|||Product Type">
						<% small_help "Product_Type" %></td>
			</tr>
	
			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>Specials</b></td>
				<td width="75%" colspan=2>Not effective when quantity discounts are enabled.</font></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Special Price</td>
				<td width="80%" class="inputvalue">Discount
						<input name="Retail_Price_special_Discount"  size="5" onKeyPress="return goodchars(event,'-0123456789.')">%
						<INPUT type="hidden"  name="Retail_Price_special_Discount_C" value="Op|Integer|||||Discount">
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
						<input name="Special_Start_Date" size="10" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')">
						<A HREF="#" onClick="cal1.select(document.forms[0].Special_Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Special_Start_Date.value=='')?document.forms[0].Special_Start_Date.value:null); return false;" TITLE="Start Date" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>
				     <INPUT type="hidden"  name="Special_Start_Date_C" value="Op|date|||||Special Start Date"> and
						<input name="Special_End_Date" size="10" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')">
						<A HREF="#" onClick="cal1.select(document.forms[0].Special_End_Date,'anchor2','M/d/yyyy',(document.forms[0].Special_End_Date.value=='')?document.forms[0].Special_End_Date.value:null); return false;" TITLE="End Date" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>
		           <INPUT type="hidden"  name="Special_end_Date_C" value="Op|date|||||Special End Date">
						<% small_help "Dates" %></td>
				</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>Electronic Software Delivery</b></td>
				<td width="75%" colspan=2>Fill this out if your item is a software file, ie MP3, Game, etc.</td>
			</tr>
	  
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Filename</td>
				<td width="80%" class="inputvalue">
							<input type="text" name="File_Location"  size="30" maxlength=250>
							<input type="hidden" name="File_Location_C" value="Op|String|0|250|||Filename">
							<a class="link" href="javascript:goFilePicker('File_Location')"><img border="0" src="images/image.gif" width="23" height="22" alt="File Picker"></a>
							<% small_help "Filename" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Pin Delivery</td>
				<td width="80%" class="inputvalue">
							<input class="image" name="Item_pin" type="radio" value="-1">Yes&nbsp;
							<input class="image" name="Item_pin" type="radio" value="0">No&nbsp;
							<input class="image" name="Item_pin" type="radio" value="" checked>Don't Update&nbsp;
							<% small_help "Item_pin" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">Item Remarks<BR><font size="1">Not seen by customers</font></td>
				<td width="80%" class="inputvalue">
					<input readonly type=text name=remLenRem size=3 class=char maxlength=3 value="200" class=image ><font size=1><I>characters left</i></font>
					<textarea name="Item_Remarks" cols="55" rows="2" onKeyDown="textCounter(this.form.Item_Remarks,this.form.remLenRem,200);" onKeyUp="textCounter(this.form.Item_Remarks,this.form.remLenRem,200);"></textarea>
							<input type="hidden" name="Item_Remarks_C" value="Op|String|0|200|||Item Remarks">
							<% small_help "Item Remarks" %></td>
			</tr>
	
			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>User Definable Fields</b></td>
				<td width="75%" colspan=2>Custom named fields, provides a textbox for the user to fill in when purchasing item.</td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">User Field 1</td>
				<td width="80%" class="inputvalue">Use
							<input class="image" name="U_d_1" type="radio" value="-1">Yes&nbsp;
							<input class="image" name="U_d_1" type="radio" value="0">No&nbsp;
							<input class="image" name="U_d_1" type="radio" value="" checked>Don't Update&nbsp;
							<br>Name
							<input	name="U_d_1_name"  size="30" maxlength=250>
							<input type="hidden" name="U_d_1_C" value="Op|String|0|250|||User Field 1">
							<% small_help "User Field" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">User Field 2</td>
				<td width="80%" class="inputvalue">Use
						<input class="image" name="U_d_2" type="radio" value="-1">Yes&nbsp;
						<input class="image" name="U_d_2" type="radio" value="0">No&nbsp;
						<input class="image" name="U_d_2" type="radio" value="" checked>Don't Update&nbsp;
						<br>
						Name
						<input	name="U_d_2_name"  size="30" maxlength=250>
						<input type="hidden" name="U_d_2_C" value="Op|String|0|250|||User Field 2">
						<% small_help "User Field" %></td>
			</tr>
		
			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">User Field 3</td>
				<td width="80%" class="inputvalue">Use
							<input class="image" name="U_d_3" type="radio" value="-1">Yes&nbsp;
							<input class="image" name="U_d_3" type="radio" value="0">No&nbsp;
							<input class="image" name="U_d_3" type="radio" value="" checked>Don't Update&nbsp;
							<br>
							Name
							<input	name="U_d_3_name" size="30" maxlength=250>
							<input type="hidden" name="U_d_3_C" value="Op|String|0|250|||User Field 3">
							<% small_help "User Field" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">User Field 4</td>
				<td width="80%" class="inputvalue">Use
							<input class="image" name="U_d_4" type="radio" value="-1">Yes&nbsp;
							<input class="image" name="U_d_4" type="radio" value="0">No&nbsp;
							<input class="image" name="U_d_4" type="radio" value="" checked>Don't Update&nbsp;
							<br>
							Name
							<input	name="U_d_4_name"  size="30" maxlength=250>
							<input type="hidden" name="U_d_4_C" value="Op|String|0|250|||User Field 4">
							<% small_help "User Field" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="20%" class="inputname">User Field 5</td>
				<td width="80%" class="inputvalue">Use
							<input class="image" name="U_d_5" type="radio" value="-1">Yes&nbsp;
							<input class="image" name="U_d_5" type="radio" value="0">No&nbsp;
							<input class="image" name="U_d_5" type="radio" value="" checked>Don't Update&nbsp;
							<br>
							Name
							<input	name="U_d_5_name" size="30" maxlength=250>
							<input type="hidden" name="U_d_5_C" value="Op|String|0|250|||User Field 5">
							<% small_help "User Field" %></td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="30%"><b>Extended Fields</b></td>
				<td width="75%" colspan=2>Extended fields, are displayed in addition to other item details</td>
			</tr>

			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname" colspan=2>Extended Field 1<BR>
				<textarea rows='12' name='M_d_1' cols='83' id='M_d_1'></textarea>

					<input type="hidden" name="M_d_1_C" value="Op|String|0|4000|||Extended Field 1">
				<% small_help "Extended Field 1" %></td>
			</tr>

			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname" colspan=2>Extended Field 2<BR>
				<textarea rows='12' name='M_d_2' cols='83' id='M_d_2'></textarea>
	<input type="hidden" name="M_d_2_C" value="Op|String|0|4000|||Extended Field 2">
				<% small_help "Extended Field 2" %></td>
			</tr>
			
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname" colspan=2>Extended Field 3<BR>
				<textarea rows='12' name='M_d_3' cols='83' id='M_d_3'></textarea>

					<input type="hidden" name="M_d_3_C" value="Op|String|0|4000|||Extended Field 3">
				<% small_help "Extended Field 3" %></td>
			</tr>
			
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname" colspan=2>Extended Field 4<BR>
				<textarea rows='12' name='M_d_4' cols='83' id='M_d_4'></textarea>

					<input type="hidden" name="M_d_4_C" value="Op|String|0|4000|||Extended Field 4">
				<% small_help "Extended Field 4" %></td>
			</tr>

			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname" colspan=2>Extended Field 5<BR>
				<textarea rows='12' name='M_d_5' cols='83' id='M_d_5'></textarea>

					<input type="hidden" name="M_d_5_C" value="Op|String|0|4000|||Extended Field 5">
				<% small_help "Extended Field 5" %></td>
			</tr>			
				 <% end if %>
<% createFoot thisRedirect, 1%>
<% end if %>

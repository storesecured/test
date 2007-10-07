<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sql_select = "select * from Store_Shipping_Labels WHERE store_id = "&Store_Id
rs_Store.open sql_select,conn_store,1,1
  Page_Width = rs_Store("Page_Width")
  Page_Height = rs_Store("Page_Height")
  Address_Width = rs_Store("Address_Width")
  Address_Height = rs_Store("Address_Height")
  Spacing_Width = rs_Store("Spacing_Width")
  Spacing_Height = rs_Store("Spacing_Height")
  Top_Margin = rs_Store("Top_Margin")
  Bottom_Margin = rs_Store("Bottom_Margin")
  Left_Margin = rs_Store("Left_Margin")
  Right_Margin = rs_Store("Right_Margin")
  Total_Rows = rs_Store("Total_Rows")
  Total_Cols = rs_Store("Total_Cols")
  Image_Name = rs_Store("Image_Name")
  Image_Pos = rs_Store("Image_Pos")
rs_Store.close

sFormAction = "Store_Settings.asp"
sName = "Shipping_Labels"
sCommonName = "Shipping Labels"
sFormName = "Shipping_Labels"
sTitle = "Shipping Labels"
sFullTitle = "Shipping > Labels"
sSubmitName = "Update_Cookies"
thisRedirect = "shipping_labels.asp"
sTopic = "Store_manage"
addPicker=1
sMenu = "shipping"
sQuestion_Path = "design/shipping_labels.htm"
sTextHelp="shipping/shipping-labels.doc"

createHead thisRedirect
%>

				<tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Page Width</b></td>
				<td class="inputvalue">
						<input name="Page_Width" value="<%= Page_Width %>" size="10" onKeyPress="return goodchars(event,'0123456789.')"> inches
						<input type="Hidden" name="Page_Width_C" value="Re|Integer|0|20000|||Page Width">
						<% small_help "Page_Width" %></td>
				</tr>
		          <tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Page Height</b></td>
				<td class="inputvalue">
						<input name="Page_Height" value="<%= Page_Height %>" size="10" onKeyPress="return goodchars(event,'0123456789.')"> inches
						<input type="Hidden" name="Page_Height_C" value="Re|Integer|0|20000|||Page Height">
						<% small_help "Page_Height" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Address Width</b></td>
				<td class="inputvalue">
						<input name="Address_Width" value="<%= Address_Width %>" size="10" onKeyPress="return goodchars(event,'0123456789.')"> inches
						<input type="Hidden" name="Address_Width_C" value="Re|Integer|0|20000|||Address Width">
						<% small_help "Address_Width" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Address Height</b></td>
				<td class="inputvalue">
						<input name="Address_Height" value="<%= Address_Height %>" size="10" onKeyPress="return goodchars(event,'0123456789.')"> inches
						<input type="Hidden" name="Address_Height_C" value="Re|Integer|0|20000|||Address Height">
						<% small_help "Address_Height" %></td>
				</tr>
                    <tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Spacing Width</b></td>
				<td class="inputvalue">
						<input name="Spacing_Width" value="<%= Spacing_Width %>" size="10" onKeyPress="return goodchars(event,'0123456789.')"> inches
						<input type="Hidden" name="Spacing_Width_C" value="Re|Integer|0|20000|||Spacing Width">
						<% small_help "Spacing_Width" %></td>
				</tr>
                    <tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Spacing Height</b></td>
				<td class="inputvalue">
						<input name="Spacing_Height" value="<%= Spacing_Height %>" size="10" onKeyPress="return goodchars(event,'0123456789.')"> inches
						<input type="Hidden" name="Spacing_Height_C" value="Re|Integer|0|20000|||Spacing Height">
						<% small_help "Spacing_Height" %></td>
				</tr>
                    <tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Top Margin</b></td>
				<td class="inputvalue">
						<input name="Top_Margin" value="<%= Top_Margin %>" size="10" onKeyPress="return goodchars(event,'0123456789.')"> inches
						<input type="Hidden" name="Top_Margin_C" value="Re|Integer|0|20000|||Top Margin">
						<% small_help "Top_Margin" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Bottom Margin</b></td>
				<td class="inputvalue">
						<input name="Bottom_Margin" value="<%= Bottom_Margin %>" size="10" onKeyPress="return goodchars(event,'0123456789.')"> inches
						<input type="Hidden" name="Bottom_Margin_C" value="Re|Integer|0|20000|||Bottom Margin">
						<% small_help "Bottom_Margin" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Left Margin</b></td>
				<td class="inputvalue">
						<input name="Left_Margin" value="<%= Left_Margin %>" size="10" onKeyPress="return goodchars(event,'0123456789.')"> inches
						<input type="Hidden" name="Left_Margin_C" value="Re|Integer|0|20000|||Left Margin">
						<% small_help "Left_Margin" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Right Margin</b></td>
				<td class="inputvalue">
						<input name="Right_Margin" value="<%= Right_Margin %>" size="10" onKeyPress="return goodchars(event,'0123456789.')"> inches
						<input type="Hidden" name="Right_Margin_C" value="Re|Integer|0|20000|||Right Margin">
						<% small_help "Right_Margin" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Total Rows</b></td>
				<td class="inputvalue">
						<input name="Total_Rows" value="<%= Total_Rows %>" size="10">
						<input type="Hidden" name="Total_Rows_C" value="Re|Integer|0|20000|||Total Rows" onKeyPress="return goodchars(event,'0123456789.')">
						<% small_help "Total_Rows" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Total Cols</b></td>
				<td class="inputvalue">
						<input name="Total_Cols" value="<%= Total_Cols %>" size="10">
						<input type="Hidden" name="Total_Cols_C" value="Re|Integer|0|20000|||Total Cols" onKeyPress="return goodchars(event,'0123456789.')">
						<% small_help "Total_Cols" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Image Name</b></td>
				<td class="inputvalue">
						<input name="Image_Name" value="<%= Image_Name %>" size="60" maxlength=100>
						<input type="Hidden" name="Image_Name_C" value="Re|String|0|100|||Image Name"><a href="javascript:goImagePicker('Image_Name')"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
						<% small_help "Image Name" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Image Pos</b></td>
				<td class="inputvalue">
						<select name="Image_Pos">
						<% if Image_Pos = 0 then
						      Pos0= "selected"
						   elseif Image_Pos = 1 then
						      Pos1 = "selected"
						   elseif Image_Pos = 2 then
						      Pos2 = "selected"
						   elseif Image_Pos = 3 then
						      Pos3 = "selected"
						   end if %>
                  <option <%=Pos0 %> value="0">No Image
                  <option <%=Pos1 %> value="1">Replace From
                  <option <%=Pos2 %> value="2">Replace Invoice
                  <option <%=Pos3 %> value="3">Replace To
                  </select>
						<input type="Hidden" name="Image_Pos_C" value="Re|Integer|0|20000|||Total Cols">
						<% small_help "Image Pos" %></td>
				</tr>
<% createFoot thisRedirect,1 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Page_Width","req","Please enter a page width.");
 frmvalidator.addValidation("Page_Height","req","Please enter a page height.");
 frmvalidator.addValidation("Address_Width","req","Please enter a address width.");
 frmvalidator.addValidation("Address_Height","req","Please enter a address height.");
 frmvalidator.addValidation("Spacing_Width","req","Please enter a spacing width.");
 frmvalidator.addValidation("Spacing_Height","req","Please enter a spacing height.");
 frmvalidator.addValidation("Top_Margin","req","Please enter a top margin.");
 frmvalidator.addValidation("Bottom_Margin","req","Please enter a bottom margin.");
 frmvalidator.addValidation("Left_Margin","req","Please enter a left margin.");
 frmvalidator.addValidation("Right_Margin","req","Please enter a right margin.");
 frmvalidator.addValidation("Total_Rows","req","Please enter the total number of rows..");
 frmvalidator.addValidation("Total_Cols","req","Please enter the total number of columns.");
 frmvalidator.addValidation("Image_Name","req","Please enter a image name.");

</script>

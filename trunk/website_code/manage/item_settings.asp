<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="help/item_settings.asp"-->

<%
if Reload_Attr <> 0 then
	checked_Reload_Attr = "checked"
else
	checked_Reload_Attr = ""
end if

if Unverified_Reduce <> 0 then
	checked_Unverified_Reduce = "checked"
else
	checked_Unverified_Reduce = ""
end if

if Inventory_Reduce <> 0 then
	checked_Inventory_Reduce = "checked"
else
	checked_Inventory_Reduce = ""
end if

if Show_Jump <>	0	then
	 Show_Jump_Checked = "Checked"
else
	 Show_Jump_Checked = ""
end if

if Hide_Empty_Depts <>	0	then
	 Hide_Empty_Depts_Checked = "Checked"
else
	 Hide_Empty_Depts_Checked = ""
end if

if Hide_OutofStock_Items <>	0	then
	 Hide_OutofStock_Items_Checked = "Checked"
else
	 Hide_OutofStock_Items_Checked = ""
end if

if Show_TopNav <>   0  then
	 Show_TopNav_Checked = "Checked"
else
	 Show_TopNav_Checked = ""
end if

if Show_ItemNav <>   0  then
	 Show_ItemNav_Checked = "Checked"
else
	 Show_ItemNav_Checked = ""
end if

if Show_SubDept <>   0  then
	 Show_SubDept_Checked = "Checked"
else
	 Show_SubDept_Checked = ""
end if


sNeedTabs=1
sFormAction = "Store_Settings.asp"
sName = "Inventory_Configuration"
sFormName = "item_settings"
sTitle = "Inventory Settings"
sFullTitle = "Inventory > Settings"
sCommonName = "Inventory Settings"
sCancel = "edit_items.asp"
sSubmitName = "Update"
thisRedirect = "item_settings.asp"
sTopic="Inventory Configuration"
sMenu="inventory"
createHead thisRedirect
if Service_Type < 5	then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		SILVER Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
	<% createFoot thisRedirect, 0%>
	</td>
	</tr>
<% else %>


<tr bgcolor='#FFFFFF'><td width="724" align=center valign=top height=35>
<table border=0 cellspacing=0 cellpadding=0 width=724>

	<!-- TAB MENU STARTS HERE -->

		<tr>
		<td align="center" valign=top height=45 width='100%'><br>
		<script type="text/javascript" language="JavaScript1.2" src="include/tabs-xp.js"></script>
		<script language="javascript1.2">
		var bselectedItem   = 0;
		var bmenuItems =
		[
		["Settings",  "content4",,,"Settings","Settings"],
		["User Defined Fields",  "content9",,,"User Defined Fields","User Defined Fields"],
		];

		apy_tabsInit();
		</script>
		</td>
		</tr>
		
		<TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='25'>
		
		
		<!-- CONTENT 4 STARTS HERE -->
			<div id="content4" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>

				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname" rowspan=2>Update stock automatically</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Inventory_Reduce" value="-1" <%= checked_Inventory_Reduce %>> Verified Orders
				<% small_help "Update Stock" %></td></tr>
                
				<tr bgcolor='#FFFFFF'><td><input class="image" type="checkbox" name="Unverified_Reduce" value="-1" <%= checked_Unverified_Reduce %>> Unverified Orders
				<% small_help "Update Stock" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">Hide Out of Stock Items</td>
					<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Hide_OutofStock_Items" value="-1" <%= Hide_OutofStock_Items_Checked %>>&nbsp;
					<% small_help "Hide Empty Stock" %></td>
					</tr>
				
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Display A-Z Links</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Show_Jump" value="-1" <%= Show_Jump_Checked %>>&nbsp;
				<% small_help "Show Jump" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Hide Empty Depts</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Hide_Empty_Depts" value="-1" <%= Hide_Empty_Depts_Checked %>>&nbsp;
				<% small_help "Hide Empty Depts" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Show Top Dept Navigation</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Show_TopNav" value="-1" <%= Show_TopNav_Checked %>>&nbsp;
				<% small_help "Show_TopNav" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Item Detail Prev-Next</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Show_ItemNav" value="-1" <%= Show_ItemNav_Checked %>>&nbsp;
				<% small_help "Show_ItemNav" %></td>
				</tr>
				
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Sub Dept Above Items</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Show_SubDept" value="-1" <%= Show_SubDept_Checked %>>&nbsp;
				<% small_help "Show_SubDept" %></td>
				</tr>
				
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Reload Attributes in IE</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Reload_Attr" value="-1" <%= checked_Reload_Attr %>>&nbsp;
				<% small_help "Reload_Attr" %></td>
				</tr>
				
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname" rowspan=2>Items Per Page</td>
				<td width="70%" class="inputvalue"><select name=item_display>
				<option value=<%= item_display %> selected><%= item_display+1 %></option>
				<option value=0>1</option>
				<option value=1>2</option>
				<option value=2>3</option>
				<option value=3>4</option>
				<option value=4>5</option>
				</select> Col(s)
				<INPUT type="hidden"  name=item_display_C value="Re|Integer|0|5|||Item">
				<% small_help "Show_SubDept" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="70%" class="inputvalue"><input name="item_rows" type="text" value="<%= item_rows %>" size=3 maxlength=3 onKeyPress="return goodchars(event,'0123456789')"> Row(s)
				<INPUT type="hidden"  name=item_rows_C value="Re|Integer|1|50|||Item Rows">
				<% small_help "Show_SubDept" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname" rowspan=2>Featured Items Per Page</td>
				<td width="70%" class="inputvalue"><select name=item_f_display>
				<option value=<%= item_f_display %> selected><%= item_f_display+1 %></option>
				<option value=0>1</option>
				<option value=1>2</option>
				<option value=2>3</option>
				<option value=3>4</option>
				<option value=4>5</option></select> Col(s)
				<INPUT type="hidden"  name=item_f_display_C value="Re|Integer|0|5|||Item">
				<% small_help "Show_SubDept" %></td>
				</tr>
				
				<TR bgcolor='#FFFFFF'>
				<td width="70%" class="inputvalue"><input name="item_f_rows" type="text" value="<%= item_f_rows %>" size=3 maxlength=3 onKeyPress="return goodchars(event,'0123456789')"> Row(s)
				<INPUT type="hidden"  name=item_f_rows_C value="Re|Integer|1|50|||Featured Item Rows">
				<% small_help "Show_SubDept" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname" rowspan=2>Departments Per Page</td>
				<td width="70%" class="inputvalue"><select name=dept_display>
				<option value=<%= dept_display %> selected><%= dept_display+1 %></option>
				<option value=0>1</option>
				<option value=1>2</option>
				<option value=2>3</option>
				<option value=3>4</option>
				<option value=4>5</option></select> Col(s)
				<INPUT type="hidden"  name=dept_display_C value="Re|Integer|0|5|||Item">
				<% small_help "Show_SubDept" %></td>
				</tr>
				
				<TR bgcolor='#FFFFFF'>
				<td width="70%" class="inputvalue"><input name="dept_rows" type="text" value="<%= dept_rows %>" size=3 maxlength=3 onKeyPress="return goodchars(event,'0123456789')"> Row(s)
				<INPUT type="hidden"  name=dept_rows_C value="Re|Integer|1|50|||Department Rows">
				<% small_help "Show_SubDept" %></td>
				</tr>
				

			</table>
			</div>		
			<!-- CONTENT 4 ENDS HERE -->

			<!-- CONTENT 9 STARTS HERE -->
			<div id="content9" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>

                                <TR bgcolor='#FFFFFF'>
					<td colspan=3>By default when user defined fields are enabled they are shown as text areas.  You may override this 
                                        default behavior by entering in the properties of your input text ield.  Ie size=30 maxlength=30 etc</td>
				</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">User Field 1</td>
					<td width="70%" class="inputvalue"><input type="text" name="User_Defined_Fields" value="<%= User_Defined_Fields %>" size=60>&nbsp;
					<input type="hidden" name="User_Defined_Fields_C" value="Op|String|0|100|||User Field 1"><% small_help "User Field 1" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">User Field 2</td>
					<td width="70%" class="inputvalue"><input type="text" name="User_Defined_Fields_2" value="<%= User_Defined_Fields_2 %>" size=60>&nbsp;
					<input type="hidden" name="User_Defined_Fields_2_C" value="Op|String|0|100|||User Field 2"><% small_help "User Field 2" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">User Field 3</td>
					<td width="70%" class="inputvalue"><input type="text" name="User_Defined_Fields_3" value="<%= User_Defined_Fields_3 %>" size=60>&nbsp;
					<input type="hidden" name="User_Defined_Fields_3_C" value="Op|String|0|100|||User Field 3"><% small_help "User Field 3" %></td>
					</tr>

				<TR bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">User Field 4</td>
					<td width="70%" class="inputvalue"><input type="text" name="User_Defined_Fields_4" value="<%= User_Defined_Fields_4 %>" size=60>&nbsp;
					<input type="hidden" name="User_Defined_Fields_4_C" value="Op|String|0|100|||User Field 4"><% small_help "User Field 4" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">User Field 5</td>
					<td width="70%" class="inputvalue"><input type="text" name="User_Defined_Fields_5" value="<%= User_Defined_Fields_5 %>" size=60>&nbsp;
					<input type="hidden" name="User_Defined_Fields_5_C" value="Op|String|0|100|||User Field 5"><% small_help "User Field 5" %></td>
				</tr>

			</table>
			</div>
			<!-- CONTENT 9 ENDS HERE -->

		</td>
		</tr>	
			 
</table></td></tr>					
<tr bgcolor=#FFFFFF>
<td colspan='4' class='tpage'>		
<% createFoot thisRedirect,1 %>
</td>
</tr>

<% end if %>

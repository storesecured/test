<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%

select case dept_display
		case 0
			checked_dept0="checked"
		case 1
			checked_dept1="checked"
		case 2
			checked_dept2="checked"
		case 3
			checked_dept3="checked"
		case 4
			checked_dept4="checked"
		case 5
			checked_dept5="checked"
		end select
		
	select case item_display
		case 0
			checked_item0="checked"
		case 1
			checked_item1="checked"
		case 2
			checked_item2="checked"
		case 3
			checked_item3="checked"
		case 4
			checked_item4="checked"
		case 5
			checked_item5="checked"
	end select
	
	select case item_f_display
		case 0
			checked_itemf0="checked"
		case 1
			checked_itemf1="checked"
		case 2
			checked_itemf2="checked"
		case 3
			checked_itemf3="checked"
		case 4
			checked_itemf4="checked"
		case 5
			checked_itemf5="checked"
	end select
	
	if index = -1 then 
		checked1 = "checked"
	else 
		checked0 = "checked"
	end if

rs_Store.Close

sInstructions="Use this page to configure your display settings for departments and items.  You should always preview with current theme if selecting anything other than 1 Across to ensure products are arranged correctly.  Remember that some shoppers will have small monitors and making your pages too wide may make the shopping experience unpleasant for them if they have to scroll sideways."

sFormAction = "Store_Settings.asp"
sName = "Inventory_display"
sFormName = "Inventory_display"
sTitle = "Inventory Display"
sSubmitName = "Inventory_display_Update"
thisRedirect = "inventory_display.asp"
addPicker = 1
sTopic = "inventory_display"
sMenu = "design"
sQuestion_Path = "design/inventory_display.htm"
createHead thisRedirect
if Service_Type < 3 then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		BRONZE Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0%>

<% else %>


				<TR bgcolor='#FFFFFF'>
					<td></td><td><b>Departments</b></td><td><b>Items</b></td><td><b>Featured Items</b></td><td></td>
				</tr>
				
				<TR bgcolor='#FFFFFF'> 
				<td width="10%" class="inputname"><B>1 Across</b></td>
				<td width="26%" class="inputvalue"><input class="image" name="dept_display" type="radio" value="0" <%= checked_dept0 %>>&nbsp; </td>
				<td width="26%" class="inputvalue"><input class="image" name="item_display" type="radio" value="0" <%= checked_item0 %>>&nbsp; </td>
				<td width="26%" class="inputvalue"><input class="image" name="item_f_display" type="radio" value="0" <%= checked_itemf0 %>>&nbsp;
				<INPUT type="hidden"  name=dept_display_C value="Re|Integer|0|5|||Departments">
				<INPUT type="hidden"  name=item_display_C value="Re|Integer|0|5|||Item">
				<INPUT type="hidden"  name=item_f_display_C value="Re|Integer|0|5|||Featured Items">
				<% small_help "1 Across" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><B>2 Across</b></td>
				<td width="26%" class="inputvalue"><input class="image" name="dept_display" type="radio" value="1" <%= checked_dept1 %>>&nbsp; </td>
				<td width="26%" class="inputvalue"><input class="image" name="item_display" type="radio" value="1" <%= checked_item1 %>>&nbsp; </td>
				<td width="26%" class="inputvalue"><input class="image" name="item_f_display" type="radio" value="1" <%= checked_itemf1 %>>&nbsp;
				<% small_help "2 Across" %></td>

				</tr>
		  
				<TR bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><B>3 Across</b></td>
				<td width="26%" class="inputvalue"><input class="image" name="dept_display" type="radio" value="2" <%= checked_dept2 %>>&nbsp; </td>
				<td width="26%" class="inputvalue"><input class="image" name="item_display" type="radio" value="2" <%= checked_item2 %>>&nbsp; </td>
				<td width="26%" class="inputvalue"><input class="image" name="item_f_display" type="radio" value="2" <%= checked_itemf2 %>>&nbsp;
				<% small_help "3 Across" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><B>4 Across</b></td>
				<td width="26%" class="inputvalue"><input class="image" name="dept_display" type="radio" value="3" <%= checked_dept3 %>>&nbsp; </td>
				<td width="26%" class="inputvalue"><input class="image" name="item_display" type="radio" value="3" <%= checked_item3 %>>&nbsp; </td>
				<td width="26%" class="inputvalue"><input class="image" name="item_f_display" type="radio" value="3" <%= checked_itemf3 %>>&nbsp;
				<% small_help "4 Across" %></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><B>5 Across</b></td>
				<td width="26%" class="inputvalue"><input class="image" name="dept_display" type="radio" value="4" <%= checked_dept4 %>>&nbsp; </td>
				<td width="26%" class="inputvalue"><input class="image" name="item_display" type="radio" value="4" <%= checked_item4 %>>&nbsp; </td>
				<td width="26%" class="inputvalue"><input class="image" name="item_f_display" type="radio" value="4" <%= checked_itemf4 %>>&nbsp;
				<% small_help "5 Across" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><B>Max Rows</b></td>
				<td width="26%" class="inputvalue"><input name="dept_rows" type="text" value="<%= dept_rows %>" size=3 maxlength=3 onKeyPress="return goodchars(event,'0123456789')">
				<INPUT type="hidden"  name=dept_rows_C value="Re|Integer|1|50|||Dept Rows"></td>
				<td width="26%" class="inputvalue"><input name="item_rows" type="text" value="<%= item_rows %>" size=3 maxlength=3 onKeyPress="return goodchars(event,'0123456789')">
				<INPUT type="hidden"  name=item_rows_C value="Re|Integer|1|50|||Item Rows"></td>
				<td width="26%" class="inputvalue"><input name="item_f_rows" type="text" value="<%= item_f_rows %>" size=3 maxlength=3 onKeyPress="return goodchars(event,'0123456789')">
				<INPUT type="hidden"  name=item_f_rows_C value="Re|Integer|1|50|||Item Featured Rows">
				<% small_help "Max Rows" %></td>
				</tr>




<% createFoot thisRedirect,1 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("dept_rows","greaterthan=0","Please enter maximum department rows between 1 and 50.");
 frmvalidator.addValidation("dept_rows","lessthan=51","Please enter maximum department rows between 1 and 50.");
 frmvalidator.addValidation("dept_rows","req","Please enter maximum department rows between 1 and 50.");
 frmvalidator.addValidation("item_rows","greaterthan=0","Please enter maximum item rows between 1 and 50.");
 frmvalidator.addValidation("item_rows","lessthan=51","Please enter maximum item rows between 1 and 50.");
 frmvalidator.addValidation("item_rows","req","Please enter maximum department rows between 1 and 50.");
 frmvalidator.addValidation("item_f_rows","greaterthan=0","Please enter maximum item featured rows between 1 and 50.");
 frmvalidator.addValidation("item_f_rows","lessthan=51","Please enter maximum item featured rows between 1 and 50.");
 frmvalidator.addValidation("item_f_rows","req","Please enter maximum department rows between 1 and 50.");

</script>
<% end if %>

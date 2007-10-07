<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->
<!--#include virtual="common/common_functions.asp"-->


<%

Switch_Name=Site_Name

on error resume next
if Request.Form("Form_Name") = "Dept_Edit" then
	'ERROR CHECKING
	
	if not CheckReferer then
		Response.Redirect "admin_error.asp?message_id=2"
	end if

	If Form_Error_Handler(Request.Form) <> "" then 
		Error_Log = Form_Error_Handler(Request.Form)
		%><!--#include file="Include/Error_Template.asp"--><%
	       response.end
        else
		'RETRIEVE FORM DATA
		Department_ID = Request.Form("Department_ID")
		Belong_to = Request.Form("Belong_to")

		if isNumeric(Department_ID) and isNumeric(Belong_to) then
			Department_Name = checkStringForQ(Request.Form("Department_Name"))
			Department_Image_Path = checkStringForQ(Request.Form("Department_Image_Path"))
			Department_Description = replace(Request.Form("Department_Description"),"'","''")
			Department_HTML = replace(Request.Form("Department_HTML"),"'","''")
			Department_HTML_Bottom = replace(Request.Form("Department_HTML_Bottom"),"'","''")
			Navig_Button_Menu=request.form("navig_button_menu")
			Navig_Link_Menu=request.form("navig_link_menu")
			Meta_Keywords = checkStringForQ(Request.Form("Meta_Keywords"))
			Meta_Description = checkStringForQ(Request.Form("Meta_Description"))
			Meta_Title = checkStringForQ(Request.Form("Meta_Title"))
			Visible = request.form("Visible")
			if Visible <> "" then
				Visible = 1
			else
				Visible = 0
			end if
			Show_Name = request.form("Show_Name")
			if Show_Name <> "" then
				Show_Name = 1
			else
				Show_Name = 0
			end if
			Protect_Page = request.form("Protect_Page")
			if Protect_Page <> "" then
				Protect_Page = 1
			else
				Protect_Page = 0
			end if
			Customer_Group = request.form("Customer_Group")

			View_Order = request.form("View_Order")
			'UPDATE DEPARTMENT DATA
			sql_select = "SELECT count(Department_ID) as count_of_children from Store_Dept where Belong_to = "&Department_ID&" and Store_id="&Store_id&""
		rs_Store.open sql_select,conn_store,1,1 
			count_of_children = rs_store("count_of_children")
		rs_Store.close
		if count_of_children > 0 then
			iLast_Level = 0
		else
			  iLast_Level = 1
		end if
                        sql_update="update store_dept set Navig_Button_Menu="&Navig_Button_Menu&",Navig_Link_Menu="&Navig_Link_Menu&",Meta_Title='"&Meta_Title&"', Meta_Keywords='"&Meta_Keywords&"', Meta_Description='"&Meta_Description&"', Last_Level = "&iLast_Level&", Belong_to = '"&Belong_To&"', Department_HTML = '"&Department_HTML&"', Department_HTML_Bottom='"&Department_HTML_Bottom&"',Department_Name = '"&Department_Name&"', Department_Description = '"&Department_Description&"', Department_Image_Path = '"&Department_Image_Path&"', Visible="&Visible&", Show_Name="&Show_Name&",Protect_Page="&Protect_Page&", Customer_Group="&Customer_Group&",View_Order="&View_Order&" where Department_ID = "&Department_ID&" and Store_id="&Store_id&""
			conn_store.Execute sql_update

			sql_update="update store_dept set Last_Level = 0 where Department_ID = "&Belong_to&" and Department_ID <> 0 and store_id="&store_id
			conn_store.Execute sql_update

			domainName = lcase(Request.ServerVariables("SERVER_NAME"))
			Session("Department_List:"&Store_Id)=""
			
			server.execute "reset_design.asp"
			response.redirect "department_manager.asp"
		else
			Response.Redirect "admin_error.asp?message_id=1"
		end if
	end if
end if

Department_ID=Request.querystring("id")
if Department_ID="0" then
   fn_error "The UnAssigned department is a default department which cannot be modified."
end if
if isNumeric(Department_ID) then
	'SELECT STORE DEPARTMENT
	sql_select = "SELECT Navig_Button_Menu, Navig_Link_Menu, Full_name,Meta_Description,Meta_Keywords, Meta_Title, View_Order, Visible, Show_Name,Customer_Group,Protect_Page, Department_ID, Department_HTML, Department_HTML_Bottom, Department_Name,Department_Description,Department_Image_Path,Belong_to,Last_Level from store_dept where Department_ID = "&Department_ID&" and Store_id="&Store_id&" and department_id<>0"
	rs_Store.open sql_select,conn_store,1,1
		Full_name = rs_store("Full_name")
		Department_Name = rs_store("Department_Name")
		Department_Image_Path = rs_store("Department_Image_Path")
		Department_Description = rs_store("Department_Description")
		Department_ID = rs_store("Department_ID")
		Department_HTML = rs_Store("Department_HTML")
		Department_HTML_Bottom = rs_Store("Department_HTML_Bottom")
		Navig_Button_Menu=rs_Store("navig_button_menu")
		if Navig_Button_Menu=0 then
			Navig_Button_Menu_Text = "None"
		else
			Navig_Button_Menu_Text = Navig_Button_Menu
		end if
		Navig_Link_Menu=rs_Store("navig_link_menu")
		if Navig_Link_Menu=0 then
			Navig_Link_Menu_Text = "None"
		else
		     Navig_Link_Menu_Text = Navig_Link_Menu
		end if
		Belong_to = rs_Store("Belong_to")
		Visible = rs_Store("Visible")
		Show_Name = rs_Store("Show_Name")
		CustomerGroup = rs_Store("Customer_Group")
		Protect_Page = rs_Store("Protect_Page")
		View_Order = rs_Store("View_Order")
		Meta_Description = rs_Store("Meta_Description")
		Meta_Keywords = rs_Store("Meta_Keywords")
		Meta_Title = rs_Store("Meta_Title")
		if Visible <> 0 or isNull(Visible) then
			visible_checked="checked"
		else
			visible_checked = ""
		end if
		if Show_Name <> 0 or isNull(Show_Name) then
			showname_checked="checked"
		else
			showname = ""
		end if
		if Protect_Page <> 0 or isNull(Protect_Page) then
			protect_checked="checked"
		else
			protect_checked = ""
		end if
	rs_Store.close
else
	Response.Redirect "admin_error.asp?message_id=1"
end if

sNeedTabs=1
addPicker=1
sFormAction = "store_dept_edit.asp"
sName = "Edit_store_dept"
sTitle = "Advanced Edit Department - "&Department_Name
sFullTitle = "Inventory > <a href=department_manager.asp class=white>Departments</a> > Advanced Edit - "&Department_Name
sFormName = "Dept_Edit"
sCommonName="Department"
sCancel="department_manager.asp"
thisRedirect = "store_dept_edit.asp"
sSubmitName = "Edit"
sMenu="inventory"
createHead thisRedirect
 %>

<%
	strUpgradeMsg = "This feature is not available at your current level of service.<br><br>"
	strGold = "GOLD"
	strSilver = "SILVER"
	strBronze = "BRONZE"
	strMsg = " Service or higher is required. <a href='billing.asp' class=link>Click here to upgrade now.</a>"
%>

<tr bgcolor='#FFFFFF'>
		<td width="100%" height="1">
                <input type="button" class="Buttons" value="Department List" OnClick='JavaScript:self.location="department_manager.asp"'>
                <input type="button" class="Buttons" value="Basic Department Edit" OnClick='JavaScript:self.location="store_dept_basic_edit.asp?Id=<%=Department_ID%> "'>
                <input class=buttons type="button" value="Preview Dept" name="Editor" OnClick="javascript:goPreview('<%= fn_dept_url(full_name,"") %>')">
			 </td>
</tr>

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
		["Main", "content1",,,"Main","Main"],
		["Content",  "content2",,,"Content","Content"],
		["Options", "content3",,,"Options","Options"],
		["Search Engines", "content4",,,"Search Engines","Search Engines"],
		];

		apy_tabsInit();
		</script>
		</td>
		</tr>
		
		<TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='25'>
		
		
		<!-- CONTENT 1 MAIN -->
			<div id="content1" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>

				
				


				<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><B>Department</B></td>
				<td width="90%" class="inputvalue">
				<input type="text" name="Department_Name" value="<%= Department_Name %>" size="60" maxlength=50>
				<input type="hidden" name="Department_Name_C" value="Re|String|0|50|||Department">
				<input type="hidden" name="Department_ID" value="<%= Department_ID %>">
				<input type="hidden" name="Department_ID_C" value="Re|Integer|||||Department ID">
				<% small_help "Department" %></td>
				</tr>
			
				<% if Service_Type >= 3 then
                  %>
      				<tr bgcolor='#FFFFFF'>
      				<td width="100%" class="inputname" colspan=2><B>Description</B><BR>
      				<%= create_editor ("Department_Description",Department_Description,"") %>
                               <input type="hidden" name="Department_Description_C" value="Op|String|0|2500|||Description">
      						<% small_help "Description" %></td>
      				</tr>
      
      				<tr bgcolor='#FFFFFF'>
      				<td width="10%" class="inputname"><B>Image</B></td>
      				<td width="90%" class="inputvalue">
      						<input type="text" name="Department_Image_Path" value="<%= Department_Image_Path %>" size="60" maxlength=100>
      						<input type="hidden" name="Department_Image_Path_C" value="Op|String|0|100|||Image">
      						<a href="JavaScript:goImagePicker('Department_Image_Path');"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
      						<a class="link" href="JavaScript:goFileUploader('Department_Image_Path');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
      				<% small_help "Image" %></td>
      				</tr>

				<%
				sql_select = "SELECT Department_ID, Department_Name,Full_Name,Belong_to,Last_Level from store_dept where Department_ID <> 0 and Store_id="&Store_id&" order by Full_Name"
				set myfields=server.createobject("scripting.dictionary")
				Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords) %>

				<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><B>Parent</B></td>
				<td width="90%" class="inputvalue">
						<select size="1" name="Belong_to">

						<option selected value="0">Top ( No Parent )</option>

						<% 
						if noRecords = 0 then
						FOR rowcounter= 0 TO myfields("rowcount")
							if Department_ID <> mydata(myfields("department_id"),rowcounter) and Department_ID <> mydata(myfields("belong_to"),rowcounter) then
								sName=mydata(myfields("full_name"),rowcounter)
								if len(sName)>70 then
                                                                   sName=" . . . "&right(sName,67)
                                                                end if
                                                                if Belong_to = mydata(myfields("department_id"),rowcounter) then
									response.write "<option selected value='"&Belong_to&"'>"&sName&"</option>"
								Else
									response.write "<option value='"&mydata(myfields("department_id"),rowcounter)&"'>"&sName&"</option>"
								End If
							end if 
						Next
						end if%>
				</select>
				<% small_help "Parent" %></td>
				</tr>
				</table>
				</div>
				<!-- CONTENT 1 ENDS HERE -->
				
				<!-- CONTENT 2 STARTS HERE -->
				<div id="content2" style="visibility: hidden;" class="tabPage">
				<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>

				<tr bgcolor='#FFFFFF'>
				<td width="100%" class="inputname" colspan=2><B>Department Listings Top HTML</B><BR>
				<%= create_editor ("Department_HTML",Department_HTML,"") %>
                               <input type="hidden" name="Department_HTML_C" value="Op|String|||||Department Listings Top HTML">
							<% small_help "Department_HTML" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
				<td width="100%" class="inputname" colspan=2><B>Department Listings Bottom HTML</B><BR>
				<%= create_editor ("Department_HTML_Bottom",Department_HTML_Bottom,"") %>
                                			<input type="hidden" name="Department_HTML_Bottom_C" value="Op|String|||||Department Listings Bottom HTML">
							<% small_help "Department_HTML_Bottom" %></td>
				</tr>	
				</table>
				</div>
				
				<div id="content3" style="visibility: hidden;" class="tabPage">
				<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>

					<tr bgcolor='#FFFFFF'>
				<td width="20%" class="inputname"><B>Navigation Link Menu</B></td>
				<td width="60%" class="inputvalue">
						<select name=navig_link_menu>
				<option value=<%= navig_link_menu %> selected><%= navig_link_menu_text %></option>
				<option value=0>None</option>
				<option value=1>1</option>
				<option value=2>2</option>
				<option value=3>3</option>
				<option value=4>4</option>
				<option value=5>5</option>
				</select>
				 <% small_help "Navig Link" %></td></tr>
				 <tr bgcolor='#FFFFFF'>
				<td width="20%" class="inputname"><B>Navigation Button Menu</B></td>
				<td width="60%" class="inputvalue">
						<select name=navig_button_menu>
				<option value=<%= navig_button_menu %> selected><%= navig_button_menu_text %></option>
				<option value=0>None</option>
				<option value=1>1</option>
				<option value=2>2</option>
				<option value=3>3</option>
				<option value=4>4</option>
				<option value=5>5</option>
				</select>
				<% small_help "Navig Button" %></td></tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><B>Visible</B></td>
					<td width="90%" class="inputvalue">
					<input type=checkbox class=image name=Visible <%= visible_checked %>>
					<% small_help "Visible" %></td>
					</tr>
					<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><B>Show Name</B></td>
					<td width="90%" class="inputvalue">
					<input type=checkbox class=image name=Show_Name <%= showname_checked %>>
					<% small_help "Show_Name" %></td>
					</tr>

					<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><B>View Order</B></td>
					<td width="90%" class="inputvalue">
					<input type="text" name="View_Order" value="<%= View_Order %>" size="5" onKeyPress="return goodchars(event,'0123456789.')">
					<input type="hidden" name="View_Order_C" value="Re|Integer|0|9999|||View Order">
					<% small_help "View Order" %></td>
					</tr>

				
			
			<% if Service_Type >= 5 then %>
				<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><B>Protect Dept</B></td>
				<td width="90%" class="inputvalue">
							<input type=checkbox class=image name=Protect_Page <%= protect_checked %>>
							<% small_help "Protect" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Customer Group</b></td>
				<td width="60%" class="inputvalue"><select size="1" name="Customer_group">
						<option

									<% if CustomerGroup=0 then %>
										selected
									<% end if %>

							value="0">All Customers</option>
						<%
						sql_groups = "select Group_Name, Group_Id from Store_Customers_Groups where Store_id = "&Store_id&""
						set myfields=server.createobject("scripting.dictionary")
						Call DataGetrows(conn_store,sql_groups,mydata,myfields,noRecords)
						if noRecords = 0 then
						FOR rowcounter= 0 TO myfields("rowcount")

							response.write "<option value='"&mydata(myfields("group_id"),rowcounter)&"'"

								if cint(CustomerGroup)=mydata(myfields("group_id"),rowcounter) then
									response.write "selected"
								end if
							
							response.write ">"&mydata(myfields("group_name"),rowcounter)&"</option>"
						Next
						end if
						set myfields= nothing
						%>
						</select><% small_help "Customer Groups" %></td>
				                    </tr>
				</table>
				</div>
				<!-- CONTENT 3 ENDS HERE -->
				
			<div id="content4" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				
			<tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname"><B>Search Title</b></td>
			<td width="70%" class="inputvalue">
						<input type=text	name="Meta_Title" value="<%= Meta_Title %>" maxlength=100 size=60>
						<input type="hidden" name="Meta_Title_C" value="Op|String|0|100|||Meta Title">
						<% small_help "Meta_Title" %></td>
			</tr>
			
			
				<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname" colspan=2><B>Search Keywords</b>

				<input readonly type=text name=remLenKey size=3 class=char maxlength=3 value="<%= 250-len(Meta_Keywords) %>" class=image><font size=1><I>characters left</i></font><BR>
				<BR><textarea	name="Meta_Keywords" rows=3 cols=83 onKeyDown="textCounter(this.form.Meta_Keywords,this.form.remLenKey,250);" onKeyUp="textCounter(this.form.Meta_Keywords,this.form.remLenKey,250);"><%= Meta_Keywords %></textarea>
				<input type="hidden" name="Meta_Keywords_C" value="Op|String|0|250|||Meta Keywords">
				<% small_help "Meta_Keywords" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname" colspan=2><B>Search Description</b>
				<input readonly type=text name=remLenDesr size=3 class=char maxlength=3 value="<%= 500-len(Meta_Description) %>" class=image><font size=1><I>characters left</i></font><BR>
				<BR><textarea	name="Meta_Description" rows=5 cols=83 onKeyDown="textCounter(this.form.Meta_Description,this.form.remLenDesr,500);" onKeyUp="textCounter(this.form.Meta_Description,this.form.remLenDesr,500);"><%= Meta_Description %></textarea>
				<input type="hidden" name="Meta_Description_C" value="Op|String|0|500|||Meta Description">
				<% small_help "Meta_Description" %></td>
				</tr>	
			</table>
			</div>
			
			
			<% else %>
			<input type="hidden" name="Meta_Keywords" value="">
			<input type="hidden" name="Meta_Description" value="">
			<input type="hidden" name="Meta_Title" value="">
			<input type="hidden" name="Protect" value=0>
			<input type="hidden" name="Customer_group" value=0>
			</table>
				</div>
				<!--	 CONTENT 3 LOOP END -->
				
			<div id="content4" style="height: 0%; visibility: hidden;" class="tabPage">
			<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
			<tr bgcolor='#FFFFFF'>
			<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strSilver%><%=strMsg%></td>
			</tr>
			</table>
			</div>
			<% end if%>
			
<% else %>
<input type="hidden" name="Meta_Keywords" value="">
<input type="hidden" name="Meta_Description" value="">
<input type="hidden" name="Meta_Title" value="">
<input type="hidden" name="View_Order" value="1">
<input type="hidden" name="Visible" value=1>
<input type="hidden" name="Show_Name" value=1>
<input type="hidden" name="Department_HTML" value="">
<input type="hidden" name="Department_HTML_Bottom" value="">
<input type="hidden" name="Belong_to" value="0">
<input type="hidden" name="Department_Image_Path" value="">
<input type="hidden" name="Protect" value=0>
<input type="hidden" name="Customer_Group" value=0>
</table>
</div>

<div id="content2" style="height: 0%; visibility: hidden;" class="tabPage">
<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
<tr bgcolor='#FFFFFF'>
<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strBronze%><%=strMsg%></td>
</tr>
</table>
</div>

<div id="content3" style="height: 0%; visibility: hidden;" class="tabPage">
<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
<tr bgcolor='#FFFFFF'>
<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strBronze%><%=strMsg%></td>
</tr>
</table>
</div>

<div id="content4" style="height: 0%; visibility: hidden;" class="tabPage">
<table width="650" border='0' cellpadding=0 cellspacing=0 class=tpage>
<tr bgcolor='#FFFFFF'>
<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strSilver%><%=strMsg%></td>
</tr>
</table>
</div>				
<!-- CONTENT 1 ENDS IN LOOP -->
<% end if %>
			
			<!-- CONTENT 1 ENDS HERE -->
		
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
 frmvalidator.addValidation("Department_Name","req","Please enter a department name.");
 frmvalidator.addValidation("View_Order","req","Please enter a view order.");
</script>

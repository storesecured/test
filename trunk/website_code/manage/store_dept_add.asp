<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->


<%

sFlashHelp="departments/departments.htm"
sMediaHelp="departments/departments.wmv"
sZipHelp="departments/departments.zip"

'FIRST INSERT THE UN ASSIGNED DEPT INTO STORE_DEPT TABLE (DEPT_ID=0)
sql_check="select Department_ID from store_dept where Department_ID = 0 and Store_id="&Store_id&""
rs_Store.open sql_check,conn_store,1,1 
if rs_Store.Bof = true then
	rs_Store.close
	SQL_INSERT = "INSERT INTO store_dept (store_id,Department_ID,Department_Name,Department_Description,Department_Image_Path,Belong_to,Last_Level) values ("&store_id&",0,'Un Assigned','NA',0,-1,-1)"
	conn_store.Execute SQL_INSERT
else
	rs_Store.close
end if

if Request.Form("Form_Name") = "add_dept" then

	if not CheckReferer then
		Response.Redirect "admin_error.asp?message_id=2"
	end if

	'ERROR CHECKING
	If Form_Error_Handler(Request.Form) <> "" then
		Error_Log = Form_Error_Handler(Request.Form)
		%><!--#include file="Include/Error_Template.asp"--><% 
	else
		Belong_to = Request.Form("Belong_to")
		if isNumeric(Belong_to) then
			Department_Name = checkStringForQ(Request.Form("Department_Name"))
			Department_Image_Path = checkStringForQ(Request.Form("Department_Image_Path"))
			Department_Description = Request.Form("Department_Description")
			Department_HTML = Request.Form("Department_HTML")
			Department_HTML_Bottom = Request.Form("Department_HTML_Bottom")
			Navig_Button_Menu=request.form("navig_button_menu")
			Navig_Link_Menu=request.form("navig_link_menu")
			View_Order = request.form("View_Order")
			Meta_Keywords = checkStringForQ(request.form("Meta_Keywords"))
            Meta_Title = checkStringForQ(request.form("Meta_Title"))
			Meta_Description = checkStringForQ(request.form("Meta_Description"))
            Visible= checkStringForQ(request.form("Visible"))
            if Visible <> "" then
				Visible = 1
			else
				Visible = 0
			end if
			Show_Name= checkStringForQ(request.form("Show_Name"))
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

			'FIND NEW CATEG ID
			sql_max="select max(Department_ID)+1 as max_categ from store_dept where Store_id="&Store_id&""
			rs_Store.open sql_max,conn_store,1,1
				max_categ = rs_store("max_categ")
			rs_Store.close
			if IsNull(max_categ) then
				max_categ = 1
			end if
			'INSERT THE NEW CATEGORY
			SQL_INSERT = "INSERT INTO store_dept (store_id,Department_ID,Department_Name,Department_Description,Department_Image_Path,Belong_to,Last_Level,Department_HTML,Department_HTML_Bottom,View_Order,Meta_Keywords,Meta_Description,Meta_Title,Visible,Protect_Page,Customer_Group,Show_Name,Navig_button_menu,navig_link_menu) values ("&store_id&","&max_categ&",'"&Department_Name&"','"&nullifyQ(Department_Description)&"','"&Department_Image_Path&"',"&Belong_to&",1,'"&nullifyQ(Department_HTML)&"','"&nullifyQ(Department_HTML_Bottom)&"',"&View_Order&",'"&Meta_Keywords&"','"&Meta_Description&"','"&Meta_Title&"', "&Visible&","&Protect_Page&","&Customer_Group&","&Show_Name&","&navig_button_menu&","&navig_link_menu&")"
			conn_store.Execute SQL_INSERT
			
			'UPDATING PARENT CATEGORY
			sql_update="update store_dept set Last_Level = 0 where Department_ID = "&Belong_to&" and Department_ID <> 0 and store_id="&store_id
			conn_store.Execute sql_update
	        Session("Department_List:"&Store_Id)=""
			domainName = lcase(Request.ServerVariables("SERVER_NAME"))
			server.execute "reset_design.asp"
            response.redirect request.form("redirect")
			else
			Response.Redirect "admin_error.asp?message_id=1"
		end if
	end if
end if
if Form_Error_Handler(Request.Form) = "" then
if Request.QueryString("categ_id") <> "" then 
	categ_id = Request.QueryString("categ_id")
	if isNumeric(categ_id) then
		sql_select = "SELECT Department_ID, Department_Name, Full_Name from store_dept where Department_ID =	"&Request.QueryString("categ_id")&" and Store_id="&Store_id&" "
		rs_Store.open sql_select,conn_store,1,1 
			s_dept_id = rs_store( "Department_ID")
			s_dept_name = rs_store( "Department_Name")
			s_full_name = rs_store("Full_Name")
		rs_Store.close
	else
		Response.Redirect "admin_error.asp?message_id=1"
	end if
end if

sNeedTabs=1
addPicker=1
sFormAction = "store_dept_add.asp?categ_id="&Request.QueryString("categ_id")
sName = "Store_Activation"
sFormName = "add_dept"
sTitle = "Advanced Add Department"
sFullTitle = "Inventory > <a href=department_manager.asp class=white>Departments</a> > Advanced Add"
sCommonName="Department"
sCancel="department_manager.asp"
thisRedirect = "store_dept_add.asp"
sMenu = "inventory"
sQuestion_Path = "inventory/departments.htm"
createHead thisRedirect
%>
<!--#include file="include/department_list.asp"-->
<%
sTrialAdd = 1

	  if Service_Type < 10 then
		  Depts = 0
		  sql_select = "Select Count(Department_Id) as Num_Depts from store_dept where Store_Id=" & Store_Id
		  rs_Store.open sql_select,conn_store,1,1
		  if rs_store.bof = false then
			 Depts = rs_Store("Num_Depts")
		  end if
		  rs_store.close
		  Depts = Depts - 1

		  if Depts >= 1 and Service_Type = 0 then
		sTrialAdd = 0
		sLimit = 1
		  elseif Depts >= 10 and Service_Type = 3 then
		sTrialAdd = 0
		sLimit = 10
		  elseif Depts >= 20 and Service_Type = 5 then
		sTrialAdd = 0
		sLimit = 20
		  elseif Depts >= 100 and Service_Type = 7 then
		sTrialAdd = 0
		sLimit = 100
		elseif Depts >= 200 and Service_Type = 9 then
		sTrialAdd = 0
		sLimit = 200
		  end if
	end if
	if sTrialAdd = 0 then %>
		<tr bgcolor='#FFFFFF'>
		<td colspan=2>
			<% if Request.Form("store_dept_add") <> "" then %>
					<b>&nbsp;&nbsp; Added successfully</b>
				<% End If %>
			
			You may not add any more departments, your level of service is limited to <%= sLimit %>, your store currently has <%= Depts %> departments.
			<BR><BR>
			<a href=billing.asp class=link>Click here to upgrade now.</a>

		</td></tr>
	<% createFoot thisRedirect, 0%>
	<% else
		'SELECT STORE DEPARTMENTS
		sql_select = "SELECT Department_ID, Department_Name,Full_Name,Belong_to,Last_Level from store_dept where Department_ID <> 0 and Store_id="&Store_id&" order by Full_Name"
		set myfields=server.createobject("scripting.dictionary")
		Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords) %>



	
		
		<tr bgcolor='#FFFFFF'>
	<td width="100%" height="1">
				<input type="Hidden" name="store_dept_add" value="store_dept_add">
				<input type="button" class="Buttons" value="Department List" OnClick='JavaScript:self.location="department_manager.asp"'>
				<% if Request.Form("store_dept_add") <> "" then %>
					<b>&nbsp;&nbsp; Added successfully</b>
				<% End If %>
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
				<tr bgcolor=#FFFFFF>
			<td width="20%" class="inputname"><B>Department Name</B></td>
			<td width="80%" class="inputvalue">
				<input type="text" name="Department_Name" size="60" maxlength=50>
				<input type="hidden" name="Department_Name_C" value="Re|String|0|50|||Department">
				<% small_help "Department" %></td>
			</tr>
	
			<% if Service_Type >= 3 then %>
			         
                  <tr bgcolor='#FFFFFF'>
			<td width="100%" class="inputname" colspan=2><B>Description</B><BR>
			<%= create_editor ("Department_Description",Department_Description,"") %>
                        <input type="hidden" name="Department_Description_C" value="Op|String|0|2500|||Description">
					<% small_help "Description" %></td>
			</tr>
			  
			<tr bgcolor='#FFFFFF'>
			<td width="20%" class="inputname"><B>Image</B></td>
			<td width="80%" class="inputvalue">
					<input type="text" name="Department_Image_Path" size="60" maxlength=100><a href="JavaScript:goImagePicker('Department_Image_Path');"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
					<input type="hidden" name="Department_Image_Path_C" value="Op|String|0|100|||Image">
					<a class="link" href="JavaScript:goFileUploader('Department_Image_Path');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
			<% small_help "Image" %></td>
			</tr>
  
			<tr bgcolor='#FFFFFF'>
			<td width="20%" class="inputname"><B>Parent</B></td>
			<td width="80%" class="inputvalue">
			<%s_dept_id=0 %>
			<%= create_dept_list ("Belong_to",s_dept_id,1,"0|{}|Top (No Parent)|{new}|") %>
					
			<% small_help "Parent" %></td>
			</tr>
		
			</table>
			</div>
			<!-- CONTENT 1 MAIN -->	
			
			
			<!-- CONTENT 2 MAIN -->
			<div id="content2" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				<tr bgcolor=#FFFFFF>
			<td width="100%" class="inputname" colspan=2><B>Department Listings Top HTML</B><BR>
					<%= create_editor ("Department_HTML",Department_HTML,"") %>
                                       <input type="hidden" name="Department_HTML_C" value="Op|String|||||HTML">
					<% small_help "Department_HTML" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="100%" class="inputname" colspan=2><B>Department Listings Bottom HTML</B><BR>
                 	<%= create_editor ("Department_HTML_Bottom",Department_HTML_Bottom,"") %>
						<input type="hidden" name="Department_HTML_Bottom_C" value="Op|String|||||HTML">
						<% small_help "Department_HTML_Bottom" %></td>
			</tr>	
			</table>
			</div>
			<!-- CONTENT 2 MAIN -->
			
			
			<!-- CONTENT 3 MAIN -->
			<div id="content3" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
			<tr bgcolor='#FFFFFF'>
				<td width="20%" class="inputname"><B>Navigation Link Menu</B></td>
				<td width="60%" class="inputvalue">
						<select name=navig_link_menu>
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
				<option value=0>None</option>
				<option value=1>1</option>
				<option value=2>2</option>
				<option value=3>3</option>
				<option value=4>4</option>
				<option value=5>5</option>
				</select>
				<% small_help "Navig Button" %></td></tr>
			<tr bgcolor=#FFFFFF>
			<td width="20%" class="inputname"><B>View Order</B></td>
			<td width="80%" class="inputvalue">
				<input type="text" name="View_Order" value="0" size="5" onKeyPress="return goodchars(event,'0123456789.')">
				<input type="hidden" name="View_Order_C" value="Re|Integer|0|9999|||View Order">
				<% small_help "View Order" %></td>
			</tr>
			
			<tr bgcolor='#FFFFFF'>
			<td width="20%" class="inputname"><B>Visible</B></td>
			<td width="80%" class="inputvalue">
				<input type=checkbox class=image name=Visible checked>
				<% small_help "Visible" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="20%" class="inputname"><B>Show Name</B></td>
			<td width="80%" class="inputvalue">
				<input type=checkbox class=image name=Show_Name checked>
				<% small_help "Show_Name" %></td>
			</tr>

			<% if Service_Type >= 5 then %>
				<tr bgcolor='#FFFFFF'>
				<td width="20%" class="inputname"><B>Protect Dept</B></td>
				<td width="80%" class="inputvalue">
				<input type=checkbox class=image name=Protect_Page>
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
			<!-- CONTENT 3 MAIN -->
			
			<!-- CONTENT 4 MAIN -->
			<div id="content4" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				<tr bgcolor=#FFFFFF>
			<td width="30%" class="inputname"><B>Search Title</b></td>
			<td width="70%" class="inputvalue">
					<input type=text	name="Meta_Title" value="<%= Meta_Title %>" maxlength=100 size=60>
					<input type="hidden" name="Meta_Title_C" value="Op|String|0|100|||Meta Title">
					<% small_help "Meta_Title" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="10%" class="inputname" colspan=2><B>Search Keywords</b>

					<input readonly type=text name=remLenKey size=3 class=char maxlength=3 value="<%= 250-len(Meta_Keywords) %>" class=image><font size=1><I>characters left</i></font>
			<BR><textarea	name="Meta_Keywords" rows=3 cols=83 onKeyDown="textCounter(this.form.Meta_Keywords,this.form.remLenKey,250);" onKeyUp="textCounter(this.form.Meta_Keywords,this.form.remLenKey,250);"><%= Meta_Keywords %></textarea>
					<input type="hidden" name="Meta_Keywords_C" value="Op|String|0|250|||Meta Keywords">
					<% small_help "Meta_Keywords" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
			<td width="10%" class="inputname" colspan=2><B>Search Description</b>

					<input readonly type=text name=remLenDesr size=3 class=char maxlength=3 value="<%= 500-len(Meta_Description) %>" class=image><font size=1><I>characters left</i></font>
			<BR><textarea	name="Meta_Description" rows=5 cols=83 onKeyDown="textCounter(this.form.Meta_Description,this.form.remLenDesr,500);" onKeyUp="textCounter(this.form.Meta_Description,this.form.remLenDesr,500);"><%= Meta_Description %></textarea>
					<input type="hidden" name="Meta_Description_C" value="Op|String|0|500|||Meta Description">
					<% small_help "Meta_Description" %></td>
			</tr>
			</table>
			</div>
			<!-- CONTENT 4 MAIN -->
			
			<% else %>
			
			<input type="hidden" name="Meta_Keywords" value="">
			<input type="hidden" name="Meta_Description" value="">
			<input type="hidden" name="Meta_Title" value="">
			<input type="hidden" name="Protect" value=0>
			<input type="hidden" name="Customer_Group" value=0>
			</table>
				</div>
			<!--	 CONTENT 3 LOOP END -->
				
				
			<!--	 CONTENT 4 LOOP END -->	
			<div id="content4" style="height: 0%; visibility: hidden;" class="tabPage">
			<table width="827" border='0' cellpadding=0 cellspacing=0 class=tpage>
			<tr bgcolor='#FFFFFF'>
			<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strSilver%><%=strMsg%></td>
			</tr>
			</table>
			</div>
			<!-- content 4 -->
			
			
				
			<% end if %>	
							
		<% else %>
		
		<input type="hidden" name="Meta_Keywords" value="">
		<input type="hidden" name="Meta_Title" value="">
		<input type="hidden" name="Meta_Description" value="">
		<input type="hidden" name="View_Order" value="1">
		<input type="hidden" name="Department_HTML" value="">
		<input type="hidden" name="Department_HTML_Bottom" value="">
		<input type="hidden" name="Belong_to" value="0">
		<input type="hidden" name="Department_Image_Path" value="">
		<input type="hidden" name="Protect" value=0>
		<input type="hidden" name="Visible" value=1>
		<input type="hidden" name="Show_Name" value=1>
		
		</table>
		</div>
		<!-- CONTENT 1 MAIN END LOOP -->
		
		<div id="content2" style="height: 0%; visibility: hidden;" class="tabPage">
		<table width="827" border='0' cellpadding=0 cellspacing=0 class=tpage>
		<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strBronze%><%=strMsg%></td>
		</tr>
		</table>
		</div>

		<div id="content3" style="height: 0%; visibility: hidden;" class="tabPage">
		<table width="827" border='0' cellpadding=0 cellspacing=0 class=tpage>
		<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strBronze%><%=strMsg%></td>
		</tr>
		</table>
		</div>

		<div id="content4" style="height: 0%; visibility: hidden;" class="tabPage">
		<table width="827" border='0' cellpadding=0 cellspacing=0 class=tpage>
		<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan='3'><%=strUpgradeMsg%><%=strSilver%><%=strMsg%></td>
		</tr>
		</table>
		</div>	
		
		<% end if %>
		</td>
		</tr>	
			 
</table></td></tr>					
<tr bgcolor=#FFFFFF>
<td class='tpage'>		
<% createFoot thisRedirect,1 %>
</td>
</tr>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Department_Name","req","Please enter a department name.");
 frmvalidator.addValidation("View_Order","req","Please enter a view order.");
frmvalidator.setAddnlValidationFunction("DoCustomValidation2");

 function DoCustomValidation2()
{
  var frm = document.forms["<%= sName %>"];

	if (frm.Belong_to.options[frm.Belong_to.selectedIndex].value == "")
		{ 
		alert('Please choose a parent department.');
		return false
		}

  return true

}
</script>
<% end if %>
<% end if %>

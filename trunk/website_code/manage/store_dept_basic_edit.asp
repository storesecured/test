<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->
<!--#include virtual="common/common_functions.asp"-->
<!--#include file="help/store_dept_basic_edit.asp"-->


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
		%><!--#include virtual="common/Error_Template.asp"--><%
	       response.end
        else
		'RETRIEVE FORM DATA
		Department_ID = Request.Form("Department_ID")
		Belong_to = Request.Form("Belong_to")

		if isNumeric(Department_ID) and isNumeric(Belong_to) then
			Department_Name = checkStringForQ(Request.Form("Department_Name"))

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
                        sql_update="update store_dept set Last_Level = "&iLast_Level&", Belong_to = '"&Belong_To&"', Department_Name = '"&Department_Name&"' where Department_ID = "&Department_ID&" and Store_id="&Store_id&""
                        conn_store.Execute sql_update

			sql_update="update store_dept set Last_Level = 0 where Department_ID = "&Belong_to&" and Department_ID <> 0 and store_id="&store_id
			conn_store.Execute sql_update

			server.execute "reset_links.asp"
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
	sql_select = "SELECT Full_Name,Meta_Description,Meta_Keywords, Meta_Title, View_Order, Visible, Show_Name,Customer_Group,Protect_Page, Department_ID, Department_HTML, Department_HTML_Bottom, Department_Name,Department_Description,Department_Image_Path,Belong_to,Last_Level from store_dept where Department_ID = "&Department_ID&" and Store_id="&Store_id&" and department_id<>0"
	rs_Store.open sql_select,conn_store,1,1
		Full_Name = rs_store("Full_Name")
		Department_Name = rs_store("Department_Name")
		Department_Image_Path = rs_store("Department_Image_Path")
		Department_Description = rs_store("Department_Description")
		Department_ID = rs_store("Department_ID")
		Department_HTML = rs_Store("Department_HTML")
		Department_HTML_Bottom = rs_Store("Department_HTML_Bottom")
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

sLink = Site_Name&"Browse_dept_items.asp/categ_id/"&Department_ID&"/parent_ids/"&Belong_to

sNeedTabs=1
addPicker=1
sFormAction = "store_dept_basic_edit.asp"
sName = "Edit_store_dept"
sTitle = "Edit Department - "&Department_Name
sFullTitle = "Inventory > <a href=department_manager.asp class=white>Departments</a> > Edit - "&Department_Name
sFormName = "Dept_Edit"
sCommonName="Department"
sCancel="department_manager.asp"
thisRedirect = "store_dept_edit.asp"
sSubmitName = "Edit"
sMenu="inventory"
createHead thisRedirect

	strUpgradeMsg = "This feature is not available at your current level of service.<br><br>"
	strGold = "GOLD"
	strSilver = "SILVER"
	strBronze = "BRONZE"
	strMsg = " Service or higher is required. <a href='billing.asp' class=link>Click here to upgrade now.</a>"
%>

<tr bgcolor='#FFFFFF'>
		<td width="100%" height="1" colspan=3>
                <input type="button" class="Buttons" value="Department List" OnClick='JavaScript:self.location="department_manager.asp"'>
                <input type="button" class="Buttons" value="Advanced Department Edit" OnClick='JavaScript:self.location="store_dept_edit.asp?Id=<%=Department_ID%> "'>
                <input class=buttons type="button" value="Preview Dept" name="Editor" OnClick="javascript:goPreview('<%= fn_dept_url(full_name,"") %>')">
			 </td>
</tr>






				<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><B>Department</B></td>
				<td width="90%" class="inputvalue">
				<input type="text" name="Department_Name" value="<%= Department_Name %>" size="60" maxlength=50>
				<input type="hidden" name="Department_Name_C" value="Re|String|0|50|||Department">
				<input type="hidden" name="Department_ID" value="<%= Department_ID %>">
				<input type="hidden" name="Department_ID_C" value="Re|Integer|||||Department ID">
				<% small_help "Department" %></td>
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


			

		


<% createFoot thisRedirect,1 %>

<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Department_Name","req","Please enter a department name.");
</script>

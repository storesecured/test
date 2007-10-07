<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%

sFlashHelp="departments/departments.htm"
sMediaHelp="departments/departments.wmv"
sZipHelp="departments/departments.zip"

'FIRST INSERT THE UN ASSIGNED DEPT INTO STORE_DEPT TABLE (DEPT_ID=0)
sql_check="select Department_ID from store_dept where Department_ID = 0 and Store_id="&Store_id&""
rs_Store.open sql_check,conn_store,1,1 
if rs_Store.Bof = true then
	rs_Store.close
	SQL_INSERT = "INSERT INTO store_dept (store_id,Department_ID,Department_Name,Belong_to,Last_Level,Navig_button_menu,navig_link_menu) values ("&store_id&",0,'Un Assigned',-1,-1,0,0)"
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
		%><!--#include virtual="common/Error_Template.asp"--><% 
	else
		Belong_to = Request.Form("Belong_to")
		if isNumeric(Belong_to) then
			Department_Name = checkStringForQ(Request.Form("Department_Name"))

			'FIND NEW CATEG ID
			sql_max="select max(Department_ID)+1 as max_categ from store_dept where Store_id="&Store_id&""
			rs_Store.open sql_max,conn_store,1,1
				max_categ = rs_store("max_categ")
			rs_Store.close
			if IsNull(max_categ) then
				max_categ = 1
			end if
			'INSERT THE NEW CATEGORY
			SQL_INSERT = "INSERT INTO store_dept (store_id,Department_ID,Department_Name,Belong_to,Last_Level,navig_button_menu, navig_link_menu) values ("&store_id&","&max_categ&",'"&Department_Name&"',"&Belong_to&",1,1,1)"

            conn_store.Execute SQL_INSERT
			'UPDATING PARENT CATEGORY
			sql_update="update store_dept set Last_Level = 0 where Department_ID = "&Belong_to&" and Department_ID <> 0 and store_id="&store_id
			conn_store.Execute sql_update
			server.execute "reset_design.asp"
            response.redirect "department_manager.asp"
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
sFormAction = "store_dept_basic.asp"
sName = "Store_Activation"
sFormName = "add_dept"
sTitle = "Add Department"
sFullTitle = "Inventory > <a href=department_manager.asp class=white>Departments</a> > Add"
sCommonName="Department"
sCancel="department_manager.asp"
thisRedirect = "store_dept_basic.asp"
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
	<td width="100%" height="1" colspan=3>
				<input type="Hidden" name="store_dept_add" value="store_dept_add">
				<input type="button" class="Buttons" value="Department List" OnClick='JavaScript:self.location="department_manager.asp"'>
				<% if Request.Form("store_dept_add") <> "" then %>
					<b>&nbsp;&nbsp; Added successfully</b>
				<% End If %>
				</td>
	</tr>




		
		<tr bgcolor=#FFFFFF>
			<td width="20%" class="inputname"><B>Department Name</B></td>
			<td width="80%" class="inputvalue">
				<input type="text" name="Department_Name" size="60" maxlength=50>
				<input type="hidden" name="Department_Name_C" value="Re|String|0|50|||Department">
				<% small_help "Department" %></td>
			</tr>


			<tr bgcolor='#FFFFFF'>
			<td width="20%" class="inputname"><B>Parent</B></td>
			<td width="80%" class="inputvalue">
			<%= create_dept_list ("Belong_to",s_dept_id,1,"0|{}|Top (No Parent)|{new}|") %>
					
			<% small_help "Parent" %></td>
			</tr>
		

<% createFoot thisRedirect,1 %>

<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Department_Name","req","Please enter a department name.");
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

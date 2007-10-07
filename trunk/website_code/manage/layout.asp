<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->

<%

sql_select = "Select Department_Layout_Id, Department_Layout from Store_Settings where Store_Id="&Store_Id
rs_Store.open sql_select,conn_store,1,1
if not rs_Store.bof and not rs_Store.eof then
	Department_Layout_Id = rs_Store("Department_Layout_Id")
	Department_Layout = rs_Store("Department_Layout")
end if
rs_Store.close

if Department_Layout_Id = 1 then
	deptchecked1 = "checked"
elseif Department_Layout_Id = 2 then
	deptchecked2 = "checked"
elseif Department_Layout_Id = 3 then
	deptchecked3 = "checked"
elseif Department_Layout_Id = 4 then
	deptchecked4 = "checked"
elseif Department_Layout_Id = 5 then
	deptchecked5 = "checked"
elseif Department_Layout_Id = 6 then
	deptchecked6 = "checked"
elseif Department_Layout_Id = 7 then
	deptchecked7 = "checked"
end if


dim sDeptTemplate(5)
sql_select = "select * from Sys_Inventory_Template where Template_Type=1"
rs_Store.open sql_select,conn_store,1,1

do while not rs_Store.eof
	sTemplate = rs_Store("Template")
	sTemplate = Replace(sTemplate,"OBJ_IMAGE_OBJ","<IMG Src=images/spacer.gif height=40 width=40 border=1>")
	sTemplate = Replace(sTemplate,"OBJ_DEPT_DESC_OBJ","Department description will go here")
	sTemplate = Replace(sTemplate,"OBJ_DEPT_NAME_OBJ","Department Name")
	sDeptTemplate(rs_Store("Template_Id")) = sTemplate
	response.write Template_Id
	rs_Store.movenext
loop
rs_Store.close

sInstructions="Modify the layout for your departments.  This will change the location of departments on your site.	We have provided a variety of layouts to fit everyones needs and for those who know html you may customize the layout to your exact specifications."		
					
addPicker=1
sFormAction = "Store_Settings.asp"
sName = "Store_Dept_Layout"
sFormName = "Store_Dept_Layout"
sTitle = "Department Layout"
sFullTitle = "Inventory > <a href=department_manager.asp class=white>Departments</a> > Layout"
sCommonName="Department Layout"
sCancel="department_manager.asp"
sSubmitName = "Store_Dept_Layout"
thisRedirect = "layout.asp"
sTopic="Store_Dept_Layout"
sMenu="inventory"
sQuestion_Path = "design/dept_layout.htm"
createHead thisRedirect
if Service_Type < 5	then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		SILVER Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0%>

<% else %>
<script>
function showHideRadio(div,field,nest){
  if (document.forms[0].Department_Layout_Id[5].checked == true)
  {
  	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0;
	  obj.display='block'
  }
  	else
  {
  	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0; 
	  obj.display='none'
  }
}
</script>


					<tr bgcolor='#FFFFFF'>
							<td width="1%" valign="top" class="inputvalue"><input class="image" name="Department_Layout_Id" type="radio" value="1" OnClick="showHideRadio('custom','Department_Layout_Id');" <%= deptchecked1 %>>&nbsp;
							</td><td width=90% class="inputvalue"><%= sDeptTemplate(1) %>
							<% small_help "1" %></td>

						</tr>
			 
					<tr bgcolor='#FFFFFF'>
							<td width="1%" valign="top" class="inputvalue"><input class="image" name="Department_Layout_Id" type="radio" value="2" OnClick="showHideRadio('custom','Department_Layout_Id');" <%= deptchecked2 %>>&nbsp;
							</td><td width=90% class="inputvalue"><%= sDeptTemplate(2) %>
							<% small_help "2" %></td>

						</tr>
					<tr bgcolor='#FFFFFF'>
							<td width="1%" valign="top" class="inputvalue"><input class="image" name="Department_Layout_Id" type="radio" value="3" OnClick="showHideRadio('custom','Department_Layout_Id');" <%= deptchecked3 %>>&nbsp;
							</td><td width=90% class="inputvalue"><%= sDeptTemplate(3) %>
							<% small_help "3" %></td>

						</tr>
					<tr bgcolor='#FFFFFF'>
							<td width="1%" valign="top" class="inputvalue"><input class="image" name="Department_Layout_Id" type="radio" value="4" OnClick="showHideRadio('custom','Department_Layout_Id');" <%= deptchecked4 %>>&nbsp;
							</td><td width=90% class="inputvalue"><%= sDeptTemplate(4) %>
							<% small_help "4" %></td>

						</tr>
					<tr bgcolor='#FFFFFF'>
							<td width="1%" valign="top" class="inputvalue"><input class="image" name="Department_Layout_Id" type="radio" value="5" OnClick="showHideRadio('custom','Department_Layout_Id');" <%= deptchecked5 %>>&nbsp;
							</td><td width=90% class="inputvalue"><%= sDeptTemplate(5) %>
							<% small_help "5" %></td>

						</tr>
					<tr bgcolor='#FFFFFF'>
							<td width="1%" valign="top" class="inputvalue"><input class="image" name="Department_Layout_Id" type="radio" value="6" OnClick="showHideRadio('custom','Department_Layout_Id');" <%= deptchecked6 %>>&nbsp;
							</td><td width="8%" class="inputname"><B>Custom</b><BR>
							Design your own department layout.  Do not select this option unless
							you are familiar with HTML, modifying the template incorrectly could cause your pages to look distorted.
							<BR>
							Use the following codes to define your custom html.
										<BR>OBJ_IMAGE_OBJ - Dept Image (if any)
										<BR>OBJ_DEPT_NAME_OBJ - Item Name
										<BR>OBJ_DEPT_DESC_OBJ - Department Description (if any)
										<BR>OBJ_DEPT_COUNT_OBJ - Department Item Count
										<BR>OBJ_DEPT_STATUS_OBJ - Department Status, ie View Items, No Items
							<BR><%= create_editor ("Department_Layout",Department_Layout,"[""Dept Image"",""OBJ_IMAGE_OBJ""],[""Dept Name"",""OBJ_DEPT_NAME_OBJ""],[""Dept Description"",""OBJ_DEPT_DESC_OBJ""]") %>
								<input type="hidden" name="Department_Layout_C" value="Op|String|0|1000|||Department Layout">
							<% small_help "Custom" %></td>

						</tr>

				



<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

</script>
<% end if %>

<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->

<%

Template_Id=Request.Querystring("Id")
If Template_Id = "" then
	fn_error "Please select a template"
end if

sql_select = "Select Nav_Link_HTML, Nav_Button_HTML from Store_Design_template where Store_Id="&Store_Id&" and template_id="&Template_Id
rs_Store.open sql_select,conn_store,1,1
if not rs_Store.bof and not rs_Store.eof then
	Nav_Link_HTML = rs_Store("Nav_Link_HTML")
	Nav_Button_HTML = rs_Store("Nav_Button_HTML")
end if
rs_Store.close
sInstructions="Modify layout for navigation buttons and links.<BR><BR>Use the following code to define your custom html.<BR>%OBJ_LINK_OBJ% will be replaced with the correct link."

addPicker=1
sFormAction = "process_form.asp"
sName = "Store_Nav_Layout"
sFormName = "Store_Nav_Layout"
sTitle = "Navigation Layout"
sSubmitName = "Store_Nav_Layout"
thisRedirect = "nav_layout.asp"
sTopic="Store_Nav_Layout"
sMenu = "design"
sQuestion_Path = "design/nav_layout.htm"
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
				<TR bgcolor='#FFFFFF'>
						 <td width="100%" colspan="5" height="11">
						 <input type="button" OnClick=JavaScript:self.location="template_list.asp" class="Buttons" value="Template List" name="Create_new_Page">&nbsp;<input type="button" OnClick=JavaScript:self.location="layout_design.asp?op=edit&Id=<%=Template_ID%>" class="Buttons" value="Template Detail" name="Create_new_Page"></td>
					</tr>
					<TR bgcolor='#FFFFFF'>
							<td colspan="5" align="center">
								<BR>
								<font color="#ff0000" face="verdana" size="1"><b>Click the Apply Template button to apply this design to your store.</b></font>

							</td>
						</tr>
				<!--Added for bringing the Preview and apply Template buttons-->
				<TR bgcolor='#FFFFFF'>
					<td colspan=5 align="center">
							<table cellspacing="5" cellpadding="0">
									<TR bgcolor='#FFFFFF'>
										<td>	
											<!--<a href="layout_settings.asp?Gen=1&template_id=<%=Template_Id%>">Apply Template</a>-->
											<input type="button" class="buttons" value="Apply Template" onClick="window.location='layout_settings.asp?Gen=1&template_id='+<%=Template_Id%>">							
										</td>
										<td>
											<!--<a href="layout_settings.asp?Gen=1&preview_id=1&template_id=<%=Template_Id%>">Preview Template</a>-->
											<input type="button" class="buttons" value="Preview Template" onClick="window.location='layout_settings.asp?Gen=1&preview_id=1&template_id='+<%=Template_Id%>">
										</td>
									</tr>
								</table>
					</td>
				</TR>
					<input type="hidden" name="Template_Id" value="<%= Template_Id%>">

                  <tr bgcolor='#FFFFFF'>
							<td class="inputname" colspan=4>
								<b>Navigation Button Layout</b><BR>
								<%= create_editor ("Nav_Button_HTML",Nav_Button_HTML,"[""Link Text"",""%OBJ_LINK_OBJ%""]") %>
								<input type="hidden" name="Nav_Button_HTML_C" value="Op|String|0|1500|||Navigation Button">
			<% small_help "Navigation Button" %></td>
						</tr>

						<tr bgcolor='#FFFFFF'>
							<td class="inputname" colspan=4>
								<b>Navigation Link Layout</b><BR>
								<%= create_editor ("Nav_Link_HTML",Nav_Link_HTML,"[""Link Text"",""%OBJ_LINK_OBJ%""]") %>
								<input type="hidden" name="Nav_Link_HTML_C" value="Op|String|0|1500|||Navigation Link">
			<% small_help "Navigation Link" %></td>
						</tr>


<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

</script>
<% end if %>

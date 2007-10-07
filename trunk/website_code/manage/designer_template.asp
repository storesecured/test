<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->



<%

Template_Id=Request.Querystring("Id")
If Template_Id = "" then
	fn_error "Please select a template"
else
	op=Request.Querystring("op")
	sql_select = "Select template_name, template_html, template_head From store_design_template where Store_id = "&Store_id&" and template_id="&Template_Id
	rs_Store.open sql_select,conn_store,1,1
	if not rs_Store.EOF	then
		template_html = Rs_store("template_html")
		template_head = Rs_store("template_head")
		templ_name = Rs_store("Template_Name")	
	end if
	rs_Store.Close
	template_html=replace(replace(template_html,"<textarea","<OBJ_TEXTAREA_START"),"</textarea","<OBJ_TEXTAREA_END")
	template_html=replace(replace(template_html,"<TEXTAREA","<OBJ_TEXTAREA_START"),"</TEXTAREA","<OBJ_TEXTAREA_END")
     template_html=replace(template_html,"OBJ_SWITCH_NAME",Site_Name)
	sArticleHelp="CustomizeLook.htm"
            sFormAction = "Store_Settings.asp"
	sName = "Designer_template"
	sFormName = "Designer_template"
	sTitle = "Edit Template Header Footer - "&templ_name
	sFullTitle = "Design > <a href=template_list.asp class=white>Template</a> > <a href=layout_design.asp?Id="&Template_Id&" class=white>Edit</a> > Header Footer - "&templ_name
	sCommonName = "Header Footer"
	sCancel = "layout_design.asp?Id="&Template_Id
    sSubmitName = "Update"
	thisRedirect = "designer_template.asp"
	addPicker=1
	sMenu = "design"
	sQuestion_Path = "design/header_and_footer.htm"
	createHead thisRedirect

		if Service_Type < 1 then %>
			<TR bgcolor='#FFFFFF'>
			<td colspan=2>
				This feature is not available at your current level of service.<BR><BR>
				PEARL Service or higher is required.
				<a href=billing.asp class=link>Click here to upgrade now.</a>
			</td></tr>
			<% createFoot thisRedirect,0 %>

		<% else %>



				<TR bgcolor='#FFFFFF'>					
					 <td width="100%" colspan="3" height="11">
					 <input type="button" OnClick=JavaScript:self.location="template_list.asp" class="Buttons" value="Template List" name="Create_new_Page">&nbsp;
					 <input type="button" OnClick=JavaScript:self.location="layout_design.asp?Id=<%=Template_ID%>" class="Buttons" value="Template Detail" name="Create_new_Page">


               </td>
				</tr>
				<TR bgcolor='#FFFFFF'>
						<td colspan="3" align="center">
							<BR>
							<font color="#ff0000" face="verdana" size="1"><b>Click the Apply Template button to apply this design to your store.</b></font>

						</td>
					</tr>

		

			<!--Added for bringing the Preview and apply Template buttons-->
			<TR bgcolor='#FFFFFF'>
				<td colspan=3 align="center">
						<table cellspacing="5" cellpadding="0">
								<TR bgcolor='#FFFFFF'>
									<td>
										<input type="button" class="buttons" value="Apply Template" onClick="window.location='layout_settings.asp?Gen=1&template_id='+<%=Template_Id%>">
									</td>
									<td>
										<input type="button" class="buttons" value="Preview Template" onClick="window.location='layout_settings.asp?Gen=1&preview_id=1&template_id='+<%=Template_Id%>">
									</td>
									
								</tr>
							</table>
				</td>
			</TR>
			<!--Added for bringing the Preview and apply Template buttons-->
			

				<TR bgcolor='#FFFFFF'><TD colspan=3 class=instructions>Use this page if you would like to modify the base html of the template you have selected.
				<BR><BR>
				<a href="http://server.iad.liveperson.net/hc/s-7400929/cmd/kbresource/kb-7985615771295727019/view_question!PAGETYPE?sq=variables&sf=101113&sg=0&st=422096&documentid=157238&action=view" class=link target=_blank>Click here to view keywords which can be used in template</a>
				
						
						</td></tr>
						<TR bgcolor='#FFFFFF'>
								<td width="100%" class="inputname"><B>Store HTML Template Head</b><BR>
								<table cellpadding=2 cellspacing=0 border=1><tr><td bgcolor=yellow>The template head section should be modified only by ADVANCED HTML users.  Removing, or changing existing content can cause your store to not function if done incorrectly.  Please proceed with edits to the template head under extreme caution.</td></tr></table>
								  <textarea rows='12' name='template_head' cols='83'><%= template_head %></textarea>
						<% small_help "Store Head" %></td>
						</tr>
						
									<TR bgcolor='#FFFFFF'>
								<td width="100%" class="inputname"><B>Store HTML Template Body</b><BR>
								<% sVariables = "[""Store Name"",""%OBJ_NAME_OBJ%""],[""Path to system image"",""%OBJ_SYS_IMAGES_OBJ%""],[""Path to store image"",""%OBJ_STORE_IMAGES_OBJ%""],[""Navigation Buttons"",""%OBJ_NAV_BUTTONS_OBJ%""],[""Navigation Links"",""%OBJ_NAV_LINKS_OBJ%""],"&_
									"[""System Date"",""%OBJ_DATE_OBJ%""],[""Center Content"",""%OBJ_CENTER_CONTENT_OBJ%""],[""Search Box"",""%OBJ_SEARCH_BOX_OBJ%""],[""Small Cart no total"",""%OBJ_SMALL_CARTNT_OBJ%""],[""Small Cart with total"",""%OBJ_SMALL_CART_OBJ%""],"&_
									"[""Login Box"",""%OBJ_LOGIN_OBJ%""],[""Department Select Box"",""%OBJ_SELECT_BOX_DEPTS_OBJ%""],[""Banner Image"",""%OBJ_BANNER_OBJ%""],[""Users First Name"",""%OBJ_FIRSTNAME_OBJ%""],[""Users Last Name"",""%OBJ_LASTNAME_OBJ%""],"&_
									"[""Top Area Objects"",""%OBJ_TOP_DES_OBJECTS_OBJ%""],[""Top background color"",""%OBJ_TOP_BG%""],[""Top border"",""%OBJ_TOP_BORDER%""],[""Top width"",""%OBJ_TOP_WIDTH%""],[""Top Height"",""%OBJ_TOP_HEIGHT%""],"&_
									"[""Bottom Area Objects"",""%OBJ_BOTTOM_DES_OBJECTS_OBJ%""],[""Bottom background color"",""%OBJ_BOTTOM_BG%""],[""Bottom border"",""%OBJ_BOTTOM_BORDER%""],[""Bottom width"",""%OBJ_BOTTOM_WIDTH%""],[""Bottom Height"",""%OBJ_BOTTOM_HEIGHT%""],"&_
									"[""Left Area Objects"",""%OBJ_LEFT_DES_OBJECTS_OBJ%""],[""Left background color"",""%OBJ_LEFT_BG%""],[""Left border"",""%OBJ_LEFT_BORDER%""],[""Left width"",""%OBJ_LEFT_WIDTH%""],[""Left Height"",""%OBJ_LEFT_HEIGHT%""],"&_
									"[""Right Area Objects"",""%OBJ_RIGHT_DES_OBJECTS_OBJ%""],[""Right background color"",""%OBJ_RIGHT_BG%""],[""Right border"",""%OBJ_RIGHT_BORDER%""],[""Right width"",""%OBJ_RIGHT_WIDTH%""],[""Right Height"",""%OBJ_RIGHT_HEIGHT%""],"&_
									"[""Center Top Area Objects"",""%OBJ_CENTOP_DES_OBJECTS_OBJ%""],[""Center Top background color"",""%OBJ_CENTOP_BG%""],[""Center Top border"",""%OBJ_CENTOP_BORDER%""],[""Center Top width"",""%OBJ_CENTOP_WIDTH%""],[""Center Top Height"",""%OBJ_CENTOP_HEIGHT%""],"&_
									"[""Center Bottom Area Objects"",""%OBJ_CENBOT_DES_OBJECTS_OBJ%""],[""Center Bottom background color"",""%OBJ_CENBOT_BG%""],[""Center Bottom border"",""%OBJ_CENBOT_BORDER%""],[""Center Bottom width"",""%OBJ_CENBOT_WIDTH%""],[""Center Bottom Height"",""%OBJ_CENBOT_HEIGHT%""],"&_
									"[""Center Area Objects"",""%OBJ_CENCEN_DES_OBJECTS_OBJ%""],[""Center background color"",""%OBJ_CENCEN_BG%""],[""Center border"",""%OBJ_CENCEN_BORDER%""],[""Center width"",""%OBJ_CENCEN_WIDTH%""],[""Center Height"",""%OBJ_CENCEN_HEIGHT%""]"
									%>
									<%= create_editor_button ("Store_Header",template_html,sVariables) %>
								  <% small_help "Store Template" %></td>
						</tr>

					<input type="hidden" name="Template_Id" value="<%= Template_Id%>">

				<% createFoot thisRedirect,1 %>
				<SCRIPT language="JavaScript">
				 var frmvalidator  = new Validator(0);

				</script>
		<% end if %>
      
<% end if %>

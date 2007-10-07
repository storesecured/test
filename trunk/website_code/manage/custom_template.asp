<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
op=Request.QueryString("op")
if op="edit" then
	Template_ID = Request.QueryString("ID")
	if not isNumeric(Template_ID) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	
	sql_select="select template_id,template_name,template_description from store_design_template where template_id=" & Template_ID & " and Store_Id=" &Store_Id
	rs_Store.open sql_select,conn_store,1,1
	if not rs_store.eof then
        template_id=rs_store("template_id")
	    template_name=checkStringForQBack(rs_store("template_name"))
	    template_description=rs_store("template_description")
    end if
    rs_Store.close

end if

if op="edit" then
	sTitle = "Edit Template General Info - "&template_name
	sFullTitle = "Design > <a href=template_list.asp class=white>Template</a> > <a href=layout_design.asp?Id="&Template_Id&" class=white>Edit</a> > General Info - "&template_name
        sCancel = "layout_design.asp?Id="&Template_Id
        sCommonName = "General Info"

else
	sTitle = "Add Template"
	sFullTitle = "Design > <a href=template_list.asp class=white>Template</a> > Add"
        sCancel = "template_list.asp"
        sCommonName = "Template"

end if
sFormAction = "custom_template_action.asp"
thisRedirect = "custom_template.asp"
sFormName ="Create_Template"
sMenu = "design"
addPicker=1
createHead thisRedirect
if Service_Type < 1 then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service<BR><BR>
		PEARL Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0%>

<% else %>



<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Template_ID" value="<%=Template_ID%>">

				 
		
			
			<TR bgcolor='#FFFFFF'>
				 <td width="100%" colspan="3" height="11">
				 <input type="button" OnClick=JavaScript:self.location="template_list.asp" class="Buttons" value="Template List" name="Create_new_Page">
				 <% if op="edit" then %>
					 &nbsp;<input type="button" OnClick=JavaScript:self.location="layout_design.asp?op=edit&Id=<%=Template_ID%>" class="Buttons" 	value="Template Detail" name="Create_new_Page">
				<% end if %>
				
				<BR></td>
					
			</tr>

			
		 <% if op="edit" then %>
		 <TR bgcolor='#FFFFFF'>
							<td colspan="3" align="center">
								<BR>
								<font color="#ff0000" face="verdana" size="1"><b>Click the Apply Template button to apply this design to your store.</b></font>

							</td>
						</tr>
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
			 <% end if%>
				


			<TR bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Name</b></td>
				<td width="60%" class="inputvalue">
					<input type="text" name="template_name" size="60" maxlength=250 value="<%= template_name %>">
						<INPUT type="hidden"  name=template_name_C value="Re|String|0|250|||Name">
					<% small_help "Template Name" %></td>
			</tr>

 			<TR bgcolor='#FFFFFF'>
				<td width="40%" class="inputname" colspan=2><B>Description</b><input readonly type=text name=remLenDesTempl size=3 class=char maxlength=3 value="<%= 200-len(template_description) %>" class=image><font size=1><I>characters left</i></font>
					<BR>
					<textarea name="template_description" cols="83" rows="5" onKeyDown="textCounter(this.form.template_description,this.form.remLenDesTempl,200);" onKeyUp="textCounter(this.form.template_description,this.form.remLenDesTempl,200);" ><%= template_description %></textarea>
						<INPUT type="hidden"  name=template_description_C value="Op|String|0|250|||Name">
					<% small_help "Template Description" %></td>
			</tr>

<% createFoot thisRedirect, 1

end if

function checkStringForQBack(theString)
	tmpResult = theString
	tmpResult = replace(tmpResult,"&#8243;","&Prime;")
	tmpResult = replace(tmpResult,"&#8242;","&prime;")
	checkStringForQBack = tmpResult
end function
%>



<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->



<%
on error resume next
Theme_Id = Request.Querystring("Theme")

if Theme_Id <> "" and Theme_Id then
	 'UPDATE CURRENT THEME

	 if cint(Theme_Id)>0 then

		sql_update = "wsp_design_copy "&store_Id&",-1,"& Theme_Id
		conn_store.Execute(sql_update)
        
        sql_update = "wsp_design_apply "&store_Id&",-1"
        conn_store.Execute(sql_update)
		
		 Session("Preview")="0"
         server.execute "reset_design.asp"

		 response.redirect "new_theme_manager.asp?Selected=Yes"&sAddString
	 end if
end if

sql_select="select sys_theme_id from store_design_template where store_id="&store_id&" and last_selected=1"
rs_Store.open sql_select,conn_store,1,1
    sys_theme_id = rs_store("sys_theme_id")
rs_Store.close
' =======================================

sFlashHelp="modifydesign/modifydesign.htm"
sMediaHelp="modifydesign/modifydesign.wmv"
sZipHelp="modifydesign/modifydesign.zip"
sTextHelp="templates/choose_template.doc"

sInstructions="Themes control the look and feel of your website.	You can use the drop down boxes below to search for themes that match certain criteria or just hit Search to view them all."

addPicker = 1
sFormAction = "new_theme_manager.asp"
sTitle = "Choose Existing Template"
sFullTitle = "Design > <a href=template_list.asp class=white>Template</a> > Choose Existing"
sSubmitName = "submit"
thisRedirect = "new_theme_manager.asp"
sMenu = "design"
sQuestion_Path = "design/theme_manager.htm"
createHead thisRedirect
if Service_Type < 1  then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		PEARL Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>

	</td></tr>

<% else %>
<tr bgcolor='#FFFFFF'>
			<td width="100%" colspan="4" height="21">
				<input type="button" class="Buttons" OnClick=JavaScript:self.location="template_list.asp" value="Modify Template" name="Create_New_Group">
			</td>
	</tr>

				<TR bgcolor='#D4DEE5'>
					<td height="18" colspan=4>
						<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
					<tr bgcolor='#FFFFFF'><td class="inputname">Color</td>
					<td class="inputvalue">
					<select size="1" name="Color_id">
			 <% Category_id = Request.Form("Category_id")
			 Color_id = Request.Form("Color_id")
			 Page_size = Request.Form("Page_size")
			 sql_color = "SELECT * FROM Sys_template_color_id;"
			 set myfields=server.createobject("scripting.dictionary")
			 Call DataGetrows(conn_store,sql_color,mydata,myfields,noRecords) %>
				<Option selected value="0">Any</option>
				<% if noRecords = 0 then %>
				<% FOR rowcounter= 0 TO myfields("rowcount") %>

					<Option value=<%= mydata(myfields("id"),rowcounter) %>> <%= mydata(myfields("color"),rowcounter) %></option>

				<% Next %>
				<% end if %>
			</select> 
					<% small_help "Color" %></td></tr>


					<tr bgcolor='#FFFFFF'><td class="inputname">Category</td>
					<td class="inputvalue">
					<select size="1" name="Category_id">
				<% sql_category = "SELECT * FROM Sys_category;" %>
			<% set myfields=server.createobject("scripting.dictionary")
			 Call DataGetrows(conn_store,sql_category,mydata,myfields,noRecords) %>
				<Option selected value="0">Any</option>
			  <% if noRecords = 0 then %>
				<% FOR rowcounter= 0 TO myfields("rowcount") %>
					<Option value="<%= mydata(myfields("category_id"),rowcounter) %>"> <%= mydata(myfields("category"),rowcounter) %></option>
				<% Next %>
				<% end if %>
			</select>
			<% small_help "Category" %></td></tr>

			<input type=hidden name="Page_size" value=0>

			<tr bgcolor='#FFFFFF'>
						<td valign="top" colspan=3 align=center><input type="submit" class="Buttons" value="Search Templates" name="<%= sSubmitName %>"></td>
					</tr></table>
					</td></tr>


						<tr bgcolor='#FFFFFF'><td>
						<% if Request.Form <> "" then

						 sql_theme_html = "exec wsp_template_search "&category_id&","&color_id&";"
 
						 set myfields=server.createobject("scripting.dictionary")
						 Call DataGetrows(conn_store,sql_theme_html,mydata,myfields,noRecords)
								if noRecords=1 then %>
								<tr bgcolor='#FFFFFF'><td>No matching records were found, try loosening your search criteria and searching again.</td></tr>
								<% else %>
								<% FOR rowcounter= 0 TO myfields("rowcount") %>
								<tr bgcolor='#FFFFFF'>
									<td valign="top" width=1%>
									<img src="images/themes/<%= mydata(myfields("theme_id"),rowcounter) %>.gif" border=1>
									</td>
									<td><table><tr bgcolor='#FFFFFF'><td>
									<font size="1">
									<a class="link" href="javascript:goShowPicture('<%= mydata(myfields("theme_id"),rowcounter) %>')">Preview</a><br>
									<% if cint(Store_Theme) = cint(mydata(myfields("theme_id"),rowcounter)) then %>
									<a class="link" href=new_theme_manager.asp?Theme=<%= mydata(myfields("theme_id"),rowcounter) %><%= sAddString %>>SELECTED</a></font>
									<% else %>
									<a class="link" href=new_theme_manager.asp?Theme=<%= mydata(myfields("theme_id"),rowcounter) %><%= sAddString %>>Select</a></font>
									<% end if %>
									</td></tr></table></td>
									<% rowcounter = rowcounter + 1 %>


									<% if rowcounter <= myfields("rowcount") then %>
										<td valign="top" width=1%>
										<img src="images/themes/<%= mydata(myfields("theme_id"),rowcounter) %>.gif" border=1>
										</td>
										<td><table><tr bgcolor='#FFFFFF'><td>
										<font size="1">
										<a class="link" href="javascript:goShowPicture('<%=mydata(myfields("theme_id"),rowcounter) %>')">Preview</a><br>
										<% if cint(Store_Theme) = cint(mydata(myfields("theme_id"),rowcounter)) then %>
											<a class="link" href=new_theme_manager.asp?Theme=<%= mydata(myfields("theme_id"),rowcounter) %><%= sAddString %>>SELECTED</a></font>
										<% else %>
											<a class="link" href=new_theme_manager.asp?Theme=<%= mydata(myfields("theme_id"),rowcounter) %><%= sAddString %>>Select</a></font>
										<% end if %>
										</td></tr></table></td>
									<% else %>
									<td colspan=2>&nbsp;</td>
									<% end if %>
								</tr>
							<% Next %>
							
							 </td></tr>
							
							 <% end if %>
					<% end if %>
						
<tr bgcolor='#FFFFFF'><td colspan=4>
<BR><BR>All of the templates are completely customizable.  Once you choose a template <a href=template_list.asp class=link>click here to modify it</a>.
<% end if %>
<% createFoot thisRedirect, 2%>



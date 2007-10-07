<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


	'CHECK THE MODE IF EDIT GET THE VALUES

	if request.querystring("op") = "edit" then

		sql_select = "select page_name, page_description, file_name, navig_button_menu, navig_link_menu, view_order from store_pages where is_link=1 and page_id = " &  request.querystring("id")  & " and store_id =" & store_id
		set rs_page   = server.createobject("ADODB.Recordset")
		rs_page.open sql_select, conn_store, 2,3
			if not rs_page.eof  then
				page_name = rs_page.fields("page_name")
				file_name = rs_page.fields("file_name")
				page_description = rs_page.fields("page_description")
				view_order = rs_page.fields("view_order")
				Navig_Button_Menu=rs_page.fields("navig_button_menu")
			    if Navig_Button_Menu=0 then
				    Navig_Button_Menu_Text = "None"
			    else
				    Navig_Button_Menu_Text = Navig_Button_Menu
			    end if
			    Navig_Link_Menu=rs_page.fields("navig_link_menu")
			    if Navig_Link_Menu=0 then
				    Navig_Link_Menu_Text = "None"
			    else
			         Navig_Link_Menu_Text = Navig_Link_Menu
			    end if
				page_id = request.querystring("id")
				op = "edit"
			end if
		rs_page.close
		set rs_page= nothing
	else
        navig_link_menu=0
        navig_link_menu_text=0
        navig_button_menu=0
        navig_button_menu_text=0
        view_order=0
	end if

if request.querystring("op") = "edit" then
	sFullTitle ="Design > <a href=page_links_manager.asp class=white>Links</a> > Edit - "&page_name
        sTitle = "Edit Link - "&page_name
else
	sFullTitle ="Design > <a href=page_links_manager.asp class=white>Links</a> > Add"
        sTitle = "Add Link"
end if

sCommonName="Link"
sCancel="page_links_manager.asp"
sFormAction = "new_page_link_action.asp"
thisRedirect = "new_page_link.asp"
sMenu = "design"
createHead thisRedirect


%>
<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="page_id" value="<%=page_id%>">
<input type="hidden" name="Navig_Button_Menu_Old" value="<%=Navig_Button_Menu%>">
<input type="hidden" name="Navig_Link_Menu_Old" value="<%=Navig_Link_Menu%>">
<input type="hidden" name="Add_page_link" value="Save">
	 


						
				<tr bgcolor='#FFFFFF'>
						<td width="40%" class="inputname"><b>Name</b></td>
						<td width="60%" class="inputvalue">
						<input type="text" name="page_name" size="60" value="<%= page_name %>" maxlength=250>
							<INPUT type="hidden"  name=page_name_C value="Re|String|0|250|||Name">
                                   <% small_help "Link Name" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
						<td width="40%" class="inputname"><b>URL</b></td>
						<td width="60%" class="inputvalue">
						<input type="text" name="file_name" size="60" value="<%= file_name %>" maxlength=250>
							<INPUT type="hidden"  name=file_name_C value="Re|String|0|250|||Name">
                                   <% small_help "URL" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
						<td width="40%" class="inputname"><b>Description</b></td>
						<td width="60%" class="inputvalue">
							<input type=text name="page_description" value="<%= page_description %>" size=60>
							<INPUT type="hidden"  name=page_description_C value="Op|String|0|250|||Description"></td>
                                   <% small_help "Description" %></td>
				</tr>

				
				<tr bgcolor='#FFFFFF'>
						<td width="40%" class="inputname"><b>View Order</b></td>
						<td width="60%" class="inputvalue">
						<input type=text name="view_order" value="<%= view_order %>" size=3>
						<INPUT type="hidden"  name=view_order_C value="Re|Integer|||||View Order">
                         <% small_help "View Order" %></td>
				</tr>


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


					

<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("page_name","req","Please enter a Link's name");
 frmvalidator.addValidation("file_name","req","Please enter a Link or URL");
 frmvalidator.addValidation("view_order","req","Please enter a view order.");

</script>

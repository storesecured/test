<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->
<!--#include file="help/new_page.asp"-->
<%
View_Order=0
Page_Id=request.Querystring("Id")

if Page_Id<>"" then
	op="edit"
	if not isNumeric(Page_Id) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	'IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlPages="select center_area,use_template,protect_page, customer_group, Meta_Title, Meta_Keywords, "&_
	    "Meta_Description, p.File_Name, Page_Name, Page_Description, Navig_Button_Menu, Navig_Link_Menu, "&_
	    "Allow_Link, View_Order, Page_Head, Page_Top, Page_Bottom "&_
	    " from store_pages p WITH (NOLOCK) LEFT JOIN sys_pages sp on p.file_name=sp.file_name "&_
	    " where p.store_id=" & store_id & "AND p.Page_Id=" & Page_Id 

	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sqlPages,mydata,myfields,noRecords)

	if noRecords = 0 then
		FOR rowcounter= 0 TO myfields("rowcount")
			Page_Name=mydata(myfields("page_name"),rowcounter)
			Center_Area=mydata(myfields("center_area"),rowcounter)
			if isnull(center_area) then
				Center_Area="No predefined content"
			end if
			File_Name=mydata(myfields("file_name"),rowcounter)
			Page_Description=mydata(myfields("page_description"),rowcounter)
			Navig_Button_Menu=mydata(myfields("navig_button_menu"),rowcounter)
			if Navig_Button_Menu=0 then
				Navig_Button_Menu_Text = "None"
			else
				Navig_Button_Menu_Text = Navig_Button_Menu
			end if
			Navig_Link_Menu=mydata(myfields("navig_link_menu"),rowcounter)
			if Navig_Link_Menu=0 then
				Navig_Link_Menu_Text = "None"
			else
			     Navig_Link_Menu_Text = Navig_Link_Menu
			end if
			Allow_Link=mydata(myfields("allow_link"),rowcounter)
			View_Order=mydata(myfields("view_order"),rowcounter)
			Meta_Title=mydata(myfields("meta_title"),rowcounter)
			Meta_Description=mydata(myfields("meta_description"),rowcounter)
			Meta_Keywords=mydata(myfields("meta_keywords"),rowcounter)
			Protect_page=mydata(myfields("protect_page"),rowcounter)
			Use_Template=mydata(myfields("use_template"),rowcounter)
			CustomerGroup=mydata(myfields("customer_group"),rowcounter)
			Page_Head=mydata(myfields("page_head"),rowcounter)
			Page_Content_Top=replace(replace(mydata(myfields("page_top"),rowcounter),"<textarea","<OBJ_TEXTAREA_START"),"</textarea","<OBJ_TEXTAREA_END")
			Page_Content_Top=replace(replace(Page_Content_Top,"<TEXTAREA","<OBJ_TEXTAREA_START"),"</TEXTAREA","<OBJ_TEXTAREA_END")
			Page_Content_Bottom=replace(replace(mydata(myfields("page_bottom"),rowcounter),"<textarea","<OBJ_TEXTAREA_START"),"</textarea","<OBJ_TEXTAREA_END")
			Page_Content_Bottom=replace(replace(Page_Content_Bottom,"<TEXTAREA","<OBJ_TEXTAREA_START"),"</TEXTAREA","<OBJ_TEXTAREA_END")

		Next
		set myfields=nothing
	else
		set myfields=nothing
		Response.Redirect "admin_error.asp?message_id=1"
	end if

	if Protect_Page<>0 then
		protected_checked = "checked"
	else
		protected_checked = ""
	end if
	if Use_Template<>0 then
		template_checked = "checked"
	else
		template_checked = ""
	end if
else
    navig_link_menu=0
    navig_link_menu_text=0
    navig_button_menu=0
    navig_button_menu_text
    template_checked="checked"
end if


if op="edit" then
	sFullTitle = "Design > <a href=page_manager.asp class=white>Pages</a> > Edit - "&Page_Name
        sTitle = "Edit Page - "&Page_Name
else
	sFullTitle = "Design > <a href=page_manager.asp class=white>Pages</a> > Add"
        sTitle = "Add Page"

end if

sNeedTabs=1
sFormAction = "Page_action.asp"
thisRedirect = "new_page.asp"
sFormName ="Create_Page"
sCommonName="Page"
sCancel="page_manager.asp"
sMenu = "design"
addPicker=1
createHead thisRedirect
if Trial_Version and op<>"edit" then %>
       <tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available for free trial users.  
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0%>

<% else %>
	

	<!-- TAB MENU STARTS HERE -->

	<% if op="edit" then %>
	<tr bgcolor='#FFFFFF'>
	<td width="100%" height="11">

		<input class=buttons type="button" name="copy_page" value ="Copy Page" onclick=javascript:self.location="copy_page.asp?page_Id=<%= page_id %>&file_name=<% response.write server.urlencode(file_name)%>">
		<input class=buttons type="button" value="Preview Page" name="Editor" OnClick="javascript:goPreview('<%= Site_Name %><%= File_Name %>')">
			    

	</td></tr>
	<% end if %>

<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Navig_Button_Menu_Old" value="<%=Navig_Button_Menu%>">
<input type="hidden" name="Navig_Link_Menu_Old" value="<%=Navig_Link_Menu%>">

<input type="hidden" name="Page_Id" value="<%=Page_Id%>">

<tr bgcolor='#FFFFFF'><td width="724" align=center valign=top height=35>
<table border=0 cellspacing=0 cellpadding=0 width=724>

			 
	<!-- TAB MENU STARTS HERE -->

		<tr bgcolor='#FFFFFF'>
		<td align="center" valign=top height=45 width='100%'><br>
		<script type="text/javascript" language="JavaScript1.2" src="include/tabs-xp.js"></script>
		<script language="javascript1.2">
		var bselectedItem   = 0;
		var bmenuItems =
		[
		["Main", "content1",,,"Main","0"],
		["Options", "content3",,,"Options","0"],
		["Search Engines", "content4",,,"Search Engines","0"],
		];

		apy_tabsInit();
		</script>
		</td>
		</tr>
		
		<TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='25'>
		
		<!-- CONTENT 1 MAIN -->
			<div id="content1" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				
		<% if op="edit" then %>
			<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Page Link</b></td>
				<td width="60%" class="inputvalue" colspan=2>
            <%= Site_Name %><%= File_Name %>
            </td>
			</tr>
		<% end if %>
			<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Name</b></td>
				<td width="60%" class="inputvalue">
					<input type="text" name="Page_name" size="60" maxlength=250 value="<%= Page_Name %>">
						<INPUT type="hidden"  name=Page_name_C value="Re|String|0|250|||Name">
					<% small_help "Name" %></td>
			</tr>
	 
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Description</B></td>
				<td width="60%" class="inputvalue">
					<input type="text" name="Page_Description" size="60" maxlength=250 value="<%= Page_Description %>">
						<INPUT type="hidden"  name=Page_Description_C value="Op|String|0|250|||Description">
					<% small_help "Description" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Filename</B></td>
				<td width="60%" class="inputvalue">
				<% if page_id="" or page_id>=50 then %>
					<input onKeyPress="return goodchars(event,'0123456789_abcdefghijklmnopqrstuvwxyz-')" type="text" name="File_name" size="60" maxlength=250 value="<%= replace(File_name,".asp","") %>">.asp
				        <INPUT type="hidden"  name=File_name_C value="Re|String|0|250|||File_name">
				<% else %>
				    <%=File_name %><INPUT type="hidden"  name=File_name value="<%= replace(File_name,".asp","") %>">
				<% end if  %>
                                
				<% small_help "File_name" %></td>
			</tr>
			<tr bgcolor='#FFFFFF'><td class="inputname" colspan=2><B>Page Content Top</b><BR>
       <%= create_editor ("Page_Content_Top",Page_Content_Top,"") %>
				<% small_help "Content" %></td></tr>
				<tr bgcolor='#FFFFFF'><td colspan=3 class=instructions>Below is a short description of the content automatically
                                shown in the center of this page.  Please note
                                that you do not have the ability to directly edit this content.  You 
                                may however add your own content above or below the existing content 
                                using the page content top or page content bottom fields.</td></tr>
                                <tr bgcolor='#FFFFFF'><td class=inputname><B>Center</b></td><td class=inputvalue colspan=2><%= Center_area %></td></tr>


				<tr bgcolor='#FFFFFF'><td class="inputname" colspan=2><B>Page Content Bottom</b><BR>
				<%= create_editor ("Page_Content_Bottom",Page_Content_Bottom,"") %>
				<% small_help "Content" %></td></tr>
			</table>
			</div>
			<!-- CONTENT 1 ENDS HERE -->
			
			
			<!-- CONTENT 3 STARTS HERE -->
			<div id="content3" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
			<% if Allow_Link or Page_Id = "" then %>
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
				<td width="40%" class="inputname"><B>View Order</B></td>
				<td width="60%" class="inputvalue">
					<input name="View_Order" size="3" value="<%= View_Order %>"  onKeyPress="return goodchars(event,'0123456789-.')">
						<INPUT type="hidden"  name="View_Order_C" value="Re|Integer|||||View Order">
					<% small_help "View Order" %></td></tr>
					<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Use Template</B></td>
				<td width="60%" class="inputvalue">
					<input class=image name="Use_Template" value="1" <%= template_checked %> type=checkbox>
					<% small_help "Use_Template" %></td></tr>
					<% if File_Name <> "check_out.asp" then %>
          <tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Protect</B></td>
				<td width="60%" class="inputvalue">
					<input class=image name="Protect_Page" value="1" <%= protected_checked %> type=checkbox>
					<% small_help "Protect" %></td></tr>

					<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Customer Group</b></td>
				<td width="60%" class="inputvalue"><select size="1" name="Customer_group">
						<option
								 <% if op="edit" then %>
									<% if CustomerGroup=0 then %>
										selected
									<% end if %>
								<% else %>
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
							if op="edit" then
								if cint(CustomerGroup)=mydata(myfields("group_id"),rowcounter) then
									response.write "selected"
								end if
							end if
							response.write ">"&mydata(myfields("group_name"),rowcounter)&"</option>"
						Next
						end if
						set myfields= nothing
						%>
						</select><% small_help "Customer Groups" %></td>
				</tr>
					<% end if %>
               <% else %>
					<tr bgcolor='#FFFFFF'><td>This is an internal system page which cannot be added to nav buttons or links.</td></tr>
                                        <input type=hidden name=Use_Template value=1>
                                        <input type=hidden name=View_Order value=0>
                                        <input type=hidden name=Customer_Group value=0>
                                        <input type=hidden name=Protect_page value="">
                                        <input type=hidden name=Navig_button_menu value=0>
                                        <input type=hidden name=Navig_link_menu value=0>
					<% end if %>
			</table>
			</div>
			<!-- CONTENT 3 ENDS HERE -->
			
			
			<!-- CONTENT 4 STARTS HERE -->
            <div id="content4" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>	

            
			
				
				   <tr bgcolor='#FFFFFF'>
			<td width="30%" class="inputname"><B>Search Title</b></td>
			<td width="70%" class="inputvalue">
						<input type=text	name="Meta_Title" value="<%= Meta_Title %>" maxlength=100 size=60>
						<input type="hidden" name="Meta_Title_C" value="Op|String|0|100|||Meta Title">
						<% small_help "Meta_Title" %></td>
			</tr>
				<tr bgcolor='#FFFFFF'><td class="inputname" colspan=2><B>Search Keywords</b>
				<input readonly type=text name=remLenKey size=3 class=char maxlength=3 value="<%= 350-len(Meta_Keywords) %>" class=image><font size=1><I>characters left</i></font>
        <BR><textarea rows="3" name="Meta_Keywords" cols="83"  onKeyDown="textCounter(this.form.Meta_Keywords,this.form.remLenKey,350);" onKeyUp="textCounter(this.form.Meta_Keywords,this.form.remLenKey,350);"><%= Meta_Keywords %></textarea>
				<input type="hidden" name="Meta_Keywords_C" value="Op|String|0|350|||Keywords">
            <% small_help "Meta_Keywords" %></td></tr>

				<tr bgcolor='#FFFFFF'><td class="inputname" colspan=2><B>Search Description</b>
				<input readonly type=text name=remLenDes size=3 class=char maxlength=3 value="<%= 500-len(Meta_Description) %>" class=image><font size=1><I>characters left</i></font>
        <BR><textarea rows="5" name="Meta_Description" cols="83"  onKeyDown="textCounter(this.form.Meta_Description,this.form.remLenDes,500);" onKeyUp="textCounter(this.form.Meta_Description,this.form.remLenDes,500);"><%= Meta_Description %></textarea>
				<input type="hidden" name="Meta_Description_C" value="Op|String|0|500|||Keywords">
            <% small_help "Meta_Description" %></td></tr>
            <tr bgcolor='#FFFFFF'><td class="inputname" colspan=2><B>Page Head</b>
				<input readonly type=text name=remLenHead size=3 class=char maxlength=3 value="<%= 4000-len(Page_Head) %>" class=image><font size=1><I>characters left</i></font><br>
        <BR><textarea rows="5" name="Page_Head" cols="83"  onKeyDown="textCounter(this.form.Page_Head,this.form.remLenHead,4000);" onKeyUp="textCounter(this.form.Page_Head,this.form.remLenHead,4000);"><%= Page_Head %></textarea>
				<input type="hidden" name="Page_Head_C" value="Op|String|0|4000|||Page Head">
            <% small_help "Page_Head" %></td></tr>



            
			</table>
			</div>	
			<!-- CONTENT 4 ENDS HERE -->
</td>
		</tr>	
			 
</table></td></tr>	
<tr bgcolor='#FFFFFF'>
<td colspan='4' class='tpage'>            
<% createFoot thisRedirect, 1%>
</td>
</tr>

<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("View_Order","req","Please enter a view order.");
 frmvalidator.addValidation("Page_name","req","Please enter a page name.");
 frmvalidator.addValidation("Meta_Keywords","maxlen=350","The keyword length cannot exceed 350 characters.");
 frmvalidator.addValidation("Meta_Description","maxlen=500","The description length cannot exceed 500 characters.");

</script>
<% end if %>

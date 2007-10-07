<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->

<%

sql_select = "select page_top from store_pages where store_id="&store_id&" and Page_Id=4"
rs_Store.open sql_select,conn_store,1,1
Store_Close_Message = rs_store("page_top")
rs_Store.close

if Store_active = -1 then
	checked1 = "checked"
else
	checked0 = "checked"
end if

sFlashHelp="activatestore/activatestore.htm"
sMediaHelp="activatestore/activatestore.wmv"
sZipHelp="activatestore/activatestore.zip"
sInstructions="Close store for maintenance, or any other time when you would like to suspend orders.  When your store is closed you will still be able to login to the admin area and make changes however you will not be allowed to preview those changes."
sTextHelp="activation/open_close_store.doc"


addPicker=1
sLevelRequired=5
sFormAction = "Store_Settings.asp"
sName = "Store_Activation"
sFormName = "activation"
sTitle = "Open/Close Store"
sFullTitle = "General > Open/Close Store"
sSubmitName = "Store_Activation_Update"
thisRedirect = "activation.asp"
sTopic="Activate"
sMenu = "general"
sQuestion_Path = "advanced/activate_store.htm"
createHead thisRedirect
%>

					<TR bgcolor='#FFFFFF'>
							<td width="7%" class="inputname"><b>Open</b></td>
							<td width="93%" class="inputvalue"><input class="image" name="Store_active" type="radio" value="-1" <%= checked1 %>>&nbsp;
						<% small_help "Open" %></td></tr>

					<TR bgcolor='#FFFFFF'>
							<td width="7%" class="inputname"><b>Closed</b></td>
							<td width="93%" class="inputvalue"><input class="image" name="Store_active" type="radio" value="0" <%= checked0 %>>&nbsp;
							<% small_help "Closed" %></td>
						
						</tr>
			 
					<TR bgcolor='#FFFFFF'><TD class="inputvalue" colspan=2><b>Closed Message</b><BR>
								<%= create_editor ("Store_Close_Message",Store_Close_Message,"") %>
								<input type="hidden" name="Store_Close_Message_C" value="Op|String|||||Closed Message">
			<% small_help "Closed Message" %></td>
						</tr>

<% createFoot thisRedirect, 1%>

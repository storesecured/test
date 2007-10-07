<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->
<!--#include file="help/homepage.asp"-->

<%

sql_select = "select page_top from store_pages where store_id="&Store_id&" and Page_Id=5"
rs_Store.open sql_select,conn_store,1,1
sHomePage = rs_store("page_top")
rs_Store.close

addPicker=1
sFormAction = "Store_Settings.asp"
sName = "Store_Home"
sCommonName = "Homepage"
sFormName = "homepage"
sTitle = "Homepage"
sFullTitle = "General > Homepage"
sSubmitName = "store_home"
thisRedirect = "homepage.asp"
sTopic = "Home"
sMenu = "general"
createHead thisRedirect
%>

					<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Default Page URL</td>
				<td width="70%" class="inputvalue"><input type="text" name="Store_Homepage" value="<%= Store_Homepage %>" size=60>&nbsp;
				<input type="hidden" name="Store_Homepage_C" value="Op|String|0|255|||Default Page"><% small_help "Store_Homepage" %></td>
				</tr>
				
                                 <TR bgcolor='#FFFFFF'>
				<td colspan=3 align=center><BR><B>OR</b><BR></td>
				</tr>
				<TR bgcolor=FFFFFF>
							<td class="inputvalue" colspan=2><b>Store Homepage</b><BR>
								<%= create_editor ("sHomePage",sHomePage,"") %>
								<input type="hidden" name="sHomePage_C" value="Op|String|||||Homepage">
							 <% small_help "Store Homepage" %>
						</td></tr>

					



<% createFoot thisRedirect, 1%>

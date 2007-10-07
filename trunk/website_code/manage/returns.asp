<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->
<!--#include file="help/returns.asp"-->


<%

sql_select = "select page_top from store_pages where store_id="&Store_id&" and Page_Id=14"
rs_Store.open sql_select,conn_store,1,1
Return_Policy = rs_store("page_top")
rs_Store.close

if enable_rma <> 0 then
	Enable_RMA_Checked = "checked"
else
	Enable_RMA_Checked = ""
end if
if auto_rma_days="" then
	auto_rma_days=0
end if

addPicker=1
sFormAction = "Store_Settings.asp"
sName = "Store_Returns"
sFormName = "returns"
sCommonName="Return Policy"
sTitle = "Return Policy"
sFullTitle = "General > Return Policy"
sSubmitName = "Store_Returns"
thisRedirect = "returns.asp"
sTopic = "Returns"
sMenu = "general"
createHead thisRedirect
%>

					<TR bgcolor='#FFFFFF'>
				<td width="35%" class="inputname">Automatic RMA</td>
				<td width="35%" class="inputvalue">Enabled <input class="image" type="checkbox" name="Enable_RMA" value="-1" <%= Enable_RMA_Checked %>>
				</td><td width="35%" class="inputvalue">If Within <input type="text" name="Auto_RMA_Days" value="<%= Auto_RMA_Days %>" size=3 onKeyPress="return goodchars(event,'0123456789')"> Days</DIV>
				<input type="hidden" name="Enable_RMA_C" value="Op|Integer|||||Auto Return Days">
				<% small_help "Enable Automatic RMA" %></td>
				</tr>
				
				<TR bgcolor=FFFFFF>
							<td class="inputvalue" colspan=3><b>Store Return Policy</b><BR>
								<%= create_editor ("Return_Policy",Return_Policy,"") %>
								<input type="hidden" name="Return_Policy_C" value="Op|String|||||Returns">
							 <% small_help "Store Return Policy" %>
						</td></tr>

					



<% createFoot thisRedirect, 1%>

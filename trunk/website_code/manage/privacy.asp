<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->

<%

sql_select = "select page_top from store_pages where store_id="&Store_id&" and Page_Id=18"
rs_Store.open sql_select,conn_store,1,1
Privacy_Policy = rs_store("page_top")
rs_Store.close

sFlashHelp="returnprivacy/returnprivacy.htm"
sMediaHelp="returnprivacy/returnprivacy.wmv"
sZipHelp="returnprivacy/returnprivacy.zip"
sInstructions="A privacy policy tells your customers what you will do with their sensitive information."
sTextHelp="returnprivacy/privacy_policy.doc"


addPicker=1
sFormAction = "Store_Settings.asp"
sName = "Store_Privacy"
sFormName = "privacy"
sCommonName="Privacy Policy"
sTitle = "Privacy Policy"
sFullTitle = "General > Privacy Policy"
sSubmitName = "Store_Privacy"
thisRedirect = "privacy.asp"
sTopic = "Privacy"
sMenu = "general"
sQuestion_Path = "general/privacy_policy.htm"
createHead thisRedirect
%>


						 <TR bgcolor=FFFFFF>
							<td class="inputvalue" colspan=2><b>Store Privacy Policy</b><BR>
								 
								<%= create_editor ("Privacy_Policy",Privacy_Policy,"") %>
								<input type="hidden" name="Privacy_Policy_C" value="Op|String|||||Privacy">
							<% small_help "Store Privacy Policy" %></td>
						</tr>




<% createFoot thisRedirect, 1%>

<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/sub.asp"-->


<%

if Session("Super_User") <> 1 then
     Response.redirect "noaccess.asp"
end if

sFormAction = "custom_purchase_admin.asp"
sTitle = "Custom Purchase"
thisRedirect = "custom_purchase_admin.asp"
sMenu="none"
createHead thisRedirect

sAmount=request("amount")
sDescription=request("description")
if sAmount <> "" then
  sql_update = "update store_settings set Custom_Amount="&sAmount&",custom_description='"&sDescription&"' where Store_id ="&Store_id
  conn_store.Execute sql_update

end if

%>
<tr bgcolor='#FFFFFF'>
					<td width="24%" height="23" class="inputname"><B>Amount Due</B></td>
					<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Amount" value="" size="5" maxlength=200>
							</td>
					</tr>
<tr bgcolor='#FFFFFF'>
					<td width="24%" height="23" class="inputname"><B>Description</B></td>
					<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Description" value="" size="50" maxlength=50>
							</td>
					</tr>

<% createFoot thisRedirect,1 %>



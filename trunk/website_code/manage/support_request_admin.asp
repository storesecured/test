<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%
on error goto 0
if request.form <> "" then
	Store_Id=checkStringForQ(request.form("Store_Id"))
	Subject = checkStringForQ(request.form("Subject"))
	Name = checkStringForQ(request.form("Name"))
	Email = checkStringForQ(request.form("Email"))
	Detail = checkStringForQ(request.form("Detail"))
	IP_Address= Request.ServerVariables("REMOTE_ADDR")

	sql_insert = "New_Support_Request1 "&Store_Id&",'"&Subject&"','"&Detail&"','"&Name&"','"&Email&"','"&IP_Address&"'"
	conn_store.Execute sql_insert

	response.redirect "support_list_admin.asp"
end if

sTitle = "Request Support"
sFormAction = "support_request_admin.asp"
thisRedirect = "support_request_admin.asp"
sMenu = "account"
createHead thisRedirect
%>


				<tr>
					 <td colspan="3"><input type="button" value="View all Requests" OnClick=JavaScript:self.location="support_list.asp"></td>
				 </tr>

				<tr>
					<td width="10%" class="inputname"><b>Name</b></td>
					<td width="90%" class="inputvalue"><input type="text" name="Name" value="" size="30">
					<INPUT type="hidden"  name="Name_C" value="Re|String|0|150|||Name">
					<% small_help "Name" %></td>
				</tr>
				<tr>
					<td width="10%" class="inputname"><b>Store Id</b></td>
					<td width="90%" class="inputvalue"><input type="text" name="Store_Id" value="" size="30">
					<INPUT type="hidden"  name="Store_Id_C" value="Re|String|0|150|||Store_Id">
					<% small_help "Store_Id" %></td>
				</tr>
				<tr>
					<td width="10%" class="inputname"><b>Email</b></td>
					<td width="90%" class="inputvalue"><input type="text" name="Email" value="<%= Store_Email %>" size="30">
					<INPUT type="hidden"  name="Email_C" value="Re|Email|0|150|||Email">
					<% small_help "Email" %></td>
				</tr>
				<tr>
					<td width="10%" class="inputname"><b>Subject</b></td>
					<td width="90%" class="inputvalue"><input type="text" name="Subject" value="" size="30">
					<INPUT type="hidden"  name="Subject_C" value="Re|String|0|150|||Subject">
					<% small_help "Subject" %></td>
				</tr>
				<tr>
					<td width="10%" class="inputname"><b>Detail</b></td>
					<td width="90%" class="inputvalue"><textarea name="Detail" rows=10 cols=55></textarea>
					<INPUT type="hidden"  name="Detail_C" value="Re|String|0|3000|||Detail">
					<% small_help "Detail" %></td>
				</tr>


<% createFoot thisRedirect, 1%>

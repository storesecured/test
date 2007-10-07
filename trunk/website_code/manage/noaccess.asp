<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "No Access"
thisRedirect = "noaccess.asp"
sMenu="logout"
createHead thisRedirect  %>



		<tr bgcolor='#FFFFFF'><td>You do not have access to options under the <%= request.querystring("sMenu") %> menu.
		<BR><BR>Please contact your store admin to gain access.</td></tr>


<% createFoot thisRedirect, 0 %>



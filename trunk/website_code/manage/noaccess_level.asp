<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "No Access"
thisRedirect = "noaccess.asp"
sMenu="logout"
createHead thisRedirect  %>



		<tr bgcolor='#FFFFFF'><td>You do not have access to this feature at your current level of service.
		<BR><BR>Unlimited Store Service is required.
		<BR><BR><a href=billing.asp class=link>Click here to upgrade now.</a>
		</td></tr>


<% createFoot thisRedirect, 0 %>



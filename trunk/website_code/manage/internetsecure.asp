<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "Internet Secure Instructions"
thisRedirect = "internetsecure.asp"
Secure = Replace(Site_Name,"http://","https://")
sMenu="advanced"
createHead thisRedirect  %>



		<TR><td><B>Instructions for setting up Internet Secure with Easystorecreator</td></tr>
		<TR><TD nowrap>1. Login to your Internet Secure account<BR>
		2. Select Export Scrips --> Export Script Options<BR>
		3. Set Server Domain Name to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Secure %></I><BR>
		4. Set Web Page to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>/include/internetsecure/internetsecure.asp?Store_Id=<%= Store_ID %></I><BR>
		5. Check the box for Send Approvals Only:<BR>
		6. Select Update
		</td></tr>

	  <TR><TD></td></tr>

<% createFoot thisRedirect, 0 %>



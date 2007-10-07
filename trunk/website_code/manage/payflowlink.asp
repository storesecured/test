<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "Payflow Link Instructions"
thisRedirect = "payflowlink.asp"
Secure = "https://"&Secure_Name&"/"
sMenu="general"
createHead thisRedirect  %>



		<TR bgcolor='#FFFFFF'><td><B>Instructions for setting up Payflow Link with Easystorecreator</td></tr>
		<TR bgcolor='#FFFFFF'><TD nowrap>1. Login to your Payflow Link account<BR>
		2. Select Account Info --> Payflow Link Info<BR>
		3. Set Return URL Method to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>Post</I><BR>
		4. Set Return URL to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Secure %>include/verisign/payflowlink.asp</I><BR>
		5. Set Silent Post URL to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Secure %>include/verisign/payflowlink.asp</I><BR>
		6. And enable the checkbox for Silent Post URL<BR>
		7. Make sure all of the Accepted URL fields are empty<BR>
      8. Make sure force silent post confirmation is unchecked<BR>
      9. Select save changes
		<BR><BR>
      <a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>
      </td></tr>


<% createFoot thisRedirect, 0 %>



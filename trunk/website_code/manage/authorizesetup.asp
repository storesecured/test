<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "Authorize.Net Instructions"
thisRedirect = "authorizesetup.asp"
Secure = Replace(Site_Name,"http://","https://")
sMenu="general"
createHead thisRedirect  %>



		<TR bgcolor='#FFFFFF'><td><B>Instructions for setting up Authorize.net with Easystorecreator</td></tr>
		<TR bgcolor='#FFFFFF'><TD>1. Login to your Authorize.net account<BR>
		2. In the left hand menu select Settings<BR>
		3. Towards the bottom select Obtain Transaction Key<BR>
		4. Create a new transaction key<BR>
		5. Copy the transaction key and paste it in the x_tran_key field under Use Authorize.Net<BR>
		6. Note: if you decide to type in the transaction key make sure it is typed in correctly, 1 wrong character or case will cause your transactions to be rejected by Authorize.Net<BR>
		7. Enter your Authorize.net user id in the space above the x_tran_key<BR>
		8. Make sure "Use Authorize.Net" is selected<BR>
		9. Hit Save
		<BR><BR>
      <a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>

		</td></tr>

	  <TR bgcolor='#FFFFFF'><TD></td></tr>

<% createFoot thisRedirect, 0 %>



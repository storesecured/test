<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "NoChex Instructions"
thisRedirect = "nochex.asp"
Secure = "https://"&Secure_Name&"/"
sMenu="general"
createHead thisRedirect  %>


      
		<TR bgcolor='#FFFFFF'><td><B>Instructions for setting up NoChex Version</td></tr>
		<TR bgcolor='#FFFFFF'><TD nowrap>1. Login to your NoChex account<BR>
		2. Select Control Panel --> Payment Page Details<BR>
		3. Set Callback URL and Success URL to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Secure %>include/nochex/nochex.asp</I><BR>
		4. Check the box labelled Auto Redirect<BR>
                5. Change other settings as needed  <BR>
                6. Select save changes <BR>

		<BR><BR>
      <a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>
      </td></tr>


<% createFoot thisRedirect, 0 %>



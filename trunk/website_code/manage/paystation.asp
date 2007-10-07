<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "Paystation Instructions"
thisRedirect = "paystation.asp"
Secure = "https://"&Secure_Name&"/"
sMenu="general"
createHead thisRedirect  %>



		<TR bgcolor='#FFFFFF'><td><B>Instructions for setting up PayStation</td></tr>
		<TR bgcolor='#FFFFFF'><TD>Email Paystation <I>info@paystation.co.nz</i> and ask that your return url be set to<BR><I><%= Secure %>include/paystation/paystation.asp</i><BR>

      <BR><a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>
      </td></tr>


<% createFoot thisRedirect, 0 %>



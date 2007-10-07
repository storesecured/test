<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "PaySystems Instructions"
thisRedirect = "paysystems.asp"
Secure = Replace(Site_Name,"http://","https://")
sMenu="advanced"
createHead thisRedirect  %>



		<TR><td><B>Instructions for setting up PaySystems with Easystorecreator</td></tr>
		<TR><TD>1. Ensure that you account is setup as a TPP-PRO account, this can be changed at anytime by calling PaySystems
		<BR><BR>
      <a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>
</td></tr>


<% createFoot thisRedirect, 0 %>



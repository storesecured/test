<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "Protx Instructions"
thisRedirect = "protx.asp"
Secure = Replace(Site_Name,"http://","https://")
sMenu="advanced"
createHead thisRedirect  %>



		<TR><td><B>Instructions for setting up Protx</td></tr>
		<TR><TD>Protx requires all users to start out in test mode and submit test transactions before approval.<BR>
		<BR>
                1. First enter your VendorName, Password and Currency<BR>
                2. Make sure Use Protx is selected<BR>
                3. Hit Save<BR>
                4. Submit a store support request asking to have your protx gateway switched to test mode<BR>
                5. We will confirm when the switch is ready<BR>
		6. Perform your test transactions<BR>
		7. After protx has approved your test transactions submit a support request asking for test mode to be removed.<BR>

		<BR><BR>
      <a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>

		</td></tr>



<% createFoot thisRedirect, 0 %>



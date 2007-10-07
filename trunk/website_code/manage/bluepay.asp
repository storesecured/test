<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "BluePay Instructions"
thisRedirect = "bluepay.asp"
Secure = Replace(Site_Name,"http://","https://")
sMenu="advanced"
createHead thisRedirect  %>



		<TR><td><B>Instructions for getting BluePay information to use in your store</td></tr>
		<TR><TD>1. Login to your BluePay account<BR>
		2. Select the administration tab<BR>
		3. From the dropdown select Accounts and then List<BR>
		4. On the right where the options are shown select the first option which is a face with no mouth.<BR>
		5. Under the 3rd heading titled Website Integration you should see your secret key.<BR>
		6. Under the 1st heading title Account Info find your account id which normally starts with a 1 and then a set of 0's<BR>
		7. Enter your BluePay Account Id and Secret key in the spots indicated<BR>
		8. Make sure "Use BluePay" is selected<BR>
		9. Hit Save
		<BR><BR>
      <a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>

		</td></tr>



<% createFoot thisRedirect, 0 %>



<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "Bank of America Instructions"
thisRedirect = "boa.asp"
Secure = "https://"&Secure_name&"/"
sMenu="general"
createHead thisRedirect  %>



		<TR bgcolor='#FFFFFF'><td><B>Instructions for setting up Bank of America Settle Up with Easystorecreator</td></tr>
		<TR bgcolor='#FFFFFF'><TD>1. Login to your Bank of America Settle Up account<BR>
		2. Select Manage Stores --> Order Rules --> Configure Options<BR>
		3. Set your store front URL:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Site_Name %></I><BR>
		4. Check the radio button next to:<BR>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>Redirect to merchant's approval/decline pages</i><BR>
		5. Set Approvals to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Secure %>include/boa/boa.asp</i><BR>
		6. Set Rejections to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Secure %>include/boa/boareject.asp</i><BR>
		7. Check the radio button next to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>Handle multiple concurrent authorization requests per shopper cookie.</i><BR>
		8. Uncheck the checkbox next to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>Use ioc_merchant_order_id to enforce unique orders posting (this option enforces uniqueness for a 7-day period)</i><BR>
		9. Uncheck the checkbox next to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>Click here if you want to decline orders taken if the connection to the payment processor is not available</i><BR>
		10. Select Update<BR>
		11. Publish your changes
		<BR><BR>
      <a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>
</td></tr>

<% createFoot thisRedirect, 0 %>



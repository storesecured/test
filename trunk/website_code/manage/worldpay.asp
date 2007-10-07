<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "WorldPay Instructions"
thisRedirect = "worldpay.asp"
Secure = "https://"&Secure_Name&"/"
sMenu="general"
createHead thisRedirect  %>



		<TR bgcolor='#FFFFFF'><td><B>Instructions for setting up WorldPay</td></tr>
		<TR bgcolor='#FFFFFF'><TD>1. Login to your WorldPay account<BR>
		2. Select Configuration Options under the heading Installations for XXX<BR>
		3. Make sure the Integration type is listed as Select Junior, if not please contact WorldPay to change<BR>
		4. Set Site URL to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Site_Name %></I><BR>
		5. Set Store Builder used to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>StoreSecured</I><BR>
		6. Set Callback URL to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Secure %>include/worldpay/worldpay.asp</I><BR>
		7. Check the box marked Callback Enabled<BR>
		8. Select save changes<BR>
		9. Enter your WorldPay Installation Id and default currency on the Payment Processor Page
		<BR><BR>
      <a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>
</td></tr>


<% createFoot thisRedirect, 0 %>



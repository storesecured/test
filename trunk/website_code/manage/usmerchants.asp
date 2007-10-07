<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "US Merchant Credit Card Processing"
thisRedirect = "usmerchants.asp"
sMenu="none"
createHead thisRedirect  %>


			  </form>
		<TR bgcolor="#FFFFFF"><TD colspan=2><HR></td></tr>
			 <tr bgcolor="#FFFFFF"><td>
		<UL><LI>No Application Fee
		<LI>$25 Monthly Fee
		<LI>2.25% Discount Rate
		<LI>25¢ Transaction Fee
		<LI>$25 Monthly Minimum
		<LI>Approval in 2-3 days
		</UL>
		</td><td align=center width=200>Recommended for businesses who will be processing > $2000 per month<BR>
		<form action=https://www.bluemerchant.com/bluepay/merchantform/form0.asp?hostID=506302><input type=submit name=TMS value=Apply></form></td></tr>

		<TR bgcolor="#FFFFFF"><td colspan=2><HR></td></tr>


			 <tr bgcolor="#FFFFFF"><td><UL><LI>$149 Application Fee
			 <LI>$29.95 Monthly Fee
			 <LI>2.39% Discount Rate
			 <LI>35¢ Transaction Fee
			 <LI>$25 Monthly Minimum
			 <LI>eCheck, American Express and Discover included
			 <LI>Instant approval for credit score over 600
			 </UL></td>
			 <td align=center width=200>Recommended for businesses who need to be processing today and want an easy application<BR>
		<form action="http://www.authorizenet.com/launch/ips_main.php" name="EMS" method="Post">
<input type="hidden" name="anUserID" value="easystorecreateapp">
<input type=submit name="Authorize.Net" value="Apply">
<BR><BR></form></td></tr>
		<TR bgcolor="#FFFFFF"><TD colspan=2><HR></td></tr>
			 <tr bgcolor="#FFFFFF"><td>
			 <UL><LI>No Application Fee
			 <LI>$20 Monthly Fee
			 <LI>2.9% Discount Rate
			 <LI>30¢ Transaction Fee
			 <LI>No Monthly Minimum
			 <LI>Free processing on express checkout till 1/1/2008
			 </UL></td>
			 <td align=center width=200>Recommended for businesses who will be processing < $2000 per month<form action=http://www.paypal-promo.com/searchmarketing/tracking/media_source/easystore.html><input type=submit name=PlugnPay value=Apply></form></td></tr>




<% createFoot thisRedirect, 0 %>



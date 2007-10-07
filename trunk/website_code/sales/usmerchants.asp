<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "US Merchant Credit Card Processing"
thisRedirect = "usmerchants.asp"
sMenu="none"
createHead thisRedirect  %>


			  </form>
		<TR><TD colspan=2><HR></td></tr>
			 <tr><td>
		<UL><LI>$0 Application Fee
		<LI>$25 Monthly Fee
		<LI>2.25% Discount Rate
		<LI>25¢ Transaction fee
		<LI>$25 Monthly Minimum
		<LI>Electronic application
		<LI>Approval in 2-3 days
		<LI>Includes merchant account
		</UL>
		</td><td><form action=http://webpaymentsolutions.com/application/?id=104479&pid=11048><input type=submit name=TMS value=Apply></form></td></tr>
		
		<TR><TD colspan=2><HR></td></tr>
			 <tr><td>
		<UL><LI>$0 Application Fee
		<LI>$49 Annual Fee
		<LI>3.79% Discount Rate
		<LI>50¢ Transaction fee
			 <LI>No monthly minimums
			 <LI>Includes merchant account
		</UL>
		</td><td><form action=http://www.payquake.com/pl_lib/PSignUp.asp?SourceEntityID=11789><input type=submit name=PayQuake value=Apply></form></td></tr>
		
		<TR><td colspan=2><HR></td></tr>


			 <tr><td><UL><LI>$149 Application Fee
			 <LI>$29.95 Monthly Fee
			 <LI>2.39% Discount Rate
			 <LI>35¢ transaction fee
			 <LI>eCheck included
			 <LI>Electronic Application
			 <LI>Instant approval for credit score over 600
			 <LI>Includes merchant account
			 </UL></td>
			 <td><form action="http://www.authorizenet.com/launch/ips_main.php" name="EMS" method="Post">
<input type="hidden" name="anUserID" value="easystorecreateapp">
<input type=submit name="Authorize.Net" value="Apply">
<BR><BR></form></td></tr>


		<TR><TD colspan=2><HR></td></tr>
			 <tr><td>
		<UL><LI>$150 Application Fee
		<LI>$10 Monthly Fee (waived if no transactions)
		<LI>2.69% Discount Rate
		<LI>30¢ transaction fee
		<LI>eCheck included
		<LI>Site inspection required
			 <LI>Mail-in application
			 <LI>No monthly minimums
			 <LI>Includes merchant account
		</UL>
		</td><td><form action=http://www<%= root %>/formmail_echo.asp><input type=submit name=Echo value=Apply></form></td></tr>
		<TR><TD colspan=2><HR></td></tr>
			 <tr><td>
			 <UL><LI>$75 Application Fee
			 <LI>$15 Monthly Fee
			 <LI>First 100 transactions free
			 <LI>After 100 each transaction 15¢
			 <LI>Electronic Applications
			 <LI>Approval in approx 2 days
			 <LI>Requires separate merchant account
			 </UL></td>
			 <td><form action=https://www.onlinedatacorp.com/onlineapp/entry.asp?hostID=504899><input type=submit name=PlugnPay value=Apply></form></td></tr>
			 <TR><TD colspan=2><HR></td></tr>
			 <tr><td>
			 <UL><LI>$49 Application Fee
			 <LI>$0 Monthly Fee
			 <LI>5.5% Discount Rate
			 <LI>45¢ transaction fee
			 <LI>Instant Approval
			 <LI>Customers will be redirected on checkout to 3rd party site for payment processing
			 <LI>Customers billing statment will not show merchants name
			 <LI>High Risk Merchants Accepted
			 </UL></td>
			 <td><form action=http://www.2checkout.com/cgi-bin/aff.2c?affid=74224><input type=submit name=2Checkout value=Apply></form></td></tr>

	  <TR><TD></td></tr>

<% createFoot thisRedirect, 0 %>


